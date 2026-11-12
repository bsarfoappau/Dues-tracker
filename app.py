"""
Dues Tracker (Supabase edition)

* Supabase Auth handles sign up and sign in (the Group Name is captured at sign up).
* Supabase Postgres stores everything. Row Level Security (RLS) makes sure every
  account can only ever read and write its own rows.
* Normal pages only use the public "anon" key plus the signed-in user's own token.
* The Admin page is the one exception. It uses the service_role key (read from secrets, on the
  server only) to manage accounts, and only after the signed-in user's role has been re-checked
  as "admin" in the database. Never expose that key anywhere else.
* Passwords can be reset from the sign-in screen (Supabase emails a link back to this app).
"""
import io
import re
import copy
import json
import html
import time
import hashlib
import inspect
import threading
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, date, timezone, timedelta

import altair as alt
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# ─────────────────────────────────────────────────────────────
# Page config (must be first Streamlit call)
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dues Tracker",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="collapsed",
)

try:
    from supabase import create_client
except ImportError:  # pragma: no cover
    st.error("The supabase package is missing. Run: pip install -r requirements.txt")
    st.stop()


def _full(fn) -> dict:
    """Full-width keyword that works on older and newer Streamlit releases."""
    try:
        major, minor = (int(x) for x in st.__version__.split(".")[:2])
        newer = (major, minor) >= (1, 50)
    except Exception:
        newer = False
    try:
        has_old = "use_container_width" in inspect.signature(fn).parameters
    except (TypeError, ValueError):
        has_old = True
    if newer or not has_old:
        return {"width": "stretch"}
    return {"use_container_width": True}


FULL_BTN = _full(st.button)
FULL_SUBMIT = _full(st.form_submit_button)
FULL_DL = _full(st.download_button)
FULL_POP = _full(st.popover)
FULL_CHART = _full(st.altair_chart)

# ─────────────────────────────────────────────────────────────
# Palette (solid, vivid colours)
# ─────────────────────────────────────────────────────────────
COBALT = "#2F5BFF"
COBALT_DARK = "#1E45D8"
MINT = "#00B386"
CORAL = "#FF4D6D"
ORANGE = "#FF8A1F"
VIOLET = "#7A4DFF"
SUN = "#FFC93C"
INDIGO = "#231B6B"
INDIGO_SOFT = "#3A2F95"
PAGE_BG = "#F3F4FA"
AVATARS = [COBALT, MINT, CORAL, ORANGE, VIOLET, "#0096C7"]

# One vivid colour per contribution, used consistently in every chart
CHART_COLORS = [COBALT, "#00C48C", "#FF5C7A", "#FF9F1C", "#8B5CF6",
                "#00B4D8", "#FFC93C", "#E83E8C", "#7BD389", "#FF6B35"]

STATUS = {
    "paid": ("Paid", MINT),
    "partial": ("Partial", ORANGE),
    "unpaid": ("Unpaid", CORAL),
    "waived": ("Waived", VIOLET),
}

# Light fills used in the Excel and Word exports
STATUS_FILL = {"Paid": "C6EFCE", "Partial": "FFE2C2", "Unpaid": "FFC7CE", "Waived": "E4DBFF", "Not in": "ECECEC"}


def status_label(s: str) -> str:
    """Export label for a status. "out" means the member is not part of that contribution."""
    return "Not in" if s == "out" else STATUS[s][0]

# ─────────────────────────────────────────────────────────────
# Styling (written with __TOKENS__ so CSS braces stay readable)
# ─────────────────────────────────────────────────────────────
_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@600;700;800&family=Manrope:wght@400;500;600;700&display=swap');

.stApp { font-family: 'Manrope', sans-serif; background: __PAGE_BG__; }
.stApp h1, .stApp h2, .stApp h3, .stApp h4 { font-family: 'Sora', sans-serif; letter-spacing: -0.02em; }
html { -webkit-text-size-adjust: 100%; }
#MainMenu, footer, .stDeployButton, [data-testid="stToolbar"] { visibility: hidden; }
header[data-testid="stHeader"] { display: none; }
section[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"],
[data-testid="collapsedControl"] { display: none !important; }
.block-container { padding: 1.2rem 1.2rem 2rem !important; max-width: 920px !important; }
/* the style-only markdown block should not leave a blank gap */
.stApp [data-testid="stElementContainer"]:has(> [data-testid="stMarkdownContainer"] > style) { display: none; }

/* Buttons (thumb-sized) */
.stApp [data-testid^="stBaseButton-primary"] {
    background: __COBALT__; color: #fff; border: 2px solid __COBALT__;
    border-radius: 999px; font-weight: 700; padding: .35rem 1rem;
}
.stApp [data-testid^="stBaseButton-primary"]:hover {
    background: __COBALT_DARK__; border-color: __COBALT_DARK__; color: #fff;
}
.stApp [data-testid^="stBaseButton-secondary"],
.stApp [data-testid="stPopover"] > button,
.stApp [data-testid="stPopoverButton"] {
    background: #fff; color: __COBALT__; border: 2px solid __COBALT__;
    border-radius: 999px; font-weight: 700; padding: .35rem .8rem;
}
.stApp [data-testid^="stBaseButton-secondary"]:hover,
.stApp [data-testid="stPopover"] > button:hover,
.stApp [data-testid="stPopoverButton"]:hover {
    background: __COBALT__; color: #fff; border-color: __COBALT__;
}
.stApp .stButton button, .stApp [data-testid="stPopover"] > button,
.stApp [data-testid="stDownloadButton"] button { min-height: 44px; }
.stApp .stButton button p, .stApp [data-testid="stPopover"] button p { color: inherit; }

/* Never cut button labels short */
.stApp .stButton button p,
.stApp .stButton button div,
.stApp [data-testid="stPopover"] button p,
.stApp [data-testid="stPopover"] button div,
.stApp [data-testid="stPopoverButton"] p,
.stApp [data-testid="stPopoverButton"] div,
.stApp [data-testid="stDownloadButton"] button p {
    overflow: visible !important; text-overflow: clip !important; white-space: nowrap !important;
}
.stApp [data-testid="stPopover"] button [data-testid="stIconMaterial"] { display: none; }

.stApp [data-testid="stDownloadButton"] button {
    background: __MINT__; color: #fff; border: 2px solid __MINT__;
    border-radius: 999px; font-weight: 700; padding: .45rem 1rem;
}
.stApp [data-testid="stDownloadButton"] button:hover { background: #008A63; border-color: #008A63; color: #fff; }
.stApp [data-testid="stDownloadButton"] button p { color: inherit; }

/* Colour-coded action buttons (matched by widget key) */
.stApp [class*="st-key-pay_"] button,
.stApp [class*="st-key-markall_"] button {
    background: __MINT__ !important; border-color: __MINT__ !important; color: #fff !important;
}
.stApp [class*="st-key-pay_"] button:hover,
.stApp [class*="st-key-markall_"] button:hover { background: #008A63 !important; border-color: #008A63 !important; }
.stApp [class*="st-key-del_"] button {
    background: __CORAL__ !important; border-color: __CORAL__ !important; color: #fff !important;
}
.stApp [class*="st-key-del_"] button:hover { background: #E0304F !important; border-color: #E0304F !important; }
.stApp [class*="st-key-undo_"] button {
    background: #fff !important; border-color: __CORAL__ !important; color: __CORAL__ !important;
}
.stApp [class*="st-key-undo_"] button:hover { background: __CORAL__ !important; color: #fff !important; }
.stApp .st-key-logout button { background: __CORAL__ !important; border-color: __CORAL__ !important; color: #fff !important; }
.stApp .st-key-logout button:hover { background: #E0304F !important; border-color: #E0304F !important; }

/* Pop-up dialogs use the same button colours (they render outside .stApp) */
div[role="dialog"] [class*="st-key-del_"] button {
    background: __CORAL__ !important; border: 2px solid __CORAL__ !important; color: #fff !important;
    border-radius: 999px; font-weight: 700; min-height: 44px;
}
div[role="dialog"] [class*="st-key-del_"] button:hover { background: #E0304F !important; border-color: #E0304F !important; }
div[role="dialog"] [class*="st-key-dlg_cancel"] button {
    background: #fff !important; border: 2px solid __COBALT__ !important; color: __COBALT__ !important;
    border-radius: 999px; font-weight: 700; min-height: 44px;
}
div[role="dialog"] button p { color: inherit !important; }

/* Inputs: 16px text stops iOS zooming the page on focus */
.stApp input, .stApp textarea, .stApp [data-baseweb="select"] *,
div[data-testid="stPopoverBody"] input, div[data-testid="stPopoverBody"] textarea { font-size: 16px !important; }
.stApp [data-baseweb="input"], .stApp [data-baseweb="base-input"],
.stApp [data-baseweb="textarea"], .stApp [data-baseweb="select"] > div { border-radius: 14px !important; }
div[data-testid="stPopoverBody"] { max-width: min(94vw, 440px) !important; }

/* Readable on phones in dark mode too: tabs, inputs, forms, expanders */
.stApp [data-baseweb="tab-list"] { gap: .4rem; }
.stApp [data-baseweb="tab-list"] button p { color: #5B5E80 !important; font-weight: 700; font-size: 1rem; }
.stApp [data-baseweb="tab-list"] button[aria-selected="true"] p { color: __COBALT__ !important; }
.stApp [data-baseweb="tab-highlight"] { background-color: __COBALT__ !important; }
.stApp [data-baseweb="tab-border"] { background-color: #E4E6F4 !important; }
.stApp [data-baseweb="input"], .stApp [data-baseweb="base-input"],
.stApp [data-baseweb="textarea"], .stApp [data-baseweb="select"] > div {
    background: #F3F4FA !important; border: 1px solid #D5D8EE !important;
}
.stApp input, .stApp textarea { color: #17153B !important; -webkit-text-fill-color: #17153B !important; background: transparent !important; }
.stApp input::placeholder, .stApp textarea::placeholder { color: #7C7FA3 !important; opacity: 1 !important; -webkit-text-fill-color: #7C7FA3 !important; }
.stApp [data-testid="stWidgetLabel"] p, .stApp label p { color: #17153B !important; }
.stApp [data-testid="stForm"] { border: 1px solid #E4E6F4 !important; border-radius: 18px; background: #fff; }
.stApp [data-testid="stExpander"] details { background: #fff; border: 2px solid #E4E6F4 !important; border-radius: 14px; }
.stApp [data-testid="stExpander"] summary p, .stApp [data-testid="stExpander"] summary span,
.stApp [data-testid="stExpander"] summary svg { color: #17153B !important; fill: #17153B !important; }
.stApp [class*="st-key-hrow_report"] button { padding: .45rem .5rem !important; }

/* Cards */
.stApp [class*="st-key-card_"], .stApp [class*="st-key-mcard_"] {
    background: #fff; border: 2px solid #E4E6F4; border-radius: 20px;
    padding: .7rem 1.1rem; box-shadow: 0 4px 14px rgba(35, 27, 107, .06);
}
.stApp [class*="st-key-mcard_"] { gap: .55rem !important; padding: .75rem .9rem .8rem; }
.stApp .st-key-mlist { gap: .65rem !important; }
.stApp [class*="st-key-card_chart"] { box-shadow: none; padding: 1rem 1.2rem 1.1rem; }
.stApp [data-testid="stVegaLiteChart"], .stApp [data-testid="stArrowVegaLiteChart"] { filter: none !important; }

/* Rows that must stay side by side on phones (Streamlit would stack them) */
.stApp [class*="st-key-hrow_"] [data-testid="stHorizontalBlock"],
.stApp .st-key-tnav [data-testid="stHorizontalBlock"] { flex-wrap: nowrap !important; gap: .5rem !important; }
.stApp [class*="st-key-hrow_"] [data-testid="stColumn"],
.stApp [class*="st-key-hrow_"] [data-testid="column"],
.stApp .st-key-tnav [data-testid="stColumn"],
.stApp .st-key-tnav [data-testid="column"] { flex: 1 1 0 !important; width: auto !important; min-width: 0 !important; }
.stApp .st-key-hrow_month [data-testid="stColumn"]:nth-child(2),
.stApp .st-key-hrow_month [data-testid="column"]:nth-child(2) { flex: 3 1 0 !important; }
.stApp .st-key-hrow_msearch [data-testid="stColumn"]:nth-child(1),
.stApp .st-key-hrow_msearch [data-testid="column"]:nth-child(1) { flex: 2 1 0 !important; }
.stApp .st-key-hrow_mhead [data-testid="stColumn"]:nth-child(1),
.stApp .st-key-hrow_mhead [data-testid="column"]:nth-child(1) { flex: 1.3 1 0 !important; }

.stApp .st-key-hrow_top [data-testid="stColumn"]:nth-child(1),
.stApp .st-key-hrow_top [data-testid="column"]:nth-child(1) { flex: 1 1 0 !important; }
.stApp .st-key-hrow_top [data-testid="stColumn"]:nth-child(2),
.stApp .st-key-hrow_top [data-testid="column"]:nth-child(2) { flex: 0 0 52px !important; }
.stApp .st-key-hrow_top [data-testid="stPopover"] > button { min-height: 48px; padding: 0; border-radius: 14px; }
.stApp .st-key-hrow_top [data-testid="stPopover"] > button p { font-size: 1.5rem; line-height: 1; }

/* Member card action buttons: sized to their label, left aligned */
.stApp [class*="st-key-mcard_"] [data-testid="stHorizontalBlock"] {
    flex-wrap: wrap !important; justify-content: flex-start !important; gap: .5rem !important; margin-top: .3rem;
}
.stApp [class*="st-key-mcard_"] [data-testid="stColumn"],
.stApp [class*="st-key-mcard_"] [data-testid="column"] { flex: 0 0 auto !important; width: auto !important; min-width: 0 !important; }
.stApp [class*="st-key-mcard_"] button { min-width: 96px; min-height: 40px; padding: .25rem 1.1rem; }
/* Streamlit pulls markdown blocks up by 1rem; cancel that inside cards so text is never covered */
.stApp [class*="st-key-mcard_"] [data-testid="stMarkdownContainer"] { margin-bottom: 0 !important; }

/* Menu: a short list of actions */
[class*="st-key-mnu_"] button {
    justify-content: flex-start !important; text-align: left !important; background: #fff !important;
    border: 0 !important; border-bottom: 1px solid #EEF0FA !important; border-radius: 0 !important;
    color: #17153B !important; font-weight: 600 !important; min-height: 46px !important; padding: .4rem .3rem !important;
}
[class*="st-key-mnu_"] button:hover { background: #F3F4FA !important; }
[class*="st-key-mnu_"] button div, [class*="st-key-mnu_"] button p { justify-content: flex-start !important; text-align: left !important; width: 100%; }
.st-key-mnu_back button { color: __COBALT__ !important; border-bottom: 0 !important; min-height: 38px !important; }
.menu-head { font-family: 'Sora', sans-serif; font-weight: 700; color: #17153B; margin: .1rem 0 .4rem; }
.menu-sub { color: #6B6E8F; font-size: .82rem; font-weight: 600; padding: .1rem .3rem .3rem; }
div[data-testid="stPopoverBody"] { min-width: min(86vw, 260px); }

/* Top navigation: sits in the normal page flow, so nothing is fixed over the content on phones */
.stApp .st-key-tnav {
    background: #fff; border: 2px solid #E4E6F4; border-radius: 18px;
    padding: .35rem .45rem; margin-bottom: 1rem; box-shadow: 0 4px 14px rgba(35, 27, 107, .06);
}
.stApp .st-key-tnav button {
    background: transparent !important; border: 2px solid transparent !important; color: #6B6E8F !important;
    border-radius: 12px; padding: .2rem 0; min-height: 44px;
}
.stApp .st-key-tnav button:hover { background: #F3F4FA !important; color: __COBALT__ !important; }
.stApp .st-key-tnav button p {
    white-space: nowrap !important; text-align: center; line-height: 1.2; font-size: .82rem;
    font-weight: 700; color: inherit !important;
}
.stApp .st-key-tnav button[data-testid="stBaseButton-primary"] { background: #EEF1FF !important; color: __COBALT__ !important; }

/* Group banner (top of the dashboard) */
.group-banner {
    background: linear-gradient(135deg, __INDIGO__, __INDIGO_SOFT__); color: #fff;
    border-radius: 20px; padding: .85rem 1.2rem; margin-bottom: 1rem;
    box-shadow: 0 6px 18px rgba(35, 27, 107, .18);
}
.group-banner .gb-label { font-size: .7rem; font-weight: 700; letter-spacing: .1em; text-transform: uppercase; opacity: .75; }
.group-banner .gb-name { font-family: 'Sora', sans-serif; font-weight: 800; font-size: 1.55rem; line-height: 1.15;
                         overflow-wrap: anywhere; }

/* Page headings */
.page-title { font-family: 'Sora', sans-serif; font-weight: 800; font-size: 2rem; color: #17153B; line-height: 1.1; }
.page-sub { color: #5B5E80; font-weight: 500; margin: .25rem 0 1rem; }
.section-title { font-family: 'Sora', sans-serif; font-weight: 700; font-size: 1.05rem; color: #17153B; margin-bottom: .4rem; }
.big-title { font-family: 'Sora', sans-serif; font-weight: 800; font-size: 1.5rem; color: #17153B; margin: .3rem 0 .2rem; }
.month-label { font-family: 'Sora', sans-serif; font-weight: 700; color: #17153B; text-align: center; font-size: 1rem; }

/* Tiles: 4 across on desktop, 2 x 2 on phones */
.tiles { display: grid; grid-template-columns: repeat(4, 1fr); gap: .8rem; }
.tile { border-radius: 22px; padding: 1.1rem 1.3rem; min-height: 124px; min-width: 0; }
.tile-label { font-weight: 700; font-size: .95rem; opacity: .95; }
.tile-value { font-family: 'Sora', sans-serif; font-weight: 800; font-size: 2.1rem; line-height: 1.2; margin-top: .15rem;
              overflow-wrap: anywhere; }
.tile-sub { font-weight: 600; font-size: .85rem; opacity: .9; margin-top: .1rem; }

/* Progress */
.bar { background: #DADDF0; height: 14px; border-radius: 999px; overflow: hidden; margin: .5rem 0 .2rem; }
.bar-fill { background: __MINT__; height: 100%; border-radius: 999px; }
.bar-label { color: #5B5E80; font-weight: 600; font-size: .9rem; }

/* Member rows */
.mrow { display: flex; align-items: center; justify-content: space-between; gap: .6rem; }
.mr { text-align: right; flex: none; max-width: 46%; }
.mr .chip { margin-bottom: .2rem; }
.mr .amt { font-size: .95rem; line-height: 1.2; }
.person { display: flex; align-items: center; gap: .7rem; min-width: 0; flex: 1 1 auto; }
.person > div:last-child { min-width: 0; }
.avatar { width: 42px; height: 42px; border-radius: 50%; color: #fff; font-weight: 800;
          display: flex; align-items: center; justify-content: center; flex: none; font-family: 'Sora', sans-serif; }
.p-name { font-weight: 700; color: #17153B; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.p-sub { color: #6B6E8F; font-size: .85rem; font-weight: 500; }
.ell { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.chip { display: inline-block; color: #fff; font-weight: 700; font-size: .8rem; padding: .22rem .75rem; border-radius: 999px; }
.amt { font-weight: 800; color: #17153B; font-family: 'Sora', sans-serif; }
.mnote { display: flex; justify-content: space-between; gap: .6rem; }
.empty { background: #fff; border: 2px dashed #C9CDE8; border-radius: 22px; padding: 2.2rem; text-align: center; color: #5B5E80; }
.empty b { color: #17153B; font-family: 'Sora', sans-serif; font-size: 1.2rem; }

/* Titles inside cards and rows */
.ctitle { font-family: 'Sora', sans-serif; font-weight: 700; font-size: 1.05rem; color: #17153B; line-height: 1.25; }
.csub { color: #6B6E8F; font-size: .85rem; font-weight: 500; margin: .1rem 0 .7rem; }
.mamt { display: flex; justify-content: space-between; align-items: baseline; gap: .6rem;
        padding-top: .45rem; border-top: 1px solid #EEF0FA; }
.mamt .amt { font-size: 1rem; }

/* Donut (pure CSS so it can never be clipped) */
.donut-wrap { display: flex; align-items: center; justify-content: center; gap: 1.6rem; flex-wrap: wrap; padding: .2rem 0 .3rem; }
.donut { width: 150px; height: 150px; border-radius: 50%; display: flex; align-items: center; justify-content: center; flex: none; }
.donut-hole { width: 96px; height: 96px; border-radius: 50%; background: #fff; display: flex; flex-direction: column;
              align-items: center; justify-content: center; text-align: center; line-height: 1.15; }
.donut-hole b { font-family: 'Sora', sans-serif; font-size: 1.55rem; color: #17153B; }
.donut-hole span { font-size: .72rem; color: #6B6E8F; font-weight: 600; }
.legend { display: grid; gap: .5rem; min-width: 150px; }
.lg-row { display: flex; align-items: center; gap: .55rem; }
.lg-dot { width: 12px; height: 12px; border-radius: 50%; flex: none; }
.lg-name { flex: 1; font-weight: 600; color: #17153B; font-size: .9rem; }
.lg-n { font-weight: 800; font-family: 'Sora', sans-serif; color: #17153B; }

.copyright { text-align: center; color: #8A8DAB; font-size: .8rem; font-weight: 600; padding: 1.6rem 0 .4rem; }

/* History lists */
.reclist { background: #fff; border: 2px solid #E4E6F4; border-radius: 18px; padding: .2rem 1rem;
           box-shadow: 0 4px 14px rgba(35, 27, 107, .06); }
.rec { padding: .7rem 0; border-bottom: 1px solid #EEF0FA; }
.rec:last-child { border-bottom: 0; }
.rec-top { display: flex; justify-content: space-between; align-items: baseline; gap: .6rem; }
.rec-bot { display: flex; align-items: center; gap: .5rem; margin-top: .3rem; flex-wrap: wrap; }

/* Phones */
@media (max-width: 640px) {
    .block-container { padding: .8rem .75rem 1.5rem !important; }
    .page-title { font-size: 1.55rem; }
    .page-sub { font-size: .9rem; margin-bottom: .8rem; }
    .big-title { font-size: 1.2rem; }
    .section-title { font-size: 1rem; }
    .group-banner { padding: .7rem 1rem; border-radius: 16px; }
    .group-banner .gb-name { font-size: 1.25rem; }
    .tiles { grid-template-columns: repeat(2, 1fr); gap: .6rem; }
    .tile { min-height: 0; padding: .8rem .9rem; border-radius: 18px; }
    .tile-label { font-size: .85rem; }
    .tile-value { font-size: clamp(1.05rem, 5.4vw, 1.4rem); }
    .tile-sub { font-size: .78rem; }
    .stApp [class*="st-key-card_"] { padding: .6rem .8rem; border-radius: 16px; }
    .stApp [class*="st-key-card_chart"] { padding: .8rem .8rem .9rem; }
    .empty { padding: 1.4rem; }
    .stApp .st-key-tnav { padding: .3rem .3rem; border-radius: 14px; margin-bottom: .8rem; }
    .stApp .st-key-tnav [data-testid="stHorizontalBlock"] { gap: .2rem !important; }
    .stApp .st-key-tnav button p { font-size: .68rem; }
}
</style>
"""

_LOGIN_CSS = """
<style>
.stApp { background: #FFFFFF; }
.st-key-login_card {
    max-width: 440px; margin: 6vh auto 0; background: #fff; border: 2px solid #E4E6F4;
    border-radius: 28px; padding: 2rem 1.6rem 1.6rem; box-shadow: 0 12px 40px rgba(35, 27, 107, .08);
}
.login-title { font-family: 'Sora', sans-serif; font-weight: 800; font-size: 1.7rem; color: __INDIGO__; text-align: center; }
.login-sub { color: #5B5E80; text-align: center; font-weight: 500; margin: .3rem 0 1rem; }
.stApp .st-key-forgot_open button {
    background: transparent !important; border: 0 !important; color: __COBALT__ !important;
    text-decoration: underline; min-height: 40px; padding: .2rem .4rem !important; font-weight: 600;
}
.stApp .st-key-forgot_open button:hover { background: transparent !important; color: __COBALT_DARK__ !important; }
@media (max-width: 640px) { .st-key-login_card { margin-top: 3vh; } }
</style>
"""


def _paint(css: str) -> str:
    tokens = {"PAGE_BG": PAGE_BG, "COBALT": COBALT, "COBALT_DARK": COBALT_DARK, "MINT": MINT,
              "CORAL": CORAL, "INDIGO": INDIGO, "INDIGO_SOFT": INDIGO_SOFT}
    for k, v in tokens.items():
        css = css.replace(f"__{k}__", v)
    return css


def copyright_html() -> str:
    return f'<div class="copyright">&copy; {datetime.now().year} AppaTech. All rights reserved.</div>'


BASE_CSS = _paint(_CSS)
LOGIN_CSS = _paint(_LOGIN_CSS)


# ─────────────────────────────────────────────────────────────
# Settings, limits and small helpers
# ─────────────────────────────────────────────────────────────
DEFAULT_SETTINGS = {"group_name": "My Group", "currency": "GH₵", "default_dues": 0.0}
SETTINGS = dict(DEFAULT_SETTINGS)          # replaced with the signed-in user's profile after login

# Protected actions: after MAX_ATTEMPTS wrong passwords the user is locked out for LOCKOUT_SECONDS
MAX_ATTEMPTS = 3
LOCKOUT_SECONDS = 300
MIN_PASSWORD = 8

esc = html.escape


def mid(x) -> int:
    """Id as int (safe for ids that came back as floats or dicts)."""
    if isinstance(x, dict):
        x = x.get("id", 0)
    try:
        return int(x)
    except Exception:
        return 0


def month_label(m: str) -> str:
    try:
        return datetime.strptime(m + "-01", "%Y-%m-%d").strftime("%B %Y")
    except Exception:
        return m


def month_short(m: str) -> str:
    try:
        return datetime.strptime(m + "-01", "%Y-%m-%d").strftime("%b %y")
    except Exception:
        return m


def shift_month(m: str, delta: int) -> str:
    y, mo = map(int, m.split("-"))
    idx = y * 12 + (mo - 1) + delta
    return f"{idx // 12:04d}-{idx % 12 + 1:02d}"


def norm_status(s) -> str:
    s = str(s or "").strip().lower()
    if s in ("pending", "not recorded", "", "nan", "none"):
        return "unpaid"
    return s if s in STATUS else "unpaid"


def to_float(v) -> float:
    try:
        return float(v or 0)
    except Exception:
        return 0.0


def now_utc() -> datetime:
    return datetime.now(timezone.utc)


def now_iso() -> str:
    return now_utc().isoformat(timespec="seconds")


def parse_ts(s):
    """Parse a timestamp from Postgres into an aware datetime (or None)."""
    if not s:
        return None
    try:
        dt = datetime.fromisoformat(str(s).replace("Z", "+00:00"))
        return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)
    except Exception:
        return None


def fmt_ts(s) -> str:
    return str(s or "")[:19].replace("T", " ")


def flash(msg: str, icon: str = "✅"):
    st.session_state["_flash"] = (msg, icon)


def show_flash():
    if "_flash" in st.session_state:
        msg, icon = st.session_state.pop("_flash")
        st.toast(msg)


def initials(name: str) -> str:
    parts = [p for p in str(name).split() if p]
    if not parts:
        return "?"
    if len(parts) == 1:
        return parts[0][:2].upper()
    return (parts[0][0] + parts[-1][0]).upper()


def color_of(c: dict) -> str:
    return CHART_COLORS[(mid(c) - 1) % len(CHART_COLORS)]


def is_monthly(c: dict) -> bool:
    return c.get("frequency") == "monthly"


# ─────────────────────────────────────────────────────────────
# Supabase client
# One client per browser session. After sign in it carries the user's JWT, so every
# query is evaluated by Postgres as that user and Row Level Security applies.
# ─────────────────────────────────────────────────────────────
def get_secret(key, default=""):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


def supabase_config() -> tuple:
    """Reads [supabase] url / anon_key (or SUPABASE_URL / SUPABASE_ANON_KEY) from secrets."""
    url = key = ""
    section = get_secret("supabase", None)
    if section:
        try:
            url = str(section.get("url", ""))
            key = str(section.get("anon_key", ""))
        except Exception:
            pass
    url = url or str(get_secret("SUPABASE_URL", ""))
    key = key or str(get_secret("SUPABASE_ANON_KEY", ""))
    return url.strip(), key.strip()


def new_client(implicit: bool = False):
    url, key = supabase_config()
    if not url or not key:
        st.error("Supabase is not configured. Add url and anon_key under [supabase] in "
                 ".streamlit/secrets.toml (see the README) and restart the app.")
        st.stop()
    try:
        try:
            from supabase import ClientOptions
        except ImportError:
            from supabase.lib.client_options import ClientOptions
        opts = {"auto_refresh_token": False}
        if implicit:
            opts["flow_type"] = "implicit"
        return create_client(url, key, options=ClientOptions(**opts))
    except Exception:
        return create_client(url, key)


def service_role_key() -> str:
    """Server-side only. Read from secrets, never shown in the UI or sent to the browser."""
    key = ""
    section = get_secret("supabase", None)
    if section:
        try:
            key = str(section.get("service_role_key", ""))
        except Exception:
            pass
    return (key or str(get_secret("SUPABASE_SERVICE_ROLE_KEY", ""))).strip()


def new_service_client():
    """Supabase Admin client (service_role). It bypasses Row Level Security, so it is created only
    by admin_client(), which first confirms the signed-in user is an admin."""
    url, _ = supabase_config()
    key = service_role_key()
    if not url or not key:
        raise RuntimeError("The service role key is not configured. Add service_role_key under "
                           "[supabase] in .streamlit/secrets.toml.")
    try:
        try:
            from supabase import ClientOptions
        except ImportError:
            from supabase.lib.client_options import ClientOptions
        return create_client(url, key, options=ClientOptions(auto_refresh_token=False, persist_session=False))
    except Exception:
        return create_client(url, key)


def sb():
    return st.session_state["sb"]


def uid() -> str:
    return st.session_state.get("uid", "")


def refresh_auth() -> bool:
    """Make sure the signed-in session is still valid (refreshes an expired token)
    and that database calls use the newest access token."""
    client = st.session_state.get("sb")
    if client is None:
        return False
    try:
        session = client.auth.get_session()
        if session is None:
            return False
        client.postgrest.auth(session.access_token)
        return True
    except Exception:
        return False


def run_failed(e: Exception):
    msg = str(e)
    if "jwt" in msg.lower() and "expired" in msg.lower():
        reset_session("Your session expired. Please sign in again.")
        st.rerun()
    st.error(f"Database error: {msg}")
    st.stop()


def run(query):
    """Execute a Supabase query. Shows a clear message instead of a stack trace on failure."""
    try:
        return query.execute()
    except Exception as e:
        run_failed(e)


# ─────────────────────────────────────────────────────────────
# Per-run cache. The script reruns after every click, and many helpers need the same
# tables, so each table is fetched at most once per run. Writes drop the cache entry.
# ─────────────────────────────────────────────────────────────
CACHE_TTL = 60   # seconds. Writes drop the entry at once; this only bounds staleness from other devices.


def _store() -> dict:
    return st.session_state.setdefault("_db", {})


def touch(*names):
    store = _store()
    for n in names:
        store.pop(n, None)


def cached(name: str, loader):
    """Return a table from the session cache, loading it when missing or older than CACHE_TTL.
    The returned lists are shared, so treat them as read-only."""
    store = _store()
    hit = store.get(name)
    if hit is None or time.time() - hit[0] > CACHE_TTL:
        hit = (time.time(), loader())
        store[name] = hit
    return hit[1]


PAGE = 1000   # PostgREST returns at most 1000 rows per request, so page through


def _fetch_table(client, user_id: str, table: str, order: str = "id") -> list:
    rows, start = [], 0
    while True:
        res = (client.table(table).select("*").eq("user_id", user_id).order(order)
               .range(start, start + PAGE - 1).execute())
        batch = res.data or []
        rows += batch
        if len(batch) < PAGE:
            return rows
        start += PAGE


def _fetch_profile(client, user_id: str):
    res = client.table("profiles").select("*").eq("id", user_id).limit(1).execute()
    return (res.data or [None])[0]


def fetch_all(table: str, order: str = "id") -> list:
    try:
        return _fetch_table(sb(), uid(), table, order)
    except Exception as e:
        run_failed(e)


def prefetch():
    """Load the profile, members, contributions and payments at the same time instead of one after
    another. On a slow connection this turns four round trips into roughly one."""
    store, now = _store(), time.time()
    jobs = {
        "profile": lambda c, u: _fetch_profile(c, u),
        "members": lambda c, u: _fetch_table(c, u, "members"),
        "contribs": lambda c, u: _fetch_table(c, u, "contributions"),
        "payments": lambda c, u: _fetch_table(c, u, "payments"),
    }
    need = [n for n in jobs if n not in store or now - store[n][0] > CACHE_TTL]
    if len(need) < 2:
        return
    client, user_id = sb(), uid()
    try:
        with ThreadPoolExecutor(max_workers=4) as pool:
            futures = {n: pool.submit(jobs[n], client, user_id) for n in need}
            results = {n: f.result() for n, f in futures.items()}
        for n, value in results.items():
            store[n] = (time.time(), value)
    except Exception:
        pass   # fall back to loading each table when it is first needed


# ─────────────────────────────────────────────────────────────
# Profile (group name, currency, default dues, lockout counters)
# ─────────────────────────────────────────────────────────────
def load_profile():
    def _load():
        res = run(sb().table("profiles").select("*").eq("id", uid()).limit(1))
        return (res.data or [None])[0]
    return cached("profile", _load)


def update_profile(fields: dict):
    run(sb().table("profiles").update(fields).eq("id", uid()))
    touch("profile")


def ensure_profile(group_hint: str = ""):
    """The database trigger normally creates the profile at sign up. This covers accounts
    that were created before the trigger existed."""
    if load_profile():
        return
    run(sb().table("profiles").insert({"id": uid(), "group_name": group_hint.strip() or "My Group"}))
    touch("profile")


def get_settings() -> dict:
    p = load_profile() or {}
    return {
        "group_name": (p.get("group_name") or DEFAULT_SETTINGS["group_name"]).strip(),
        "currency": p.get("currency") or DEFAULT_SETTINGS["currency"],
        "default_dues": to_float(p.get("default_dues")),
    }


# ─────────────────────────────────────────────────────────────
# Roles (RBAC). The role lives in profiles.role ("user" or "admin"). Users cannot change it:
# the SQL schema only lets the signed-in role write the other profile columns.
# ─────────────────────────────────────────────────────────────
ROLES = ("user", "admin")


def current_role() -> str:
    role = str((load_profile() or {}).get("role") or "user").lower()
    return role if role in ROLES else "user"


def is_admin() -> bool:
    """Cheap check (cached profile), used only to decide what to show in the navigation."""
    return current_role() == "admin"


def verify_admin_now() -> bool:
    """Authoritative check, repeated on every admin render and action. It re-reads the role from the
    database with the user's own token (RLS lets a user read only their own row), so a stale cache
    can never grant access."""
    try:
        prof = _fetch_profile(sb(), uid())
    except Exception:
        return False
    if prof:
        _store()["profile"] = (time.time(), prof)
    return bool(prof) and str(prof.get("role") or "").lower() == "admin"


def admin_client():
    if not verify_admin_now():
        raise PermissionError("Administrator access is required.")
    return new_service_client()


# ─────────────────────────────────────────────────────────────
# Authentication (Supabase Auth)
# ─────────────────────────────────────────────────────────────
def friendly_auth_error(e: Exception) -> str:
    msg = str(e)
    low = msg.lower()
    if "invalid login" in low or "invalid credentials" in low:
        return "That email or password is not right. Try again."
    if "not confirmed" in low:
        return "Please confirm your email first. Check your inbox for the confirmation link."
    if "rate limit" in low or "too many" in low:
        return "Too many attempts. Please wait a few minutes and try again."
    if "already registered" in low or "already exists" in low:
        return "An account with this email already exists. Try signing in."
    if "different from the old" in low or "same password" in low:
        return "Choose a password you have not used before."
    if "token has expired" in low or "otp_expired" in low or "token is invalid" in low:
        return "That reset link is invalid or has expired. Request a new one."
    if "password" in low and ("weak" in low or "least" in low or "short" in low):
        return f"Choose a stronger password (at least {MIN_PASSWORD} characters)."
    return msg or "Something went wrong. Please try again."


def reset_session(notice: str = ""):
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    if notice:
        st.session_state["_login_notice"] = notice


def do_logout():
    log_activity("Signed out", wait=True)
    client = st.session_state.get("sb")
    if client is not None:
        try:
            client.auth.sign_out({"scope": "local"})   # only this device, not other sessions
        except Exception:
            pass
    reset_session()


def finish_login(client, res):
    """Store the signed-in session and load (or create) the profile."""
    user = res.user
    st.session_state["sb"] = client
    st.session_state["uid"] = str(user.id)
    st.session_state["email"] = user.email or ""
    st.session_state["user"] = user.email or "User"
    st.session_state["auth"] = True
    st.session_state["_db"] = {}
    try:
        client.postgrest.auth(res.session.access_token)
    except Exception:
        pass
    meta = getattr(user, "user_metadata", None) or {}
    ensure_profile(str(meta.get("group_name", "")))
    log_activity("Signed in")
    st.rerun()


RESET_COOLDOWN = 60   # seconds between reset emails from one browser session


def _set_forgot(flag: bool):
    st.session_state["_forgot"] = flag


def clear_query_params():
    try:
        st.query_params.clear()
    except Exception:
        pass


def hash_bridge():
    """Supabase's default reset email sends the user back as /#access_token=...&type=recovery. A browser
    never sends the part after # to the server, so this tiny script (it runs in the browser) copies it into
    the query string, which Streamlit can read. It does nothing on any other page load."""
    js = """<script>
    try {
      var w = window.parent, h = w.location.hash || "";
      if ((h.indexOf("access_token=") > -1 && h.indexOf("type=recovery") > -1) || h.indexOf("error_code=") > -1) {
        w.location.replace(w.location.pathname + "?" + h.substring(1));
      }
    } catch (e) {}
    </script>"""
    if hasattr(st, "iframe"):
        st.iframe(js, height=1)
    else:
        components.html(js, height=0)


def consume_recovery_link():
    """Handles the reset link in either form: ?access_token=...&refresh_token=...&type=recovery (default
    Supabase email, via hash_bridge) or ?token_hash=...&type=recovery (custom email template). The tokens
    are used once and removed from the address bar straight away."""
    qp = st.query_params
    token_hash = str(qp.get("token_hash", "") or "")
    access = str(qp.get("access_token", "") or "")
    refresh = str(qp.get("refresh_token", "") or "")
    kind = str(qp.get("type", "") or "")
    err = str(qp.get("error_code", "") or qp.get("error_description", "") or "")
    if not (token_hash or access or err):
        return
    clear_query_params()
    if st.session_state.get("auth"):
        return
    if err:
        st.session_state["_login_notice"] = "That reset link is invalid or has expired. Request a new one."
        return
    if kind != "recovery":
        return
    client = new_client()
    try:
        if token_hash:
            res = client.auth.verify_otp({"token_hash": token_hash, "type": "recovery"})
        else:
            res = client.auth.set_session(access, refresh)
        st.session_state["_recovery"] = {"client": client, "email": getattr(res.user, "email", "") or ""}
    except Exception:
        st.session_state["_login_notice"] = "That reset link is invalid or has expired. Request a new one."


def render_recovery_card():
    """Shown after the user opens the link in the reset email: choose a new password."""
    rec = st.session_state["_recovery"]
    with st.container(key="login_card"):
        st.markdown('<div class="login-title">New password</div>'
                    f'<div class="login-sub">Choose a new password for {esc(rec.get("email", ""))}.</div>',
                    unsafe_allow_html=True)
        with st.form("recovery_form"):
            p1 = st.text_input("New password", type="password", placeholder=f"New password (min {MIN_PASSWORD} characters)",
                               label_visibility="collapsed", key="rec_pw1")
            p2 = st.text_input("Confirm new password", type="password", placeholder="Confirm new password",
                               label_visibility="collapsed", key="rec_pw2")
            go = st.form_submit_button("Update password", type="primary", **FULL_SUBMIT)
        if go:
            if len(p1) < MIN_PASSWORD:
                st.error(f"Use a password of at least {MIN_PASSWORD} characters.")
            elif p1 != p2:
                st.error("The two passwords do not match.")
            else:
                try:
                    rec["client"].auth.update_user({"password": p1})
                except Exception as e:
                    st.error(friendly_auth_error(e))
                else:
                    try:
                        rec["client"].auth.sign_out()
                    except Exception:
                        pass
                    reset_session("Password updated. Sign in with your new password.")
                    st.rerun()
        st.button("Cancel", key="rec_cancel", on_click=lambda: st.session_state.pop("_recovery", None), **FULL_BTN)


def forgot_password_form():
    st.markdown('<div class="menu-head">Reset your password</div>'
                '<div class="p-sub" style="margin-bottom:.5rem">Enter your email and we will send you a link '
                'to choose a new password.</div>', unsafe_allow_html=True)
    with st.form("forgot_form"):
        email = st.text_input("Email", placeholder="Email", label_visibility="collapsed", key="fp_email")
        go = st.form_submit_button("Send reset link", type="primary", **FULL_SUBMIT)
    if go:
        addr = email.strip()
        wait = RESET_COOLDOWN - (time.time() - st.session_state.get("_reset_sent_at", 0.0))
        if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", addr):
            st.error("Enter a valid email address.")
        elif wait > 0:
            st.warning(f"Please wait {int(wait) + 1} seconds before asking for another link.")
        else:
            try:
                new_client(implicit=True).auth.reset_password_for_email(addr)
            except Exception as e:
                st.error(friendly_auth_error(e))
            else:
                st.session_state["_reset_sent_at"] = time.time()
                # Same message whether or not the email has an account, so accounts cannot be guessed
                st.success("If an account exists for that email, a reset link is on its way. "
                           "Check your inbox and spam folder.")
    st.button("Back to sign in", key="forgot_back", on_click=_set_forgot, args=(False,), **FULL_BTN)


def sign_in_panel():
    if st.session_state.get("_forgot"):
        forgot_password_form()
        return
    with st.form("login_form"):
        email = st.text_input("Email", placeholder="Email", label_visibility="collapsed", key="login_email")
        pw = st.text_input("Password", type="password", placeholder="Password",
                           label_visibility="collapsed", key="login_pw")
        go_in = st.form_submit_button("Sign in", type="primary", **FULL_SUBMIT)
    if go_in:
        if not email.strip() or not pw:
            st.error("Enter your email and password.")
        else:
            client = new_client()
            try:
                res = client.auth.sign_in_with_password({"email": email.strip(), "password": pw})
            except Exception as e:
                st.error(friendly_auth_error(e))
            else:
                finish_login(client, res)
    st.button("Forgot password?", key="forgot_open", on_click=_set_forgot, args=(True,))


def require_login():
    consume_recovery_link()
    if st.session_state.get("auth"):
        if refresh_auth():
            return
        reset_session("Your session expired. Please sign in again.")

    st.markdown(BASE_CSS + LOGIN_CSS, unsafe_allow_html=True)
    hash_bridge()
    if st.session_state.get("_recovery"):
        render_recovery_card()
        st.markdown(copyright_html(), unsafe_allow_html=True)
        st.stop()
    with st.container(key="login_card"):
        st.markdown('<div class="login-title">Dues Tracker</div>'
                    '<div class="login-sub">Sign in, or create an account for your group.</div>',
                    unsafe_allow_html=True)
        notice = st.session_state.pop("_login_notice", "")
        if notice:
            st.info(notice)

        tab_in, tab_up = st.tabs(["Sign in", "Create account"])

        with tab_in:
            sign_in_panel()

        with tab_up:
            with st.form("signup_form"):
                group = st.text_input("Group name", placeholder="Group name (e.g. Class of 2026, Finance Dept)",
                                      label_visibility="collapsed", key="su_group", max_chars=80)
                email_u = st.text_input("Email", placeholder="Email", label_visibility="collapsed", key="su_email")
                pw1 = st.text_input("Password", type="password", placeholder=f"Password (min {MIN_PASSWORD} characters)",
                                    label_visibility="collapsed", key="su_pw1")
                pw2 = st.text_input("Confirm password", type="password", placeholder="Confirm password",
                                    label_visibility="collapsed", key="su_pw2")
                go_up = st.form_submit_button("Create account", type="primary", **FULL_SUBMIT)
            if go_up:
                if not group.strip():
                    st.error("Enter your group name.")
                elif not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email_u.strip()):
                    st.error("Enter a valid email address.")
                elif len(pw1) < MIN_PASSWORD:
                    st.error(f"Use a password of at least {MIN_PASSWORD} characters.")
                elif pw1 != pw2:
                    st.error("The two passwords do not match.")
                else:
                    client = new_client()
                    try:
                        # group_name travels as user metadata; a database trigger copies it into profiles
                        res = client.auth.sign_up({"email": email_u.strip(), "password": pw1,
                                                   "options": {"data": {"group_name": group.strip()}}})
                    except Exception as e:
                        st.error(friendly_auth_error(e))
                    else:
                        if getattr(res, "session", None):
                            finish_login(client, res)
                        else:
                            st.success("Account created. Check your email to confirm your address, "
                                       "then come back and sign in.")
    st.markdown(copyright_html(), unsafe_allow_html=True)
    st.stop()


# ─────────────────────────────────────────────────────────────
# Password check for protected actions, with a lockout after repeated mistakes
# Protected: Reset all data, Clear activity log, Delete payment records.
# The password is verified by signing in again with a throw-away client, so the main
# session is not touched. Counters live in the profiles table (pw_fails, pw_locked_until).
# ─────────────────────────────────────────────────────────────
def check_password(pw: str) -> bool:
    email = st.session_state.get("email", "")
    if not pw or not email:
        return False
    try:
        new_client().auth.sign_in_with_password({"email": email, "password": pw})
        return True
    except Exception:
        return False


def _fmt_wait(seconds: float) -> str:
    seconds = max(int(seconds + 0.999), 1)
    m, s = divmod(seconds, 60)
    return f"{m} min {s:02d} sec" if m else f"{s} sec"


def lockout_message() -> str:
    """Text to show while the signed-in user is locked out, or "" when they are not."""
    until = parse_ts((load_profile() or {}).get("pw_locked_until"))
    if until:
        wait = (until - now_utc()).total_seconds()
        if wait > 0:
            return (f"Too many wrong passwords. Protected actions are locked for {_fmt_wait(wait)}. "
                    "Try again after that.")
    return ""


def verify_protected(pw: str, action: str):
    """Check the password for a protected action. Returns (ok, message)."""
    blocked = lockout_message()
    if blocked:
        return False, blocked

    prof = load_profile() or {}
    fails = int(to_float(prof.get("pw_fails")))

    if check_password(pw):
        if fails or prof.get("pw_locked_until"):
            update_profile({"pw_fails": 0, "pw_locked_until": None})
        return True, ""

    fails += 1
    if fails >= MAX_ATTEMPTS:
        update_profile({"pw_fails": 0,
                        "pw_locked_until": (now_utc() + timedelta(seconds=LOCKOUT_SECONDS)).isoformat()})
        log_activity("Locked out after wrong passwords", action)
        return False, (f"Wrong password {MAX_ATTEMPTS} times. Protected actions are locked for "
                       f"{_fmt_wait(LOCKOUT_SECONDS)}.")
    update_profile({"pw_fails": fails, "pw_locked_until": None})
    log_activity("Wrong password", action)
    left = MAX_ATTEMPTS - fails
    return False, f"That password is not right. {left} attempt(s) left before a timeout."


# ─────────────────────────────────────────────────────────────
# Data access. Every query also filters on user_id; Row Level Security enforces the
# same rule inside Postgres, so a mistake here can never expose another account's rows.
# ─────────────────────────────────────────────────────────────
def load_members():
    return cached("members", lambda: fetch_all("members"))


def load_contribs():
    return cached("contribs", lambda: fetch_all("contributions"))


def load_payments():
    return cached("payments", lambda: fetch_all("payments"))


def load_activity():
    def _load():
        res = run(sb().table("activity").select("*").eq("user_id", uid()).order("ts", desc=True).limit(300))
        return [{"timestamp": fmt_ts(r.get("ts")), "user": r.get("actor", ""),
                 "action": r.get("action", ""), "details": r.get("details", "")} for r in (res.data or [])]
    return cached("activity", _load)


def _insert_activity(client, row: dict):
    try:
        client.table("activity").insert(row).execute()
    except Exception:
        pass   # the audit trail must never block or break the action the user just did


def log_activity(action: str, details: str = "", wait: bool = False):
    """Record an action. It is sent in the background so the screen does not wait for it."""
    row = {"user_id": uid(), "actor": st.session_state.get("user", ""), "action": action, "details": details}
    touch("activity")
    if wait:
        _insert_activity(sb(), row)
    else:
        threading.Thread(target=_insert_activity, args=(sb(), row), daemon=True).start()


# ─────────────────────────────────────────────────────────────
# Members
# ─────────────────────────────────────────────────────────────
def find_member(name: str):
    n = name.strip().lower()
    return next((m for m in load_members() if m.get("name", "").strip().lower() == n), None)


def add_member(name: str, contact: str = "", dues=None, notes: str = ""):
    name = name.strip()
    if not name:
        return False, "Enter a name."
    if find_member(name):
        return False, f"{name} is already in the group."
    run(sb().table("members").insert({"user_id": uid(), "name": name, "contact": contact.strip(),
                                      "notes": notes.strip(), "dues": dues}))
    touch("members")
    log_activity("Added member", name)
    return True, f"Added {name}."


def add_many(text: str):
    existing = {m.get("name", "").strip().lower() for m in load_members()}
    rows, skipped = [], 0
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        parts = [p.strip() for p in line.split(",", 1)]
        name = parts[0]
        if not name or name.lower() in existing:
            skipped += 1
            continue
        existing.add(name.lower())
        rows.append({"user_id": uid(), "name": name, "contact": parts[1] if len(parts) > 1 else "",
                     "notes": "", "dues": None})
    if rows:
        run(sb().table("members").insert(rows))
        touch("members")
        log_activity("Added members", f"{len(rows)} added")
    return len(rows), skipped


def update_member(member_id: int, name: str, contact: str, dues, notes: str):
    name = name.strip()
    if not name:
        return False, "Name cannot be empty."
    members = load_members()
    clash = next((m for m in members if mid(m) != member_id and m.get("name", "").strip().lower() == name.lower()), None)
    if clash:
        return False, f"Another member is already called {name}."
    old_rec = next((m for m in members if mid(m) == member_id), None)
    if old_rec is None:
        return False, "Member not found."
    old = old_rec.get("name", "")
    run(sb().table("members").update({"name": name, "contact": contact.strip(), "notes": notes.strip(),
                                      "dues": dues}).eq("id", member_id))
    run(sb().table("payments").update({"member_name": name}).eq("member_id", member_id))
    touch("members", "payments")
    log_activity("Edited member", f"{old} to {name}" if old != name else name)
    return True, f"Saved {name}."


def delete_member(member_id: int):
    target = next((m for m in load_members() if mid(m) == member_id), None)
    if not target:
        return
    # Payments are removed by the database (ON DELETE CASCADE). Participant lists are cleaned here.
    for c in load_contribs():
        ids = participant_ids(c)
        if ids is not None and member_id in ids:
            run(sb().table("contributions").update({"participants": sorted(ids - {member_id})}).eq("id", mid(c)))
    run(sb().table("members").delete().eq("id", member_id))
    touch("members", "contribs", "payments")
    log_activity("Deleted member", target.get("name", ""))


# ─────────────────────────────────────────────────────────────
# Contributions
# ─────────────────────────────────────────────────────────────
def contrib_name_clash(name: str, ignore_id: int = 0) -> str:
    n = name.strip().lower()
    for c in load_contribs():
        if mid(c) != ignore_id and c.get("name", "").strip().lower() == n:
            return f"A contribution called {name.strip()} already exists."
    return ""


def add_contribution(name: str, frequency: str = "one-off", amount=0.0, notes: str = ""):
    name = name.strip()
    if not name:
        return False, "Give the contribution a name.", None
    clash = contrib_name_clash(name)
    if clash:
        return False, clash, None
    res = run(sb().table("contributions").insert({
        "user_id": uid(), "name": name, "amount": float(amount or 0),
        "frequency": "monthly" if frequency == "monthly" else "one-off",
        "member_amounts": False, "notes": notes.strip()}))
    touch("contribs")
    log_activity("Added contribution", name)
    new_id = mid((res.data or [{}])[0])
    return True, f"Added {name}.", new_id


def update_contribution(cid: int, name: str, amount, notes: str):
    name = name.strip()
    if not name:
        return False, "Name cannot be empty."
    clash = contrib_name_clash(name, ignore_id=cid)
    if clash:
        return False, clash
    c = next((x for x in load_contribs() if mid(x) == cid), None)
    if c is None:
        return False, "Contribution not found."
    fields = {"name": name, "notes": notes.strip()}
    if not c.get("member_amounts"):
        fields["amount"] = float(amount or 0)
    run(sb().table("contributions").update(fields).eq("id", cid))
    touch("contribs")
    old = c.get("name", "")
    log_activity("Edited contribution", f"{old} to {name}" if old != name else name)
    return True, f"Saved {name}."


def delete_contribution(cid: int):
    target = next((c for c in load_contribs() if mid(c) == cid), None)
    if not target:
        return
    run(sb().table("contributions").delete().eq("id", cid))   # its payments cascade
    touch("contribs", "payments")
    log_activity("Removed contribution", target.get("name", ""))


def participant_ids(c: dict):
    """Ids of the members who take part in a contribution, or None when everyone in the group does."""
    v = c.get("participants")
    return None if v is None else {mid(x) for x in v}


def roster(c: dict, members: list, period=None) -> list:
    """Members to show for a contribution: its participants, plus anyone who already has a payment
    record for it (for the given period, or any period when period is None), so history is never hidden."""
    ids = participant_ids(c)
    if ids is None:
        return list(members)
    cid = mid(c)
    have = {mid(p.get("member_id")) for p in load_payments()
            if mid(p.get("contribution_id")) == cid and (period is None or p.get("period", "") == period)}
    return [m for m in members if mid(m) in ids or mid(m) in have]


def join_contribution(cid: int, member_ids: list):
    members = load_members()
    for c in load_contribs():
        if mid(c) == cid:
            cur = participant_ids(c)
            if cur is None:          # everyone is already in
                return
            new = {mid(x) for x in member_ids}
            run(sb().table("contributions").update({"participants": sorted(cur | new)}).eq("id", cid))
            touch("contribs")
            names = ", ".join(x.get("name", "") for x in members if mid(x) in new)
            log_activity("Added to contribution", f"{names} | {c.get('name')}")
            return


def leave_contribution(cid: int, member_ids: list, purge: bool = False):
    members = load_members()
    gone = {mid(x) for x in member_ids}
    for c in load_contribs():
        if mid(c) == cid:
            cur = participant_ids(c)
            if cur is None:
                cur = {mid(x) for x in members}
            run(sb().table("contributions").update({"participants": sorted(cur - gone)}).eq("id", cid))
            if purge:
                run(sb().table("payments").delete().eq("contribution_id", cid).in_("member_id", sorted(gone)))
            touch("contribs", "payments")
            names = ", ".join(x.get("name", "") for x in members if mid(x) in gone)
            log_activity("Removed from contribution",
                         f"{names} | {c.get('name')}" + (" | records deleted" if purge else ""))
            return


def dues_of(m: dict) -> float:
    v = m.get("dues")
    try:
        if v is None or v == "" or pd.isna(v):
            return float(SETTINGS["default_dues"])
        return float(v)
    except Exception:
        return float(SETTINGS["default_dues"])


def due_for(c: dict, m: dict) -> float:
    """Amount a member is expected to give for this contribution (per period)."""
    if c.get("member_amounts"):
        return dues_of(m)
    return to_float(c.get("amount"))


def money(x) -> str:
    return f"{SETTINGS['currency']} {float(x or 0):,.2f}"


# ─────────────────────────────────────────────────────────────
# Payments (all contributions)
# period is "YYYY-MM" for monthly contributions and "" for one-off ones
# ─────────────────────────────────────────────────────────────
PAY_CONFLICT = "contribution_id,member_id,period"


def pay_map_for(cid: int, period: str) -> dict:
    return {mid(p.get("member_id")): p for p in load_payments()
            if mid(p.get("contribution_id")) == cid and p.get("period", "") == period}


def status_of(m: dict, pay_map: dict) -> str:
    p = pay_map.get(mid(m))
    return norm_status(p.get("status")) if p else "unpaid"


def counted(rec, status: str) -> float:
    """Amount that counts as money received (paid or partial only)."""
    return to_float(rec.get("amount")) if rec and status in ("paid", "partial") else 0.0


def _pay_rec(c: dict, member: dict, period: str, status: str, amount: float, notes: str = "") -> dict:
    return {
        "user_id": uid(),
        "contribution_id": mid(c),
        "member_id": mid(member),
        "member_name": member.get("name", ""),
        "period": period,
        "amount": float(amount or 0),
        "status": status,
        "notes": notes.strip(),
        "recorded_at": now_iso(),
        "recorded_by": st.session_state.get("user", ""),
    }


def set_payment(c: dict, member: dict, period: str, status: str, amount: float, notes: str = ""):
    rec = _pay_rec(c, member, period, status, amount, notes)
    run(sb().table("payments").upsert(rec, on_conflict=PAY_CONFLICT))
    touch("payments")
    when = f" | {period}" if period else ""
    log_activity("Payment", f"{c.get('name')}{when} | {member.get('name')} | {status} | {rec['amount']:.2f}")


def set_payments_bulk(c: dict, period: str, pairs: list) -> int:
    """Mark many members paid in one request. pairs is a list of (member, amount)."""
    if not pairs:
        return 0
    recs = [_pay_rec(c, m, period, "paid", amt) for m, amt in pairs]
    run(sb().table("payments").upsert(recs, on_conflict=PAY_CONFLICT))
    touch("payments")
    when = f" | {period}" if period else ""
    log_activity("Payment", f"{c.get('name')}{when} | marked {len(recs)} member(s) as paid")
    return len(recs)


def clear_payment(c: dict, member: dict, period: str):
    run(sb().table("payments").delete().eq("contribution_id", mid(c)).eq("member_id", mid(member))
        .eq("period", period))
    touch("payments")
    when = f" | {period}" if period else ""
    log_activity("Cleared payment", f"{c.get('name')}{when} | {member.get('name')}")


def delete_payments(ids: list):
    ids = [int(i) for i in ids]
    for i in range(0, len(ids), 100):
        run(sb().table("payments").delete().in_("id", ids[i:i + 100]))
    touch("payments")
    log_activity("Deleted payment records", f"{len(ids)} removed")


def clear_activity():
    run(sb().table("activity").delete().eq("user_id", uid()))
    touch("activity")


def reset_all_data():
    for table in ("payments", "contributions", "members", "activity"):
        run(sb().table(table).delete().eq("user_id", uid()))
    touch("payments", "contribs", "members", "activity")


def save_settings(name: str, currency: str, dues: float):
    update_profile({"group_name": name.strip() or DEFAULT_SETTINGS["group_name"],
                    "currency": currency.strip() or DEFAULT_SETTINGS["currency"],
                    "default_dues": float(dues or 0)})


def contrib_stats(c: dict, members: list, period: str) -> dict:
    pm = pay_map_for(mid(c), period)
    sts = {mid(m): status_of(m, pm) for m in members}
    paid_n = sum(1 for s in sts.values() if s == "paid")
    waived_n = sum(1 for s in sts.values() if s == "waived")
    collected = sum(counted(pm.get(mid(m)), sts[mid(m)]) for m in members)
    expected = sum(due_for(c, m) for m in members if sts[mid(m)] != "waived")
    return {
        "pay_map": pm,
        "statuses": sts,
        "paid_n": paid_n,
        "waived_n": waived_n,
        "unpaid_n": len(members) - paid_n - waived_n,
        "countable": max(len(members) - waived_n, 1),
        "collected": collected,
        "expected": expected,
    }


# ─────────────────────────────────────────────────────────────
# Reports (Excel and Word) with member / month / contribution filters
# Every export starts with the Group Name.
# Monthly contributions follow the month filter (or every month with records).
# One-off contributions are not tied to a month, so they always show everything recorded.
# ─────────────────────────────────────────────────────────────
def export_months(extra: str = "") -> list:
    months = set()
    by_id = {mid(c): c for c in load_contribs()}
    for p in load_payments():
        c = by_id.get(mid(p.get("contribution_id")))
        if not c:
            continue
        bucket = p.get("period", "") if is_monthly(c) else str(p.get("recorded_at", ""))[:7]
        if bucket:
            months.add(bucket)
    if extra:
        months.add(extra)
    return sorted(months, reverse=True)


def build_report(member_id=None, month=None, contrib_id=None) -> dict:
    """Everything a report needs. None means "all" for the member, the month and the contribution."""
    members = sorted(load_members(), key=lambda x: x.get("name", "").lower())
    contribs = sorted(load_contribs(), key=mid)
    pays = load_payments()
    member = next((m for m in members if mid(m) == member_id), None) if member_id is not None else None
    contrib = next((c for c in contribs if mid(c) == contrib_id), None) if contrib_id is not None else None
    chosen = [c for c in contribs if contrib_id is None or mid(c) == contrib_id]

    lines, by_contrib, breakdowns = [], [], []
    for c in chosen:
        cid = mid(c)
        monthly = is_monthly(c)
        people = roster(c, members)
        if member_id is not None:
            people = [m for m in people if mid(m) == member_id]
        if monthly:
            periods = [month] if month else sorted({p.get("period") for p in pays
                                                    if mid(p.get("contribution_id")) == cid and p.get("period")})
        else:
            periods = [""]

        c_lines = []
        for per in periods:
            pm = pay_map_for(cid, per)
            for m in people:
                s = status_of(m, pm)
                rec = pm.get(mid(m))
                due = due_for(c, m)
                paid = counted(rec, s)
                expected = 0.0 if s == "waived" else due
                balance = 0.0 if s in ("paid", "waived") else max(expected - paid, 0.0)
                c_lines.append({
                    "Period": month_label(per) if per else "One-off",
                    "Contribution": c.get("name", ""),
                    "Member": m.get("name", ""),
                    "Status": STATUS[s][0],
                    "Due": expected,
                    "Paid": paid,
                    "Balance": balance,
                    "Recorded": fmt_ts(rec.get("recorded_at")) if rec else "",
                    "Recorded by": rec.get("recorded_by", "") if rec else "",
                    "Notes": rec.get("notes", "") if rec else "",
                    "_status": s, "_period": per, "_mid": mid(m), "_contact": m.get("contact", ""),
                })
        lines += c_lines

        if not monthly:
            period_text = "One-off"
        elif month:
            period_text = month_label(month)
        else:
            period_text = f"{len(periods)} month(s)" if periods else "No months yet"
        by_contrib.append({
            "Contribution": c.get("name", ""),
            "Type": "Monthly" if monthly else "One-off",
            "Period": period_text,
            "Members": len({x["_mid"] for x in c_lines}),
            "Expected": sum(x["Due"] for x in c_lines),
            "Collected": sum(x["Paid"] for x in c_lines),
            "Balance": sum(x["Balance"] for x in c_lines),
        })

        # Month-by-month grid for a monthly contribution when every month is included
        if monthly and not month and len(periods) > 1 and people:
            cols = ["Name"] + [month_short(x) for x in periods] + ["Total"]
            rows = []
            for m in people:
                row, tot = {"Name": m.get("name", "")}, 0.0
                for per in periods:
                    a = sum(x["Paid"] for x in c_lines if x["_mid"] == mid(m) and x["_period"] == per)
                    row[month_short(per)] = a
                    tot += a
                row["Total"] = tot
                rows.append(row)
            totals = {"Name": "TOTAL"}
            for col in cols[1:]:
                totals[col] = sum(r[col] for r in rows)
            rows.append(totals)
            breakdowns.append({"name": c.get("name", ""), "columns": cols, "rows": rows})

    per_member = {}
    for x in lines:
        d = per_member.setdefault(x["_mid"], {"Member": x["Member"], "Contact": x["_contact"], "Contributions": set(),
                                              "Expected": 0.0, "Collected": 0.0, "Balance": 0.0})
        d["Contributions"].add(x["Contribution"])
        d["Expected"] += x["Due"]
        d["Collected"] += x["Paid"]
        d["Balance"] += x["Balance"]
    by_member = sorted(({**d, "Contributions": len(d["Contributions"])} for d in per_member.values()),
                       key=lambda d: d["Member"].lower())

    if member and contrib:
        title = f"{member.get('name', '')}: {contrib.get('name', '')}"
    elif member:
        title = f"Member statement: {member.get('name', '')}"
    elif contrib:
        title = f"Contribution report: {contrib.get('name', '')}"
    else:
        title = "Contributions report"

    return {
        "title": title,
        "filters": [("Member", member.get("name", "") if member else "All members"),
                    ("Month", month_label(month) if month else "All months"),
                    ("Contribution", contrib.get("name", "") if contrib else "All contributions")],
        "lines": lines, "by_contrib": by_contrib, "by_member": by_member, "breakdowns": breakdowns,
        "expected": sum(x["Due"] for x in lines),
        "collected": sum(x["Paid"] for x in lines),
        "balance": sum(x["Balance"] for x in lines),
        "has_one_off": any(not is_monthly(c) for c in chosen),
        # Summary tables only make sense across several contributions / members
        "show_by_contrib": contrib_id is None and len(chosen) > 1,
        "show_by_member": contrib_id is None and member_id is None and len(chosen) > 1,
    }


# ─────────────────────────────────────────────────────────────
# Excel and Word writers
# ─────────────────────────────────────────────────────────────
SHEET_HEADER_ROW = 4        # rows 1-2 hold the Group Name banner, row 3 is blank, row 4 the column headings


def style_sheet(ws, df, status_cols=(), money_cols=(), total_row=False, header_row=1):
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    head_fill = PatternFill("solid", fgColor="2F5BFF")
    thin = Side(style="thin", color="D9DCEF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for cell in ws[header_row]:
        cell.fill = head_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[header_row].height = 32

    for idx, col in enumerate(df.columns, start=1):
        longest = max([len(str(col))] + [len(str(v)) for v in df[col].tolist()]) if len(df) else len(str(col))
        ws.column_dimensions[get_column_letter(idx)].width = min(max(12, longest + 2), 42)
        for row in range(header_row + 1, ws.max_row + 1):
            cell = ws.cell(row=row, column=idx)
            cell.border = border
            if col in money_cols:
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right")
            if col in status_cols:
                cell.alignment = Alignment(horizontal="center")
                if cell.value in STATUS_FILL:
                    cell.fill = PatternFill("solid", fgColor=STATUS_FILL[cell.value])

    if total_row and ws.max_row > header_row:
        for cell in ws[ws.max_row]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="E8EBFA")
    ws.freeze_panes = f"B{header_row + 1}"


def banner(ws, subtitle: str):
    """Group Name in big letters on row 1 and what the sheet shows on row 2."""
    from openpyxl.styles import Font
    ws["A1"] = SETTINGS["group_name"]
    ws["A1"].font = Font(bold=True, size=16, color="2F5BFF")
    ws["A2"] = subtitle
    ws["A2"].font = Font(italic=True, size=10, color="555555")
    ws.row_dimensions[1].height = 26


def safe_sheet_name(name: str, used: set) -> str:
    base = re.sub(r"[\[\]\:\*\?\/\\]", "", name).strip()[:24] or "Sheet"
    candidate = f"{base} by month"
    n = 2
    while candidate in used:
        candidate = f"{base} by month {n}"[:31]
        n += 1
    used.add(candidate)
    return candidate


LEDGER_COLS = ["Period", "Contribution", "Member", "Status", "Due", "Paid", "Balance", "Recorded", "Recorded by", "Notes"]


def create_excel_report(member_id=None, month=None, contrib_id=None, user: str = "") -> bytes:
    data = build_report(member_id, month, contrib_id)
    cur = SETTINGS["currency"]
    out = io.BytesIO()
    used = {"Report", "By contribution", "By member", "Ledger"}

    def write(writer, df, sheet, label, **style):
        df.to_excel(writer, sheet_name=sheet, index=False, startrow=SHEET_HEADER_ROW - 1)
        ws = writer.sheets[sheet]
        banner(ws, f"{data['title']}  |  {label}")
        style_sheet(ws, df, header_row=SHEET_HEADER_ROW, **style)
        return ws

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        info = [("Group", SETTINGS["group_name"]), ("Report", data["title"])] + list(data["filters"]) + [
            ("Expected", f"{cur} {data['expected']:,.2f}"),
            ("Collected", f"{cur} {data['collected']:,.2f}"),
            ("Outstanding", f"{cur} {data['balance']:,.2f}"),
            ("Generated", datetime.now().strftime("%d %B %Y, %H:%M")),
            ("Generated by", user or "-"),
            ("Developed by", "AppaTech"),
        ]
        idf = pd.DataFrame([{"Item": k, "Value": v} for k, v in info])
        ws = write(writer, idf, "Report", "Summary")
        ws.column_dimensions["B"].width = 44

        if data["show_by_contrib"] and data["by_contrib"]:
            cdf = pd.DataFrame(data["by_contrib"])
            cdf.loc[len(cdf)] = ["TOTAL", "", "", int(cdf["Members"].sum()), cdf["Expected"].sum(),
                                 cdf["Collected"].sum(), cdf["Balance"].sum()]
            write(writer, cdf, "By contribution", "By contribution",
                  money_cols=("Expected", "Collected", "Balance"), total_row=True)

        if data["show_by_member"] and data["by_member"]:
            mdf = pd.DataFrame(data["by_member"])
            mdf.loc[len(mdf)] = ["TOTAL", "", "", mdf["Expected"].sum(), mdf["Collected"].sum(), mdf["Balance"].sum()]
            mdf = mdf[["Member", "Contact", "Contributions", "Expected", "Collected", "Balance"]]
            write(writer, mdf, "By member", "By member",
                  money_cols=("Expected", "Collected", "Balance"), total_row=True)

        if data["lines"]:
            ldf = pd.DataFrame(data["lines"])[LEDGER_COLS]
            ldf.loc[len(ldf)] = ["TOTAL", "", "", "", ldf["Due"].sum(), ldf["Paid"].sum(), ldf["Balance"].sum(), "", "", ""]
            write(writer, ldf, "Ledger", "Ledger", status_cols=("Status",),
                  money_cols=("Due", "Paid", "Balance"), total_row=True)

        for b in data["breakdowns"]:
            bdf = pd.DataFrame(b["rows"], columns=b["columns"])
            name = safe_sheet_name(b["name"], used)
            write(writer, bdf, name, f"{b['name']} by month", money_cols=b["columns"][1:], total_row=True)
    return out.getvalue()


def create_word_report(member_id=None, month=None, contrib_id=None, user: str = "") -> bytes:
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.section import WD_ORIENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    RIGHT, CENTER, LEFT = WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT
    GREY = RGBColor(0x55, 0x55, 0x55)

    def shade(cell, hex_color):
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), hex_color)
        cell._tc.get_or_add_tcPr().append(shd)

    def put(cell, text, bold=False, color=None, size=9, align=None):
        cell.text = ""
        para = cell.paragraphs[0]
        run_ = para.add_run(str(text))
        run_.font.size = Pt(size)
        run_.bold = bold
        if color is not None:
            run_.font.color.rgb = color
        if align is not None:
            para.alignment = align

    def header_row(table, headers, size=9):
        for i, h in enumerate(headers):
            c = table.rows[0].cells[i]
            put(c, h, bold=True, color=RGBColor(255, 255, 255), size=size, align=CENTER)
            shade(c, "2F5BFF")

    def new_table(headers, size=9):
        t = doc.add_table(rows=1, cols=len(headers))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        header_row(t, headers, size)
        return t

    def money_text(v):
        return f"{v:,.2f}"

    data = build_report(member_id, month, contrib_id)
    cur = SETTINGS["currency"]

    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Cm(1.5))

    # Group Name first, then what the report is
    doc.add_heading(SETTINGS["group_name"], level=0).alignment = CENTER
    p = doc.add_paragraph()
    p.alignment = CENTER
    r = p.add_run(data["title"])
    r.bold = True
    r.font.size = Pt(14)
    for text in ("   |   ".join(f"{k}: {v}" for k, v in data["filters"]),
                 f"Generated {datetime.now().strftime('%d %B %Y, %H:%M')}" + (f" by {user}" if user else "")):
        p = doc.add_paragraph()
        p.alignment = CENTER
        r = p.add_run(text)
        r.font.size = Pt(10)
        r.font.color.rgb = GREY

    # Summary figures
    s = doc.add_paragraph()
    s.add_run("Summary\n").bold = True
    s.add_run(f"Expected: {cur} {money_text(data['expected'])}\n")
    s.add_run(f"Collected: {cur} {money_text(data['collected'])}\n")
    s.add_run(f"Outstanding: {cur} {money_text(data['balance'])}").bold = True

    if not data["lines"]:
        doc.add_paragraph("There are no records for this selection.")
    else:
        if data["show_by_contrib"]:
            # By contribution
            doc.add_heading("By contribution", level=2)
            t = new_table(["Contribution", "Type", "Period", "Members", "Expected", "Collected", "Balance"])
            for o in data["by_contrib"]:
                c = t.add_row().cells
                put(c[0], o["Contribution"], bold=True)
                put(c[1], o["Type"])
                put(c[2], o["Period"])
                put(c[3], o["Members"], align=RIGHT)
                put(c[4], money_text(o["Expected"]) if o["Expected"] else "-", align=RIGHT)
                put(c[5], money_text(o["Collected"]), align=RIGHT)
                put(c[6], money_text(o["Balance"]) if o["Expected"] else "-", align=RIGHT)

        # By member (skipped for a single-member statement)
        if data["show_by_member"]:
            doc.add_heading("By member", level=2)
            t = new_table(["Member", "Contributions", "Expected", "Collected", "Balance"])
            for o in data["by_member"]:
                c = t.add_row().cells
                put(c[0], o["Member"], bold=True)
                put(c[1], o["Contributions"], align=RIGHT)
                put(c[2], money_text(o["Expected"]) if o["Expected"] else "-", align=RIGHT)
                put(c[3], money_text(o["Collected"]), align=RIGHT)
                put(c[4], money_text(o["Balance"]) if o["Expected"] else "-", align=RIGHT)

        # Ledger
        doc.add_heading("Details", level=2)
        spec = [("Period", "Period", LEFT), ("Contribution", "Contribution", LEFT), ("Member", "Member", LEFT),
                ("Status", "Status", CENTER), ("Due", "Due", RIGHT), ("Paid", "Paid", RIGHT),
                ("Balance", "Balance", RIGHT), ("Recorded by", "Recorded by", LEFT)]
        if member_id is not None:
            spec = [x for x in spec if x[0] != "Member"]
        if contrib_id is not None:
            spec = [x for x in spec if x[0] != "Contribution"]
        t = new_table([h for h, _, _ in spec])
        for x in data["lines"]:
            cells = t.add_row().cells
            for i, (h, key, al) in enumerate(spec):
                v = x[key]
                if key in ("Due", "Paid", "Balance"):
                    v = money_text(v) if v else "-"
                put(cells[i], v, bold=(key == "Member"), align=al)
                if key == "Status":
                    shade(cells[i], STATUS_FILL[x["Status"]])
        cells = t.add_row().cells
        for cell in cells:
            shade(cell, "E8EBFA")
        put(cells[0], "TOTAL", bold=True)
        for i, (h, key, al) in enumerate(spec):
            if key in ("Due", "Paid", "Balance"):
                put(cells[i], money_text(data["expected"] if key == "Due" else
                                         data["collected"] if key == "Paid" else data["balance"]),
                    bold=True, align=RIGHT)

        # Month-by-month grids
        for b in data["breakdowns"]:
            doc.add_heading(f"{b['name']} by month", level=2)
            tb = new_table(b["columns"], size=8)
            for ri, row in enumerate(b["rows"]):
                is_total = ri == len(b["rows"]) - 1
                cells = tb.add_row().cells
                for ci, col in enumerate(b["columns"]):
                    v = row[col]
                    if ci == 0:
                        put(cells[ci], v, bold=True, size=8)
                    else:
                        put(cells[ci], money_text(v) if v else "-", bold=is_total or col == "Total", size=8, align=RIGHT)
                    if is_total:
                        shade(cells[ci], "E8EBFA")

        legend = doc.add_paragraph()
        note = "Shading: green = paid, orange = partial, red = unpaid, purple = waived. "
        if data["has_one_off"] and month:
            note += "One-off contributions are not tied to a month, so they show everything recorded. "
        lr = legend.add_run(note + f"Amounts are in {cur}.")
        lr.font.size = Pt(8)
        lr.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    foot = sec.footer.paragraphs[0]
    foot.alignment = CENTER
    fr = foot.add_run(f"{SETTINGS['group_name']}  |  © {datetime.now().year} AppaTech. All rights reserved.")
    fr.font.size = Pt(8)
    fr.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────
# UI building blocks
# ─────────────────────────────────────────────────────────────
def page_header(title: str, sub: str = ""):
    st.markdown(f'<div class="page-title">{esc(title)}</div><div class="page-sub">{esc(sub)}</div>',
                unsafe_allow_html=True)


def group_banner():
    """The signed-in user's Group Name, shown at the top of the dashboard."""
    st.markdown(f'<div class="group-banner"><div class="gb-label">Group</div>'
                f'<div class="gb-name">{esc(SETTINGS["group_name"])}</div></div>', unsafe_allow_html=True)


def tile(label, value, sub, bg, fg="#fff"):
    return (f'<div class="tile" style="background:{bg};color:{fg}">'
            f'<div class="tile-label">{esc(label)}</div>'
            f'<div class="tile-value">{esc(str(value))}</div>'
            f'<div class="tile-sub">{esc(sub)}</div></div>')


def tiles_html(items: list) -> str:
    return '<div class="tiles">' + "".join(tile(*t) for t in items) + "</div>"


def person_html(m: dict) -> str:
    color = AVATARS[mid(m) % len(AVATARS)]
    contact = m.get("contact") or "No contact saved"
    return (f'<div class="person"><div class="avatar" style="background:{color}">{esc(initials(m.get("name", "")))}</div>'
            f'<div><div class="p-name">{esc(m.get("name", ""))}</div>'
            f'<div class="p-sub ell">{esc(contact)}</div></div></div>')


def chip_html(status: str) -> str:
    label, color = STATUS[status]
    return f'<span class="chip" style="background:{color}">{label}</span>'


def amount_line_html(status: str, p, due: float) -> str:
    """Caption on the left, amount on the right, one aligned line under the member."""
    if status in ("paid", "partial"):
        left = f"Given, of {money(due)}" if due else "Given"
        right = money(p.get("amount", 0) if p else 0)
    elif status == "waived":
        left, right = "Not required", ""
    elif due > 0:
        left, right = "Due", money(due)
    else:
        left, right = "Nothing recorded", ""
    return (f'<div class="mamt"><span class="p-sub">{esc(left)}</span>'
            f'<span class="amt">{esc(right)}</span></div>')


def card_title(title: str, sub: str = "") -> str:
    return (f'<div class="ctitle">{esc(title)}</div>'
            + (f'<div class="csub">{esc(sub)}</div>' if sub else '<div class="csub"></div>'))


def go_to(page: str):
    st.session_state["page"] = page
    st.session_state["menu_view"] = None


def progress_html(pct: int, text: str) -> str:
    return (f'<div class="bar"><div class="bar-fill" style="width:{pct}%"></div></div>'
            f'<div class="bar-label">{esc(text)}</div>')


def pick_one(label: str, options: list, key: str, default=None):
    """Pills where available (they wrap on small screens), radio buttons on older Streamlit."""
    if hasattr(st, "pills"):
        kw = {"default": default} if default is not None else {}
        return st.pills(label, options, selection_mode="single", label_visibility="collapsed", key=key, **kw)
    return st.radio(label, options, horizontal=True, label_visibility="collapsed", key=key)


def data_signature() -> str:
    """Changes whenever the user's data or settings change, so cached exports are rebuilt only when needed."""
    blob = json.dumps([load_members(), load_contribs(), load_payments(), SETTINGS], sort_keys=True, default=str)
    return hashlib.md5(blob.encode("utf-8")).hexdigest()


# The cache key includes the user id, so one account can never be served another account's export.
@st.cache_data(show_spinner=False, max_entries=16)
def cached_excel(member_id, month, contrib_id, user_id, user, sig) -> bytes:
    return create_excel_report(member_id, month, contrib_id, user)


@st.cache_data(show_spinner=False, max_entries=16)
def cached_word(member_id, month, contrib_id, user_id, user, sig) -> bytes:
    return create_word_report(member_id, month, contrib_id, user)



# ─────────────────────────────────────────────────────────────
# Confirmation pop-ups for removing people
# Uses a real pop-up window (st.dialog) when Streamlit has it, otherwise shows inline.
# ─────────────────────────────────────────────────────────────
_dialog = getattr(st, "dialog", None) or getattr(st, "experimental_dialog", None)


def _popup(title: str):
    def wrap(fn):
        return _dialog(title)(fn) if _dialog is not None else fn
    return wrap


@_popup("Delete member?")
def confirm_delete_member(member_id: int, name: str, count: int):
    st.warning(f"You are about to delete **{name}**. This also removes {count} payment record(s) "
               f"and takes them out of every contribution. This cannot be undone.")
    c1, c2 = st.columns(2)
    if c1.button("Yes, delete", key=f"del_dlg_{member_id}", **FULL_BTN):
        delete_member(member_id)
        flash(f"Deleted {name}.")
        st.rerun()
    if c2.button("Cancel", key=f"dlg_cancel_{member_id}", **FULL_BTN):
        st.rerun()


@_popup("Remove member?")
def confirm_leave_member(cid: int, member_id: int, member_name: str, contrib_name: str, purge: bool):
    extra = (" Their payment records for this contribution will also be deleted." if purge
             else " Their past payments will stay in the history.")
    st.warning(f"Remove **{member_name}** from **{contrib_name}**?{extra}")
    c1, c2 = st.columns(2)
    if c1.button("Yes, remove", key=f"del_dlgl_{cid}_{member_id}", **FULL_BTN):
        leave_contribution(cid, [member_id], purge)
        flash(f"Removed {member_name} from {contrib_name}.")
        _menu_done()
        st.rerun()
    if c2.button("Cancel", key=f"dlg_cancell_{cid}_{member_id}", **FULL_BTN):
        st.rerun()


# ─────────────────────────────────────────────────────────────
# Top navigation (replaces the sidebar)
# ─────────────────────────────────────────────────────────────
PAGES = [("dashboard", "Home"), ("members", "Members"), ("reports", "Reports"),
         ("history", "History"), ("settings", "Settings")]


def nav_pages() -> list:
    pages = list(PAGES)
    if is_admin():
        pages.append(("admin", "Admin"))
    return pages


def render_nav():
    """Rendered at the top of every page, before the page content."""
    pages = nav_pages()
    with st.container(key="tnav"):
        cols = st.columns(len(pages))
        for col, (key, label) in zip(cols, pages):
            active = st.session_state["page"] == key
            col.button(label, key=f"nav_{key}", type="primary" if active else "secondary",
                       on_click=go_to, args=(key,), **FULL_BTN)


# ─────────────────────────────────────────────────────────────
# Charts
# ─────────────────────────────────────────────────────────────
def chart_empty(title: str, text: str):
    st.markdown(f'<div class="empty" style="padding:1.2rem"><b>{esc(title)}</b><br>{esc(text)}</div>',
                unsafe_allow_html=True)


def chart_collections(month: str, contribs: list, span: int = 6):
    months = [shift_month(month, -k) for k in range(span - 1, -1, -1)]
    by_id = {mid(c): c for c in contribs}
    amounts = {}
    for p in load_payments():
        c = by_id.get(mid(p.get("contribution_id")))
        if not c or norm_status(p.get("status")) not in ("paid", "partial"):
            continue
        bucket = p.get("period", "") if is_monthly(c) else str(p.get("recorded_at", ""))[:7]
        if bucket in months:
            key = (bucket, c["name"])
            amounts[key] = amounts.get(key, 0.0) + to_float(p.get("amount"))

    present = [c for c in contribs if any(amounts.get((m, c["name"]), 0) for m in months)]
    if not present:
        chart_empty("No payments yet", "Collections will appear here once payments are recorded.")
        return

    order = [month_short(m) for m in months]
    df = pd.DataFrame([{"Month": month_short(m), "Contribution": c["name"],
                        "Amount": amounts.get((m, c["name"]), 0.0)} for m in months for c in present])
    chart = (
        alt.Chart(df)
        .mark_bar(cornerRadiusTopLeft=7, cornerRadiusTopRight=7)
        .encode(
            x=alt.X("Month:N", sort=order, axis=alt.Axis(labelAngle=0, title=None, labelOverlap=True)),
            y=alt.Y("Amount:Q", title=None),
            color=alt.Color("Contribution:N",
                            scale=alt.Scale(domain=[c["name"] for c in present], range=[color_of(c) for c in present]),
                            legend=alt.Legend(orient="bottom", title=None, columns=2)),
            tooltip=["Month", "Contribution", alt.Tooltip("Amount:Q", format=",.2f")],
        )
        .properties(height=230)
        .configure_view(strokeWidth=0)
    )
    st.altair_chart(chart, **FULL_CHART)


def chart_status_donut(statuses: dict):
    counts = {k: 0 for k in STATUS}
    for s in statuses.values():
        counts[s] = counts.get(s, 0) + 1
    total = sum(counts.values()) or 1
    stops, start = [], 0.0
    for k, (_, color) in STATUS.items():
        if counts[k]:
            end = start + 100.0 * counts[k] / total
            stops.append(f"{color} {start:.3f}% {end:.3f}%")
            start = end
    gradient = ", ".join(stops) or "#DADDF0 0% 100%"
    due_n = max(total - counts["waived"], 1)
    legend = "".join(
        f'<div class="lg-row"><span class="lg-dot" style="background:{color}"></span>'
        f'<span class="lg-name">{label}</span><span class="lg-n">{counts[k]}</span></div>'
        for k, (label, color) in STATUS.items())
    st.markdown(
        f'<div class="donut-wrap"><div class="donut" style="background:conic-gradient({gradient})">'
        f'<div class="donut-hole"><b>{counts["paid"]}/{due_n}</b><span>paid</span></div></div>'
        f'<div class="legend">{legend}</div></div>', unsafe_allow_html=True)


def chart_all_contributions(month: str, contribs: list, members: list):
    rows = []
    for c in contribs:
        per_c = month if is_monthly(c) else ""
        s = contrib_stats(c, roster(c, members, per_c), per_c)
        rows.append({"Contribution": c["name"], "Collected": s["collected"],
                     "Expected": max(s["expected"], s["collected"])})
    df = pd.DataFrame(rows)
    if df[["Collected", "Expected"]].to_numpy().sum() == 0:
        chart_empty("Nothing collected yet", "Record a payment to see each contribution here.")
        return
    domain = list(df["Contribution"])
    colour = alt.Color("Contribution:N", scale=alt.Scale(domain=domain, range=[color_of(c) for c in contribs]),
                       legend=None)
    y = alt.Y("Contribution:N", title=None, sort=domain, axis=alt.Axis(labelLimit=110))
    tip = ["Contribution", alt.Tooltip("Collected:Q", format=",.2f"), alt.Tooltip("Expected:Q", format=",.2f")]
    expected = (alt.Chart(df)
                .mark_bar(size=28, opacity=0.22, cornerRadiusTopRight=9, cornerRadiusBottomRight=9)
                .encode(y=y, x=alt.X("Expected:Q", title=None), color=colour, tooltip=tip))
    collected = (alt.Chart(df)
                 .mark_bar(size=15, cornerRadiusTopRight=8, cornerRadiusBottomRight=8)
                 .encode(y=y, x=alt.X("Collected:Q", title=None), color=colour, tooltip=tip))
    chart = alt.layer(expected, collected).properties(height=max(110, 52 * len(contribs) + 30)) \
        .configure_view(strokeWidth=0)
    st.altair_chart(chart, **FULL_CHART)
    st.caption("The solid bar is what has been collected. The light bar is what is expected.")


def render_dashboard_charts(month: str, contribs: list, members: list, c: dict, stats: dict, period: str):
    a, b = st.columns([3, 2])  # side by side on a laptop, stacked on a phone
    with a:
        with st.container(key="card_chart_collect"):
            st.markdown(card_title("Money collected", "Last 6 months"), unsafe_allow_html=True)
            chart_collections(month, contribs)
    with b:
        with st.container(key="card_chart_status"):
            label = f"{c.get('name', '')}" + (f", {month_label(period)}" if period else "")
            st.markdown(card_title("Who has paid", label), unsafe_allow_html=True)
            chart_status_donut(stats["statuses"])

    with st.container(key="card_chart_all"):
        st.markdown(card_title("All contributions", "Collected against expected"), unsafe_allow_html=True)
        chart_all_contributions(month, contribs, members)


# ─────────────────────────────────────────────────────────────
# Menu (opened from the three-line button on the dashboard)
# A short list of actions. Choosing one shows only that form, with a Back button.
# ─────────────────────────────────────────────────────────────
MENU_ITEMS = [("add", "Add a member"), ("leave", "Remove a member"), ("new", "New contribution"),
              ("edit", "Edit contribution"), ("del", "Remove contribution")]


def _set_menu(view):
    st.session_state["menu_view"] = view


def _menu_done():
    st.session_state["menu_view"] = None


def members_join_form(c: dict, all_members: list):
    cid = mid(c)
    ids = participant_ids(c)
    outside = {mid(m): m.get("name", "") for m in all_members if ids is not None and mid(m) not in ids}
    with st.form(f"cjoin_{cid}", clear_on_submit=True):
        chosen = None
        if outside:
            chosen = st.selectbox("Existing member", [None] + list(outside), key=f"cj_sel_{cid}",
                                  format_func=lambda v: "Choose a member" if v is None else outside[v])
        st.caption("Or add someone new to the group." if outside else
                   "Everyone is already in. You can still add someone new.")
        new_name = st.text_input("New person's name", key=f"cj_name_{cid}")
        new_contact = st.text_input("Contact (optional)", key=f"cj_contact_{cid}")
        if st.form_submit_button("Add to contribution", type="primary", **FULL_SUBMIT):
            to_add = [chosen] if chosen is not None else []
            if new_name.strip():
                person = find_member(new_name)
                if person is None:
                    add_member(new_name, new_contact)
                    person = find_member(new_name)
                if person is not None:
                    to_add.append(mid(person))
            if not to_add:
                st.warning("Choose a member or type a name.")
            else:
                join_contribution(cid, to_add)
                flash(f"Added {len(set(to_add))} member(s) to {c.get('name')}.")
                _menu_done()
                st.rerun()


def members_leave_form(c: dict, all_members: list):
    cid = mid(c)
    ids = participant_ids(c)
    current = {mid(m): m.get("name", "") for m in all_members if ids is None or mid(m) in ids}
    if not current:
        st.caption("No one is in this contribution.")
        return
    with st.form(f"cleave_{cid}", clear_on_submit=True):
        chosen = st.selectbox("Who is leaving?", [None] + list(current), key=f"cl_sel_{cid}",
                              format_func=lambda v: "Choose a member" if v is None else current[v])
        purge = st.checkbox("Also delete their payment records", key=f"cl_purge_{cid}")
        st.caption("Otherwise past payments stay in the history.")
        go = st.form_submit_button("Remove from contribution", type="primary", **FULL_SUBMIT)
    if go:
        if chosen is None:
            st.warning("Choose a member.")
        else:
            confirm_leave_member(cid, chosen, current[chosen], c.get("name", ""), purge)


def new_contrib_form():
    with st.form("add_contrib", clear_on_submit=True):
        name = st.text_input("Name", placeholder="e.g. Funeral, Welfare, Building Fund", key="nc_name")
        freq = st.selectbox("How often", ["One-off", "Monthly"], key="nc_freq",
                            help="Monthly contributions are tracked month by month, like Dues.")
        amount = st.number_input(f"Amount per member ({SETTINGS['currency']})", min_value=0.0, step=1.0,
                                 format="%.2f", value=0.0, key="nc_amount",
                                 help="Leave at 0 if members give different amounts.")
        notes = st.text_input("Notes (optional)", key="nc_notes")
        if st.form_submit_button("Create contribution", type="primary", **FULL_SUBMIT):
            ok, msg, _ = add_contribution(name, "monthly" if freq == "Monthly" else "one-off", amount, notes)
            if ok:
                st.session_state["_pending_contrib"] = name.strip()
                flash(msg)
                _menu_done()
                st.rerun()
            else:
                st.error(msg)


def edit_contrib_form(c: dict):
    cid = mid(c)
    per = to_float(c.get("amount"))
    with st.form(f"cedit_{cid}"):
        n = st.text_input("Name", value=c.get("name", ""), key=f"cen_{cid}")
        if c.get("member_amounts"):
            st.caption("Amounts come from each member's dues setting on the Members page, "
                       "and the default amount in Settings.")
            a = 0.0
        else:
            a = st.number_input(f"Amount per member ({SETTINGS['currency']})", min_value=0.0, step=1.0,
                                format="%.2f", value=per, key=f"cea_{cid}",
                                help="Leave at 0 if members give different amounts.")
        nt = st.text_input("Notes", value=c.get("notes", ""), key=f"cenote_{cid}")
        if st.form_submit_button("Save changes", type="primary", **FULL_SUBMIT):
            ok, msg = update_contribution(cid, n, a, nt)
            if ok:
                st.session_state["_pending_contrib"] = n.strip()
                flash(msg)
                _menu_done()
                st.rerun()
            else:
                st.error(msg)


def remove_contrib_body(c: dict):
    cid = mid(c)
    count = len([p for p in load_payments() if mid(p.get("contribution_id")) == cid])
    st.write(f"Remove **{c.get('name')}**? This also deletes its {count} payment record(s).")
    if st.button("Yes, remove", key=f"del_contrib_{cid}", **FULL_BTN):
        delete_contribution(cid)
        st.session_state["_pending_contrib"] = None
        flash(f"Removed {c.get('name')}.")
        _menu_done()
        st.rerun()


def render_menu(c, all_members: list):
    """Body of the menu popover. c is None when no contribution exists yet."""
    if c is None:
        st.markdown('<div class="menu-head">New contribution</div>', unsafe_allow_html=True)
        new_contrib_form()
        return
    view = st.session_state.get("menu_view")
    titles = dict(MENU_ITEMS)
    if view not in titles:
        st.markdown(f'<div class="menu-sub">{esc(c.get("name", ""))}</div>', unsafe_allow_html=True)
        for key, label in MENU_ITEMS:
            st.button(label, key=f"mnu_{key}", on_click=_set_menu, args=(key,), **FULL_BTN)
        return
    st.button("‹ Back", key="mnu_back", on_click=_set_menu, args=(None,))
    st.markdown(f'<div class="menu-head">{esc(titles[view])}</div>', unsafe_allow_html=True)
    if view == "add":
        members_join_form(c, all_members)
    elif view == "leave":
        members_leave_form(c, all_members)
    elif view == "new":
        new_contrib_form()
    elif view == "edit":
        edit_contrib_form(c)
    else:
        remove_contrib_body(c)


def _shift(delta: int):
    st.session_state["month"] = shift_month(st.session_state["month"], delta)


def _this_month():
    st.session_state["month"] = date.today().strftime("%Y-%m")


def month_navigator(month: str):
    with st.container(key="hrow_month"):
        a, b, c = st.columns([1, 3, 1], vertical_alignment="center")
        a.button("‹", key="mn_prev", on_click=_shift, args=(-1,), **FULL_BTN)
        b.markdown(f'<div class="month-label">{esc(month_label(month))}</div>', unsafe_allow_html=True)
        c.button("›", key="mn_next", on_click=_shift, args=(1,), **FULL_BTN)
    if month != date.today().strftime("%Y-%m"):
        st.button("Back to this month", key="mn_now", on_click=_this_month, **FULL_BTN)


# ─────────────────────────────────────────────────────────────
# Page: Dashboard
# ─────────────────────────────────────────────────────────────
def page_dashboard(month: str):
    contribs = sorted(load_contribs(), key=mid)

    group_banner()

    # Top bar: title on the left, menu button (three stacked lines) on the right
    with st.container(key="hrow_top"):
        left, right = st.columns([1, 1], vertical_alignment="center")
        left.markdown('<div class="page-title" style="margin-bottom:.1rem">Dashboard</div>'
                      '<div class="page-sub" style="margin:0">Record and track contributions.</div>',
                      unsafe_allow_html=True)
        menu = right.popover("☰", **FULL_POP)

    if not contribs:
        with menu:
            render_menu(None, [])
        st.markdown('<div class="empty"><b>No contributions yet</b><br>'
                    'Open the menu at the top right and choose New contribution, for example Dues or Funeral.</div>',
                    unsafe_allow_html=True)
        return

    # Contribution switcher
    names = [c["name"] for c in contribs]
    if "_pending_contrib" in st.session_state:
        pending = st.session_state.pop("_pending_contrib")
        if pending is None:
            st.session_state.pop("contrib_toggle", None)
        else:
            st.session_state["contrib_toggle"] = pending
    last = st.session_state.get("_last_contrib")
    if st.session_state.get("contrib_toggle") not in names:
        st.session_state["contrib_toggle"] = last if last in names else names[0]

    picked = pick_one("Contribution", names, key="contrib_toggle")
    sel_name = picked if picked in names else (last if last in names else names[0])
    st.session_state["_last_contrib"] = sel_name

    c = next(x for x in contribs if x["name"] == sel_name)
    cid = mid(c)
    monthly = is_monthly(c)
    period = month if monthly else ""
    per = to_float(c.get("amount"))

    # Menu contents (the popover was created first, filled now that the contribution is known)
    with menu:
        render_menu(c, load_members())

    if monthly:
        month_navigator(month)

    if c.get("member_amounts"):
        amount_text = "amount set for each member"
    elif per > 0:
        amount_text = f"{money(per)} per member"
    else:
        amount_text = "no fixed amount"
    title = c["name"] + (f", {month_label(month)}" if monthly else "")
    st.markdown(f'<div class="big-title" style="margin:.2rem 0 0">{esc(title)}</div>'
                f'<div class="p-sub" style="margin-bottom:.7rem">'
                f'{esc(("Monthly" if monthly else "One-off") + ", " + amount_text)}</div>',
                unsafe_allow_html=True)

    if c.get("notes"):
        st.caption(c["notes"])

    all_members = load_members()
    if not all_members:
        st.markdown('<div class="empty"><b>No members yet</b><br>Add your first member to start tracking.</div>',
                    unsafe_allow_html=True)
        st.button("Add members", type="primary", on_click=go_to, args=("members",), key="empty_add")
        return
    members = roster(c, all_members, period)
    if not members:
        st.markdown('<div class="empty"><b>No one in this contribution yet</b><br>'
                    'Open the menu at the top right and choose Add a member.</div>', unsafe_allow_html=True)
        return

    stats = contrib_stats(c, members, period)
    statuses, pay_map = stats["statuses"], stats["pay_map"]
    n = len(members)
    pct = round(100 * stats["paid_n"] / stats["countable"])
    expected = stats["expected"]

    st.markdown(tiles_html([
        ("Members", n, f"{stats['waived_n']} waived" if stats["waived_n"] else "in the group", COBALT),
        ("Paid", stats["paid_n"], f"of {stats['countable']} due", MINT),
        ("Outstanding", stats["unpaid_n"], "still to give", CORAL),
        ("Collected", money(stats["collected"]),
         f"of {money(expected)} expected" if expected else "no fixed amount", SUN, "#17153B"),
    ]), unsafe_allow_html=True)
    st.markdown(progress_html(pct, f"{pct}% of members are settled"), unsafe_allow_html=True)

    if c.get("member_amounts") and SETTINGS["default_dues"] == 0 and not any(m.get("dues") for m in members):
        st.caption("Tip: set a default dues amount in Settings so Mark paid fills in the amount for you.")
    elif not c.get("member_amounts") and per == 0:
        st.caption("This contribution has no fixed amount. Use Edit on a member's row to record what they gave.")

    # Member list
    with st.container(key="hrow_mhead"):
        h1, h2 = st.columns([1.3, 1], vertical_alignment="center")
        h1.markdown('<div class="section-title" style="margin:0">Members</div>', unsafe_allow_html=True)
        if any(due_for(c, m) > 0 for m in members):
            with h2.popover("Mark all paid", **FULL_POP):
                st.write(f"Mark every unpaid member as paid for {title}?")
                if st.button("Yes, mark all paid", key=f"markall_{cid}_{period}", **FULL_BTN):
                    pairs = [(m, due_for(c, m)) for m in members
                             if statuses[mid(m)] == "unpaid" and due_for(c, m) > 0]
                    count = set_payments_bulk(c, period, pairs)
                    flash(f"Marked {count} member(s) as paid.")
                    st.rerun()

    q = st.text_input("Search", placeholder="Search members", label_visibility="collapsed", key=f"dq_{cid}")
    flt = pick_one("Show", ["All", "Unpaid", "Paid"], key=f"df_{cid}", default="All") or "All"

    shown = [m for m in members
             if q.strip().lower() in m.get("name", "").lower()
             and (flt == "All"
                  or (flt == "Paid" and statuses[mid(m)] == "paid")
                  or (flt == "Unpaid" and statuses[mid(m)] in ("unpaid", "partial")))]

    if not shown:
        st.markdown('<div class="empty"><b>Nothing to show</b><br>No members match this search or filter.</div>',
                    unsafe_allow_html=True)
    else:
        st.caption(f"Showing {len(shown)} of {len(members)} members")
        # Open by default for small groups; larger groups can collapse it to reach the charts
        open_list = len(members) <= 10 or bool(q.strip()) or flt != "All"
        with st.expander("Member list", expanded=open_list):
            with st.container(key="mlist"):
                for m in sorted(shown, key=lambda x: x.get("name", "").lower()):
                    i = mid(m)
                    s = statuses[i]
                    p = pay_map.get(i)
                    due = due_for(c, m)
                    with st.container(key=f"mcard_{cid}_{period}_{i}"):
                        st.markdown(f'<div class="mrow">{person_html(m)}{chip_html(s)}</div>'
                                    f'{amount_line_html(s, p, due)}', unsafe_allow_html=True)

                        can_pay = due > 0 and s not in ("paid", "waived")
                        cols = st.columns(1 + int(can_pay) + int(bool(p)))
                        k = 0
                        if can_pay:
                            if cols[k].button("Mark paid", key=f"pay_{cid}_{period}_{i}"):
                                set_payment(c, m, period, "paid", due)
                                flash(f"{m.get('name')} marked as paid.")
                                st.rerun()
                            k += 1

                        with cols[k].popover("Edit"):
                            with st.form(f"pf_{cid}_{period}_{i}"):
                                labels = [STATUS[x][0] for x in STATUS]
                                start = s if p else "paid"
                                sel = st.selectbox("Status", labels, index=list(STATUS).index(start),
                                                   key=f"ps_{cid}_{period}_{i}")
                                amt = st.number_input("Amount given", min_value=0.0, step=1.0, format="%.2f",
                                                      value=to_float(p.get("amount", due)) if p else float(due),
                                                      key=f"pa_{cid}_{period}_{i}")
                                note = st.text_input("Note", value=(p.get("notes") or "") if p else "",
                                                     key=f"pn_{cid}_{period}_{i}")
                                if st.form_submit_button("Save", type="primary", **FULL_SUBMIT):
                                    key = [x for x in STATUS if STATUS[x][0] == sel][0]
                                    set_payment(c, m, period, key, amt, note)
                                    flash(f"Saved {m.get('name')}.")
                                    st.rerun()
                        k += 1

                        if p and cols[k].button("Undo", key=f"undo_{cid}_{period}_{i}"):
                            clear_payment(c, m, period)
                            flash(f"Cleared {m.get('name')}'s record.")
                            st.rerun()

    # Charts sit below the list so recording payments stays one scroll away
    st.write("")
    render_dashboard_charts(month, contribs, all_members, c, stats, period)


# ─────────────────────────────────────────────────────────────
# Page: Members
# ─────────────────────────────────────────────────────────────
def page_members():
    members = load_members()
    page_header("Members", f"{len(members)} in the group. Add, edit or remove people here.")

    with st.container(key="hrow_msearch"):
        t1, t2 = st.columns([2, 1], vertical_alignment="center")
        q = t1.text_input("Search", placeholder="Search members", label_visibility="collapsed", key="mem_q")
        with t2.popover("Add", **FULL_POP):
            tab_one, tab_many = st.tabs(["One person", "Paste a list"])
            with tab_one:
                with st.form("add_one", clear_on_submit=True):
                    name = st.text_input("Name", key="add_name")
                    contact = st.text_input("Contact (phone or email)", key="add_contact")
                    dues = st.number_input(f"Dues amount ({SETTINGS['currency']})", min_value=0.0, step=1.0,
                                           format="%.2f", value=float(SETTINGS["default_dues"]), key="add_dues")
                    notes = st.text_input("Notes", key="add_notes")
                    if st.form_submit_button("Add member", type="primary", **FULL_SUBMIT):
                        d = None if dues == SETTINGS["default_dues"] else dues
                        ok, msg = add_member(name, contact, d, notes)
                        if ok:
                            flash(msg)
                            st.rerun()
                        else:
                            st.error(msg)
            with tab_many:
                with st.form("add_many", clear_on_submit=True):
                    st.caption("One person per line. Add a contact after a comma if you like.")
                    text = st.text_area("Names", height=150, placeholder="Ama Mensah, 024 000 0000\nKofi Boateng",
                                        key="add_many_text")
                    if st.form_submit_button("Add all", type="primary", **FULL_SUBMIT):
                        a, s = add_many(text)
                        flash(f"Added {a} member(s)" + (f", skipped {s} duplicate or empty line(s)." if s else "."))
                        st.rerun()

    shown = [m for m in members if q.strip().lower() in m.get("name", "").lower()]
    if not members:
        st.markdown('<div class="empty"><b>No members yet</b><br>Use Add to create the first one.</div>',
                    unsafe_allow_html=True)
        return
    if not shown:
        st.markdown('<div class="empty"><b>No match</b><br>Try a different name.</div>', unsafe_allow_html=True)
        return

    pay_counts = {}
    for p in load_payments():
        pay_counts[mid(p.get("member_id"))] = pay_counts.get(mid(p.get("member_id")), 0) + 1

    with st.container(key="mlist"):
        for m in sorted(shown, key=lambda x: x.get("name", "").lower()):
            i = mid(m)
            with st.container(key=f"mcard_mem_{i}"):
                right = (f'<div class="mr"><div class="amt">{esc(money(dues_of(m)))}</div>'
                         f'<div class="p-sub">dues amount</div></div>')
                st.markdown(f'<div class="mrow">{person_html(m)}{right}</div>', unsafe_allow_html=True)
                if m.get("notes"):
                    st.markdown(f'<div class="p-sub">{esc(m["notes"])}</div>', unsafe_allow_html=True)

                c1, c2 = st.columns(2)
                with c1.popover("Edit"):
                    with st.form(f"ef_{i}"):
                        n = st.text_input("Name", value=m.get("name", ""), key=f"en_{i}")
                        c = st.text_input("Contact", value=m.get("contact", "") or "", key=f"ec_{i}")
                        d = st.number_input(f"Dues amount ({SETTINGS['currency']})", min_value=0.0, step=1.0,
                                            format="%.2f", value=dues_of(m), key=f"ed_{i}")
                        nt = st.text_input("Notes", value=m.get("notes", "") or "", key=f"eno_{i}")
                        if st.form_submit_button("Save changes", type="primary", **FULL_SUBMIT):
                            stored = None if d == SETTINGS["default_dues"] else d
                            ok, msg = update_member(i, n, c, stored, nt)
                            if ok:
                                flash(msg)
                                st.rerun()
                            else:
                                st.error(msg)

                # Opens a warning pop-up; nothing is deleted until the user confirms there
                if c2.button("Delete", key=f"del_mem_{i}"):
                    confirm_delete_member(i, m.get("name", ""), pay_counts.get(i, 0))


# ─────────────────────────────────────────────────────────────
# Page: Reports
# ─────────────────────────────────────────────────────────────
def report_filename(member, month, contrib, ext: str) -> str:
    parts = [SETTINGS["group_name"], member or "all_members", month or "all_months", contrib or "all_contributions"]
    slug = re.sub(r"[^A-Za-z0-9]+", "_", "_".join(parts)).strip("_").lower()
    return f"report_{slug}.{ext}"


def page_reports():
    page_header("Reports", "Choose what to include, then download it as Excel or Word.")
    members = sorted(load_members(), key=lambda x: x.get("name", "").lower())
    contribs = sorted(load_contribs(), key=mid)
    if not members or not contribs:
        st.markdown('<div class="empty"><b>Nothing to report yet</b><br>'
                    'Add members and a contribution first.</div>', unsafe_allow_html=True)
        return

    m_by = {mid(m): m.get("name", "") for m in members}
    c_by = {mid(c): c.get("name", "") for c in contribs}
    months = export_months(st.session_state.get("month", ""))

    # forget a stored choice that no longer exists (for example a deleted member)
    for key, valid in (("rp_member", m_by), ("rp_contrib", c_by), ("rp_month", months)):
        if st.session_state.get(key) not in [None] + list(valid):
            st.session_state.pop(key, None)

    with st.container(key="card_report"):
        a, b, c3 = st.columns(3)
        who = a.selectbox("Member", [None] + list(m_by), key="rp_member",
                          format_func=lambda v: "All members" if v is None else m_by[v])
        when = b.selectbox("Month", [None] + months, key="rp_month",
                           format_func=lambda v: "All months" if v is None else month_label(v))
        what = c3.selectbox("Contribution", [None] + list(c_by), key="rp_contrib",
                            format_func=lambda v: "All contributions" if v is None else c_by[v])

    data = build_report(who, when, what)
    if when and data["has_one_off"]:
        st.caption("One-off contributions are not tied to a month, so they show everything recorded.")

    if not data["lines"]:
        st.markdown('<div class="empty"><b>No records match</b><br>Try a different member, month or contribution.</div>',
                    unsafe_allow_html=True)
        return

    st.markdown(f'<div class="p-sub" style="margin-top:.4rem">{esc(SETTINGS["group_name"])}</div>'
                f'<div class="big-title" style="margin:0 0 .6rem">{esc(data["title"])}</div>', unsafe_allow_html=True)
    st.markdown(tiles_html([
        ("Expected", money(data["expected"]), "for this selection", COBALT),
        ("Collected", money(data["collected"]), "paid so far", MINT),
        ("Outstanding", money(data["balance"]), "still to give", CORAL),
        ("Records", len(data["lines"]), f"{len(data['by_contrib'])} contribution(s)", SUN, "#17153B"),
    ]), unsafe_allow_html=True)

    st.write("")
    sig = data_signature()
    user = st.session_state.get("user", "")
    user_id = uid()
    m_name = m_by.get(who)
    c_name = c_by.get(what)
    with st.container(key="hrow_report"):
        d1, d2 = st.columns(2)
        try:
            d1.download_button("Download Excel", cached_excel(who, when, what, user_id, user, sig),
                               file_name=report_filename(m_name, when, c_name, "xlsx"),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="dl_xlsx", **FULL_DL)
        except ImportError:
            d1.warning("Excel needs openpyxl. Run: pip install openpyxl")
        except Exception as e:
            d1.warning(f"Excel report failed: {e}")
        try:
            d2.download_button("Download Word", cached_word(who, when, what, user_id, user, sig),
                               file_name=report_filename(m_name, when, c_name, "docx"),
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               key="dl_docx", **FULL_DL)
        except ImportError:
            d2.warning("Word needs python-docx. Run: pip install python-docx")
        except Exception as e:
            d2.warning(f"Word report failed: {e}")

    with st.expander("Preview"):
        items = []
        for x in data["lines"][:25]:
            items.append(
                f'<div class="rec"><div class="rec-top"><span class="p-name">{esc(x["Member"])}</span>'
                f'<span class="amt">{esc(money(x["Paid"] if x["Paid"] else x["Due"]))}</span></div>'
                f'<div class="p-sub">{esc(x["Contribution"])} · {esc(x["Period"])}</div>'
                f'<div class="rec-bot">{chip_html(x["_status"])}</div></div>')
        st.markdown('<div class="reclist">' + "".join(items) + "</div>", unsafe_allow_html=True)
        if len(data["lines"]) > 25:
            st.caption(f"Showing 25 of {len(data['lines'])} lines. The download has all of them.")


# ─────────────────────────────────────────────────────────────
# Page: History
# ─────────────────────────────────────────────────────────────
HIST_STEP = 40


def _more_history():
    st.session_state["hist_limit"] = st.session_state.get("hist_limit", HIST_STEP) + HIST_STEP


def page_history():
    page_header("History", "Every payment record and every action taken in the app.")
    tab_pay, tab_act = st.tabs(["Payments", "Activity"])

    with tab_pay:
        payments = load_payments()
        cnames = {mid(c): c.get("name", "") for c in load_contribs()}
        payments = [p for p in payments if mid(p.get("contribution_id")) in cnames]
        if not payments:
            st.markdown('<div class="empty"><b>No payments yet</b><br>Records appear here once you mark someone as paid.</div>',
                        unsafe_allow_html=True)
        else:
            c_opts = ["All contributions"] + sorted(set(cnames.values()))
            m_opts = ["Everyone"] + sorted({p.get("member_name", "") for p in payments})
            periods = sorted({p.get("period", "") for p in payments if p.get("period", "")}, reverse=True)
            p_opts = [None] + periods + ([""] if any(p.get("period", "") == "" for p in payments) else [])
            with st.expander("Filters"):
                which = st.selectbox("Contribution", c_opts, key="h_which")
                who = st.selectbox("Member", m_opts, key="h_who")
                when = st.selectbox("Period", p_opts, key="h_when",
                                    format_func=lambda v: "All periods" if v is None
                                    else ("One-off" if v == "" else month_label(v)))

            rows = [p for p in payments
                    if (which == "All contributions" or cnames.get(mid(p.get("contribution_id"))) == which)
                    and (who == "Everyone" or p.get("member_name") == who)
                    and (when is None or p.get("period", "") == when)]
            rows.sort(key=lambda p: (p.get("period", ""), str(p.get("recorded_at", ""))), reverse=True)

            if not rows:
                st.info("No records match these filters.")
            else:
                total = sum(to_float(p.get("amount")) for p in rows)
                st.markdown(f'<div class="p-sub" style="margin-bottom:.5rem">{len(rows)} record(s), '
                            f'{esc(money(total))} in total</div>', unsafe_allow_html=True)

                limit = st.session_state.get("hist_limit", HIST_STEP)
                items = []
                for p in rows[:limit]:
                    cname = cnames.get(mid(p.get("contribution_id")), "")
                    per = month_label(p["period"]) if p.get("period") else "One-off"
                    by = f'<span class="p-sub">by {esc(p["recorded_by"])}</span>' if p.get("recorded_by") else ""
                    note = by + (f'<span class="p-sub">{esc(p["notes"])}</span>' if p.get("notes") else "")
                    items.append(
                        f'<div class="rec"><div class="rec-top"><span class="p-name">{esc(p.get("member_name", ""))}</span>'
                        f'<span class="amt">{esc(money(to_float(p.get("amount"))))}</span></div>'
                        f'<div class="p-sub">{esc(cname)} · {esc(per)}</div>'
                        f'<div class="rec-bot">{chip_html(norm_status(p.get("status")))}{note}</div></div>')
                st.markdown('<div class="reclist">' + "".join(items) + "</div>", unsafe_allow_html=True)
                if len(rows) > limit:
                    st.write("")
                    st.button(f"Show {min(HIST_STEP, len(rows) - limit)} more", key="hist_more",
                              on_click=_more_history, **FULL_BTN)

                with st.expander("Delete records"):
                    label_of = {p.get("id"): (f'{p.get("member_name", "")} · {cnames.get(mid(p.get("contribution_id")), "")} · '
                                              f'{month_label(p["period"]) if p.get("period") else "One-off"} · '
                                              f'{money(to_float(p.get("amount")))}') for p in rows}
                    picked = st.multiselect("Choose the records to remove", list(label_of),
                                            format_func=lambda x: label_of[x], key=f"hist_del_{which}_{who}_{when}")
                    blocked = lockout_message()
                    if blocked:
                        st.error(blocked)
                    else:
                        with st.form(f"hist_del_form_{which}_{who}_{when}", clear_on_submit=True):
                            pw = st.text_input("Enter your password to confirm", type="password",
                                               key=f"hist_pw_{which}_{who}_{when}")
                            go = st.form_submit_button(f"Delete selected ({len(picked)})", **FULL_SUBMIT)
                        if go:
                            if not picked:
                                st.warning("Choose at least one record.")
                            else:
                                ok, msg = verify_protected(pw, "Delete payment records")
                                if ok:
                                    delete_payments(picked)
                                    flash(f"Deleted {len(picked)} record(s).")
                                    st.rerun()
                                else:
                                    st.error(msg)

    with tab_act:
        activity = load_activity()
        if not activity:
            st.markdown('<div class="empty"><b>No activity yet</b></div>', unsafe_allow_html=True)
        else:
            people = sorted({a.get("user", "") for a in activity if a.get("user")})
            who = "Everyone"
            if people:
                with st.expander("Filter by user"):
                    who = st.selectbox("User", ["Everyone"] + people, key="act_user")
            rows = [a for a in activity if who == "Everyone" or a.get("user") == who]
            shown = rows[:100]
            if not shown:
                st.info("No activity for this user.")
            items = []
            for a in shown:
                by = f'{esc(a["user"])} · ' if a.get("user") else ""
                items.append(
                    f'<div class="rec"><div class="rec-top"><span class="p-name">{esc(str(a.get("action", "")))}</span>'
                    f'<span class="p-sub">{esc(str(a.get("timestamp", "")))}</span></div>'
                    f'<div class="p-sub">{by}{esc(str(a.get("details", "")))}</div></div>')
            if items:
                st.markdown('<div class="reclist">' + "".join(items) + "</div>", unsafe_allow_html=True)
            if len(rows) > len(shown):
                st.caption(f"Showing the latest {len(shown)} of {len(rows)} entries.")
            st.write("")
            with st.popover("Clear activity log", **FULL_POP):
                st.write("Remove all activity entries? Payments and members are not affected.")
                blocked = lockout_message()
                if blocked:
                    st.error(blocked)
                else:
                    with st.form("clear_log_form", clear_on_submit=True):
                        pw = st.text_input("Enter your password to confirm", type="password", key="clear_log_pw")
                        go = st.form_submit_button("Yes, clear log", **FULL_SUBMIT)
                    if go:
                        ok, msg = verify_protected(pw, "Clear activity log")
                        if ok:
                            clear_activity()
                            log_activity("Cleared the activity log")
                            flash("Activity log cleared.")
                            st.rerun()
                        else:
                            st.error(msg)


# ─────────────────────────────────────────────────────────────
# Page: Settings
# ─────────────────────────────────────────────────────────────
def page_settings():
    page_header("Settings", "Group details and data controls. Reports have their own page.")

    with st.container(key="card_set_group"):
        st.markdown("#### Group")
        with st.form("settings_form"):
            name = st.text_input("Group name", value=SETTINGS["group_name"], key="set_name", max_chars=80,
                                 help="Shown at the top of the dashboard and on every report.")
            cur = st.text_input("Currency symbol", value=SETTINGS["currency"], key="set_cur")
            dues = st.number_input("Default dues amount", min_value=0.0, step=1.0, format="%.2f",
                                   value=float(SETTINGS["default_dues"]), key="set_dues")
            if st.form_submit_button("Save settings", type="primary", **FULL_SUBMIT):
                save_settings(name, cur, dues)
                flash("Settings saved.")
                st.rerun()

    with st.container(key="card_set_account"):
        st.markdown("#### Account")
        who = st.session_state.get("email", "")
        if who:
            st.markdown(f'<div class="p-sub" style="margin-bottom:.5rem">Signed in as <b>{esc(who)}</b></div>',
                        unsafe_allow_html=True)
        if st.button("Refresh data", key="refresh_data", **FULL_BTN):
            st.session_state["_db"] = {}
            flash("Data refreshed.")
            st.rerun()
        if st.button("Log out", key="logout", **FULL_BTN):
            do_logout()
            st.rerun()

    with st.container(key="card_set_danger"):
        st.markdown("#### Danger zone")
        st.caption("Removes every member, contribution, payment and activity entry in this account. "
                   "This cannot be undone. "
                   f"Your password is needed, and {MAX_ATTEMPTS} wrong tries lock this for "
                   f"{LOCKOUT_SECONDS // 60} minutes.")
        with st.popover("Reset all data", **FULL_POP):
            blocked = lockout_message()
            if blocked:
                st.error(blocked)
            else:
                with st.form("reset_form", clear_on_submit=True):
                    confirm = st.text_input("Type RESET to confirm", key="reset_confirm")
                    pw = st.text_input("Enter your password", type="password", key="reset_pw")
                    go = st.form_submit_button("Erase everything", **FULL_SUBMIT)
                if go:
                    if confirm != "RESET":
                        st.error("Type RESET exactly as shown.")
                    else:
                        ok, msg = verify_protected(pw, "Reset all data")
                        if ok:
                            reset_all_data()
                            st.session_state["_pending_contrib"] = None
                            log_activity("Reset all data")
                            flash("All data erased.")
                            st.rerun()
                        else:
                            st.error(msg)


# ─────────────────────────────────────────────────────────────
# Page: Admin (administrators only)
# Uses the Supabase Admin client (service_role). Every function below goes through admin_client(),
# which re-checks the role in the database first. Deleting an auth user removes their profile,
# members, contributions, payments and activity through ON DELETE CASCADE in the schema.
# ─────────────────────────────────────────────────────────────
ADMIN_PAGE = 200


def _paged_rows(client, table: str, columns: str, order: str = "id") -> list:
    rows, start = [], 0
    while True:
        res = client.table(table).select(columns).order(order).range(start, start + PAGE - 1).execute()
        batch = res.data or []
        rows += batch
        if len(batch) < PAGE:
            return rows
        start += PAGE


def list_accounts() -> list:
    admin = admin_client()
    users, page = [], 1
    while True:
        res = admin.auth.admin.list_users(page=page, per_page=ADMIN_PAGE)
        batch = list(res if isinstance(res, list) else (getattr(res, "users", None) or []))
        users += batch
        if len(batch) < ADMIN_PAGE:
            break
        page += 1
    profiles = {str(p["id"]): p for p in _paged_rows(admin, "profiles", "id,group_name,role")}
    rows = []
    for u in users:
        p = profiles.get(str(u.id), {})
        role = str(p.get("role") or "user").lower()
        rows.append({
            "id": str(u.id),
            "email": u.email or "(no email)",
            "group": p.get("group_name") or "No profile yet",
            "role": role if role in ROLES else "user",
            "created": fmt_ts(getattr(u, "created_at", "")),
            "last_seen": fmt_ts(getattr(u, "last_sign_in_at", "")) or "Never",
        })
    rows.sort(key=lambda r: r["email"].lower())
    return rows


def set_account_role(target_id: str, email: str, role: str):
    if role not in ROLES:
        return False, "Unknown role."
    if target_id == uid():
        return False, "You cannot change your own role."
    try:
        admin = admin_client()
        admin.table("profiles").upsert({"id": target_id, "role": role}, on_conflict="id").execute()
    except PermissionError:
        return False, "Administrator access is required."
    except Exception as e:
        return False, f"Could not change the role: {e}"
    touch("admin_accounts")
    log_activity("Changed account role", f"{email} | {role}")
    return True, f"{email} is now {'an admin' if role == 'admin' else 'a regular user'}."


def delete_account(target_id: str, email: str):
    if target_id == uid():
        return False, "You cannot delete your own account here."
    try:
        admin = admin_client()
        res = admin.table("profiles").select("role").eq("id", target_id).limit(1).execute()
        if res.data and str(res.data[0].get("role") or "").lower() == "admin":
            return False, "Remove this account's admin access first, then delete it."
        admin.auth.admin.delete_user(target_id)
    except PermissionError:
        return False, "Administrator access is required."
    except Exception as e:
        return False, f"Could not delete the account: {e}"
    touch("admin_accounts")
    log_activity("Deleted account", email)
    return True, f"Deleted {email} and all of its data."


@_popup("Change role?")
def confirm_role_change(target_id: str, email: str, new_role: str):
    verb = "give administrator access to" if new_role == "admin" else "remove administrator access from"
    st.warning(f"You are about to {verb} **{email}**.")
    blocked = lockout_message()
    if blocked:
        st.error(blocked)
        return
    pw = st.text_input("Enter your password to confirm", type="password", key=f"role_pw_{target_id}")
    c1, c2 = st.columns(2)
    if c1.button("Confirm", key=f"role_go_{target_id}", **FULL_BTN):
        ok, msg = verify_protected(pw, "Change account role")
        if ok:
            ok, msg = set_account_role(target_id, email, new_role)
        if ok:
            flash(msg)
            st.rerun()
        else:
            st.error(msg)
    if c2.button("Cancel", key=f"dlg_cancel_role_{target_id}", **FULL_BTN):
        st.rerun()


@_popup("Delete account?")
def confirm_delete_account(target_id: str, email: str, group: str):
    st.warning(f"This permanently deletes **{email}** (group: {group}) together with all of its members, "
               "contributions, payments and activity. This cannot be undone.")
    blocked = lockout_message()
    if blocked:
        st.error(blocked)
        return
    typed = st.text_input("Type the account's email to confirm", key=f"del_email_{target_id}")
    pw = st.text_input("Enter your password", type="password", key=f"del_pw_{target_id}")
    c1, c2 = st.columns(2)
    if c1.button("Delete account", key=f"del_acct_go_{target_id}", **FULL_BTN):
        if typed.strip().lower() != email.strip().lower():
            st.error("The email does not match.")
        else:
            ok, msg = verify_protected(pw, "Delete user account")
            if ok:
                ok, msg = delete_account(target_id, email)
            if ok:
                flash(msg)
                st.rerun()
            else:
                st.error(msg)
    if c2.button("Cancel", key=f"dlg_cancel_acct_{target_id}", **FULL_BTN):
        st.rerun()


def page_admin():
    page_header("Admin", "Manage accounts. Only administrators can open this page.")
    if not verify_admin_now():
        st.error("You do not have access to this page.")
        return
    if not service_role_key():
        st.warning("The Admin page needs the service role key. Add service_role_key under [supabase] in "
                   ".streamlit/secrets.toml (or SUPABASE_SERVICE_ROLE_KEY) and restart the app.")
        return
    try:
        accounts = cached("admin_accounts", list_accounts)
    except Exception as e:
        st.error(f"Could not load accounts: {e}")
        return

    admins = sum(1 for a in accounts if a["role"] == "admin")
    st.markdown(tiles_html([
        ("Accounts", len(accounts), "registered", COBALT),
        ("Admins", admins, "full access", VIOLET),
    ]), unsafe_allow_html=True)
    st.write("")

    q = st.text_input("Search", placeholder="Search by email or group", label_visibility="collapsed", key="adm_q")
    needle = q.strip().lower()
    shown = [a for a in accounts if needle in a["email"].lower() or needle in a["group"].lower()]
    st.caption(f"Showing {len(shown)} of {len(accounts)} accounts")
    if not shown:
        st.markdown('<div class="empty"><b>No match</b><br>Try a different email or group.</div>', unsafe_allow_html=True)
        return

    with st.expander("Accounts", expanded=len(accounts) <= 10 or bool(needle)):
        with st.container(key="mlist"):
            for a in shown:
                aid = a["id"]
                is_admin_row = a["role"] == "admin"
                color = AVATARS[sum(map(ord, aid)) % len(AVATARS)]
                chip = (f'<span class="chip" style="background:{VIOLET if is_admin_row else COBALT}">'
                        f'{"Admin" if is_admin_row else "User"}</span>')
                with st.container(key=f"mcard_adm_{aid}"):
                    st.markdown(
                        f'<div class="mrow"><div class="person"><div class="avatar" style="background:{color}">'
                        f'{esc(initials(a["email"].split("@")[0]))}</div><div>'
                        f'<div class="p-name">{esc(a["email"])}</div>'
                        f'<div class="p-sub ell">{esc(a["group"])}</div></div></div>{chip}</div>'
                        f'<div class="mamt"><span class="p-sub">Joined {esc(a["created"][:10])}</span>'
                        f'<span class="p-sub">Last sign in {esc(a["last_seen"][:16])}</span></div>',
                        unsafe_allow_html=True)
                    if aid == uid():
                        st.caption("This is your account. It cannot be changed or deleted here.")
                    elif is_admin_row:
                        if st.button("Remove admin", key=f"adm_role_{aid}"):
                            confirm_role_change(aid, a["email"], "user")
                        st.caption("Remove admin access before deleting this account.")
                    else:
                        c1, c2 = st.columns(2)
                        if c1.button("Make admin", key=f"adm_role_{aid}"):
                            confirm_role_change(aid, a["email"], "admin")
                        if c2.button("Delete", key=f"del_acct_{aid}"):
                            confirm_delete_account(aid, a["email"], a["group"])


# ─────────────────────────────────────────────────────────────
# App entry
# ─────────────────────────────────────────────────────────────
require_login()
prefetch()
SETTINGS = get_settings()

st.markdown(BASE_CSS, unsafe_allow_html=True)
st.session_state.setdefault("page", "dashboard")
st.session_state.setdefault("month", date.today().strftime("%Y-%m"))

show_flash()

# Only admins can reach the Admin page; anyone else is sent back to the dashboard
if st.session_state["page"] == "admin" and not is_admin():
    st.session_state["page"] = "dashboard"

render_nav()

page = st.session_state["page"]
if page == "dashboard":
    page_dashboard(st.session_state["month"])
elif page == "members":
    page_members()
elif page == "reports":
    page_reports()
elif page == "history":
    page_history()
elif page == "admin":
    page_admin()
else:
    page_settings()

st.markdown(copyright_html(), unsafe_allow_html=True)
