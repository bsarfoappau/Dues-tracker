# Dues Tracker (Supabase edition)

A Streamlit app for tracking dues and other contributions for a group (a department, class, association, church group and so on).

* **Sign up / sign in** with Supabase Auth (email and password).
* **Group Name** is captured at sign up, shown at the top of the dashboard and at the top of every Excel, Word and CSV report.
* **Cloud storage** in Supabase Postgres, so data survives restarts and redeploys.
* **Multi-tenant by design**: Row Level Security (RLS) means each account can only read and change its own rows.

## 1. Project layout

```
dues-tracker/
├── app.py
├── requirements.txt
├── supabase_schema.sql        # run once in Supabase
├── .gitignore
└── .streamlit/
    └── secrets.toml           # you create this; it is git-ignored
```

## 2. Create the Supabase project

1. Go to https://supabase.com, create a project and wait for it to finish provisioning.
2. Open Project Settings > API and note the Project URL and the anon / public key.

Use only the anon / public key in this app. Never use the service_role key: it bypasses RLS and would expose every account's data.

## 3. Create the tables and RLS policies

1. In the Supabase dashboard open SQL Editor > New query.
2. Paste the full contents of supabase_schema.sql and click Run.

This creates the tables profiles, members, contributions, payments and activity. Every data table has a user_id that references auth.users(id) with on delete cascade, so deleting an account removes its data.

How the isolation works:
* RLS is enabled on all five tables.
* Each table's policies allow an operation only when user_id = auth.uid() (for profiles, id = auth.uid()), for both reading and writing.
* The anon role (signed-out visitors) has all privileges revoked.
* payments uses composite foreign keys, so a payment can never reference another account's member or contribution.

How the Group Name is stored: at sign up the app sends it as user metadata. A database trigger (handle_new_user) copies it into profiles.group_name and creates the default monthly Dues contribution. It can be changed later under Settings > Group.

## 4. Configure Supabase Auth

In Authentication > Providers > Email:
* Confirm email ON (recommended for production): new users must click a link in an email before signing in.
* Confirm email OFF (handy while testing): users are signed in straight after creating the account.

In Authentication > URL Configuration, set Site URL to your deployed app URL (use http://localhost:8501 while developing).

## 5. Configure secrets

Create .streamlit/secrets.toml:

    [supabase]
    url = "https://YOUR-PROJECT-REF.supabase.co"
    anon_key = "YOUR-ANON-PUBLIC-KEY"

This file is listed in .gitignore so it is never committed. Check with git status before your first commit.

On Streamlit Community Cloud: open your app, then Settings > Secrets and paste the same two lines.

## 6. Run the app

    python -m venv .venv
    # Windows:  .venv\Scripts\activate
    # macOS/Linux:  source .venv/bin/activate

    pip install -r requirements.txt
    streamlit run app.py

Open http://localhost:8501, choose Create account, enter your Group Name, email and password, then start adding members.

## 7. Check that RLS works

1. Create two accounts (for example in two different browsers).
2. Add a member in account A. Confirm account B does not see them.
3. In Table Editor (which uses the admin role) you will see rows from both accounts, each with a different user_id. That is expected.

## 8. Reports

Reports page: pick a member, month and contribution, then download Excel, Word or CSV.
* Excel: every sheet starts with the Group Name in large letters on row 1.
* Word: the Group Name is the document title and also appears in the footer.
* CSV: the first line is "Group,<your group name>", followed by the report details, then the ledger.

Need a PDF? Open the Word report and choose Save as PDF.

## 9. Things worth knowing

* Refreshing the browser signs you out. Your data is safe in Supabase; just sign in again.
* Protected actions (reset all data, clear the activity log, delete payment records) ask for your password again. Three wrong tries lock them for five minutes.
* Existing local data from the old version (data/*.json) is not migrated.
* To reset a forgotten password, use Authentication > Users in the Supabase dashboard.

## 10. Troubleshooting

* "Supabase is not configured": secrets.toml is missing or the [supabase] keys are misspelled.
* "relation public.profiles does not exist": supabase_schema.sql has not been run.
* "new row violates row-level security policy": sign out and in again, and re-run the SQL file.
* Sign up works but no email arrives: check spam, or turn off Confirm email while testing.
* Group name shows "My Group": the account was created before the trigger existed. Change it under Settings > Group.
* "Invalid API key": wrong key, or a stray space or quote in secrets.toml.

(c) AppaTech. All rights reserved.