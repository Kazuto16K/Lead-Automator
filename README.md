# Lead Automator

Streamlit app: the manager types a note in plain English, Gemini turns it into JSON, and the
app fills `Local_Lead_Tracker.xlsx`. The View tab and the sidebar **Download Excel** button always
show the latest saved file.

## Run
```
pip install -r requirements.txt
streamlit run app.py
```
Gemini key: never shown in the UI. Locally put `GEMINI_API_KEY=...` in `.env`; on Streamlit Cloud add it
under App settings -> Secrets. Never commit `.env` or secrets.

## Flow
1. Pick the niche and which step(s) the note covers (Find / Audit / Qualify).
2. For Audit/Qualify, pick the lead (or let it auto-detect from the business name).
3. Write the note, click **Parse**, check/edit the table, click **Confirm**.

Automatic: Lead ID row, Date Added, Researched By, Stage (New -> Audited, never backwards unless
the note says Qualified/Disqualified), dropdown-value matching. Score, Priority and Audit Status are
Excel formulas (recalculated when the file is opened in Excel; the app shows the same values).
A timestamped backup goes in `backups/` before every write. Close the workbook in Excel before writing.

## Login
The app shows a username and password page first. Users live only in Streamlit Secrets (never in the repo):
```toml
GEMINI_API_KEY = "..."
GITHUB_TOKEN = "..."
GITHUB_REPO = "owner/data-repo"

[users]
manager = "a-password"
soumava = "another-password"
```
Keep the `[users]` table at the BOTTOM of the Secrets box: every line after a `[table]` header belongs to that table.
Usernames are not case sensitive. With no `[users]` set, nobody can log in (the app fails closed).
After a correct login the browser keeps a signed "stay signed in" cookie for 30 days (`REMEMBER_DAYS` in `auth.py`), so
refreshing or reopening the app does not ask again. Sign out deletes it. Changing a user's password in Secrets signs that
user out everywhere; removing a user locks them out. To add, remove or change a user, edit Secrets and save (the app restarts). For local runs copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml`.

## Deploy on Streamlit Community Cloud
GitHub Pages cannot run this app (it only serves static files). Use Streamlit Community Cloud (free):
1. Put the contents of this folder in a GitHub repository (app.py, excel_io.py, llm_parser.py, schema.py,
   requirements.txt, Local_Lead_Tracker.xlsx, .streamlit/config.toml). Do not commit .env or secrets.
   Make the repository private if the leads are sensitive.
2. share.streamlit.io -> sign in with GitHub -> Create app -> pick the repo, branch, main file `app.py`.
3. Advanced settings -> Secrets: `GEMINI_API_KEY = "your key"`.
4. Deploy. Under the app's Share settings, restrict viewing to your manager's email.
5. Open the app URL on the phone and use "Add to Home Screen" for an app-like icon.

## Make the data permanent with GitHub (recommended on Streamlit Cloud)
Streamlit Cloud's disk is temporary, so the app keeps its master copy of the Excel in a second, private GitHub repo:
the app downloads it when it starts and uploads it after every save or delete (one commit per change = full history).

1. Create a new PRIVATE repo for data only, for example `lead-data`. It must NOT be the repo the app deploys from,
   because a push to that repo restarts the app.
2. Create a token: GitHub -> Settings -> Developer settings -> Personal access tokens -> Fine-grained tokens ->
   Generate. Repository access: only `lead-data`. Permissions: Contents = Read and write. Copy the token.
3. In the Streamlit app's Secrets add:
```toml
GITHUB_TOKEN = "github_pat_..."
GITHUB_REPO  = "your-user/lead-data"
```
4. Reboot the app. The first start uploads the current Excel to `lead-data`. After that GitHub is the master copy.
5. Daily dated copies: add `github_backup/daily-backup.yml` to the data repo as `.github/workflows/daily-backup.yml`.
   Every night at 23:30 India time it saves `daily/Local_Lead_Tracker_YYYY-MM-DD.xlsx`. Run it once by hand from the
   repo's Actions tab ("Run workflow") to check it works.

If a GitHub upload fails the change is still saved in the app and a warning appears; download the Excel then.
Do not edit the Excel inside the data repo by hand while the app is running.

## Files
`app.py` UI, `auth.py` login, `neon.css` theme · `llm_parser.py` Gemini prompt + validation · `excel_io.py` Excel read/write, `github_sync.py` GitHub backup, `github_backup/` nightly workflow · `schema.py` field/column map
