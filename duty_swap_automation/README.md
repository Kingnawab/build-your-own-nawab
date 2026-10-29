# Duty Swap Automation (CrewAI)

A CrewAI Flow that runs every morning and:

1. Reads recent unread emails in your Outlook inbox (Microsoft Graph).
2. Uses a CrewAI agent to decide which emails are duty-swap / cover requests and extract who, which dates and which duties.
3. Validates each request against the Excel duty roster on your OneDrive / SharePoint (the person must actually be on that date, the date must not have passed, and the requester must be part of the swap or a manager).
4. Swaps the names in the roster sheet, appends a row to a `Swap Log` sheet, and
5. Replies in the same email thread to the requester (CC'ing the other person and anyone in `NOTIFY_CC`) confirming the change, or explaining why it couldn't be made / asking for missing details.

The AI only *reads* emails. Every edit to the workbook is made by plain Python after validation, so a bad parse is rejected rather than written to your roster.

```
fetch_emails ──> parse_emails (CrewAI agent) ──> apply_and_notify ──> finish
   Outlook          ParsedEmail (pydantic)        Excel edits + replies    summary
```

Processed emails are marked read and tagged with the Outlook category `Duty Swap Processed` (or `Duty Swap Needs Attention` if anything was rejected or unclear), so they're never handled twice. Emails that aren't swap requests are left untouched and unread.

## Expected workbook layout

Sheet `Roster` (a title row above the header is fine; column names are configurable):

| Date       | Day     | Duty         | Assigned     |
|------------|---------|--------------|--------------|
| 2026-11-03 | Tuesday | Morning Gate | Alice Nguyen |
| 2026-11-04 | Wednesday | Morning Gate | Bob Martinez |

Optional sheet `Staff` with `Name` and `Email` columns. It's used to recognize who sent the email and whom to CC. Optional sheet `Swap Log` is the audit trail.

If your roster is a grid instead (names down the side, dates across the top), the only code that needs changing is `load_roster` / `plan_changes` in `roster.py`.

## Setup

```bash
cd duty_swap_automation
pip install uv && uv venv && uv pip install -e ".[dev]"   # or: crewai install
cp .env.example .env    # then fill it in
```

### 1. Register an app in Microsoft Entra ID (Azure AD)

Azure portal → **App registrations** → New registration.

- **Option A, unattended (recommended for a daily job):** add *Application* permissions `Mail.ReadWrite`, `Mail.Send`, `Files.ReadWrite.All` (or `Sites.Selected` for a single SharePoint site) and have an admin **grant consent**. Create a client secret and set `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`, `GRAPH_USER`. Ask IT to scope mailbox access to your mailbox with an [application access policy](https://learn.microsoft.com/graph/auth-limit-mailbox-access).
- **Option B, sign in once yourself:** runs on your own computer as you.
  1. [App registrations](https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) → *New registration*. Name it e.g. "Duty Swap Assistant", choose *Accounts in this organizational directory only*, leave Redirect URI empty, *Register*.
  2. Copy the **Application (client) ID** into `AZURE_CLIENT_ID` (and optionally the Directory (tenant) ID into `AZURE_TENANT_ID`).
  3. *Authentication* → *Advanced settings* → **Allow public client flows: Yes** → *Save*.
  4. *API permissions* → *Add a permission* → *Microsoft Graph* → *Delegated* → add `Mail.ReadWrite`, `Mail.Send`, `Files.ReadWrite.All`, `User.Read`.
  5. Leave `AZURE_CLIENT_SECRET` and `GRAPH_USER` blank and run `authorize`. It prints a code: open https://microsoft.com/devicelogin, enter it, sign in with your work account and accept the permissions. It then checks it can read your inbox and the roster file.

  The refresh token is saved in `.graph_token_cache.json` and the daily run refreshes it silently. Each run keeps it alive; if the job stops running for a long time (typically 90 days) or your password/MFA policy revokes sessions, just run `authorize` again. Keep that file private: it grants access to your mail and files.

  If you can't register apps, or sign-in says *Need admin approval*, your organisation blocks user consent, and you'll need IT to approve the app (they can grant consent on the same *API permissions* page).

### 2. Point it at the roster

Set `ROSTER_FILE_PATH` to the file's path in your OneDrive (e.g. `Documents/Duty Roster.xlsx`). For a SharePoint/Teams library also set `ROSTER_SITE_ID` (from `https://graph.microsoft.com/v1.0/sites/{host}:/sites/{site-name}`). You can also set `ROSTER_DRIVE_ID` + `ROSTER_ITEM_ID`.

### 3. Try it safely

```bash
DRY_RUN=true kickoff        # reads mail + roster, logs the replies it would send, changes nothing
kickoff                     # real run
```

Local sandbox with no Microsoft account (still needs an LLM key for the agent):

```bash
make_sample_data            # writes sample_data/duty_roster.xlsx + two demo emails
DUTY_SWAP_BACKEND=local kickoff
```

Tests (no network or LLM needed): `pytest`

## Running it every morning

- **CrewAI AMP / Studio:** push this folder to GitHub and deploy it from your CrewAI project (it is a standard `type = "flow"` CrewAI project). Add the `.env` values as environment variables in the deployment, then schedule the kickoff for the morning. The flow also accepts a trigger payload: `{"dry_run": true}`.
- **GitHub Actions:** copy `deploy/github-actions-daily.yml` into `.github/workflows/` (app-only auth required).
- **cron** on an always-on machine (cron uses the machine's local time zone):

  ```cron
  0 7 * * 1-5 cd /path/to/duty_swap_automation && .venv/bin/kickoff >> run.log 2>&1
  ```

- **Windows Task Scheduler:** daily trigger running `C:\path\to\duty_swap_automation\.venv\Scripts\kickoff.exe` with "Start in" set to the project folder.

## Configuration reference

See `.env.example`. Key settings:

| Variable | Purpose |
|---|---|
| `DRY_RUN` | Don't edit the workbook or send email |
| `SUBJECT_KEYWORDS` | Only emails containing these words go to the AI (saves cost, avoids sending unrelated mail to the LLM) |
| `LOOKBACK_HOURS` | How far back to look for unread mail (72 covers a weekend) |
| `REQUIRE_REQUESTER_IS_PARTY` / `MANAGER_EMAILS` | Who may request a swap |
| `NOTIFY_CC` | Always CC these people on confirmations |
| `ALLOW_PAST_DATES` | Allow edits to dates that have already passed |
