# bs

## Daily Google Sheet + Metabase sync (Apps Script)

The daily job runs **inside the spreadsheet** with Google Apps Script. No service account and no GitHub Action is required.

- Spreadsheet: [telesales_test](https://docs.google.com/spreadsheets/d/1STzx1zHsztQ1LNAE_0eF9FU0Crfcbes5Q62ZlxaJG6Q/edit?gid=292013814#gid=292013814)
- Tab gid: `292013814`
- Metabase question: `590`

Each run:

1. Keeps the header row.
2. Deletes every data row whose **column P** is blank.
3. Pulls Metabase question 590 and **appends** those rows as the new day's data.

### Install in the sheet

1. Open the spreadsheet → **Extensions → Apps Script**.
2. Paste `google_apps_script/Code.gs` into `Code.gs`.
3. Add a second file and paste `google_apps_script/SyncLogic.js`.
4. **Project Settings → Script properties**:
   - `METABASE_USERNAME`
   - `METABASE_PASSWORD`
5. Run `runSelfTests`, then `dailySync` (approve permissions).
6. Run `installDailyTrigger` once. It fires every day at **04:00 Asia/Baghdad**.

After reload, the sheet has a **Metabase sync** menu. Full notes: `google_apps_script/README.md`.

### Optional Python fallback

`sheets_daily_sync.py` implements the same rules for local CSV fixtures or GitHub Actions (`workflow_dispatch` only, so it does not compete with the in-sheet trigger).
