# In-sheet daily sync (Apps Script)

This is the intended daily runner. It lives on the Google Sheet itself: no service account and no GitHub Action required.

Target tab: [telesales_test](https://docs.google.com/spreadsheets/d/1STzx1zHsztQ1LNAE_0eF9FU0Crfcbes5Q62ZlxaJG6Q/edit?gid=292013814#gid=292013814)  
Metabase question: `590`

Each run keeps the header, deletes data rows with a blank **column P**, then appends question 590.

## Install

1. Open the spreadsheet.
2. **Extensions → Apps Script**.
3. Delete the stub `Code.gs` contents. Add two files:
   - `Code.gs` ← paste `google_apps_script/Code.gs`
   - `SyncLogic.js` ← paste `google_apps_script/SyncLogic.js` (Apps Script accepts `.js` in the editor; you can also name the file `SyncLogic.gs`)
4. **Project Settings → Script properties** and add:
   - `METABASE_USERNAME` = your Metabase login
   - `METABASE_PASSWORD` = your Metabase password
5. **Run → runSelfTests** once. Approve the Sheets permission prompt.
6. **Run → dailySync** once. Approve the external-request prompt (Metabase) if asked.
7. **Run → installDailyTrigger**. This schedules `dailySync` every day at **04:00 Asia/Baghdad**.

A **Metabase sync** menu also appears after you reload the sheet.

If Metabase is not reachable from Google (`UrlFetchApp` blocked, or login fails), check Executions in the Apps Script editor for the error body.
