# bs

## Daily Google Sheet scripts (Apps Script)

Two separate scripts on [telesales_test](https://docs.google.com/spreadsheets/d/1STzx1zHsztQ1LNAE_0eF9FU0Crfcbes5Q62ZlxaJG6Q/edit?gid=292013814#gid=292013814):

1. **Check, then remove** rows whose **column L** is empty (`checkEmptyColumnL` / `removeEmptyColumnL`).
2. **Extract Metabase question 590** and **append** it under the rows that are already on the sheet (`appendMetabaseData`).

### Install in the sheet

1. Open the spreadsheet → **Extensions → Apps Script**.
2. Paste the files from `google_apps_script/`:
   - `Code.gs`, `Shared.gs`, `1_RemoveBlankColumnL.gs`, `2_AppendMetabase.gs`, `SyncLogic.js`
3. Script properties: `METABASE_USERNAME`, `METABASE_PASSWORD`
4. Run `runSelfTests`, then each script once (approve permissions).
5. Run `installDailyTriggers` (cleanup 04:00, append 04:15 Asia/Baghdad).

Details: `google_apps_script/README.md`.
