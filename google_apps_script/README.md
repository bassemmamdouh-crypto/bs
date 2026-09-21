# In-sheet telesales scripts (Apps Script)

Two independent scripts on [telesales_test](https://docs.google.com/spreadsheets/d/1STzx1zHsztQ1LNAE_0eF9FU0Crfcbes5Q62ZlxaJG6Q/edit?gid=292013814#gid=292013814):

| Script | Function | What it does |
| --- | --- | --- |
| 1 | `checkEmptyColumnL` | Inspects the tab and reports rows whose **column L** is empty. Does not delete. |
| 1 | `removeEmptyColumnL` | Runs that check, logs the empty row numbers, then deletes those rows. Keeps the header. |
| 2 | `appendMetabaseData` | Pulls Metabase question **590** and **appends** it under the rows already on the sheet. Does not delete. |

## Install

1. Open the spreadsheet → **Extensions → Apps Script**.
2. Paste these files into the project:

   - `Code.gs`
   - `Shared.gs`
   - `1_RemoveBlankColumnL.gs`
   - `2_AppendMetabase.gs`
   - `SyncLogic.js` (or name it `SyncLogic.gs`)

3. **Project Settings → Script properties**:
   - `METABASE_USERNAME`
   - `METABASE_PASSWORD`
4. Run `runSelfTests`.
5. Run `checkEmptyColumnL`, then `removeEmptyColumnL`, then `appendMetabaseData` (approve permissions).
6. Run `installDailyTriggers` once:
   - Script 1 at **04:00** Asia/Baghdad
   - Script 2 at **04:15** Asia/Baghdad

Reload the sheet for the **Telesales scripts** menu.
