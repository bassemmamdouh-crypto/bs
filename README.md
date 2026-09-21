# bs

## Daily Google Sheet + Metabase sync

`sheets_daily_sync.py` runs a daily refresh against one Google Sheet tab:

1. Keep the header row.
2. Delete every data row whose **column P** (the 16th column) is blank.
3. Pull a Metabase question and **append** those rows as the new day's data.

Rows that already have a value in column P stay on the sheet, so completed days accumulate.

### One-time Google setup

1. Create a Google Cloud service account and download its JSON key.
2. Enable the **Google Sheets API** and **Google Drive API** on that project.
3. Share the target spreadsheet with the service account email (`...@...iam.gserviceaccount.com`) as Editor.
4. Copy the spreadsheet ID from the URL: `https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit`.

### GitHub Actions (runs every day at 01:00 UTC)

Add these repository secrets:

| Secret | Purpose |
| --- | --- |
| `IRAQ_METABASE_USERNAME` | Metabase login |
| `IRAQ_METABASE_PASSWORD` | Metabase password |
| `METABASE_QUESTION_ID` | Card / question id to export |
| `GOOGLE_SPREADSHEET_ID` | Spreadsheet id from the URL |
| `GOOGLE_WORKSHEET_NAME` | Tab name (optional; first tab if empty) |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | Full service-account JSON key |

The workflow is `.github/workflows/daily-sheet-sync.yml`. Use **Run workflow** on the Actions tab to test it once secrets are in place.

### Run locally

```bash
pip install -r requirements.txt
export IRAQ_METABASE_USERNAME=...
export IRAQ_METABASE_PASSWORD=...
export METABASE_QUESTION_ID=123
export GOOGLE_SPREADSHEET_ID=...
export GOOGLE_SERVICE_ACCOUNT_FILE=/path/to/service-account.json
python sheets_daily_sync.py --worksheet "Sheet1"
```

Dry-run a CSV fixture without touching Google or Metabase:

```bash
python sheets_daily_sync.py \
  --local-input path/to/current_sheet.csv \
  --local-new path/to/metabase_export.csv \
  --local-output path/to/result.csv
```
