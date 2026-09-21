# bs

## Daily Google Sheet + Metabase sync

`sheets_daily_sync.py` refreshes the **telesales_test** tab every day:

- Spreadsheet: [telesales workbook](https://docs.google.com/spreadsheets/d/1STzx1zHsztQ1LNAE_0eF9FU0Crfcbes5Q62ZlxaJG6Q/edit?gid=292013814#gid=292013814)
- Tab: `telesales_test` (gid `292013814`)
- Metabase question: `590`

Each run:

1. Keeps the header row.
2. Deletes every data row whose **column P** (the 16th column) is blank.
3. Pulls Metabase question 590 and **appends** those rows as the new day's data.

Rows that already have a value in column P stay on the sheet, so completed days accumulate.

### One-time Google setup

1. Create a Google Cloud service account and download its JSON key.
2. Enable the **Google Sheets API** and **Google Drive API** on that project.
3. Share `1STzx1zHsztQ1LNAE_0eF9FU0Crfcbes5Q62ZlxaJG6Q` with the service account email as Editor.

### GitHub Actions (runs every day at 01:00 UTC)

Required repository secrets:

| Secret | Purpose |
| --- | --- |
| `IRAQ_METABASE_USERNAME` | Metabase login |
| `IRAQ_METABASE_PASSWORD` | Metabase password |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | Full service-account JSON key |

Question id `590` and the telesales spreadsheet/tab are already the script defaults. Optional secrets `METABASE_QUESTION_ID`, `GOOGLE_SPREADSHEET_ID`, and `GOOGLE_WORKSHEET_NAME` override them.

Use **Run workflow** on the Actions tab to test once secrets are in place.

### Run locally

```bash
pip install -r requirements.txt
export IRAQ_METABASE_USERNAME=...
export IRAQ_METABASE_PASSWORD=...
export GOOGLE_SERVICE_ACCOUNT_FILE=/path/to/service-account.json
python sheets_daily_sync.py
```

Dry-run a CSV fixture without touching Google or Metabase:

```bash
python sheets_daily_sync.py \
  --local-input path/to/current_sheet.csv \
  --local-new path/to/metabase_export.csv \
  --local-output path/to/result.csv
```
