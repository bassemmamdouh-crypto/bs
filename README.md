# Orders Union Script (Google Sheets)

This repository includes a Google Apps Script to scan order data from multiple
sheets and union all rows into one target sheet.

## File

- `orders_union.gs`

## How to use

1. Open your Google Sheet.
2. Go to **Extensions -> Apps Script**.
3. Paste the contents of `orders_union.gs` into the script editor.
4. Update `UNION_ORDERS_CONFIG` as needed:
   - `targetSheetName`: destination sheet for merged data
   - `includeSheets`: explicit source sheets (optional)
   - `excludeSheets`: sheets to skip
   - `sourceNamePrefix`: only include sheets with this prefix (optional)
5. Save and run `unionOrdersData()`.
6. Optionally use the custom menu: **Orders Tools -> Union Orders**.
