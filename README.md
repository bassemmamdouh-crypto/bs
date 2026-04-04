# Orders Union Script (Python)

This repository includes a Python script to scan order data from multiple Excel
sheets and union all rows into one target sheet.

## File

- `orders_union.py`

## Requirements

- Python 3.9+
- `openpyxl`

Install:

```bash
pip3 install openpyxl
```

## How to use

Run:

```bash
python3 orders_union.py --input orders.xlsx --output orders_merged.xlsx --target-sheet "All Orders"
```

### Optional arguments

- `--include "Orders Jan,Orders Feb"`: only read listed sheets
- `--exclude "Notes,Archive"`: skip listed sheets
- `--source-prefix "Orders "`: only read sheets with this prefix
- `--header-row 1`: header row index (1-based)
- `--no-source-column`: do not add source sheet name column
- `--no-clear-target`: keep existing rows in target sheet

## Example

```bash
python3 orders_union.py \
  --input orders.xlsx \
  --output orders_merged.xlsx \
  --target-sheet "All Orders" \
  --exclude "Summary,Notes"
```
