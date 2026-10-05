# Retailer sheet consolidation

Builds one table from an Excel workbook that has **one tab per agent**.

## Output columns

| Column | Source |
| --- | --- |
| `retailer_id` | `كود` |
| `mobile_number` | `رقم الهاتف` |
| `market_name` | `الاسم التجاري` |
| `agent_name` | Tab name **outside** the brackets |
| `supervisor_name` | Tab name **inside** the brackets |

Example tab title: `أحمد علي (خالد حسن)`

- agent_name = `أحمد علي`
- supervisor_name = `خالد حسن`

## Run

```bash
python3 retailers/consolidate_retailer_sheets.py path/to/agents.xlsx \
  -o retailers/output/massabeh_retailers_consolidated.xlsx
```

Re-run this against the original `.xlsx` (multiple tabs, Arabic text intact). The file that was uploaded here was already flattened to a single CSV, and Arabic characters were replaced with `?`, so agent/supervisor cannot be recovered from tab names on that dump.
