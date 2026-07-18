# Bundle Sales Planner

This repository includes an Excel template to build bundle actions for slow movers using your updated schema (last 4 months sales + stock fields).

## Files

- `bundle_planning_template.xlsx`: Ready-to-use workbook with input, scoring, bundle recommendation, and logic sheets.
- `create_bundle_workbook.py`: Script that generates/regenerates the workbook.
- `generate_all_possible_bundles.py`: Builds all possible bundles (max 3 products), then keeps only the top 15 by bundle priority score.

## How to use

1. Open `bundle_planning_template.xlsx`.
2. In `Input_Data`, paste your product data:
   - `clean_product_name`, `category`, `product_id`
   - `purchased_item_count`
   - `purchased_qty_m1` to `purchased_qty_m4` (current month excluded)
   - `cont_from_total` (as a percentage)
   - `stock`, `reserved_stock`, `available_stock`
3. Review `Scoring_Model`:
   - Movement, contribution, and stock pressure percentiles are auto-calculated.
   - `stock_data_check` flags rows where available stock does not match `stock - reserved_stock`.
   - Each product gets a bundle priority score and one of:
     - High Value Bundle
     - Medium Value Bundle
     - Low Value Bundle
4. Go to `Bundle_Recommendations`:
   - Focus on rows with `Low Movement`.
   - Every candidate bundle includes an anchor product to attract purchase.
   - Use slight discount guidance to control burn.
5. Generate bundle combinations (maximum 3 products) and keep the top 15:

```bash
python3 generate_all_possible_bundles.py bundle_planning_template.xlsx
```

This builds every possible bundle, ranks them by `bundle_priority_score`, and
writes only the top 15 to `All_Possible_Bundles` (controlled by `TOP_N_BUNDLES`):
- 2-product bundles: `1 anchor + 1 slow mover`
- 3-product bundles: `1 anchor + 2 slow movers`

6. Read `Logic` for the full scoring formula and clustering rules.

## Core logic

- `sold_qty_last_4m = SUM(m1:m4)`
- `Movement_Percentile = PERCENTRANK(sold_qty_last_4m)`
- `Slow_Mover_Score = 1 - Movement_Percentile`
- `stock_coverage_months = available_stock / avg_monthly_qty_4m`
- `Stock_Pressure_Percentile = PERCENTRANK(stock_coverage_months)`
- `Contribution_Percentile = PERCENTRANK(cont_from_total)`
- `Bundle_Priority_Score = 0.45*Slow_Mover_Score + 0.35*Stock_Pressure_Percentile + 0.20*Contribution_Percentile`

Anchor and candidate rules:
- Anchor product: high movement + high contribution + available stock
- Candidate product: low movement + available stock + minimum stock coverage

Discount policy (slight by design):
- High Value Bundle: 5%
- Medium Value Bundle: 7%
- Low Value Bundle: 10%
- Add +2% only for very high stock coverage (capped at 12%)

## Regenerate workbook

```bash
python3 create_bundle_workbook.py
```
