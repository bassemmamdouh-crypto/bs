# Bundle Sales Planner

This repository now includes an Excel template to help you create bundles for unmoved products using 6 months of sales data.

## Files

- `bundle_planning_template.xlsx`: Ready-to-use workbook with input, scoring, bundle recommendation, and logic sheets.
- `create_bundle_workbook.py`: Script that generates/regenerates the workbook.

## How to use

1. Open `bundle_planning_template.xlsx`.
2. In `Input_Data`, paste your product data:
   - Product ID, Product Name, Category
   - Purchased quantity for months M1-M6
   - Contribution % of total sold amount (as percentage)
3. Review `Scoring_Model`:
   - Movement and contribution percentiles are auto-calculated.
   - Each product gets a priority score and one of:
     - High Value Bundle
     - Medium Value Bundle
     - Low Value Bundle
4. Go to `Bundle_Recommendations`:
   - Focus on rows with `Low Movement`.
   - Use suggested anchor products and discount guidance.
5. Read `Logic` for the full scoring formula and clustering rules.

## Regenerate workbook

```bash
python3 create_bundle_workbook.py
```
