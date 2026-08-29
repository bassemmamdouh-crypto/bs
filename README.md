# Area item sales targets

Turn an Excel dump of **purchasing behavior by area** into **next-month item targets**.

For each area the planner takes the last 3 months of sales, measures every item's share of that area's total, and applies the same mix to the area's next-month total.

```
item next-month target = area next-month total × (item sales last 3 months / area sales last 3 months)
```

Area next-month totals default to the **average of the months actually present** in that window (if the file only has June and July, the mix and the average use those two months). Put a commercial number in `Area_Targets.next_month_target` and rerun with `--area-targets` to lock the mix to a different headline.

Arabic area workbooks are supported: `المنطقة`, `الصنف (ROUT)`, `العبوة (Container)`, and month columns like `شهر 6` / `شهر 7`. The all-areas sheet is preferred when several tabs exist. Each item is keyed as `container | product`. `Cost_Center_Targets` splits the area/item target by cost-center share of that item.

## Install

```bash
pip install -r requirements.txt
```

## Run

```bash
python create_sample_purchasing_workbook.py
python area_item_target_planner.py sample_purchasing_behavior.xlsx \
  --sheet Purchasing_Long \
  -o area_item_targets.xlsx
```

Optional: override area totals from a second sheet/file that has `area` and `next_month_target`.

```bash
python area_item_target_planner.py sample_purchasing_behavior.xlsx \
  --area-targets area_item_targets.xlsx \
  -o area_item_targets.xlsx
```

## Input Excel

Long format:

| area | sku | item | month | sales | qty |
| Baghdad | P330 | Pepsi 330ml | 2026-07 | 125000 | 250 |

Wide format:

| area | sku | item | 2026-05 | 2026-06 | 2026-07 |
| Baghdad | P330 | Pepsi 330ml | 120000 | 130000 | 125000 |

`qty` is optional. When it is present, a quantity target is inferred from the 3-month average selling price.

Column names are matched loosely (`area` / `polygon` / `city`, `sales` / `net_amount` / `revenue`, `May 2026`, …).

## Output Excel

- `Instructions` — what the numbers mean
- `Area_Targets` — headline total per area (edit `next_month_target` to override)
- `Item_Targets` — mix + sales/qty targets for every area/item
- `Monthly_Mix` — each of the last 3 months, for audit
- one sheet per area — same item targets, filtered

`weighted_contribution` is the mix used for targeting (share of 3-month sales). `simple_avg_contribution` is the unweighted average of the three monthly shares, so a thin month is visible.

## Tests

```bash
python -m pytest tests/test_area_item_target_planner.py -q
```
