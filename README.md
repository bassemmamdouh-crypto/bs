# Delivery Run Routing Planner

This project contains a Python script to generate:

1. **Run sheet**: order-level assignment to delivery runs.
2. **Load summary**: product quantities to load per run/vehicle.

The routing logic is built around your requested rules:

- Try to keep each run in **one polygon** first.
- If a vehicle still has free capacity, add the **closest retailers/orders**.
- Vehicle capacity is configurable per **supply chain**.
- Output is an Excel file with two sheets: `run_sheet` and `load_summary`.

## Input expectations

You provide an Excel file with order lines that include at least:

- Order ID
- Retailer ID
- Retailer name
- Retailer latitude
- Retailer longitude
- Product
- Quantity
- Supply chain

Column names are configurable in JSON.

## Setup

Python 3.10+ is recommended.

Install dependencies:

```bash
python3 -m pip install pandas openpyxl
```

## Configuration

Use `config/routing_config.example.json` as a template.

Key sections:

- `columns`: maps your Excel column names.
- `routing.default_vehicle_capacity`: fallback capacity.
- `routing.capacity_by_supply_chain`: supply-chain-specific vehicle capacities.
- `routing.allow_cross_polygon_fill`: allow filling runs with orders from other polygons.
- `routing.max_cross_fill_distance_km`: optional max distance when cross-filling.
- `polygons`: list of polygons, each with `id` and `points` (`[lat, lon]`).

## Run

```bash
python3 delivery_routing/run_planner.py \
  --input path/to/orders.xlsx \
  --config config/routing_config.example.json \
  --output path/to/routing_output.xlsx
```

Optional:

- `--sheet SHEET_NAME` to override the configured input sheet.
- `--settings path/to/run_planner_settings.json` to load input/config/output from a JSON file.

## Jupyter / Notebook mode (no CLI args required)

If your notebook and script are in the same folder, you can avoid argument issues by
using a settings file.

1) Copy `delivery_routing/run_planner_settings.example.json` to:

`delivery_routing/run_planner_settings.json`

2) Edit values, for example:

```json
{
  "input": "E:/Marbah Products/Other Scripts/Delivery Plan/orders_in_delivery.xlsx",
  "config": "E:/Marbah Products/Other Scripts/Delivery Plan/config/routing_config.example.json",
  "output": "E:/Marbah Products/Other Scripts/Delivery Plan/routing_output.xlsx",
  "sheet": null
}
```

3) Run from notebook:

```python
%run delivery_routing/run_planner.py
```

Or explicitly pass settings path:

```python
%run delivery_routing/run_planner.py --settings "E:/Marbah Products/Other Scripts/Delivery Plan/delivery_routing/run_planner_settings.json"
```

Notes:
- In your previous command, config path missed a slash: `configrouting...` should be `config/routing...`.
- Prefer `/` in notebook paths on Windows to avoid escaping issues.

## Output sheets

### 1) `run_sheet`

Contains original order lines plus assignment fields such as:

- `run_id`
- `stop_sequence`
- `vehicle_capacity`
- `used_capacity`
- `remaining_capacity`
- `over_capacity`
- `primary_polygon`
- `mixed_polygons`

### 2) `load_summary`

Aggregated by:

- `run_id`
- `supply_chain`
- `product`

With:

- `total_quantity`

This is the summary sheet used for loading each vehicle with product quantities.
