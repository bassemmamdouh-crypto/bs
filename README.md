# bs

## `retailer_agent_daily_target_vs_achievement.sql`

Two views that answer, for every day a retailer was dispatched to an agent: how
much target that retailer should have hit by that day, and how much of it was
actually achieved.

- `retailer_agent_daily_target_vs_achievement` — one row per dispatch day, per
  retailer, per agent.
- `agent_daily_target_vs_achievement` — the same measures rolled up to one row
  per agent per day.

Fridays are not working days. The monthly target is spread evenly over the
non-Friday days of the month, so a Friday adds no target and carries the same
target-to-date as the Thursday before it. Sales made on a Friday still count
towards the achievement.

    daily_target   = month_target / working_days_in_month
    target_to_date = daily_target * working_days_elapsed

Only section 1 (`SOURCES`) of the file is schema specific: point `dispatch_raw`,
`retailer_targets`, `brand_scope` and `retailer_brand_daily_sales` at the real
objects in your warehouse.

### Verifying it

    pip install -r requirements-dev.txt
    python3 tests/test_daily_target_vs_achievement.py

The test executes the `.sql` file verbatim against a fixture warehouse in
DuckDB and compares every output row with an independent Python implementation
of the working-day calendar, then checks that Snowflake and DuckDB parse the
file identically.
