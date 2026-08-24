# Monthly active-retailer report

`monthly_active_retailers_by_brand.sql` is a faster rewrite of the Metabase
question that reports, per month, the base retailer universe, how many of those
retailers bought each brand group, and total NMV.

All retailer count columns return exactly the same numbers as the original
query. The only intentional change in output is `total_nmv` (see
[Behaviour changes](#behaviour-changes)).

## Why the original was slow

1. **`CROSS JOIN` before aggregation.** `months CROSS JOIN base_retailers LEFT
   JOIN orders_data` builds one row per month per base-retailer row per order
   line, then collapses it with nine `COUNT(DISTINCT ...)` passes. Every extra
   month multiplies the whole intermediate result.
2. **`base_retailers` returns several rows per retailer.** A retailer with two
   addresses, two attribute rows and a polygon that appears twice in
   `sales_structure` contributes eight rows. `COUNT(DISTINCT ...)` hides the
   duplication but the planner still has to carry all of it.
3. **Nine `COUNT(DISTINCT ...)` on the widest intermediate.** Postgres cannot
   share work between distinct-aggregates, so each one is its own sort or hash
   over the fanned-out rows.
4. **The date predicate was not sargable.** `so.created_at::date >= ...` casts
   the column, so no index on `created_at` can be used. In the benchmark below
   the original scans all 60M rows and throws away 57M of them, even with the
   index in place.
5. **Work that never reaches the output.** The `districts` join, `polygon_name`
   and `district_name` are selected but unused, and the `sales_structure` and
   agent-structure joins run once per retailer address rather than once per
   polygon.

## What the rewrite does

- Aggregates orders to **one row per retailer per month** first, turning the
  brand columns into cheap `bool_or(...)` flags and a single `sum(...)`.
- Reduces `base_retailers` to **one row per retailer** with `GROUP BY r.id`, so
  nothing downstream fans out.
- Replaces the nine `COUNT(DISTINCT ...)` with `count(*) FILTER (WHERE flag)`
  over the pre-aggregated rows.
- Counts the base universe **once** instead of once per month, and joins the
  monthly numbers back by month, so adding months no longer multiplies work.
- Replaces the `retailer_attributes` join with `EXISTS`, which keeps the same
  filter semantics without duplicating retailers.
- Resolves warehouse and route **per polygon** (a few hundred rows) instead of
  per retailer address.
- Drops the unused `districts` join and the unused output columns.
- Makes the date filter sargable:
  `created_at >= {{Start_date}}::date AND created_at < {{end_date}}::date + 1`
  instead of `created_at::date BETWEEN ...`, which covers the same days and can
  use an index.
- Derives `months` from the already-aggregated orders instead of scanning the
  fact table a second time.

## Measured results

PostgreSQL 16.15, 4 vCPU, 15 GB RAM, default `work_mem` (4 MB). Synthetic data
shaped like the real report: 60M order lines (3M in the current month, 18
months of history), 200k retailers, 285k addresses, 250k attribute rows, 400
dispatching polygons.

| Scenario | Original | Rewrite | Rewrite, `DISTINCT` dropped |
| --- | --- | --- | --- |
| Current month, no extra indexes | 14.4 s | 5.4 s | 4.6 s |
| Current month, with the indexes below | 14.8 s | **3.9 s** | **2.4 s** |
| 4-month range, no extra indexes | 19.4 s | 16.9 s | 5.0 s |
| 4-month range, with the indexes below | 19.0 s | 14.4 s | **5.0 s** |

The original gets no benefit at all from the indexes because of the `::date`
cast: its plan stays a sequential scan (`Rows Removed by Filter: 57000000`),
while the rewrite switches to `Index Only Scan ... Heap Fetches: 0` and reads
only the reporting month.

## Verified against the original

Both queries were rendered through the same Metabase template renderer and run
against the same database. The retailer count columns matched exactly in every
parameter combination:

| Parameters | Result |
| --- | --- |
| none (defaults to the current month) | identical counts |
| `Start_date` / `end_date` spanning 4 months | identical counts, 4 rows |
| `area` set | identical counts |
| `route` set | identical counts |
| `area` + `route` matching no retailer | both return 0 rows |

## Behaviour changes

`total_nmv` is now lower than the original. In the original, revenue is summed
across the fanned-out `base_retailers` rows, so a retailer's orders are counted
once per address/attribute/warehouse row it has. On the benchmark data the
original reports `399,499,082.75` against a true value of `262,816,480.25` — a
1.52x overstatement. The rewrite sums each order line once. Every retailer
count column is unchanged.

## Optional: drop the `DISTINCT`

`SELECT DISTINCT retailer_id, created_at, total_price, brand_id` is the single
most expensive remaining step: it sorts every order line in the range and
spills to disk. It is kept only so the numbers match the original exactly.

It also silently merges two genuinely different order lines when the same
retailer buys two products of the same brand at the same price in the same
order, which understates NMV. Check whether it removes anything at all:

```sql
SELECT count(*) - count(DISTINCT (retailer_id, created_at, total_price, brand_id))
           AS rows_removed_by_distinct
FROM   <sales order source>
WHERE  created_at >= date_trunc('month', current_date);
```

If that returns `0`, delete the `DISTINCT` keyword in `orders_data`: the output
is identical and the 4-month range gets roughly 3x faster again.

## Indexes

See `recommended_indexes.sql`. The important one is the covering index on the
fact table's `created_at`, which is what turns the month filter into an
index-only scan.
