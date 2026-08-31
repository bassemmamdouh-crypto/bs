# Reporting queries

Faster rewrites of two Metabase native questions, plus the indexes they need.
Both keep the original output; the only intentional difference anywhere is
`total_nmv` in the first report, explained below.

| File | Report |
| --- | --- |
| `monthly_active_retailers_by_brand.sql` | Per month: base retailer universe, brand-active retailers, NMV |
| `brand_order_frequency_buckets.sql` | Retailers per brand section bucketed by order frequency |
| `brand_order_frequency_buckets_prebuilt.sql` | The same report, ~2.5x faster again, reading precomputed section flags |
| `polygon_brand_sections_mv.sql` | Materialized view the `_prebuilt` variant reads |
| `recommended_indexes.sql` | Supporting indexes |

## Benchmark environment

Everything below was measured on PostgreSQL 16.15, 4 vCPU, 15 GB RAM, default
`work_mem` (4 MB) and `shared_buffers` (128 MB), against synthetic data shaped
like the real reports:

- 57M order lines / 19M sales orders — 19 months x 1M orders, ~3 lines per
  order, 3M lines in the current month. `id` repeats across the lines of an
  order, and a brand can repeat within an order.
- 200k retailers, 190k active, 143,333 in the base universe after the
  market-type filter.
- 285k retailer addresses, 250k attribute rows, 400 dispatching polygons, 458
  sales-structure rows (some polygons deliberately appear twice).

Original and rewrite were rendered through the same Metabase template renderer
(card references resolved to tables, `[[ ]]` blocks kept or dropped per
parameter) and run against the same database.

---

## 1. Monthly active retailers by brand

All retailer count columns return exactly the same numbers as the original.

### Why the original was slow

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
   the column, so no index on `created_at` can be used.
5. **Work that never reaches the output.** The `districts` join,
   `polygon_name` and `district_name` are selected but unused, and the
   `sales_structure` and agent-structure joins run once per retailer address
   rather than once per polygon.

### What the rewrite does

- Aggregates orders to **one row per retailer per month** first, turning the
  brand columns into cheap `bool_or(...)` flags and a single `sum(...)`.
- Reduces `base_retailers` to **one row per retailer** with `GROUP BY r.id`, so
  nothing downstream fans out.
- Replaces the nine `COUNT(DISTINCT ...)` with `count(*) FILTER (WHERE flag)`.
- Counts the base universe **once** instead of once per month.
- Replaces the `retailer_attributes` join with `EXISTS`, same filter semantics
  without duplicating retailers.
- Resolves warehouse and route **per polygon** instead of per address.
- Drops the unused `districts` join and unused output columns.
- Makes the date filter sargable:
  `created_at >= {{Start_date}}::date AND created_at < {{end_date}}::date + 1`.
- Derives `months` from the already-aggregated orders instead of scanning the
  fact table a second time.

### Measured

| Scenario | Original | Rewrite | Rewrite, `DISTINCT` dropped |
| --- | --- | --- | --- |
| Current month, no extra indexes | 13.9 s | 5.2 s | 4.5 s |
| Current month, with indexes | 13.7 s | **4.1 s** | **2.5 s** |
| 4-month range, no extra indexes | 17.0 s | 15.2 s | 4.6 s |
| 4-month range, with indexes | 16.9 s | 13.4 s | **4.4 s** |

The original gets no benefit from the indexes because of the `::date` cast: its
plan stays a sequential scan discarding 54M of 57M rows, while the rewrite
switches to `Index Only Scan ... Heap Fetches: 0`.

### Verified

Retailer count columns matched exactly with no parameters, with a 4-month
`Start_date`/`end_date` range, with `area` set, with `route` set, and with a
filter combination matching no retailer (both return 0 rows).

### Behaviour change: `total_nmv`

`total_nmv` is now lower. The original sums revenue across the fanned-out
`base_retailers` rows, so a retailer's orders are counted once per
address/attribute/warehouse row it has. The rewrite sums each order line once.
Every retailer count column is unchanged.

### Optional: drop the `DISTINCT`

`SELECT DISTINCT retailer_id, created_at, total_price, brand_id` is the most
expensive remaining step: it sorts every order line in the range and spills to
disk. It is kept only for exact parity with the original.

It also silently merges two genuinely different order lines when a retailer
buys two products of the same brand at the same price in one order. Check
whether it removes anything:

```sql
SELECT count(*) - count(DISTINCT (retailer_id, created_at, total_price, brand_id))
           AS rows_removed_by_distinct
FROM   <sales order source>
WHERE  created_at >= date_trunc('month', current_date);
```

If that returns `0`, delete the `DISTINCT` keyword in `orders_data`: identical
output, and the 4-month range gets about 3x faster again.

---

## 2. Brand order-frequency buckets

Output is byte-identical to the original in every parameter combination tested.

### Why the original was slow

1. **`historical_section_flags` replayed the whole fact table per address row.**
   It joins every order line ever (no date bound) to the base CTE, which has
   several rows per retailer, and only then reduces to one row per polygon with
   `MAX(CASE ...)`. On the benchmark data that join produced tens of millions of
   rows to compute 385 output rows.
2. **`current_orders` joined, then `order_level` undid the join.** The join to
   the base CTE multiplied each order line by the retailer's address/polygon
   row count, and `SELECT DISTINCT` then removed exactly those duplicates.
3. **Frequencies were computed once per (retailer, polygon).** A retailer in two
   polygons had its seven `COUNT(DISTINCT sales_order_id)` evaluated twice, over
   identical input, since `order_level` joins on `retailer_id` alone.
4. **Seven `COUNT(DISTINCT ...)` per group** over the un-aggregated order rows.
5. **`combined` was scanned seven times** by the `UNION ALL` unpivot.
6. **The date predicate was not sargable**, as in the first report.
7. **The base CTE was named `retailers`**, shadowing the physical `retailers`
   table it selects from. Postgres resolves this the way the author intended
   (a non-recursive CTE is not visible inside its own definition), but it reads
   as a self-reference and is one keyword away from breaking.

### What the rewrite does

- Aggregates the fact table's history to **one row per retailer** first, then
  joins that to the retailer/polygon map to get the per-polygon section flags.
  Same flags, without replaying order lines per address row.
- Aggregates the reporting window to **one row per order**, so the seven
  `COUNT(DISTINCT sales_order_id)` become `count(*) FILTER (WHERE flag)`.
- Drops the `current_orders`/`order_level` dedup pass entirely: `DISTINCT`
  there only removed rows that `bool_or` and `GROUP BY` ignore anyway.
- Computes frequencies **once per retailer**, then joins to the (retailer,
  polygon) rows.
- Collapses to **one row per retailer** before the unpivot, so the final
  aggregate is a plain `count(*) FILTER (...)` instead of a `COUNT(DISTINCT)`
  that had to sort every (retailer, polygon, brand) combination. See
  [the zero-count subtlety](#the-zero-count-subtlety) for what that has to
  preserve.
- Deduplicates the base CTE with `DISTINCT`, which the original's `GROUP BY`
  did later anyway.
- Replaces the seven-branch `UNION ALL` with one `CROSS JOIN LATERAL (VALUES
  ...)`, so the rows are scanned once.
- Makes the date filter sargable, and renames the base CTE to
  `base_retailers`.
- Adds `frequency_bucket` to the `ORDER BY`. The original leaves ordering
  within a brand unspecified (that part is commented out); the buckets sort as
  text into the intended `0, 1, 2, 3, 4+` order.

### Measured

All with the recommended indexes. Run-to-run variation is around 10%.

| Scenario | Original | Rewrite | Rewrite + precomputed flags |
| --- | --- | --- | --- |
| Default (current month) | 21.9 s | 8.9 s | **3.6 s** |
| 4-month range | 26.1 s | 14.0 s | **8.7 s** |
| `area` set | 12.5 s | 3.8 s | **2.7 s** |
| `route` set | 2.3 s | 2.7 s | 2.4 s |
| `warehouse` set | 19.1 s | 7.6 s | **1.8 s** |
| `area` + `warehouse` | 5.0 s | 2.3 s | **0.8 s** |

With the fact-table indexes dropped, the first two columns run 23.7 / 27.3 /
21.0 / 19.6 / 21.6 / 13.3 s and 20.5 / 24.5 / 10.2 / 6.8 / 19.4 / 7.0 s.
**Without precomputed flags this report needs
`idx_sales_order_retailer_brand`**: the section flags come from the entire
order history, so otherwise both versions spend most of their time reading the
whole table.

### Precomputed section flags

`brand_order_frequency_buckets_prebuilt.sql` reads
`materialized_views.polygon_brand_sections` (built by
`polygon_brand_sections_mv.sql`) instead of deriving the section flags inline.

That step is worth removing because it reads **every order line ever recorded**
on every run — no date bound — to produce a few hundred rows. In the default
scenario it was 10.8 s of a 16 s plan: a full 57M-row index-only scan feeding a
merge join that emitted 39.6M rows.

The flags also do not depend on any report parameter. `area` and `route` filter
which polygons survive, not which retailers feed a surviving polygon's flags,
so one view serves every parameter combination — confirmed by the equivalence
run below, where the filtered scenarios match exactly.

The cost is staleness: a polygon recording its very first order of a brand
section, or a retailer newly linked to a polygon, shows up only after the next
refresh. The flags are monotonic and slow moving, so a nightly
`REFRESH MATERIALIZED VIEW CONCURRENTLY` is normally enough. Building the view
from scratch took 6 s on the benchmark data. If you need always-live flags, use
`brand_order_frequency_buckets.sql`.

### Verified

Output compared row for row (sorted, since the original's within-brand order is
unspecified) with no parameters, a 4-month range, `area` set, `route` set,
`warehouse` set, and `area` + `warehouse` together. Both rewrites are identical
to the original in all six, including the intermediate `1`/`2`/`3` buckets and
the zero-count row described next.

### The zero-count subtlety

`PEPSI`, `AQUAFINA` and `ALYOUM` count `pepsi_retailers`, which is NULL for
districts 13/14/15/32, and `COUNT(DISTINCT ...)` skips NULLs. So the original
can emit a bucket row whose count is `0`: the group exists because in-section
rows landed in that bucket, but every one of them had a NULL id. Collapsing to
one row per retailer drops such groups unless membership and counting are
tracked separately, which is why `retailer_eligibility` carries both a
`*_section` and a `*_counts` flag. The `area` scenario produces exactly this
case (`ALYOUM | 2 | 0`), and it caught the first version of the collapse.

### Optional server-side settings

`jit = off` and a larger `work_mem` on the Metabase role are worth about 17% on
the default scenario (3.55 s to 2.96 s) and nothing measurable on the 4-month
range. Not required, and only worth doing if the same settings suit the rest of
the workload.

### A note on the fact-table semi-join

The rewrite deliberately does **not** restrict the reporting-window aggregate to
base retailers, even though the original joined the base CTE there. `combined`
already discards retailers outside the base universe, so the filter is
redundant — and adding it makes things dramatically worse: `SELECT DISTINCT`
over a CTE gives the planner a fixed 200-row estimate, which was enough for it
to drive the fact table by `retailer_id` and turn the date range into 13.6M
single-row heap lookups (57 s instead of 15 s on the 4-month range).

The historical aggregate keeps its `EXISTS` against `base_retailers`, because
there the per-retailer path is index-only and is what makes the filtered
scenarios fast.
