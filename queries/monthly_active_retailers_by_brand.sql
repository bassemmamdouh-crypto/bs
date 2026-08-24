-- Monthly base vs. brand-active retailers and NMV (Metabase, PostgreSQL).
--
-- Optimized rewrite of the original report. See queries/README.md for the
-- reasoning, the measured speed-up and the recommended indexes.
--
-- Behaviour notes:
--   * All retailer count columns return exactly the same numbers as the
--     original query.
--   * total_nmv no longer multiplies revenue by the number of
--     address/polygon/attribute rows a retailer happens to have, so it is
--     lower than (and more correct than) the original figure.

WITH orders_data AS (
    SELECT DISTINCT
        so.retailer_id,
        so.created_at,
        so.total_price,
        so.brand_id
    FROM {{#417-products-sales-order-data}} so
    -- Sargable range predicate: no cast on created_at, so an index on
    -- created_at can be used. Keep this on one line for the [[ ... --]] trick.
    WHERE [[so.created_at >= {{Start_date}}::date AND so.created_at < {{end_date}}::date + 1 --]] so.created_at >= date_trunc('month', current_date)
),

-- One row per retailer per month, so the per-brand flags are computed once
-- instead of nine COUNT(DISTINCT ...) passes over a fanned-out join.
orders_agg AS (
    SELECT
        o.retailer_id,
        date_trunc('month', o.created_at)::date AS month,
        bool_or(o.brand_id IN (3,4,5,23,24,25,26,27,28)) AS has_lays,
        bool_or(o.brand_id = 8)                          AS has_pepsi,
        bool_or(o.brand_id = 9)                          AS has_aquafina,
        bool_or(o.brand_id = 10)                         AS has_youmy,
        bool_or(o.brand_id IN (1,2,6))                   AS has_alyoum,
        bool_or(o.brand_id BETWEEN 11 AND 22)            AS has_maraii,
        sum(o.total_price)                               AS nmv
    FROM orders_data o
    GROUP BY 1, 2
),

months AS (
    SELECT DISTINCT month
    FROM orders_agg
),

-- dispatching_polygons is small, so resolve warehouse/route once per polygon
-- instead of once per retailer address.
polygons AS (
    SELECT
        dp.id,
        w."Warehouse" AS area,
        st.route,
        coalesce(dp.district_id NOT IN (13,14,15,32), false) AS pepsi_base
    FROM dispatching_polygons dp
    LEFT JOIN materialized_views.sales_structure w ON w.polygon_name_ar = dp.name_ar
    LEFT JOIN {{#577-agents-structuer}} st ON st.name_en = dp.name_en
),

-- Exactly one row per retailer in the base universe.
base_retailers AS (
    SELECT
        r.id AS retailer_id,
        coalesce(bool_or(p.pepsi_base), false) AS pepsi_base
    FROM retailers r
    LEFT JOIN retailer_addresses ra             ON ra.retailer_id = r.id
    LEFT JOIN retailer_dispatching_polygons rdp ON rdp.retailer_address_id = ra.id
    LEFT JOIN polygons p                        ON p.id = rdp.dispatching_polygon_id
    WHERE r.activation
      AND EXISTS (
          SELECT 1
          FROM retailer_attributes rat
          WHERE rat.retailer_id = r.id
            AND rat.market_type_id NOT IN (5,2)
      )
      [[AND p.area = {{area}} ]]
      [[AND p.route = {{route}} ]]
    GROUP BY r.id
),

-- The base universe is month independent, so count it once.
base_counts AS (
    SELECT
        count(*)                           AS total_retailers,
        count(*) FILTER (WHERE pepsi_base) AS pepsi_total_base
    FROM base_retailers
),

monthly AS (
    SELECT
        o.month,
        count(*) FILTER (WHERE o.has_lays)     AS lays_active_retailers,
        count(*) FILTER (WHERE o.has_pepsi)    AS pepsi_active_retailers,
        count(*) FILTER (WHERE o.has_aquafina) AS aquafina_active_retailers,
        count(*) FILTER (WHERE o.has_youmy)    AS youmy_active_retailers,
        count(*) FILTER (WHERE o.has_alyoum)   AS alyoum_active_retailers,
        count(*) FILTER (WHERE o.has_maraii)   AS maraii_active_retailers,
        count(*)                               AS total_active_retailers,
        sum(o.nmv)                             AS total_nmv
    FROM orders_agg o
    JOIN base_retailers b ON b.retailer_id = o.retailer_id
    GROUP BY o.month
)

SELECT
    m.month,
    c.total_retailers,
    c.pepsi_total_base,
    coalesce(mo.lays_active_retailers, 0)     AS lays_active_retailers,
    coalesce(mo.pepsi_active_retailers, 0)    AS pepsi_active_retailers,
    coalesce(mo.aquafina_active_retailers, 0) AS aquafina_active_retailers,
    coalesce(mo.youmy_active_retailers, 0)    AS youmy_active_retailers,
    coalesce(mo.alyoum_active_retailers, 0)   AS alyoum_active_retailers,
    coalesce(mo.maraii_active_retailers, 0)   AS maraii_active_retailers,
    coalesce(mo.total_active_retailers, 0)    AS total_active_retailers,
    mo.total_nmv
FROM months m
CROSS JOIN base_counts c
LEFT JOIN monthly mo ON mo.month = m.month
ORDER BY 1;
