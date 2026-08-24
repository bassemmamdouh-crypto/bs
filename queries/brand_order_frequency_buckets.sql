-- Order-frequency buckets per brand section (Metabase, PostgreSQL).
--
-- Optimized rewrite of the original report. Same output; see
-- queries/README.md for the reasoning and the measured speed-up.
--
-- The base CTE is called base_retailers, not retailers: the original name
-- shadowed the physical `retailers` table, which reads as a self-reference.

WITH base_retailers AS (
    -- One row per (retailer, polygon). DISTINCT collapses the duplicates the
    -- address / attribute / warehouse joins produce; the original collapsed
    -- them later in GROUP BY, so this changes nothing but the row count.
    SELECT DISTINCT
        r.id AS retailer_id,
        dp.name_en AS polygon_name,
        CASE WHEN dp.district_id IN (13,14,15,32) THEN NULL ELSE r.id END AS pepsi_retailers
    FROM retailers r
    LEFT JOIN retailer_addresses ra                ON ra.retailer_id = r.id
    LEFT JOIN retailer_dispatching_polygons rdp    ON rdp.retailer_address_id = ra.id
    LEFT JOIN dispatching_polygons dp              ON dp.id = rdp.dispatching_polygon_id
    LEFT JOIN materialized_views.sales_structure w ON w.polygon_name_ar = dp.name_ar
    LEFT JOIN {{#577-agents-structuer}} st         ON st.name_en = dp.name_en
    WHERE r.activation
      AND EXISTS (
          SELECT 1
          FROM retailer_attributes rat
          WHERE rat.retailer_id = r.id
            AND rat.market_type_id NOT IN (5,2)
      )
      [[AND w."Warehouse" = {{area}} ]]
      [[AND st.route = {{route}} ]]
),

-- Historical brand flags per retailer. Aggregating the fact table on its own
-- and joining afterwards avoids replaying every order line once per
-- address/polygon row of its retailer.
--
-- The EXISTS runs against base_retailers rather than a DISTINCT-ed CTE of it:
-- SELECT DISTINCT over a CTE gives the planner a fixed 200-row guess, which
-- was enough to make it pick a per-retailer nested loop over the fact table.
retailer_history AS (
    SELECT
        so.retailer_id,
        bool_or(so.brand_id IN (3,4,5,23,24,25,26,27,28)) AS has_lays,
        bool_or(so.brand_id IN (8,9,10))                  AS has_pepsi,
        bool_or(so.brand_id IN (1,2,6))                   AS has_alyoum,
        bool_or(so.brand_id BETWEEN 11 AND 22)            AS has_maraii
    FROM {{#417-products-sales-order-data}} so
    WHERE EXISTS (
        SELECT 1 FROM base_retailers b WHERE b.retailer_id = so.retailer_id
    )
    GROUP BY so.retailer_id
),

historical_section_flags AS (
    SELECT
        b.polygon_name,
        bool_or(h.has_lays)   AS has_lays_section,
        bool_or(h.has_pepsi)  AS has_pepsi_section,
        bool_or(h.has_alyoum) AS has_alyoum_section,
        bool_or(h.has_maraii) AS has_maraii_section
    FROM base_retailers b
    JOIN retailer_history h ON h.retailer_id = b.retailer_id
    GROUP BY b.polygon_name
),

-- One row per order in the reporting window, so the seven
-- COUNT(DISTINCT sales_order_id) below become plain counts.
--
-- This replaces the original current_orders + order_level pair: the DISTINCT
-- there only removed rows that bool_or and GROUP BY ignore anyway, so it was a
-- full extra dedup pass over the whole date range for nothing.
order_flags AS (
    SELECT
        so.retailer_id,
        so.id AS sales_order_id,
        bool_or(so.brand_id IN (3,4,5,23,24,25,26,27,28)) AS is_lays,
        bool_or(so.brand_id = 8)                          AS is_pepsi,
        bool_or(so.brand_id = 9)                          AS is_aquafina,
        bool_or(so.brand_id = 10)                         AS is_youmy,
        bool_or(so.brand_id IN (1,2,6))                   AS is_alyoum,
        bool_or(so.brand_id BETWEEN 11 AND 22)            AS is_maraii
    FROM {{#417-products-sales-order-data}} so
    -- No retailer restriction here on purpose. The original joined the base
    -- CTE at this point, but `combined` already discards retailers outside the
    -- base universe, so filtering here changes nothing -- and doing it means
    -- the planner drives the fact table by retailer_id, which turns the date
    -- range into millions of single-row heap lookups.
    --
    -- Sargable range predicate: no cast on created_at, so an index on
    -- created_at can be used. Keep this on one line for the [[ ... --]] trick.
    WHERE [[so.created_at >= {{Start_date}}::date AND so.created_at < {{end_date}}::date + 1 --]] so.created_at >= date_trunc('month', current_date)
      [[AND so.warehouse_id = {{warehouse}}]]
    GROUP BY 1, 2
),

-- Frequencies depend only on the retailer, so they are computed once per
-- retailer rather than once per (retailer, polygon) group.
retailer_frequency AS (
    SELECT
        retailer_id,
        count(*)                            AS total_freq,
        count(*) FILTER (WHERE is_lays)     AS lays_freq,
        count(*) FILTER (WHERE is_pepsi)    AS pepsi_freq,
        count(*) FILTER (WHERE is_aquafina) AS aquafina_freq,
        count(*) FILTER (WHERE is_youmy)    AS youmy_freq,
        count(*) FILTER (WHERE is_alyoum)   AS alyoum_freq,
        count(*) FILTER (WHERE is_maraii)   AS maraii_freq
    FROM order_flags
    GROUP BY 1
),

combined AS (
    SELECT
        b.retailer_id,
        b.pepsi_retailers,
        coalesce(hf.has_lays_section,   false) AS has_lays_section,
        coalesce(hf.has_pepsi_section,  false) AS has_pepsi_section,
        coalesce(hf.has_alyoum_section, false) AS has_alyoum_section,
        coalesce(hf.has_maraii_section, false) AS has_maraii_section,
        coalesce(f.total_freq, 0)    AS total_freq,
        coalesce(f.lays_freq, 0)     AS lays_freq,
        coalesce(f.pepsi_freq, 0)    AS pepsi_freq,
        coalesce(f.aquafina_freq, 0) AS aquafina_freq,
        coalesce(f.youmy_freq, 0)    AS youmy_freq,
        coalesce(f.alyoum_freq, 0)   AS alyoum_freq,
        coalesce(f.maraii_freq, 0)   AS maraii_freq
    FROM base_retailers b
    LEFT JOIN historical_section_flags hf ON hf.polygon_name = b.polygon_name
    LEFT JOIN retailer_frequency f        ON f.retailer_id = b.retailer_id
),

-- One pass over combined instead of seven UNION ALL branches over it. The
-- counted_id column carries the original's
--   CASE WHEN brand IN ('PEPSI','AQUAFINA','ALYOUM')
--        THEN pepsi_retailers ELSE retailer_id END
-- so the final aggregate stays a single COUNT(DISTINCT ...).
unpivoted AS (
    SELECT v.brand, v.freq, v.counted_id
    FROM combined c
    CROSS JOIN LATERAL (VALUES
        ('TOTAL',    c.total_freq,    c.retailer_id,     true),
        ('LAYS',     c.lays_freq,     c.retailer_id,     c.has_lays_section),
        ('PEPSI',    c.pepsi_freq,    c.pepsi_retailers, c.has_pepsi_section),
        ('AQUAFINA', c.aquafina_freq, c.pepsi_retailers, c.has_pepsi_section),
        ('YOUMY',    c.youmy_freq,    c.retailer_id,     c.has_pepsi_section),
        ('ALYOUM',   c.alyoum_freq,   c.pepsi_retailers, c.has_alyoum_section),
        ('MARAII',   c.maraii_freq,   c.retailer_id,     c.has_maraii_section)
    ) AS v(brand, freq, counted_id, section_ok)
    WHERE v.section_ok
)

SELECT
    brand,
    CASE
        WHEN freq = 0 THEN '0'
        WHEN freq = 1 THEN '1'
        WHEN freq = 2 THEN '2'
        WHEN freq = 3 THEN '3'
        ELSE '4+'
    END AS frequency_bucket,
    count(DISTINCT counted_id) AS retailer_count
FROM unpivoted
GROUP BY 1, 2
ORDER BY
    CASE brand
        WHEN 'TOTAL'    THEN 1
        WHEN 'LAYS'     THEN 2
        WHEN 'PEPSI'    THEN 3
        WHEN 'AQUAFINA' THEN 4
        WHEN 'YOUMY'    THEN 5
        WHEN 'ALYOUM'   THEN 6
        WHEN 'MARAII'   THEN 7
    END,
    frequency_bucket;
