-- Order-frequency buckets per brand section (Metabase, PostgreSQL).
--
-- Same report as brand_order_frequency_buckets.sql, but reading the section
-- flags from materialized_views.polygon_brand_sections instead of deriving
-- them from the whole order history on every run. Create that view with
-- polygon_brand_sections_mv.sql and refresh it on a schedule.
--
-- The two files are otherwise identical. Use the other one if you need the
-- flags to always reflect the live tables.

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

-- One row per order in the reporting window, so the seven
-- COUNT(DISTINCT sales_order_id) below become plain counts.
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
    -- Sargable range predicate: no cast on created_at, so an index on
    -- created_at can be used. Keep this on one line for the [[ ... --]] trick.
    WHERE [[so.created_at >= {{Start_date}}::date AND so.created_at < {{end_date}}::date + 1 --]] so.created_at >= date_trunc('month', current_date)
      [[AND so.warehouse_id = {{warehouse}}]]
    GROUP BY 1, 2
),

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

retailer_eligibility AS (
    SELECT
        b.retailer_id,
        bool_or(hf.has_lays_section)   AS lays_section,
        bool_or(hf.has_pepsi_section)  AS pepsi_section,
        bool_or(hf.has_alyoum_section) AS alyoum_section,
        bool_or(hf.has_maraii_section) AS maraii_section,
        bool_or(hf.has_pepsi_section  AND b.pepsi_retailers IS NOT NULL) AS pepsi_counts,
        bool_or(hf.has_alyoum_section AND b.pepsi_retailers IS NOT NULL) AS alyoum_counts
    FROM base_retailers b
    LEFT JOIN materialized_views.polygon_brand_sections hf
        ON hf.polygon_name = b.polygon_name
    GROUP BY b.retailer_id
),

unpivoted AS (
    SELECT v.brand, v.freq, v.counts
    FROM retailer_eligibility e
    LEFT JOIN retailer_frequency f ON f.retailer_id = e.retailer_id
    CROSS JOIN LATERAL (VALUES
        ('TOTAL',    coalesce(f.total_freq, 0),    true,             true),
        ('LAYS',     coalesce(f.lays_freq, 0),     e.lays_section,   true),
        ('PEPSI',    coalesce(f.pepsi_freq, 0),    e.pepsi_section,  e.pepsi_counts),
        ('AQUAFINA', coalesce(f.aquafina_freq, 0), e.pepsi_section,  e.pepsi_counts),
        ('YOUMY',    coalesce(f.youmy_freq, 0),    e.pepsi_section,  true),
        ('ALYOUM',   coalesce(f.alyoum_freq, 0),   e.alyoum_section, e.alyoum_counts),
        ('MARAII',   coalesce(f.maraii_freq, 0),   e.maraii_section, true)
    ) AS v(brand, freq, section_ok, counts)
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
    count(*) FILTER (WHERE counts) AS retailer_count
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
