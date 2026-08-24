-- Pre-computed "has this polygon ever sold this brand section" flags, for
-- brand_order_frequency_buckets_prebuilt.sql.
--
-- WHY THIS EXISTS
-- The frequency report derives its section flags from the entire order history
-- with no date bound, which means reading every order line ever recorded on
-- every run. That is the single largest cost in the report and it produces only
-- a few hundred rows. Those rows also do not depend on any report parameter:
-- `area` and `route` filter which polygons survive, not which retailers feed a
-- surviving polygon's flags, so the same table serves every parameter
-- combination.
--
-- TRADE-OFF
-- The flags are a snapshot. A polygon that records its very first order of a
-- brand section, or a retailer newly linked to a polygon, is reflected only
-- after the next refresh. The flags are monotonic and slow moving ("has ever
-- sold"), so a nightly refresh is normally enough. If you need
-- always-current flags instead, use brand_order_frequency_buckets.sql, which
-- computes them inline.
--
-- Replace products_sales_order_data with the table behind Metabase card #417.

-- IF NOT EXISTS so re-running this file is a no-op. To change the definition,
-- drop it explicitly first:
--   DROP MATERIALIZED VIEW materialized_views.polygon_brand_sections;
-- The report returns no section flags while the view is missing, so do that
-- outside reporting hours.
CREATE MATERIALIZED VIEW IF NOT EXISTS materialized_views.polygon_brand_sections AS
WITH base AS (
    SELECT DISTINCT
        r.id AS retailer_id,
        dp.name_en AS polygon_name
    FROM retailers r
    LEFT JOIN retailer_addresses ra             ON ra.retailer_id = r.id
    LEFT JOIN retailer_dispatching_polygons rdp ON rdp.retailer_address_id = ra.id
    LEFT JOIN dispatching_polygons dp           ON dp.id = rdp.dispatching_polygon_id
    WHERE r.activation
      AND EXISTS (
          SELECT 1
          FROM retailer_attributes rat
          WHERE rat.retailer_id = r.id
            AND rat.market_type_id NOT IN (5,2)
      )
),
history AS (
    SELECT
        so.retailer_id,
        bool_or(so.brand_id IN (3,4,5,23,24,25,26,27,28)) AS has_lays,
        bool_or(so.brand_id IN (8,9,10))                  AS has_pepsi,
        bool_or(so.brand_id IN (1,2,6))                   AS has_alyoum,
        bool_or(so.brand_id BETWEEN 11 AND 22)            AS has_maraii
    FROM products_sales_order_data so
    WHERE EXISTS (
        SELECT 1 FROM base b WHERE b.retailer_id = so.retailer_id
    )
    GROUP BY so.retailer_id
)
SELECT
    b.polygon_name,
    bool_or(h.has_lays)   AS has_lays_section,
    bool_or(h.has_pepsi)  AS has_pepsi_section,
    bool_or(h.has_alyoum) AS has_alyoum_section,
    bool_or(h.has_maraii) AS has_maraii_section
FROM base b
JOIN history h ON h.retailer_id = b.retailer_id
GROUP BY b.polygon_name;

-- Required for REFRESH ... CONCURRENTLY.
CREATE UNIQUE INDEX IF NOT EXISTS polygon_brand_sections_polygon_name
    ON materialized_views.polygon_brand_sections (polygon_name);

-- Refresh on a schedule (nightly is plenty). CONCURRENTLY keeps the report
-- readable while the refresh runs, at the cost of taking longer.
-- REFRESH MATERIALIZED VIEW CONCURRENTLY materialized_views.polygon_brand_sections;
