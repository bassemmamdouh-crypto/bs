-- Indexes that support the queries in this directory.
-- Create them on the physical tables (the sales-order indexes belong on the
-- table behind Metabase card #417, not on the card itself).
--
-- CONCURRENTLY keeps the tables writable while the index is built; it cannot
-- run inside a transaction block, so run this file with psql, not as one
-- statement in a BI tool.

-- Reporting-window access path, used by both reports. Lets the selected month
-- be read with an index-only scan instead of a sequential scan of the whole
-- fact table. Wide, because it covers every column the reports read: expect it
-- to be roughly the size of the table itself. Drop `warehouse_id` if the
-- frequency report's warehouse filter is never used, and `total_price` if the
-- NMV report is not deployed.
CREATE INDEX CONCURRENTLY IF NOT EXISTS idx_sales_order_created_at_cover
    ON products_sales_order_data (created_at)
    INCLUDE (retailer_id, id, brand_id, total_price, warehouse_id);

-- If the fact table is append-only and physically ordered by created_at, a
-- BRIN index is a far smaller alternative to the btree above, at the cost of
-- reading the heap:
-- CREATE INDEX CONCURRENTLY IF NOT EXISTS idx_sales_order_created_at_brin
--     ON products_sales_order_data USING brin (created_at) WITH (pages_per_range = 32);

-- brand_order_frequency_buckets.sql derives its "section" flags from the whole
-- order history, with no date bound. This index is what makes that step an
-- index-only scan of a narrow structure instead of a full heap scan, and it is
-- the single biggest win for that report.
CREATE INDEX CONCURRENTLY IF NOT EXISTS idx_sales_order_retailer_brand
    ON products_sales_order_data (retailer_id, brand_id);

-- Base-universe joins.
CREATE INDEX CONCURRENTLY IF NOT EXISTS idx_retailer_addresses_retailer
    ON retailer_addresses (retailer_id);

CREATE INDEX CONCURRENTLY IF NOT EXISTS idx_retailer_dispatching_polygons_address
    ON retailer_dispatching_polygons (retailer_address_id);

-- Supports the EXISTS market-type check.
CREATE INDEX CONCURRENTLY IF NOT EXISTS idx_retailer_attributes_retailer_market
    ON retailer_attributes (retailer_id, market_type_id);

-- Polygon lookups by name.
CREATE INDEX CONCURRENTLY IF NOT EXISTS idx_sales_structure_polygon_name_ar
    ON materialized_views.sales_structure (polygon_name_ar);
