create or replace view retailer_bucket_month_start_vs_now as
with
/* Map SALES_OPS polygons to sales structure (best retailer coverage per polygon). */
polygons as (
    select
        sales_ops_polygon_name_en,
        polygon_name_en,
        district_name_en,
        city_name_en,
        region_name_en,
        agent_id,
        partner_integration_id,
        agent_name,
        final_polygon
    from (
        select
            rpso.sales_ops_polygon_name_en,
            rp.polygon_name_en,
            d.name_en as district_name_en,
            c.name_en as city_name_en,
            r.name_en as region_name_en,
            st.agent_id,
            a.partner_integration_id,
            a.name as agent_name,
            case
                when lower(r.name_en) like '%delta%'
                  or lower(r.name_en) like '%upper%'
                    then coalesce(st.cluster, rpso.sales_ops_polygon_name_en)
                else rpso.sales_ops_polygon_name_en
            end as final_polygon,
            count(distinct rpso.retailer_id) as retailers,
            row_number() over (
                partition by rpso.sales_ops_polygon_name_en
                order by count(distinct rpso.retailer_id) desc
            ) as rn
        from materialized_views.retailer_polygon_sales_ops rpso
        left join materialized_views.retailer_polygon rp
            on rp.retailer_id = rpso.retailer_id
        left join districts d
            on d.id = rp.district_id
        left join cities c
            on c.id = d.city_id
        left join states s
            on s.id = c.state_id
        left join regions r
            on r.id = s.region_id
        left join materialized_views.maxman_polygons st
            on st.sales_ops_polygon_name = rpso.sales_ops_polygon_name_en
        left join agents a
            on a.id = st.agent_id
        where rpso.sales_ops_polygon_name_en is not null
          and rp.polygon_name_en is not null
        group by
            rpso.sales_ops_polygon_name_en,
            rp.polygon_name_en,
            d.name_en,
            c.name_en,
            r.name_en,
            st.agent_id,
            a.partner_integration_id,
            a.name,
            st.cluster
    )
    qualify rn = 1
),

/* Unified orders from Instock + Marketplace (excluding cancelled/failed). */
orders_data as (
    select
        so.created_at::date as so_date,
        so.retailer_id,
        'instock' as vertical,
        so.channel,
        max(so.parent_sales_order_id) as order_id,
        min(so.created_at) as created_at,
        sum(pso.total_price) as total_price,
        count(distinct pso.product_id) as sku_count,
        count(distinct p.brand_id) as brand_count
    from sales_orders so
    inner join product_sales_order pso
        on pso.sales_order_id = so.id
    inner join products p
        on p.id = pso.product_id
    where so.sales_order_status_id not in (7, 12)
    group by 1, 2, 3, 4

    union all

    select
        so.created_at::date as so_date,
        so.retailer_id,
        'mp' as vertical,
        so.channel,
        max(so.id) as order_id,
        min(so.created_at) as created_at,
        sum(pso.total_price) as total_price,
        count(distinct pso.product_id) as sku_count,
        count(distinct p.brand_id) as brand_count
    from egypt_marketplace.sales_orders so
    inner join egypt_marketplace.sales_order_products pso
        on pso.order_id = so.id
    inner join products p
        on p.id = pso.product_id
    where so.status not in (3, 7, 8)
    group by 1, 2, 3, 4
),

/* Retailer value segmentation (HV/MV/LV) over the last 18 months. */
retailer_statuses as (
    select
        retailer_id,
        case
            when sum(case when rank_bucket = 'high' then 1 else 0 end) > 0 then 'HV'
            when sum(case when rank_bucket = 'med' then 1 else 0 end) > 0 then 'MV'
            else 'LV'
        end as retailer_value,
        max(
            case
                when month = dateadd(month, -1, date_trunc('month', current_date))
                    then 1
                else 0
            end
        ) as active_lm
    from (
        select
            month,
            retailer_id,
            case
                when pct < 0.30 then 'low'
                when pct between 0.30 and 0.75 then 'med'
                when pct between 0.75 and 1 then 'high'
            end as rank_bucket
        from (
            select
                month,
                retailer_id,
                percent_rank() over (partition by month order by nmv) as pct
            from (
                select
                    date_trunc('month', created_at::date) as month,
                    retailer_id,
                    sum(total_price) as nmv,
                    count(distinct order_id) as order_count
                from orders_data
                where created_at >= dateadd(month, -18, date_trunc('month', current_date))
                group by 1, 2
            ) monthly_nmv
        ) ranked
    ) scored
    group by retailer_id
),

/* Replace this source with the physical fulfilled-visits dataset used in your environment. */
visits_data as (
    select
        retailer_id,
        arrive_time
    from logistics_scheme_sales_data_raw
    where arrive_time is not null
),

/* Two snapshots: first day of current month vs now. */
snapshots as (
    select
        'month_start' as snapshot_key,
        date_trunc('month', current_date)::timestamp as snapshot_ts
    union all
    select
        'current' as snapshot_key,
        current_timestamp()::timestamp as snapshot_ts
),

/* Orders attached to visits (visit order window: -15 min to +24h). */
sales_team_orders as (
    select
        s.snapshot_key,
        s.snapshot_ts,
        v.retailer_id,
        v.arrive_time,
        so.order_id as visit_order_id,
        so.created_at::date as visit_order_date,
        so.brand_count,
        case when so.brand_count >= 5 then so.order_id else null end as golden_order_id
    from snapshots s
    inner join visits_data v
        on v.arrive_time >= s.snapshot_ts - interval '30 day'
       and v.arrive_time < s.snapshot_ts
    left join orders_data so
        on so.retailer_id = v.retailer_id
       and so.created_at between v.arrive_time - interval '15 minute'
                             and v.arrive_time + interval '24 hour'
       and so.created_at <= s.snapshot_ts
    where so.order_id is not null
),

/* Snapshot metrics needed to classify next channel bucket. */
bucket_inputs as (
    select
        v.snapshot_key,
        v.snapshot_ts,
        v.retailer_id,
        max(v.arrive_time) as last_visit_order,
        datediff('day', max(v.arrive_time), v.snapshot_ts) as days_since_last_visit_order,
        count(distinct v.visit_order_date) as visits_orders,
        count(distinct v.golden_order_id) as visits_golden_orders,
        count(distinct so.order_id) as after_visit_orders,
        count(distinct case when so.brand_count >= 5 then so.order_id end) as after_visit_golden_orders
    from sales_team_orders v
    left join orders_data so
        on so.retailer_id = v.retailer_id
       and so.created_at > v.arrive_time + interval '24 hour'
       and so.created_at <= v.snapshot_ts
    group by 1, 2, 3
),

channel_by_snapshot as (
    select
        snapshot_key,
        snapshot_ts,
        retailer_id,
        last_visit_order,
        days_since_last_visit_order,
        visits_orders,
        visits_golden_orders,
        after_visit_orders,
        after_visit_golden_orders,
        case
            when days_since_last_visit_order between 7 and 14
                 and after_visit_orders = 0
                 and visits_golden_orders >= 1 then 'telesales'
            when days_since_last_visit_order > 14
                 and after_visit_orders = 0
                 and visits_golden_orders >= 1 then 'sales_2'
            when days_since_last_visit_order > 7
                 and after_visit_orders = 0
                 and visits_golden_orders = 0 then 'sales_1'
            when visits_golden_orders + after_visit_golden_orders >= 1
                 and after_visit_orders >= 1 then 'growth'
            else 'tbd'
        end as next_visit_channel
    from bucket_inputs
),

bucket_comparison as (
    select
        retailer_id,
        max(case when snapshot_key = 'month_start' then next_visit_channel end) as bucket_on_month_start,
        max(case when snapshot_key = 'current' then next_visit_channel end) as bucket_now,
        max(case when snapshot_key = 'month_start' then last_visit_order end) as last_visit_order_month_start,
        max(case when snapshot_key = 'current' then last_visit_order end) as last_visit_order_now
    from channel_by_snapshot
    group by retailer_id
),

retailer_polygons as (
    select
        retailer_id,
        final_polygon,
        polygon_name_en,
        district_name_en,
        city_name_en,
        region_name_en,
        agent_id,
        partner_integration_id,
        agent_name
    from (
        select
            rpso.retailer_id,
            p.final_polygon,
            p.polygon_name_en,
            p.district_name_en,
            p.city_name_en,
            p.region_name_en,
            p.agent_id,
            p.partner_integration_id,
            p.agent_name,
            row_number() over (
                partition by rpso.retailer_id
                order by p.final_polygon
            ) as rn
        from materialized_views.retailer_polygon_sales_ops rpso
        left join polygons p
            on p.sales_ops_polygon_name_en = rpso.sales_ops_polygon_name_en
    )
    where rn = 1
)

select
    bc.retailer_id,
    coalesce(bc.bucket_on_month_start, 'no_bucket') as bucket_on_month_start,
    coalesce(bc.bucket_now, 'no_bucket') as bucket_now,
    case
        when coalesce(bc.bucket_on_month_start, 'no_bucket')
           = coalesce(bc.bucket_now, 'no_bucket')
            then 'same'
        else 'changed'
    end as bucket_movement,
    bc.last_visit_order_month_start,
    bc.last_visit_order_now,
    rs.retailer_value,
    rs.active_lm,
    rp.final_polygon,
    rp.polygon_name_en,
    rp.district_name_en,
    rp.city_name_en,
    rp.region_name_en,
    rp.agent_id,
    rp.partner_integration_id,
    rp.agent_name,
    current_timestamp() as view_refreshed_at
from bucket_comparison bc
left join retailer_statuses rs
    on rs.retailer_id = bc.retailer_id
left join retailer_polygons rp
    on rp.retailer_id = bc.retailer_id
;
