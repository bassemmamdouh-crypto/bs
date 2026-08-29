/* ============================================================================
   Retailer category achievement vs monthly target, with MTD target.
   ----------------------------------------------------------------------------
   Source: retailers_targets_ach (one or more rows per retailer / category).

   Working-day rule
   ----------------
   Fridays are not working days (ISODOW = 5). The monthly target is spread
   evenly over the non-Friday days of the current month:

       daily_target = month_target / working_days_in_month
       mtd_target   = daily_target * working_days_elapsed

   working_days_elapsed counts non-Friday days from the 1st of the month
   through current_date inclusive. A Friday does not raise the MTD target
   (it stays at Thursday's value). Achievement is not restricted: sales
   booked on Friday still sit in *_ach.

   Metabase optional filters use [[ ... ]] so an unset variable drops out
   and leaves `1=1`.
   ============================================================================ */

with calendar as (
    select d.day::date as day
    from generate_series(
        date_trunc('month', current_date)::date,
        (date_trunc('month', current_date) + interval '1 month' - interval '1 day')::date,
        interval '1 day'
    ) as d(day)
),
working_days as (
    select
        count(*) filter (where extract(isodow from day) <> 5) as working_days_in_month,
        count(*) filter (
            where extract(isodow from day) <> 5
              and day <= current_date
        ) as working_days_elapsed
    from calendar
)
select
    r.retailer_id,
    r.retailer_name,
    r.market_name,
    r.polygon_name,
    wd.working_days_elapsed,
    wd.working_days_in_month,

    max(case when r.category_name = 'Lays' then r.tier end) as lays_tier,
    coalesce(sum(case when r.category_name = 'Lays' then r.sold_cases end), 0) as lays_ach,
    coalesce(sum(case when r.category_name = 'Lays' then r.retailer_month_target end), 0) as lays_target,
    coalesce(
        sum(case when r.category_name = 'Lays' then r.sold_cases end)::numeric
        / nullif(sum(case when r.category_name = 'Lays' then r.retailer_month_target end), 0),
        0
    ) as lays_ach_percentage,
    coalesce(
        sum(case when r.category_name = 'Lays' then r.retailer_month_target end)::numeric
        * wd.working_days_elapsed
        / nullif(wd.working_days_in_month, 0),
        0
    ) as lays_mtd_target,

    max(case when r.category_name = 'Pepsi' then r.tier end) as pepsi_tier,
    coalesce(sum(case when r.category_name = 'Pepsi' then r.sold_cases end), 0) as pepsi_ach,
    coalesce(sum(case when r.category_name = 'Pepsi' then r.retailer_month_target end), 0) as pepsi_target,
    coalesce(
        sum(case when r.category_name = 'Pepsi' then r.sold_cases end)::numeric
        / nullif(sum(case when r.category_name = 'Pepsi' then r.retailer_month_target end), 0),
        0
    ) as pepsi_ach_percentage,
    coalesce(
        sum(case when r.category_name = 'Pepsi' then r.retailer_month_target end)::numeric
        * wd.working_days_elapsed
        / nullif(wd.working_days_in_month, 0),
        0
    ) as pepsi_mtd_target,

    max(case when r.category_name = 'Aquafina' then r.tier end) as aquafina_tier,
    coalesce(sum(case when r.category_name = 'Aquafina' then r.sold_cases end), 0) as aquafina_ach,
    coalesce(sum(case when r.category_name = 'Aquafina' then r.retailer_month_target end), 0) as aquafina_target,
    coalesce(
        sum(case when r.category_name = 'Aquafina' then r.sold_cases end)::numeric
        / nullif(sum(case when r.category_name = 'Aquafina' then r.retailer_month_target end), 0),
        0
    ) as aquafina_ach_percentage,
    coalesce(
        sum(case when r.category_name = 'Aquafina' then r.retailer_month_target end)::numeric
        * wd.working_days_elapsed
        / nullif(wd.working_days_in_month, 0),
        0
    ) as aquafina_mtd_target

from retailers_targets_ach r
cross join working_days wd
where
    [[polygon_name = {{polygon_name}} --]] 1=1
    and [[warehouse = {{warehouse}} --]] 1=1
    and [[segment = {{segment}} --]] segment = 'Retail'
group by
    r.retailer_id,
    r.retailer_name,
    r.market_name,
    r.polygon_name,
    wd.working_days_elapsed,
    wd.working_days_in_month
