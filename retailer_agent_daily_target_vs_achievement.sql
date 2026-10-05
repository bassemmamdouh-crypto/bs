/* ============================================================================
   Daily target-to-date vs achievement-to-date, per dispatched retailer/agent.
   ----------------------------------------------------------------------------
   For every day a retailer is dispatched to an agent, this reports:
     * the month-to-date target the retailer should have reached by that day,
     * how much of it the retailer actually achieved by that day,
     * the gap, the achievement rate, the month-end projection and the daily
       run rate still required,
   for Pepsi and for Lay's separately, plus a combined total.

   Working-day rule
   ----------------
   Fridays are not working days and carry no target. The monthly target is
   spread evenly over the non-Friday days of the month:

       daily_target   = month_target / working_days_in_month
       target_to_date = daily_target * working_days_elapsed

   working_days_elapsed counts non-Friday days from the 1st of the month up to
   and including the dispatch date, so a Friday adds nothing and inherits the
   same target_to_date as the Thursday before it.

   Achievement is deliberately not restricted to working days: whatever the
   retailer buys on a Friday still counts towards the month-to-date
   achievement, it just never raises the target.

   Adapting this to your warehouse
   -------------------------------
   Only section 1 (SOURCES) should need editing. Point dispatch_raw,
   retailer_targets, brand_scope and retailer_brand_daily_sales at the real
   objects; everything after that is calendar and target arithmetic that does
   not depend on the physical schema.
   ============================================================================ */

create or replace view retailer_agent_daily_target_vs_achievement as
with

/* ---------------------------------------------------------------------------
   1. SOURCES  (the only block that is schema specific)
   --------------------------------------------------------------------------- */

/* One row per retailer dispatched to an agent on a given day. */
dispatch_raw as (
    select
        tv.date::date as dispatch_date,
        tv.retailer_id,
        tv.agent_id,
        tv.updated_at
    from target_visits tv
    where tv.date is not null
      and tv.retailer_id is not null
      and tv.agent_id is not null
),

/* Monthly target and tier per retailer. If this table is versioned by month,
   add its month column here and to the join in the `dispatch_enriched` CTE. */
retailer_targets as (
    select
        rt.retailer_id,
        rt.pepsi_tier,
        rt.lays_tier,
        rt.pepsi_month_target,
        rt.lays_month_target
    from retailer_tiers rt
),

/* Which brands roll up into which target. Replace the name matching with an
   explicit list of brand ids once they are known: it is cheaper and safer. */
brand_scope as (
    select
        b.id as brand_id,
        case
            when lower(b.name) like '%pepsi%' then 'pepsi'
            else 'lays'
        end as target_brand
    from brands b
    where lower(b.name) like '%pepsi%'
       or lower(b.name) like 'lay%'
),

/* Achievement source: value sold per retailer, per day, per target brand.
   Swap sum(pso.total_price) for a quantity column if the monthly targets are
   expressed in cartons/units rather than in money. */
retailer_brand_daily_sales as (
    select
        so.created_at::date as sales_date,
        so.retailer_id,
        bs.target_brand,
        sum(pso.total_price) as sales_value
    from sales_orders so
    inner join product_sales_order pso
        on pso.sales_order_id = so.id
    inner join products p
        on p.id = pso.product_id
    inner join brand_scope bs
        on bs.brand_id = p.brand_id
    where so.sales_order_status_id not in (7, 12)   -- cancelled / failed
    group by 1, 2, 3
),

/* ---------------------------------------------------------------------------
   2. WORKING-DAY CALENDAR
   --------------------------------------------------------------------------- */

/* Collapse to one row per retailer/agent/day; keep the latest touch. */
dispatch as (
    select
        dispatch_date,
        retailer_id,
        agent_id,
        max(updated_at) as updated_at
    from dispatch_raw
    group by 1, 2, 3
),

dispatch_months as (
    select distinct date_trunc('month', dispatch_date)::date as month_start
    from dispatch
),

day_offsets as (
    select day_offset
    from (
        values
            (0), (1), (2), (3), (4), (5), (6), (7), (8), (9),
            (10), (11), (12), (13), (14), (15), (16), (17), (18), (19),
            (20), (21), (22), (23), (24), (25), (26), (27), (28), (29),
            (30)
    ) as t (day_offset)
),

/* Every calendar day of every month that has at least one dispatch. */
month_calendar as (
    select
        m.month_start,
        (m.month_start + o.day_offset)::date as calendar_date,
        /* dayname() is 'Fri' on Snowflake and 'Friday' elsewhere. */
        case
            when dayname((m.month_start + o.day_offset)::date) ilike 'fri%' then 0
            else 1
        end as is_working_day
    from dispatch_months m
    cross join day_offsets o
    where (m.month_start + o.day_offset)::date <= last_day(m.month_start)
),

month_working_days as (
    select
        month_start,
        sum(is_working_day) as working_days_in_month
    from month_calendar
    group by month_start
),

calendar_progress as (
    select
        c.month_start,
        c.calendar_date,
        c.is_working_day,
        w.working_days_in_month,
        sum(c.is_working_day) over (
            partition by c.month_start
            order by c.calendar_date
            rows between unbounded preceding and current row
        ) as working_days_elapsed
    from month_calendar c
    inner join month_working_days w
        on w.month_start = c.month_start
),

/* ---------------------------------------------------------------------------
   3. MONTH-TO-DATE ACHIEVEMENT PER DISPATCHED DAY
   --------------------------------------------------------------------------- */

dispatch_days as (
    select distinct dispatch_date, retailer_id
    from dispatch
),

dispatch_actuals as (
    select
        d.dispatch_date,
        d.retailer_id,
        sum(case when s.target_brand = 'pepsi' then s.sales_value else 0 end) as pepsi_achieved_to_date,
        sum(case when s.target_brand = 'lays' then s.sales_value else 0 end) as lays_achieved_to_date,
        sum(case when s.target_brand = 'pepsi' and s.sales_date = d.dispatch_date then s.sales_value else 0 end) as pepsi_achieved_on_day,
        sum(case when s.target_brand = 'lays' and s.sales_date = d.dispatch_date then s.sales_value else 0 end) as lays_achieved_on_day
    from dispatch_days d
    left join retailer_brand_daily_sales s
        on s.retailer_id = d.retailer_id
       and s.sales_date >= date_trunc('month', d.dispatch_date)::date
       and s.sales_date <= d.dispatch_date
    group by 1, 2
),

/* ---------------------------------------------------------------------------
   4. TARGET SPREAD OVER WORKING DAYS
   --------------------------------------------------------------------------- */

dispatch_enriched as (
    select
        d.dispatch_date,
        d.retailer_id,
        d.agent_id,
        d.updated_at,
        cp.month_start,
        cp.is_working_day,
        cp.working_days_in_month,
        cp.working_days_elapsed,
        cp.working_days_in_month - cp.working_days_elapsed as working_days_remaining,
        rt.pepsi_tier,
        rt.lays_tier,
        rt.pepsi_month_target,
        rt.lays_month_target,
        cast(rt.pepsi_month_target as double) / nullif(cp.working_days_in_month, 0) as pepsi_daily_target,
        cast(rt.lays_month_target as double) / nullif(cp.working_days_in_month, 0) as lays_daily_target,
        cast(rt.pepsi_month_target as double) * cp.working_days_elapsed
            / nullif(cp.working_days_in_month, 0) as pepsi_target_to_date,
        cast(rt.lays_month_target as double) * cp.working_days_elapsed
            / nullif(cp.working_days_in_month, 0) as lays_target_to_date,
        a.pepsi_achieved_to_date,
        a.lays_achieved_to_date,
        a.pepsi_achieved_on_day,
        a.lays_achieved_on_day
    from dispatch d
    inner join calendar_progress cp
        on cp.calendar_date = d.dispatch_date
    left join retailer_targets rt
        on rt.retailer_id = d.retailer_id
    left join dispatch_actuals a
        on a.dispatch_date = d.dispatch_date
       and a.retailer_id = d.retailer_id
)

select
    e.dispatch_date,
    e.retailer_id,
    e.agent_id,
    e.updated_at,
    e.month_start,
    e.is_working_day,
    e.working_days_in_month,
    e.working_days_elapsed,
    e.working_days_remaining,

    e.pepsi_tier,
    e.pepsi_month_target,
    e.pepsi_daily_target,
    e.pepsi_target_to_date,
    e.pepsi_achieved_on_day,
    e.pepsi_achieved_to_date,
    e.pepsi_achieved_to_date - e.pepsi_target_to_date as pepsi_gap_to_date,
    e.pepsi_achieved_to_date / nullif(e.pepsi_target_to_date, 0) as pepsi_achievement_rate,
    e.pepsi_achieved_to_date * e.working_days_in_month
        / nullif(e.working_days_elapsed, 0) as pepsi_projected_month_end,
    (e.pepsi_month_target - e.pepsi_achieved_to_date)
        / nullif(e.working_days_remaining, 0) as pepsi_required_daily_run_rate,

    e.lays_tier,
    e.lays_month_target,
    e.lays_daily_target,
    e.lays_target_to_date,
    e.lays_achieved_on_day,
    e.lays_achieved_to_date,
    e.lays_achieved_to_date - e.lays_target_to_date as lays_gap_to_date,
    e.lays_achieved_to_date / nullif(e.lays_target_to_date, 0) as lays_achievement_rate,
    e.lays_achieved_to_date * e.working_days_in_month
        / nullif(e.working_days_elapsed, 0) as lays_projected_month_end,
    (e.lays_month_target - e.lays_achieved_to_date)
        / nullif(e.working_days_remaining, 0) as lays_required_daily_run_rate,

    e.pepsi_target_to_date + e.lays_target_to_date as total_target_to_date,
    e.pepsi_achieved_to_date + e.lays_achieved_to_date as total_achieved_to_date,
    (e.pepsi_achieved_to_date + e.lays_achieved_to_date)
        - (e.pepsi_target_to_date + e.lays_target_to_date) as total_gap_to_date,
    (e.pepsi_achieved_to_date + e.lays_achieved_to_date)
        / nullif(e.pepsi_target_to_date + e.lays_target_to_date, 0) as total_achievement_rate
from dispatch_enriched e
;


/* ============================================================================
   Agent-level roll-up: one row per agent per day, over the retailers that were
   dispatched to that agent on that day.
   ============================================================================ */

create or replace view agent_daily_target_vs_achievement as
select
    dispatch_date,
    agent_id,
    month_start,
    is_working_day,
    working_days_in_month,
    working_days_elapsed,
    working_days_remaining,
    count(distinct retailer_id) as retailers_dispatched,
    count(distinct case when pepsi_month_target is null then retailer_id end) as retailers_without_target,
    sum(pepsi_target_to_date) as pepsi_target_to_date,
    sum(pepsi_achieved_to_date) as pepsi_achieved_to_date,
    sum(pepsi_achieved_to_date) - sum(pepsi_target_to_date) as pepsi_gap_to_date,
    sum(pepsi_achieved_to_date) / nullif(sum(pepsi_target_to_date), 0) as pepsi_achievement_rate,
    sum(lays_target_to_date) as lays_target_to_date,
    sum(lays_achieved_to_date) as lays_achieved_to_date,
    sum(lays_achieved_to_date) - sum(lays_target_to_date) as lays_gap_to_date,
    sum(lays_achieved_to_date) / nullif(sum(lays_target_to_date), 0) as lays_achievement_rate,
    sum(total_target_to_date) as total_target_to_date,
    sum(total_achieved_to_date) as total_achieved_to_date,
    sum(total_achieved_to_date) - sum(total_target_to_date) as total_gap_to_date,
    sum(total_achieved_to_date) / nullif(sum(total_target_to_date), 0) as total_achievement_rate
from retailer_agent_daily_target_vs_achievement
group by 1, 2, 3, 4, 5, 6, 7
;
