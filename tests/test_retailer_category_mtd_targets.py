"""Working-day MTD target math for retailer_category_mtd_targets.sql."""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import duckdb
import pytest

SQL_PATH = Path(__file__).resolve().parents[1] / "queries" / "retailer_category_mtd_targets.sql"


def working_days(as_of: dt.date) -> tuple[int, int]:
    start = as_of.replace(day=1)
    if start.month == 12:
        end = dt.date(start.year + 1, 1, 1) - dt.timedelta(days=1)
    else:
        end = dt.date(start.year, start.month + 1, 1) - dt.timedelta(days=1)
    in_month = elapsed = 0
    day = start
    while day <= end:
        if day.isoweekday() != 5:
            in_month += 1
            if day <= as_of:
                elapsed += 1
        day += dt.timedelta(days=1)
    return elapsed, in_month


def test_august_2026_friday_exclusion():
    elapsed, in_month = working_days(dt.date(2026, 8, 29))
    assert in_month == 27
    assert elapsed == 25
    # Friday 28 Aug adds nothing vs Thursday 27 Aug.
    elapsed_thu, _ = working_days(dt.date(2026, 8, 27))
    elapsed_fri, _ = working_days(dt.date(2026, 8, 28))
    assert elapsed_thu == elapsed_fri == 24


def test_mtd_target_scales_month_target_by_elapsed_share():
    elapsed, in_month = working_days(dt.date(2026, 8, 29))
    month_target = 2700
    mtd = month_target * elapsed / in_month
    assert mtd == pytest.approx(2500)


def _run_query(as_of: dt.date):
    raw = SQL_PATH.read_text()
    raw = raw.replace("current_date", f"DATE '{as_of.isoformat()}'")
    raw = raw.replace("[[polygon_name = {{polygon_name}} --]] 1=1", "1=1")
    raw = raw.replace("[[warehouse = {{warehouse}} --]] 1=1", "1=1")
    raw = raw.replace("[[segment = {{segment}} --]] segment = 'Retail'", "segment = 'Retail'")
    con = duckdb.connect()
    con.execute(
        """
        create table retailers_targets_ach as
        select * from (
            values
                (101, 'R1', 'M1', 'الحرية', 'Retail', 'Lays', 'A', 100, 2700),
                (101, 'R1', 'M1', 'الحرية', 'Retail', 'Pepsi', 'B', 400, 5400),
                (101, 'R1', 'M1', 'الحرية', 'Retail', 'Aquafina', 'C', 50, 810),
                (102, 'R2', 'M1', 'شعلة', 'Retail', 'Pepsi', 'A', 10, 2700),
                (103, 'R3', 'M1', 'الحرية', 'Wholesale', 'Pepsi', 'A', 999, 9999)
        ) t(retailer_id, retailer_name, market_name, polygon_name, segment,
            category_name, tier, sold_cases, retailer_month_target)
        """
    )
    return con.execute(raw).fetchdf()


def test_sql_adds_mtd_target_per_category():
    as_of = dt.date(2026, 8, 29)
    elapsed, in_month = working_days(as_of)
    df = _run_query(as_of)
    assert set(df["retailer_id"]) == {101, 102}
    row = df.set_index("retailer_id").loc[101]
    assert row["lays_mtd_target"] == pytest.approx(2700 * elapsed / in_month)
    assert row["pepsi_mtd_target"] == pytest.approx(5400 * elapsed / in_month)
    assert row["aquafina_mtd_target"] == pytest.approx(810 * elapsed / in_month)
    assert row["working_days_elapsed"] == elapsed
    assert row["working_days_in_month"] == in_month
    assert "pepsi_tier" in df.columns
    assert list(df.columns).count("lays_tier") == 1
