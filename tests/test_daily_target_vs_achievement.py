"""
Runs retailer_agent_daily_target_vs_achievement.sql against DuckDB on a fixture
warehouse and checks every output row against an independent Python
implementation of the working-day target spread.

    python3 tests/test_daily_target_vs_achievement.py

The SQL file is executed verbatim, so the calendar arithmetic, the Friday rule
and the month-to-date joins are all exercised as written.
"""

import calendar
import datetime as dt
import os
import sys

import duckdb


SQL_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "retailer_agent_daily_target_vs_achievement.sql",
)

FRIDAY = 4
TOLERANCE = 1e-9

# --------------------------------------------------------------------------
# Fixture warehouse
# --------------------------------------------------------------------------

BRANDS = [
    (1, "Pepsi"),
    (2, "Lay's"),
    (3, "Aquafina"),          # in the portfolio but outside both targets
    (4, "Malaysian Snacks"),  # contains "lay" mid-word, must not match Lay's
]

PRODUCTS = [
    (11, 1),
    (12, 1),
    (21, 2),
    (31, 3),
    (41, 4),
]

RETAILER_TIERS = [
    # retailer, pepsi_tier, lays_tier, pepsi_month_target, lays_month_target
    (101, "A", "A", 2600.0, 1300.0),
    (102, "B", "C", 5200.0, 650.0),
    # 103 is deliberately absent: a dispatched retailer with no tier row
]

# (order_id, retailer_id, date, status_id, [(product_id, total_price), ...])
SALES_ORDERS = [
    (1, 101, "2026-04-30", 1, [(11, 900.0)]),    # previous month, must not leak
    (2, 101, "2026-05-04", 1, [(11, 150.0)]),
    (3, 101, "2026-05-04", 7, [(11, 400.0)]),    # cancelled, must be ignored
    (4, 101, "2026-05-06", 1, [(21, 100.0)]),
    (5, 101, "2026-05-08", 1, [(11, 500.0)]),    # a Friday: counts as achievement
    (6, 101, "2026-05-11", 1, [(31, 700.0)]),    # Aquafina, outside both targets
    (7, 101, "2026-05-11", 1, [(41, 300.0)]),    # "Malaysian Snacks", not Lay's
    (8, 102, "2026-02-03", 1, [(12, 800.0), (21, 250.0)]),
    (9, 102, "2026-02-27", 1, [(11, 1200.0)]),
    (10, 102, "2026-05-20", 1, [(11, 4000.0), (21, 900.0)]),
    (11, 103, "2026-05-04", 1, [(11, 75.0)]),    # retailer without a tier row
    (12, 101, "2026-05-12", 12, [(11, 999.0)]),  # failed, must be ignored
]

# (date, retailer_id, agent_id, updated_at)
TARGET_VISITS = [
    ("2026-05-01", 101, 9001, "2026-05-01 08:10:00"),  # month starts on a Friday
    ("2026-05-04", 101, 9001, "2026-05-04 09:00:00"),
    ("2026-05-04", 101, 9001, "2026-05-04 17:30:00"),  # duplicate touch, same day
    ("2026-05-04", 103, 9001, "2026-05-04 10:15:00"),
    ("2026-05-04", 102, 9002, "2026-05-04 11:00:00"),
    ("2026-05-06", 101, 9001, "2026-05-06 09:05:00"),
    ("2026-05-07", 101, 9001, "2026-05-07 09:20:00"),
    ("2026-05-08", 101, 9001, "2026-05-08 09:40:00"),  # Friday
    ("2026-05-09", 101, 9002, "2026-05-09 09:15:00"),  # Saturday is a working day
    ("2026-05-11", 101, 9001, "2026-05-11 08:55:00"),
    ("2026-05-20", 102, 9002, "2026-05-20 12:00:00"),
    ("2026-05-31", 101, 9001, "2026-05-31 13:00:00"),  # last day of the month
    ("2026-02-03", 102, 9002, "2026-02-03 08:30:00"),  # 28-day month
    ("2026-02-28", 102, 9002, "2026-02-28 08:30:00"),  # last day of a 28-day month
]


def build_fixture_warehouse(con):
    con.execute(
        """
        create table brands (id integer, name varchar);
        create table products (id integer, brand_id integer);
        create table sales_orders (
            id integer,
            retailer_id integer,
            created_at timestamp,
            sales_order_status_id integer
        );
        create table product_sales_order (
            sales_order_id integer,
            product_id integer,
            total_price double
        );
        create table retailer_tiers (
            retailer_id integer,
            pepsi_tier varchar,
            lays_tier varchar,
            pepsi_month_target double,
            lays_month_target double
        );
        create table target_visits (
            date date,
            retailer_id integer,
            agent_id integer,
            updated_at timestamp
        );
        """
    )

    con.executemany("insert into brands values (?, ?)", BRANDS)
    con.executemany("insert into products values (?, ?)", PRODUCTS)
    con.executemany("insert into retailer_tiers values (?, ?, ?, ?, ?)", RETAILER_TIERS)
    con.executemany("insert into target_visits values (?, ?, ?, ?)", TARGET_VISITS)

    for order_id, retailer_id, order_date, status_id, lines in SALES_ORDERS:
        con.execute(
            "insert into sales_orders values (?, ?, ?, ?)",
            [order_id, retailer_id, f"{order_date} 12:00:00", status_id],
        )
        for product_id, total_price in lines:
            con.execute(
                "insert into product_sales_order values (?, ?, ?)",
                [order_id, product_id, total_price],
            )


# --------------------------------------------------------------------------
# Independent reference implementation
# --------------------------------------------------------------------------

BRAND_OF_PRODUCT = {product_id: brand_id for product_id, brand_id in PRODUCTS}
TARGET_OF_BRAND = {1: "pepsi", 2: "lays"}
EXCLUDED_STATUSES = {7, 12}
TIERS_BY_RETAILER = {row[0]: row for row in RETAILER_TIERS}


def working_days_in_month(day):
    days_in_month = calendar.monthrange(day.year, day.month)[1]
    return sum(
        1
        for d in range(1, days_in_month + 1)
        if dt.date(day.year, day.month, d).weekday() != FRIDAY
    )


def working_days_elapsed(day):
    return sum(
        1
        for d in range(1, day.day + 1)
        if dt.date(day.year, day.month, d).weekday() != FRIDAY
    )


def achieved_to_date(retailer_id, day, brand_key):
    month_start = day.replace(day=1)
    total = 0.0
    for _, order_retailer, order_date, status_id, lines in SALES_ORDERS:
        if order_retailer != retailer_id or status_id in EXCLUDED_STATUSES:
            continue
        sales_date = dt.date.fromisoformat(order_date)
        if not (month_start <= sales_date <= day):
            continue
        for product_id, total_price in lines:
            if TARGET_OF_BRAND.get(BRAND_OF_PRODUCT[product_id]) == brand_key:
                total += total_price
    return total


def expected_row(day, retailer_id):
    in_month = working_days_in_month(day)
    elapsed = working_days_elapsed(day)
    tier = TIERS_BY_RETAILER.get(retailer_id)

    expected = {
        "working_days_in_month": in_month,
        "working_days_elapsed": elapsed,
        "working_days_remaining": in_month - elapsed,
        "is_working_day": 0 if day.weekday() == FRIDAY else 1,
        "pepsi_achieved_to_date": achieved_to_date(retailer_id, day, "pepsi"),
        "lays_achieved_to_date": achieved_to_date(retailer_id, day, "lays"),
    }

    if tier is None:
        expected["pepsi_target_to_date"] = None
        expected["lays_target_to_date"] = None
    else:
        expected["pepsi_target_to_date"] = tier[3] * elapsed / in_month
        expected["lays_target_to_date"] = tier[4] * elapsed / in_month
    return expected


# --------------------------------------------------------------------------
# Harness
# --------------------------------------------------------------------------

def fetch_records(con, query):
    result = con.execute(query)
    columns = [description[0] for description in result.description]
    records = []
    for row in result.fetchall():
        record = dict(zip(columns, row))
        record["dispatch_date"] = as_date(record["dispatch_date"])
        records.append(record)
    return records, columns


def as_date(value):
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    return dt.date.fromisoformat(str(value)[:10])


def split_statements(sql_text):
    """Split on statement terminators, ignoring semicolons inside comments and
    string literals."""
    statements = []
    buffer = []
    index = 0
    length = len(sql_text)
    while index < length:
        char = sql_text[index]
        pair = sql_text[index:index + 2]
        if pair == "/*":
            end = sql_text.find("*/", index + 2)
            end = length if end == -1 else end + 2
            buffer.append(sql_text[index:end])
            index = end
        elif pair == "--":
            end = sql_text.find("\n", index)
            end = length if end == -1 else end
            buffer.append(sql_text[index:end])
            index = end
        elif char in "'\"":
            end = sql_text.find(char, index + 1)
            end = length if end == -1 else end + 1
            buffer.append(sql_text[index:end])
            index = end
        elif char == ";":
            statements.append("".join(buffer))
            buffer = []
            index += 1
        else:
            buffer.append(char)
            index += 1
    statements.append("".join(buffer))
    return [s for s in statements if s.strip()]


# --------------------------------------------------------------------------
# Checks
# --------------------------------------------------------------------------

class Checker:
    def __init__(self):
        self.failures = []
        self.checks = 0

    def equal(self, label, actual, expected):
        self.checks += 1
        if isinstance(expected, float) or isinstance(actual, float):
            ok = (
                actual is not None
                and expected is not None
                and abs(float(actual) - float(expected)) < TOLERANCE
            )
        else:
            ok = actual == expected
        if not ok:
            self.failures.append(f"{label}: got {actual!r}, expected {expected!r}")

    def is_none(self, label, actual):
        self.checks += 1
        if actual is not None:
            self.failures.append(f"{label}: got {actual!r}, expected None")


def check_dialect_portability(check, sql_text):
    """The view runs on Snowflake in production but is verified here on DuckDB,
    so assert that both dialects read it the same way."""
    try:
        import sqlglot
    except ImportError:
        print("sqlglot is not installed: skipping the Snowflake dialect check")
        return

    as_snowflake = [s.sql(dialect="snowflake")
                    for s in sqlglot.parse(sql_text, read="snowflake") if s]
    as_duckdb = [s.sql(dialect="snowflake")
                 for s in sqlglot.parse(sql_text, read="duckdb") if s]
    check.equal("statements parsed as Snowflake", len(as_snowflake), 2)
    check.equal("Snowflake and DuckDB parse to the same tree", as_duckdb, as_snowflake)


def run():
    con = duckdb.connect()
    build_fixture_warehouse(con)

    with open(SQL_PATH) as handle:
        sql_text = handle.read()
    for statement in split_statements(sql_text):
        con.execute(statement)

    # fetchall() rather than fetchdf(): pandas turns SQL NULL into NaN, which
    # would hide the difference between "no target" and "a target of zero".
    records, columns = fetch_records(
        con,
        """
        select *
        from retailer_agent_daily_target_vs_achievement
        order by dispatch_date, retailer_id, agent_id
        """,
    )

    check = Checker()
    check_dialect_portability(check, sql_text)

    # One output row per (day, retailer, agent): the duplicate 2026-05-04 touch
    # for retailer 101 must collapse to a single row.
    check.equal("row count", len(records), len(TARGET_VISITS) - 1)
    check.equal(
        "duplicate dispatch collapsed",
        len([r for r in records
             if str(r["dispatch_date"]) == "2026-05-04" and r["retailer_id"] == 101]),
        1,
    )
    check.equal(
        "latest updated_at kept",
        str(next(r["updated_at"] for r in records
                 if str(r["dispatch_date"]) == "2026-05-04" and r["retailer_id"] == 101)),
        "2026-05-04 17:30:00",
    )

    # Every row against the independent implementation.
    for record in records:
        day = record["dispatch_date"]
        retailer_id = record["retailer_id"]
        label = f"{day} r{retailer_id} a{record['agent_id']}"
        expected = expected_row(day, retailer_id)

        for field, value in expected.items():
            if value is None:
                check.is_none(f"{label} {field}", record[field])
            else:
                check.equal(f"{label} {field}", record[field], value)

        # Derived measures stay consistent with the base ones.
        if expected["pepsi_target_to_date"] is not None:
            gap = expected["pepsi_achieved_to_date"] - expected["pepsi_target_to_date"]
            check.equal(f"{label} pepsi_gap_to_date", record["pepsi_gap_to_date"], gap)
            if expected["pepsi_target_to_date"] == 0:
                check.is_none(f"{label} pepsi_achievement_rate", record["pepsi_achievement_rate"])
            else:
                check.equal(
                    f"{label} pepsi_achievement_rate",
                    record["pepsi_achievement_rate"],
                    expected["pepsi_achieved_to_date"] / expected["pepsi_target_to_date"],
                )

    # Every (day, retailer) pair is dispatched to exactly one agent in the fixture.
    by_key = {(str(r["dispatch_date"]), r["retailer_id"]): r for r in records}

    # May 2026 starts on a Friday: 31 days, 5 Fridays, 26 working days.
    may_01 = by_key[("2026-05-01", 101)]
    check.equal("2026-05-01 working_days_in_month", may_01["working_days_in_month"], 26)
    check.equal("2026-05-01 is a Friday", may_01["is_working_day"], 0)
    check.equal("2026-05-01 no working day elapsed", may_01["working_days_elapsed"], 0)
    check.equal("2026-05-01 no target yet", may_01["pepsi_target_to_date"], 0.0)
    check.is_none("2026-05-01 rate undefined", may_01["pepsi_achievement_rate"])

    # 2600 / 26 = 100 per working day. By Mon 4 May three working days elapsed
    # (Sat 2, Sun 3, Mon 4) because Fri 1 May does not count.
    may_04 = by_key[("2026-05-04", 101)]
    check.equal("2026-05-04 elapsed", may_04["working_days_elapsed"], 3)
    check.equal("2026-05-04 daily target", may_04["pepsi_daily_target"], 100.0)
    check.equal("2026-05-04 target to date", may_04["pepsi_target_to_date"], 300.0)
    check.equal("2026-05-04 achieved", may_04["pepsi_achieved_to_date"], 150.0)
    check.equal("2026-05-04 gap", may_04["pepsi_gap_to_date"], -150.0)
    check.equal("2026-05-04 rate", may_04["pepsi_achievement_rate"], 0.5)

    # A Friday must inherit Thursday's target to date unchanged, while sales
    # made on that Friday still add to the achievement.
    may_07 = by_key[("2026-05-07", 101)]
    may_08 = by_key[("2026-05-08", 101)]
    check.equal("Thu 7 May elapsed", may_07["working_days_elapsed"], 6)
    check.equal("Fri 8 May elapsed unchanged", may_08["working_days_elapsed"], 6)
    check.equal(
        "Fri 8 May target to date unchanged",
        may_08["pepsi_target_to_date"],
        may_07["pepsi_target_to_date"],
    )
    check.equal("Fri 8 May target to date", may_08["pepsi_target_to_date"], 600.0)
    check.equal("Fri 8 May pepsi sold on the day", may_08["pepsi_achieved_on_day"], 500.0)
    check.equal("Fri 8 May achievement includes Friday", may_08["pepsi_achieved_to_date"], 650.0)
    check.equal("Fri 8 May ahead of target", may_08["pepsi_gap_to_date"], 50.0)

    # Saturday is a normal working day, so the target advances again.
    may_09 = by_key[("2026-05-09", 101)]
    check.equal("Sat 9 May elapsed", may_09["working_days_elapsed"], 7)
    check.equal("Sat 9 May target to date", may_09["pepsi_target_to_date"], 700.0)

    # Out-of-scope brands and previous-month orders never reach the achievement.
    may_11 = by_key[("2026-05-11", 101)]
    check.equal("Aquafina and Malaysian Snacks excluded",
                may_11["pepsi_achieved_to_date"], 650.0)
    check.equal("April order excluded", may_11["lays_achieved_to_date"], 100.0)

    # By the last day of the month, target to date equals the full month target.
    may_31 = by_key[("2026-05-31", 101)]
    check.equal("31 May elapsed", may_31["working_days_elapsed"], 26)
    check.equal("31 May remaining", may_31["working_days_remaining"], 0)
    check.equal("31 May target to date", may_31["pepsi_target_to_date"], 2600.0)
    check.is_none("31 May run rate undefined", may_31["pepsi_required_daily_run_rate"])

    # Short month: February 2026 has 28 days and 4 Fridays.
    feb_28 = by_key[("2026-02-28", 102)]
    check.equal("Feb 2026 working days", feb_28["working_days_in_month"], 24)
    check.equal("28 Feb elapsed", feb_28["working_days_elapsed"], 24)
    check.equal("28 Feb target to date", feb_28["pepsi_target_to_date"], 5200.0)
    check.equal("28 Feb achieved", feb_28["pepsi_achieved_to_date"], 2000.0)

    # A dispatched retailer with no tier row keeps its achievement but has no
    # target, instead of silently reporting zero or dropping out of the report.
    no_tier = by_key[("2026-05-04", 103)]
    check.is_none("no tier -> no month target", no_tier["pepsi_month_target"])
    check.is_none("no tier -> no target to date", no_tier["pepsi_target_to_date"])
    check.equal("no tier -> achievement still reported",
                no_tier["pepsi_achieved_to_date"], 75.0)

    # Month-end projection: 650 achieved over 6 working days of 26.
    check.equal(
        "Fri 8 May projection",
        may_08["pepsi_projected_month_end"],
        650.0 * 26 / 6,
    )
    # Run rate needed over the 20 working days left after 8 May.
    check.equal("Fri 8 May remaining", may_08["working_days_remaining"], 20)
    check.equal(
        "Fri 8 May required run rate",
        may_08["pepsi_required_daily_run_rate"],
        (2600.0 - 650.0) / 20,
    )

    # Agent roll-up sums the retailers dispatched to an agent that day.
    agent_rows, _ = fetch_records(
        con,
        """
        select *
        from agent_daily_target_vs_achievement
        order by dispatch_date, agent_id
        """,
    )
    agent_9001_may_04 = next(
        r for r in agent_rows
        if r["dispatch_date"] == dt.date(2026, 5, 4) and r["agent_id"] == 9001
    )
    check.equal("agent 9001 retailers on 4 May",
                agent_9001_may_04["retailers_dispatched"], 2)
    check.equal("agent 9001 retailers without target",
                agent_9001_may_04["retailers_without_target"], 1)
    check.equal("agent 9001 pepsi target to date",
                agent_9001_may_04["pepsi_target_to_date"], 300.0)
    check.equal("agent 9001 pepsi achieved to date",
                agent_9001_may_04["pepsi_achieved_to_date"], 225.0)

    print(f"columns returned ({len(columns)}): {', '.join(columns)}")
    print()
    print(
        con.execute(
            """
            select
                dispatch_date,
                retailer_id,
                agent_id,
                is_working_day,
                working_days_elapsed,
                working_days_in_month,
                round(pepsi_target_to_date, 2) as pepsi_target_to_date,
                round(pepsi_achieved_to_date, 2) as pepsi_achieved_to_date,
                round(pepsi_gap_to_date, 2) as pepsi_gap_to_date,
                round(pepsi_achievement_rate, 4) as pepsi_achievement_rate,
                round(lays_target_to_date, 2) as lays_target_to_date,
                round(lays_achieved_to_date, 2) as lays_achieved_to_date
            from retailer_agent_daily_target_vs_achievement
            order by retailer_id, dispatch_date, agent_id
            """
        ).fetchdf().to_string(index=False, na_rep='NULL')
    )
    print()
    print(
        con.execute(
            """
            select
                dispatch_date,
                agent_id,
                retailers_dispatched,
                retailers_without_target,
                round(total_target_to_date, 2) as total_target_to_date,
                round(total_achieved_to_date, 2) as total_achieved_to_date,
                round(total_achievement_rate, 4) as total_achievement_rate
            from agent_daily_target_vs_achievement
            order by dispatch_date, agent_id
            """
        ).fetchdf().to_string(index=False, na_rep='NULL')
    )
    print()

    if check.failures:
        print(f"FAILED {len(check.failures)} of {check.checks} checks:")
        for failure in check.failures:
            print(f"  - {failure}")
        return 1

    print(f"PASSED all {check.checks} checks.")
    return 0


def test_daily_target_vs_achievement():
    assert run() == 0


if __name__ == "__main__":
    sys.exit(run())
