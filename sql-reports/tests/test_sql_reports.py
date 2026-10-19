import sqlite3
import sys
from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import load_workbook

DEMO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(DEMO))
import sql_reports as sr  # noqa: E402
import make_sample_customers as customers_demo  # noqa: E402

sales_report = sr.sales_report  # the cleaning step the database is built from
AS_OF = date(2026, 10, 19)


def rows_of(con, file_name):
    report = sr.run_report(con, sr.QUERIES / file_name)
    return [dict(zip(report.columns, row)) for row in report.rows]


# ---------- a tiny database, small enough to work out by hand ----------

@pytest.fixture()
def tiny():
    """Three customers and eight orders. Only O1, O2, O3, O5 and O8 are sales."""
    con = sqlite3.connect(":memory:")
    sr.create_schema(con)
    con.executemany("INSERT INTO customers VALUES (?, ?, ?)",
                    [("C1", "Ann", "NSW"), ("C2", "Bob", "VIC"), ("C3", "Cy", "NSW")])
    con.executemany("INSERT INTO products VALUES (?, ?)", [("JAR", "Jar"), ("LID", "Lid")])
    con.executemany("INSERT INTO orders VALUES (?, ?, ?, ?, ?)", [
        ("O1", "2026-04-10", "Shipped", "NSW", "C1"),
        ("O2", "2026-04-20", "Delivered", "VIC", "C2"),
        ("O3", "2026-05-05", "Shipped", "NSW", "C1"),
        ("O4", "2026-05-06", "Cancelled", "NSW", "C3"),
        ("O5", "2026-06-01", "Delivered", "NSW", "C3"),
        ("O6", "2026-06-02", "Refunded", "VIC", "C2"),
        ("O7", "2026-06-15", "Pending", "NSW", "C1"),
        ("O8", "2026-06-20", "Shipped", "NSW", "C1"),
    ])
    con.executemany("INSERT INTO order_items VALUES (?, ?, ?, ?)", [
        ("O1", "JAR", 2, 1000), ("O1", "LID", 1, 500), ("O2", "JAR", 1, 1000), ("O3", "LID", 4, 500),
        ("O4", "JAR", 1, 1000), ("O5", "JAR", 3, 1000), ("O6", "LID", 2, 500), ("O7", "JAR", 1, 1000),
        ("O8", "JAR", 1, 1000),
    ])
    yield con
    con.close()


def test_view_keeps_only_shipped_delivered_and_completed_orders(tiny):
    sold = {order for (order,) in tiny.execute("SELECT DISTINCT order_id FROM sold_lines")}
    assert sold == {"O1", "O2", "O3", "O5", "O8"}


def test_headline_numbers(tiny):
    (row,) = rows_of(tiny, "00_headline.sql")
    assert row == {
        "Orders": 5, "Units sold": 12, "Net sales": 95.0, "Average order": 19.0,
        "Customers who bought": 3, "Repeat customers": 1, "Repeat customer rate": 0.3333,
        "Share of sales from repeat customers": 0.5789,  # Ann: $55 of $95
        "First order": "2026-04-10", "Latest order": "2026-06-20",
    }


def test_sales_by_month(tiny):
    assert rows_of(tiny, "01_sales_by_month.sql") == [
        {"Month": "2026-04", "Orders": 2, "Units": 4, "Sales": 35.0, "Average order": 17.5,
         "Change on previous month": None},  # nothing before April to compare with
        {"Month": "2026-05", "Orders": 1, "Units": 4, "Sales": 20.0, "Average order": 20.0,
         "Change on previous month": -0.4286},
        {"Month": "2026-06", "Orders": 2, "Units": 4, "Sales": 40.0, "Average order": 20.0,
         "Change on previous month": 1.0},
    ]


def test_top_products(tiny):
    assert rows_of(tiny, "02_top_products.sql") == [
        {"Rank": 1, "SKU": "JAR", "Product": "Jar", "Units": 7, "Orders": 4, "Sales": 70.0,
         "Share of sales": 0.7368, "Running share": 0.7368},
        {"Rank": 2, "SKU": "LID", "Product": "Lid", "Units": 5, "Orders": 2, "Sales": 25.0,
         "Share of sales": 0.2632, "Running share": 1.0},
    ]


def test_repeat_customers(tiny):
    # Bob's second order was refunded and Cy's was cancelled, so neither has ordered twice.
    assert rows_of(tiny, "03_repeat_customers.sql") == [
        {"Customer ID": "C1", "Customer": "Ann", "State": "NSW", "Orders": 3, "First order": "2026-04-10",
         "Latest order": "2026-06-20", "Days between first and latest": 71, "Total spent": 55.0,
         "Average order": 18.33},
    ]


def test_totals_by_state(tiny):
    assert rows_of(tiny, "04_totals_by_state.sql") == [
        {"State": "NSW", "Orders": 4, "Customers": 2, "Units": 11, "Sales": 85.0, "Average order": 21.25,
         "Share of sales": 0.8947},
        {"State": "VIC", "Orders": 1, "Customers": 1, "Units": 1, "Sales": 10.0, "Average order": 10.0,
         "Share of sales": 0.1053},
    ]


def test_the_database_refuses_bad_rows(tiny):
    for sql in (
        "INSERT INTO orders VALUES ('X1', '2026-04-10', 'Shipped', 'NSW', 'NOBODY')",  # unknown customer
        "INSERT INTO orders VALUES ('X2', '10/04/2026', 'Shipped', 'NSW', 'C1')",  # date not YYYY-MM-DD
        "INSERT INTO order_items VALUES ('O1', 'JAR', 0, 1000)",  # nothing ordered
        "INSERT INTO order_items VALUES ('O1', 'NOPE', 1, 1000)",  # unknown product
    ):
        with pytest.raises(sqlite3.IntegrityError):
            tiny.execute(sql)


def test_adding_a_sql_file_adds_a_report(tiny, tmp_path):
    for path in sr.QUERIES.glob("*.sql"):
        (tmp_path / path.name).write_text(path.read_text(encoding="utf-8"), encoding="utf-8")
    (tmp_path / "05_by_status.sql").write_text(
        "-- title: Orders by status\n-- about: Every order, counted by its status.\n"
        "SELECT status AS \"Status\", COUNT(*) AS \"Orders\" FROM orders GROUP BY status ORDER BY status;\n",
        encoding="utf-8")
    reports = sr.run_all(tiny, tmp_path)
    assert [r.title for r in reports] == ["Summary", "Sales by month", "Top products", "Repeat customers",
                                          "Totals by state", "Orders by status"]
    assert reports[-1].rows == [("Cancelled", 1), ("Delivered", 2), ("Pending", 1), ("Refunded", 1), ("Shipped", 3)]


def test_sheet_titles_are_checked(tiny, tmp_path):
    path = tmp_path / "09_long.sql"
    path.write_text("-- title: " + "x" * 32 + "\nSELECT 1;\n", encoding="utf-8")
    with pytest.raises(ValueError, match="can't be an Excel sheet name"):
        sr.run_report(tiny, path)
    assert sr.header_value("-- title: Sales by month\n-- about: Yes.\nSELECT 1;", "title") == "Sales by month"
    assert sr.header_value("SELECT 1;", "title") is None


# ---------- the full sample database, checked against pandas ----------

@pytest.fixture(scope="module")
def full(tmp_path_factory):
    con = sr.build_database(tmp_path_factory.mktemp("db") / "store.db")
    yield con
    con.close()


@pytest.fixture(scope="module")
def everything():
    """Every cleaned order line, with its customer."""
    clean, _ = sales_report.clean_orders(sales_report.load_export(sr.ORDERS))
    owner = dict(sr.read_csv(sr.ORDER_CUSTOMERS, ["Order ID", "Customer ID"]))
    clean["Customer ID"] = clean["Order ID"].map(owner)
    return clean


@pytest.fixture(scope="module")
def sold(everything):
    return everything[everything["Order Status"].isin(sales_report.SOLD_STATUSES)]


def test_database_holds_the_whole_export(full, everything):
    count = lambda table: full.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]  # noqa: E731
    assert count("order_items") == len(everything) == 797
    assert count("orders") == everything["Order ID"].nunique() == 572
    assert count("products") == 8
    assert full.execute("PRAGMA foreign_key_check").fetchall() == []
    assert full.execute("SELECT DISTINCT typeof(unit_price_cents) FROM order_items").fetchall() == [("integer",)]
    # every customer lives in the state their orders ship to
    assert full.execute("SELECT COUNT(*) FROM orders o JOIN customers c USING (customer_id) "
                        "WHERE o.ship_state <> c.state").fetchone()[0] == 0


def test_headline_matches_the_python_sales_report(full, everything):
    (row,) = rows_of(full, "00_headline.sql")
    python = sales_report.summarise(everything)
    assert row["Net sales"] == pytest.approx(python["sales"], abs=0.005)
    assert row["Orders"] == python["orders"]
    assert row["Units sold"] == python["units"]
    assert row["Average order"] == pytest.approx(python["average_order"], abs=0.005)
    assert row["First order"] == "2026-04-01" and row["Latest order"] <= "2026-09-30"


def test_sales_by_month_matches_pandas(full, sold):
    month = sold["Order Date"].dt.strftime("%Y-%m")
    expected = sold.groupby(month).agg(orders=("Order ID", "nunique"), units=("Qty", "sum"),
                                       sales=("Line Total", "sum"))
    rows = rows_of(full, "01_sales_by_month.sql")
    assert [r["Month"] for r in rows] == list(expected.index) == sorted(expected.index)
    assert [r["Orders"] for r in rows] == list(expected["orders"])
    assert [r["Units"] for r in rows] == list(expected["units"])
    assert [r["Sales"] for r in rows] == pytest.approx(list(expected["sales"]), abs=0.005)
    assert rows[0]["Change on previous month"] is None
    first, second = rows[0]["Sales"], rows[1]["Sales"]
    assert rows[1]["Change on previous month"] == pytest.approx((second - first) / first, abs=0.0001)
    assert sum(r["Sales"] for r in rows) == pytest.approx(sold["Line Total"].sum(), abs=0.005)


def test_top_products_match_pandas(full, sold):
    expected = sold.groupby("SKU")["Line Total"].sum().sort_values(ascending=False)
    rows = rows_of(full, "02_top_products.sql")
    assert [r["SKU"] for r in rows] == list(expected.index)
    assert [r["Sales"] for r in rows] == pytest.approx(list(expected), abs=0.005)
    assert [r["Rank"] for r in rows] == list(range(1, len(rows) + 1))
    assert sum(r["Share of sales"] for r in rows) == pytest.approx(1.0, abs=0.001)
    assert rows[-1]["Running share"] == 1.0
    assert [r["Running share"] for r in rows] == sorted(r["Running share"] for r in rows)
    assert rows[0]["Product"] == "Glass Meal Prep Set (10 pack)"
    assert len(rows) <= 10


def test_repeat_customers_match_pandas(full, sold):
    per_order = sold.groupby(["Customer ID", "Order ID"]).agg(day=("Order Date", "first"), cents=("Line Total", "sum"))
    per_customer = per_order.groupby("Customer ID").agg(orders=("day", "size"), total=("cents", "sum"),
                                                        first=("day", "min"), last=("day", "max"))
    expected = per_customer[per_customer["orders"] >= 2]
    rows = rows_of(full, "03_repeat_customers.sql")
    assert {r["Customer ID"] for r in rows} == set(expected.index)
    assert len(rows) == 102
    for r in rows:
        want = expected.loc[r["Customer ID"]]
        assert r["Orders"] == want["orders"] and r["Orders"] >= 2
        assert r["Total spent"] == pytest.approx(want["total"], abs=0.005)
        assert r["First order"] == f"{want['first']:%Y-%m-%d}" and r["Latest order"] == f"{want['last']:%Y-%m-%d}"
        assert r["Days between first and latest"] == (want["last"] - want["first"]).days
    # most orders first, then biggest spender
    assert [(r["Orders"], r["Total spent"]) for r in rows] == sorted(
        [(r["Orders"], r["Total spent"]) for r in rows], reverse=True)


def test_totals_by_state_match_pandas(full, sold):
    expected = sold.groupby("Ship State").agg(orders=("Order ID", "nunique"), customers=("Customer ID", "nunique"),
                                              units=("Qty", "sum"), sales=("Line Total", "sum"))
    rows = rows_of(full, "04_totals_by_state.sql")
    assert {r["State"] for r in rows} == set(expected.index) and len(rows) == 8
    for r in rows:
        want = expected.loc[r["State"]]
        assert (r["Orders"], r["Customers"], r["Units"]) == (want["orders"], want["customers"], want["units"])
        assert r["Sales"] == pytest.approx(want["sales"], abs=0.005)
    assert [r["Sales"] for r in rows] == sorted((r["Sales"] for r in rows), reverse=True)
    assert sum(r["Orders"] for r in rows) == sold["Order ID"].nunique()  # an order ships to one state
    assert sum(r["Sales"] for r in rows) == pytest.approx(sold["Line Total"].sum(), abs=0.005)


# ---------- the Excel workbook ----------

@pytest.fixture(scope="module")
def built(full):
    reports = sr.run_all(full)
    return reports, sr.build_workbook(reports, "Sample Store", "orders_export_messy.csv", AS_OF)


def test_workbook_has_a_sheet_per_query(built, tmp_path):
    reports, wb = built
    wb.save(tmp_path / "report.xlsx")
    loaded = load_workbook(tmp_path / "report.xlsx")
    assert loaded.sheetnames == ["Summary", "Sales by month", "Top products", "Repeat customers",
                                 "Totals by state", "Queries"]
    for report in reports[1:]:
        ws = loaded[report.title]
        assert [c.value for c in ws[1]] == report.columns
        assert ws.max_row == 1 + len(report.rows)
        assert ws.freeze_panes == "A2"


def test_workbook_formats_and_values(built):
    reports, wb = built
    months = wb["Sales by month"]
    assert months["A2"].value == "Apr 2026"
    assert months["D2"].number_format == sr.MONEY
    assert months["F2"].value is None and months["F3"].number_format == sr.CHANGE
    states = wb["Totals by state"]
    assert states["G2"].number_format == sr.PERCENT and states["A2"].value == "NSW"
    repeat = wb["Repeat customers"]
    assert repeat["B2"].value == "Benjamin Mitchell" and repeat["D2"].value == 6  # the most orders
    assert repeat["E2"].value == date(2026, 4, 4) and repeat["E2"].number_format == "dd/mm/yyyy"  # a date, not text
    assert wb["Repeat customers"].max_row == 1 + 102


def test_summary_sheet(built, sold):
    reports, wb = built
    ws = wb["Summary"]
    labels = {ws.cell(row=r, column=1).value: ws.cell(row=r, column=2).value for r in range(4, 12)}
    assert list(labels) == ["Orders", "Units sold", "Net sales", "Average order", "Customers who bought",
                            "Repeat customers", "Repeat customer rate", "Share of sales from repeat customers"]
    assert labels["Net sales"] == pytest.approx(sold["Line Total"].sum(), abs=0.005)
    assert labels["Orders"] == sold["Order ID"].nunique()
    assert "01 Apr 2026 to 30 Sep 2026" in ws["A2"].value and "19 Oct 2026" in ws["A2"].value
    assert len(ws._charts) == 2  # net sales by month, net sales by state


def test_queries_sheet_shows_the_sql(built):
    reports, wb = built
    ws = wb["Queries"]
    assert ws.max_row == 1 + len(reports)
    for r, report in enumerate(reports, start=2):
        assert ws.cell(row=r, column=2).value == report.file
        assert ws.cell(row=r, column=4).value == report.sql
    assert "LAG(cents) OVER" in ws["D3"].value


# ---------- the sample files and the loader ----------

def test_sample_customer_files_are_what_the_script_writes(tmp_path):
    customers, mapping = customers_demo.write_files(tmp_path)

    def text(path):  # git may check the files out with Windows line endings
        return path.read_bytes().replace(b"\r\n", b"\n")

    assert text(customers) == text(DEMO / "sample_data" / "customers.csv")
    assert text(mapping) == text(DEMO / "sample_data" / "order_customers.csv")


def test_orders_without_a_customer_are_refused(tmp_path):
    lines = sr.ORDER_CUSTOMERS.read_text(encoding="utf-8").splitlines()
    short = tmp_path / "short.csv"
    short.write_text("\n".join(lines[:-3]) + "\n", encoding="utf-8")
    with pytest.raises(ValueError, match=r"3 orders have no customer in short\.csv"):
        sr.build_database(":memory:", order_customers_csv=short)

    unknown = tmp_path / "unknown.csv"
    unknown.write_text("Order ID,Customer ID\n" + "\n".join(f"{o},C9999" for o, _ in
                       sr.read_csv(sr.ORDER_CUSTOMERS, ["Order ID", "Customer ID"])) + "\n", encoding="utf-8")
    with pytest.raises(ValueError, match="no customer C9999"):
        sr.build_database(":memory:", order_customers_csv=unknown)


def test_command_line_builds_both_files_and_can_be_rerun(tmp_path, capsys):
    for _ in range(2):  # the second run replaces the database instead of failing on existing tables
        sr.main(["-o", str(tmp_path / "out"), "--as-of", "2026-10-19", "--store", "Test Shop"])
    printed = capsys.readouterr().out
    assert "Net sales $25,546.25 from 519 orders; 102 of 376 customers ordered again." in printed
    assert load_workbook(tmp_path / "out" / "sql_report.xlsx")["Summary"]["A1"].value == "SQL sales report: Test Shop"
    database = sqlite3.connect(tmp_path / "out" / "store.db")
    assert database.execute("SELECT COUNT(*) FROM order_items").fetchone() == (797,)
    database.close()


def test_the_committed_report_is_what_a_run_writes(built):
    reports, wb = built

    def values(book):  # a saved workbook gives dates back as datetimes
        return {ws.title: [[c.value.date() if isinstance(c.value, datetime) else c.value for c in row]
                           for row in ws.iter_rows()] for ws in book}

    assert values(wb) == values(load_workbook(DEMO / "output" / "sql_report.xlsx"))
    database = sqlite3.connect(DEMO / "output" / "store.db")
    assert database.execute("SELECT COUNT(*) FROM orders").fetchone() == (572,)
    database.close()
