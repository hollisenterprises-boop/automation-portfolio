"""Fast tests for the price tracker. No network: pages come from a hand-written fixture."""

import random
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))

import price_tracker as pt  # noqa: E402
from make_demo_history import make_history  # noqa: E402

PAGE_URL = "https://books.toscrape.com/catalogue/page-2.html"
LISTING = (HERE / "fixtures" / "listing.html").read_text(encoding="utf-8")
LAST_PAGE = """<article class="product_pod"><p class="star-rating Two"></p>
<h3><a href="tea-for-one_77/index.html" title="Tea for One">Tea for One</a></h3>
<p class="price_color">£9.99</p><p class="instock availability">In stock</p></article>"""


def snapshot(*rows):
    """A tiny snapshot from (url, title, price, in_stock) rows."""
    return pd.DataFrame([
        {"url": url, "title": title, "price": price, "currency": "£", "rating": 3,
         "stock": "In stock" if in_stock else "Out of stock", "in_stock": in_stock,
         "scraped_at": "2026-10-18 09:00:00"}
        for url, title, price, in_stock in rows
    ], columns=pt.SNAPSHOT_COLUMNS)


OLD = snapshot(("u/apple", "Apple", 10.00, True), ("u/banana", "Banana", 20.00, True),
               ("u/cherry", "Cherry", 15.00, False), ("u/date", "Date", 5.00, True))
NEW = snapshot(("u/apple", "Apple", 8.00, True), ("u/banana", "Banana", 25.00, False),
               ("u/cherry", "Cherry", 15.00, True), ("u/elder", "Elderberry", 7.50, True))


# ---------- reading pages ----------

def test_listing_page_gives_every_product():
    products, _ = pt.parse_listing(LISTING, PAGE_URL)
    assert [p["title"] for p in products] == [
        "The Blue Teapot Handbook",
        "Gardening on a Windy Balcony: A Practical Guide for Small Spaces",  # full title, not the cut link text
        "A Short History of Paperclips",
    ]
    assert products[0]["url"] == "https://books.toscrape.com/catalogue/the-blue-teapot-handbook_901/index.html"
    assert [p["price"] for p in products] == [12.50, 45.00, 7.99]
    assert {p["currency"] for p in products} == {"£"}
    assert [p["rating"] for p in products] == [4, 1, 5]
    assert [p["stock"] for p in products] == ["In stock", "Out of stock", "In stock"]
    assert [p["in_stock"] for p in products] == [True, False, True]


def test_next_page_link_is_made_absolute():
    _, next_url = pt.parse_listing(LISTING, PAGE_URL)
    assert next_url == "https://books.toscrape.com/catalogue/page-3.html"


def test_last_page_has_no_next_page():
    products, next_url = pt.parse_listing(LAST_PAGE, PAGE_URL)
    assert next_url is None
    assert products[0]["title"] == "Tea for One"


@pytest.mark.parametrize("text, expected", [
    ("£51.77", (51.77, "£")),
    ("£1,299.00", (1299.0, "£")),
    ("$8", (8.0, "$")),
    ("Call for price", (None, "")),
])
def test_parse_price(text, expected):
    assert pt.parse_price(text) == expected


class FakeResponse:
    def __init__(self, text):
        self.text, self.encoding = text, None

    def raise_for_status(self):
        pass


class FakeSession:
    """Stands in for requests.Session: serves saved pages and records each request."""

    def __init__(self, pages):
        self.pages, self.requests, self.responses = pages, [], []

    def get(self, url, timeout):
        self.requests.append((url, timeout))
        self.responses.append(FakeResponse(self.pages[url]))
        return self.responses[-1]


def test_scrape_waits_between_pages_and_stops_after_the_last():
    session = FakeSession({pt.LISTING_URL.format(n=1): LISTING, pt.LISTING_URL.format(n=2): LAST_PAGE})
    waits = []
    products = pt.scrape_catalogue(pages=5, delay=1.5, session=session, sleep=waits.append)
    assert [url for url, _ in session.requests] == [pt.LISTING_URL.format(n=1), pt.LISTING_URL.format(n=2)]
    assert {timeout for _, timeout in session.requests} == {pt.TIMEOUT}
    assert waits == [1.5]  # one pause, between the two requests
    assert {r.encoding for r in session.responses} == {"utf-8"}
    assert list(products.columns) == pt.SNAPSHOT_COLUMNS
    assert len(products) == 4


# ---------- snapshots ----------

def test_snapshot_round_trip_keeps_pound_sign_and_types(tmp_path):
    path = pt.save_snapshot(NEW, pt.snapshot_path(tmp_path, datetime(2026, 10, 18, 9, 5)))
    assert path.name == "2026-10-18_0905.csv"
    loaded = pt.load_snapshot(path)
    assert loaded["currency"].tolist() == ["£"] * 4
    assert loaded["price"].tolist() == [8.0, 25.0, 15.0, 7.5]
    assert loaded["in_stock"].tolist() == [True, False, True, True]
    assert loaded["rating"].tolist() == [3] * 4


def test_previous_snapshot_is_the_latest_one_before(tmp_path):
    for name in ["2026-10-01_0900.csv", "2026-10-11_0900.csv", "2026-10-18_0900.csv", "notes.csv"]:
        (tmp_path / name).write_text("url\n")
    assert [p.name for p in pt.list_snapshots(tmp_path)] == [
        "2026-10-01_0900.csv", "2026-10-11_0900.csv", "2026-10-18_0900.csv"]
    assert pt.previous_snapshot(tmp_path, tmp_path / "2026-10-18_0900.csv").name == "2026-10-11_0900.csv"
    assert pt.previous_snapshot(tmp_path, tmp_path / "2026-10-01_0900.csv") is None


# ---------- comparing ----------

def test_price_changes_and_their_maths():
    moves = pt.compare(OLD, NEW)["price_changes"]
    assert moves["Product"].tolist() == ["Apple", "Banana"]  # Cherry's price didn't move; sorted by change %
    assert moves["Old price"].tolist() == [10.0, 20.0]
    assert moves["New price"].tolist() == [8.0, 25.0]
    assert moves["Change"].tolist() == [-2.0, 5.0]
    assert moves["Change %"].tolist() == pytest.approx([-0.20, 0.25])


def test_new_and_gone_products():
    changes = pt.compare(OLD, NEW)
    assert changes["new_products"]["Product"].tolist() == ["Elderberry"]
    assert changes["gone"]["Product"].tolist() == ["Date"]
    assert changes["gone"]["Last price"].tolist() == [5.0]


def test_stock_changes():
    stock = pt.compare(OLD, NEW)["stock_changes"]
    assert stock[["Product", "Change", "Was", "Now"]].values.tolist() == [
        ["Cherry", "Back in stock", "Out of stock", "In stock"],
        ["Banana", "Out of stock", "In stock", "Out of stock"],
    ]


def test_summary_counts():
    summary = pt.summarise(NEW, pt.compare(OLD, NEW))
    assert summary["products"] == 4
    assert summary["average_price"] == pytest.approx((8 + 25 + 15 + 7.5) / 4)
    assert (summary["drops"], summary["rises"], summary["new"], summary["gone"]) == (1, 1, 1, 1)
    assert (summary["back_in_stock"], summary["out_of_stock"]) == (1, 1)


def test_short_title_for_chart_labels():
    assert pt.short_title("Tea for One") == "Tea for One"
    cut = pt.short_title("Gardening on a Windy Balcony: A Practical Guide for Small Spaces")
    assert cut.endswith("…") and len(cut) <= 30


# ---------- the report ----------

SHEETS = ["Summary", "Price changes", "New products", "Gone", "Stock changes", "All products"]


def test_report_has_every_sheet_and_the_numbers(tmp_path):
    changes = pt.compare(OLD, NEW)
    wb = pt.build_workbook(NEW, changes, pt.summarise(NEW, changes),
                           datetime(2026, 10, 18, 9, 0), datetime(2026, 10, 11, 9, 0))
    assert wb.sheetnames == SHEETS
    ws = wb["Summary"]
    kpis = {label: value for label, value in ws.iter_rows(min_row=4, max_row=11, max_col=2, values_only=True)}
    assert kpis["Price drops"] == 1 and kpis["Price rises"] == 1 and kpis["Went out of stock"] == 1
    assert len(ws._charts) == 1
    assert wb["Price changes"]["E2"].number_format == pt.PERCENT
    assert wb["Price changes"]["F2"].hyperlink.target == "u/apple"
    wb.save(tmp_path / "report.xlsx")


def test_report_without_an_earlier_snapshot_says_so():
    wb = pt.build_workbook(NEW, None, pt.summarise(NEW, None), datetime(2026, 10, 18, 9, 0))
    assert wb.sheetnames == SHEETS
    assert "nothing earlier to compare" in wb["Summary"]["A2"].value
    assert wb["Price changes"]["A2"].value == "No earlier snapshot to compare with yet."
    assert wb["All products"].max_row == 5  # header + 4 products


# ---------- demo history ----------

def test_demo_history_gives_the_promised_changes():
    latest = snapshot(*[(f"u/{i}", f"Book {i}", 10.0 + i, True) for i in range(30)])
    history = make_history(latest, random.Random(7), days=7)
    changes = pt.compare(history, latest)
    moves = changes["price_changes"]["Change %"]
    assert len(moves) == 12
    assert moves.abs().between(0.049, 0.251).all()
    assert (moves < 0).sum() == 6
    assert len(changes["new_products"]) == 4 and changes["gone"].empty
    assert changes["stock_changes"]["Change"].tolist() == ["Back in stock"] * 2
    assert set(history["scraped_at"]) == {"2026-10-11 09:00:00"}
