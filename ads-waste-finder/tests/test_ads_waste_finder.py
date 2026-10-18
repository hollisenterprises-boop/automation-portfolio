"""Tests for the ads waste finder. Run from the demo folder with:  pytest"""

import sys
from datetime import date
from io import BytesIO
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import ads_waste_finder as awf  # noqa: E402

AMAZON_HEADERS = ["Date", "Campaign Name", "Ad Group Name", "Targeting", "Match Type",
                  "Customer Search Term", "Impressions", "Clicks", "Spend",
                  "7 Day Total Sales ", "7 Day Total Orders (#)"]


def raw_report(*rows):
    """A tiny report as it comes out of a CSV: Amazon headers, every cell text.
    Each row: date, campaign, ad group, targeting, match type, search term,
    impressions, clicks, spend, sales, orders."""
    return pd.DataFrame([[str(v) for v in row] for row in rows], columns=AMAZON_HEADERS)


def analysed(*rows):
    return awf.analyse(awf.read_rows(raw_report(*rows)))


# ---------- columns ----------

@pytest.mark.parametrize("header, key", [
    ("7 Day Total Orders (#)", "7 day total orders"),
    ("  7 Day Total Sales ", "7 day total sales"),
    ("Customer_Search-Term", "customer search term"),
    ("CAMPAIGN NAME", "campaign name"),
])
def test_normalise_header(header, key):
    assert awf.normalise_header(header) == key


def test_columns_match_despite_odd_headers():
    found = awf.match_columns([" Campaign Name ", "AD GROUP NAME", "Customer Search Term",
                               "Impressions", "Clicks", "Spend", "7 Day Total Sales ",
                               "7 Day Total Orders (#)", "Currency"])
    assert found["Campaign"] == " Campaign Name "
    assert found["Ad Group"] == "AD GROUP NAME"
    assert found["Sales"] == "7 Day Total Sales "
    assert found["Orders"] == "7 Day Total Orders (#)"
    assert "Date" not in found and "Targeting" not in found  # optional


def test_14_day_and_plain_total_columns_are_accepted():
    found = awf.match_columns(["Campaign Name", "Ad Group Name", "Customer Search Term",
                               "Impressions", "Clicks", "Spend", "14 Day Total Sales",
                               "Total Orders"])
    assert found["Sales"] == "14 Day Total Sales"
    assert found["Orders"] == "Total Orders"


def test_7_day_columns_win_when_both_are_there():
    found = awf.match_columns(["Campaign Name", "Ad Group Name", "Customer Search Term",
                               "Impressions", "Clicks", "Spend", "14 Day Total Sales",
                               "7 Day Total Sales ", "7 Day Total Orders (#)"])
    assert found["Sales"] == "7 Day Total Sales "


def test_missing_columns_are_named():
    with pytest.raises(ValueError, match="missing these columns: Spend, 7 Day Total Sales"):
        awf.match_columns(["Campaign Name", "Ad Group Name", "Customer Search Term",
                           "Impressions", "Clicks", "7 Day Total Orders (#)"])


# ---------- single values ----------

@pytest.mark.parametrize("text, amount", [
    ("1.23", 1.23), ("$1.23", 1.23), ("A$1.23", 1.23), ("AU$1.23", 1.23), ("AUD 1.23", 1.23),
    ("1,234.56", 1234.56), ("$ -", 0.0), ("", 0.0), ("  ", 0.0), (None, 0.0), (2.5, 2.5),
])
def test_parse_money(text, amount):
    assert awf.parse_money(text) == amount


def test_parse_money_refuses_text():
    with pytest.raises(ValueError):
        awf.parse_money("n/a")


@pytest.mark.parametrize("text, count", [("12", 12), ("1,234", 1234), ("12.0", 12), ("", 0)])
def test_parse_count(text, count):
    assert awf.parse_count(text) == count


def test_unreadable_number_names_the_column_and_row():
    raw = raw_report(("2026-09-01", "Auto", "All", "close-match", "-", "glass jars",
                      100, 3, "abc", 0, 0))
    with pytest.raises(ValueError, match=r"'Spend'.*row 2"):
        awf.read_rows(raw)


@pytest.mark.parametrize("text", ["Sep 01, 2026", "2026-09-01", "01/09/2026", "2026-09-01 00:00:00"])
def test_parse_date(text):
    assert awf.parse_date(text) == date(2026, 9, 1)


def test_unreadable_date_is_none():
    assert awf.parse_date("next Tuesday") is None


@pytest.mark.parametrize("term, expected", [
    ("b08kx7q2lm", True), ("B08KX7Q2LM", True), (" b0c3m5r7tq ", True),
    ("glass jars", False), ("b08kx7q2l", False), ("b08kx7q2lmx", False), ("a08kx7q2lm", False),
])
def test_is_asin(term, expected):
    assert awf.is_asin(term) is expected


# ---------- rules ----------

@pytest.mark.parametrize("orders, clicks, spend, sales, outcome", [
    (0, 9, 14.99, 0, awf.KEEP_WATCHING),     # under both no-sale limits
    (0, 10, 5.00, 0, awf.NO_SALES),          # 10 clicks is enough to judge
    (0, 3, 15.00, 0, awf.NO_SALES),          # so is $15
    (1, 20, 45.00, 100, awf.KEEP_WATCHING),  # ACOS 45% is not above 1.5 x target
    (1, 20, 45.10, 100, awf.TOO_EXPENSIVE),  # 45.1% is
    (1, 20, 30.00, 100, awf.PROFITABLE),     # exactly on target
    (1, 20, 30.10, 100, awf.KEEP_WATCHING),  # 30.1%: near target
    (1, 2, 0.90, 3.00, awf.PROFITABLE),      # 0.9 / 3 is 0.30000000000000004 in floating point
])
def test_classify_at_the_boundaries(orders, clicks, spend, sales, outcome):
    assert awf.classify(orders, clicks, spend, sales) == outcome


def test_classify_uses_the_thresholds_given():
    assert awf.classify(0, 5, 2.00, 0, min_clicks=5) == awf.NO_SALES
    assert awf.classify(0, 4, 8.00, 0, min_spend=8.0) == awf.NO_SALES
    assert awf.classify(1, 20, 25.00, 100, target_acos=0.20) == awf.KEEP_WATCHING  # limit is 30%
    assert awf.classify(1, 20, 31.00, 100, target_acos=0.20) == awf.TOO_EXPENSIVE


@pytest.mark.parametrize("cpc, acos, bid", [
    (1.00, 0.60, 0.50),
    (1.23, 0.75, 0.50),  # 0.492 rounds to the nearest 5 cents
    (1.10, 0.80, 0.40),  # 0.4125 rounds down
    (0.95, 0.60, 0.50),  # 0.475 is exactly halfway and rounds up
    (0.40, 0.90, 0.20),  # 0.133 is below the 20 cent minimum
])
def test_suggest_bid(cpc, acos, bid):
    assert awf.suggest_bid(cpc, acos, 0.30) == bid


def test_action_for_each_outcome():
    terms = analysed(
        ("2026-09-01", "Auto", "All", "loose-match", "-", "mason jars", 500, 12, 10.00, 0, 0),
        ("2026-09-01", "Auto", "All", "substitutes", "-", "B08KX7Q2LM", 500, 12, 10.00, 0, 0),
        ("2026-09-01", "Auto", "All", "loose-match", "-", "food containers", 500, 20, 24.00, 40.00, 1),
        ("2026-09-01", "Auto", "All", "loose-match", "-", "silicone lids", 50, 2, 1.50, 0, 0),
        ("2026-09-01", "Auto", "All", "close-match", "-", "glass jars", 50, 10, 10.00, 25.00, 1),
    ).set_index("Search Term")
    assert terms.loc["mason jars", "Action"] == awf.NEGATIVE_EXACT
    assert terms.loc["B08KX7Q2LM", "Action"] == awf.NEGATIVE_PRODUCT
    expensive = terms.loc["food containers"]  # ACOS 60%
    assert (expensive["Outcome"], expensive["Action"]) == (awf.TOO_EXPENSIVE, awf.LOWER_BID)
    assert expensive["Suggested Bid"] == 0.60  # CPC 1.20 x 30% / 60%
    assert terms.loc["silicone lids", "Action"] == awf.NOT_ENOUGH_DATA
    assert terms.loc["glass jars", "Action"] == awf.NEAR_TARGET  # ACOS 40%


EXACT_ROW = ("2026-09-01", "Exact", "Meal Prep", "Glass  Meal Prep Containers", "EXACT",
             "glass meal prep containers", 100, 10, 10.00, 100.00, 2)


def test_exact_keywords_ignore_case_and_spacing():
    rows = awf.read_rows(raw_report(
        EXACT_ROW,
        ("2026-09-01", "Broad", "Core", "meal prep", "BROAD", "glass jars", 100, 10, 10.00, 100.00, 2),
    ))
    assert awf.exact_keywords(rows) == {"glass meal prep containers"}


def test_winner_needs_two_orders_and_no_exact_keyword_yet():
    terms = analysed(
        EXACT_ROW,
        ("2026-09-01", "Broad", "Core", "meal prep", "BROAD",
         "glass meal prep containers", 100, 10, 10.00, 100.00, 2),
        ("2026-09-01", "Broad", "Core", "glass containers", "BROAD",
         "glass lunch containers", 100, 10, 10.00, 100.00, 2),
        ("2026-09-01", "Broad", "Core", "glass containers", "BROAD",
         "glass bento box", 100, 10, 10.00, 100.00, 1),
    ).set_index(["Campaign", "Search Term"])
    assert terms.loc[("Exact", "glass meal prep containers"), "Action"] == awf.KEEP_RUNNING
    # Already an exact keyword in another campaign, so not a new winner here either.
    assert terms.loc[("Broad", "glass meal prep containers"), "Action"] == awf.KEEP_RUNNING
    assert terms.loc[("Broad", "glass lunch containers"), "Action"] == awf.ADD_EXACT
    assert terms.loc[("Broad", "glass bento box"), "Action"] == awf.KEEP_RUNNING  # 1 order


# ---------- adding up ----------

def test_days_are_added_up_per_campaign_ad_group_and_search_term():
    rows = awf.read_rows(raw_report(
        ("2026-09-01", "Auto", "All", "close-match", "-", "glass jars", 100, 3, "$1.10", 0, 0),
        ("2026-09-02", "Auto", "All", "loose-match", "-", "glass jars", 50, 2, "A$2.20", "$29.95", 1),
        ("2026-09-03", "Auto", "All", "close-match", "-", "glass jars", "1,000", 5, "0.30", "", ""),
        ("2026-09-01", "Auto", "Other", "close-match", "-", "glass jars", 10, 1, "0.50", 0, 0),
    ))
    terms = awf.aggregate_terms(rows)
    assert len(terms) == 2
    jars = terms[terms["Ad Group"] == "All"].iloc[0]
    assert (jars["Impressions"], jars["Clicks"], jars["Orders"]) == (1150, 10, 1)
    assert jars["Spend"] == 3.60  # 1.10 + 2.20 + 0.30 without floating point dust
    assert jars["Sales"] == 29.95
    assert jars["Targeting"] == "close-match"  # the first one is kept
    assert jars["CPC"] == 0.36
    assert jars["ACOS"] == 0.120


def test_summary_adds_up_and_scales_waste_to_30_days():
    rows = awf.read_rows(raw_report(
        ("2026-09-01", "Auto", "All", "loose-match", "-", "mason jars", 500, 12, 20.00, 0, 0),
        ("2026-09-10", "Auto", "All", "close-match", "-", "glass jars", 500, 10, 10.00, 50.00, 2),
        ("2026-09-05", "Auto", "All", "loose-match", "-", "food containers", 500, 10, 30.00, 50.00, 1),
    ))
    terms = awf.analyse(rows)
    summary = awf.summarise(terms, rows)
    assert summary["days"] == 10
    assert (summary["spend"], summary["sales"], summary["wasted"]) == (60.00, 100.00, 20.00)
    assert summary["by_outcome"]["Spend"].sum() == pytest.approx(summary["spend"])
    assert summary["wasted_share"] == pytest.approx(1 / 3)
    assert summary["wasted_per_30_days"] == pytest.approx(60.00)
    assert (summary["negatives"], summary["bids"], summary["winners"]) == (1, 1, 1)


def test_report_without_dates_still_works():
    raw = raw_report(("", "Auto", "All", "loose-match", "-", "mason jars", 500, 12, 20.00, 0, 0))
    rows = awf.read_rows(raw.drop(columns=["Date"]))
    terms = awf.analyse(rows)
    summary = awf.summarise(terms, rows)
    assert summary["days"] is None and summary["wasted_per_30_days"] is None
    wb = awf.build_workbook(terms, summary, 0.30, 10, 15.0, "Test Store", "report.csv")
    assert wb["Summary"]["A4"].value == "Ad spend"


# ---------- files ----------

def test_workbook_sheets_and_chart_colours():
    rows = awf.read_rows(raw_report(
        ("2026-09-01", "Auto", "All", "loose-match", "-", "mason jars", 500, 12, 20.00, 0, 0),
        ("2026-09-02", "Auto", "All", "close-match", "-", "glass jars", 500, 10, 10.00, 50.00, 2),
    ))
    terms = awf.analyse(rows)
    wb = awf.build_workbook(terms, awf.summarise(terms, rows), 0.30, 10, 15.0,
                            "Test Store", "report.csv")
    assert wb.sheetnames == ["Summary", "Negatives to add", "Bids to lower",
                             "Winners to add", "All search terms"]
    assert wb["Negatives to add"]["C2"].value == "mason jars"
    assert wb["Winners to add"]["C2"].value == "glass jars"
    points = wb["Summary"]._charts[0].series[0].dPt
    assert [p.graphicalProperties.solidFill.srgbClr for p in points] == \
        ["D03B3B", "EC835A", "898781", "0CA30C"]
    wb.save(BytesIO())


def test_reads_an_excel_report(tmp_path):
    path = tmp_path / "report.xlsx"
    pd.DataFrame({
        "Date": pd.to_datetime(["2026-09-01", "2026-09-02"]),
        "Campaign Name": ["Auto", "Auto"],
        "Ad Group Name": ["All", "All"],
        "Customer Search Term": ["glass jars", "glass jars"],
        "Impressions": [100, 200],
        "Clicks": [3, 4],
        "Spend": [1.5, 2.25],
        "7 Day Total Sales ": [0, 29.95],
        "7 Day Total Orders (#)": [0, 1],
    }).to_excel(path, index=False)
    rows = awf.read_rows(awf.load_report(path))
    assert rows["Date"].tolist() == [date(2026, 9, 1), date(2026, 9, 2)]
    assert (rows["Clicks"].sum(), rows["Spend"].sum(), rows["Sales"].sum()) == (7, 3.75, 29.95)
