"""Make the sample Amazon search term report used by the ads waste finder demo.

    python make_sample_data.py

Invents 30 days of daily Sponsored Products search term rows (September 2026)
for a fictional Australian store selling glass food containers. Each search's
clicks and orders for the month are set in TERMS below, so the story never
changes; the seeded random generator only spreads them over the days. The
headers copy Amazon's, trailing spaces and "(#)" included. All data is made up.
"""

from __future__ import annotations

import argparse
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 2026
FIRST_DAY = date(2026, 9, 1)
DAYS = 30
OUTPUT = Path(__file__).resolve().parent / "sample_data" / "search_term_report.csv"

# Campaign name, ad group name, match type.
AUTO = ("SP Auto - Glass Containers", "Auto - All Sets", "-")
BROAD = ("SP Broad - Glass Containers", "Broad - Core Terms", "BROAD")
EXACT = ("SP Exact - Meal Prep", "Exact - Meal Prep", "EXACT")

# (campaign, targeting, search term, clicks in the month, orders in the month,
#  average cost per click, price per unit, click-through rate).
# Exact keywords target their own text, so their targeting is None.
TERMS = [
    # Strong sellers that already have an exact keyword
    (EXACT, None, "glass meal prep containers", 150, 22, 1.25, 44.95, 0.011),
    (EXACT, None, "meal prep containers glass", 72, 9, 1.15, 44.95, 0.009),
    (EXACT, None, "glass meal prep containers 3 compartment", 34, 5, 1.05, 49.95, 0.010),
    (AUTO, "close-match", "glass meal prep containers", 48, 6, 1.10, 44.95, 0.008),
    (BROAD, "meal prep containers", "glass meal prep containers", 40, 5, 1.20, 44.95, 0.008),
    # Strong sellers found by auto and broad, not exact keywords yet
    (AUTO, "close-match", "glass food storage containers", 96, 12, 1.05, 44.95, 0.009),
    (AUTO, "close-match", "borosilicate glass containers", 42, 5, 0.95, 54.95, 0.008),
    (BROAD, "food containers", "glass food containers", 130, 15, 1.15, 44.95, 0.008),
    (BROAD, "glass containers", "glass containers with lids", 100, 13, 1.05, 34.95, 0.009),
    (BROAD, "glass containers", "glass containers with bamboo lids", 36, 5, 0.95, 54.95, 0.008),
    # Generic searches that sell, but at a high cost
    (EXACT, None, "meal prep containers", 92, 4, 1.45, 44.95, 0.006),
    (BROAD, "food containers", "food containers", 84, 3, 1.25, 34.95, 0.005),
    (AUTO, "loose-match", "food containers", 88, 3, 1.10, 34.95, 0.005),
    (BROAD, "meal prep containers", "meal prep containers", 50, 2, 1.30, 44.95, 0.005),
    # Close to target
    (AUTO, "close-match", "glass lunch containers", 40, 2, 0.95, 49.95, 0.007),
    (BROAD, "glass jars", "glass jars with lids", 46, 4, 1.00, 29.95, 0.007),
    (BROAD, "food containers", "food storage containers", 30, 2, 1.20, 44.95, 0.006),
    # Too few clicks to judge yet
    (EXACT, None, "glass meal prep containers 2 compartment", 7, 0, 1.05, 49.95, 0.008),
    (AUTO, "substitutes", "b0c3m5r7tq", 6, 0, 0.80, 44.95, 0.003),
    (AUTO, "complements", "silicone lids", 7, 0, 0.65, 19.95, 0.004),
    (BROAD, "glass containers", "glass containers for fridge", 8, 0, 0.95, 34.95, 0.006),
    (BROAD, "glass containers", "airtight glass containers", 9, 1, 1.00, 44.95, 0.007),
    (BROAD, "meal prep containers", "meal prep containers glass 5 pack", 5, 0, 1.00, 44.95, 0.006),
    # Clicks that never sell
    (AUTO, "loose-match", "plastic food containers", 80, 0, 0.95, 34.95, 0.004),
    (AUTO, "loose-match", "kids lunch box", 60, 0, 0.90, 34.95, 0.004),
    (AUTO, "loose-match", "mason jars", 72, 0, 0.85, 29.95, 0.004),
    (AUTO, "substitutes", "b08kx7q2lm", 48, 0, 0.80, 44.95, 0.003),
    (AUTO, "substitutes", "b09tz4wn8r", 32, 0, 0.75, 44.95, 0.003),
    (AUTO, "complements", "freezer bags", 28, 0, 0.70, 19.95, 0.004),
    (BROAD, "food containers", "plastic food containers", 66, 0, 1.10, 34.95, 0.004),
    (BROAD, "meal prep containers", "disposable meal prep containers", 68, 0, 0.95, 44.95, 0.004),
    (BROAD, "glass jars", "glass jars for candles", 64, 0, 0.85, 29.95, 0.004),
]

WEEKDAY_TRAFFIC = [1.10, 1.00, 0.95, 0.95, 0.90, 1.05, 1.15]  # Monday first


def daily_rows(rng, campaign, targeting, term, clicks, orders, cpc, price, ctr) -> list[dict]:
    """Spread one search's month over the days; a day with no clicks gets no row."""
    name, ad_group, match_type = campaign
    days = [FIRST_DAY + timedelta(days=n) for n in range(DAYS)]
    weights = np.array([WEEKDAY_TRAFFIC[d.weekday()] for d in days]) * rng.lognormal(0, 0.35, DAYS)
    day_clicks = rng.multinomial(clicks, weights / weights.sum())
    # Pick which of the month's clicks turned into orders.
    order_days = rng.choice(np.repeat(np.arange(DAYS), day_clicks), size=orders, replace=False)
    day_orders = np.bincount(order_days, minlength=DAYS)

    rows = []
    for day, n_clicks, n_orders in zip(days, day_clicks, day_orders):
        if n_clicks == 0:
            continue
        spend = round(n_clicks * cpc * rng.uniform(0.85, 1.15), 2)
        impressions = max(int(n_clicks), round(n_clicks / ctr * rng.lognormal(0, 0.2)))
        units = n_orders + rng.binomial(n_orders, 0.1)  # now and then two in one order
        sales = round(units * price, 2)
        rows.append({
            "Date": day,
            "Currency": "AUD",
            "Campaign Name": name,
            "Ad Group Name": ad_group,
            "Targeting": targeting or term,
            "Match Type": match_type,
            "Customer Search Term": term,
            "Impressions": impressions,
            "Clicks": n_clicks,
            "Click-Thru Rate (CTR)": f"{n_clicks / impressions:.2%}",
            "Cost Per Click (CPC)": f"{spend / n_clicks:.2f}",
            "Spend": f"{spend:.2f}",
            "7 Day Total Sales ": f"{sales:.2f}",
            "Total Advertising Cost of Sales (ACOS) ": f"{spend / sales:.2%}" if sales else "",
            "Total Return on Advertising Spend (ROAS)": f"{sales / spend:.2f}",
            "7 Day Total Orders (#)": n_orders,
            "7 Day Total Units (#)": units,
            "7 Day Conversion Rate": f"{n_orders / n_clicks:.2%}",
        })
    return rows


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("-o", "--output", type=Path, default=OUTPUT)
    args = parser.parse_args(argv)

    rng = np.random.default_rng(SEED)
    rows = [row for spec in TERMS for row in daily_rows(rng, *spec)]
    report = pd.DataFrame(rows).sort_values(
        ["Date", "Campaign Name", "Customer Search Term"], kind="stable")
    report["Date"] = report["Date"].map(lambda d: f"{d:%b %d, %Y}")  # e.g. Sep 01, 2026
    args.output.parent.mkdir(parents=True, exist_ok=True)
    report.to_csv(args.output, index=False, lineterminator="\n")

    spend = report["Spend"].astype(float).sum()
    print(f"{len(report):,} daily rows, {len(TERMS)} search terms, ${spend:,.2f} ad spend. "
          f"Written to {args.output}")


if __name__ == "__main__":
    main()
