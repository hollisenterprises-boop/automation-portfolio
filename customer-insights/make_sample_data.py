"""Cut the small test sample from the full UCI Online Retail file.

    python make_sample_data.py "path/to/Online Retail.xlsx"

Writes sample_data/online_retail_sample.csv: every line, unchanged, for about
300 real customers, plus a few real lines with no customer ID, so the tests run
without the 23 MB download. Customers are drawn at random (fixed seed) across
one to seven-plus orders, from those with at most 25 lines to keep the file
small, plus a few picked because they show a cleaning rule or an odd case.
"""

from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd

import analyse as an

HERE = Path(__file__).parent
OUTPUT = HERE / "sample_data" / "online_retail_sample.csv"
SEED = 2011
MAX_LINES = 25
BANDS = [(1, 1, 100), (2, 2, 80), (3, 3, 55), (4, 6, 45), (7, None, 10)]  # orders from, to, customers
NAMED = {
    12346: "bought 74,215 storage jars and cancelled them the same day (net spend £0)",
    15098: "a £38,970 order refunded with a Manual line (M), not a product return",
    16446: "an order of 80,995 paper birdies cancelled 12 minutes later",
}


def fewest_lines(ids, line_counts: pd.Series) -> int:
    """The customer with the fewest lines (lowest ID on a tie)."""
    counts = line_counts.reindex(pd.Index(ids).unique()).sort_index(kind="stable")
    return int(counts.sort_values(kind="stable").index[0])


def pick_customers(lines: pd.DataFrame) -> list[int]:
    known = lines[lines["CustomerID"].notna()]
    line_counts = known.groupby("CustomerID").size()
    sales = known[~an.is_return(known)]
    orders = sales.groupby("CustomerID")["InvoiceNo"].nunique().reindex(line_counts.index, fill_value=0)
    code = known["StockCode"].str.upper()

    shows = {  # one customer for each rule or case, the one with the fewest lines
        "exact duplicate line": known.loc[known.duplicated(subset=an.COLUMNS, keep=False), "CustomerID"],
        "postage": known.loc[code == "POST", "CustomerID"],
        "carriage": known.loc[code == "C2", "CustomerID"],
        "bank charges": known.loc[code == "BANK CHARGES", "CustomerID"],
        "cushion pads at £0.001": known.loc[code == "PADS", "CustomerID"],
        "discount": known.loc[code == "D", "CustomerID"],
        "free item": sales.loc[sales["UnitPrice"] <= 0, "CustomerID"],
        "returns only": orders[orders == 0].index.to_series(),
    }
    chosen = {fewest_lines(ids, line_counts) for ids in shows.values()} | set(NAMED)

    rng = np.random.default_rng(SEED)
    small = line_counts[line_counts <= MAX_LINES].index
    for low, high, count in BANDS:
        in_band = orders[(orders >= low) & ((orders <= high) if high else True)].index
        pool = sorted(set(in_band) & set(small) - chosen)
        chosen |= {int(c) for c in rng.choice(pool, size=min(count, len(pool)), replace=False)}
    return sorted(chosen)


def no_customer_lines(lines: pd.DataFrame) -> pd.Index:
    """A few real lines with no customer ID: small orders, stock write-offs and fees."""
    rng = np.random.default_rng(SEED)
    unknown = lines[lines["CustomerID"].isna()]
    sizes = unknown.groupby("InvoiceNo").size()
    small = sorted(sizes[(sizes <= 4) & ~sizes.index.str.startswith("C")].index)
    invoices = rng.choice(small, size=3, replace=False)
    picks = [unknown.index[unknown["InvoiceNo"].isin(invoices)]]
    code = unknown["StockCode"].str.upper()
    write_offs = unknown[(unknown["Quantity"] < 0) & (unknown["UnitPrice"] == 0)].index
    picks.append(pd.Index(rng.choice(write_offs, size=6, replace=False)))
    for wanted in ("AMAZONFEE", "B", "DOT"):
        picks.append(unknown.index[code == wanted][:1])
    picks.append(unknown.index[unknown["StockCode"].str.match(an.GIFT_VOUCHER)][:1])
    return picks[0].append(picks[1:])


def sample(raw: pd.DataFrame) -> pd.DataFrame:
    lines = an.prepare(raw)
    customers = pick_customers(lines)
    rows = lines.index[lines["CustomerID"].isin(customers)].append(no_customer_lines(lines))
    return raw.loc[sorted(set(rows))]


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("data", type=Path, nargs="?", default=an.DEFAULT_DATA, help="the full Online Retail file")
    parser.add_argument("-o", "--output", type=Path, default=OUTPUT)
    args = parser.parse_args(argv)
    raw = an.read_raw(args.data)
    picked = sample(raw)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    picked.to_csv(args.output, index=False, date_format="%Y-%m-%d %H:%M", float_format="%.15g")
    customers = picked["CustomerID"].nunique()
    size = args.output.stat().st_size / 1024
    print(f"{len(picked):,} lines, {customers} customers -> {args.output} ({size:,.0f} KB)")


if __name__ == "__main__":
    main()
