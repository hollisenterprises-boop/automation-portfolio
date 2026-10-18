"""Write an earlier, edited copy of the newest snapshot so the demo report has changes to show.

    python make_demo_history.py --days 7 --seed 7

books.toscrape.com is a practice shop whose prices never change, so two real
runs would always report "no changes". This script takes the newest snapshot
and writes a copy dated 7 days earlier with seeded edits: 12 prices moved 5-25%
(half down, half up), 4 products left out (they show as new) and 2 products
marked out of stock (they show as back in stock). The copy is generated demo
data, not a real price history. Afterwards run:

    python price_tracker.py --compare-only
"""

from __future__ import annotations

import argparse
import random
from datetime import timedelta
from pathlib import Path

import pandas as pd

from price_tracker import HERE, list_snapshots, load_snapshot, save_snapshot, snapshot_path, snapshot_time

PRICE_CHANGES = 12
LEFT_OUT = 4
OUT_OF_STOCK = 2


def make_history(latest: pd.DataFrame, rng: random.Random, days: int = 7) -> pd.DataFrame:
    """An edited copy of `latest` that looks like a snapshot taken `days` earlier."""
    history = latest.copy()
    candidates = sorted(history.loc[history["price"].notna() & history["in_stock"], "url"])
    needed = PRICE_CHANGES + LEFT_OUT + OUT_OF_STOCK
    if len(candidates) < needed:
        raise ValueError(f"Need at least {needed} in-stock products with a price, found {len(candidates)}.")
    picks = rng.sample(candidates, needed)
    moved = picks[:PRICE_CHANGES]
    left_out = picks[PRICE_CHANGES:PRICE_CHANGES + LEFT_OUT]
    sold_out = picks[PRICE_CHANGES + LEFT_OUT:]

    signs = [-1, 1] * (PRICE_CHANGES // 2) + [-1] * (PRICE_CHANGES % 2)
    rng.shuffle(signs)
    for url, sign in zip(moved, signs):
        change = sign * rng.uniform(0.05, 0.25)
        row = history["url"] == url
        # Work back from today's price so the report shows this % change from old to new.
        history.loc[row, "price"] = round(history.loc[row, "price"].iat[0] / (1 + change), 2)

    sold = history["url"].isin(sold_out)
    history.loc[sold, "stock"] = "Out of stock"
    history.loc[sold, "in_stock"] = False
    history = history[~history["url"].isin(left_out)].copy()
    shifted = pd.to_datetime(history["scraped_at"]) - pd.Timedelta(days=days)
    history["scraped_at"] = shifted.dt.strftime("%Y-%m-%d %H:%M:%S")
    return history.reset_index(drop=True)


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--snapshots", type=Path, default=HERE / "snapshots",
                        help="snapshot folder (default: snapshots/ next to this script)")
    parser.add_argument("--days", type=int, default=7, help="how many days before the newest snapshot (default 7)")
    parser.add_argument("--seed", type=int, default=7, help="random seed, so the same edits come out each time")
    args = parser.parse_args(argv)

    saved = list_snapshots(args.snapshots)
    if not saved:
        raise SystemExit(f"No snapshots in {args.snapshots}. Run price_tracker.py first.")
    newest = saved[-1]
    history = make_history(load_snapshot(newest), random.Random(args.seed), args.days)
    path = save_snapshot(history, snapshot_path(args.snapshots, snapshot_time(newest) - timedelta(days=args.days)))
    print(f"Wrote {path.name}: generated demo history, {args.days} days before {newest.name}.")
    print(f"{PRICE_CHANGES} prices moved, {LEFT_OUT} products left out (they show as new), "
          f"{OUT_OF_STOCK} marked out of stock. Next: python price_tracker.py --compare-only")


if __name__ == "__main__":
    main()
