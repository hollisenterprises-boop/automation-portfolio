"""Write a deliberately messy order export, like the ones online stores produce.

    python make_sample_data.py

All data is made up. The random seed is fixed, so every run gives the same file.
"""

import csv
import random
from datetime import date, timedelta
from pathlib import Path

PRODUCTS = [  # sku, name, price, popularity
    ("GLS-JAR-500", "Glass Storage Jar 500ml", 14.95, 6),
    ("GLS-JAR-1000", "Glass Storage Jar 1L", 19.95, 5),
    ("MP-SET-5", "Glass Meal Prep Set (5 pack)", 39.95, 8),
    ("MP-SET-10", "Glass Meal Prep Set (10 pack)", 69.95, 3),
    ("LID-BAMBOO-4", "Bamboo Lids (4 pack)", 12.95, 4),
    ("SPICE-12", "Spice Jar Set (12)", 29.95, 3),
    ("OIL-SPRAY", "Glass Oil Sprayer", 17.95, 4),
    ("LUNCH-2C", "Two-Compartment Lunch Box", 24.95, 5),
]
STATES = [  # code, full name, share of orders
    ("NSW", "New South Wales", 30), ("VIC", "Victoria", 25), ("QLD", "Queensland", 20),
    ("WA", "Western Australia", 10), ("SA", "South Australia", 7), ("TAS", "Tasmania", 3),
    ("ACT", "Australian Capital Territory", 3), ("NT", "Northern Territory", 2),
]
HEADER = [" Order ID", "Order Date ", "SKU", "Product Name", "Qty", "Item Price",
          "Order Status", "Ship State"]
FIRST_DAY, LAST_DAY = date(2026, 4, 1), date(2026, 9, 30)
OUT = Path(__file__).parent / "sample_data" / "orders_export_messy.csv"

rng = random.Random(42)


def pick(options):
    """options: list of (value, weight)."""
    values, weights = zip(*options)
    return rng.choices(values, weights=weights)[0]


def messy_date(d):
    return pick([(d.isoformat(), 50), (d.strftime("%d/%m/%Y"), 25),
                 (f"{d.day} {d:%b %Y}", 15),
                 (f"{d.isoformat()} {rng.randint(7, 22):02d}:{rng.randint(0, 59):02d}:{rng.randint(0, 59):02d}", 10)])


def messy_price(p):
    return pick([(f"{p:.2f}", 55), (f"${p:.2f}", 30), (f"AUD {p:.2f}", 10), (f" {p:.2f} ", 5)])


def messy_qty(q):
    return pick([(str(q), 85), (f"{q} unit" + ("s" if q > 1 else ""), 10), (f"x{q}", 5)])


def messy_status(status):
    variants = {
        "Shipped": ["Shipped", "shipped", "SHIPPED"],
        "Delivered": ["Delivered", "delivered"],
        "Cancelled": ["Cancelled", "CANCELED", "canceled"],
        "Refunded": ["Refunded", "refunded"],
        "Pending": ["Pending", "pending "],
    }[status]
    return pick([(variants[0], 70)] + [(v, 30 / (len(variants) - 1)) for v in variants[1:]])


def messy_state(code, name):
    options = [(code, 60), (name, 25), (code.lower(), 10)]
    if len(code) == 2:
        options.append((".".join(code) + ".", 5))
    return pick(options)


def messy_sku(sku):
    return pick([(sku, 85), (sku.lower(), 8), (f" {sku} ", 7)])


def messy_name(name):
    if rng.random() < 0.05:
        return name.replace(" ", "  ", 1)
    return name


def order_lines():
    """Yield (order_id, day, sku, name, qty, price, status, state code, state name)."""
    order_no = 100100
    day = FIRST_DAY
    while day <= LAST_DAY:
        progress = (day - FIRST_DAY).days / (LAST_DAY - FIRST_DAY).days
        for _ in range(max(0, round(rng.gauss(2.2 + 1.6 * progress, 1.2)))):
            order_no += 1
            if day >= LAST_DAY - timedelta(days=5) and rng.random() < 0.7:
                status = "Pending"
            else:
                status = pick([("Shipped", 46.75), ("Delivered", 46.75),
                               ("Cancelled", 3.5), ("Refunded", 3)])
            code, name, _ = pick([((c, n, w), w) for c, n, w in STATES])
            count = pick([(1, 70), (2, 22), (3, 8)])
            products = rng.sample(PRODUCTS, k=count)
            for sku, product, price, _ in products:
                qty = pick([(1, 80), (2, 15), (3, 5)])
                yield (f"SS-{order_no}", day, sku, product, qty, price, status, code, name)
        day += timedelta(days=1)


def render(line):
    order_id, day, sku, product, qty, price, status, code, name = line
    return [order_id, messy_date(day), messy_sku(sku), messy_name(product), messy_qty(qty),
            messy_price(price), messy_status(status), messy_state(code, name)]


def main():
    rows = []
    for line in order_lines():
        rows.append(render(line))
        if rng.random() < 0.02:  # the same line exported twice, formatted differently
            rows.append(render(line))
    for _ in range(6):
        rows.insert(rng.randrange(len(rows)), [])
    for index in rng.sample(range(len(rows)), k=3):
        if rows[index]:
            rows[index][5] = "TBC"
    rows.append([f"Total rows: {len(rows)}", "", "", "", "", "", "", ""])

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with OUT.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(HEADER)
        writer.writerows(rows)
    print(f"Wrote {len(rows):,} rows to {OUT}")


if __name__ == "__main__":
    main()
