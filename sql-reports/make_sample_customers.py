"""Write the made-up customer files that go with the sample store orders.

    python make_sample_customers.py

The order export in ../sales-report/sample_data has no customer column, so this
makes two small files to go with it:

- sample_data/customers.csv: Customer ID, Name, State
- sample_data/order_customers.csv: Order ID, Customer ID (who placed each order)

Every customer lives in the state their orders ship to. About 3 in 10 orders
come from a customer who has ordered before, so there are repeat customers for
the SQL to find. All names are made up. The random seed is fixed, so every run
gives the same files.
"""

from __future__ import annotations

import csv
import random
import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent / "sales-report"))
import sales_report  # noqa: E402  (its cleaning step reads the messy export)

ORDERS = HERE.parent / "sales-report" / "sample_data" / "orders_export_messy.csv"
OUT = HERE / "sample_data"
REPEAT_CHANCE = 0.30

FIRST = ["Olivia", "Jack", "Charlotte", "Noah", "Amelia", "William", "Isla", "Oliver", "Mia", "Thomas",
         "Grace", "James", "Chloe", "Lucas", "Ava", "Henry", "Sophie", "Liam", "Ruby", "Samuel",
         "Ella", "Benjamin", "Harper", "Daniel", "Matilda", "Joshua", "Priya", "Arjun", "Mei", "Tariq",
         "Aroha", "Wei", "Fatima", "Kai", "Leila", "Marcus", "Nadia", "Declan", "Imogen", "Rohan"]
LAST = ["Smith", "Nguyen", "Brown", "Wilson", "Taylor", "Johnson", "White", "Martin", "Anderson", "Thompson",
        "Walker", "Harris", "Lee", "Ryan", "Robinson", "Kelly", "King", "Davies", "Wright", "Mitchell",
        "Campbell", "Scott", "Murphy", "Hughes", "Edwards", "Clarke", "Price", "Turner", "Nair", "Patel",
        "Singh", "Chen", "Tran", "Hassan", "Williams", "Cooper", "Bennett", "Morgan", "Reid", "Stewart"]


def make_customers(orders: list[tuple[str, str]], seed: int = 2027) -> tuple[list[tuple], list[tuple]]:
    """Give every order a customer. orders is (order ID, state) in date order.

    Returns (customers as (ID, name, state), orders as (order ID, customer ID)).
    """
    rng = random.Random(seed)
    customers: list[tuple[str, str, str]] = []
    by_state: dict[str, list[str]] = {}  # state -> customer IDs
    order_count: dict[str, int] = {}
    names: set[str] = set()  # only checked for membership, never looped over
    assigned = []
    for order_id, state in orders:
        pool = by_state.setdefault(state, [])
        if pool and rng.random() < REPEAT_CHANCE:
            # People who have ordered more often are more likely to order again.
            customer_id = rng.choices(pool, weights=[1 + order_count[c] for c in pool])[0]
        else:
            name = f"{rng.choice(FIRST)} {rng.choice(LAST)}"
            while name in names:
                name = f"{rng.choice(FIRST)} {rng.choice(LAST)}"
            names.add(name)
            customer_id = f"C{len(customers) + 1:04d}"
            customers.append((customer_id, name, state))
            pool.append(customer_id)
            order_count[customer_id] = 0
        order_count[customer_id] += 1
        assigned.append((order_id, customer_id))
    return customers, assigned


def write_files(out_dir: Path = OUT, orders_csv: Path = ORDERS) -> tuple[Path, Path]:
    clean, _ = sales_report.clean_orders(sales_report.load_export(orders_csv))
    orders = (clean.drop_duplicates("Order ID").sort_values(["Order Date", "Order ID"], kind="stable")
              [["Order ID", "Ship State"]].itertuples(index=False, name=None))
    customers, assigned = make_customers(list(orders))
    out_dir.mkdir(parents=True, exist_ok=True)
    customers_path, mapping_path = out_dir / "customers.csv", out_dir / "order_customers.csv"
    for path, header, rows in ((customers_path, ["Customer ID", "Name", "State"], customers),
                               (mapping_path, ["Order ID", "Customer ID"], assigned)):
        with path.open("w", newline="", encoding="utf-8") as handle:
            writer = csv.writer(handle, lineterminator="\n")
            writer.writerow(header)
            writer.writerows(rows)
    return customers_path, mapping_path


if __name__ == "__main__":
    for written in write_files():
        print(f"Wrote {written}")
