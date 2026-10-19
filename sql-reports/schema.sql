-- The store database: four tables and one view (SQLite).
--
-- Money is stored as whole cents, so sums are exact. The queries turn it back
-- into dollars at the end. Dates are stored as text in YYYY-MM-DD form, which
-- sorts correctly and works with SQLite's date functions.

PRAGMA foreign_keys = ON;

CREATE TABLE customers (
    customer_id  TEXT PRIMARY KEY,
    name         TEXT NOT NULL,
    state        TEXT NOT NULL
);

CREATE TABLE products (
    sku           TEXT PRIMARY KEY,
    product_name  TEXT NOT NULL
);

CREATE TABLE orders (
    order_id     TEXT PRIMARY KEY,
    order_date   TEXT NOT NULL
                 CHECK (order_date GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
    status       TEXT NOT NULL,
    ship_state   TEXT NOT NULL,
    customer_id  TEXT NOT NULL REFERENCES customers (customer_id)
);

CREATE TABLE order_items (
    order_id          TEXT    NOT NULL REFERENCES orders (order_id),
    sku               TEXT    NOT NULL REFERENCES products (sku),
    qty               INTEGER NOT NULL CHECK (qty > 0),
    unit_price_cents  INTEGER NOT NULL CHECK (unit_price_cents >= 0)
);

CREATE INDEX idx_orders_customer    ON orders (customer_id);
CREATE INDEX idx_orders_date        ON orders (order_date);
CREATE INDEX idx_order_items_order  ON order_items (order_id);

-- What counts as a sale: lines on shipped, delivered and completed orders.
-- Cancelled, refunded and pending orders are left out. Every report reads
-- from this view, so the rule lives in one place (it is the same rule the
-- sales report demo uses).
CREATE VIEW sold_lines AS
SELECT o.order_id,
       o.order_date,
       o.ship_state,
       o.customer_id,
       i.sku,
       i.qty,
       i.unit_price_cents,
       i.qty * i.unit_price_cents AS line_cents
FROM orders      AS o
JOIN order_items AS i ON i.order_id = o.order_id
WHERE o.status IN ('Shipped', 'Delivered', 'Completed');
