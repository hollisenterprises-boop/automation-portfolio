-- title: Repeat customers
-- about: Customers with two or more orders: when they first and last ordered, and how much they spent.
-- Reads the sold_lines view (shipped, delivered and completed orders), see schema.sql.
WITH customer_orders AS (
    SELECT customer_id,
           order_id,
           order_date,
           SUM(line_cents) AS order_cents
    FROM sold_lines
    GROUP BY customer_id, order_id, order_date
)
SELECT c.customer_id                                          AS "Customer ID",
       c.name                                                 AS "Customer",
       c.state                                                AS "State",
       COUNT(*)                                               AS "Orders",
       MIN(co.order_date)                                     AS "First order",
       MAX(co.order_date)                                     AS "Latest order",
       CAST(julianday(MAX(co.order_date))
            - julianday(MIN(co.order_date)) AS INTEGER)       AS "Days between first and latest",
       ROUND(SUM(co.order_cents) / 100.0, 2)                  AS "Total spent",
       ROUND(SUM(co.order_cents) / 100.0 / COUNT(*), 2)       AS "Average order"
FROM customer_orders AS co
JOIN customers       AS c ON c.customer_id = co.customer_id
GROUP BY c.customer_id, c.name, c.state
HAVING COUNT(*) >= 2
ORDER BY "Orders" DESC, "Total spent" DESC, "Customer ID";
