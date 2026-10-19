-- title: Summary
-- about: The headline numbers on the Summary sheet: orders, units, net sales, and how many customers came back.
-- Sales are shipped, delivered and completed orders (the sold_lines view in schema.sql).
WITH per_customer AS (
    SELECT customer_id,
           COUNT(DISTINCT order_id) AS orders,
           SUM(line_cents)          AS cents
    FROM sold_lines
    GROUP BY customer_id
)
SELECT (SELECT COUNT(DISTINCT order_id) FROM sold_lines)                       AS "Orders",
       (SELECT SUM(qty) FROM sold_lines)                                       AS "Units sold",
       ROUND((SELECT SUM(line_cents) FROM sold_lines) / 100.0, 2)              AS "Net sales",
       ROUND((SELECT SUM(line_cents) FROM sold_lines) / 100.0
             / (SELECT COUNT(DISTINCT order_id) FROM sold_lines), 2)           AS "Average order",
       (SELECT COUNT(*) FROM per_customer)                                     AS "Customers who bought",
       (SELECT COUNT(*) FROM per_customer WHERE orders >= 2)                   AS "Repeat customers",
       ROUND(1.0 * (SELECT COUNT(*) FROM per_customer WHERE orders >= 2)
                 / (SELECT COUNT(*) FROM per_customer), 4)                     AS "Repeat customer rate",
       ROUND(1.0 * (SELECT SUM(cents) FROM per_customer WHERE orders >= 2)
                 / (SELECT SUM(cents) FROM per_customer), 4)                   AS "Share of sales from repeat customers",
       (SELECT MIN(order_date) FROM sold_lines)                                AS "First order",
       (SELECT MAX(order_date) FROM sold_lines)                                AS "Latest order";
