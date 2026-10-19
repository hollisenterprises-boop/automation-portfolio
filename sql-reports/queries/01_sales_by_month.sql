-- title: Sales by month
-- about: Orders, units, sales and average order for each month, with the change on the month before.
-- Reads the sold_lines view (shipped, delivered and completed orders), see schema.sql.
WITH monthly AS (
    SELECT strftime('%Y-%m', order_date) AS month,
           COUNT(DISTINCT order_id)      AS orders,
           SUM(qty)                      AS units,
           SUM(line_cents)               AS cents
    FROM sold_lines
    GROUP BY month
)
SELECT month                                                        AS "Month",
       orders                                                       AS "Orders",
       units                                                        AS "Units",
       ROUND(cents / 100.0, 2)                                      AS "Sales",
       ROUND(cents / 100.0 / orders, 2)                             AS "Average order",
       ROUND(1.0 * (cents - LAG(cents) OVER (ORDER BY month))
                 / LAG(cents) OVER (ORDER BY month), 4)             AS "Change on previous month"
FROM monthly
ORDER BY month;
