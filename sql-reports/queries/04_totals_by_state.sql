-- title: Totals by state
-- about: Orders, customers, units and sales for each ship-to state, biggest first.
-- Reads the sold_lines view (shipped, delivered and completed orders), see schema.sql.
SELECT ship_state                                                     AS "State",
       COUNT(DISTINCT order_id)                                       AS "Orders",
       COUNT(DISTINCT customer_id)                                    AS "Customers",
       SUM(qty)                                                       AS "Units",
       ROUND(SUM(line_cents) / 100.0, 2)                              AS "Sales",
       ROUND(SUM(line_cents) / 100.0 / COUNT(DISTINCT order_id), 2)   AS "Average order",
       ROUND(1.0 * SUM(line_cents) / SUM(SUM(line_cents)) OVER (), 4) AS "Share of sales"
FROM sold_lines
GROUP BY ship_state
ORDER BY SUM(line_cents) DESC, ship_state;
