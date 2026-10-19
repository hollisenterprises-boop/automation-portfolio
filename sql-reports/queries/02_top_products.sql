-- title: Top products
-- about: The 10 best-selling products by sales, with each one's share of all sales and the running total.
-- Reads the sold_lines view (shipped, delivered and completed orders), see schema.sql.
SELECT RANK() OVER (ORDER BY SUM(l.line_cents) DESC)                       AS "Rank",
       l.sku                                                               AS "SKU",
       p.product_name                                                      AS "Product",
       SUM(l.qty)                                                          AS "Units",
       COUNT(DISTINCT l.order_id)                                          AS "Orders",
       ROUND(SUM(l.line_cents) / 100.0, 2)                                 AS "Sales",
       ROUND(1.0 * SUM(l.line_cents) / SUM(SUM(l.line_cents)) OVER (), 4)  AS "Share of sales",
       ROUND(1.0 * SUM(SUM(l.line_cents)) OVER (
                 ORDER BY SUM(l.line_cents) DESC, l.sku
                 ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW)
                 / SUM(SUM(l.line_cents)) OVER (), 4)                      AS "Running share"
FROM sold_lines AS l
JOIN products   AS p ON p.sku = l.sku
GROUP BY l.sku, p.product_name
ORDER BY SUM(l.line_cents) DESC, l.sku
LIMIT 10;
