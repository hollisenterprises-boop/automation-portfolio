# SQL reports in Excel

**Problem:** the answers a shop owner needs sit in a database or an order export, and every new question means someone writing a query, pasting the result into Excel and formatting it by hand. Next month the numbers don't match, because nobody remembers exactly which orders were counted.

**What this does:** loads the store's orders into a SQLite database (four tables, plus one view that defines what counts as a sale), runs one .sql file for each question, and writes the results to a formatted Excel workbook with a summary page, charts, and the SQL behind every sheet. To add a report, add a .sql file.

## What you get (one Excel file and the database)

| Sheet | What's on it |
|---|---|
| Summary | Orders, units, net sales, average order, customers who bought, repeat customers and their share of sales, plus charts of sales by month and by state |
| Sales by month | Orders, units, sales and average order for each month, with the change on the month before |
| Top products | The best sellers by sales, each one's share, and a running total |
| Repeat customers | Every customer with two or more orders: first and latest order, days between, total spent, average order |
| Totals by state | Orders, customers, units, sales and share of sales for each ship-to state |
| Queries | The SQL behind every sheet, so anyone on your team can re-run it |

## How to run it

```bash
pip install -r ../requirements.txt
python make_sample_customers.py
python sql_reports.py -o output --as-of 2026-10-19
```

The finished workbook is [output/sql_report.xlsx](output/sql_report.xlsx) and the database is [output/store.db](output/store.db) (open it in any SQLite tool). Tests: `python -m pytest tests`.

## The queries

The SQL is in [queries/](queries/), one file for each report, and the tables are in [schema.sql](schema.sql). Every report reads one view, `sold_lines`, which holds the rule for what counts as a sale (shipped, delivered and completed orders), so the rule lives in one place.

| File | What it shows | SQL it uses |
|---|---|---|
| 00_headline.sql | The Summary sheet's numbers | subqueries, a CTE |
| 01_sales_by_month.sql | Sales by month with the change on the month before | CTE, `LAG` window function |
| 02_top_products.sql | The 10 best sellers, share and running share | join, `RANK`, window sums over a group |
| 03_repeat_customers.sql | Customers with two or more orders | CTE, join, `HAVING`, date maths |
| 04_totals_by_state.sql | Orders, customers, units and sales for each state | `COUNT(DISTINCT ...)`, share of the total |

Money is stored as whole cents, so sums are exact. The database also refuses bad rows: an order for a customer that doesn't exist, a quantity of zero, a date that isn't YYYY-MM-DD.

## How it's checked

- The numbers agree with the Python sales report from the first demo, to the cent: net sales $25,546.25 from 519 orders, and every monthly, product and state total.
- Each query is also tested on a tiny database of eight orders whose answers are worked out by hand, including a cancelled, a refunded and a pending order that must not count.
- The tests compare every query with an independent calculation in pandas.

## About the data

The orders are the made-up export from the [sales report builder](../sales-report/), cleaned by that demo's code (572 orders, 797 order lines, April to September 2026). That export has no customer column, so `make_sample_customers.py` adds 408 made-up customers, each living in the state their orders ship to, so there are repeat customers to find. In this sample, 102 of the 376 customers who bought ordered again. All data is made up.

## For your business

I'd point it at your real data (a database, a shop export or a set of spreadsheets), write the queries for the questions you actually ask (margins, stock, customers, ad spend, anything), and set it to produce the report on a schedule or on demand.
