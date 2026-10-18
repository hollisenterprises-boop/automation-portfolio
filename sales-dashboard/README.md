# Store sales dashboard

**Problem:** a spreadsheet report answers last month's questions. A dashboard lets you, or your team, check sales any day, filter by state and see what's selling without opening Excel.

**What this does:** turns the cleaned order data from the [sales report builder](../sales-report/) into one interactive web page: KPI cards, net sales by month, top products, sales by state and the latest orders, with a state filter and hover details. It's a single HTML file, so you can open it in any browser, email it or host it.

## How to run it

```bash
pip install -r ../requirements.txt
python build_dashboard.py
```

Then open [output/dashboard.html](output/dashboard.html). Add `?state=WA` to the address to open it filtered. Tests: `python -m pytest tests` (every order line is in the page, and net sales match the Excel report to the cent). `python make_demo_gif.py` records the short animated demo in ../images/ (needs Chrome).

## How it decides

- **Net sales** count shipped and delivered orders only, the same rule as the Excel report. Refunds, cancellations and pending orders get their own cards.
- **Charts** use one colour each and a single axis, with exact values on hover.

All data is made up.

## For your business

I'd connect it to your store (Shopify, WooCommerce, Amazon, eBay) or your spreadsheets, refresh it on a schedule, and add the numbers you care about, such as profit, ad spend or stock.
