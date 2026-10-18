# Sales report builder

**Problem:** your store's order export is messy. Dates come in three formats, prices have "$" or "AUD" stuck to them, statuses are written five ways, and there are blank lines, duplicate rows and a "Total rows" line at the bottom. Turning that into a weekly or monthly report by hand takes hours, and it's easy to get wrong.

**What this does:** reads the export exactly as it comes out of the store, cleans it, and writes a formatted Excel report in seconds.

## What you get (one Excel file)

| Sheet | What's on it |
|---|---|
| Summary | Net sales, orders, units, average order, refunds, cancellations and pending orders, the top 5 products, and a monthly sales chart |
| By Month | Orders, units, sales and average order for each month |
| By Product | Units and sales for each product, with its share of total sales |
| By State | Orders and sales for each state |
| Clean Data | Every order line, cleaned, ready for pivot tables |
| Fixes Log | Every value it fixed and every row it removed, with the reason, so nothing changes silently |

## How to run it

```bash
pip install -r ../requirements.txt
python make_sample_data.py
python sales_report.py sample_data/orders_export_messy.csv -o output/sales_report.xlsx
```

The finished report from the sample data is in [output/sales_report.xlsx](output/sales_report.xlsx).

## How it decides

- **Sales** count shipped and delivered orders only. Refunded, cancelled and pending orders are shown separately.
- **Dates** like 03/04/2026 are read the Australian way (3 April).
- **Duplicates:** a line with the same order, product, quantity and price as an earlier one is treated as an export glitch and removed, even if it was formatted differently.
- **Unreadable rows** (for example a price of "TBC") are left out and listed in the Fixes Log rather than guessed.

All sample data is made up.

## For your store

I'd match it to your platform's export (Shopify, WooCommerce, eBay, Amazon and others), add the totals you care about (GST, shipping, profit per product), and set it to run on a schedule or with one click.
