# Customer insights (RFM segmentation)

**Problem:** most online shops treat every customer the same: one newsletter, one discount. The best customers get nothing special, and good customers who quietly stop ordering aren't noticed until they've gone.

**What this does:** reads a shop's order history, cleans it, and scores every customer on how recently they ordered, how often, and how much they spent after returns. Each customer lands in one of nine segments, from Champions to Lost, with a suggested action, all in one Excel workbook.

This demo runs on **real public data**, not made-up data: 541,909 order lines from a UK online gift shop, 1 Dec 2010 to 9 Dec 2011.

## What it found

- **4,319 customers** analysed, with £8,214,279 of net revenue (sales minus returns).
- **691 Champions** (16% of customers) bring **62.7% of net revenue**.
- **631 customers at risk** (At risk and Can't lose them): good customers who haven't ordered for months. They spent £856,583 in the period.
- 65.6% of customers ordered more than once.
- The dataset's page says it has no missing values, but a quarter of the lines (135,080) have no customer ID.

## What you get (one Excel file)

| Sheet | What's on it |
|---|---|
| Summary | Headline numbers, revenue by segment and a chart |
| Segments | Customers, revenue, share of revenue and a suggested action for each segment, plus how segments are decided |
| Customers | One row per customer: days since last order, orders, sales, returns, net spend, scores and segment, biggest spenders first |
| Monthly | Net revenue, orders, active customers and new customers by month |
| Cleaning log | Every cleaning rule with the lines it removed, and the customers left out |

## How to run it

The dataset is 23 MB, so it isn't in this repo. Download [online+retail.zip](https://archive.ics.uci.edu/static/public/352/online+retail.zip) from the [dataset's page](https://archive.ics.uci.edu/dataset/352/online+retail) and unzip `Online Retail.xlsx` into a `data` folder here. Then:

```bash
pip install -r ../requirements.txt
python analyse.py
```

Reading the Excel file takes about a minute and a half; the analysis takes seconds. To read a file somewhere else: `python analyse.py "path/to/Online Retail.xlsx"`.

The finished workbook is in [output/customer_insights.xlsx](output/customer_insights.xlsx). Tests: `python -m pytest tests`. They need no download: [sample_data/online_retail_sample.csv](sample_data/online_retail_sample.csv) holds every line for 297 real customers (plus 13 lines with no customer), cut from the full file by `make_sample_data.py`.

## How it decides

**Cleaning**, in this order (each line is counted once, under the first rule it meets):

1. No customer ID: 135,080 lines can't be tied to a customer, so they're left out.
2. Exact duplicates: 5,225 lines identical to an earlier line are removed.
3. Not a product: 1,378 lines of postage, carriage, bank charges, charity commission and other lines that aren't goods.
4. Free items: 39 lines with a price of £0.
5. Returns: 8,758 lines on invoices starting with C (cancellations and returns, £589,612) are kept apart.

That leaves 391,429 sale lines.

**Returns** are taken off the customer's spend, but they aren't orders, so they don't count towards how often or how recently someone ordered. 53 customers are left out of the segments because they have no net spend: their returns equal or exceed what they bought (most only have returns in this period), or nothing was left after cleaning. They're listed in the Cleaning log.

**Manual lines and discounts** (stock codes M and D) stay in spend. They aren't stock items, but they're money the customer paid or got back. In this file one manual refund of £38,970 cancels an order of 60 picnic baskets at £649.50 each; leaving it out would overstate that customer's spend by £38,970.

**Scores**, each 1 to 5, cut at every fifth of the customers (5 is the best):

- Recency: days from the last order to 10 Dec 2011, the day after the last invoice. Here 5 means 1 to 14 days, 1 means 180 days or more.
- Frequency: the number of different invoices. A third of customers ordered once, so the bands are 1, 2, 3, 4 to 6, and 7 or more orders.
- Monetary: net spend in pounds. 5 means over £1,999, 1 means up to £243.

Equal values always get the same score, so a tied group can be bigger than a fifth. The exact bands are on the Segments sheet.

**Segments:** a value score (the average of the frequency and spend scores, rounded up) and the recency score place each customer in one cell of this grid:

| Recency | Value 1 | Value 2 | Value 3 | Value 4 | Value 5 |
|---|---|---|---|---|---|
| 5 | New | Potential loyalists | Potential loyalists | Loyal | Champions |
| 4 | New | Potential loyalists | Potential loyalists | Loyal | Champions |
| 3 | Need attention | Need attention | Need attention | Loyal | Loyal |
| 2 | Hibernating | Hibernating | At risk | At risk | Can't lose them |
| 1 | Lost | Lost | At risk | At risk | Can't lose them |

For example, Champions get early access and a thank-you rather than discounts; Can't lose them get a personal call or email from the owner; Lost get one last win-back email and then no more spend.

**One limit:** the file doesn't say which order a return belongs to, so an order that was later cancelled in full still counts as an order. That's about 1 order in 200 here.

## The data

Online Retail dataset, UCI Machine Learning Repository (Daqing Chen, 2015), CC BY 4.0. [doi.org/10.24432/C5BW33](https://doi.org/10.24432/C5BW33)

Customer IDs in the dataset are anonymised numbers; there are no names or contact details. Amounts are in pounds. The dataset's page notes that many of the shop's customers are wholesalers, which is part of why a small group brings most of the revenue.

## For your business

I'd connect it to your shop's orders (Shopify, WooCommerce, Amazon, eBay or a spreadsheet), tune the segments and actions to your customers, and send each segment to your email tool as a ready-made list. Run every month, it keeps those lists current, so you can reach customers who are slipping away while there's still time to win them back.
