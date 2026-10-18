# Amazon ads waste finder

Amazon charges for every click on a Sponsored Products ad, whether the shopper
buys or not. Some searches collect clicks for weeks and never sell, and they are
hard to spot in the Amazon Ads console. This tool reads your search term report
and gives you one Excel file that says which searches to block, which to promote
and which bids to lower.

On the sample month it finds $462 of $1,855 in ad spend (25%) going on searches
that never sold.

## What the report shows

- **Summary**: ad spend, sales, ACOS, the money spent on searches that never sold
  (and what that comes to per month), and a chart of spend by outcome.
- **Negatives to add**: searches with clicks but no sales, to add as negative keywords.
- **Bids to lower**: searches that sell but cost too much, each with a suggested bid.
- **Winners to add**: searches that sell at a good ACOS but aren't exact keywords yet.
- **All search terms**: every search with its numbers and decision, biggest spend first.

## How to run it

    pip install pandas openpyxl
    python make_sample_data.py
    python ads_waste_finder.py sample_data/search_term_report.csv -o output/ads_report.xlsx

For your own numbers, download a Sponsored Products **Search term** report with
time unit **Daily** from Amazon Ads (CSV or Excel) and use its file name in the
last command. Tests: `pip install pytest`, then `pytest`.

## How it decides

ACOS is ad spend divided by the sales the ads brought in. Each search is added up
over the whole report, separately for each campaign and ad group, then put into
one of four groups:

- **No sales**: no orders after 10 or more clicks, or $15 or more spend. Add it as
  a negative exact keyword. If the search is a product code (like b08kx7q2lm, a
  click from another product's page), add it as a negative product target.
- **Too expensive**: it sells, but ACOS is above 45% (1.5 times the 30% target).
  Lower the bid. Suggested bid = cost per click x 30% / ACOS, rounded to the
  nearest 5 cents, never below $0.20. Amazon sets bids on the keyword or target,
  so apply it there (the Targeting column shows which one matched the search).
- **Profitable**: ACOS is 30% or less. With 2 or more orders and no exact keyword
  for it yet, it's a winner: add it as an exact keyword. Otherwise keep it running.
- **Keep watching**: everything else, either too few clicks to judge or close to target.

Change the limits with `--target-acos 0.25`, `--min-clicks 15` and `--min-spend 20`.

All sample data is made up.

**For a real store** it would be set up with your own thresholds, write an Amazon
bulk-upload file so the changes go in with one upload, and run weekly on a schedule.
