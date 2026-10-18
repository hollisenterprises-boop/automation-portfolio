# Portfolio demos

Small, working tools of the kind I build for online sellers and small businesses. Each one comes with its finished output, so you can open the result without running anything. Most run on made-up sample data; the customer insights demo runs on a real public dataset of 541,909 order lines.

Live: the [demo site](https://hollisenterprises-boop.github.io/automation-portfolio/) and the [interactive sales dashboard](https://hollisenterprises-boop.github.io/automation-portfolio/sales-dashboard/output/dashboard.html).

| Demo | What it does |
|---|---|
| [Sales report builder](sales-report/) | Cleans a messy store order export and builds a formatted Excel sales report, with a log of every fix |
| [Store sales dashboard](sales-dashboard/) | Turns the cleaned order data into an interactive web dashboard (one HTML file) with KPI cards, charts, a state filter and the latest orders |
| [Amazon ads waste finder](ads-waste-finder/) | Finds the searches wasting ad money, the bids to lower and the winners to promote, from an Amazon search term report |
| [Competitor price tracker](price-tracker/) | Scrapes a product catalogue on a schedule and reports price, stock and range changes in Excel |
| [Contact list cleaner](list-cleaner/) | Merges messy customer lists from different sources, fixes names, emails and phone numbers, removes duplicates, and keeps marketing consent strict |
| [PDF invoices to Excel](invoice-extractor/) | Reads a folder of supplier invoice PDFs in different layouts, checks the maths and GST, flags duplicates, and writes one Excel workbook |
| [Customer insights (RFM)](customer-insights/) | Finds a shop's best customers and the ones slipping away: scores every customer on how recently and often they order and what they spend after returns, and puts each in a segment with a suggested action (real public data) |

To run any of them: install Python 3.12, then `pip install -r requirements.txt`, then follow the demo's README.

Portfolio images: `python make_images.py` writes them to [images/](images/).
