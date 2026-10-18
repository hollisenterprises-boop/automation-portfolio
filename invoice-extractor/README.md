# PDF invoices to Excel

**Problem:** supplier invoices arrive as PDFs, and every supplier lays them out differently. Someone retypes them into a spreadsheet, which takes hours, and mistakes slip through: a total that doesn't add up, GST charged at the wrong rate, or the same invoice paid twice because it was emailed twice.

**What this does:** reads a whole folder of invoice PDFs, pulls out the supplier, invoice number, dates, line items, subtotal, GST and total, checks the maths, and writes one Excel workbook. Anything it can't read, or that doesn't add up, is listed for a person to check rather than guessed.

## What you get (one Excel file)

| Sheet | What's on it |
|---|---|
| Summary | Invoices, total spend, GST to claim, invoices to check, what's due in the next 14 days, and spend by supplier |
| Invoices | One row per PDF: supplier, number, dates, due status, subtotal, GST, total, and OK or Check |
| Line Items | Every line from every invoice, ready for pivot tables |
| Checks | Each problem found, in plain words |

## How to run it

```bash
pip install -r ../requirements.txt
python make_sample_invoices.py
python extract.py sample_data/invoices -o output/invoices.xlsx --as-of 2026-10-18
```

The finished workbook from the sample invoices is in [output/invoices.xlsx](output/invoices.xlsx). Tests: `python -m pytest tests`.

## What it checks

- **Each line:** quantity × unit price = the line amount.
- **Items:** the lines add up to the subtotal.
- **Totals:** subtotal + GST = total.
- **GST:** 10% of the subtotal (Australian GST; set to suit your country).
- **Duplicates:** the same supplier and invoice number in two files is counted once and flagged.
- **Due dates:** read from the invoice, or worked out from the payment terms ("Terms: 14 days").

The sample has three planted problems, and it catches all three: a total $10.00 too high, GST charged at 12.5%, and one invoice saved twice under different file names.

All invoices are made up.

## For your business

I'd set it up for your suppliers' layouts, add OCR for scanned invoices, and send the results where you need them: Excel, Google Sheets, an import file for your accounting software, or an email when something needs checking.
