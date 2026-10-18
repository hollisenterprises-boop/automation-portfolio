# Contact list cleaner

**Problem:** your customers are spread across several lists: the online store's export, a sign-up form, an old CRM. Each one writes names, emails and phone numbers differently, and the same people appear in more than one. Merging them by hand takes hours, and it's easy to email someone who unsubscribed.

**What this does:** reads every CSV and Excel list in a folder, cleans each value, finds the same person across lists, and writes one clean list in Excel, with every merge and every problem shown so a person can check them.

## What you get (one Excel file)

| Sheet | What's on it |
|---|---|
| Summary | People in the clean list, duplicates merged and why, values fixed, problems to check, and who you can market to |
| Clean List | One row per person: name, email, phone, suburb, state, postcode, marketing consent, which lists they were in |
| Merged Rows | Every original row that was merged, as it was written, so merges can be checked (name-only merges are shaded) |
| To Check | Emails and phone numbers it couldn't fix, and people with no way to contact them |

## How to run it

```bash
pip install -r ../requirements.txt
python make_sample_data.py
python clean_list.py sample_data -o output/clean_contacts.xlsx
```

The finished workbook from the sample lists is in [output/clean_contacts.xlsx](output/clean_contacts.xlsx). Tests: `python -m pytest tests`.

## How it decides

- **Columns:** worked out from the headers ("E-mail address", "PHONE_NO", "Keep me updated" and so on), so each list can be laid out differently.
- **Names:** proper case (O'Brien, McDonald), titles removed, "Smith, Olivia" turned round.
- **Emails:** lower case, spaces and "mailto:" removed; anything still not a valid address goes to To Check.
- **Phones:** every Australian format (+61, brackets, dashes, a lost leading zero) written one way: 08 5550 1234, 0491 570 156.
- **Same person:** the same email or phone. As a last resort, a similar first name (Kate and Katherine) with the same surname and suburb, but only when the rows have no conflicting email or phone. Those merges are shaded for checking.
- **Marketing consent:** one "no" or "unsubscribed" in any list means no.

On the sample (197 rows from 3 lists, built from 138 made-up people), it finds all 138 people, merges 57 duplicate rows, and lists 4 values to check.

All data is made up: emails use example.com and similar, and phone numbers are in the ranges Australia reserves for fiction.

## For your business

I'd match it to your lists (Shopify, Mailchimp, a CRM, spreadsheets), add your own rules (which list wins when two disagree), and export straight back into your email or CRM tool.
