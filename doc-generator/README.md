# Certificates and letters from a spreadsheet

**Problem:** every time a course finishes, someone copies each person's name, course and date from a spreadsheet into a certificate template, then into a letter, one at a time. It takes an afternoon, names get mistyped, and afterwards nobody can say which ones were done.

**What this does:** reads a spreadsheet with a name, course, date and amount on each row, and makes a PDF certificate and a Word letter for every complete row. A row with something missing is skipped and listed rather than guessed. Every file is opened again after it's written, to check the name, course and certificate number are really on it, and an Excel log lists everything it made and everything it skipped.

## What you get

| Output | What's on it |
|---|---|
| certificates/ | One PDF per person (A4 landscape): name, course, date, certificate number, signature line and seal |
| letters/ | One Word letter per person (A4): greeting, congratulations, a details table with the fee paid, and a sign-off. They open and edit like any Word file |
| generation_log.xlsx | **Summary** (rows read, files made, rows skipped, fees, made by course), **Files made** (every file, with its re-open check) and **Skipped rows** (the row number, what the sheet had and what's wrong) |

## How to run it

```bash
pip install -r ../requirements.txt
python make_sample_completions.py
python generate_docs.py sample_data/course_completions.xlsx -o output --as-of 2026-10-19
```

The sample spreadsheet has 40 rows. The run takes about 12 seconds and makes 37 certificates and 37 letters. [output/](output/) keeps the log from that run and three example certificates and letters; the other 68 files are made when you run it. Tests: `python -m pytest tests`.

Use a new output folder for each run. If files from an earlier run are still in the folder, the tool says so.

## What it checks

- **Missing data:** a blank name, course, date or amount skips the row, and the log says which.
- **Unreadable data:** a date like "TBC", an amount like "n/a" or a negative amount is listed, not guessed. Dates can be real Excel dates or typed text (13/03/2026, 2026-03-13, 13 March 2026); slashed dates are read day-first.
- **Names that won't print:** letters the PDF's built-in fonts can't draw (Vietnamese, for example) or a name too long for one line are flagged instead of printing as boxes. Long names shrink to fit; accents, apostrophes and hyphens are fine.
- **After writing:** each PDF and Word file is opened again, and the name, course and certificate number (and the fee, on the letter) must be found in it.
- **Blank rows** in the middle of the sheet are ignored and counted; blank rows at the bottom are ignored.
- **One bad row never stops the run.** It is skipped and listed with the reason.

The sample has three planted problems, and it catches all three: a row with no name, one with no amount, and one with no course or date (the log shows what each row did have). It also has a name with an accent, one with an apostrophe, a hyphenated one, a 50-character name, and dates typed four different ways.

All names and data are made up. Each file carries a small note saying so; `--note ""` removes it, and `--org` and `--signer` change the organisation and the name under the signature line.

## For your business

I'd put your logo, wording, fonts and signature on the certificate and the letter, match your spreadsheet's columns, and send the finished files where you need them: a shared folder, or an email to each person with their two files attached.
