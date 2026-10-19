"""Draw a bold cover image (1600x1200) for each demo, made to read at thumbnail size.

    python make_covers.py

Covers go to images/cover-<demo>.png and use the numbers from each demo's
committed Excel output, so they stay true. The detailed images from
make_images.py stay as each portfolio item's second picture.
"""

import sqlite3
from pathlib import Path

import pandas as pd
from PIL import Image, ImageDraw, ImageFont

from make_images import load_summary

HERE = Path(__file__).parent
OUT = HERE / "images"
FONTS = Path("C:/Windows/Fonts")
W, H = 1600, 1200

BG_TOP, BG_BOTTOM = (11, 18, 32), (22, 40, 72)
WHITE, MUTED, FAINT = (241, 245, 249), (165, 180, 203), (52, 76, 116)
BAD = (190, 70, 70)


def font(name, size):
    return ImageFont.truetype(str(FONTS / name), size)


def hex_rgb(value):
    value = value.lstrip("#")
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))


def wrap(draw, text, fnt, width):
    lines, line = [], ""
    for word in text.split():
        trial = f"{line} {word}".strip()
        if draw.textlength(trial, font=fnt) <= width:
            line = trial
        else:
            lines.append(line)
            line = word
    return lines + [line]


def background(img):
    draw = ImageDraw.Draw(img)
    for y in range(H):
        t = y / (H - 1)
        draw.line([(0, y), (W, y)], fill=tuple(round(a + (b - a) * t) for a, b in zip(BG_TOP, BG_BOTTOM)))
    for x in range(0, W, 40):
        for y in range(0, H, 40):
            draw.point((x, y), fill=FAINT)


def grid(draw, x, y, accent, messy):
    """A small table: ragged and broken (before) or neat with a header (after)."""
    cols, rows, cw, rh, gap = 5, 4, 74, 34, 8
    for r in range(rows):
        for c in range(cols):
            x0, y0 = x + c * (cw + gap), y + r * (rh + gap)
            if messy:
                jitter = ((r * 7 + c * 13) % 5 - 2) * 4
                width = cw - ((r + c * 3) % 4) * 12
                if (r, c) in {(1, 3), (2, 1), (3, 4), (0, 2)}:
                    draw.rounded_rectangle([x0 + jitter, y0, x0 + jitter + width, y0 + rh], radius=6,
                                           fill=BAD)
                elif (r, c) in {(2, 3), (1, 0)}:
                    draw.rounded_rectangle([x0, y0, x0 + cw, y0 + rh], radius=6, outline=FAINT, width=3)
                else:
                    draw.rounded_rectangle([x0 + jitter, y0, x0 + jitter + width, y0 + rh], radius=6,
                                           fill=(70, 92, 130))
            else:
                fill = accent if r == 0 else (58, 82, 122)
                draw.rounded_rectangle([x0, y0, x0 + cw, y0 + rh], radius=6, fill=fill)
    return cols * (cw + gap) - gap, rows * (rh + gap) - gap


def bars(draw, x, y, accent, width, height):
    """A small bar chart in the same box as a table (the 'after' of a dashboard)."""
    heights, gap = [0.5, 0.68, 0.62, 1.0, 0.72, 0.66], 16
    bw = (width - gap * (len(heights) - 1)) / len(heights)
    for i, h in enumerate(heights):
        x0 = x + i * (bw + gap)
        draw.rounded_rectangle([x0, y + height * (1 - h), x0 + bw, y + height], radius=6, fill=accent)


def pages(draw, x, y, accent, width, height):
    """A small certificate and a letter side by side (the 'after' of a document generator)."""
    cw, ch = 200, height - 18  # the landscape certificate
    cx0, cy0 = x, y + (height - ch) // 2
    draw.rounded_rectangle([cx0, cy0, cx0 + cw, cy0 + ch], radius=8, fill=WHITE)
    draw.rounded_rectangle([cx0 + 8, cy0 + 8, cx0 + cw - 8, cy0 + ch - 8], radius=4, outline=accent, width=3)
    mid = cx0 + cw // 2
    draw.rounded_rectangle([mid - 52, cy0 + 22, mid + 52, cy0 + 32], radius=4, fill=(24, 79, 149))
    draw.rounded_rectangle([mid - 66, cy0 + 56, mid + 66, cy0 + 68], radius=5, fill=(70, 92, 130))
    draw.rounded_rectangle([mid - 40, cy0 + 82, mid + 40, cy0 + 90], radius=4, fill=(150, 165, 190))
    draw.ellipse([mid - 15, cy0 + ch - 50, mid + 15, cy0 + ch - 20], fill=accent)
    draw.line([(cx0 + 24, cy0 + ch - 24), (cx0 + 64, cy0 + ch - 24)], fill=(120, 135, 160), width=2)
    draw.line([(cx0 + cw - 64, cy0 + ch - 24), (cx0 + cw - 24, cy0 + ch - 24)], fill=(120, 135, 160), width=2)

    lx0, lw = x + cw + 24, 108  # the portrait letter
    draw.rounded_rectangle([lx0, y, lx0 + lw, y + height], radius=8, fill=WHITE)
    draw.rounded_rectangle([lx0 + 12, y + 14, lx0 + 66, y + 22], radius=4, fill=(24, 79, 149))
    draw.line([(lx0 + 12, y + 30), (lx0 + lw - 12, y + 30)], fill=accent, width=2)
    for i, share in enumerate((0.55, 0.75, 0.9, 1.0, 0.6)):
        top = y + 42 + i * 12
        draw.rounded_rectangle([lx0 + 12, top, lx0 + 12 + (lw - 24) * share, top + 6], radius=3, fill=(150, 165, 190))
    for i in range(3):
        top = y + 108 + i * 14
        draw.rectangle([lx0 + 12, top, lx0 + lw - 12, top + 11], outline=(150, 165, 190), width=1)
        draw.rectangle([lx0 + 12, top, lx0 + 42, top + 11], fill=(214, 224, 240))
    return cw + 24 + lw, height


def cover(name, label, headline, stats, accent_hex, tools, after="table", data="made-up sample data"):
    accent = hex_rgb(accent_hex)
    img = Image.new("RGB", (W, H))
    background(img)
    draw = ImageDraw.Draw(img)
    left = 90

    draw.text((left, 80), label, font=font("seguisb.ttf", 40), fill=accent)
    head_font = font("seguisb.ttf", 104)
    y = 140
    for line in wrap(draw, headline, head_font, W - 2 * left):
        draw.text((left, y), line, font=head_font, fill=WHITE)
        y += 128

    tile_w, top = (W - 2 * left - 2 * 40) // 3, max(y + 50, 470)
    lab_font = font("segoeui.ttf", 38)
    for i, (number, text) in enumerate(stats):
        x0 = left + i * (tile_w + 40)
        size = 112
        while font("segoeuib.ttf", size).getlength(number) > tile_w - 72:
            size -= 6
        draw.rounded_rectangle([x0, top, x0 + tile_w, top + 250], radius=28, fill=(26, 44, 78))
        draw.text((x0 + 36, top + 30), number, font=font("segoeuib.ttf", size), fill=accent)
        draw.text((x0 + 38, top + 176), text, font=lab_font, fill=MUTED)

    strip_y = top + 320
    gw, gh = grid(draw, left, strip_y, accent, messy=True)
    ax = left + gw + 70
    mid = strip_y + gh // 2
    draw.line([(ax, mid), (ax + 150, mid)], fill=accent, width=12)
    draw.polygon([(ax + 150, mid - 30), (ax + 200, mid), (ax + 150, mid + 30)], fill=accent)
    after_x = ax + 270
    if after == "bars":
        bars(draw, after_x, strip_y, accent, gw, gh)
    elif after == "pages":
        pages(draw, after_x, strip_y, accent, gw, gh)
    else:
        grid(draw, after_x, strip_y, accent, messy=False)
    caption = font("segoeui.ttf", 36)
    draw.text((left + gw // 2, strip_y + gh + 40), "Before", font=caption, fill=MUTED, anchor="mm")
    draw.text((after_x + gw // 2, strip_y + gh + 40), "After", font=caption, fill=MUTED, anchor="mm")

    draw.text((left, H - 70), f"{tools}  \u00b7  tested  \u00b7  {data}", font=font("segoeui.ttf", 34),
              fill=MUTED, anchor="lm")
    OUT.mkdir(exist_ok=True)
    path = OUT / f"cover-{name}.png"
    img.save(path, optimize=True)
    print(f"Wrote {path}")


def main():
    sales = load_summary(HERE / "sales-report" / "output" / "sales_report.xlsx")
    ads = load_summary(HERE / "ads-waste-finder" / "output" / "ads_report.xlsx")
    prices = load_summary(HERE / "price-tracker" / "output" / "price_report.xlsx")
    invoices = load_summary(HERE / "invoice-extractor" / "output" / "invoices.xlsx")
    lists = load_summary(HERE / "list-cleaner" / "output" / "clean_contacts.xlsx")
    customers = load_summary(HERE / "customer-insights" / "output" / "customer_insights.xlsx")
    docs = load_summary(HERE / "doc-generator" / "output" / "generation_log.xlsx")
    sql_report = HERE / "sql-reports" / "output" / "sql_report.xlsx"
    sql = load_summary(sql_report)
    with sqlite3.connect(HERE / "sql-reports" / "output" / "store.db") as db:
        orders_loaded = db.execute("SELECT COUNT(*) FROM orders").fetchone()[0]
    sql_queries = len(pd.read_excel(sql_report, sheet_name="Queries"))
    pdf_count = len(list((HERE / "invoice-extractor" / "sample_data" / "invoices").glob("*.pdf")))

    cover("sales-report", "EXCEL REPORTING", "Messy store export \u2192 clean Excel report", [
        (f"${sales['Net sales (shipped and delivered)']:,.0f}", "net sales reported"),
        (f"{int(sales['Values fixed in the export']):,}", "values fixed"),
        (f"{int(sales['Rows removed (blank, footer, duplicate, unreadable)'])}", "junk rows removed"),
    ], "#22C55E", "Python \u00b7 pandas \u00b7 Excel")
    clean = pd.read_excel(HERE / "sales-report" / "output" / "sales_report.xlsx", sheet_name="Clean Data")
    cover("sales-dashboard", "DASHBOARDS", "Store orders \u2192 a live sales dashboard", [
        (f"${sales['Net sales (shipped and delivered)']:,.0f}", "net sales tracked"),
        (f"{clean['Ship State'].nunique()}", "states to filter"),
        ("1", "HTML file, no install"),
    ], "#818CF8", "Python \u00b7 pandas \u00b7 Chart.js", after="bars")
    cover("ads-waste-finder", "AMAZON ADS", "Stop paying for ad clicks that never sell", [
        (f"${ads['Wasted spend per 30 days']:,.0f}", "wasted per month"),
        (f"{ads['Share of ad spend wasted']:.0%}", "of ad spend"),
        (f"{int(ads['Negatives to add'])}", "searches to block"),
    ], "#FB923C", "Python \u00b7 pandas \u00b7 Excel")
    cover("price-tracker", "WEB SCRAPING", "Competitor prices and stock, tracked daily", [
        (f"{int(prices['Products tracked'])}", "products tracked"),
        (f"{int(prices['Price drops'] + prices['Price rises'])}", "price changes found"),
        (f"{int(prices['New products'])}", "new products spotted"),
    ], "#38BDF8", "Python \u00b7 requests \u00b7 Excel")
    cover("invoice-extractor", "PDF DATA EXTRACTION", "Supplier PDF invoices \u2192 one checked spreadsheet", [
        (f"{pdf_count}", "PDFs read"),
        (f"{int(invoices['Invoices to check'])}/3", "planted errors caught"),
        (f"${invoices['Total spend, inc GST']:,.0f}", "spend totalled"),
    ], "#F59E0B", "Python \u00b7 pdfplumber \u00b7 Excel")
    cover("list-cleaner", "DATA CLEANING", "Three messy contact lists \u2192 one clean list", [
        (f"{int(lists['People in the clean list'])}", "real people found"),
        (f"{int(lists['Duplicate rows merged'])}", "duplicates merged"),
        (f"{int(lists['Said no or unsubscribed'])}", "opt-outs respected"),
    ], "#14B8A6", "Python \u00b7 Excel")
    at_risk = int(customers["Customers at risk (At risk + Can't lose them)"])
    cover("customer-insights", "CUSTOMER ANALYTICS", "Find your best customers and the ones slipping away", [
        (f"{int(customers['Customers analysed']):,}", "customers analysed"),
        (f"{customers['Share of revenue from Champions']:.0%}", "of revenue: Champions"),
        (f"{at_risk:,}", "customers at risk"),
    ], "#F472B6", "Python \u00b7 pandas \u00b7 Excel", after="bars", data="real public data")
    files_made = int(docs["Certificates made (PDF)"] + docs["Letters made (Word)"])
    cover("doc-generator", "DOCUMENT AUTOMATION", "Spreadsheet \u2192 certificates and letters", [
        (f"{int(docs['Rows read'])}", "spreadsheet rows"),
        (f"{files_made}", "PDF and Word files"),
        (f"{int(docs['Rows skipped (data missing or unreadable)'])}", "rows flagged"),
    ], "#C084FC", "Python \u00b7 reportlab \u00b7 python-docx", after="pages")
    cover("sql-reports", "SQL REPORTING", "Orders in SQL \u2192 a formatted Excel report", [
        (f"{orders_loaded}", "orders loaded"),
        (f"{sql_queries}", "SQL queries"),
        (f"${sql['Net sales']:,.0f}", "net sales, checked"),
    ], "#A3E635", "Python \u00b7 SQLite \u00b7 Excel", after="bars")


if __name__ == "__main__":
    main()
