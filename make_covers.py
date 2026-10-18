"""Draw a bold cover image (1600x1200) for each demo, made to read at thumbnail size.

    python make_covers.py

Covers go to images/cover-<demo>.png and use the numbers from each demo's
committed Excel output, so they stay true. The detailed images from
make_images.py stay as each portfolio item's second picture.
"""

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


def cover(name, label, headline, stats, accent_hex, tools, after="table"):
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
    else:
        grid(draw, after_x, strip_y, accent, messy=False)
    caption = font("segoeui.ttf", 36)
    draw.text((left + gw // 2, strip_y + gh + 40), "Before", font=caption, fill=MUTED, anchor="mm")
    draw.text((after_x + gw // 2, strip_y + gh + 40), "After", font=caption, fill=MUTED, anchor="mm")

    draw.text((left, H - 70), f"{tools}  \u00b7  tested  \u00b7  made-up sample data", font=font("segoeui.ttf", 34),
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


if __name__ == "__main__":
    main()
