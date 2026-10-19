"""Draw Fiverr gig images (1280x769, Fiverr's recommended size): two per gig.

    python make_gig_images.py

images/gig-<name>-cover.png is the thumbnail: a headline of 10 words or fewer
(Fiverr's limit for gig images) over a real screenshot of the demo's result.
images/gig-<name>.png is the second image: the demo's numbers, in the same look
as the portfolio covers (make_covers.py). The numbers come from the demos'
committed output, so they stay true.
"""

from pathlib import Path

from PIL import Image, ImageDraw

from make_covers import BG_BOTTOM, BG_TOP, FAINT, MUTED, OUT, WHITE, font, hex_rgb, wrap
from make_images import load_summary

HERE = Path(__file__).parent
IMAGES = HERE / "images"
W, H = 1280, 769


def background():
    img = Image.new("RGB", (W, H))
    draw = ImageDraw.Draw(img)
    for y in range(H):
        t = y / (H - 1)
        draw.line([(0, y), (W, y)], fill=tuple(round(a + (b - a) * t) for a, b in zip(BG_TOP, BG_BOTTOM)))
    for x in range(0, W, 32):
        for y in range(0, H, 32):
            draw.point((x, y), fill=FAINT)
    return img


def cover(name, headline, screenshot, accent_hex):
    """Thumbnail: a short headline over the top of a real demo screenshot."""
    assert len(headline.split()) <= 10, "Fiverr allows 10 words at most on gig images"
    img = background()
    draw = ImageDraw.Draw(img)
    draw.text((48, 72), headline, font=font("seguisb.ttf", 54), fill=WHITE, anchor="lm")
    draw.rounded_rectangle([48, 116, 168, 124], radius=4, fill=hex_rgb(accent_hex))
    shot = Image.open(IMAGES / f"{screenshot}.png").convert("RGB")
    box_w, box_h = W - 96, H - 150 - 24
    shot = shot.crop((0, 0, shot.width, round(shot.width * box_h / box_w))).resize((box_w, box_h), Image.LANCZOS)
    img.paste(shot, (48, 150))
    path = OUT / f"gig-{name}-cover.png"
    img.save(path, optimize=True)
    print(f"Wrote {path}")


def gig(name, label, headline, stats, accent_hex, footer):
    accent = hex_rgb(accent_hex)
    img = background()
    draw = ImageDraw.Draw(img)
    left = 64

    draw.text((left, 56), label, font=font("seguisb.ttf", 32), fill=accent)
    head = font("seguisb.ttf", 68)
    y = 104
    for line in wrap(draw, headline, head, W - 2 * left):
        draw.text((left, y), line, font=head, fill=WHITE)
        y += 84

    gap, top = 28, max(y + 40, 330)
    tile_w = (W - 2 * left - 2 * gap) // 3
    for i, (number, text) in enumerate(stats):
        x0 = left + i * (tile_w + gap)
        size = 96
        while font("segoeuib.ttf", size).getlength(number) > tile_w - 60:
            size -= 6
        draw.rounded_rectangle([x0, top, x0 + tile_w, top + 210], radius=24, fill=(26, 44, 78))
        draw.text((x0 + 30, top + 26), number, font=font("segoeuib.ttf", size), fill=accent)
        draw.text((x0 + 32, top + 150), text, font=font("segoeui.ttf", 30), fill=MUTED)

    draw.text((left, top + 270), "Fixed price  ·  Built and tested on your data  ·  Plain instructions",
              font=font("seguisb.ttf", 34), fill=WHITE)
    draw.text((left, H - 52), footer, font=font("segoeui.ttf", 30), fill=MUTED, anchor="lm")
    OUT.mkdir(exist_ok=True)
    path = OUT / f"gig-{name}.png"
    img.save(path, optimize=True)
    print(f"Wrote {path}")


def main():
    sales = load_summary(HERE / "sales-report" / "output" / "sales_report.xlsx")
    prices = load_summary(HERE / "price-tracker" / "output" / "price_report.xlsx")
    invoices = load_summary(HERE / "invoice-extractor" / "output" / "invoices.xlsx")
    customers = load_summary(HERE / "customer-insights" / "output" / "customer_insights.xlsx")
    pdf_count = len(list((HERE / "invoice-extractor" / "sample_data" / "invoices").glob("*.pdf")))

    cover("excel-reports", "Excel and Google Sheets reports, automated", "sales-report", "#22C55E")
    cover("pdf-to-excel", "PDF invoices to Excel, checked and totalled", "invoice-extractor", "#F59E0B")
    cover("web-scraping", "Website data scraped into Excel or Sheets", "price-tracker", "#38BDF8")
    cover("store-insights", "Your best customers, and who's slipping away", "customer-insights", "#F472B6")

    gig("excel-reports", "EXCEL & GOOGLE SHEETS", "Your Excel or Google Sheets reports, built automatically", [
        (f"{int(sales['Values fixed in the export']):,}", "messy values fixed"),
        (f"{int(sales['Rows removed (blank, footer, duplicate, unreadable)'])}", "junk rows removed"),
        (f"${sales['Net sales (shipped and delivered)']:,.0f}", "sales reported"),
    ], "#22C55E", "From one of my demos  ·  Python, pandas, Excel, Apps Script  ·  tested")
    gig("pdf-to-excel", "PDF TO EXCEL", "PDF invoices and forms → one checked spreadsheet", [
        (f"{pdf_count}", "PDFs read, 3 layouts"),
        (f"{int(invoices['Invoices to check'])}/3", "planted errors caught"),
        (f"${invoices['Total spend, inc GST']:,.0f}", "spend totalled"),
    ], "#F59E0B", "From one of my demos  ·  Python, pdfplumber, Excel  ·  tested")
    gig("web-scraping", "WEB SCRAPING", "Website data scraped into Excel, CSV or Google Sheets", [
        (f"{int(prices['Products tracked'])}", "products tracked"),
        (f"{int(prices['Price drops'] + prices['Price rises'])}", "price changes found"),
        (f"{int(prices['New products'])}", "new products spotted"),
    ], "#38BDF8", "From one of my demos  ·  Python, requests, Excel  ·  public data only")
    at_risk = int(customers["Customers at risk (At risk + Can't lose them)"])
    gig("store-insights", "SALES & CUSTOMER INSIGHTS", "Find your best customers and the ones slipping away", [
        (f"{int(customers['Order lines in the file']):,}", "real order lines"),
        (f"{customers['Share of revenue from Champions']:.0%}", "from top customers"),
        (f"{at_risk:,}", "customers at risk"),
    ], "#F472B6", "From one of my demos  ·  Python, pandas, Excel  ·  real public data")


if __name__ == "__main__":
    main()
