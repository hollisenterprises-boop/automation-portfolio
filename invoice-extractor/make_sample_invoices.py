"""Make the sample supplier invoices (PDFs) for the invoice extractor demo.

    python make_sample_invoices.py

Writes 13 made-up invoices from three suppliers, each with its own layout, to
sample_data/invoices/. A few have deliberate problems for the extractor to
catch: a total that doesn't add up, GST charged at the wrong rate, and the
same invoice saved twice under different file names.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

OUT_DIR = Path(__file__).parent / "sample_data" / "invoices"
BILL_TO = "Example Store Pty Ltd, 7 Sample St, Perth WA 6000"


@dataclass
class Invoice:
    layout: str  # "table", "lines" or "services"
    number: str
    issued: date
    lines: list[tuple[str, float, float]]  # description, quantity, unit price
    terms_days: int = 30
    gst_rate: float = 0.10  # what the supplier actually charged
    total_error: float = 0.0  # added to the printed total
    file_name: str = ""
    extra: dict = field(default_factory=dict)

    @property
    def subtotal(self) -> float:
        return round(sum(round(q * p, 2) for _, q, p in self.lines), 2)

    @property
    def gst(self) -> float:
        return round(self.subtotal * self.gst_rate, 2)

    @property
    def total(self) -> float:
        return round(self.subtotal + self.gst + self.total_error, 2)

    @property
    def due(self) -> date:
        return self.issued + timedelta(days=self.terms_days)


def money(value: float) -> str:
    return f"${value:,.2f}"


def unit_price(value: float) -> str:
    """Unit prices keep a third decimal when they have one ($0.035 a label)."""
    return f"${value:,.3f}" if round(value, 2) != round(value, 3) else money(value)


def styles():
    base = getSampleStyleSheet()
    return {
        "name": ParagraphStyle("name", parent=base["Title"], alignment=0, fontSize=20, spaceAfter=2),
        "small": ParagraphStyle("small", parent=base["Normal"], fontSize=9, textColor=colors.HexColor("#444444")),
        "head": ParagraphStyle("head", parent=base["Heading2"], fontSize=14, spaceBefore=6),
        "body": base["Normal"],
    }


def grid_table(rows, widths, header_colour):
    table = Table(rows, colWidths=widths)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(header_colour)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#BBBBBB")),
        ("ALIGN", (1, 1), (-1, -1), "RIGHT"),
        ("FONTSIZE", (0, 0), (-1, -1), 9.5),
    ]))
    return table


def coastal_packaging(inv: Invoice, s) -> list:
    """Layout "table": label: value fields and a ruled item table."""
    rows = [["Description", "Qty", "Unit Price", "Amount"]]
    rows += [[d, f"{q:g}", unit_price(p), money(q * p)] for d, q, p in inv.lines]
    return [
        Paragraph("Coastal Packaging Supplies", s["name"]),
        Paragraph("12 Harbour Rd, Fremantle WA 6160 | accounts@coastal-packaging.example", s["small"]),
        Paragraph("TAX INVOICE", s["head"]),
        Paragraph(f"Invoice No: {inv.number}", s["body"]),
        Paragraph(f"Invoice Date: {inv.issued:%d/%m/%Y}", s["body"]),
        Paragraph(f"Due Date: {inv.due:%d/%m/%Y}", s["body"]),
        Paragraph(f"Bill To: {BILL_TO}", s["body"]),
        Spacer(1, 6 * mm),
        grid_table(rows, [90 * mm, 18 * mm, 30 * mm, 30 * mm], "#1F4E79"),
        Spacer(1, 5 * mm),
        Paragraph(f"Subtotal: {money(inv.subtotal)}", s["body"]),
        Paragraph(f"GST: {money(inv.gst)}", s["body"]),
        Paragraph(f"Total: {money(inv.total)}", s["body"]),
        Spacer(1, 5 * mm),
        Paragraph("Please pay by bank transfer within 30 days. Thank you for your business.", s["small"]),
    ]


def northside_print(inv: Invoice, s) -> list:
    """Layout "lines": items written as '250 x Thing @ $0.18 = $45.00', ISO dates."""
    parts = [
        Paragraph("Northside Print Co.", s["name"]),
        Paragraph("Unit 4, 88 Scarborough Beach Rd, Osborne Park WA 6017", s["small"]),
        Paragraph(f"Tax Invoice # {inv.number}", s["head"]),
        Paragraph(f"Issued: {inv.issued:%Y-%m-%d}", s["body"]),
        Paragraph(f"Payment due: {inv.due:%Y-%m-%d}", s["body"]),
        Paragraph(f"Customer: {BILL_TO}", s["body"]),
        Spacer(1, 6 * mm),
        Paragraph("Items", s["head"]),
    ]
    for d, q, p in inv.lines:
        parts.append(Paragraph(f"{q:g} x {d} @ {unit_price(p)} = {money(q * p)}", s["body"]))
    parts += [
        Spacer(1, 5 * mm),
        Paragraph(f"Subtotal (ex GST): {money(inv.subtotal)}", s["body"]),
        Paragraph(f"GST: {money(inv.gst)}", s["body"]),
        Paragraph(f"Total due: {money(inv.total)}", s["body"]),
    ]
    return parts


def bright_spark(inv: Invoice, s) -> list:
    """Layout "services": hours x rate table, written-out dates, due from payment terms."""
    rows = [["Service", "Hours", "Rate", "Line total"]]
    rows += [[d, f"{q:g}", unit_price(p), money(q * p)] for d, q, p in inv.lines]
    return [
        Paragraph("Bright Spark Electrical", s["name"]),
        Paragraph("Licensed electrical contractor | 3/21 Example Way, Joondalup WA 6027", s["small"]),
        Spacer(1, 4 * mm),
        Paragraph(f"Invoice Number: {inv.number}", s["body"]),
        Paragraph(f"Invoice Date: {inv.issued:%d %b %Y}", s["body"]),
        Paragraph(f"Terms: {inv.terms_days} days", s["body"]),
        Paragraph(f"Client: {BILL_TO}", s["body"]),
        Spacer(1, 6 * mm),
        grid_table(rows, [90 * mm, 18 * mm, 30 * mm, 30 * mm], "#8A3B12"),
        Spacer(1, 5 * mm),
        Paragraph(f"Total ex GST: {money(inv.subtotal)}", s["body"]),
        Paragraph(f"GST 10%: {money(inv.gst)}", s["body"]),
        Paragraph(f"Amount Payable: {money(inv.total)}", s["body"]),
    ]


LAYOUTS = {"table": coastal_packaging, "lines": northside_print, "services": bright_spark}


def sample_invoices() -> list[Invoice]:
    d = date
    invoices = [
        Invoice("table", "INV-10421", d(2026, 8, 4), [("Shipping boxes 40x30x20cm", 200, 1.15), ("Packing tape 48mm, 6 pack", 12, 14.90), ("Void fill paper roll", 4, 32.50)]),
        Invoice("table", "INV-10477", d(2026, 8, 18), [("Mailer bags 25x35cm, 100 pack", 10, 18.40), ("Fragile labels, roll of 500", 3, 21.00)]),
        Invoice("table", "INV-10512", d(2026, 9, 1), [("Shipping boxes 40x30x20cm", 300, 1.15), ("Bubble wrap 500mm x 50m", 6, 39.95)], total_error=10.00),
        Invoice("table", "INV-10563", d(2026, 9, 15), [("Mailer bags 25x35cm, 100 pack", 15, 18.40), ("Packing tape 48mm, 6 pack", 10, 14.90)]),
        Invoice("table", "INV-10609", d(2026, 9, 29), [("Shipping boxes 30x20x15cm", 400, 0.92), ("Void fill paper roll", 6, 32.50)]),
        Invoice("lines", "NP-2291", d(2026, 8, 10), [("Business cards, double sided", 500, 0.09), ("Thank-you insert cards A6", 1000, 0.07)], terms_days=14),
        Invoice("lines", "NP-2318", d(2026, 8, 31), [("Product labels 50mm round", 2000, 0.035), ("Flyers A5 gloss", 500, 0.22)], terms_days=14),
        Invoice("lines", "NP-2344", d(2026, 9, 21), [("Thank-you insert cards A6", 1500, 0.07)], terms_days=14, file_name="northside_NP-2344.pdf"),
        Invoice("lines", "NP-2344", d(2026, 9, 21), [("Thank-you insert cards A6", 1500, 0.07)], terms_days=14, file_name="scan_0921_reprint.pdf"),
        Invoice("lines", "NP-2370", d(2026, 10, 9), [("Product labels 50mm round", 3000, 0.035)], terms_days=14),
        Invoice("services", "BSE-0087", d(2026, 8, 22), [("Warehouse lighting fault find", 2.5, 95.00), ("LED batten replacement", 4, 95.00)], terms_days=14),
        Invoice("services", "BSE-0102", d(2026, 9, 12), [("Switchboard upgrade labour", 6.5, 95.00), ("Test and tag, 40 items", 1, 180.00)], terms_days=14, gst_rate=0.125),
        Invoice("services", "BSE-0119", d(2026, 10, 3), [("After-hours callout", 1, 165.00), ("Power point install", 2, 95.00)], terms_days=7),
    ]
    for inv in invoices:
        if not inv.file_name:
            prefix = {"table": "coastal", "lines": "northside", "services": "brightspark"}[inv.layout]
            inv.file_name = f"{prefix}_{inv.number}.pdf"
    return invoices


def write_invoice(inv: Invoice, folder: Path) -> Path:
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / inv.file_name
    doc = SimpleDocTemplate(str(path), pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm, topMargin=16 * mm)
    doc.build(LAYOUTS[inv.layout](inv, styles()))
    return path


def main() -> None:
    for old in OUT_DIR.glob("*.pdf"):
        old.unlink()
    for inv in sample_invoices():
        write_invoice(inv, OUT_DIR)
    print(f"Wrote {len(list(OUT_DIR.glob('*.pdf')))} invoices to {OUT_DIR}")


if __name__ == "__main__":
    main()
