"""Make a PDF certificate and a Word letter for every row of a spreadsheet.

    python generate_docs.py sample_data/course_completions.xlsx -o output --as-of 2026-10-19

Reads a spreadsheet (.xlsx or .csv) with a name, course, date and amount on
each row. Every complete row gets a PDF certificate and a Word letter. A row
with something missing or unreadable is skipped and listed, never guessed.
After writing, each file is opened again to check that the name, course and
certificate number are on it. An Excel log lists every file made and every row
skipped, with the reason.
"""

from __future__ import annotations

import argparse
import csv
import math
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import pdfplumber
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import simpleSplit
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

# Column headings the spreadsheet may use (case and extra spaces don't matter).
COLUMNS = {
    "name": ("name", "full name", "student", "participant", "attendee"),
    "course": ("course", "course name", "program", "programme", "workshop"),
    "date": ("date", "completed", "date completed", "completion date", "course date"),
    "amount": ("amount", "amount paid", "fee", "fee paid", "course fee", "price", "paid"),
}

# Slashed dates are read day-first (Australian format).
DATE_FORMATS = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d %b %Y", "%d %B %Y", "%b %d %Y", "%B %d %Y")

NAVY, GOLD, INK, GREY = HexColor("#184F95"), HexColor("#B7791F"), HexColor("#1F2937"), HexColor("#6B7280")
NAVY_RGB, GREY_RGB = RGBColor(0x18, 0x4F, 0x95), RGBColor(0x6B, 0x72, 0x80)
PAGE = landscape(A4)
NAME_FONT = "Times-BoldItalic"
COURSE_FONT = "Helvetica-Bold"
NAME_MAX_WIDTH, COURSE_MAX_WIDTH = 620, 600

MONEY = '"$"#,##0.00'
HEADER_FILL = PatternFill("solid", fgColor="184F95")
HEADER_FONT = Font(bold=True, color="FFFFFF")


@dataclass
class Settings:
    """What changes from one business to the next."""
    org: str = "Sample Skills Academy"
    address: str = "1 Example Street, Perth WA 6000"
    signer: str = "Alex Morgan"
    role: str = "Course Director"
    note: str = "Sample document. The name and all details are made up."


@dataclass
class Row:
    number: int  # row number in the spreadsheet (the heading is row 1)
    name: str = ""
    course: str = ""
    completed: date | None = None
    amount: float | None = None
    shown: dict = field(default_factory=dict)  # what the sheet had in each column, as text
    problems: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.problems


@dataclass
class Made:
    row: Row
    number: str  # certificate number
    certificate: str  # file path relative to the output folder
    letter: str
    check: str  # "OK", or what the re-opened file got wrong


@dataclass
class Result:
    source: str
    rows: list[Row]
    made: list[Made]
    skipped: list[Row]
    blank_rows: int
    leftover: int  # files from an earlier run still sitting in the output folder


# ---------- reading the spreadsheet ----------

def is_blank(value) -> bool:
    return value is None or str(value).strip() == ""


def clean_text(value) -> str:
    return "" if value is None else " ".join(str(value).split())


def show(value) -> str:
    """A cell as plain text, for the skipped-rows list."""
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return f"{value:%d/%m/%Y}"
    if isinstance(value, float):
        return f"{value:.2f}"
    return clean_text(value)


def parse_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = clean_text(value).rstrip(".")
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def parse_amount(value) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
    else:
        text = re.sub(r"(?i)aud|a\$|\$|,|\s", "", str(value))
        try:
            number = float(text)
        except ValueError:
            return None
    return round(number, 2) if math.isfinite(number) else None


def unprintable(text: str) -> list[str]:
    """Characters the PDF's built-in fonts can't draw (they cover Western European text)."""
    bad = []
    for char in text:
        try:
            char.encode("cp1252")
        except UnicodeEncodeError:
            if char not in bad:
                bad.append(char)
    return bad


def name_size(name: str, biggest: int = 46, smallest: int = 20) -> int | None:
    """The largest font size at which the name fits on one line, or None if it never does."""
    for size in range(biggest, smallest - 1, -1):
        if stringWidth(name, NAME_FONT, size) <= NAME_MAX_WIDTH:
            return size
    return None


def course_lines(course: str) -> list[str]:
    return simpleSplit(course, COURSE_FONT, 26, COURSE_MAX_WIDTH)


def build_row(number: int, values: dict) -> Row:
    """Turn one spreadsheet row into a Row, noting everything wrong with it."""
    row = Row(number=number, shown={key: show(value) for key, value in values.items()})
    row.name, row.course = clean_text(values["name"]), clean_text(values["course"])
    missing = [label for label in ("name", "course") if not getattr(row, label)]
    problems = []

    if is_blank(values["date"]):
        missing.append("date")
    else:
        row.completed = parse_date(values["date"])
        if row.completed is None:
            problems.append(f'Could not read the date "{row.shown["date"]}"')

    if is_blank(values["amount"]):
        missing.append("amount")
    else:
        row.amount = parse_amount(values["amount"])
        if row.amount is None:
            problems.append(f'Could not read the amount "{row.shown["amount"]}"')
        elif row.amount < 0:
            problems.append(f"The amount is negative ({row.shown['amount']})")

    for label, text in (("name", row.name), ("course", row.course)):
        bad = unprintable(text)
        if bad:
            problems.append(f"The {label} has characters the certificate font can't print: {' '.join(bad)}")
    if row.name and not unprintable(row.name) and name_size(row.name) is None:
        problems.append("The name is too long to fit on the certificate")
    if row.course and not unprintable(row.course) and len(course_lines(row.course)) > 2:
        problems.append("The course name is too long to fit on the certificate")

    row.problems = ([f"Missing: {', '.join(missing)}"] if missing else []) + problems
    return row


def find_columns(header: list) -> dict[str, int]:
    names = [" ".join(str(cell).lower().split()) if cell is not None else "" for cell in header]
    found = {}
    for key, options in COLUMNS.items():
        index = next((names.index(option) for option in options if option in names), None)
        if index is None:
            raise ValueError(f"The spreadsheet has no {key} column (looked for: {', '.join(options)}).")
        found[key] = index
    return found


def read_table(path: Path) -> list[list]:
    if path.suffix.lower() == ".csv":
        with path.open(newline="", encoding="utf-8-sig") as handle:
            return [list(record) for record in csv.reader(handle)]
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        return [list(record) for record in workbook.active.iter_rows(values_only=True)]
    finally:
        workbook.close()


def read_rows(path: Path) -> tuple[list[Row], int]:
    """Return (every row that has data, number of blank rows between them)."""
    table = read_table(Path(path))
    if not table:
        raise ValueError("The spreadsheet is empty.")
    columns = find_columns(table[0])
    entries = []
    for number, cells in enumerate(table[1:], start=2):
        entries.append((number, {key: cells[i] if i < len(cells) else None for key, i in columns.items()}))
    while entries and all(is_blank(v) for v in entries[-1][1].values()):
        entries.pop()  # blank rows at the end of the sheet are just empty space
    rows, blanks = [], 0
    for number, values in entries:
        if all(is_blank(v) for v in values.values()):
            blanks += 1
        else:
            rows.append(build_row(number, values))
    return rows, blanks


# ---------- names and wording ----------

def long_date(value: date) -> str:
    return f"{value.day} {value:%B %Y}"


def money(value: float) -> str:
    return f"${value:,.2f}"


def certificate_number(row: Row) -> str:
    return f"C-{row.completed.year}-{row.number:03d}"


def slug(text: str, limit: int = 40) -> str:
    """A name made safe for a file name: letters and digits, joined by underscores."""
    plain = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()
    plain = plain.replace("'", "").replace("’", "")
    return "_".join(re.findall(r"[A-Za-z0-9]+", plain))[:limit].strip("_") or "unnamed"


def first_name(name: str) -> str:
    return name.split()[0]


def file_names(row: Row) -> tuple[str, str]:
    stem = f"{row.number:03d}_{slug(row.name)}"
    return f"certificate_{stem}.pdf", f"letter_{stem}.docx"


# ---------- the PDF certificate ----------

def draw_seal(pdf: canvas.Canvas, x: float, y: float, radius: float, year: int) -> None:
    points = 24
    path = pdf.beginPath()
    for i in range(points * 2):
        angle = math.pi * i / points
        r = radius if i % 2 == 0 else radius * 0.86
        px, py = x + r * math.cos(angle), y + r * math.sin(angle)
        path.moveTo(px, py) if i == 0 else path.lineTo(px, py)
    path.close()
    pdf.setFillColor(GOLD)
    pdf.drawPath(path, fill=1, stroke=0)
    pdf.setStrokeColor(white)
    pdf.setLineWidth(1.2)
    pdf.circle(x, y, radius * 0.68, stroke=1, fill=0)
    pdf.setFillColor(white)
    pdf.setFont("Helvetica-Bold", 6.5)
    pdf.drawCentredString(x, y + 2, "COMPLETED")
    pdf.setFont("Helvetica", 6.5)
    pdf.drawCentredString(x, y - 9, str(year))


def draw_certificate(path: Path, row: Row, number: str, settings: Settings, issued: date) -> None:
    width, height = PAGE
    mid = width / 2
    pdf = canvas.Canvas(str(path), pagesize=PAGE)
    pdf.setTitle(f"Certificate of Completion: {row.name}")
    pdf.setAuthor(settings.org)
    pdf.setSubject(row.course)
    pdf.setCreator("generate_docs.py")

    pdf.setStrokeColor(NAVY)
    pdf.setLineWidth(4)
    pdf.rect(24, 24, width - 48, height - 48)
    pdf.setStrokeColor(GOLD)
    pdf.setLineWidth(1)
    pdf.rect(34, 34, width - 68, height - 68)
    pdf.setFillColor(GOLD)
    for cx, cy in ((34, 34), (width - 34, 34), (34, height - 34), (width - 34, height - 34)):
        diamond = pdf.beginPath()
        diamond.moveTo(cx, cy + 7)
        diamond.lineTo(cx + 7, cy)
        diamond.lineTo(cx, cy - 7)
        diamond.lineTo(cx - 7, cy)
        diamond.close()
        pdf.drawPath(diamond, fill=1, stroke=0)

    pdf.setFillColor(NAVY)
    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawCentredString(mid, height - 84, settings.org.upper(), charSpace=2.5)
    pdf.setFont("Times-Bold", 42)
    pdf.drawCentredString(mid, height - 150, "Certificate of Completion")
    pdf.setStrokeColor(GOLD)
    pdf.setLineWidth(1.5)
    pdf.line(mid - 70, height - 166, mid + 70, height - 166)

    pdf.setFillColor(GREY)
    pdf.setFont("Times-Italic", 17)
    pdf.drawCentredString(mid, height - 205, "This is to certify that")

    size = name_size(row.name) or 20
    pdf.setFillColor(INK)
    pdf.setFont(NAME_FONT, size)
    pdf.drawCentredString(mid, height - 262, row.name)
    underline = min(stringWidth(row.name, NAME_FONT, size) + 60, NAME_MAX_WIDTH + 20)
    pdf.setStrokeColor(GOLD)
    pdf.setLineWidth(0.75)
    pdf.line(mid - underline / 2, height - 276, mid + underline / 2, height - 276)

    pdf.setFillColor(GREY)
    pdf.setFont("Times-Italic", 17)
    pdf.drawCentredString(mid, height - 312, "has successfully completed")

    pdf.setFillColor(NAVY)
    pdf.setFont(COURSE_FONT, 26)
    y = height - 356
    for line in course_lines(row.course):
        pdf.drawCentredString(mid, y, line)
        y -= 32
    pdf.setFillColor(INK)
    pdf.setFont("Times-Roman", 18)
    pdf.drawCentredString(mid, y - 8, f"on {long_date(row.completed)}")

    draw_seal(pdf, mid, 118, 42, row.completed.year)

    pdf.setStrokeColor(INK)
    pdf.setLineWidth(0.75)
    pdf.line(90, 112, 280, 112)
    pdf.line(width - 280, 112, width - 90, 112)
    pdf.setFillColor(INK)
    pdf.setFont("Times-Roman", 13)
    pdf.drawCentredString(185, 96, settings.signer)
    pdf.drawCentredString(width - 185, 96, f"Issued {long_date(issued)}")
    pdf.setFillColor(GREY)
    pdf.setFont("Helvetica", 9)
    pdf.drawCentredString(185, 83, settings.role)
    pdf.drawCentredString(width - 185, 83, f"Certificate No. {number}")
    if settings.note:
        pdf.setFont("Helvetica", 8)
        pdf.drawCentredString(mid, 52, settings.note)
    pdf.showPage()
    pdf.save()


# ---------- the Word letter ----------

def add_bottom_border(paragraph, colour: str) -> None:
    properties = paragraph._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for key, value in (("w:val", "single"), ("w:sz", "8"), ("w:space", "6"), ("w:color", colour)):
        bottom.set(qn(key), value)
    borders.append(bottom)
    properties.insert_element_before(
        borders, "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
        "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid",
        "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
        "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle",
        "w:rPr", "w:sectPr", "w:pPrChange")


def shade(cell, fill: str) -> None:
    shading = OxmlElement("w:shd")
    for key, value in (("w:val", "clear"), ("w:color", "auto"), ("w:fill", fill)):
        shading.set(qn(key), value)
    cell._tc.get_or_add_tcPr().append(shading)


def draw_letter(path: Path, row: Row, number: str, settings: Settings, issued: date) -> None:
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Cm(21.0), Cm(29.7)  # A4 (python-docx starts on US Letter)
    section.left_margin = section.right_margin = Cm(2.5)
    section.top_margin = section.bottom_margin = Cm(2.2)

    normal = doc.styles["Normal"]
    normal.font.name, normal.font.size = "Calibri", Pt(11)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    normal.paragraph_format.space_after = Pt(8)
    normal.paragraph_format.line_spacing = 1.15

    head = doc.add_paragraph()
    run = head.add_run(settings.org)
    run.bold, run.font.size, run.font.color.rgb = True, Pt(20), NAVY_RGB
    head.paragraph_format.space_after = Pt(0)
    address = doc.add_paragraph()
    run = address.add_run(settings.address)
    run.font.size, run.font.color.rgb = Pt(9.5), GREY_RGB
    add_bottom_border(address, "B7791F")
    address.paragraph_format.space_after = Pt(22)

    doc.add_paragraph(long_date(issued))
    doc.add_paragraph().add_run(row.name).bold = True
    doc.add_paragraph(f"Dear {first_name(row.name)},")
    doc.add_paragraph().add_run("Your certificate of completion").bold = True

    para = doc.add_paragraph()
    para.add_run("Congratulations on completing ")
    para.add_run(row.course).bold = True
    para.add_run(" on ")
    para.add_run(long_date(row.completed)).bold = True
    para.add_run(". Your certificate is enclosed with this letter.")
    doc.add_paragraph(f"We have received your course fee of {money(row.amount)}. "
                      "The details are below for your records.")

    details = [("Course", row.course), ("Completed", long_date(row.completed)),
               ("Certificate number", number), ("Fee paid", money(row.amount))]
    table = doc.add_table(rows=len(details), cols=2)
    table.style = "Table Grid"
    table.autofit = False
    for record, (label, value) in zip(table.rows, details):
        left, right = record.cells
        left.width, right.width = Cm(4.5), Cm(11.5)
        left.text, right.text = label, value
        shade(left, "EAF1FB")
        left.paragraphs[0].runs[0].bold = True
        for cell in (left, right):
            cell.paragraphs[0].paragraph_format.space_after = Pt(2)
            cell.paragraphs[0].paragraph_format.space_before = Pt(2)

    closing = doc.add_paragraph("If anything on this page is wrong, please reply and we will correct it.")
    closing.paragraph_format.space_before = Pt(14)
    doc.add_paragraph("Yours sincerely,")
    signer = doc.add_paragraph()
    signer.paragraph_format.space_before = Pt(26)
    signer.paragraph_format.space_after = Pt(0)
    signer.add_run(settings.signer).bold = True
    doc.add_paragraph(f"{settings.role}, {settings.org}")

    if settings.note:
        footer = section.footer.paragraphs[0]
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = footer.add_run(settings.note)
        run.font.size, run.font.color.rgb = Pt(8.5), GREY_RGB

    props = doc.core_properties
    props.title = f"Course completion letter: {row.name}"
    props.subject = row.course
    props.author = props.last_modified_by = settings.org
    props.comments = ""
    props.created = props.modified = datetime.combine(issued, datetime.min.time())
    doc.save(path)


# ---------- checking what was written ----------

def squash(text: str) -> str:
    return unicodedata.normalize("NFC", " ".join(text.split()))


def check_certificate(path: Path, row: Row, number: str) -> str:
    """Open the PDF again and confirm the details are on its one page; return "OK" or the problem."""
    try:
        with pdfplumber.open(path) as pdf:
            if len(pdf.pages) != 1:
                return f"The certificate has {len(pdf.pages)} pages, not 1"
            text = squash(pdf.pages[0].extract_text() or "")
    except Exception as error:  # a file that won't open is the problem to report
        return f"The certificate would not open ({error})"
    for label, wanted in (("name", row.name), ("course", row.course), ("number", number),
                          ("date", long_date(row.completed))):
        if squash(wanted) not in text:
            return f"The {label} is not on the certificate"
    return "OK"


def check_letter(path: Path, row: Row, number: str) -> str:
    """Open the Word file again and confirm the details are in it; return "OK" or the problem."""
    try:
        doc = Document(path)
        parts = [p.text for p in doc.paragraphs]
        parts += [cell.text for table in doc.tables for record in table.rows for cell in record.cells]
    except Exception as error:
        return f"The letter would not open ({error})"
    text = squash(" ".join(parts))
    for label, wanted in (("name", row.name), ("course", row.course), ("number", number),
                          ("fee", money(row.amount)), ("greeting", f"Dear {first_name(row.name)},")):
        if squash(wanted) not in text:
            return f"The {label} is not in the letter"
    return "OK"


# ---------- running it ----------

def generate(source: Path, out_dir: Path, settings: Settings | None = None, issued: date | None = None) -> Result:
    """Make every certificate and letter, check them, and write the log."""
    settings, issued = settings or Settings(), issued or date.today()
    source, out_dir = Path(source), Path(out_dir)
    rows, blank_rows = read_rows(source)
    certificates, letters = out_dir / "certificates", out_dir / "letters"
    certificates.mkdir(parents=True, exist_ok=True)
    letters.mkdir(parents=True, exist_ok=True)

    made, skipped = [], []
    for row in rows:
        if not row.ok:
            skipped.append(row)
            continue
        number = certificate_number(row)
        cert_name, letter_name = file_names(row)
        try:
            draw_certificate(certificates / cert_name, row, number, settings, issued)
            draw_letter(letters / letter_name, row, number, settings, issued)
        except Exception as error:  # one bad row must not stop the rest
            for leftover in (certificates / cert_name, letters / letter_name):
                leftover.unlink(missing_ok=True)
            row.problems.append(f"Could not make the files ({error})")
            skipped.append(row)
            continue
        problems = [check_certificate(certificates / cert_name, row, number),
                    check_letter(letters / letter_name, row, number)]
        made.append(Made(row, number, f"certificates/{cert_name}", f"letters/{letter_name}",
                         next((p for p in problems if p != "OK"), "OK")))

    wanted = {m.certificate for m in made} | {m.letter for m in made}
    older = [p for folder in (certificates, letters) for p in folder.iterdir()
             if f"{folder.name}/{p.name}" not in wanted]
    result = Result(source.name, rows, made, skipped, blank_rows, len(older))
    build_log(result, issued).save(out_dir / "generation_log.xlsx")
    return result


# ---------- the Excel log ----------

def write_table(ws, headers: list[str], records: list[list], start_row: int = 1, formats: dict | None = None) -> None:
    formats = formats or {}
    for col, name in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=col, value=name)
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for r, record in enumerate(records, start=start_row + 1):
        for col, value in enumerate(record, start=1):
            cell = ws.cell(row=r, column=col, value=value)
            if headers[col - 1] in formats:
                cell.number_format = formats[headers[col - 1]]
            elif isinstance(value, date):
                cell.number_format = "dd/mm/yyyy"
    for col, name in enumerate(headers, start=1):
        width = max([len(str(name))] + [len(str(record[col - 1])) for record in records[:500]])
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width or 0, min(width + 3, 60))


def build_log(result: Result, issued: date) -> Workbook:
    made, skipped = result.made, result.skipped
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    files = wb.create_sheet("Files made")
    write_table(files, ["Row", "Name", "Course", "Completed", "Fee", "Certificate", "Letter", "Checked"],
                [[m.row.number, m.row.name, m.row.course, m.row.completed, m.row.amount, m.certificate,
                  m.letter, m.check] for m in made], formats={"Fee": MONEY})
    skips = wb.create_sheet("Skipped rows")
    write_table(skips, ["Row", "Name", "Course", "Date", "Amount", "What is wrong"],
                [[r.number, r.shown["name"], r.shown["course"], r.shown["date"], r.shown["amount"],
                  "; ".join(r.problems)] for r in skipped])
    for sheet in (files, skips):
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions

    ws["A1"] = f"Document run: {result.source}"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"] = f"Made on {issued:%d %b %Y}. Every file was opened again after writing to check the details on it."
    ws["A2"].font = Font(color="52514E")

    checked_ok = sum(2 for m in made if m.check == "OK")
    kpis = [
        ("Rows read", len(result.rows), "#,##0"),
        ("Certificates made (PDF)", len(made), "#,##0"),
        ("Letters made (Word)", len(made), "#,##0"),
        ("Rows skipped (data missing or unreadable)", len(skipped), "#,##0"),
        ("Files that passed the re-open check", checked_ok, "#,##0"),
        ("Files that failed the re-open check", len(made) * 2 - checked_ok, "#,##0"),
        ("Blank rows ignored", result.blank_rows, "#,##0"),
        ("Fees on the documents made", round(sum(m.row.amount for m in made), 2), MONEY),
        ("Fees on skipped rows (nothing made)", round(sum(r.amount for r in skipped if r.amount), 2), MONEY),
    ]
    for offset, (label, value, fmt) in enumerate(kpis):
        ws.cell(row=4 + offset, column=1, value=label)
        cell = ws.cell(row=4 + offset, column=2, value=value)
        cell.number_format, cell.font = fmt, Font(bold=True)
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 16

    by_course: dict[str, list] = {}
    for m in made:
        entry = by_course.setdefault(m.row.course, [0, 0.0])
        entry[0] += 1
        entry[1] += m.row.amount
    ws.cell(row=16, column=1, value="Made, by course").font = Font(bold=True, size=12)
    table = [[course, count, round(fees, 2)] for course, (count, fees) in
             sorted(by_course.items(), key=lambda item: -item[1][1])]
    write_table(ws, ["Course", "Certificates and letters", "Fees"], table, start_row=17,
                formats={"Fees": MONEY, "Certificates and letters": "#,##0"})
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 14
    return wb


def main(argv=None) -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("spreadsheet", type=Path, help=".xlsx or .csv with name, course, date and amount columns")
    parser.add_argument("-o", "--output", type=Path, default=Path("output"), help="folder for the files and the log")
    parser.add_argument("--as-of", type=date.fromisoformat, default=None, metavar="YYYY-MM-DD",
                        help="date printed on the letters and certificates (default: today)")
    defaults = Settings()
    parser.add_argument("--org", default=defaults.org, help="organisation name on every file")
    parser.add_argument("--signer", default=defaults.signer, help="name above the signature line")
    parser.add_argument("--note", default=defaults.note, help="small footer on every file (\"\" for none)")
    args = parser.parse_args(argv)

    settings = Settings(org=args.org, signer=args.signer, note=args.note)
    result = generate(args.spreadsheet, args.output, settings, args.as_of)
    failed = [m for m in result.made if m.check != "OK"]
    print(f"Read {len(result.rows)} rows from {result.source}: made {len(result.made)} certificates and "
          f"{len(result.made)} letters, skipped {len(result.skipped)}.")
    for row in result.skipped:
        print(f"  Row {row.number} skipped: {'; '.join(row.problems)}")
    for made in failed:
        print(f"  Row {made.row.number} failed its check: {made.check}")
    if result.leftover:
        print(f"Note: {result.leftover} files from an earlier run are still in {args.output}. "
              "Use a new folder for each run.")
    print(f"Log: {args.output / 'generation_log.xlsx'}")


if __name__ == "__main__":
    main()
