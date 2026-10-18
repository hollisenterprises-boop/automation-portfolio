"""Turn a messy online-store order export into a clean, formatted Excel sales report.

    python sales_report.py sample_data/orders_export_messy.csv -o output/sales_report.xlsx

Reads the CSV exactly as the store exports it and fixes what it can: dates in
mixed formats, prices with currency text, quantities like "2 units", status and
state names written different ways, and stray spaces. It drops blank rows,
footer lines and exact duplicates, then writes an Excel report with sales by
month, product and state, plus a log of every fix it made.
"""

from __future__ import annotations

import argparse
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

REQUIRED_COLUMNS = [
    "Order ID", "Order Date", "SKU", "Product Name",
    "Qty", "Item Price", "Order Status", "Ship State",
]

SOLD_STATUSES = {"Shipped", "Delivered", "Completed"}

# Slashed dates are read day-first (Australian format).
DATE_FORMATS = (
    "%Y-%m-%d",
    "%Y-%m-%d %H:%M:%S",
    "%d/%m/%Y",
    "%d/%m/%Y %H:%M",
    "%d-%m-%Y",
    "%d %b %Y",
    "%d %B %Y",
    "%b %d %Y",
    "%B %d %Y",
)

STATUS_MAP = {
    "shipped": "Shipped", "delivered": "Delivered", "completed": "Completed", "complete": "Completed",
    "cancelled": "Cancelled", "canceled": "Cancelled",
    "refunded": "Refunded", "returned": "Refunded",
    "pending": "Pending", "processing": "Pending", "awaiting shipment": "Pending",
}

STATE_MAP = {
    "wa": "WA", "western australia": "WA",
    "nsw": "NSW", "new south wales": "NSW",
    "vic": "VIC", "victoria": "VIC",
    "qld": "QLD", "queensland": "QLD",
    "sa": "SA", "south australia": "SA",
    "tas": "TAS", "tasmania": "TAS",
    "act": "ACT", "australian capital territory": "ACT",
    "nt": "NT", "northern territory": "NT",
}

ORDER_ID_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9-]*$")
FIX_COLUMNS = ["Row", "Column", "Was", "Now", "What happened"]

BLANK = "Blank row removed"
FOOTER = "Not an order (summary or footer line), removed"
UNREADABLE = "Could not read this value, row left out"
DUPLICATE = "Exact duplicate of an earlier row, removed"
REMOVALS = {BLANK, FOOTER, UNREADABLE, DUPLICATE}

MONEY = '"$"#,##0.00'
WHOLE_MONEY = '"$"#,##0'
COUNT = "#,##0"
PERCENT = "0.0%"
HEADER_FILL = PatternFill("solid", fgColor="184F95")
HEADER_FONT = Font(bold=True, color="FFFFFF")
BAR_COLOUR = "2A78D6"


# ---------- reading single values ----------

def parse_date(value: str) -> date | None:
    text = " ".join(str(value).split())
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def parse_money(value: str) -> float | None:
    text = re.sub(r"(?i)aud|a\$|\$|,", "", str(value)).strip()
    if not text:
        return None
    try:
        return round(float(text), 2)
    except ValueError:
        return None


def parse_qty(value: str) -> int | None:
    match = re.search(r"\d+", str(value))
    return int(match.group()) if match else None


def normalize_status(value: str) -> str:
    return STATUS_MAP.get(" ".join(str(value).lower().split()), "Unknown")


def normalize_state(value: str) -> str:
    key = " ".join(re.sub(r"[^a-z ]", "", str(value).lower()).split())
    return STATE_MAP.get(key, "Unknown")


# ---------- cleaning ----------

def load_export(source) -> pd.DataFrame:
    """Read the export as text, keeping blank lines so they can be counted."""
    raw = pd.read_csv(source, dtype=str, keep_default_na=False,
                      skip_blank_lines=False, encoding="utf-8-sig")
    raw.columns = [str(c).strip() for c in raw.columns]
    return raw.fillna("")


def _fix(row, column, was, now, what):
    return {"Row": row, "Column": column, "Was": was, "Now": now, "What happened": what}


def _log_changes(log, line, raw, clean):
    checks = [
        ("Order Date", clean["Order Date"].isoformat(), "Date written as YYYY-MM-DD"),
        ("Item Price", f"{clean['Item Price']:.2f}", "Currency text removed"),
        ("Qty", str(clean["Qty"]), "Quantity read as a number"),
        ("SKU", clean["SKU"], "Spaces and capitals fixed"),
        ("Product Name", clean["Product Name"], "Extra spaces removed"),
        ("Order Status", clean["Order Status"], "Status written one way"),
        ("Ship State", clean["Ship State"], "State written as its code"),
    ]
    for column, now, what in checks:
        if raw[column] != now:
            if now == "Unknown":
                what = "Not recognised, marked Unknown"
            log.append(_fix(line, column, raw[column], now, what))


def clean_orders(raw: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return (clean order lines, log of every fix and removal)."""
    missing = [c for c in REQUIRED_COLUMNS if c not in raw.columns]
    if missing:
        raise ValueError("The export is missing these columns: " + ", ".join(missing))

    log: list[dict] = []
    rows: list[dict] = []
    for index, record in raw.iterrows():
        line = index + 2  # header is line 1
        values = {c: str(record[c]) for c in REQUIRED_COLUMNS}
        if not any(v.strip() for v in values.values()):
            log.append(_fix(line, "", "", "", BLANK))
            continue
        if not ORDER_ID_RE.match(values["Order ID"].strip()):
            log.append(_fix(line, "Order ID", values["Order ID"], "", FOOTER))
            continue

        parsed = {
            "Order Date": parse_date(values["Order Date"]),
            "Item Price": parse_money(values["Item Price"]),
            "Qty": parse_qty(values["Qty"]),
        }
        unreadable = [col for col, val in parsed.items() if val is None]
        if unreadable:
            for col in unreadable:
                log.append(_fix(line, col, values[col], "", UNREADABLE))
            continue

        clean = {
            "Order ID": values["Order ID"].strip(),
            "Order Date": parsed["Order Date"],
            "SKU": values["SKU"].strip().upper(),
            "Product Name": " ".join(values["Product Name"].split()),
            "Qty": parsed["Qty"],
            "Item Price": parsed["Item Price"],
            "Order Status": normalize_status(values["Order Status"]),
            "Ship State": normalize_state(values["Ship State"]),
        }
        _log_changes(log, line, values, clean)
        clean["Source Row"] = line
        rows.append(clean)

    clean = pd.DataFrame(rows, columns=REQUIRED_COLUMNS + ["Source Row"])
    # Same order, product, quantity and price twice = an export glitch.
    duplicate = clean.duplicated(subset=REQUIRED_COLUMNS, keep="first")
    for line in clean.loc[duplicate, "Source Row"]:
        log.append(_fix(int(line), "", "", "", DUPLICATE))
    clean = clean.loc[~duplicate].reset_index(drop=True)
    clean["Order Date"] = pd.to_datetime(clean["Order Date"])
    clean["Line Total"] = (clean["Qty"] * clean["Item Price"]).round(2)

    fixes = pd.DataFrame(log, columns=FIX_COLUMNS).sort_values("Row", kind="stable")
    return clean, fixes.reset_index(drop=True)


# ---------- summarising ----------

def summarise(clean: pd.DataFrame) -> dict:
    done = clean[clean["Order Status"].isin(SOLD_STATUSES)]
    sales = round(done["Line Total"].sum(), 2)
    orders = done["Order ID"].nunique()

    month = done["Order Date"].dt.to_period("M")
    by_month = done.groupby(month).agg(
        Orders=("Order ID", "nunique"), Units=("Qty", "sum"), Sales=("Line Total", "sum"))
    by_month["Average Order"] = by_month["Sales"] / by_month["Orders"]
    by_month.index = [p.strftime("%b %Y") for p in by_month.index]
    by_month = by_month.rename_axis("Month").reset_index()

    names = done.groupby("SKU")["Product Name"].agg(lambda s: s.mode().iat[0])
    by_product = done.groupby("SKU").agg(Units=("Qty", "sum"), Sales=("Line Total", "sum"))
    by_product.insert(0, "Product", names)
    by_product["Share"] = by_product["Sales"] / sales
    by_product = by_product.sort_values("Sales", ascending=False).reset_index()

    by_state = done.groupby("Ship State").agg(
        Orders=("Order ID", "nunique"), Sales=("Line Total", "sum"))
    by_state["Share"] = by_state["Sales"] / sales
    by_state = by_state.sort_values("Sales", ascending=False).rename_axis("State").reset_index()

    def value_of(status):
        return round(clean.loc[clean["Order Status"] == status, "Line Total"].sum(), 2)

    return {
        "first_day": clean["Order Date"].min(),
        "last_day": clean["Order Date"].max(),
        "sales": sales,
        "orders": orders,
        "units": int(done["Qty"].sum()),
        "average_order": sales / orders if orders else 0.0,
        "refunded": value_of("Refunded"),
        "cancelled_orders": clean.loc[clean["Order Status"] == "Cancelled", "Order ID"].nunique(),
        "pending": value_of("Pending"),
        "by_month": by_month,
        "by_product": by_product,
        "by_state": by_state,
    }


# ---------- writing Excel ----------

def write_table(ws, df: pd.DataFrame, start_row: int = 1, formats: dict | None = None) -> int:
    """Write df with a styled header at start_row; return the last row written."""
    formats = formats or {}
    for col, name in enumerate(df.columns, start=1):
        cell = ws.cell(row=start_row, column=col, value=name)
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for r, record in enumerate(df.itertuples(index=False), start=start_row + 1):
        for col, value in enumerate(record, start=1):
            if isinstance(value, pd.Timestamp):
                value = value.date()
            cell = ws.cell(row=r, column=col, value=value)
            fmt = formats.get(df.columns[col - 1])
            if fmt:
                cell.number_format = fmt
            elif isinstance(value, date):
                cell.number_format = "dd/mm/yyyy"
    for col, name in enumerate(df.columns, start=1):
        width = max([len(str(name))] + [len(str(v)) for v in df.iloc[:, col - 1].head(500)])
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = max(ws.column_dimensions[letter].width or 0,
                                                 min(width + 3, 48))
    return start_row + len(df)


def build_workbook(clean, fixes, summary, store, source_name) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    months = wb.create_sheet("By Month")
    write_table(months, summary["by_month"],
                formats={"Orders": COUNT, "Units": COUNT, "Sales": MONEY, "Average Order": MONEY})
    products = wb.create_sheet("By Product")
    write_table(products, summary["by_product"],
                formats={"Units": COUNT, "Sales": MONEY, "Share": PERCENT})
    states = wb.create_sheet("By State")
    write_table(states, summary["by_state"],
                formats={"Orders": COUNT, "Sales": MONEY, "Share": PERCENT})
    data = wb.create_sheet("Clean Data")
    write_table(data, clean, formats={"Item Price": MONEY, "Line Total": MONEY})
    log = wb.create_sheet("Fixes Log")
    write_table(log, fixes)
    for sheet in (months, products, states, data, log):
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions

    first, last = summary["first_day"], summary["last_day"]
    ws["A1"] = f"Sales report: {store}"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"] = f"{first:%d %b %Y} to {last:%d %b %Y}. Built from {source_name} on {date.today():%d %b %Y}."
    ws["A2"].font = Font(color="52514E")

    removal = fixes["What happened"].isin(REMOVALS)
    kpis = [
        ("Net sales (shipped and delivered)", summary["sales"], MONEY),
        ("Orders", summary["orders"], COUNT),
        ("Units sold", summary["units"], COUNT),
        ("Average order", summary["average_order"], MONEY),
        ("Refunded", summary["refunded"], MONEY),
        ("Cancelled orders", summary["cancelled_orders"], COUNT),
        ("Pending (not shipped yet)", summary["pending"], MONEY),
        ("Values fixed in the export", int((~removal).sum()), COUNT),
        ("Rows removed (blank, footer, duplicate, unreadable)",
         int(fixes.loc[removal, "Row"].nunique()), COUNT),
    ]
    for offset, (label, value, fmt) in enumerate(kpis):
        row = 4 + offset
        ws.cell(row=row, column=1, value=label)
        cell = ws.cell(row=row, column=2, value=value)
        cell.number_format = fmt
        cell.font = Font(bold=True)
    ws.column_dimensions["A"].width = 50
    ws.column_dimensions["B"].width = 16

    top = summary["by_product"].head(5)[["Product", "Units", "Sales"]]
    ws.cell(row=15, column=1, value="Top 5 products").font = Font(bold=True, size=12)
    write_table(ws, top, start_row=16, formats={"Units": COUNT, "Sales": MONEY})
    ws.column_dimensions["A"].width = 50

    chart = BarChart()
    chart.title = "Net sales by month"
    chart.legend = None
    chart.height, chart.width = 8, 16
    count = len(summary["by_month"])
    chart.add_data(Reference(months, min_col=4, min_row=1, max_row=count + 1), titles_from_data=True)
    chart.set_categories(Reference(months, min_col=1, min_row=2, max_row=count + 1))
    chart.series[0].graphicalProperties.solidFill = BAR_COLOUR
    chart.series[0].graphicalProperties.line.solidFill = BAR_COLOUR
    chart.y_axis.numFmt = WHOLE_MONEY
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.gapWidth = 60
    ws.add_chart(chart, "D4")
    return wb


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("export", type=Path, help="order export CSV")
    parser.add_argument("-o", "--output", type=Path, default=Path("sales_report.xlsx"))
    parser.add_argument("--store", default="Sample Store", help="name shown on the report")
    args = parser.parse_args(argv)

    clean, fixes = clean_orders(load_export(args.export))
    summary = summarise(clean)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    build_workbook(clean, fixes, summary, args.store, args.export.name).save(args.output)

    print(f"Net sales ${summary['sales']:,.2f} from {summary['orders']:,} orders "
          f"({summary['first_day']:%d %b %Y} to {summary['last_day']:%d %b %Y})")
    print(f"{len(fixes):,} fixes and removals logged. Report: {args.output}")


if __name__ == "__main__":
    main()
