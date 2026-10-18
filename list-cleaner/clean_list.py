"""Merge messy contact lists into one clean list, with duplicates found and merged.

    python clean_list.py sample_data -o output/clean_contacts.xlsx

Reads every CSV and Excel file in the folder, works out which column is which
from the headers, and cleans each value: names (case, titles, "Smith, Olivia"),
emails (case, spaces, "mailto:"), Australian phone numbers (one format), and
suburb, state and postcode. Then it finds the same person across lists by
email or phone, and by a similar name in the same suburb as a careful last
resort. Marketing consent is treated strictly: one "no" or "unsubscribed"
anywhere means no. Values it can't fix are listed for a person to check.
"""

from __future__ import annotations

import argparse
import csv
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

HEADER_ALIASES = {
    "first": {"first name", "firstname", "given name"},
    "last": {"last name", "lastname", "surname", "family name"},
    "name": {"name", "full name", "contact name", "customer name"},
    "email": {"email", "e-mail", "email address", "e-mail address"},
    "phone": {"phone", "mobile", "phone no", "phone number", "mobile number", "mobile phone"},
    "suburb": {"suburb", "city", "town", "location"},
    "state": {"state", "region"},
    "opt_in": {"accepts marketing", "keep me updated", "subscribed", "opt in", "marketing"},
    "opt_out": {"unsubscribed", "opted out", "do not contact"},
}
STATES = {"WA": "Western Australia", "NSW": "New South Wales", "VIC": "Victoria", "QLD": "Queensland",
          "SA": "South Australia", "TAS": "Tasmania", "ACT": "Australian Capital Territory",
          "NT": "Northern Territory"}
STATE_BY_NAME = {name.lower(): code for code, name in STATES.items()}
TITLES = {"mr", "mrs", "ms", "miss", "dr", "mx", "prof"}
NICKNAMES = {"kate": "katherine", "kathy": "katherine", "mike": "michael", "liz": "elizabeth",
             "beth": "elizabeth", "rob": "robert", "bob": "robert", "jen": "jennifer", "alex": "alexander",
             "pat": "patrick", "will": "william", "bill": "william", "tom": "thomas", "ben": "benjamin",
             "dan": "daniel", "josh": "joshua", "sam": ("samuel", "samantha"), "jim": "james", "chris": "christopher"}
TRUTHY = {"yes", "y", "true", "1", "x", "subscribed"}
FALSY = {"no", "n", "false", "0"}

EMAIL_RE = re.compile(r"^[a-z0-9._%+'-]+@[a-z0-9-]+(\.[a-z0-9-]+)*\.[a-z]{2,}$")
PLACE_RE = re.compile(r"^(?P<suburb>[A-Za-z .'-]+?),?\s+(?P<state>[A-Za-z]{2,3})(?:\s+(?P<postcode>\d{4}))?$")

HEADER_FILL = PatternFill("solid", fgColor="184F95")
HEADER_FONT = Font(bold=True, color="FFFFFF")
CHECK_FILL = PatternFill("solid", fgColor="FDE2E1")
EXACT, SIMILAR = "Same email or phone", "Similar name, same suburb"


@dataclass
class Record:
    source: str
    row: int
    raw: dict
    first: str = ""
    last: str = ""
    email: str = ""
    phone: str = ""
    suburb: str = ""
    state: str = ""
    postcode: str = ""
    consent: bool | None = None
    problems: list[str] = field(default_factory=list)
    fixes: list[str] = field(default_factory=list)


@dataclass
class Contact:
    first: str
    last: str
    email: str
    phone: str
    suburb: str
    state: str
    postcode: str
    marketing: str  # Yes, No or Unknown
    records: list[Record]
    match: str  # "Single record", EXACT or SIMILAR

    @property
    def sources(self) -> str:
        return ", ".join(sorted({r.source for r in self.records}))


# ---------------------------------------------------------------- cleaning values

def proper(word: str) -> str:
    """Title-case a name, keeping O'Brien, McDonald and Smith-Jones right."""
    if not word or (not word.isupper() and not word.islower()):
        return word  # already mixed case: trust it
    parts = []
    for piece in word.lower().split("-"):
        piece = piece[:1].upper() + piece[1:]
        if piece.startswith("Mc") and len(piece) > 2:
            piece = "Mc" + piece[2].upper() + piece[3:]
        if piece.startswith("O'") and len(piece) > 2:
            piece = "O'" + piece[2].upper() + piece[3:]
        parts.append(piece)
    return "-".join(parts)


def clean_name(full: str) -> tuple[str, str]:
    """'  MRS olivia SMITH ' or 'Smith, Olivia' -> ('Olivia', 'Smith')."""
    text = " ".join(full.replace(" ", " ").split()).strip(" .,;")
    if "," in text:
        last, first = (p.strip() for p in text.split(",", 1))
        text = f"{first} {last}"
    tokens = text.split()
    while len(tokens) > 1 and tokens[0].lower().rstrip(".") in TITLES:
        tokens = tokens[1:]
    if not tokens:
        return "", ""
    return proper(tokens[0]), " ".join(proper(t) for t in tokens[1:])


def clean_email(raw: str) -> tuple[str, str | None]:
    text = raw.strip()
    if text.lower().startswith("mailto:"):
        text = text[7:]
    text = text.strip().lower()
    if not text:
        return "", None
    if EMAIL_RE.match(text):
        return text, None
    return "", f'Email "{raw.strip()}" is not a valid address'


def clean_phone(raw: str) -> tuple[str, str | None]:
    digits = re.sub(r"\D", "", raw)
    if not digits:
        return "", None
    if digits.startswith("61") and len(digits) == 11:
        digits = "0" + digits[2:]
    if len(digits) == 9 and digits[0] in "23478":
        digits = "0" + digits  # leading zero lost, often by Excel
    if len(digits) == 10 and digits.startswith("0") and digits[1] in "23478":
        if digits.startswith("04"):
            return f"{digits[:4]} {digits[4:7]} {digits[7:]}", None
        return f"{digits[:2]} {digits[2:6]} {digits[6:]}", None
    return "", f'Phone "{raw.strip()}" is not a valid Australian number'


def clean_state(raw: str) -> str:
    text = raw.strip().rstrip(".")
    if text.upper() in STATES:
        return text.upper()
    return STATE_BY_NAME.get(text.lower(), "")


def clean_place(suburb_raw: str, state_raw: str) -> tuple[str, str, str]:
    suburb, state, postcode = " ".join(suburb_raw.split()), clean_state(state_raw), ""
    match = PLACE_RE.match(suburb)
    if match and clean_state(match["state"]):
        suburb, postcode = match["suburb"].strip(), match["postcode"] or ""
        state = state or clean_state(match["state"])
    suburb = " ".join(proper(w) for w in suburb.split())
    return suburb, state, postcode


def flag(value: str) -> bool | None:
    text = str(value).strip().lower()
    if text in TRUTHY:
        return True
    if text in FALSY:
        return False
    return None


# ---------------------------------------------------------------- reading files

def read_rows(path: Path) -> list[dict]:
    if path.suffix.lower() == ".csv":
        with path.open(encoding="utf-8-sig", newline="") as f:
            return list(csv.DictReader(f))
    ws = load_workbook(path, read_only=True, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    header = [str(h or "").strip() for h in rows[0]]
    return [{h: "" if v is None else str(v) for h, v in zip(header, r)} for r in rows[1:]]


def map_columns(headers: list[str]) -> dict[str, str]:
    mapping = {}
    for header in headers:
        key = " ".join(header.replace("_", " ").lower().split())
        for field_name, aliases in HEADER_ALIASES.items():
            if key in aliases and field_name not in mapping:
                mapping[field_name] = header
    return mapping


def to_record(source: str, row_number: int, row: dict, cols: dict[str, str]) -> Record:
    get = lambda name: str(row.get(cols.get(name, ""), "") or "")  # noqa: E731
    rec = Record(source=source, row=row_number, raw=row)
    if "first" in cols or "last" in cols:
        first, last = clean_name(f"{get('first')} {get('last')}")
        if get("first").strip() and get("last").strip():
            first, last = proper(" ".join(get("first").split())), proper(" ".join(get("last").split()))
    else:
        first, last = clean_name(get("name"))
    rec.first, rec.last = first, last
    rec.email, problem = clean_email(get("email"))
    if problem:
        rec.problems.append(problem)
    elif rec.email and rec.email != get("email"):
        rec.fixes.append("email")
    rec.phone, problem = clean_phone(get("phone"))
    if problem:
        rec.problems.append(problem)
    elif rec.phone and rec.phone != get("phone"):
        rec.fixes.append("phone")
    rec.suburb, rec.state, rec.postcode = clean_place(get("suburb"), get("state"))
    if "opt_out" in cols:
        rec.consent = False if flag(get("opt_out")) else None
    elif "opt_in" in cols:
        rec.consent = flag(get("opt_in")) if get("opt_in").strip() else None
    return rec


def read_folder(folder: Path) -> tuple[list[Record], dict[str, int], int]:
    records, rows_per_file, blanks = [], {}, 0
    for path in sorted(p for p in folder.iterdir() if p.suffix.lower() in {".csv", ".xlsx"}):
        rows = read_rows(path)
        rows_per_file[path.name] = len(rows)
        cols = map_columns(list(rows[0]) if rows else [])
        for n, row in enumerate(rows, start=2):
            if not any(str(v).strip() for v in row.values()):
                blanks += 1
                continue
            records.append(to_record(path.name, n, row, cols))
    return records, rows_per_file, blanks


# ---------------------------------------------------------------- finding duplicates

def same_first_name(a: str, b: str) -> bool:
    a, b = a.lower(), b.lower()
    if a == b:
        return True
    formal = lambda n: NICKNAMES.get(n, n)  # noqa: E731
    fa, fb = formal(a), formal(b)
    options_a = set(fa) if isinstance(fa, tuple) else {fa}
    options_b = set(fb) if isinstance(fb, tuple) else {fb}
    if options_a & options_b:
        return True
    return SequenceMatcher(None, a, b).ratio() >= 0.85


def group_records(records: list[Record]) -> list[tuple[list[Record], str]]:
    parent = list(range(len(records)))
    reason = {}

    def root(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def join(i, j, why):
        ri, rj = root(i), root(j)
        if ri != rj:
            parent[rj] = ri
            reason[ri] = SIMILAR if SIMILAR in (why, reason.get(ri), reason.get(rj)) else EXACT

    for key in ("email", "phone"):
        first_seen = {}
        for i, rec in enumerate(records):
            value = getattr(rec, key)
            if value:
                if value in first_seen:
                    join(first_seen[value], i, EXACT)
                else:
                    first_seen[value] = i

    def details(r):
        members = [records[k] for k in range(len(records)) if root(k) == r]
        return {m.email for m in members if m.email}, {m.phone for m in members if m.phone}

    buckets = defaultdict(list)
    for i, rec in enumerate(records):
        if rec.last and rec.suburb:
            buckets[(rec.last.lower(), rec.suburb.lower())].append(i)
    for members in buckets.values():
        for x in range(len(members)):
            for y in range(x + 1, len(members)):
                i, j = members[x], members[y]
                if root(i) == root(j) or not same_first_name(records[i].first, records[j].first):
                    continue
                (emails_i, phones_i), (emails_j, phones_j) = details(root(i)), details(root(j))
                conflict = (emails_i and emails_j and not emails_i & emails_j) or \
                           (phones_i and phones_j and not phones_i & phones_j)
                if not conflict:  # different contact details means different people
                    join(i, j, SIMILAR)

    groups = defaultdict(list)
    for i in range(len(records)):
        groups[root(i)].append(records[i])
    return [(members, reason.get(r, "Single record") if len(members) > 1 else "Single record")
            for r, members in groups.items()]


def most_common(values) -> str:
    values = [v for v in values if v]
    return Counter(values).most_common(1)[0][0] if values else ""


def merge(members: list[Record], match: str) -> Contact:
    firsts = [m.first for m in members if m.first]
    first = max(firsts, key=lambda n: (len(n), firsts.count(n))) if firsts else ""  # Katherine over Kate
    consents = [m.consent for m in members if m.consent is not None]
    marketing = "No" if False in consents else "Yes" if True in consents else "Unknown"
    return Contact(first=first, last=most_common(m.last for m in members),
                   email=most_common(m.email for m in members), phone=most_common(m.phone for m in members),
                   suburb=most_common(m.suburb for m in members), state=most_common(m.state for m in members),
                   postcode=most_common(m.postcode for m in members), marketing=marketing,
                   records=members, match=match)


def clean(folder: Path):
    records, rows_per_file, blanks = read_folder(folder)
    contacts = [merge(members, match) for members, match in group_records(records)]
    contacts.sort(key=lambda c: (c.last.lower(), c.first.lower()))
    return contacts, records, rows_per_file, blanks


# ---------------------------------------------------------------- Excel output

def table(ws, headers, rows, widths):
    ws.append(headers)
    for cell in ws[ws.max_row]:
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
    for row in rows:
        ws.append(row)
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(ws.max_row, 2)}"


def build_workbook(contacts, records, rows_per_file, blanks) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    merged = [c for c in contacts if len(c.records) > 1]
    no_contact = [c for c in contacts if not c.email and not c.phone]
    record_problems = [(r, p) for r in records for p in r.problems]
    ws["A1"] = "Contact list clean-up"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"] = f"{sum(rows_per_file.values())} rows from {len(rows_per_file)} files"
    lines = [("People in the clean list", len(contacts)),
             ("Duplicate rows merged", len(records) - len(contacts)),
             ("  merged on the same email or phone", sum(len(c.records) - 1 for c in merged if c.match == EXACT)),
             ("  merged on a similar name in the same suburb (check these)", sum(len(c.records) - 1 for c in merged if c.match == SIMILAR)),
             ("Blank rows removed", blanks),
             ("Emails tidied (case, spaces, mailto:)", sum("email" in r.fixes for r in records)),
             ("Phone numbers put in one format", sum("phone" in r.fixes for r in records)),
             ("Values to check (bad email or phone)", len(record_problems)),
             ("People with no email or phone", len(no_contact)),
             ("OK to market to", sum(c.marketing == "Yes" for c in contacts)),
             ("Said no or unsubscribed", sum(c.marketing == "No" for c in contacts))]
    for n, (label, value) in enumerate(lines, start=4):
        ws.cell(n, 1, label)
        ws.cell(n, 2, value).font = Font(bold=True)
    ws.cell(17, 1, "Rows read").font = Font(bold=True, size=12)
    for n, (name, count) in enumerate(rows_per_file.items(), start=18):
        ws.cell(n, 1, name)
        ws.cell(n, 2, count)
    ws.column_dimensions["A"].width = 58
    ws.column_dimensions["B"].width = 12

    table(wb.create_sheet("Clean List"),
          ["First Name", "Last Name", "Email", "Phone", "Suburb", "State", "Postcode", "Marketing OK",
           "Rows merged", "Found in", "Merged because"],
          [[c.first, c.last, c.email, c.phone, c.suburb, c.state, c.postcode, c.marketing, len(c.records),
            c.sources, c.match] for c in contacts],
          [14, 16, 34, 15, 15, 7, 10, 13, 12, 40, 26])
    review = wb.create_sheet("Merged Rows")
    table(review, ["Person", "File", "Row", "Name as written", "Email as written", "Phone as written",
                   "Merged because"],
          [[f"{c.first} {c.last}", r.source, r.row,
            " ".join(str(r.raw.get(k, "")) for k in r.raw if "name" in k.lower()).strip(),
            next((str(v) for k, v in r.raw.items() if "mail" in k.lower()), ""),
            next((str(v) for k, v in r.raw.items() if k.lower().replace("_", " ") in HEADER_ALIASES["phone"]), ""),
            c.match]
           for c in merged for r in c.records],
          [22, 22, 6, 28, 36, 20, 26])
    for row in range(2, review.max_row + 1):
        if review.cell(row, 7).value == SIMILAR:
            for col in range(1, 8):
                review.cell(row, col).fill = CHECK_FILL
    table(wb.create_sheet("To Check"), ["File", "Row", "Person", "What to check"],
          [[r.source, r.row, f"{r.first} {r.last}", p] for r, p in record_problems] +
          [[c.sources, "", f"{c.first} {c.last}", "No email or phone in any list"] for c in no_contact],
          [22, 6, 24, 70])
    return wb


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("folder", type=Path, help="folder of CSV and Excel contact lists")
    parser.add_argument("-o", "--output", type=Path, default=Path("output/clean_contacts.xlsx"))
    args = parser.parse_args()
    contacts, records, rows_per_file, blanks = clean(args.folder)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    build_workbook(contacts, records, rows_per_file, blanks).save(args.output)
    print(f"{sum(rows_per_file.values())} rows -> {len(contacts)} people "
          f"({len(records) - len(contacts)} duplicates merged) -> {args.output}")


if __name__ == "__main__":
    main()
