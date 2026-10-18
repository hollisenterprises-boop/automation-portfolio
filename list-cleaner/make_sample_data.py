"""Write three messy contact lists, like a small business ends up with.

    python make_sample_data.py

- shop_customers.csv: the online store's customer export
- event_signups.xlsx: a sign-up form from a market stall
- old_crm_export.csv: a contact list from an old CRM

The same people turn up in more than one list, written differently each time.
All data is made up: emails use example.com/.net/.org, and phone numbers are in
the ranges Australia reserves for fiction (0491 570 xxx-style mobiles and
5550 xxxx landlines). The random seed is fixed, so every run gives the same files.
"""

from __future__ import annotations

import csv
import random
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook

OUT = Path(__file__).parent / "sample_data"
rng = random.Random(7)

FIRST = ["Olivia", "Jack", "Charlotte", "Noah", "Amelia", "William", "Isla", "Oliver", "Mia", "Thomas",
         "Grace", "James", "Chloe", "Lucas", "Ava", "Henry", "Sophie", "Liam", "Zoe", "Samuel",
         "Ruby", "Ethan", "Ella", "Benjamin", "Harper", "Daniel", "Matilda", "Joshua", "Katherine", "Michael",
         "Elizabeth", "Robert", "Jennifer", "Alexander", "Samantha", "Patrick"]
LAST = ["Smith", "Nguyen", "Brown", "Wilson", "Taylor", "Johnson", "White", "Martin", "Anderson", "Thompson",
        "Walker", "Harris", "Lee", "Ryan", "Robinson", "Kelly", "King", "Davies", "Wright", "Mitchell",
        "O'Brien", "McDonald", "Campbell", "Scott", "Murphy", "Hughes", "Edwards", "Clarke", "Price", "Turner"]
NICKNAMES = {"Katherine": "Kate", "Michael": "Mike", "Elizabeth": "Liz", "Robert": "Rob", "Jennifer": "Jen",
             "Alexander": "Alex", "Samantha": "Sam", "Patrick": "Pat", "William": "Will", "Thomas": "Tom",
             "Benjamin": "Ben", "Daniel": "Dan", "Joshua": "Josh", "Samuel": "Sam"}
PLACES = [("Perth", "WA", "6000"), ("Fremantle", "WA", "6160"), ("Joondalup", "WA", "6027"),
          ("Scarborough", "WA", "6019"), ("Subiaco", "WA", "6008"), ("Mandurah", "WA", "6210"),
          ("Bunbury", "WA", "6230"), ("Melbourne", "VIC", "3000"), ("Geelong", "VIC", "3220"),
          ("Brisbane", "QLD", "4000"), ("Gold Coast", "QLD", "4217"), ("Adelaide", "SA", "5000"),
          ("Sydney", "NSW", "2000"), ("Newcastle", "NSW", "2300"), ("Hobart", "TAS", "7000")]
STATE_NAMES = {"WA": "Western Australia", "VIC": "Victoria", "QLD": "Queensland", "SA": "South Australia",
               "NSW": "New South Wales", "TAS": "Tasmania"}
AREA_CODES = {"WA": "08", "SA": "08", "VIC": "03", "TAS": "03", "QLD": "07", "NSW": "02"}
DOMAINS = ["example.com", "example.net", "example.org"]


@dataclass
class Person:
    pid: int
    first: str
    last: str
    email: str
    phone: str  # digits, national format, e.g. 0855501234
    place: tuple[str, str, str]
    marketing: bool


def make_people(n: int = 150) -> list[Person]:
    people, used, emails = [], set(), set()
    while len(people) < n:
        first, last = rng.choice(FIRST), rng.choice(LAST)
        place = rng.choice(PLACES)
        local = f"{first}.{last}".lower().replace("'", "")
        email = f"{local}{rng.randint(1, 99)}@{rng.choice(DOMAINS)}"
        if (first, last, place) in used or email in emails:
            continue
        used.add((first, last, place))
        emails.add(email)
        if rng.random() < 0.15:  # a mobile in the fictional 0491 570 xxx block
            phone = f"0491570{rng.randint(100, 999)}"
        else:  # a landline in the fictional 5550 block for their state
            phone = f"{AREA_CODES[place[1]]}5550{rng.randint(0, 9999):04d}"
        people.append(Person(len(people) + 1, first, last, email, phone, place, rng.random() < 0.6))
    return people


# ------------------------------------------------------------ ways of writing it

def messy_phone(p: str) -> str:
    national = p
    if national.startswith("04"):
        styles = [f"{p[:4]} {p[4:7]} {p[7:]}", f"+61 {p[1:4]} {p[4:7]} {p[7:]}", p, f"61{p[1:]}"]
    else:
        styles = [f"({p[:2]}) {p[2:6]} {p[6:]}", f"{p[:2]} {p[2:6]}-{p[6:]}", f"+61 {p[1]} {p[2:6]} {p[6:]}", national]
    return rng.choice(styles)


def messy_name(first: str, last: str) -> str:
    style = rng.random()
    if style < 0.15:
        return f"{last}, {first}"
    if style < 0.25:
        return f"{first} {last}".upper()
    if style < 0.35:
        return f"{first} {last}".lower()
    if style < 0.42:
        return f"{rng.choice(['Mr', 'Mrs', 'Ms', 'Dr'])} {first} {last}"
    if style < 0.50:
        return f"  {first}  {last} "
    return f"{first} {last}"


def messy_email(email: str) -> str:
    style = rng.random()
    if style < 0.15:
        return email.upper()
    if style < 0.22:
        return f" {email} "
    if style < 0.27:
        return f"mailto:{email}"
    return email


def messy_place(place) -> tuple[str, str]:
    suburb, state, postcode = place
    style = rng.random()
    if style < 0.2:
        return suburb.upper(), STATE_NAMES[state]
    if style < 0.35:
        return f"{suburb}, {state} {postcode}", ""
    if style < 0.45:
        return suburb.lower(), state.lower()
    return suburb, state


# ------------------------------------------------------------ the three lists

def shop_rows(people):
    rows = []
    for p in people:
        suburb, state = messy_place(p.place)
        rows.append({"First Name": p.first if rng.random() > 0.1 else p.first.lower(), "Last Name": p.last,
                     "Email": messy_email(p.email), "Phone": messy_phone(p.phone),
                     "Suburb": suburb, "State": state, "Accepts Marketing": "yes" if p.marketing else "no"})
    return rows


def signup_rows(people):
    rows = []
    for p in people:
        first = NICKNAMES.get(p.first, p.first) if rng.random() < 0.5 else p.first
        email = messy_email(p.email) if rng.random() < 0.7 else ""
        rows.append({"Name": messy_name(first, p.last), "Email Address": email,
                     "Mobile": messy_phone(p.phone), "Suburb": messy_place(p.place)[0],
                     "Keep me updated": "Y" if p.marketing else ""})
    return rows


def crm_rows(people):
    rows = []
    for p in people:
        suburb, state = messy_place(p.place)
        rows.append({"CONTACT_NAME": messy_name(p.first, p.last), "EMAIL": messy_email(p.email),
                     "PHONE_NO": messy_phone(p.phone) if rng.random() < 0.85 else "",
                     "CITY": suburb, "STATE": state, "UNSUBSCRIBED": "" if p.marketing else "TRUE"})
    return rows


def add_problems(shop, signups, crm):
    """A few rows that can't be fixed automatically, and some junk."""
    shop[3]["Email"] = shop[3]["Email"].split("@")[0] + "@example"      # no ".com"
    shop[11]["Phone"] = "0855 501"                                       # too short
    signups[2]["Email Address"] = "see notes"
    crm[5]["EMAIL"] = crm[5]["EMAIL"].replace(".", " ", 1)               # space in the address
    crm[8]["PHONE_NO"], crm[8]["EMAIL"] = "", ""                         # no way to contact
    crm.insert(14, {k: "" for k in crm[0]})                              # blank row
    shop.insert(20, {k: "" for k in shop[0]})


def lists(people):
    """Who goes where: overlaps between the lists are the duplicates to find."""
    ids = [p.pid for p in people]
    shop_ids = set(rng.sample(ids, 95))
    signup_ids = set(rng.sample(sorted(shop_ids), 25)) | set(rng.sample([i for i in ids if i not in shop_ids], 20))
    crm_ids = set(rng.sample(sorted(shop_ids), 20)) | set(rng.sample([i for i in ids if i not in shop_ids], 35))
    by_id = {p.pid: p for p in people}
    pick = lambda chosen: [by_id[i] for i in sorted(chosen)]  # noqa: E731
    return pick(shop_ids), pick(signup_ids), pick(crm_ids)


def main(out: Path = OUT) -> None:
    rng.seed(7)  # the same files on every run
    people = make_people()
    shop_people, signup_people, crm_people = lists(people)
    shop, signups, crm = shop_rows(shop_people), signup_rows(signup_people), crm_rows(crm_people)
    add_problems(shop, signups, crm)

    out.mkdir(parents=True, exist_ok=True)
    for name, rows in (("shop_customers.csv", shop), ("old_crm_export.csv", crm)):
        with (out / name).open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=list(rows[0]))
            writer.writeheader()
            writer.writerows(rows)
    wb = Workbook()
    ws = wb.active
    ws.title = "Responses"
    ws.append(list(signups[0]))
    for row in signups:
        ws.append(list(row.values()))
    wb.save(out / "event_signups.xlsx")

    everyone = {p.pid for p in shop_people} | {p.pid for p in signup_people} | {p.pid for p in crm_people}
    print(f"{len(shop)} + {len(signups)} + {len(crm)} rows written for {len(everyone)} real people")


if __name__ == "__main__":
    main()
