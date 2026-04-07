#!/usr/bin/env python3
"""Terminal tool to import Medifox Excel exports into SQLite and export phone list CSV."""

from __future__ import annotations

import csv
import hashlib
import sqlite3
from dataclasses import dataclass, field
from pathlib import Path
import re

import openpyxl


NBSP = "\xa0"
DB_PATH = Path("phonebook.db")


@dataclass
class ContactRecord:
    lastname: str = ""
    firstname: str = ""
    relation: str = ""
    phone: str = ""
    mobile: str = ""


@dataclass
class CustomerRecord:
    lastname: str = ""
    firstname: str = ""
    carelevel: str = ""
    phone: str = ""
    mobile: str = ""
    external_id: str = ""
    contacts: list[ContactRecord] = field(default_factory=list)


def clean(value: object) -> str:
    if value is None:
        return ""
    text = str(value).replace(NBSP, " ").strip()
    return re.sub(r"\s+", " ", text)


def normalize_phone(value: object) -> str:
    text = clean(value)
    if not text:
        return ""
    # Keep only useful phone characters for matching while preserving leading +.
    text = text.replace("(0)", "0")
    return re.sub(r"[^0-9+/ -]", "", text).strip()


def normalized_phone_digits(value: str) -> str:
    return re.sub(r"\D", "", value or "")


def split_name(value: str) -> tuple[str, str]:
    value = clean(value)
    if not value:
        return "", ""
    if "," in value:
        last, first = value.split(",", 1)
        return clean(last), clean(first)
    parts = value.split()
    if len(parts) == 1:
        return parts[0], ""
    return " ".join(parts[:-1]), parts[-1]


def compute_customer_key(rec: CustomerRecord) -> str:
    raw = "|".join(
        [
            rec.external_id.lower(),
            rec.lastname.lower(),
            rec.firstname.lower(),
            normalized_phone_digits(rec.phone),
            normalized_phone_digits(rec.mobile),
        ]
    )
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


def init_db(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        PRAGMA foreign_keys = ON;

        CREATE TABLE IF NOT EXISTS customers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_key TEXT NOT NULL UNIQUE,
            external_id TEXT,
            lastname TEXT NOT NULL,
            firstname TEXT NOT NULL,
            carelevel TEXT,
            phone TEXT,
            mobile TEXT,
            active INTEGER NOT NULL DEFAULT 1,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS contacts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER NOT NULL,
            lastname TEXT NOT NULL,
            firstname TEXT NOT NULL,
            relation TEXT,
            phone TEXT,
            mobile TEXT,
            phone_digits TEXT,
            mobile_digits TEXT,
            FOREIGN KEY(customer_id) REFERENCES customers(id) ON DELETE CASCADE
        );

        CREATE INDEX IF NOT EXISTS idx_customers_phone_digits
            ON customers ((replace(replace(replace(replace(phone, ' ', ''), '-', ''), '/', ''), '+', '')));

        CREATE INDEX IF NOT EXISTS idx_customers_mobile_digits
            ON customers ((replace(replace(replace(replace(mobile, ' ', ''), '-', ''), '/', ''), '+', '')));

        CREATE INDEX IF NOT EXISTS idx_contacts_phone_digits ON contacts(phone_digits);
        CREATE INDEX IF NOT EXISTS idx_contacts_mobile_digits ON contacts(mobile_digits);
        """
    )


def detect_sheet_kind(path: Path) -> str:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    sample_text = []
    for r in range(1, min(40, ws.max_row) + 1):
        for c in range(1, min(25, ws.max_column) + 1):
            val = clean(ws.cell(r, c).value)
            if val:
                sample_text.append(val)

    blob = " | ".join(sample_text)
    if "Klienten-Nr.:" in blob and "Bezieh.:" in blob:
        return "medifox_report"

    header_tokens = {s.lower() for s in sample_text[:80]}
    if any(t in header_tokens for t in ["name", "nachname"]) and any(
        t in header_tokens for t in ["vorname", "pflegegrad", "telefon", "mobil"]
    ):
        return "tabular_contacts"

    return "unknown"


def is_client_start(ws, row: int) -> bool:
    name = clean(ws.cell(row, 2).value)
    marker = clean(ws.cell(row, 16).value)
    return bool(name and "," in name and marker == "Klienten-Nr.:")


def find_next_client_start(ws, row: int) -> int:
    for rr in range(row + 1, ws.max_row + 1):
        if is_client_start(ws, rr):
            return rr
    return ws.max_row + 1


def extract_report_customer(ws, row: int) -> CustomerRecord:
    lastname, firstname = split_name(clean(ws.cell(row, 2).value))
    phone = normalize_phone(ws.cell(row, 9).value)
    external_id = clean(ws.cell(row, 19).value)
    mobile = ""
    carelevel = ""

    for rr in range(row, min(row + 8, ws.max_row + 1)):
        labels = {clean(ws.cell(rr, c).value): c for c in range(1, ws.max_column + 1) if clean(ws.cell(rr, c).value)}
        if "Mobil:" in labels:
            col = labels["Mobil:"]
            if col + 2 <= ws.max_column:
                mobile = mobile or normalize_phone(ws.cell(rr, col + 2).value)
            if not mobile and col + 1 <= ws.max_column:
                mobile = normalize_phone(ws.cell(rr, col + 1).value)
        for key in ("Pflegegrad:", "Pflegegrad", "PG:"):
            if key in labels:
                col = labels[key]
                carelevel = (
                    carelevel
                    or clean(ws.cell(rr, min(col + 3, ws.max_column)).value)
                    or clean(ws.cell(rr, min(col + 1, ws.max_column)).value)
                )

    return CustomerRecord(
        lastname=lastname,
        firstname=firstname,
        carelevel=carelevel,
        phone=phone,
        mobile=mobile,
        external_id=external_id,
    )


def extract_report_contact(ws, row: int) -> ContactRecord:
    lastname, firstname = split_name(clean(ws.cell(row, 3).value))
    relation = clean(ws.cell(row, 7).value)
    phone = normalize_phone(ws.cell(row, 12).value)
    mobile = ""
    for rr in range(row, min(row + 3, ws.max_row + 1)):
        for c in range(1, ws.max_column + 1):
            label = clean(ws.cell(rr, c).value)
            if label == "Mobil:":
                if c + 1 <= ws.max_column:
                    mobile = mobile or normalize_phone(ws.cell(rr, c + 1).value)
                if c + 2 <= ws.max_column:
                    mobile = mobile or normalize_phone(ws.cell(rr, c + 2).value)
    return ContactRecord(lastname=lastname, firstname=firstname, relation=relation, phone=phone, mobile=mobile)


def parse_medifox_report(path: Path) -> list[CustomerRecord]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    customers: list[CustomerRecord] = []
    for start_row in range(1, ws.max_row + 1):
        if not is_client_start(ws, start_row):
            continue
        customer = extract_report_customer(ws, start_row)
        block_end = find_next_client_start(ws, start_row) - 1
        for rr in range(start_row + 1, block_end + 1):
            if clean(ws.cell(rr, 6).value) == "Bezieh.:":
                customer.contacts.append(extract_report_contact(ws, rr))
        customers.append(customer)
    return customers


def find_header_row(ws) -> int | None:
    search_for = {"name", "nachname", "vorname", "telefon", "mobil", "pflegegrad"}
    for r in range(1, min(ws.max_row, 50) + 1):
        row_tokens = {clean(ws.cell(r, c).value).lower() for c in range(1, ws.max_column + 1)}
        if len(search_for.intersection(row_tokens)) >= 2:
            return r
    return None


def parse_tabular(path: Path) -> list[CustomerRecord]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    header_row = find_header_row(ws)
    if header_row is None:
        raise ValueError("Could not find header row for tabular sheet")

    headers = {clean(ws.cell(header_row, c).value).lower(): c for c in range(1, ws.max_column + 1)}

    def col(*names: str) -> int | None:
        for n in names:
            if n in headers:
                return headers[n]
        return None

    c_last = col("nachname", "name")
    c_first = col("vorname")
    c_full = col("kunde", "klient")
    c_phone = col("telefon", "tel", "phone")
    c_mobile = col("mobil", "handy", "mobile")
    c_care = col("pflegegrad", "pg")
    c_ext = col("klienten-nr.", "kundennummer", "kunden-nr", "id")
    cc_name = col("kontakt_name", "kontakt", "bezugsperson")
    cc_first = col("kontakt_vorname")
    cc_rel = col("kontakt_relation", "beziehung", "relation")
    cc_phone = col("kontakt_telefon", "kontakt_phone")
    cc_mobile = col("kontakt_mobil", "kontakt_mobile")

    if not any([c_last, c_first, c_full]):
        raise ValueError("No customer name columns found in tabular sheet")

    customers: list[CustomerRecord] = []
    by_key: dict[str, CustomerRecord] = {}
    for r in range(header_row + 1, ws.max_row + 1):
        last = clean(ws.cell(r, c_last).value) if c_last else ""
        first = clean(ws.cell(r, c_first).value) if c_first else ""

        if c_full and not (last or first):
            full_last, full_first = split_name(clean(ws.cell(r, c_full).value))
            last = last or full_last
            first = first or full_first

        if not (last or first):
            continue

        rec = CustomerRecord(
            lastname=last,
            firstname=first,
            carelevel=clean(ws.cell(r, c_care).value) if c_care else "",
            phone=normalize_phone(ws.cell(r, c_phone).value) if c_phone else "",
            mobile=normalize_phone(ws.cell(r, c_mobile).value) if c_mobile else "",
            external_id=clean(ws.cell(r, c_ext).value) if c_ext else "",
        )

        key = compute_customer_key(rec)
        if key not in by_key:
            by_key[key] = rec
            customers.append(rec)

        target = by_key[key]
        contact_last = clean(ws.cell(r, cc_name).value) if cc_name else ""
        contact_first = clean(ws.cell(r, cc_first).value) if cc_first else ""
        if contact_last or contact_first:
            target.contacts.append(
                ContactRecord(
                    lastname=contact_last,
                    firstname=contact_first,
                    relation=clean(ws.cell(r, cc_rel).value) if cc_rel else "",
                    phone=normalize_phone(ws.cell(r, cc_phone).value) if cc_phone else "",
                    mobile=normalize_phone(ws.cell(r, cc_mobile).value) if cc_mobile else "",
                )
            )

    return customers


def parse_excel(path: Path) -> tuple[str, list[CustomerRecord]]:
    kind = detect_sheet_kind(path)
    if kind == "medifox_report":
        return kind, parse_medifox_report(path)
    if kind == "tabular_contacts":
        return kind, parse_tabular(path)
    raise ValueError(f"Unknown sheet layout in {path.name}")


def upsert_customer(conn: sqlite3.Connection, rec: CustomerRecord) -> int:
    customer_key = compute_customer_key(rec)
    conn.execute(
        """
        INSERT INTO customers (customer_key, external_id, lastname, firstname, carelevel, phone, mobile, active)
        VALUES (?, ?, ?, ?, ?, ?, ?, 1)
        ON CONFLICT(customer_key) DO UPDATE SET
            external_id = excluded.external_id,
            lastname = excluded.lastname,
            firstname = excluded.firstname,
            carelevel = excluded.carelevel,
            phone = excluded.phone,
            mobile = excluded.mobile,
            active = 1,
            updated_at = CURRENT_TIMESTAMP
        """,
        (
            customer_key,
            rec.external_id,
            rec.lastname,
            rec.firstname,
            rec.carelevel,
            rec.phone,
            rec.mobile,
        ),
    )
    row = conn.execute("SELECT id FROM customers WHERE customer_key = ?", (customer_key,)).fetchone()
    return int(row[0])


def replace_contacts(conn: sqlite3.Connection, customer_id: int, contacts: list[ContactRecord]) -> None:
    conn.execute("DELETE FROM contacts WHERE customer_id = ?", (customer_id,))
    for c in contacts:
        conn.execute(
            """
            INSERT INTO contacts (customer_id, lastname, firstname, relation, phone, mobile, phone_digits, mobile_digits)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                customer_id,
                c.lastname,
                c.firstname,
                c.relation,
                c.phone,
                c.mobile,
                normalized_phone_digits(c.phone),
                normalized_phone_digits(c.mobile),
            ),
        )


def sync_customers(conn: sqlite3.Connection, imported: list[CustomerRecord]) -> tuple[int, int]:
    if not imported:
        return 0, 0

    seen_ids: set[int] = set()
    with conn:
        for rec in imported:
            customer_id = upsert_customer(conn, rec)
            seen_ids.add(customer_id)
            replace_contacts(conn, customer_id, rec.contacts)

        placeholders = ",".join("?" for _ in seen_ids)
        conn.execute(f"UPDATE customers SET active = 0 WHERE id NOT IN ({placeholders})", tuple(seen_ids))
        conn.execute(f"UPDATE customers SET active = 1 WHERE id IN ({placeholders})", tuple(seen_ids))

    total = conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0]
    active = conn.execute("SELECT COUNT(*) FROM customers WHERE active = 1").fetchone()[0]
    return int(total), int(active)


def import_excel_flow(conn: sqlite3.Connection) -> None:
    raw = input("Excel path(s), separated by comma: ").strip()
    if not raw:
        print("No files provided.")
        return

    files = [Path(p.strip().strip('"')) for p in raw.split(",") if p.strip()]
    all_records: list[CustomerRecord] = []

    for path in files:
        if not path.exists() or not path.is_file():
            print(f"Skipping missing file: {path}")
            continue
        try:
            kind, records = parse_excel(path)
            print(f"Detected {kind} for {path.name}: {len(records)} customer(s)")
            all_records.extend(records)
        except Exception as exc:
            print(f"Failed to parse {path}: {exc}")

    if not all_records:
        print("Nothing imported.")
        return

    total, active = sync_customers(conn, all_records)
    print(f"Import done. Total in DB: {total}, active customers: {active}")


def export_active_csv(conn: sqlite3.Connection) -> None:
    default_path = Path("active_telephone_list.csv")
    raw = input(f"Output CSV path [{default_path}]: ").strip()
    out_path = Path(raw) if raw else default_path

    rows = conn.execute(
        """
        SELECT
            c.lastname,
            c.firstname,
            c.carelevel,
            c.phone,
            c.mobile,
            p.lastname,
            p.firstname,
            p.relation,
            p.phone,
            p.mobile
        FROM customers c
        LEFT JOIN contacts p ON p.customer_id = c.id
        WHERE c.active = 1
        ORDER BY c.lastname, c.firstname, p.lastname, p.firstname
        """
    ).fetchall()

    header = [
        "name",
        "firstname",
        "pflegegrad",
        "phone",
        "mobile",
        "contact_name",
        "contact_firstname",
        "contact_relation",
        "contact_phone",
        "contact_mobile",
    ]

    with out_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(rows)

    print(f"CSV exported: {out_path} ({len(rows)} row(s))")


def search_by_phone(conn: sqlite3.Connection) -> None:
    needle = normalized_phone_digits(input("Phone number to search: "))
    if not needle:
        print("No phone number entered.")
        return

    results = conn.execute(
        """
        SELECT
            c.lastname,
            c.firstname,
            c.phone,
            c.mobile,
            c.active,
            NULL AS contact_lastname,
            NULL AS contact_firstname,
            NULL AS relation,
            NULL AS contact_phone,
            NULL AS contact_mobile,
            'customer' AS source
        FROM customers c
        WHERE replace(replace(replace(replace(c.phone, ' ', ''), '-', ''), '/', ''), '+', '') LIKE '%' || ? || '%'
           OR replace(replace(replace(replace(c.mobile, ' ', ''), '-', ''), '/', ''), '+', '') LIKE '%' || ? || '%'

        UNION ALL

        SELECT
            c.lastname,
            c.firstname,
            c.phone,
            c.mobile,
            c.active,
            p.lastname,
            p.firstname,
            p.relation,
            p.phone,
            p.mobile,
            'contact' AS source
        FROM contacts p
        JOIN customers c ON c.id = p.customer_id
        WHERE p.phone_digits LIKE '%' || ? || '%'
           OR p.mobile_digits LIKE '%' || ? || '%'
        ORDER BY lastname, firstname
        """,
        (needle, needle, needle, needle),
    ).fetchall()

    if not results:
        print("No matches found.")
        return

    for row in results:
        cust_last, cust_first, cust_phone, cust_mobile, active, c_last, c_first, relation, c_phone, c_mobile, source = row
        status = "active" if active else "inactive"
        if source == "customer":
            print(f"[{status}] Customer: {cust_last}, {cust_first} | Tel: {cust_phone} | Mobil: {cust_mobile}")
        else:
            print(
                f"[{status}] Contact: {c_last}, {c_first} ({relation}) | Tel: {c_phone} | Mobil: {c_mobile} "
                f"for {cust_last}, {cust_first}"
            )


def show_stats(conn: sqlite3.Connection) -> None:
    total = conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0]
    active = conn.execute("SELECT COUNT(*) FROM customers WHERE active = 1").fetchone()[0]
    contacts = conn.execute("SELECT COUNT(*) FROM contacts").fetchone()[0]
    print(f"Customers: {total} total, {active} active | Contacts: {contacts}")


def main() -> None:
    conn = sqlite3.connect(DB_PATH)
    init_db(conn)

    while True:
        print("\n=== Medifox Phonebook Menu ===")
        print("1) Import Excel sheet(s)")
        print("2) Search by phone number")
        print("3) Export active telephone list CSV")
        print("4) Show DB stats")
        print("0) Exit")

        choice = input("Select option: ").strip()
        if choice == "1":
            import_excel_flow(conn)
        elif choice == "2":
            search_by_phone(conn)
        elif choice == "3":
            export_active_csv(conn)
        elif choice == "4":
            show_stats(conn)
        elif choice == "0":
            print("Bye.")
            break
        else:
            print("Invalid choice.")

    conn.close()


if __name__ == "__main__":
    main()
