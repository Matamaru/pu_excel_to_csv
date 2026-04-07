#!/usr/bin/env python3
"""Terminal tool to import Medifox Excel exports into SQLite and export phone list CSV."""

from __future__ import annotations

import csv
import hashlib
import os
import shutil
import sqlite3
import sys
from dataclasses import dataclass, field
from pathlib import Path
import re

import openpyxl

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog
except Exception:
    tk = None
    filedialog = None
    messagebox = None
    simpledialog = None


NBSP = "\xa0"
DB_PATH = Path("phonebook.db")
UPLOADS_DIR = Path("uploads")
EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls"}
LEGACY_HUB_NAME = "Dorsten"


@dataclass(frozen=True)
class HubRecord:
    id: int
    name: str
    slug: str


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


def create_hubs_table(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS hubs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            slug TEXT NOT NULL UNIQUE,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )


def create_customers_table(conn: sqlite3.Connection, table_name: str = "customers") -> None:
    conn.execute(
        f"""
        CREATE TABLE IF NOT EXISTS {table_name} (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            hub_id INTEGER NOT NULL,
            customer_key TEXT NOT NULL,
            external_id TEXT,
            lastname TEXT NOT NULL,
            firstname TEXT NOT NULL,
            carelevel TEXT,
            phone TEXT,
            mobile TEXT,
            active INTEGER NOT NULL DEFAULT 1,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(hub_id) REFERENCES hubs(id) ON DELETE RESTRICT,
            UNIQUE(hub_id, customer_key)
        )
        """
    )


def create_contacts_table(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
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
        )
        """
    )


def ensure_indexes(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        CREATE INDEX IF NOT EXISTS idx_customers_hub_id ON customers(hub_id);
        CREATE INDEX IF NOT EXISTS idx_customers_hub_active ON customers(hub_id, active);

        CREATE INDEX IF NOT EXISTS idx_customers_phone_digits
            ON customers ((replace(replace(replace(replace(phone, ' ', ''), '-', ''), '/', ''), '+', '')));

        CREATE INDEX IF NOT EXISTS idx_customers_mobile_digits
            ON customers ((replace(replace(replace(replace(mobile, ' ', ''), '-', ''), '/', ''), '+', '')));

        CREATE INDEX IF NOT EXISTS idx_contacts_phone_digits ON contacts(phone_digits);
        CREATE INDEX IF NOT EXISTS idx_contacts_mobile_digits ON contacts(mobile_digits);
        """
    )


def table_exists(conn: sqlite3.Connection, name: str) -> bool:
    row = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
        (name,),
    ).fetchone()
    return row is not None


def table_has_column(conn: sqlite3.Connection, table_name: str, column_name: str) -> bool:
    if not table_exists(conn, table_name):
        return False
    rows = conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    return any(row[1] == column_name for row in rows)


def slugify_hub_name(value: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", clean(value).lower()).strip("-")
    return slug or "hub"


def list_hubs(conn: sqlite3.Connection) -> list[HubRecord]:
    rows = conn.execute(
        "SELECT id, name, slug FROM hubs ORDER BY lower(name), id"
    ).fetchall()
    return [HubRecord(id=int(row[0]), name=row[1], slug=row[2]) for row in rows]


def get_hub_by_id(conn: sqlite3.Connection, hub_id: int) -> HubRecord | None:
    row = conn.execute(
        "SELECT id, name, slug FROM hubs WHERE id = ?",
        (hub_id,),
    ).fetchone()
    if row is None:
        return None
    return HubRecord(id=int(row[0]), name=row[1], slug=row[2])


def ensure_hub(conn: sqlite3.Connection, name: str) -> HubRecord:
    hub_name = clean(name)
    if not hub_name:
        raise ValueError("Hub name cannot be empty.")

    existing = conn.execute(
        "SELECT id, name, slug FROM hubs WHERE lower(name) = lower(?)",
        (hub_name,),
    ).fetchone()
    if existing is not None:
        hub = HubRecord(id=int(existing[0]), name=existing[1], slug=existing[2])
        ensure_hub_uploads_dir(hub)
        return hub

    base_slug = slugify_hub_name(hub_name)
    slug = base_slug
    suffix = 2
    while conn.execute("SELECT 1 FROM hubs WHERE slug = ?", (slug,)).fetchone() is not None:
        slug = f"{base_slug}-{suffix}"
        suffix += 1

    with conn:
        conn.execute(
            "INSERT INTO hubs (name, slug) VALUES (?, ?)",
            (hub_name, slug),
        )
        hub_id = int(conn.execute("SELECT last_insert_rowid()").fetchone()[0])

    hub = HubRecord(id=hub_id, name=hub_name, slug=slug)
    ensure_hub_uploads_dir(hub)
    return hub


def migrate_customers_to_hubs(conn: sqlite3.Connection) -> None:
    legacy_hub = ensure_hub(conn, LEGACY_HUB_NAME)
    conn.commit()
    conn.execute("PRAGMA foreign_keys = OFF")
    try:
        with conn:
            create_customers_table(conn, "customers_new")
            conn.execute(
                """
                INSERT INTO customers_new (
                    id,
                    hub_id,
                    customer_key,
                    external_id,
                    lastname,
                    firstname,
                    carelevel,
                    phone,
                    mobile,
                    active,
                    updated_at
                )
                SELECT
                    id,
                    ?,
                    customer_key,
                    external_id,
                    lastname,
                    firstname,
                    carelevel,
                    phone,
                    mobile,
                    active,
                    updated_at
                FROM customers
                """,
                (legacy_hub.id,),
            )
            conn.execute("DROP TABLE customers")
            conn.execute("ALTER TABLE customers_new RENAME TO customers")
        ensure_indexes(conn)
    finally:
        conn.execute("PRAGMA foreign_keys = ON")


def init_db(conn: sqlite3.Connection) -> None:
    conn.execute("PRAGMA foreign_keys = ON")
    create_hubs_table(conn)

    if not table_exists(conn, "customers"):
        create_customers_table(conn)
    elif not table_has_column(conn, "customers", "hub_id"):
        migrate_customers_to_hubs(conn)

    create_contacts_table(conn)
    ensure_indexes(conn)


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


def upsert_customer(conn: sqlite3.Connection, hub_id: int, rec: CustomerRecord) -> int:
    customer_key = compute_customer_key(rec)
    conn.execute(
        """
        INSERT INTO customers (hub_id, customer_key, external_id, lastname, firstname, carelevel, phone, mobile, active)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)
        ON CONFLICT(hub_id, customer_key) DO UPDATE SET
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
            hub_id,
            customer_key,
            rec.external_id,
            rec.lastname,
            rec.firstname,
            rec.carelevel,
            rec.phone,
            rec.mobile,
        ),
    )
    row = conn.execute(
        "SELECT id FROM customers WHERE hub_id = ? AND customer_key = ?",
        (hub_id, customer_key),
    ).fetchone()
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


def sync_customers(conn: sqlite3.Connection, hub_id: int, imported: list[CustomerRecord]) -> tuple[int, int]:
    if not imported:
        return 0, 0

    seen_ids: set[int] = set()
    with conn:
        for rec in imported:
            customer_id = upsert_customer(conn, hub_id, rec)
            seen_ids.add(customer_id)
            replace_contacts(conn, customer_id, rec.contacts)

        placeholders = ",".join("?" for _ in seen_ids)
        conn.execute(
            f"UPDATE customers SET active = 0 WHERE hub_id = ? AND id NOT IN ({placeholders})",
            (hub_id, *tuple(seen_ids)),
        )
        conn.execute(
            f"UPDATE customers SET active = 1 WHERE hub_id = ? AND id IN ({placeholders})",
            (hub_id, *tuple(seen_ids)),
        )

    total = conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0]
    active = conn.execute("SELECT COUNT(*) FROM customers WHERE active = 1").fetchone()[0]
    return int(total), int(active)


def clear_screen() -> None:
    os.system("cls" if os.name == "nt" else "clear")


def pause() -> None:
    input("\nPress Enter to continue...")


def ensure_uploads_dir() -> None:
    UPLOADS_DIR.mkdir(parents=True, exist_ok=True)


def hub_uploads_dir(hub: HubRecord) -> Path:
    return UPLOADS_DIR / hub.slug


def ensure_hub_uploads_dir(hub: HubRecord) -> Path:
    ensure_uploads_dir()
    path = hub_uploads_dir(hub)
    path.mkdir(parents=True, exist_ok=True)
    return path


def is_excel_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in EXCEL_SUFFIXES


def list_legacy_upload_excel_files() -> list[Path]:
    ensure_uploads_dir()
    return sorted(
        [p for p in UPLOADS_DIR.iterdir() if is_excel_file(p)],
        key=lambda p: p.name.lower(),
    )


def list_upload_excel_files(hub: HubRecord) -> list[Path]:
    files: list[Path] = []
    hub_dir = ensure_hub_uploads_dir(hub)
    files.extend(p for p in hub_dir.iterdir() if is_excel_file(p))
    if hub.name.lower() == LEGACY_HUB_NAME.lower():
        files.extend(list_legacy_upload_excel_files())

    seen: set[Path] = set()
    unique: list[Path] = []
    for path in sorted(files, key=lambda p: (p.parent.name.lower(), p.name.lower())):
        resolved = path.resolve()
        if resolved in seen:
            continue
        seen.add(resolved)
        unique.append(path)
    return unique


def parse_paths_input(raw: str) -> list[Path]:
    if not raw.strip():
        return []
    result: list[Path] = []
    for token in [t.strip().strip('"') for t in raw.split(",") if t.strip()]:
        path = Path(token)
        if path.is_dir():
            for child in sorted(path.iterdir(), key=lambda p: p.name.lower()):
                if is_excel_file(child):
                    result.append(child)
            continue
        if is_excel_file(path):
            result.append(path)
    # De-duplicate while keeping order.
    seen: set[Path] = set()
    unique: list[Path] = []
    for path in result:
        resolved = path.resolve()
        if resolved not in seen:
            seen.add(resolved)
            unique.append(path)
    return unique


def copy_files_into_uploads(files: list[Path], hub: HubRecord) -> list[Path]:
    dest_dir = ensure_hub_uploads_dir(hub)
    copied: list[Path] = []
    for source in files:
        dest = dest_dir / source.name
        if source.resolve() != dest.resolve():
            shutil.copy2(source, dest)
        copied.append(dest)
    return copied


def describe_upload_file(path: Path) -> str:
    if path.parent == UPLOADS_DIR:
        return f"{path.name} [legacy root]"
    return path.name


def choose_upload_files(hub: HubRecord) -> list[Path]:
    files = list_upload_excel_files(hub)
    if not files:
        print(f"No Excel files for hub '{hub.name}' in {ensure_hub_uploads_dir(hub)}.")
        return []

    print("Select file numbers separated by comma, or 'all'.")
    for idx, file_path in enumerate(files, 1):
        print(f"{idx}) {describe_upload_file(file_path)}")

    selection = input("Selection: ").strip().lower()
    if selection == "all":
        return files

    chosen: list[Path] = []
    for token in [t.strip() for t in selection.split(",") if t.strip()]:
        if not token.isdigit():
            continue
        index = int(token)
        if 1 <= index <= len(files):
            chosen.append(files[index - 1])
    return chosen


def prompt_create_hub_cli(conn: sqlite3.Connection) -> HubRecord | None:
    raw_name = input("New hub name: ").strip()
    if not raw_name:
        print("Hub creation cancelled.")
        return None
    try:
        hub = ensure_hub(conn, raw_name)
    except ValueError as exc:
        print(exc)
        return None
    print(f"Hub ready: {hub.name}")
    return hub


def choose_hub_cli(conn: sqlite3.Connection, *, allow_create: bool = True) -> HubRecord | None:
    while True:
        hubs = list_hubs(conn)
        if not hubs:
            print("No hubs yet. Create the first hub now.")
            if not allow_create:
                return None
            return prompt_create_hub_cli(conn)

        print("Select a hub:")
        for idx, hub in enumerate(hubs, 1):
            print(f"{idx}) {hub.name}")
        if allow_create:
            print("n) Create new hub")
        print("0) Back")

        selection = input("Hub: ").strip().lower()
        if selection == "0":
            return None
        if allow_create and selection == "n":
            hub = prompt_create_hub_cli(conn)
            if hub is not None:
                return hub
            continue
        if selection.isdigit():
            index = int(selection)
            if 1 <= index <= len(hubs):
                return hubs[index - 1]
        print("Invalid choice.")


def choose_hubs_cli(conn: sqlite3.Connection) -> list[HubRecord]:
    hubs = list_hubs(conn)
    if not hubs:
        print("No hubs available.")
        return []

    print("Select hub numbers separated by comma, or 'all'.")
    for idx, hub in enumerate(hubs, 1):
        print(f"{idx}) {hub.name}")

    selection = input("Hub selection [all]: ").strip().lower()
    if not selection or selection == "all":
        return hubs

    selected: list[HubRecord] = []
    seen_ids: set[int] = set()
    for token in [t.strip() for t in selection.split(",") if t.strip()]:
        if not token.isdigit():
            continue
        index = int(token)
        if 1 <= index <= len(hubs):
            hub = hubs[index - 1]
            if hub.id not in seen_ids:
                seen_ids.add(hub.id)
                selected.append(hub)
    return selected


def run_import(conn: sqlite3.Connection, hub: HubRecord, files: list[Path]) -> dict:
    summary = {
        "hub": hub.name,
        "files_ok": 0,
        "files_failed": 0,
        "customers_parsed": 0,
        "total": 0,
        "active": 0,
        "hub_total": 0,
        "hub_active": 0,
        "errors": [],
    }

    if not files:
        print("No valid Excel files selected.")
        return summary

    all_records: list[CustomerRecord] = []

    for path in files:
        if not path.exists() or not path.is_file():
            print(f"Skipping missing file: {path}")
            summary["files_failed"] += 1
            summary["errors"].append(f"Missing file: {path}")
            continue
        try:
            kind, records = parse_excel(path)
            print(f"Detected {kind} for {path.name}: {len(records)} customer(s)")
            all_records.extend(records)
            summary["files_ok"] += 1
            summary["customers_parsed"] += len(records)
        except Exception as exc:
            print(f"Failed to parse {path}: {exc}")
            summary["files_failed"] += 1
            summary["errors"].append(f"{path.name}: {exc}")

    if not all_records:
        print("Nothing imported.")
        return summary

    total, active = sync_customers(conn, hub.id, all_records)
    hub_total = conn.execute("SELECT COUNT(*) FROM customers WHERE hub_id = ?", (hub.id,)).fetchone()[0]
    hub_active = conn.execute(
        "SELECT COUNT(*) FROM customers WHERE hub_id = ? AND active = 1",
        (hub.id,),
    ).fetchone()[0]
    print(
        f"Import done for hub {hub.name}. "
        f"Hub customers: {hub_active}/{hub_total} active | DB total: {total}, active: {active}"
    )
    summary["total"] = total
    summary["active"] = active
    summary["hub_total"] = int(hub_total)
    summary["hub_active"] = int(hub_active)
    return summary


def import_excel_flow(conn: sqlite3.Connection) -> None:
    ensure_uploads_dir()
    hub = choose_hub_cli(conn)
    if hub is None:
        return

    print("\n=== Import Excel ===")
    print(f"Hub: {hub.name}")
    print(f"Hub uploads folder: {ensure_hub_uploads_dir(hub).resolve()}")
    print("1) Import all Excel files from uploads folder")
    print("2) Select Excel file(s) from uploads folder")
    print("3) Add file(s) to uploads folder and import")
    print("4) Import custom path(s) directly")
    print("0) Back")

    choice = input("Select option: ").strip()
    if choice == "1":
        run_import(conn, hub, list_upload_excel_files(hub))
        return
    if choice == "2":
        run_import(conn, hub, choose_upload_files(hub))
        return
    if choice == "3":
        raw = input("Source file path(s), separated by comma: ").strip()
        source_files = parse_paths_input(raw)
        copied = copy_files_into_uploads(source_files, hub)
        if copied:
            print(f"Copied {len(copied)} file(s) to {ensure_hub_uploads_dir(hub)}.")
        run_import(conn, hub, copied)
        return
    if choice == "4":
        raw = input("Excel path(s) or folder path(s), separated by comma: ").strip()
        run_import(conn, hub, parse_paths_input(raw))
        return
    if choice == "0":
        return
    print("Invalid choice.")


def summary_text(summary: dict) -> str:
    lines = [
        f"Hub: {summary['hub']}",
        f"Files processed: {summary['files_ok']}",
        f"Files failed: {summary['files_failed']}",
        f"Customers parsed: {summary['customers_parsed']}",
    ]
    if summary["total"]:
        lines.append(f"Hub customers total: {summary['hub_total']}")
        lines.append(f"Hub customers active: {summary['hub_active']}")
        lines.append(f"DB customers total: {summary['total']}")
        lines.append(f"DB customers active: {summary['active']}")
    if summary["errors"]:
        lines.append("")
        lines.append("Errors:")
        lines.extend(summary["errors"])
    return "\n".join(lines)


def search_matches(conn: sqlite3.Connection, needle: str) -> list[tuple]:
    return conn.execute(
        """
        SELECT
            h.name,
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
        JOIN hubs h ON h.id = c.hub_id
        WHERE replace(replace(replace(replace(c.phone, ' ', ''), '-', ''), '/', ''), '+', '') LIKE '%' || ? || '%'
           OR replace(replace(replace(replace(c.mobile, ' ', ''), '-', ''), '/', ''), '+', '') LIKE '%' || ? || '%'

        UNION ALL

        SELECT
            h.name,
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
        JOIN hubs h ON h.id = c.hub_id
        WHERE p.phone_digits LIKE '%' || ? || '%'
           OR p.mobile_digits LIKE '%' || ? || '%'
        ORDER BY 1, 2, 3
        """,
        (needle, needle, needle, needle),
    ).fetchall()


CSV_HEADER = [
    "hub",
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


def fetch_export_rows(conn: sqlite3.Connection, hub_ids: list[int] | None = None) -> list[tuple]:
    params: list[object] = []
    hub_filter = ""
    if hub_ids:
        placeholders = ",".join("?" for _ in hub_ids)
        hub_filter = f" AND c.hub_id IN ({placeholders})"
        params.extend(hub_ids)

    return conn.execute(
        f"""
        SELECT
            h.name,
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
        JOIN hubs h ON h.id = c.hub_id
        LEFT JOIN contacts p ON p.customer_id = c.id
        WHERE c.active = 1{hub_filter}
        ORDER BY h.name, c.lastname, c.firstname, p.lastname, p.firstname
        """,
        tuple(params),
    ).fetchall()


def write_export_csv(out_path: Path, rows: list[tuple]) -> None:
    with out_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(CSV_HEADER)
        writer.writerows(rows)


def hub_stats_rows(conn: sqlite3.Connection) -> list[tuple]:
    return conn.execute(
        """
        SELECT
            h.name,
            COUNT(c.id) AS total_customers,
            COALESCE(SUM(CASE WHEN c.active = 1 THEN 1 ELSE 0 END), 0) AS active_customers
        FROM hubs h
        LEFT JOIN customers c ON c.hub_id = h.id
        GROUP BY h.id, h.name
        ORDER BY lower(h.name), h.id
        """
    ).fetchall()


def format_stats_text(conn: sqlite3.Connection) -> str:
    total = conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0]
    active = conn.execute("SELECT COUNT(*) FROM customers WHERE active = 1").fetchone()[0]
    contacts = conn.execute("SELECT COUNT(*) FROM contacts").fetchone()[0]
    lines = [
        f"Customers: {total} total",
        f"Customers active: {active}",
        f"Contacts: {contacts}",
    ]

    hub_rows = hub_stats_rows(conn)
    if hub_rows:
        lines.append("")
        lines.append("Hubs:")
        for hub_name, hub_total, hub_active in hub_rows:
            lines.append(f"- {hub_name}: {hub_active}/{hub_total} active")
    return "\n".join(lines)


def default_export_filename(hubs: list[HubRecord]) -> str:
    if len(hubs) == 1:
        return f"active_telephone_list_{hubs[0].slug}.csv"
    return "active_telephone_list.csv"


def run_tkinter_ui(conn: sqlite3.Connection) -> None:
    ensure_uploads_dir()
    root = tk.Tk()
    root.title("Medifox Phonebook")
    root.geometry("560x430")

    frame = tk.Frame(root, padx=16, pady=16)
    frame.pack(fill="both", expand=True)

    title = tk.Label(frame, text="Medifox Phonebook", font=("Segoe UI", 16, "bold"))
    title.pack(pady=(0, 12))

    subtitle = tk.Label(frame, anchor="w", justify="left")
    subtitle.pack(fill="x", pady=(0, 12))

    def refresh_subtitle() -> None:
        hubs = list_hubs(conn)
        if hubs:
            subtitle.config(
                text=(
                    f"Uploads root: {UPLOADS_DIR.resolve()}\n"
                    f"Hubs: {', '.join(hub.name for hub in hubs)}"
                )
            )
            return
        subtitle.config(
            text=f"Uploads root: {UPLOADS_DIR.resolve()}\nNo hubs yet. Create one before importing."
        )

    def create_hub_ui(parent: tk.Misc | None = None, *, show_confirmation: bool = True) -> HubRecord | None:
        raw_name = simpledialog.askstring("Create Hub", "Hub name:", parent=parent or root)
        if raw_name is None:
            return None
        try:
            hub = ensure_hub(conn, raw_name)
        except ValueError as exc:
            messagebox.showerror("Create Hub", str(exc), parent=parent or root)
            return None
        refresh_subtitle()
        if show_confirmation:
            messagebox.showinfo(
                "Create Hub",
                f"Hub ready: {hub.name}\nUploads folder: {ensure_hub_uploads_dir(hub)}",
                parent=parent or root,
            )
        return hub

    def choose_hub_dialog(title_text: str) -> HubRecord | None:
        selected: dict[str, HubRecord | None] = {"hub": None}
        dialog = tk.Toplevel(root)
        dialog.title(title_text)
        dialog.geometry("440x340")
        dialog.transient(root)
        dialog.grab_set()

        info = tk.Label(dialog, text="Select one hub for this action:", anchor="w", justify="left")
        info.pack(fill="x", padx=12, pady=(12, 6))

        listbox = tk.Listbox(dialog, exportselection=False)
        listbox.pack(fill="both", expand=True, padx=12, pady=6)

        current_hubs: list[HubRecord] = []

        def refresh_list(select_hub_id: int | None = None) -> None:
            current_hubs[:] = list_hubs(conn)
            listbox.delete(0, "end")
            for hub in current_hubs:
                listbox.insert("end", f"{hub.name}  ({ensure_hub_uploads_dir(hub).name})")
            if current_hubs:
                target_index = 0
                if select_hub_id is not None:
                    for idx, hub in enumerate(current_hubs):
                        if hub.id == select_hub_id:
                            target_index = idx
                            break
                listbox.selection_set(target_index)
            else:
                info.config(text="No hubs yet. Create one to continue.")

        def use_selected() -> None:
            selection = listbox.curselection()
            if not selection:
                messagebox.showinfo("Select Hub", "Select a hub first.", parent=dialog)
                return
            selected["hub"] = current_hubs[int(selection[0])]
            dialog.destroy()

        def create_new() -> None:
            hub = create_hub_ui(dialog, show_confirmation=False)
            if hub is not None:
                refresh_list(hub.id)

        actions = tk.Frame(dialog)
        actions.pack(fill="x", padx=12, pady=(4, 12))

        tk.Button(actions, text="Use Selected", command=use_selected).pack(side="left")
        tk.Button(actions, text="Create Hub", command=create_new).pack(side="left", padx=(8, 0))
        tk.Button(actions, text="Cancel", command=dialog.destroy).pack(side="right")

        refresh_list()
        dialog.wait_window()
        return selected["hub"]

    def choose_export_hubs_dialog() -> list[HubRecord]:
        hubs = list_hubs(conn)
        if not hubs:
            messagebox.showinfo("Export", "No hubs available.")
            return []

        selected: list[HubRecord] = []
        dialog = tk.Toplevel(root)
        dialog.title("Select Hubs For CSV Export")
        dialog.geometry("440x360")
        dialog.transient(root)
        dialog.grab_set()

        tk.Label(
            dialog,
            text="Choose one or more hubs. All hubs are selected by default.",
            anchor="w",
            justify="left",
        ).pack(fill="x", padx=12, pady=(12, 6))

        body = tk.Frame(dialog)
        body.pack(fill="both", expand=True, padx=12, pady=6)

        variables: list[tuple[HubRecord, tk.BooleanVar]] = []
        for hub in hubs:
            var = tk.BooleanVar(value=True)
            variables.append((hub, var))
            tk.Checkbutton(body, text=hub.name, variable=var, anchor="w", justify="left").pack(fill="x")

        actions = tk.Frame(dialog)
        actions.pack(fill="x", padx=12, pady=(4, 12))

        def select_all() -> None:
            for _, var in variables:
                var.set(True)

        def clear_all() -> None:
            for _, var in variables:
                var.set(False)

        def use_selection() -> None:
            chosen = [hub for hub, var in variables if var.get()]
            if not chosen:
                messagebox.showinfo("Export", "Select at least one hub.", parent=dialog)
                return
            selected.extend(chosen)
            dialog.destroy()

        tk.Button(actions, text="Select All", command=select_all).pack(side="left")
        tk.Button(actions, text="Clear", command=clear_all).pack(side="left", padx=(8, 0))
        tk.Button(actions, text="Export", command=use_selection).pack(side="right")
        tk.Button(actions, text="Cancel", command=dialog.destroy).pack(side="right", padx=(0, 8))

        dialog.wait_window()
        return selected

    def pick_upload_files_dialog(files: list[Path]) -> list[Path]:
        selected: list[Path] = []
        if not files:
            return selected

        dialog = tk.Toplevel(root)
        dialog.title("Select Upload Files")
        dialog.geometry("520x360")
        dialog.transient(root)
        dialog.grab_set()

        info = tk.Label(dialog, text="Select one or more files from uploads:", anchor="w")
        info.pack(fill="x", padx=12, pady=(12, 6))

        listbox = tk.Listbox(dialog, selectmode="extended")
        listbox.pack(fill="both", expand=True, padx=12, pady=6)
        for file_path in files:
            listbox.insert("end", describe_upload_file(file_path))

        actions = tk.Frame(dialog)
        actions.pack(fill="x", padx=12, pady=(4, 12))

        def import_selected() -> None:
            for idx in listbox.curselection():
                selected.append(files[int(idx)])
            dialog.destroy()

        def cancel() -> None:
            dialog.destroy()

        tk.Button(actions, text="Import Selected", command=import_selected).pack(side="left")
        tk.Button(actions, text="Cancel", command=cancel).pack(side="right")

        dialog.wait_window()
        return selected

    def select_hub_for_import(title_text: str) -> HubRecord | None:
        hub = choose_hub_dialog(title_text)
        if hub is None:
            return None
        ensure_hub_uploads_dir(hub)
        return hub

    def import_from_uploads() -> None:
        hub = select_hub_for_import("Select Hub For Import")
        if hub is None:
            return
        summary = run_import(conn, hub, list_upload_excel_files(hub))
        messagebox.showinfo("Import Result", summary_text(summary))

    def import_selected_from_uploads() -> None:
        hub = select_hub_for_import("Select Hub For Upload Selection")
        if hub is None:
            return
        files = list_upload_excel_files(hub)
        if not files:
            messagebox.showinfo("Uploads", f"No Excel files for hub '{hub.name}'.")
            return
        selected = pick_upload_files_dialog(files)
        if not selected:
            return
        summary = run_import(conn, hub, selected)
        messagebox.showinfo("Import Result", summary_text(summary))

    def add_and_import() -> None:
        hub = select_hub_for_import("Select Hub For New Uploads")
        if hub is None:
            return
        selected = filedialog.askopenfilenames(
            title="Select Excel file(s)",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xltx *.xltm *.xls"), ("All files", "*.*")],
        )
        if not selected:
            return
        files = [Path(p) for p in selected if is_excel_file(Path(p))]
        copied = copy_files_into_uploads(files, hub)
        summary = run_import(conn, hub, copied)
        messagebox.showinfo("Import Result", summary_text(summary))

    def search_phone() -> None:
        raw = simpledialog.askstring("Search", "Phone number to search:", parent=root)
        needle = normalized_phone_digits(raw or "")
        if not needle:
            return
        results = search_matches(conn, needle)
        if not results:
            messagebox.showinfo("Search", "No matches found.")
            return
        lines = []
        for row in results:
            (
                hub_name,
                cust_last,
                cust_first,
                cust_phone,
                cust_mobile,
                active,
                c_last,
                c_first,
                relation,
                c_phone,
                c_mobile,
                source,
            ) = row
            status = "active" if active else "inactive"
            if source == "customer":
                lines.append(
                    f"[{hub_name}] [{status}] Customer: {cust_last}, {cust_first} | "
                    f"Tel: {cust_phone} | Mobil: {cust_mobile}"
                )
            else:
                lines.append(
                    f"[{hub_name}] [{status}] Contact: {c_last}, {c_first} ({relation}) | "
                    f"Tel: {c_phone} | Mobil: {c_mobile} for {cust_last}, {cust_first}"
                )
        messagebox.showinfo("Search Results", "\n".join(lines[:60]))

    def export_csv_ui() -> None:
        selected_hubs = choose_export_hubs_dialog()
        if not selected_hubs:
            return
        output = filedialog.asksaveasfilename(
            title="Save active telephone list CSV",
            defaultextension=".csv",
            initialfile=default_export_filename(selected_hubs),
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not output:
            return

        out_path = Path(output)
        rows = fetch_export_rows(conn, [hub.id for hub in selected_hubs])
        write_export_csv(out_path, rows)
        messagebox.showinfo(
            "Export",
            (
                f"CSV exported:\n{out_path}\n\n"
                f"Hubs: {', '.join(hub.name for hub in selected_hubs)}\n"
                f"Rows: {len(rows)}"
            ),
        )

    def stats_ui() -> None:
        messagebox.showinfo("Database Stats", format_stats_text(conn))

    buttons = [
        ("Create Hub", create_hub_ui),
        ("Import All From Hub uploads", import_from_uploads),
        ("Select From Hub uploads And Import", import_selected_from_uploads),
        ("Add File(s) To Hub And Import", add_and_import),
        ("Search By Phone", search_phone),
        ("Export Active CSV By Hubs", export_csv_ui),
        ("Show DB Stats", stats_ui),
        ("Exit", root.destroy),
    ]

    for text, cmd in buttons:
        tk.Button(frame, text=text, command=cmd, width=36, pady=6).pack(pady=3)

    refresh_subtitle()
    root.mainloop()


def export_active_csv(conn: sqlite3.Connection) -> None:
    selected_hubs = choose_hubs_cli(conn)
    if not selected_hubs:
        print("No hubs selected.")
        return

    default_path = Path(default_export_filename(selected_hubs))
    raw = input(f"Output CSV path [{default_path}]: ").strip()
    out_path = Path(raw) if raw else default_path

    rows = fetch_export_rows(conn, [hub.id for hub in selected_hubs])
    write_export_csv(out_path, rows)

    print(
        f"CSV exported: {out_path} ({len(rows)} row(s)) "
        f"for hubs: {', '.join(hub.name for hub in selected_hubs)}"
    )


def search_by_phone(conn: sqlite3.Connection) -> None:
    needle = normalized_phone_digits(input("Phone number to search: "))
    if not needle:
        print("No phone number entered.")
        return

    results = search_matches(conn, needle)

    if not results:
        print("No matches found.")
        return

    for row in results:
        (
            hub_name,
            cust_last,
            cust_first,
            cust_phone,
            cust_mobile,
            active,
            c_last,
            c_first,
            relation,
            c_phone,
            c_mobile,
            source,
        ) = row
        status = "active" if active else "inactive"
        if source == "customer":
            print(f"[{hub_name}] [{status}] Customer: {cust_last}, {cust_first} | Tel: {cust_phone} | Mobil: {cust_mobile}")
        else:
            print(
                f"[{hub_name}] [{status}] Contact: {c_last}, {c_first} ({relation}) | Tel: {c_phone} | Mobil: {c_mobile} "
                f"for {cust_last}, {cust_first}"
            )


def show_stats(conn: sqlite3.Connection) -> None:
    print(format_stats_text(conn))


def main() -> None:
    conn = sqlite3.connect(DB_PATH)
    init_db(conn)
    ensure_uploads_dir()

    if tk is not None:
        try:
            run_tkinter_ui(conn)
            conn.close()
            return
        except Exception as exc:
            print(f"GUI startup failed, falling back to terminal mode: {exc}", file=sys.stderr)

    while True:
        clear_screen()
        print("\n=== Medifox Phonebook Menu ===")
        print("1) Upload / Import Excel sheet(s)")
        print("2) Search by phone number")
        print("3) Export active telephone list CSV")
        print("4) Show DB stats")
        print("5) Create hub")
        print("0) Exit")

        choice = input("Select option: ").strip()
        if choice == "1":
            import_excel_flow(conn)
            pause()
        elif choice == "2":
            search_by_phone(conn)
            pause()
        elif choice == "3":
            export_active_csv(conn)
            pause()
        elif choice == "4":
            show_stats(conn)
            pause()
        elif choice == "5":
            prompt_create_hub_cli(conn)
            pause()
        elif choice == "0":
            print("Bye.")
            break
        else:
            print("Invalid choice.")
            pause()

    conn.close()


if __name__ == "__main__":
    main()
