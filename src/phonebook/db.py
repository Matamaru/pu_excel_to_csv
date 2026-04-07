"""Database and local storage helpers for the phonebook app."""

from __future__ import annotations

import re
import shutil
import sqlite3
from pathlib import Path

from phonebook.models import (
    ContactRecord,
    CustomerRecord,
    HubRecord,
    clean,
    compute_customer_key,
    normalized_phone_digits,
)


DB_PATH = Path("phonebook.db")
UPLOADS_DIR = Path("uploads")
EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls"}
LEGACY_HUB_NAME = "Dorsten"


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
        [path for path in UPLOADS_DIR.iterdir() if is_excel_file(path)],
        key=lambda path: path.name.lower(),
    )


def list_upload_excel_files(hub: HubRecord) -> list[Path]:
    files: list[Path] = []
    hub_dir = ensure_hub_uploads_dir(hub)
    files.extend(path for path in hub_dir.iterdir() if is_excel_file(path))
    if hub.name.lower() == LEGACY_HUB_NAME.lower():
        files.extend(list_legacy_upload_excel_files())

    seen: set[Path] = set()
    unique: list[Path] = []
    for path in sorted(files, key=lambda candidate: (candidate.parent.name.lower(), candidate.name.lower())):
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
    for token in [item.strip().strip('"') for item in raw.split(",") if item.strip()]:
        path = Path(token)
        if path.is_dir():
            for child in sorted(path.iterdir(), key=lambda candidate: candidate.name.lower()):
                if is_excel_file(child):
                    result.append(child)
            continue
        if is_excel_file(path):
            result.append(path)

    seen: set[Path] = set()
    unique: list[Path] = []
    for path in result:
        resolved = path.resolve()
        if resolved in seen:
            continue
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


def slugify_hub_name(value: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", clean(value).lower()).strip("-")
    return slug or "hub"


def list_hubs(conn: sqlite3.Connection) -> list[HubRecord]:
    rows = conn.execute(
        "SELECT id, name, slug FROM hubs ORDER BY lower(name), id"
    ).fetchall()
    return [HubRecord(id=int(row[0]), name=row[1], slug=row[2]) for row in rows]


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
    for contact in contacts:
        conn.execute(
            """
            INSERT INTO contacts (customer_id, lastname, firstname, relation, phone, mobile, phone_digits, mobile_digits)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                customer_id,
                contact.lastname,
                contact.firstname,
                contact.relation,
                contact.phone,
                contact.mobile,
                normalized_phone_digits(contact.phone),
                normalized_phone_digits(contact.mobile),
            ),
        )


def sync_customers(conn: sqlite3.Connection, hub_id: int, imported: list[CustomerRecord]) -> tuple[int, int]:
    if not imported:
        return 0, 0

    seen_ids: set[int] = set()
    with conn:
        for record in imported:
            customer_id = upsert_customer(conn, hub_id, record)
            seen_ids.add(customer_id)
            replace_contacts(conn, customer_id, record.contacts)

        placeholders = ",".join("?" for _ in seen_ids)
        conn.execute(
            f"UPDATE customers SET active = 0 WHERE hub_id = ? AND id NOT IN ({placeholders})",
            (hub_id, *tuple(seen_ids)),
        )
        conn.execute(
            f"UPDATE customers SET active = 1 WHERE hub_id = ? AND id IN ({placeholders})",
            (hub_id, *tuple(seen_ids)),
        )

    return get_total_and_active_counts(conn)


def get_total_and_active_counts(conn: sqlite3.Connection) -> tuple[int, int]:
    total = int(conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0])
    active = int(conn.execute("SELECT COUNT(*) FROM customers WHERE active = 1").fetchone()[0])
    return total, active


def get_hub_counts(conn: sqlite3.Connection, hub_id: int) -> tuple[int, int]:
    total = int(conn.execute("SELECT COUNT(*) FROM customers WHERE hub_id = ?", (hub_id,)).fetchone()[0])
    active = int(
        conn.execute(
            "SELECT COUNT(*) FROM customers WHERE hub_id = ? AND active = 1",
            (hub_id,),
        ).fetchone()[0]
    )
    return total, active


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
    total, active = get_total_and_active_counts(conn)
    contacts = int(conn.execute("SELECT COUNT(*) FROM contacts").fetchone()[0])
    lines = [
        f"Customers: {total} total",
        f"Customers active: {active}",
        f"Contacts: {contacts}",
    ]

    rows = hub_stats_rows(conn)
    if rows:
        lines.append("")
        lines.append("Hubs:")
        for hub_name, hub_total, hub_active in rows:
            lines.append(f"- {hub_name}: {hub_active}/{hub_total} active")
    return "\n".join(lines)
