import sqlite3
import tempfile
import types
import unittest
from pathlib import Path
import sys


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))
sys.modules.setdefault("openpyxl", types.ModuleType("openpyxl"))

import main  # noqa: E402


class HubAwarePhonebookTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tempdir = tempfile.TemporaryDirectory()
        self.original_uploads_dir = main.UPLOADS_DIR
        main.UPLOADS_DIR = Path(self.tempdir.name) / "uploads"

    def tearDown(self) -> None:
        main.UPLOADS_DIR = self.original_uploads_dir
        self.tempdir.cleanup()

    def test_init_db_migrates_legacy_customers_to_dorsten_hub(self) -> None:
        conn = sqlite3.connect(":memory:")
        conn.executescript(
            """
            CREATE TABLE customers (
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

            CREATE TABLE contacts (
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
            """
        )
        conn.execute(
            """
            INSERT INTO customers (
                customer_key,
                external_id,
                lastname,
                firstname,
                carelevel,
                phone,
                mobile,
                active
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            ("legacy-key", "42", "Muster", "Mina", "PG 2", "0201 123", "0170 999", 1),
        )
        conn.execute(
            """
            INSERT INTO contacts (
                customer_id,
                lastname,
                firstname,
                relation,
                phone,
                mobile,
                phone_digits,
                mobile_digits
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (1, "Kontakt", "Karl", "Sohn", "0201 444", "0171 555", "0201444", "0171555"),
        )

        main.init_db(conn)

        hubs = main.list_hubs(conn)
        self.assertEqual(["Dorsten"], [hub.name for hub in hubs])

        customer_row = conn.execute(
            """
            SELECT h.name, c.lastname, c.firstname, c.active
            FROM customers c
            JOIN hubs h ON h.id = c.hub_id
            """
        ).fetchone()
        self.assertEqual(("Dorsten", "Muster", "Mina", 1), customer_row)

        contact_count = conn.execute("SELECT COUNT(*) FROM contacts").fetchone()[0]
        self.assertEqual(1, contact_count)

    def test_sync_customers_only_deactivates_selected_hub(self) -> None:
        conn = sqlite3.connect(":memory:")
        main.init_db(conn)

        dorsten = main.ensure_hub(conn, "Dorsten")
        essen = main.ensure_hub(conn, "Essen")

        dorsten_first = main.CustomerRecord(lastname="Acker", firstname="Anna", phone="0201 111")
        dorsten_second = main.CustomerRecord(lastname="Becker", firstname="Berta", phone="0201 222")
        essen_customer = main.CustomerRecord(lastname="Cramer", firstname="Claus", phone="0201 333")

        main.sync_customers(conn, dorsten.id, [dorsten_first, dorsten_second])
        main.sync_customers(conn, essen.id, [essen_customer])
        main.sync_customers(conn, dorsten.id, [dorsten_first])

        rows = conn.execute(
            """
            SELECT h.name, c.lastname, c.active
            FROM customers c
            JOIN hubs h ON h.id = c.hub_id
            ORDER BY h.name, c.lastname
            """
        ).fetchall()
        self.assertEqual(
            [
                ("Dorsten", "Acker", 1),
                ("Dorsten", "Becker", 0),
                ("Essen", "Cramer", 1),
            ],
            rows,
        )

        dorsten_rows = main.fetch_export_rows(conn, [dorsten.id])
        combined_rows = main.fetch_export_rows(conn, [dorsten.id, essen.id])

        self.assertEqual([("Dorsten", "Acker"),], [(row[0], row[1]) for row in dorsten_rows])
        self.assertEqual(
            [("Dorsten", "Acker"), ("Essen", "Cramer")],
            [(row[0], row[1]) for row in combined_rows],
        )


if __name__ == "__main__":
    unittest.main()
