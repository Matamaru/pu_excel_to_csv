import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from phonebook.parsers import detect_sheet_kind, parse_excel  # noqa: E402


class ParserTests(unittest.TestCase):
    def create_workbook_file(self, filename: str, rows: list[list[object]]) -> Path:
        tempdir = tempfile.TemporaryDirectory()
        self.addCleanup(tempdir.cleanup)
        path = Path(tempdir.name) / filename
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        for row in rows:
            worksheet.append(row)
        workbook.save(path)
        return path

    def test_parse_tabular_groups_contacts_by_customer(self) -> None:
        path = self.create_workbook_file(
            "tabular.xlsx",
            [
                [
                    "Nachname",
                    "Vorname",
                    "Pflegegrad",
                    "Telefon",
                    "Mobil",
                    "Klienten-Nr.",
                    "Kontakt_Name",
                    "Kontakt_Vorname",
                    "Kontakt_Relation",
                    "Kontakt_Telefon",
                    "Kontakt_Mobil",
                ],
                ["Muster", "Mina", "PG 2", "0201 123", "0170 456", "42", "Kontakt", "Karl", "Sohn", "0201 555", ""],
                ["Muster", "Mina", "PG 2", "0201 123", "0170 456", "42", "Kontakt", "Klara", "Tochter", "", "0171 888"],
            ],
        )

        kind, records = parse_excel(path)

        self.assertEqual("tabular_contacts", kind)
        self.assertEqual(1, len(records))
        self.assertEqual("Muster", records[0].lastname)
        self.assertEqual("Mina", records[0].firstname)
        self.assertEqual("PG 2", records[0].carelevel)
        self.assertEqual(2, len(records[0].contacts))
        self.assertEqual("Karl", records[0].contacts[0].firstname)
        self.assertEqual("Klara", records[0].contacts[1].firstname)

    def test_detect_sheet_kind_returns_unknown_for_unrelated_workbook(self) -> None:
        path = self.create_workbook_file(
            "unknown.xlsx",
            [
                ["foo", "bar"],
                ["baz", "qux"],
            ],
        )

        self.assertEqual("unknown", detect_sheet_kind(path))
        with self.assertRaisesRegex(ValueError, "Unknown sheet layout"):
            parse_excel(path)


if __name__ == "__main__":
    unittest.main()
