"""CSV export helpers."""

from __future__ import annotations

import csv
from pathlib import Path

from phonebook.models import HubRecord


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


def write_export_csv(out_path: Path, rows: list[tuple]) -> None:
    with out_path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(CSV_HEADER)
        writer.writerows(rows)


def default_export_filename(hubs: list[HubRecord]) -> str:
    if len(hubs) == 1:
        return f"active_telephone_list_{hubs[0].slug}.csv"
    return "active_telephone_list.csv"
