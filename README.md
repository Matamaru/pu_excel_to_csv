# Medifox Excel Phonebook Tool

Terminal-based Python tool to import Medifox Excel exports, keep a synchronized SQLite customer/contact phonebook, and export an active telephone list as CSV.

## Features

- Reads Excel files exported from Medifox
- Detects sheet type automatically (currently `medifox_report` and generic tabular format)
- Stores unique customers with:
  - name
  - firstname
  - pflegegrad
  - phone
  - mobile
- Stores multiple contacts per customer with:
  - contact_name
  - contact_firstname
  - contact_relation
  - contact_phone
  - contact_mobile
- Synchronizes active status on every import batch:
  - customer in Excel -> active
  - customer missing from current Excel batch -> inactive
- Search by phone number (customer and contacts)
- Export active phone list to CSV

## Quick Start

1. Create and activate virtual environment:

   PowerShell:

   ```powershell
   py -3 -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

2. Install dependencies:

   ```powershell
   pip install -e .
   ```

3. Run:

   ```powershell
  python src/main.py
  ```

  or after install:

  ```powershell
  medifox-phonebook
   ```

## Menu

- 1: Import Excel sheet(s)
- 2: Search by phone number
- 3: Export active telephone list CSV
- 4: Show DB stats
- 0: Exit

## Suggested Scaffold

A suggested next step if you want a larger maintainable structure:

```text
pu_excel_to_csv/
  src/
    main.py
    phonebook/
      __init__.py
      cli.py
      db.py
      parsers/
        __init__.py
        medifox_report.py
        tabular.py
      models.py
      exporter.py
  pyproject.toml
  README.md
  AGENTS.md
  ROADMAP.md
  tests/
    test_parsers.py
    test_sync.py
```

## Data File

- SQLite DB is created automatically as `phonebook.db` in the project root.

## Notes

- Sheet categorization is heuristic-based and can be extended for additional Medifox export variants.
- If your export layout differs, add parser rules in `src/main.py` (or split into parser modules as per scaffold).
