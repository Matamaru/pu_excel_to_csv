# Medifox Excel Phonebook Tool

Python tool to import Medifox Excel exports, keep a synchronized SQLite customer/contact phonebook, and export an active telephone list as CSV.

## Features

- Reads Excel files exported from Medifox
- Tkinter desktop menu (default launcher)
- Terminal menu fallback if Tkinter cannot start
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
- Export deduplicated HalloLena single-phone CSV
- Dedicated `uploads/` folder workflow for Excel imports

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

4. If Tkinter cannot start, the app falls back to terminal mode automatically.

## Import Workflow

1. Put export files into `uploads/`.
2. In the GUI, use one of these options:
  - Import All From uploads
  - Select From uploads And Import
  - Add File(s) And Import
3. The import sync logic sets `active = 1` for customers in the import batch and `active = 0` for customers not in the current batch.

## Menu

- Import all Excel files from uploads folder
- Select files from uploads folder and import
- Add file(s) to uploads and import directly
- Search by phone number
- Export active telephone list CSV
- Export HalloLena single phone list CSV
- Show DB stats

## HalloLena Format

- File content is flattened to one person-phone per row (customers and contacts).
- Duplicate rows are removed.
- Output columns are:
  - `phone_e164`
  - `first_name`
  - `last_name`
  - `email` (currently empty)

Example header:

```csv
"phone_e164","first_name","last_name","email"
```

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
- `uploads/` and `phonebook.db` are git-ignored.

## Notes

- Sheet categorization is heuristic-based and can be extended for additional Medifox export variants.
- If your export layout differs, add parser rules in `src/main.py` (or split into parser modules as per scaffold).
