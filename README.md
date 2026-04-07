# Medifox Excel Phonebook Tool

Desktop-first Python utility for importing Medifox Excel exports into a searchable SQLite phonebook and exporting hub-filtered telephone book CSV files.

## What It Does

- imports Medifox Excel sheets into a local SQLite database
- keeps customers unique via a computed stable key
- stores zero to many contacts per customer
- synchronizes `active` status per hub on each successful import
- supports multiple hubs such as `Dorsten`, `Essen`, or `Bocholt`
- exports active telephone lists for one hub, multiple hubs, or all hubs
- offers a Tkinter UI when available and a terminal fallback otherwise

## Hub Model

The app now treats each hub as its own import scope.

- each customer belongs to exactly one hub
- importing hub `Dorsten` only updates `Dorsten`
- customers missing from the latest import batch for that hub become inactive for that hub
- customers in other hubs remain unchanged
- uploads are stored under `uploads/<hub-slug>/`
- CSV export can combine hubs through checkbox selection in the GUI or multi-select in terminal mode

Legacy single-hub databases are migrated automatically on first start of the hub-aware version. Existing rows are assigned to hub `Dorsten`.

## Requirements

- Python `>= 3.11`
- `openpyxl`
- optional: `tkinter` for the desktop UI

If `tkinter` is not installed, the program starts in terminal mode automatically.

## Setup

### Windows

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -e .
```

### Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -e .
```

For the Tkinter desktop UI on Debian/Ubuntu, install:

```bash
sudo apt install python3-tk
```

If you also need venv support on a minimal system:

```bash
sudo apt install python3-venv
```

## Run

From the repo:

```bash
python src/main.py
```

Or after editable install:

```bash
medifox-phonebook
```

## Typical Workflow

### 1. Create A Hub

Create a hub before the first import, for example:

- `Dorsten`
- `Essen`
- `Bocholt`

The app creates a matching uploads folder automatically.

### 2. Import Excel Files For That Hub

You can:

- import all Excel files from that hub's uploads folder
- choose selected files from that hub's uploads folder
- add external files into that hub's uploads folder and import them
- import files directly from arbitrary paths

Unknown Excel formats are handled as non-fatal parse errors. Other valid files in the batch still continue.

### 3. Search

Phone number search checks:

- customer phone
- customer mobile
- contact phone
- contact mobile

Search results show the hub and whether the customer is active or inactive.

### 4. Export Telephone Book CSV

CSV export includes active customers only.

- GUI: select one or more hubs with checkboxes
- terminal: select one, many, or `all`

The exported CSV includes a leading `hub` column so combined exports remain traceable.

## Menus

### GUI

- Create Hub
- Import All From Hub uploads
- Select From Hub uploads And Import
- Add File(s) To Hub And Import
- Search By Phone
- Export Active CSV By Hubs
- Show DB Stats

### Terminal

- Upload / Import Excel sheet(s)
- Search by phone number
- Export active telephone list CSV
- Show DB stats
- Create hub

## File Layout

```text
pu_excel_to_csv/
  src/main.py
  tests/test_hub_support.py
  uploads/
    dorsten/
    essen/
  phonebook.db
```

## Data Files

- the SQLite database is created as `phonebook.db` in the project root
- uploads are stored under `uploads/`
- `uploads/` and `phonebook.db` are git-ignored

## Development

Run a quick syntax check:

```bash
python -m py_compile src/main.py tests/test_hub_support.py
```

Run tests:

```bash
python -m unittest discover -s tests -v
```

## Notes

- sheet type detection is heuristic-based and currently supports `medifox_report` and a generic tabular format
- if your Medifox export layout differs, extend the parser logic in `src/main.py`
- contacts are replaced for a customer on each successful import of that customer
- the database is never deleted automatically
