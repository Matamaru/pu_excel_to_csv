# ROADMAP

## Phase 1 (Done)

- Terminal menu application
- SQLite schema for customers and contacts
- Excel sheet type detection (initial)
- Import synchronization with active/inactive logic
- Phone search (customer + contacts)
- Active telephone list CSV export

## Phase 2

- Split monolithic `main.py` into `src/phonebook` package modules
- Add unit tests for parser and sync logic
- Add sample fixture spreadsheets for regression checks
- Add logging to file for import diagnostics

## Phase 3

- Add parser plugins for additional Medifox layouts
- Add duplicate/merge assistant for near-identical customers
- Optional simple web UI for lookup and export

## Phase 4

- Scheduled import mode from watch folder
- Audit trail table for change history
- Data validation report before sync apply
