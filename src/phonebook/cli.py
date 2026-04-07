"""Terminal and Tkinter entrypoints for the phonebook app."""

from __future__ import annotations

import os
import sqlite3
import sys
from pathlib import Path

from phonebook.db import (
    DB_PATH,
    UPLOADS_DIR,
    copy_files_into_uploads,
    describe_upload_file,
    ensure_hub,
    ensure_hub_uploads_dir,
    ensure_uploads_dir,
    fetch_export_rows,
    format_stats_text,
    get_hub_counts,
    is_excel_file,
    list_hubs,
    list_upload_excel_files,
    parse_paths_input,
    search_matches,
    sync_customers,
    init_db,
)
from phonebook.exporter import default_export_filename, write_export_csv
from phonebook.models import CustomerRecord, HubRecord, normalized_phone_digits
from phonebook.parsers import parse_excel

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog
except Exception:
    tk = None
    filedialog = None
    messagebox = None
    simpledialog = None


def clear_screen() -> None:
    os.system("cls" if os.name == "nt" else "clear")


def pause() -> None:
    input("\nPress Enter to continue...")


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
    for token in [item.strip() for item in selection.split(",") if item.strip()]:
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
    for token in [item.strip() for item in selection.split(",") if item.strip()]:
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
    hub_total, hub_active = get_hub_counts(conn, hub.id)
    print(
        f"Import done for hub {hub.name}. "
        f"Hub customers: {hub_active}/{hub_total} active | DB total: {total}, active: {active}"
    )
    summary["total"] = total
    summary["active"] = active
    summary["hub_total"] = hub_total
    summary["hub_active"] = hub_active
    return summary


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
        files = [Path(path) for path in selected if is_excel_file(Path(path))]
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

    for text, command in buttons:
        tk.Button(frame, text=text, command=command, width=36, pady=6).pack(pady=3)

    refresh_subtitle()
    root.mainloop()


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
