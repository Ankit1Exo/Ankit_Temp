"""
Excel Combiner (Default)

Scans a folder of .xlsx / .xls files and combines every sheet of every
workbook into a single master workbook - no sheet-name filter (unlike
extract_named_sheets.py) and no pivot-table skip or name/ssn/dob header
detection (unlike excel_pivotCheck_compiler.py). This is the generic,
default combiner for any mix of sheet layouts.

  1. Every sheet in every file is read (the first row is treated as the
     header - pandas' normal default).

  2. Columns are matched across every file/sheet by header text:
       - A header name already seen in a previous sheet lines up under
         the same master column.
       - A header name that hasn't been seen before becomes a brand-new
         column appended to the right; sheets that don't have that
         column are left blank there.

  3. "Source File" and "Source Sheet" columns are added to every row so
     it can be traced back to its original workbook and worksheet.

Output: combined.xlsx in the destination file you choose, with every
sheet's rows stacked into one "Combined" sheet.

A small Tkinter GUI lets you pick the source folder and the output file,
with a progress bar tracking sheets processed.
"""

import os
import threading
import traceback

import pandas as pd

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

COMBINED_FILENAME = "combined.xlsx"
SOURCE_COLUMN = "Source File"
SOURCE_SHEET_COLUMN = "Source Sheet"


# --------------------------------------------------------------------------
# Header cleanup
# --------------------------------------------------------------------------

def clean_header_names(columns):
    """Turn a raw header row (pandas column index) into clean column
    names, filling in blanks so pandas/openpyxl don't choke on
    duplicates/NaN/"Unnamed: N" placeholders."""
    headers = []
    seen = {}
    for idx, val in enumerate(columns):
        name = "" if pd.isna(val) else str(val).strip()
        if not name or name.lower().startswith("unnamed:"):
            name = f"Column_{idx + 1}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        headers.append(name)
    return headers


# --------------------------------------------------------------------------
# Combine
# --------------------------------------------------------------------------

def combine_files(source_folder, dest_path, log, progress, status=None):
    """status(msg), if given, is called with a short description of the
    current step (which file/sheet is being read, when the final
    workbook is being written, etc.) so a GUI can show what's happening
    right now instead of just a bare percentage."""
    def _status(msg):
        if status:
            status(msg)

    if not os.path.isdir(source_folder):
        raise ValueError(f"Source folder not found: {source_folder}")

    # Exclude the output file itself from the input list, in case it was
    # saved into the same folder as the source files and this is a
    # re-run - otherwise the previous combined.xlsx would get read back
    # in as a source file on the next pass.
    dest_abs = os.path.abspath(dest_path)
    files = sorted(
        f for f in os.listdir(source_folder)
        if f.lower().endswith((".xlsx", ".xls"))
        and not f.startswith("~$")
        and os.path.abspath(os.path.join(source_folder, f)) != dest_abs
    )
    if not files:
        raise ValueError("No .xlsx or .xls files found in the source folder")

    total_files = len(files)
    log(f"Found {total_files} file(s) to combine.")
    progress(0, total_files)

    known_columns = []
    known_set = set()
    frames = []

    for file_idx, filename in enumerate(files, start=1):
        file_path = os.path.join(source_folder, filename)
        _status(f"Opening {filename}...")

        try:
            book = pd.ExcelFile(file_path)
        except Exception as exc:
            log(f"  ! Failed to open {filename}: {exc}")
            progress(file_idx, total_files)
            continue

        for sheet_name in book.sheet_names:
            _status(f"{filename}: reading sheet '{sheet_name}'...")
            try:
                data = book.parse(sheet_name=sheet_name, dtype=object)
            except Exception as exc:
                log(f"  ! Failed to read {filename} [{sheet_name}]: {exc}")
                continue

            data.columns = clean_header_names(data.columns)
            data.dropna(how="all", inplace=True)
            if data.empty:
                log(f"  - {filename} [{sheet_name}]: empty sheet, skipped")
                continue

            new_cols = [c for c in data.columns if c not in known_set]
            for c in new_cols:
                known_set.add(c)
                known_columns.append(c)

            data.insert(0, SOURCE_SHEET_COLUMN, sheet_name)
            data.insert(0, SOURCE_COLUMN, filename)

            if new_cols:
                log(f"  - {filename} [{sheet_name}]: {len(data)} row(s), "
                    f"{len(data.columns) - 2} column(s), {len(new_cols)} new -> {new_cols}")
            else:
                log(f"  - {filename} [{sheet_name}]: {len(data)} row(s), "
                    f"{len(data.columns) - 2} column(s), no new columns")

            frames.append(data)

        progress(file_idx, total_files)

    if not frames:
        raise ValueError("No sheets could be read; nothing to combine.")

    _status("Combining all sheets...")
    combined = pd.concat(frames, ignore_index=True, sort=False)
    column_order = [SOURCE_COLUMN, SOURCE_SHEET_COLUMN] + known_columns
    combined = combined[column_order]

    _status("Writing output workbook...")
    combined.to_excel(dest_path, index=False, sheet_name="Combined")

    progress(total_files, total_files)
    _status("Done.")
    log(f"Combined {len(frames)} sheet(s) across {total_files} file(s), "
        f"{len(combined)} row(s), {len(known_columns)} data column(s) -> {dest_path}")
    return dest_path


# --------------------------------------------------------------------------
# Tkinter GUI
# --------------------------------------------------------------------------

class CombinerApp:
    def __init__(self, root):
        self.root = root
        root.title("Excel Combiner (Default)")
        root.geometry("680x480")
        root.resizable(True, True)

        self.source_var = tk.StringVar()
        self.dest_var = tk.StringVar()

        pad = {"padx": 8, "pady": 6}

        frame = tk.Frame(root)
        frame.pack(fill="x", **pad)

        tk.Label(frame, text="Source folder:").grid(row=0, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.source_var, width=60).grid(row=0, column=1, padx=6)
        tk.Button(frame, text="Browse...", command=self.browse_source).grid(row=0, column=2)

        tk.Label(frame, text="Output file:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        tk.Entry(frame, textvariable=self.dest_var, width=60).grid(row=1, column=1, padx=6, pady=(8, 0))
        tk.Button(frame, text="Save As...", command=self.browse_dest).grid(row=1, column=2, pady=(8, 0))

        self.run_button = tk.Button(root, text="Run", width=20, command=self.on_run)
        self.run_button.pack(pady=10)

        progress_frame = tk.Frame(root)
        progress_frame.pack(fill="x", padx=8)
        self.progress_bar = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress_bar.pack(fill="x", side="left", expand=True)
        self.progress_label = tk.Label(progress_frame, text="0 / 0", width=10)
        self.progress_label.pack(side="left", padx=(8, 0))

        self.status_var = tk.StringVar(value="Idle")
        tk.Label(root, textvariable=self.status_var, anchor="w", fg="#444").pack(
            fill="x", padx=8, pady=(4, 0)
        )

        tk.Label(root, text="Log:").pack(anchor="w", padx=8, pady=(8, 0))
        self.log_box = scrolledtext.ScrolledText(root, height=16, state="disabled")
        self.log_box.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def browse_source(self):
        path = filedialog.askdirectory(title="Select source folder")
        if path:
            self.source_var.set(path)
            if not self.dest_var.get().strip():
                self.dest_var.set(os.path.join(path, COMBINED_FILENAME))

    def browse_dest(self):
        path = filedialog.asksaveasfilename(
            title="Save combined workbook as",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            initialfile=COMBINED_FILENAME,
        )
        if path:
            self.dest_var.set(path)

    def log(self, message):
        def append():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", message + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.root.after(0, append)

    def progress(self, done, total):
        def update():
            self.progress_bar.configure(maximum=max(total, 1), value=done)
            self.progress_label.configure(text=f"{done} / {total}")
        self.root.after(0, update)

    def status(self, message):
        self.root.after(0, lambda: self.status_var.set(message))

    def on_run(self):
        source = self.source_var.get().strip()
        dest = self.dest_var.get().strip()
        if not source or not dest:
            messagebox.showerror("Missing path", "Please select both a source folder and an output file.")
            return

        self.run_button.configure(state="disabled")
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
        self.progress(0, 0)
        self.status("Starting...")

        thread = threading.Thread(target=self._run_worker, args=(source, dest), daemon=True)
        thread.start()

    def _run_worker(self, source, dest):
        try:
            dest_path = combine_files(source, dest, self.log, self.progress, status=self.status)
            self.root.after(0, lambda: messagebox.showinfo(
                "Done", f"Combination complete.\nSaved to:\n{dest_path}"
            ))
        except Exception as exc:
            self.status(f"Failed: {exc}")
            self.log("ERROR: " + str(exc))
            self.log(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            self.root.after(0, lambda: self.run_button.configure(state="normal"))


def main():
    root = tk.Tk()
    CombinerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
