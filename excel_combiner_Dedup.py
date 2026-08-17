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

  4. Any sheet whose header row is exactly {Md Rc, Loc Name, Res Name,
     Appt Dt, Pat Name, Per Nbr, Sts} (case-insensitive, regardless of
     sheet name) is treated specially: it is collapsed to one row per
     Pat Name, with that name's Md Rc values joined by ":" and its
     Per Nbr values joined by ":". Only those three columns are carried
     forward - Loc Name, Res Name, Appt Dt and Sts are dropped.

  5. If the combined result has more rows than a worksheet can hold,
     it is written as several workbooks - "combined part 1.xlsx",
     "combined part 2.xlsx", ... - sized as evenly as possible. All
     rows from one source file always stay together in the same part.

Output: combined.xlsx in the destination file you choose, with every
sheet's rows stacked into one "Combined" sheet (or that name plus
"part N" per file, if the row limit forces a split).

A small Tkinter GUI lets you pick the source folder and the output file,
with a progress bar tracking sheets processed.
"""

import os
import re
import threading
import traceback

import pandas as pd

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

COMBINED_FILENAME = "combined.xlsx"
SOURCE_COLUMN = "Source File"
SOURCE_SHEET_COLUMN = "Source Sheet"

# A worksheet holds 1,048,576 rows; one of those goes to the header, so
# this is the most data rows that fit in a single output workbook.
MAX_DATA_ROWS = 1_048_576 - 1

# A sheet whose header row matches this set exactly (case-insensitive,
# any order) is collapsed - one row per Pat Name, with that name's
# Md Rc and Per Nbr values each joined by ":". The remaining columns
# (Loc Name, Res Name, Appt Dt, Sts) are dropped.
COLLAPSE_HEADERS = {"Md Rc", "Loc Name", "Res Name", "Appt Dt", "Pat Name", "Per Nbr", "Sts"}
COLLAPSE_KEY_COLUMN = "Pat Name"
COLLAPSE_VALUE_COLUMNS = ["Md Rc", "Per Nbr"]
COLLAPSE_JOINER = ":"


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
# Collapse sheet detection
# --------------------------------------------------------------------------

def is_collapse_target(columns):
    """True if a sheet's (cleaned) header row is exactly the collapse
    header set, case-insensitive and regardless of order."""
    return {str(c).strip().lower() for c in columns} == {h.lower() for h in COLLAPSE_HEADERS}


def cell_to_text(value):
    """Render one cell as the text that should end up inside a joined
    string - blanks as "", and whole numbers without the trailing ".0"
    pandas gives them when a column comes back as float."""
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def collapse_sheet(data):
    """One row per Pat Name, holding that name's Md Rc values joined by
    ":" and its Per Nbr values joined by ":", in the order the rows
    appeared on the sheet. Every other column is dropped.

    Names are matched case-insensitively (the first spelling seen is the
    one kept). A repeated Md Rc/Per Nbr pair for the same name is only
    listed once, so the two joined lists stay aligned position for
    position; a row where both values are blank is skipped entirely."""
    order = []
    groups = {}

    for _, row in data.iterrows():
        display_name = cell_to_text(row[COLLAPSE_KEY_COLUMN])
        values = tuple(cell_to_text(row[c]) for c in COLLAPSE_VALUE_COLUMNS)
        if not any(values):
            continue

        key = display_name.casefold()
        if key not in groups:
            groups[key] = {"name": display_name, "pairs": [], "seen": set()}
            order.append(key)

        group = groups[key]
        if values in group["seen"]:
            continue
        group["seen"].add(values)
        group["pairs"].append(values)

    rows = []
    for key in order:
        group = groups[key]
        collapsed = {COLLAPSE_KEY_COLUMN: group["name"]}
        for idx, column in enumerate(COLLAPSE_VALUE_COLUMNS):
            collapsed[column] = COLLAPSE_JOINER.join(pair[idx] for pair in group["pairs"])
        rows.append(collapsed)

    return pd.DataFrame(rows, columns=[COLLAPSE_KEY_COLUMN] + COLLAPSE_VALUE_COLUMNS)


# --------------------------------------------------------------------------
# Splitting the output across part files
# --------------------------------------------------------------------------

def part_path(dest_path, part_number):
    """'C:/out/combined.xlsx' + 2 -> 'C:/out/combined part 2.xlsx'."""
    base, ext = os.path.splitext(dest_path)
    return f"{base} part {part_number}{ext or '.xlsx'}"


def pack_blocks(sizes, capacity):
    """Walk the blocks in order and start a new part whenever the next
    block would push the current one past `capacity`. Returns a list of
    lists of block positions."""
    parts, current, current_rows = [], [], 0
    for position, size in enumerate(sizes):
        if current and current_rows + size > capacity:
            parts.append(current)
            current, current_rows = [], 0
        current.append(position)
        current_rows += size
    if current:
        parts.append(current)
    return parts


def plan_parts(sizes, limit):
    """Split blocks across parts without ever breaking a block, using as
    few parts as the row limit allows and making those parts as even as
    possible.

    First work out how many parts are unavoidable, then binary-search
    the smallest per-part capacity that still fits in that many - which
    is what stops a run from ending in one full part and one tiny one."""
    if not sizes:
        return []
    if sum(sizes) <= limit:
        return [list(range(len(sizes)))]

    # Blocks are capped at `limit` before we get here, so packing at
    # `limit` always succeeds; that count is the floor to aim for.
    parts_wanted = max(-(-sum(sizes) // limit), len(pack_blocks(sizes, limit)))

    low, high = max(sizes), limit
    while low < high:
        middle = (low + high) // 2
        if len(pack_blocks(sizes, middle)) <= parts_wanted:
            high = middle
        else:
            low = middle + 1
    return pack_blocks(sizes, low)


def source_file_blocks(combined, log):
    """Split the combined frame into one block of row positions per
    source file - the unit that must never be broken across parts. A
    single source file too big to fit in one workbook is the one case
    that has to be chopped up, and it is logged loudly when it happens."""
    by_file = {}
    for position, name in enumerate(combined[SOURCE_COLUMN]):
        by_file.setdefault(name, []).append(position)

    blocks = []
    for name, positions in by_file.items():
        if len(positions) <= MAX_DATA_ROWS:
            blocks.append((name, positions))
            continue
        chunks = -(-len(positions) // MAX_DATA_ROWS)
        log(f"  ! {name} on its own has {len(positions)} row(s), more than the "
            f"{MAX_DATA_ROWS} a worksheet holds - it cannot be kept whole and "
            f"is spread over {chunks} part(s)")
        for start in range(0, len(positions), MAX_DATA_ROWS):
            blocks.append((name, positions[start:start + MAX_DATA_ROWS]))
    return blocks


def write_parts(combined, dest_path, log, status):
    """Write the combined frame to dest_path, or to 'dest part N' files
    if it is too big for one worksheet. Returns the paths written."""
    blocks = source_file_blocks(combined, log)
    parts = plan_parts([len(positions) for _, positions in blocks], MAX_DATA_ROWS)

    if len(parts) <= 1:
        status("Writing output workbook...")
        combined.to_excel(dest_path, index=False, sheet_name="Combined")
        return [dest_path]

    log(f"{len(combined)} row(s) is past the {MAX_DATA_ROWS} row worksheet limit "
        f"- splitting into {len(parts)} part file(s), keeping each source file whole")

    written = []
    for part_number, block_positions in enumerate(parts, start=1):
        status(f"Writing part {part_number} of {len(parts)}...")
        rows = [row for block in block_positions for row in blocks[block][1]]
        target = part_path(dest_path, part_number)
        combined.iloc[rows].to_excel(target, index=False, sheet_name="Combined")
        written.append(target)
        source_count = len({blocks[block][0] for block in block_positions})
        log(f"  -> {os.path.basename(target)}: {len(rows)} row(s) "
            f"from {source_count} source file(s)")
    return written


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

    # Exclude the output from the input list, in case it was saved into
    # the same folder as the source files and this is a re-run -
    # otherwise the previous combined.xlsx (or its "part N" files) would
    # get read back in as source files on the next pass.
    dest_abs = os.path.abspath(dest_path)
    dest_stem, dest_ext = os.path.splitext(os.path.basename(dest_path))
    part_pattern = re.compile(rf"^{re.escape(dest_stem)} part \d+$", re.IGNORECASE)
    dest_in_source_folder = (os.path.abspath(os.path.dirname(dest_abs))
                             == os.path.abspath(source_folder))

    def is_own_output(name):
        if os.path.abspath(os.path.join(source_folder, name)) == dest_abs:
            return True
        stem, ext = os.path.splitext(name)
        return (dest_in_source_folder
                and ext.lower() == dest_ext.lower()
                and bool(part_pattern.match(stem)))

    files = sorted(
        f for f in os.listdir(source_folder)
        if f.lower().endswith((".xlsx", ".xls"))
        and not f.startswith("~$")
        and not is_own_output(f)
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

            if is_collapse_target(data.columns):
                before = len(data)
                data = collapse_sheet(data)
                log(f"  * {filename} [{sheet_name}]: appointment sheet detected - "
                    f"{before} row(s) -> {len(data)} {COLLAPSE_KEY_COLUMN} group(s); "
                    f"kept {COLLAPSE_KEY_COLUMN} + "
                    f"{'/'.join(COLLAPSE_VALUE_COLUMNS)} joined with "
                    f"'{COLLAPSE_JOINER}', dropped all other columns")
                if data.empty:
                    log(f"  - {filename} [{sheet_name}]: nothing left after collapsing, skipped")
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

    written = write_parts(combined, dest_path, log, _status)

    progress(total_files, total_files)
    _status("Done.")
    log(f"Combined {len(frames)} sheet(s) across {total_files} file(s), "
        f"{len(combined)} row(s), {len(known_columns)} data column(s) -> "
        + ", ".join(written))
    return written


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
            written = combine_files(source, dest, self.log, self.progress, status=self.status)
            if len(written) == 1:
                summary = f"Combination complete.\nSaved to:\n{written[0]}"
            else:
                files = "\n".join(os.path.basename(p) for p in written)
                summary = (f"Combination complete.\nToo many rows for one workbook, "
                           f"so it was split into {len(written)} parts in\n"
                           f"{os.path.dirname(written[0])}:\n\n{files}")
            self.root.after(0, lambda: messagebox.showinfo("Done", summary))
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
