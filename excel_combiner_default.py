"""
Excel Combiner (Default)

Scans a folder of spreadsheet/text files and combines every sheet of
every workbook into a single master workbook - no sheet-name filter
(unlike extract_named_sheets.py) and no name/ssn/dob header detection
(unlike excel_pivotCheck_compiler.py). This is the generic, default
combiner for any mix of file types and sheet layouts.

  1. Every readable file type is handled:
       .xlsx .xlsm .xltx .xltm   - openpyxl
       .xls                      - xlrd, then calamine, then openpyxl,
                                   then an HTML-table fallback (many
                                   "xls" exports are really HTML)
       .xlsb                     - pyxlsb, then calamine
       .ods                      - odf
       .csv .txt .tsv            - delimiter sniffed, several encodings
                                   tried before giving up
     If the library for a format is missing, the log says exactly what
     to pip install instead of silently skipping the file.

  2. The header row is found rather than assumed. Sheets that start with
     title/logo/blank rows above the real header line up correctly with
     sheets that start at row 1. (Toggle off to force row 1.)

  3. Columns are matched across every file/sheet by header text,
     ignoring case, extra spaces, underscores and punctuation, so
     "Client Name", "client_name" and "CLIENT NAME " all land in one
     column. A header that has not been seen before becomes a brand-new
     column appended to the right; sheets without it are left blank.

  4. "Source File" and "Source Sheet" columns are added to every row so
     it can be traced back to its original workbook and worksheet.

  5. Sheets containing a pivot table are skipped - only that sheet, the
     rest of the workbook is still combined. Output larger than Excel's
     row limit is split across Combined, Combined_2, ... sheets.

Output: combined.xlsx at the destination you choose.

A small Tkinter GUI lets you pick the source folder and the output file,
with a progress bar tracking files processed.
"""

import csv
import os
import re
import threading
import traceback
from datetime import datetime

import openpyxl
import pandas as pd

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

COMBINED_FILENAME = "combined.xlsx"
SOURCE_COLUMN = "Source File"
SOURCE_SHEET_COLUMN = "Source Sheet"

# Excel's hard limit is 1,048,576 rows including the header row.
MAX_ROWS_PER_SHEET = 1_048_575

EXCEL_EXTS = (".xlsx", ".xlsm", ".xltx", ".xltm", ".xls", ".xlsb", ".ods")
TEXT_EXTS = (".csv", ".txt", ".tsv")
SUPPORTED_EXTS = EXCEL_EXTS + TEXT_EXTS

# Engines to try per extension, in order. The first one that opens the
# file wins; a missing library is reported once with an install hint.
ENGINE_CHAIN = {
    ".xlsx": ["openpyxl", "calamine"],
    ".xlsm": ["openpyxl", "calamine"],
    ".xltx": ["openpyxl"],
    ".xltm": ["openpyxl"],
    ".xls": ["xlrd", "calamine", "openpyxl"],
    ".xlsb": ["pyxlsb", "calamine"],
    ".ods": ["odf"],
}

INSTALL_HINT = {
    "xlrd": "pip install xlrd",
    "pyxlsb": "pip install pyxlsb",
    "calamine": "pip install python-calamine",
    "odf": "pip install odfpy",
    "openpyxl": "pip install openpyxl",
    "lxml": "pip install lxml",
}

TEXT_ENCODINGS = ["utf-8-sig", "cp1252", "latin-1"]

# Control characters openpyxl refuses to write.
ILLEGAL_CELL_CHARS = re.compile(r"[\000-\010\013\014\016-\037]")


# --------------------------------------------------------------------------
# Column name handling
# --------------------------------------------------------------------------

def normalize_key(name):
    """Matching key for a header: case, spacing, underscores, hyphens and
    punctuation are all ignored so near-identical headers from different
    files land in the same master column."""
    key = str(name).strip().lower()
    key = re.sub(r"[\s_\-]+", " ", key)
    key = re.sub(r"[^\w ]+", "", key)
    key = re.sub(r"\s+", " ", key).strip()
    return key


def is_placeholder(name):
    return bool(re.fullmatch(r"column_\d+", str(name).strip().lower()))


def clean_header_names(values):
    """Turn a raw header row into clean, unique column names, filling in
    blanks so pandas/openpyxl don't choke on NaN or duplicates."""
    headers = []
    seen = {}
    for idx, val in enumerate(values):
        name = "" if pd.isna(val) else str(val).strip()
        name = re.sub(r"\s+", " ", name)
        if not name or name.lower().startswith("unnamed:"):
            name = f"Column_{idx + 1}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        headers.append(name)
    return headers


class ColumnRegistry:
    """Maps each sheet's headers onto a shared master column order.

    Headers that normalize to the same key reuse the first display name
    seen for that key; anything new is appended to the right."""

    def __init__(self):
        self.order = []          # master column order, display names
        self.display_by_key = {}
        self.used_displays = set()

    def resolve_sheet(self, headers):
        """Return display names for one sheet's headers, registering any
        new ones. Duplicate keys inside a single sheet stay distinct."""
        out = []
        used_here = set()
        for header in headers:
            key = normalize_key(header) or str(header).strip().lower()
            base = key
            dup = 1
            while key in used_here:
                dup += 1
                key = f"{base} ({dup})"
            used_here.add(key)

            display = self.display_by_key.get(key)
            if display is None:
                display = str(header).strip() or key
                candidate = display
                dup = 1
                while candidate in self.used_displays:
                    dup += 1
                    candidate = f"{display} ({dup})"
                display = candidate
                self.display_by_key[key] = display
                self.used_displays.add(display)
                self.order.append(display)
            out.append(display)
        return out


# --------------------------------------------------------------------------
# Header row detection
# --------------------------------------------------------------------------

def _looks_numeric(value):
    try:
        float(str(value).replace(",", ""))
        return True
    except (TypeError, ValueError):
        return False


def detect_header_row(raw, max_scan=25):
    """Index of the most header-like row in a headerless frame.

    Scores each candidate on how full it is, how text-like its values
    are, and how unique they are; requires at least one non-empty row
    beneath it. Returns 0 when nothing scores well, which reproduces
    pandas' normal 'first row is the header' behaviour."""
    if raw.empty:
        return 0

    n_cols = max(raw.shape[1], 1)
    best_idx, best_score = 0, 0.0

    for i in range(min(max_scan, len(raw))):
        values = [
            str(v).strip()
            for v in raw.iloc[i].tolist()
            if not pd.isna(v) and str(v).strip()
        ]
        if not values:
            continue
        # A header needs data under it.
        below = raw.iloc[i + 1:i + 4]
        if below.empty or below.notna().sum().sum() == 0:
            continue

        filled = len(values) / n_cols
        textual = sum(1 for v in values if not _looks_numeric(v)) / len(values)
        unique = len({v.lower() for v in values}) / len(values)
        short = sum(1 for v in values if len(v) <= 60) / len(values)

        score = filled * 0.40 + textual * 0.30 + unique * 0.20 + short * 0.10
        # Prefer earlier rows when scores are close.
        score -= i * 0.005

        if score > best_score:
            best_score, best_idx = score, i

    return best_idx if best_score >= 0.55 else 0


def frame_from_raw(raw, detect_headers):
    """Split a headerless frame into header row + data rows."""
    raw = raw.dropna(how="all")
    if raw.empty:
        return pd.DataFrame()
    raw = raw.reset_index(drop=True)

    header_idx = detect_header_row(raw) if detect_headers else 0
    headers = clean_header_names(raw.iloc[header_idx].tolist())
    data = raw.iloc[header_idx + 1:].reset_index(drop=True)
    data.columns = headers

    data = data.dropna(how="all")
    if data.empty:
        return pd.DataFrame()

    # Drop trailing/placeholder columns that carry no data at all -
    # otherwise stray empty columns pollute the master column list.
    keep = [
        c for c in data.columns
        if not (is_placeholder(c) and data[c].isna().all())
    ]
    data = data[keep]
    return data


# --------------------------------------------------------------------------
# Readers
# --------------------------------------------------------------------------

def _missing_dependency(exc):
    """Return the package name if exc is a missing-optional-dep error."""
    text = str(exc)
    if not isinstance(exc, ImportError) and "Missing optional dependency" not in text:
        return None
    patterns = (
        r"Missing optional dependency ['\"`]([A-Za-z0-9_\-]+)",
        r"[Ii]mport ['\"`]?([A-Za-z0-9_\-]+)['\"`]? failed",
        r"No module named ['\"`]([A-Za-z0-9_\-]+)",
        r"['\"`]([A-Za-z0-9_\-]+)['\"`]",
    )
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(1)
    return None


def _with_install_hint(exc):
    """Error text with a 'pip install x' hint appended when the cause is
    a missing reader library."""
    missing = _missing_dependency(exc)
    if not missing:
        return str(exc)
    hint = INSTALL_HINT.get(missing, f"pip install {missing}")
    return f"{exc} -> {hint}"


def read_html_tables(file_path):
    """Fallback for .xls files that are really HTML tables (a very common
    export format from web reporting tools)."""
    for encoding in TEXT_ENCODINGS:
        try:
            tables = pd.read_html(file_path, header=None, encoding=encoding)
        except ImportError as exc:
            raise ImportError(str(exc))
        except Exception:
            continue
        if tables:
            return [(f"Table{i + 1}", t) for i, t in enumerate(tables)]
    return []


DELIMITERS = [",", ";", "\t", "|"]


def sniff_separator(sample, ext):
    """Best delimiter for a text sample, restricted to real delimiter
    characters - pandas' own sep=None sniffing is free to pick any
    character and will happily split "Name|Amount" on the letter 'm'."""
    if ext == ".tsv":
        return "\t"
    try:
        return csv.Sniffer().sniff(sample, delimiters="".join(DELIMITERS)).delimiter
    except Exception:
        pass

    lines = [ln for ln in sample.splitlines() if ln.strip()][:20]
    best, best_score = ",", 0
    for cand in DELIMITERS:
        counts = [ln.count(cand) for ln in lines]
        if not counts or min(counts) == 0:
            continue
        # Reward a delimiter that appears the same number of times on
        # every line - that is what a real column separator looks like.
        modal = max(set(counts), key=counts.count)
        score = (counts.count(modal) / len(counts)) * min(counts)
        if score > best_score:
            best, best_score = cand, score
    return best


def read_text_file(file_path, ext):
    """Read a csv/txt/tsv as a headerless frame, sniffing the delimiter
    and falling back through several encodings."""
    if os.path.getsize(file_path) == 0:
        return pd.DataFrame()

    last_error = None
    for encoding in TEXT_ENCODINGS:
        try:
            # Strict decoding is what makes the encoding fallback work:
            # a cp1252 file must fail on utf-8 rather than silently
            # decode into replacement characters.
            with open(file_path, "r", encoding=encoding, errors="strict", newline="") as fh:
                sample = fh.read(65536)
        except UnicodeDecodeError as exc:
            last_error = exc
            continue
        except Exception as exc:
            raise exc

        if not sample.strip():
            return pd.DataFrame()

        sniffed = sniff_separator(sample, ext)
        separators = [sniffed] + [s for s in DELIMITERS if s != sniffed]
        for sep in separators:
            try:
                return pd.read_csv(
                    file_path,
                    header=None,
                    dtype=object,
                    sep=sep,
                    encoding=encoding,
                    encoding_errors="strict",
                    skip_blank_lines=True,
                    on_bad_lines="skip",
                )
            except pd.errors.EmptyDataError:
                return pd.DataFrame()
            except Exception as exc:
                last_error = exc

    raise last_error if last_error else ValueError("could not read text file")


def open_excel_file(file_path, ext, log):
    """Open a workbook with the first engine that works. Returns an
    pd.ExcelFile, or None if every engine failed (already logged)."""
    errors = []
    for engine in ENGINE_CHAIN.get(ext, ["openpyxl"]):
        try:
            return pd.ExcelFile(file_path, engine=engine)
        except Exception as exc:
            errors.append(f"{engine}: {_with_install_hint(exc)}")
    for line in errors:
        log(f"      tried {line}")
    return None


def load_raw_sheets(file_path, ext, log):
    """Yield (sheet_name, headerless DataFrame) for any supported file."""
    if ext in TEXT_EXTS:
        return [("CSV" if ext == ".csv" else "Text", read_text_file(file_path, ext))]

    book = open_excel_file(file_path, ext, log)
    if book is None:
        if ext == ".xls":
            # Last resort: HTML masquerading as .xls.
            tables = read_html_tables(file_path)
            if tables:
                log("      recovered as HTML table(s)")
                return tables
        raise ValueError("no engine could open this file (see attempts above)")

    sheets = []
    try:
        for name in book.sheet_names:
            try:
                sheets.append((name, book.parse(sheet_name=name, header=None, dtype=object)))
            except Exception as exc:
                log(f"  ! Failed to read sheet '{name}': {exc}")
    finally:
        try:
            book.close()
        except Exception:
            pass
    return sheets


# --------------------------------------------------------------------------
# Pivot table detection
# --------------------------------------------------------------------------

def find_pivot_sheets(file_path, ext):
    """Sheet names in an .xlsx/.xlsm workbook holding a PivotTable, so
    only those sheets get skipped. Other formats return an empty set
    (openpyxl can't inspect them), i.e. nothing is excluded."""
    if ext not in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        return set()
    try:
        wb = openpyxl.load_workbook(file_path, read_only=False, data_only=True)
    except Exception:
        return set()
    try:
        return {ws.title for ws in wb.worksheets if getattr(ws, "_pivots", None)}
    except Exception:
        return set()
    finally:
        try:
            wb.close()
        except Exception:
            pass


# --------------------------------------------------------------------------
# Output sanitising
# --------------------------------------------------------------------------

def clean_cell(value):
    """Strip anything openpyxl refuses to write (control characters,
    timezone-aware datetimes)."""
    if isinstance(value, str):
        return ILLEGAL_CELL_CHARS.sub("", value)
    if isinstance(value, datetime) and value.tzinfo is not None:
        return value.replace(tzinfo=None)
    return value


def write_output(combined, dest_path, log, status):
    """Write the combined frame, splitting across sheets if it exceeds
    Excel's row limit."""
    status("Cleaning values for Excel...")
    for col in combined.columns:
        if combined[col].dtype == object:
            combined[col] = combined[col].map(clean_cell)

    status("Writing output workbook...")
    chunks = [
        combined.iloc[i:i + MAX_ROWS_PER_SHEET]
        for i in range(0, max(len(combined), 1), MAX_ROWS_PER_SHEET)
    ] or [combined]

    with pd.ExcelWriter(dest_path, engine="openpyxl") as writer:
        for i, chunk in enumerate(chunks):
            sheet = "Combined" if i == 0 else f"Combined_{i + 1}"
            chunk.to_excel(writer, index=False, sheet_name=sheet)
            if i:
                log(f"  * row limit reached, continued on sheet '{sheet}'")


# --------------------------------------------------------------------------
# Combine
# --------------------------------------------------------------------------

def collect_files(source_folder, dest_abs, recursive):
    """All supported input files, skipping Excel lock files and the
    output workbook itself (so re-runs don't read their own output)."""
    found = []
    if recursive:
        walker = os.walk(source_folder)
    else:
        walker = [(source_folder, [], os.listdir(source_folder))]

    for root, _dirs, names in walker:
        for name in names:
            if name.startswith("~$") or name.startswith("."):
                continue
            if not name.lower().endswith(SUPPORTED_EXTS):
                continue
            full = os.path.join(root, name)
            if os.path.abspath(full) == dest_abs:
                continue
            found.append(full)
    return sorted(found)


def combine_files(source_folder, dest_path, log, progress, status=None,
                  recursive=True, detect_headers=True, skip_pivots=True):
    """status(msg), if given, is called with a short description of the
    current step so a GUI can show what's happening right now instead of
    just a bare percentage."""
    def _status(msg):
        if status:
            status(msg)

    if not os.path.isdir(source_folder):
        raise ValueError(f"Source folder not found: {source_folder}")

    dest_abs = os.path.abspath(dest_path)
    files = collect_files(source_folder, dest_abs, recursive)
    if not files:
        raise ValueError(
            "No supported files found in the source folder "
            f"({', '.join(SUPPORTED_EXTS)})"
        )

    total_files = len(files)
    log(f"Found {total_files} file(s) to combine"
        f"{' (including subfolders)' if recursive else ''}.")
    progress(0, total_files)

    registry = ColumnRegistry()
    frames = []
    failed_files = []
    sheet_count = 0

    for file_idx, file_path in enumerate(files, start=1):
        filename = os.path.relpath(file_path, source_folder)
        ext = os.path.splitext(file_path)[1].lower()
        _status(f"Opening {filename}...")

        try:
            raw_sheets = load_raw_sheets(file_path, ext, log)
        except Exception as exc:
            log(f"  ! Failed to open {filename}: {_with_install_hint(exc)}")
            failed_files.append(filename)
            progress(file_idx, total_files)
            continue

        if not raw_sheets:
            log(f"  - {filename}: no readable sheets, skipped")
            failed_files.append(filename)
            progress(file_idx, total_files)
            continue

        pivot_sheets = find_pivot_sheets(file_path, ext) if skip_pivots else set()

        for sheet_name, raw in raw_sheets:
            if sheet_name in pivot_sheets:
                log(f"  - {filename} [{sheet_name}]: pivot table detected, sheet excluded")
                continue

            _status(f"{filename}: reading sheet '{sheet_name}'...")
            try:
                data = frame_from_raw(raw, detect_headers)
            except Exception as exc:
                log(f"  ! Failed to parse {filename} [{sheet_name}]: {exc}")
                continue

            if data.empty:
                log(f"  - {filename} [{sheet_name}]: empty sheet, skipped")
                continue

            before = len(registry.order)
            data.columns = registry.resolve_sheet(list(data.columns))
            new_cols = registry.order[before:]

            data.insert(0, SOURCE_SHEET_COLUMN, sheet_name)
            data.insert(0, SOURCE_COLUMN, filename)

            detail = (f"{len(new_cols)} new -> {new_cols}" if new_cols
                      else "no new columns")
            log(f"  - {filename} [{sheet_name}]: {len(data)} row(s), "
                f"{len(data.columns) - 2} column(s), {detail}")

            frames.append(data)
            sheet_count += 1

        progress(file_idx, total_files)

    if not frames:
        raise ValueError("No sheets could be read; nothing to combine.")

    _status("Combining all sheets...")
    combined = pd.concat(frames, ignore_index=True, sort=False)
    combined = combined.reindex(
        columns=[SOURCE_COLUMN, SOURCE_SHEET_COLUMN] + registry.order
    )

    write_output(combined, dest_path, log, _status)

    progress(total_files, total_files)
    _status("Done.")
    log(f"Combined {sheet_count} sheet(s) across "
        f"{total_files - len(failed_files)} of {total_files} file(s), "
        f"{len(combined)} row(s), {len(registry.order)} data column(s) -> {dest_path}")
    if failed_files:
        log(f"Could not read {len(failed_files)} file(s): {', '.join(failed_files)}")
    return dest_path


# --------------------------------------------------------------------------
# Tkinter GUI
# --------------------------------------------------------------------------

class CombinerApp:
    def __init__(self, root):
        self.root = root
        root.title("Excel Combiner (Default)")
        root.geometry("720x540")
        root.resizable(True, True)

        self.source_var = tk.StringVar()
        self.dest_var = tk.StringVar()
        self.recursive_var = tk.BooleanVar(value=True)
        self.detect_headers_var = tk.BooleanVar(value=True)
        self.skip_pivots_var = tk.BooleanVar(value=True)

        pad = {"padx": 8, "pady": 6}

        frame = tk.Frame(root)
        frame.pack(fill="x", **pad)

        tk.Label(frame, text="Source folder:").grid(row=0, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.source_var, width=60).grid(row=0, column=1, padx=6)
        tk.Button(frame, text="Browse...", command=self.browse_source).grid(row=0, column=2)

        tk.Label(frame, text="Output file:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        tk.Entry(frame, textvariable=self.dest_var, width=60).grid(row=1, column=1, padx=6, pady=(8, 0))
        tk.Button(frame, text="Save As...", command=self.browse_dest).grid(row=1, column=2, pady=(8, 0))

        options = tk.Frame(root)
        options.pack(fill="x", padx=8)
        tk.Checkbutton(options, text="Include subfolders",
                       variable=self.recursive_var).pack(side="left")
        tk.Checkbutton(options, text="Auto-detect header row",
                       variable=self.detect_headers_var).pack(side="left", padx=(12, 0))
        tk.Checkbutton(options, text="Skip pivot sheets",
                       variable=self.skip_pivots_var).pack(side="left", padx=(12, 0))

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

        options = {
            "recursive": self.recursive_var.get(),
            "detect_headers": self.detect_headers_var.get(),
            "skip_pivots": self.skip_pivots_var.get(),
        }
        thread = threading.Thread(
            target=self._run_worker, args=(source, dest, options), daemon=True
        )
        thread.start()

    def _run_worker(self, source, dest, options):
        try:
            dest_path = combine_files(
                source, dest, self.log, self.progress, status=self.status, **options
            )
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
