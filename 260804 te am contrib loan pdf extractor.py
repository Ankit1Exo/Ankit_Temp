"""
Contributions/Loans PDF Extractor

Scans a folder of searchable (text-based) PDF "Contributions Loans Report"
files and pulls out just the participant identity columns:

    Member ID / Member Name   (Contributions table)
    Participant ID / Name    (Loans table)

Both pairs land in the same two output columns (ID, Name) since a given
row is only ever one or the other. Blank when the source table doesn't
provide it (e.g. Loans rows commonly have a blank Participant ID).

Column positions are auto-detected per PDF from the header row text
("Member ID", "Participant ID", "Member Name", "Name") rather than
hard-coded coordinates, so it tolerates the two known header formats and
minor layout drift between reports. If a report uses different wording
for these headers, nothing will be extracted for it -- check the log for
files that returned 0 rows.

Output columns: File Name | Page Number | ID | Name

If the combined row count exceeds Excel's per-sheet limit
(1,048,576 rows), the output is automatically split into
..._part1.xlsx, ..._part2.xlsx, etc.

REQUIREMENTS:
    pip install pymupdf openpyxl

USAGE (GUI):
    python "260804 te am contrib loan pdf extractor.py"
"""

import os
import re
import threading
import time
import traceback
from pathlib import Path

import fitz  # PyMuPDF
from openpyxl import Workbook

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# ---------------------------------------------------------------------------
# Tunables
# ---------------------------------------------------------------------------

# Regex for lines that end the current table segment.
STOP_RE = re.compile(r"\b(subtotals?|eft\s+totals?)\b", re.IGNORECASE)

# Excel's hard per-sheet row cap (1,048,576), minus 1 for the header row.
EXCEL_MAX_DATA_ROWS = 1_048_576 - 1


# ---------------------------------------------------------------------------
# PDF parsing
# ---------------------------------------------------------------------------

def get_lines(page, y_tol=3.0):
    """
    Group a page's words into reading-order lines by y-position.

    Deliberately ignores PyMuPDF's own block/line grouping: many
    report-generated PDFs (Crystal Reports, SSRS, etc.) place every table
    cell as an independent text object, so words on the same visual row
    can land in different blocks/lines. Clustering by y0 proximity instead
    reconstructs the visual row regardless of how the PDF's content
    stream is structured.
    """
    words = sorted(page.get_text("words"), key=lambda w: (w[1], w[0]))  # x0,y0,x1,y1,text,block,line,word_no

    lines = []
    for x0, y0, x1, y1, text, *_rest in words:
        for line in lines:
            if abs(line["y0"] - y0) <= y_tol:
                line["words"].append((x0, y0, x1, y1, text))
                break
        else:
            lines.append({"y0": y0, "words": [(x0, y0, x1, y1, text)]})

    for line in lines:
        line["words"].sort(key=lambda w: w[0])
    lines.sort(key=lambda l: l["y0"])
    return lines


def _norm(text):
    """Lowercase, alnum-only form of a word -- for matching OCR'd header
    tokens regardless of stray punctuation the OCR engine may introduce."""
    return re.sub(r"[^a-z0-9]", "", text.lower())


# OCR sometimes misreads "ID" as "1D"/"lD" (I/1/l confusion).
ID_TOKEN_ALIASES = ("id", "1d", "ld")


def find_header_columns(words):
    """
    If this line looks like a table header, return the ID / Name column
    x-ranges to use for subsequent data rows. Returns None otherwise.

    Column boundaries are derived purely from the x-position of the
    matched header words themselves (this word vs. the next word on the
    line) -- there is no fixed-gap/clustering assumption. That matters
    for OCR'd (e.g. ABBYY) PDFs, whose text layer can use a completely
    different point scale, spacing, and word-segmentation behavior than
    a natively-generated PDF, which broke the earlier gap-based approach.
    """
    toks = [_norm(w[4]) for w in words]
    n = len(words)
    id_col = None    # (start_index, word_span)
    name_col = None

    i = 0
    while i < n:
        t = toks[i]
        if t in ("member", "participant") and i + 1 < n and toks[i + 1] in ID_TOKEN_ALIASES:
            id_col = (i, 2)
            i += 2
            continue
        if t in ("memberid", "participantid"):
            id_col = (i, 1)
            i += 1
            continue
        if t == "member" and i + 1 < n and toks[i + 1] == "name":
            name_col = (i, 2)
            i += 2
            continue
        if t == "membername":
            name_col = (i, 1)
            i += 1
            continue
        if t == "name":
            name_col = (i, 1)
            i += 1
            continue
        i += 1

    if id_col is None and name_col is None:
        return None

    def col_range(idx, span):
        left = words[idx][0] - 3
        next_idx = idx + span
        if next_idx < n:
            right = words[next_idx][0] - 1
        else:
            right = words[idx + span - 1][2] + 300
        return (left, right)

    result = {}
    if id_col:
        result["id"] = col_range(*id_col)
    if name_col:
        result["name"] = col_range(*name_col)
    return result


def extract_range_text(words, rng):
    if rng is None:
        return ""
    lo, hi = rng
    toks = [t for (x0, y0, x1, y1, t) in words if lo <= x0 < hi]
    return " ".join(toks).strip()


def process_pdf(path):
    """Return a list of {"File Name", "Page Number", "ID", "Name"} dicts."""
    records = []
    active_cols = None

    with fitz.open(path) as doc:
        for page_no in range(len(doc)):
            page = doc[page_no]
            for line in get_lines(page):
                words = line["words"]
                if not words:
                    continue
                line_text = " ".join(w[4] for w in words)

                header_cols = find_header_columns(words)
                if header_cols:
                    active_cols = header_cols
                    continue

                if active_cols is None:
                    continue

                if STOP_RE.search(line_text):
                    active_cols = None
                    continue

                id_val = extract_range_text(words, active_cols.get("id"))
                name_val = extract_range_text(words, active_cols.get("name"))
                if not id_val and not name_val:
                    continue

                records.append({
                    "File Name": path.name,
                    "Page Number": page_no + 1,
                    "ID": id_val,
                    "Name": name_val,
                })

    return records


def diagnose(pdf_path, max_pages=2):
    """
    Print the detected line/word structure for the first few pages of one
    PDF, so header-detection failures on OCR'd files can be debugged.

    Runs entirely locally -- this prints actual extracted text (including
    names/IDs) to YOUR terminal. Do not paste that raw output into chat;
    instead describe what you see structurally (e.g. "the header line
    shows as one merged word with no spaces", "page width is 2550pt",
    "HEADER MATCH never appears").
    """
    with fitz.open(pdf_path) as doc:
        print(f"File: {pdf_path}")
        print(f"Pages: {len(doc)}")
        for page_no in range(min(max_pages, len(doc))):
            page = doc[page_no]
            print(f"\n=== Page {page_no + 1} | size {page.rect.width:.1f} x {page.rect.height:.1f} pt ===")
            lines = get_lines(page)
            total_words = sum(len(l["words"]) for l in lines)
            print(f"  {len(lines)} line(s) detected, {total_words} word(s) total")
            for li, line in enumerate(lines):
                words = line["words"]
                header_cols = find_header_columns(words)
                tag = "  <-- HEADER MATCH" if header_cols else ""
                preview = " | ".join(f"'{w[4]}'@x{w[0]:.0f}" for w in words)
                print(f"  L{li:03d} y={line['y0']:7.1f} n={len(words):3d}: {preview}{tag}")


def find_pdf_files(folder):
    for root, _dirs, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(".pdf"):
                yield Path(root) / f


# ---------------------------------------------------------------------------
# Extraction driver + Excel writer
# ---------------------------------------------------------------------------

COLUMNS = ["File Name", "Page Number", "ID", "Name"]


def run_extraction(source_folder, dest_path, log, progress, status=None):
    def _status(msg):
        if status:
            status(msg)

    if not os.path.isdir(source_folder):
        raise ValueError(f"Source folder not found: {source_folder}")

    files = sorted(find_pdf_files(source_folder))
    if not files:
        raise ValueError("No .pdf files found in the source folder (searched recursively).")

    total_files = len(files)
    log(f"Found {total_files} PDF file(s).")
    progress(0, total_files)

    all_records = []
    for idx, path in enumerate(files, start=1):
        _status(f"Reading {path.name} ({idx}/{total_files})...")
        try:
            recs = process_pdf(path)
        except Exception as exc:
            log(f"  ! Failed to process {path.name}: {exc}")
            progress(idx, total_files)
            continue

        if recs:
            log(f"  - {path.name}: {len(recs)} row(s)")
        else:
            log(f"  - {path.name}: 0 rows (headers not recognized -- check layout)")
        all_records.extend(recs)
        progress(idx, total_files)

    if not all_records:
        raise ValueError("No rows extracted from any file.")

    total_rows = len(all_records)
    import math
    num_parts = math.ceil(total_rows / EXCEL_MAX_DATA_ROWS)
    base, ext = os.path.splitext(dest_path)
    ext = ext or ".xlsx"
    if num_parts > 1:
        log(f"{total_rows:,} rows exceed Excel's per-sheet limit of "
            f"{EXCEL_MAX_DATA_ROWS:,}; splitting output into {num_parts} files.")

    _status("Writing output workbook(s)...")
    output_paths = []
    written_total = 0
    for part_num in range(1, num_parts + 1):
        part_path = f"{base}_part{part_num}{ext}" if num_parts > 1 else dest_path
        part_start = (part_num - 1) * EXCEL_MAX_DATA_ROWS
        part_end = min(part_start + EXCEL_MAX_DATA_ROWS, total_rows)

        wb = Workbook(write_only=True)
        ws = wb.create_sheet("Extract")
        ws.append(COLUMNS)
        for rec in all_records[part_start:part_end]:
            ws.append([rec["File Name"], rec["Page Number"], rec["ID"], rec["Name"]])
            written_total += 1
            if written_total % 2000 == 0 or written_total == total_rows:
                _status(f"Writing output workbook {part_num}/{num_parts}... "
                        f"{written_total:,} / {total_rows:,} rows")

        wb.save(part_path)
        output_paths.append(part_path)

    log(f"Done. {total_rows:,} row(s) across {total_files} file(s) -> {', '.join(output_paths)}")
    return output_paths


# ---------------------------------------------------------------------------
# Tkinter GUI
# ---------------------------------------------------------------------------

DEFAULT_OUTPUT_NAME = "contrib_loan_extract.xlsx"


class ExtractorApp:
    def __init__(self, root):
        self.root = root
        root.title("Contributions/Loans PDF Extractor")
        root.geometry("700x500")
        root.resizable(True, True)

        self.source_var = tk.StringVar()
        self.dest_var = tk.StringVar()

        pad = {"padx": 8, "pady": 6}

        frame = tk.Frame(root)
        frame.pack(fill="x", **pad)

        tk.Label(frame, text="Source folder (PDFs, searched recursively):").grid(row=0, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.source_var, width=60).grid(row=0, column=1, padx=6)
        tk.Button(frame, text="Browse...", command=self.browse_source).grid(row=0, column=2)

        tk.Label(frame, text="Output Excel file:").grid(row=1, column=0, sticky="w", pady=(8, 0))
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
        self.log_box = scrolledtext.ScrolledText(root, height=18, state="disabled")
        self.log_box.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def browse_source(self):
        path = filedialog.askdirectory(title="Select folder containing PDFs")
        if path:
            self.source_var.set(path)
            if not self.dest_var.get().strip():
                self.dest_var.set(os.path.join(path, DEFAULT_OUTPUT_NAME))

    def browse_dest(self):
        path = filedialog.asksaveasfilename(
            title="Save extract as",
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            initialfile=DEFAULT_OUTPUT_NAME,
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
            output_paths = run_extraction(source, dest, self.log, self.progress, status=self.status)
            if len(output_paths) == 1:
                msg = f"Extraction complete.\nSaved to:\n{output_paths[0]}"
            else:
                msg = (
                    f"Extraction complete. Output exceeded Excel's row limit, "
                    f"so it was split into {len(output_paths)} files:\n"
                    + "\n".join(output_paths)
                )
            self.status("Done.")
            self.root.after(0, lambda: messagebox.showinfo("Done", msg))
        except Exception as exc:
            self.status(f"Failed: {exc}")
            self.log("ERROR: " + str(exc))
            self.log(traceback.format_exc())
            self.root.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            self.root.after(0, lambda: self.run_button.configure(state="normal"))


def main():
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "--diagnose":
        if len(sys.argv) < 3:
            print('Usage: python "260804 te am contrib loan pdf extractor.py" --diagnose <pdf_path> [max_pages]')
            return
        max_pages = int(sys.argv[3]) if len(sys.argv) > 3 else 2
        diagnose(sys.argv[2], max_pages)
        return

    root = tk.Tk()
    ExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
