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

Contributions rows are also validated against their expected value shape:
Member ID (a plain number), Member Name (contains letters), ER Match in K
(a decimal), Elect Deferral (a decimal). If overlapping/overwritten OCR
text breaks that sequence for a given line, only that line is skipped --
extraction continues with the next line rather than emitting bad data or
aborting the table. Loans rows have no amount columns to validate against,
so they keep the more tolerant blank-Participant-ID-is-normal behavior.

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

# Excel's hard per-sheet row cap (1,048,576), minus 1 for the header row.
EXCEL_MAX_DATA_ROWS = 1_048_576 - 1


# ---------------------------------------------------------------------------
# PDF parsing
# ---------------------------------------------------------------------------

def get_lines(page, y_tol=5.0):
    """
    Group a page's words into reading-order lines by y-position. Used only
    for HEADER detection (the words within one header cell, e.g. "Member" +
    "ID", are typeset on the same physical baseline so a modest tolerance
    is safe here).

    Deliberately ignores PyMuPDF's own block/line grouping: many
    report-generated / OCR'd PDFs place every table cell as an independent
    text object, so words on the same visual row can land in different
    blocks/lines as PyMuPDF sees them. Clustering by y0 proximity instead
    reconstructs the visual row regardless of how the underlying text
    layer is structured.

    NOTE: this whole-line approach is NOT used for data rows. On OCR'd
    (e.g. ABBYY) PDFs, different columns can drift vertically relative to
    each other (scan skew / independent per-column OCR), so a data row's
    ID and Name can land several points apart in y -- see
    cluster_words_by_y() + extract_segment() below, which cluster each
    column independently and then pair rows by nearest y instead of
    requiring an exact match.
    """
    words = sorted(page.get_text("words"), key=lambda w: (w[1], w[0]))  # x0,y0,x1,y1,text,block,line,word_no

    lines = []
    for x0, y0, x1, y1, text, *_rest in words:
        for line in lines:
            if abs(line["y0"] - y0) <= y_tol:
                line["words"].append((x0, y0, x1, y1, text))
                line["y1"] = max(line["y1"], y1)
                break
        else:
            lines.append({"y0": y0, "y1": y1, "words": [(x0, y0, x1, y1, text)]})

    for line in lines:
        line["words"].sort(key=lambda w: w[0])
    lines.sort(key=lambda l: l["y0"])
    return lines


def cluster_words_by_y(words, y_tol=4.0):
    """
    Cluster a set of words -- already filtered to one column's x-range --
    into rows by y-position, independent of any other column. Returns a
    list of {"y0", "text"} sorted top to bottom.
    """
    clusters = []
    for w in sorted(words, key=lambda w: w[1]):
        for c in clusters:
            if abs(c["y0"] - w[1]) <= y_tol:
                c["words"].append(w)
                break
        else:
            clusters.append({"y0": w[1], "words": [w]})

    out = []
    for c in clusters:
        c["words"].sort(key=lambda w: w[0])
        out.append({"y0": c["y0"], "text": " ".join(w[4] for w in c["words"]).strip()})
    out.sort(key=lambda c: c["y0"])
    return out


def _norm(text):
    """Lowercase, alnum-only form of a word -- for matching OCR'd header
    tokens regardless of stray punctuation the OCR engine may introduce."""
    return re.sub(r"[^a-z0-9]", "", text.lower())


# OCR sometimes misreads "ID" as "1D"/"lD" (I/1/l confusion).
ID_TOKEN_ALIASES = ("id", "1d", "ld")


def find_header_columns(words):
    """
    If this line looks like a table header, return the column x-ranges to
    use for subsequent data rows. Returns None otherwise.

    Recognizes: Member ID / Participant ID -> "id"
                Member Name / Name         -> "name"
                ER Match in K              -> "er_match"
                Elect Deferral             -> "elect_deferral"

    The last two are only present on the Contributions header and are
    used to validate the expected number-text-decimal-decimal row
    sequence (see extract_segment's strict mode) -- they are not part of
    the output columns.

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
    er_col = None
    ed_col = None

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
        if t == "er" and i + 3 < n and toks[i + 1] == "match" and toks[i + 2] == "in" and toks[i + 3] == "k":
            er_col = (i, 4)
            i += 4
            continue
        if t in ("ermatchink", "ermatch"):
            er_col = (i, 1)
            i += 1
            continue
        if t == "elect" and i + 1 < n and toks[i + 1] == "deferral":
            ed_col = (i, 2)
            i += 2
            continue
        if t == "electdeferral":
            ed_col = (i, 1)
            i += 1
            continue
        i += 1

    if id_col is None and name_col is None and er_col is None and ed_col is None:
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
    if er_col:
        result["er_match"] = col_range(*er_col)
    if ed_col:
        result["elect_deferral"] = col_range(*ed_col)
    return result


def _is_stop_word(text):
    """True for a word that marks the end of a table segment (Subtotals)."""
    return _norm(text) in ("subtotal", "subtotals")


def _mk_record(file_name, page_no, id_val, name_val):
    return {"File Name": file_name, "Page Number": page_no, "ID": id_val, "Name": name_val}


# Expected value "shape" for a valid Contributions row: Member ID (a plain
# number), Member Name (contains letters), ER Match in K (a decimal),
# Elect Deferral (a decimal). Used only when both amount columns were
# found on the header -- i.e. Contributions-style rows, not Loans rows.
ID_NUMERIC_RE = re.compile(r"^\d+$")
NAME_TEXT_RE = re.compile(r"[A-Za-z]")
DECIMAL_RE = re.compile(r"^-?\d[\d,]*\.\d+$")


def _looks_like_id(text):
    return bool(ID_NUMERIC_RE.match(text.replace(" ", "")))


def _looks_like_name(text):
    return bool(NAME_TEXT_RE.search(text))


def _looks_like_decimal(text):
    return bool(DECIMAL_RE.match(text.strip()))


def _estimate_row_height(*row_lists):
    non_empty = [rows for rows in row_lists if rows]
    if not non_empty:
        return 14.0
    base = max(non_empty, key=len)
    ys = sorted(r["y0"] for r in base)
    gaps = [b - a for a, b in zip(ys, ys[1:]) if b - a > 1]
    return sorted(gaps)[len(gaps) // 2] if gaps else 14.0


def _nearest_unused(rows, y, tol, used):
    """Index of the row in `rows` closest to y (within tol) not already in `used`, or None."""
    best_idx, best_dy = None, None
    for idx, r in enumerate(rows):
        if idx in used:
            continue
        dy = abs(r["y0"] - y)
        if dy <= tol and (best_dy is None or dy < best_dy):
            best_idx, best_dy = idx, dy
    return best_idx


def _extract_lenient_rows(id_rows, name_rows, file_name, page_no):
    """
    Pair ID/Name rows by nearest y, tolerating a blank on either side
    (e.g. Loans rows commonly have no Participant ID at all).
    """
    pair_tol = max(_estimate_row_height(id_rows, name_rows) * 0.6, 6.0)

    records = []
    i, j = 0, 0
    while i < len(id_rows) or j < len(name_rows):
        if i < len(id_rows) and j < len(name_rows):
            dy = id_rows[i]["y0"] - name_rows[j]["y0"]
            if abs(dy) <= pair_tol:
                records.append(_mk_record(file_name, page_no, id_rows[i]["text"], name_rows[j]["text"]))
                i += 1
                j += 1
            elif dy < 0:
                records.append(_mk_record(file_name, page_no, id_rows[i]["text"], ""))
                i += 1
            else:
                records.append(_mk_record(file_name, page_no, "", name_rows[j]["text"]))
                j += 1
        elif i < len(id_rows):
            records.append(_mk_record(file_name, page_no, id_rows[i]["text"], ""))
            i += 1
        else:
            records.append(_mk_record(file_name, page_no, "", name_rows[j]["text"]))
            j += 1

    return [r for r in records if r["ID"] or r["Name"]]


def _extract_strict_rows(id_rows, name_rows, er_rows, ed_rows, file_name, page_no, stats=None):
    """
    Only emit a row when the full Member ID / Member Name / ER Match in K /
    Elect Deferral sequence is present AND each value has the expected
    shape (number - text - decimal - decimal). Anchors on the ID rows
    (Member ID is expected on every valid Contributions line); if any of
    the other three columns is missing nearby, or a matched value doesn't
    look like the expected type -- e.g. overlapping/overwritten text
    scrambled it -- that single line is skipped and the next ID row is
    tried, rather than emitting bad data or aborting the segment.
    """
    pair_tol = max(_estimate_row_height(id_rows, name_rows, er_rows, ed_rows) * 0.6, 6.0)

    records = []
    used_name, used_er, used_ed = set(), set(), set()
    for id_row in id_rows:
        y = id_row["y0"]
        ni = _nearest_unused(name_rows, y, pair_tol, used_name)
        ei = _nearest_unused(er_rows, y, pair_tol, used_er)
        di = _nearest_unused(ed_rows, y, pair_tol, used_ed)

        if ni is None or ei is None or di is None:
            if stats is not None:
                stats["rows_skipped"] += 1
            continue

        id_text = id_row["text"]
        name_text = name_rows[ni]["text"]
        er_text = er_rows[ei]["text"]
        ed_text = ed_rows[di]["text"]

        if not (_looks_like_id(id_text) and _looks_like_name(name_text)
                and _looks_like_decimal(er_text) and _looks_like_decimal(ed_text)):
            if stats is not None:
                stats["rows_skipped"] += 1
            continue

        used_name.add(ni)
        used_er.add(ei)
        used_ed.add(di)
        records.append(_mk_record(file_name, page_no, id_text, name_text))

    return records


def extract_segment(words, active, y_lo, y_hi, file_name, page_no, stats=None):
    """
    Extract ID/Name rows from the words falling between y_lo and y_hi
    (exclusive), using the column x-ranges in `active`.

    Every relevant column's words are clustered into rows INDEPENDENTLY
    (see cluster_words_by_y) rather than requiring them to share a line --
    this is what tolerates per-column vertical drift on OCR'd PDFs, where
    the ID entry and Name entry for the same row can land several points
    apart in y.

    When the header also carried "ER Match in K" and "Elect Deferral"
    (Contributions rows), strict-mode validation is applied: a row is only
    kept if all four values are present nearby AND match the expected
    number/text/decimal/decimal shape (see _extract_strict_rows). Loans
    rows (no amount columns tracked) keep the earlier lenient pairing,
    since a blank Participant ID there is normal, not an error.
    """
    id_range = active.get("id")
    name_range = active.get("name")
    er_range = active.get("er_match")
    ed_range = active.get("elect_deferral")

    seg_words = [w for w in words if y_lo < w[1] < y_hi]

    def words_in(rng):
        if not rng:
            return []
        lo, hi = rng
        return [w for w in seg_words if lo <= w[0] < hi]

    id_rows = cluster_words_by_y(words_in(id_range)) if id_range else []
    name_rows = cluster_words_by_y(words_in(name_range)) if name_range else []

    if er_range and ed_range:
        er_rows = cluster_words_by_y(words_in(er_range))
        ed_rows = cluster_words_by_y(words_in(ed_range))
        return _extract_strict_rows(id_rows, name_rows, er_rows, ed_rows, file_name, page_no, stats)

    return _extract_lenient_rows(id_rows, name_rows, file_name, page_no)


def process_pdf(path):
    """
    Return (records, stats).

    records: list of {"File Name", "Page Number", "ID", "Name"} dicts.
    stats: counts useful for diagnosing extraction issues without exposing
    any PII -- header/stop-marker counts, and how many output rows came
    out fully paired vs. missing one side.
    """
    records = []
    active = None
    stats = {"pages": 0, "headers_found": 0, "stop_markers_found": 0, "rows_skipped": 0}

    with fitz.open(path) as doc:
        stats["pages"] = len(doc)
        for page_no in range(len(doc)):
            page = doc[page_no]
            all_words = sorted(page.get_text("words"), key=lambda w: (w[1], w[0]))

            events = []  # (y, "header"|"stop", payload)
            for line in get_lines(page):
                cols = find_header_columns(line["words"])
                if cols:
                    events.append((line["y0"], "header", (cols, line["y1"])))
                    stats["headers_found"] += 1
            for w in all_words:
                if _is_stop_word(w[4]):
                    events.append((w[1], "stop", None))
                    stats["stop_markers_found"] += 1
            events.sort(key=lambda e: e[0])

            # A "stop" marker (Subtotals) only chunks the segment so the
            # row-height estimate stays local to each block of rows -- it
            # does NOT clear `active`. Some reports repeat a subtotal line
            # mid-table (e.g. a running/page subtotal) before continuing
            # the same table; only a genuinely new header line should
            # redefine the column mapping. Any trailing footer text after
            # the true last table on a page won't produce spurious rows
            # since it has no words positioned inside the ID/Name columns.
            cursor_y = 0.0
            for ev_y, kind, payload in events:
                if active:
                    records.extend(extract_segment(all_words, active, cursor_y, ev_y, path.name, page_no + 1, stats))
                if kind == "header":
                    active, header_y1 = payload
                    cursor_y = header_y1
                else:
                    cursor_y = ev_y

            if active:
                records.extend(extract_segment(all_words, active, cursor_y, float("inf"), path.name, page_no + 1, stats))

    return records, stats


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
            recs, stats = process_pdf(path)
        except Exception as exc:
            log(f"  ! Failed to process {path.name}: {exc}")
            progress(idx, total_files)
            continue

        if recs:
            paired = sum(1 for r in recs if r["ID"] and r["Name"])
            id_only = sum(1 for r in recs if r["ID"] and not r["Name"])
            name_only = sum(1 for r in recs if r["Name"] and not r["ID"])
            log(f"  - {path.name}: {len(recs)} row(s) "
                f"[{paired} paired, {id_only} ID-only, {name_only} Name-only, "
                f"{stats['rows_skipped']} skipped as malformed] -- "
                f"{stats['headers_found']} header(s), {stats['stop_markers_found']} Subtotal marker(s), "
                f"{stats['pages']} page(s)")
        else:
            log(f"  - {path.name}: 0 rows -- {stats['headers_found']} header(s) recognized, "
                f"{stats['pages']} page(s) (headers not recognized if 0 -- check layout)")
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
