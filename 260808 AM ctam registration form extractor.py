"""
CTAM Undergraduate Registration Form Extractor -- Tkinter GUI
================================================================
Extracts student SSN, name, address, contact info, and course/cost
details from Campbell University "Undergraduate Registration Form -
CTAM ONLY" PDFs (one form per PDF page) into a single XLSX.

Reference layout (top block, then a course table, then totals):
    Social Security Number <ssn>
    Last Nam <last>   First Name <first>   MI <mi>
    Other Names <other>
    Address <street>   <city>   <state>   <zip>
    Current Telephone <phone>   E-mail Address <email>
    -- course table header row --
    Term Code | Action Requested | Section Number | Subject Code |
    Catalog Number | Course Title | Course Credits | Student Class
    Cost | Total Cost of Class
    -- one or more course rows --
    Total Cost of Courses <amount>
    Total Cost to Student <amount>
    Total TA <amount>

Column layout for the course table is located dynamically from each
header label's x-position (not hardcoded coordinates), and the contact
block's labels are matched independently of which line they land on --
so minor template shifts should still parse. If a page lists more than
one course, that page's course-table columns are combined into ONE
output row (each column's values joined with "; "), not one row per
course. A page whose layout doesn't match well enough to find a given
field is left blank there and flagged in the "Extraction Notes" column
rather than silently producing wrong data -- treat any flagged row as
needing a manual look.

GUI:
    - Source folder picker (folder containing the registration-form
      PDFs; scanned recursively)
    - Destination folder picker (where the output XLSX goes)
    - Start Extraction button; progress prints to the console window

USAGE:
    python "260808 AM ctam registration form extractor.py"

REQUIREMENTS:
    pip install pymupdf pandas openpyxl tqdm

SECURITY NOTE:
    These forms contain SSNs and other personal data. Run only on an
    authorized workstation, and store the output only in an approved
    location -- delete the local copy once it has been loaded into the
    authorized system of record.
"""
from __future__ import annotations

import re
import threading
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pymupdf as fitz
import pandas as pd
from tqdm import tqdm

# ===========================================================================
# CONFIG
# ===========================================================================
MIN_TEXT_CHARS_PER_PAGE = 20
OUTPUT_XLSX_NAME = "ctam_registration_extracted.xlsx"

CONTACT_LABEL_VARIANTS = {
    "Social Security Number": [["Social Security Number"]],
    "Last Name": [["Last Nam"], ["Last Name"]],
    "First Name": [["First Name"]],
    "MI": [["MI"]],
    "Other Names": [["Other Names"]],
    "Address": [["Address"]],
    "Current Telephone": [["Current Telephone"]],
    "E-mail Address": [["E-mail Address"], ["Email Address"]],
}

COURSE_HEADERS = [
    "Term Code", "Action Requested", "Section Number", "Subject Code",
    "Catalog Number", "Course Title", "Course Credits",
    "Student Class Cost", "Total Cost of Class",
]

TOTALS_LABELS = ["Total Cost of Courses", "Total Cost to Student", "Total TA"]

OUTPUT_COLUMNS = (
    ["File Name", "Page Number"]
    + list(CONTACT_LABEL_VARIANTS.keys())
    + COURSE_HEADERS
    + TOTALS_LABELS
    + ["Extraction Notes"]
)

# Field-name fragments used to match AcroForm widget names (lowercased,
# non-alphanumerics stripped) when a page's filled-in values live in form
# fields rather than the flattened text layer -- used only as a fallback
# for the contact block; the course table's field-naming scheme for
# repeated rows can't be guessed generically, so it isn't attempted there.
WIDGET_NAME_ALIASES = {
    "Social Security Number": ["ssn", "socialsecuritynumber"],
    "Last Name": ["lastname", "lastnam"],
    "First Name": ["firstname"],
    "MI": ["mi", "middleinitial"],
    "Other Names": ["othername"],
    "Address": ["address"],
    "Current Telephone": ["telephone", "phone"],
    "E-mail Address": ["email"],
}


# ===========================================================================
# EXTRACTION LOGIC
# ===========================================================================
def group_words_into_lines(words, y_tol=3):
    """words: PyMuPDF page.get_text('words') output. Returns a list of
    lines, each a list of (x0, x1, text) tuples sorted left to right."""
    lines = {}
    for w in words:
        x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
        key = round(y0 / y_tol) * y_tol
        lines.setdefault(key, []).append((x0, x1, text))
    return [sorted(v, key=lambda t: t[0]) for _, v in sorted(lines.items())]


def find_label_spans(line, labels):
    """Locate each label (a list of label strings, each possibly
    multi-word) as an exact, case-insensitive token sequence on the line,
    left to right. Returns [(label, start_idx, end_idx_exclusive), ...] in
    the order the labels were actually found (not the order passed in)."""
    tokens = [t.strip(":").lower() for _, _, t in line]
    label_token_lists = [lbl.lower().split() for lbl in labels]
    n = len(tokens)
    spans = []
    i = 0
    remaining = list(zip(labels, label_token_lists))
    while i < n and remaining:
        matched_at = None
        for pos, (lbl, ltoks) in enumerate(remaining):
            tl = len(ltoks)
            if i + tl <= n and tokens[i:i + tl] == ltoks:
                matched_at = (pos, lbl, tl)
                break
        if matched_at:
            pos, lbl, tl = matched_at
            spans.append((lbl, i, i + tl))
            i += tl
            remaining = remaining[pos + 1:]
        else:
            i += 1
    return spans


def label_values(line, labels):
    """Returns (values, spans) where values[label] is the list of word
    tuples between that label and the next matched label (or end of
    line)."""
    spans = find_label_spans(line, labels)
    n = len(line)
    values = {}
    for idx, (lbl, s, e) in enumerate(spans):
        val_end = spans[idx + 1][1] if idx + 1 < len(spans) else n
        values[lbl] = line[e:val_end]
    return values, spans


def split_by_gap(value_words, gap_threshold=14):
    """Splits a value's word list into groups wherever the horizontal gap
    between words exceeds gap_threshold points -- used to recover the
    unlabeled street/city/state/zip sub-fields that make up the Address
    line (each sub-field is visually separated by blank underline space,
    not by its own caption)."""
    groups = []
    current = []
    prev_x1 = None
    for x0, x1, text in value_words:
        if prev_x1 is not None and (x0 - prev_x1) > gap_threshold:
            if current:
                groups.append(current)
            current = []
        current.append((x0, x1, text))
        prev_x1 = x1
    if current:
        groups.append(current)
    return groups


def normalize_field_name(name):
    return re.sub(r"[^a-z0-9]", "", (name or "").lower())


def widget_field_values(page):
    """AcroForm fallback for the contact block only (see WIDGET_NAME_ALIASES)."""
    values = {}
    try:
        widgets = list(page.widgets() or [])
    except Exception:
        return values
    for w in widgets:
        name = normalize_field_name(getattr(w, "field_name", "") or "")
        val = (getattr(w, "field_value", "") or "").strip()
        if not name or not val:
            continue
        for out_col, aliases in WIDGET_NAME_ALIASES.items():
            if any(a in name for a in aliases):
                values.setdefault(out_col, val)
                break
    return values


def parse_contact_block(lines):
    result = {k: "" for k in CONTACT_LABEL_VARIANTS}
    for line in lines:
        active = []
        for out_key, variants in CONTACT_LABEL_VARIANTS.items():
            if result[out_key]:
                continue
            for variant in variants:
                lbl = " ".join(variant)
                spans = find_label_spans(line, [lbl])
                if spans:
                    active.append((out_key, lbl, spans[0][1]))
                    break
        if not active:
            continue
        active.sort(key=lambda item: item[2])
        labels_in_order = [lbl for _, lbl, _ in active]
        vals, _ = label_values(line, labels_in_order)
        for out_key, lbl, _ in active:
            value_words = vals.get(lbl, [])
            if out_key == "Address":
                groups = split_by_gap(value_words)
                result[out_key] = ", ".join(" ".join(t for _, _, t in g) for g in groups)
            else:
                result[out_key] = " ".join(t for _, _, t in value_words)
    notes = [f"'{k}' not found/blank" for k, v in result.items() if not v]
    return result, notes


def find_header_line_idx(lines):
    for idx, line in enumerate(lines):
        text = " ".join(t for _, _, t in line)
        if "Term Code" in text and "Action Requested" in text:
            return idx
    return None


def header_column_ranges(header_line, page_width):
    _, spans = label_values(header_line, COURSE_HEADERS)
    ranges = {}
    for idx, (lbl, s, e) in enumerate(spans):
        x_start = header_line[s][0] - 3
        x_end = header_line[spans[idx + 1][1]][0] - 1 if idx + 1 < len(spans) else page_width
        ranges[lbl] = (x_start, x_end)
    return ranges


def bucket_row(line, col_ranges):
    buckets = {lbl: [] for lbl in col_ranges}
    for x0, x1, text in line:
        for lbl, (xs, xe) in col_ranges.items():
            if xs <= x0 < xe:
                buckets[lbl].append(text)
                break
    return {lbl: " ".join(v) for lbl, v in buckets.items()}


def parse_course_table(lines, page_width):
    header_idx = find_header_line_idx(lines)
    if header_idx is None:
        return {h: "" for h in COURSE_HEADERS}, ["course table header row not found"]
    header_line = lines[header_idx]
    col_ranges = header_column_ranges(header_line, page_width)
    if len(col_ranges) < len(COURSE_HEADERS):
        missing = [h for h in COURSE_HEADERS if h not in col_ranges]
        return {h: "" for h in COURSE_HEADERS}, [f"course table header missing column(s): {', '.join(missing)}"]

    rows = []
    for line in lines[header_idx + 1:]:
        text = " ".join(t for _, _, t in line).strip()
        if not text:
            continue
        if text.startswith("Total Cost of Courses"):
            break
        rows.append(bucket_row(line, col_ranges))

    if not rows:
        return {h: "" for h in COURSE_HEADERS}, ["no course row(s) found below header"]

    combined = {h: "; ".join(r[h] for r in rows if r.get(h)) for h in COURSE_HEADERS}
    return combined, []


def parse_totals(lines):
    result = {lbl: "" for lbl in TOTALS_LABELS}
    for line in lines:
        text = " ".join(t for _, _, t in line).strip()
        for lbl in TOTALS_LABELS:
            if text.startswith(lbl):
                vals, _ = label_values(line, [lbl])
                result[lbl] = " ".join(t for _, _, t in vals.get(lbl, []))
    notes = [f"'{lbl}' not found" for lbl in TOTALS_LABELS if not result[lbl]]
    return result, notes


def process_page(page):
    text = page.get_text()
    if len(text.strip()) < MIN_TEXT_CHARS_PER_PAGE:
        row = {col: "" for col in OUTPUT_COLUMNS if col not in ("File Name", "Page Number")}
        row["Extraction Notes"] = ("page has little/no extractable text -- scanned/image-only page; "
                                    "OCR it first if needed")
        return row

    words = page.get_text("words")
    lines = group_words_into_lines(words)

    contact, contact_notes = parse_contact_block(lines)
    widget_vals = widget_field_values(page)
    for col, val in widget_vals.items():
        if not contact.get(col):
            contact[col] = val
            contact_notes = [n for n in contact_notes if not n.startswith(f"'{col}'")]

    course, course_notes = parse_course_table(lines, page.rect.width)
    totals, totals_notes = parse_totals(lines)

    row = {}
    row.update(contact)
    row.update(course)
    row.update(totals)
    row["Extraction Notes"] = "; ".join(contact_notes + course_notes + totals_notes)
    return row


def process_pdf(path: Path):
    rows = []
    doc = fitz.open(str(path))
    for i, page in enumerate(doc, start=1):
        try:
            row = process_page(page)
        except Exception as e:
            row = {col: "" for col in OUTPUT_COLUMNS if col not in ("File Name", "Page Number")}
            row["Extraction Notes"] = f"ERROR: {type(e).__name__}: {e}"
        row["File Name"] = path.name
        row["Page Number"] = i
        rows.append(row)
    doc.close()
    return rows


# ===========================================================================
# EXTRACTION RUNNER (called from GUI thread)
# ===========================================================================
def run_extraction(source_folder, dest_folder, status_callback):
    src = Path(source_folder)
    dst = Path(dest_folder)
    if not src.is_dir():
        status_callback("ERROR: Source folder invalid.")
        return False
    dst.mkdir(parents=True, exist_ok=True)
    output_path = dst / OUTPUT_XLSX_NAME

    print("=" * 70)
    print(f"CTAM Registration Form Extractor")
    print(f"Source:      {src}")
    print(f"Destination: {dst}")
    print("=" * 70)

    status_callback("Scanning for PDFs...")
    pdfs = sorted(src.rglob("*.pdf"))
    if not pdfs:
        print(f"No PDF files found in {src}")
        status_callback("No PDFs found in source folder.")
        return False
    print(f"Found {len(pdfs)} PDF file(s).")
    status_callback(f"Found {len(pdfs)} PDF(s). Extracting...")

    all_rows = []
    flagged = 0
    with tqdm(pdfs, desc="Extracting", unit="pdf", ncols=100) as pbar:
        for pdf_path in pbar:
            pbar.set_postfix_str(pdf_path.name)
            rows = process_pdf(pdf_path)
            for r in rows:
                if r.get("Extraction Notes"):
                    flagged += 1
            all_rows.extend(rows)
            status_callback(f"Processed {pdf_path.name} ({len(all_rows)} row(s) so far)")

    df = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS)
    df.to_excel(output_path, index=False)

    print(f"\nDone. {len(all_rows)} row(s) from {len(pdfs)} file(s) -> {output_path}")
    if flagged:
        print(f"{flagged} row(s) have a non-empty 'Extraction Notes' -- spot-check those for missed/blank fields.")
    status_callback(f"Done. {len(all_rows)} row(s) written to {output_path.name} ({flagged} flagged for review).")
    return True


# ===========================================================================
# TKINTER GUI
# ===========================================================================
class ExtractorGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CTAM Registration Form Extractor")
        self.geometry("640x280")
        self.resizable(False, False)
        self._running = False
        self._build_widgets()

    def _build_widgets(self):
        pad = {"padx": 12, "pady": 8}

        title = ttk.Label(self, text="CTAM Registration Form Extractor",
                           font=("Segoe UI", 14, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w", **pad)

        subtitle = ttk.Label(self, text="Progress is printed to the console window.",
                              foreground="#555")
        subtitle.grid(row=1, column=0, columnspan=3, sticky="w", padx=12)

        ttk.Label(self, text="Source folder (PDFs):").grid(row=2, column=0, sticky="e", **pad)
        self.src_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.src_var, width=50).grid(row=2, column=1, sticky="we", **pad)
        ttk.Button(self, text="Browse...", command=self._pick_source).grid(row=2, column=2, **pad)

        ttk.Label(self, text="Destination folder:").grid(row=3, column=0, sticky="e", **pad)
        self.dst_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.dst_var, width=50).grid(row=3, column=1, sticky="we", **pad)
        ttk.Button(self, text="Browse...", command=self._pick_destination).grid(row=3, column=2, **pad)

        self.start_btn = ttk.Button(self, text="Start Extraction", command=self._start_clicked)
        self.start_btn.grid(row=4, column=0, columnspan=3, pady=12)

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w").grid(
            row=5, column=0, columnspan=3, sticky="we", padx=12, pady=(0, 12))

        self.columnconfigure(1, weight=1)

    def _pick_source(self):
        folder = filedialog.askdirectory(title="Select folder containing registration-form PDFs",
                                          mustexist=True)
        if folder:
            self.src_var.set(folder)
            if not self.dst_var.get():
                self.dst_var.set(folder)

    def _pick_destination(self):
        folder = filedialog.askdirectory(title="Select destination folder for the output XLSX",
                                          mustexist=False)
        if folder:
            self.dst_var.set(folder)

    def _set_status(self, msg):
        self.after(0, lambda: self.status_var.set(msg))

    def _start_clicked(self):
        if self._running:
            return
        src = self.src_var.get().strip()
        dst = self.dst_var.get().strip()

        if not src or not Path(src).is_dir():
            messagebox.showerror("Missing/invalid source", "Please select a valid source folder.")
            return
        if not dst:
            messagebox.showerror("Missing destination", "Please select a destination folder.")
            return

        self._running = True
        self.start_btn.config(state="disabled", text="Running...")
        self._set_status("Starting...")

        threading.Thread(target=self._run_in_thread, args=(src, dst), daemon=True).start()

    def _run_in_thread(self, src, dst):
        try:
            run_extraction(src, dst, self._set_status)
        except Exception as e:
            print(f"\nFATAL ERROR: {e}")
            self._set_status(f"Error: {e}")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.after(0, self._finished)

    def _finished(self):
        self._running = False
        self.start_btn.config(state="normal", text="Start Extraction")


# ===========================================================================
# ENTRY POINT
# ===========================================================================
def main():
    app = ExtractorGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
