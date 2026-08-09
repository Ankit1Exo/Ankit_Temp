"""
Transcript Extractor (two page formats) -- Tkinter GUI
==========================================================
Extracts identity fields from transcript PDFs that contain two different
page formats. The two formats are extracted INDEPENDENTLY -- nothing is
matched or merged between them -- and each recognised page produces its
own output row.

  FORMAT A -- the "XGRA / TRANSFER CREDITS" summary page. One line holds
  the name, then the DOB, Student ID and Print Date captions:

      Lastname, Firstname Middle   DOB: 14 Dec XXXX   Student ID: 1234567   Print Date: 15 Jul 2022

  Only two things are taken from this format: the Name (everything to the
  left of the "DOB:" caption, kept exactly as printed -- including the
  comma) and the Student ID.

  FORMAT B -- the "Etran Omed Only" page. The name and the address sit in
  a left-hand column; the captioned fields sit in a right-hand column at
  the same page heights:

      Mr. Firstname M. Lastname          ID Number: 1234567
      1221 Example Street                       SSN: 123-45-6789
      Apt 12-202                         Birth Date: 12/14/95
      Raleigh, NC  27606                 Birth Name:

  Name (with any Mr./Mrs./Ms./Dr. prefix removed), Street, Apt/Suite,
  City, State, Zip, SSN, Birth Date, ID Number and Birth Name are all
  taken from this format. The Apt/Suite line is optional -- when it isn't
  present the address is just street then city/state/zip.

Pages matching neither format (continuation pages, etc.) are skipped.
The course tables are not extracted.

A field that can't be located is left blank and the reason is recorded in
the "Extraction Notes" column, rather than guessing at a value -- treat
any row with a note as needing a manual look. ("Birth Name" is commonly
blank on these forms, so it is not flagged.)

GUI:
    - Source folder picker (scanned recursively for PDFs)
    - Destination folder picker (where the output XLSX goes)
    - Start Extraction button; progress prints to the console window

USAGE:
    python "260809 AM transcript extractor.py"
    python "260809 AM transcript extractor.py" --debug <pdf_or_folder> [page_number]

The --debug mode prints a per-line dump of a page with every value masked
to its digit/letter shape (#/X) and only the known form captions left
readable, so a layout problem can be diagnosed without exposing PII.

REQUIREMENTS:
    pip install pymupdf pandas openpyxl tqdm

SECURITY NOTE:
    These transcripts contain SSNs, birth dates and home addresses. Run
    only on an authorised workstation, save the output only to an
    approved location, and delete the local copy once it has been loaded
    into the authorised system of record.
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
OUTPUT_XLSX_NAME = "transcript_extracted.xlsx"

# How far below the name line to keep looking for address lines before
# giving up (the address is 2-3 lines; the extra room absorbs a blank or
# a stray line without running off down the page).
ADDRESS_SCAN_LINES = 6

# Requiring a period OR trailing whitespace matters: a bare
# "^(Mr|Mrs|Ms|Dr)" would also strike the first two letters off real
# surnames such as "Mroz" or "Drury".
HONORIFIC_RE = re.compile(r"^(?:Mr|Mrs|Ms|Miss|Dr)(?:\.\s*|\s+)", re.IGNORECASE)

CITY_STATE_ZIP_RE = re.compile(
    r"^(?P<city>.+?),?\s*(?P<state>[A-Z]{2})\.?\s+(?P<zip>\d{5}(?:-\d{4})?)\s*$"
)

APT_LINE_RE = re.compile(r"^(?:apt|apartment|suite|ste|unit|rm|room|#)\b", re.IGNORECASE)

FORMAT_B_LABELS = ["ID Number", "SSN", "Birth Date", "Birth Name"]
FORMAT_A_LABELS = ["DOB", "Student ID", "Print Date"]

OUTPUT_COLUMNS = [
    "File Name", "Page Number", "Format",
    "Name (Format A)", "Student ID (Format A)",
    "Name (Format B)", "Street", "Apt/Suite", "City", "State", "Zip",
    "SSN", "Birth Date", "ID Number", "Birth Name",
    "Extraction Notes",
]


# ===========================================================================
# LINE / LABEL HELPERS
# ===========================================================================
def group_words_into_lines(words, y_tol=3):
    """words: PyMuPDF page.get_text('words') output. Returns a list of
    lines, each a list of (x0, x1, text) tuples sorted left to right."""
    lines = {}
    for w in words:
        x0, y0, x1, text = w[0], w[1], w[2], w[4]
        key = round(y0 / y_tol) * y_tol
        lines.setdefault(key, []).append((x0, x1, text))
    return [sorted(v, key=lambda t: t[0]) for _, v in sorted(lines.items())]


def find_label_span(line, label):
    """Find `label` on `line` as a case-insensitive token sequence,
    ignoring any trailing colon. Returns (start, end_exclusive) or None.
    A multi-word label also matches its words fused into a single token
    ("IDNumber"), which PDF text extraction sometimes produces when the
    caption is tightly kerned."""
    tokens = [t.strip(":").lower() for _, _, t in line]
    ltoks = label.lower().split()
    tl = len(ltoks)
    for i in range(len(tokens)):
        if tokens[i:i + tl] == ltoks:
            return i, i + tl
        if tl > 1 and tokens[i] == "".join(ltoks):
            return i, i + 1
    return None


def label_values_on_line(line, labels):
    """{label: value} for every one of `labels` present on `line`. A
    value runs from just after its own caption to the start of the next
    caption on the same line (or to the end of the line)."""
    spans = sorted((s, e, lbl) for lbl in labels
                    for s, e in [find_label_span(line, lbl) or (None, None)] if s is not None)
    values = {}
    for idx, (_, end, lbl) in enumerate(spans):
        stop = spans[idx + 1][0] if idx + 1 < len(spans) else len(line)
        values[lbl] = " ".join(t for _, _, t in line[end:stop]).strip()
    return values


def text_left_of(line, x_limit):
    """The line's text up to (not including) x_limit -- i.e. the
    left-hand column only."""
    return " ".join(t for x0, _, t in line if x0 < x_limit).strip()


# ===========================================================================
# PAGE CLASSIFICATION
# ===========================================================================
def classify_page(lines):
    text = " ".join(t for line in lines for _, _, t in line).lower()
    if "id number" in text and "ssn" in text:
        return "B"
    if "student id" in text and "dob" in text:
        return "A"
    return None


# ===========================================================================
# FORMAT A -- Name + Student ID from the DOB caption line
# ===========================================================================
def parse_format_a(lines):
    result = {"Name (Format A)": "", "Student ID (Format A)": ""}
    notes = []

    for line in lines:
        dob_span = find_label_span(line, "DOB")
        if not dob_span or not find_label_span(line, "Student ID"):
            continue

        name = " ".join(t for _, _, t in line[:dob_span[0]]).strip()
        if name:
            result["Name (Format A)"] = name
        else:
            notes.append("Format A: no name text to the left of the 'DOB:' caption")

        values = label_values_on_line(line, FORMAT_A_LABELS)
        result["Student ID (Format A)"] = values.get("Student ID", "")
        if not result["Student ID (Format A)"]:
            notes.append("Format A: 'Student ID' caption found but its value is blank")
        return result, notes

    notes.append("Format A: no line carrying both the 'DOB' and 'Student ID' captions was found")
    return result, notes


# ===========================================================================
# FORMAT B -- Name, address and the captioned right-hand column
# ===========================================================================
def split_address(addr_lines, notes):
    """addr_lines: the left-column text of the lines below the name, in
    reading order. Returns (street, apt_suite, city, state, zip)."""
    csz_idx = next((i for i, t in enumerate(addr_lines) if CITY_STATE_ZIP_RE.match(t)), None)

    if csz_idx is None:
        notes.append("Format B: no 'City, ST ZIP'-shaped line found below the name -- "
                      "City/State/Zip left blank")
        street = addr_lines[0] if addr_lines else ""
        apt = " ".join(addr_lines[1:])
        return street, apt, "", "", ""

    m = CITY_STATE_ZIP_RE.match(addr_lines[csz_idx])
    city, state, zip_code = m.group("city").rstrip(",").strip(), m.group("state"), m.group("zip")

    before = addr_lines[:csz_idx]
    if not before:
        notes.append("Format B: city/state/zip line found but no street line above it")
        return "", "", city, state, zip_code

    # Anything between the street and the city line that doesn't look like
    # an Apt/Suite designation is folded back into Street rather than
    # mislabelled as Apt/Suite.
    street, extra = before[0], before[1:]
    apt = " ".join(t for t in extra if APT_LINE_RE.match(t))
    trailing = [t for t in extra if not APT_LINE_RE.match(t)]
    if trailing:
        street = " ".join([street] + trailing)
    return street, apt, city, state, zip_code


def parse_format_b(lines):
    result = {k: "" for k in ["Name (Format B)", "Street", "Apt/Suite", "City", "State", "Zip",
                               "SSN", "Birth Date", "ID Number", "Birth Name"]}
    notes = []

    for line in lines:
        for lbl, value in label_values_on_line(line, FORMAT_B_LABELS).items():
            if not result[lbl]:
                result[lbl] = value

    # The name shares a line with the "ID Number" caption, and the address
    # lines below sit in that same left-hand column. Cutting every line at
    # the caption's x keeps right-column values out of the name/address
    # whether or not a given address line happens to have a caption beside
    # it -- more reliable than assuming each line carries one.
    anchor_idx, label_x = None, None
    for idx, line in enumerate(lines):
        span = find_label_span(line, "ID Number")
        if span:
            anchor_idx, label_x = idx, line[span[0]][0]
            break

    if anchor_idx is None:
        notes.append("Format B: 'ID Number' caption not found -- name and address could not be located")
        return result, notes

    name_text = text_left_of(lines[anchor_idx], label_x)
    result["Name (Format B)"] = HONORIFIC_RE.sub("", name_text).strip()
    if not result["Name (Format B)"]:
        notes.append("Format B: no name text to the left of the 'ID Number' caption")

    addr_lines = []
    for idx in range(anchor_idx + 1, min(anchor_idx + 1 + ADDRESS_SCAN_LINES, len(lines))):
        text = text_left_of(lines[idx], label_x)
        if not text:
            continue
        addr_lines.append(text)
        if CITY_STATE_ZIP_RE.match(text):
            break

    if not addr_lines:
        notes.append("Format B: no address lines found below the name")
    else:
        (result["Street"], result["Apt/Suite"], result["City"],
         result["State"], result["Zip"]) = split_address(addr_lines, notes)

    # "Birth Name" is routinely blank on these forms, so it isn't flagged.
    for key in ["SSN", "Birth Date", "ID Number"]:
        if not result[key]:
            notes.append(f"Format B: '{key}' not found/blank")

    return result, notes


# ===========================================================================
# PER-PDF DRIVER -- one row per recognised page
# ===========================================================================
def process_pdf(path: Path):
    doc = fitz.open(str(path))
    rows = []
    counts = {"A": 0, "B": 0}

    for page_num, page in enumerate(doc, start=1):
        text = page.get_text()
        if len(text.strip()) < MIN_TEXT_CHARS_PER_PAGE:
            continue

        lines = group_words_into_lines(page.get_text("words"))
        kind = classify_page(lines)
        if kind is None:
            continue
        counts[kind] += 1

        try:
            values, notes = parse_format_a(lines) if kind == "A" else parse_format_b(lines)
        except Exception as e:
            values, notes = {}, [f"ERROR: {type(e).__name__}: {e}"]

        row = {col: "" for col in OUTPUT_COLUMNS}
        row.update(values)
        row["File Name"] = path.name
        row["Page Number"] = page_num
        row["Format"] = kind
        row["Extraction Notes"] = "; ".join(notes)
        rows.append(row)

    doc.close()
    return rows, counts


# ===========================================================================
# DEBUG (safe, PII-free diagnostic dump)
# ===========================================================================
KNOWN_DEBUG_LABELS = [
    "DOB", "Student ID", "Print Date", "TRANSFER CREDITS", "Hours Attempted", "Hours Passed",
    "ID Number", "SSN", "Birth Date", "Birth Name", "Etran Omed Only", "Course", "Title",
]


def mask_shape(s):
    return re.sub(r"[A-Za-z]", "X", re.sub(r"\d", "#", s))


def mask_line_except_labels(line):
    tokens = [t for _, _, t in line]
    lowered = [t.strip(":").lower() for t in tokens]
    is_label = [False] * len(tokens)
    for lbl in sorted(KNOWN_DEBUG_LABELS, key=len, reverse=True):
        ltoks = lbl.lower().split()
        tl = len(ltoks)
        for i in range(len(tokens) - tl + 1):
            if not any(is_label[i:i + tl]) and lowered[i:i + tl] == ltoks:
                for k in range(i, i + tl):
                    is_label[k] = True
    return " ".join(t if is_label[k] else mask_shape(t) for k, t in enumerate(tokens))


def debug_page(path: Path, page_num: int):
    doc = fitz.open(str(path))
    if page_num < 1 or page_num > len(doc):
        print(f"{path.name}: page {page_num} out of range (document has {len(doc)} page(s))")
        doc.close()
        return

    page = doc[page_num - 1]
    text = page.get_text()
    print(f"--- {path.name} page {page_num} ---")
    print(f"text layer: {len(text.strip())} chars "
          f"({'OK' if len(text.strip()) >= MIN_TEXT_CHARS_PER_PAGE else 'BELOW MIN -- image-only page'})")

    lines = group_words_into_lines(page.get_text("words"))
    kind = classify_page(lines)
    print(f"page format: {kind or 'UNRECOGNISED (neither Format A nor Format B captions found)'}")

    if kind:
        values, notes = parse_format_a(lines) if kind == "A" else parse_format_b(lines)
        for key, value in values.items():
            print(f"  {key}: {'(found)' if value else '(blank)'}")
        if notes:
            print(f"  notes: {'; '.join(notes)}")

    print(f"lines detected: {len(lines)}")
    for i, line in enumerate(lines):
        print(f"  [{i:>3}] {mask_line_except_labels(line)}")
    doc.close()


# ===========================================================================
# EXTRACTION RUNNER (called from the GUI thread)
# ===========================================================================
def run_extraction(source_folder, dest_folder, status_callback):
    src, dst = Path(source_folder), Path(dest_folder)
    if not src.is_dir():
        status_callback("ERROR: Source folder invalid.")
        return False
    dst.mkdir(parents=True, exist_ok=True)
    output_path = dst / OUTPUT_XLSX_NAME

    print("=" * 70)
    print("Transcript Extractor")
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

    all_rows, flagged, no_pages = [], 0, []
    with tqdm(pdfs, desc="Extracting", unit="pdf", ncols=100) as pbar:
        for pdf_path in pbar:
            pbar.set_postfix_str(pdf_path.name)
            rows, counts = process_pdf(pdf_path)
            if not rows:
                no_pages.append(pdf_path.name)
            flagged += sum(1 for r in rows if r["Extraction Notes"])
            all_rows.extend(rows)
            status_callback(f"Processed {pdf_path.name} ({len(all_rows)} row(s) so far)")

    if not all_rows:
        print("No Format A or Format B pages found in any file -- nothing to write.")
        status_callback("Done. No recognised pages found -- nothing written.")
        return False

    pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS).to_excel(output_path, index=False)

    fmt_a = sum(1 for r in all_rows if r["Format"] == "A")
    fmt_b = sum(1 for r in all_rows if r["Format"] == "B")
    print(f"\nDone. {len(all_rows)} row(s) ({fmt_a} Format A, {fmt_b} Format B) "
          f"from {len(pdfs)} file(s) -> {output_path}")
    if flagged:
        print(f"{flagged} row(s) have a non-empty 'Extraction Notes' -- spot-check those.")
    if no_pages:
        print(f"{len(no_pages)} file(s) had no recognised page: {', '.join(no_pages)}")
    status_callback(f"Done. {len(all_rows)} row(s) written to {output_path.name} "
                    f"({flagged} flagged for review).")
    return True


# ===========================================================================
# TKINTER GUI
# ===========================================================================
class ExtractorGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Transcript Extractor")
        self.geometry("640x280")
        self.resizable(False, False)
        self._running = False
        self._build_widgets()

    def _build_widgets(self):
        pad = {"padx": 12, "pady": 8}

        ttk.Label(self, text="Transcript Extractor",
                   font=("Segoe UI", 14, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(self, text="Progress is printed to the console window.",
                   foreground="#555").grid(row=1, column=0, columnspan=3, sticky="w", padx=12)

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
        folder = filedialog.askdirectory(title="Select folder containing transcript PDFs", mustexist=True)
        if folder:
            self.src_var.set(folder)
            if not self.dst_var.get():
                self.dst_var.set(folder)

    def _pick_destination(self):
        folder = filedialog.askdirectory(title="Select destination folder for the output XLSX", mustexist=False)
        if folder:
            self.dst_var.set(folder)

    def _set_status(self, msg):
        self.after(0, lambda: self.status_var.set(msg))

    def _start_clicked(self):
        if self._running:
            return
        src, dst = self.src_var.get().strip(), self.dst_var.get().strip()
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
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "--debug":
        if len(sys.argv) < 3:
            print("Usage: --debug <pdf_file_or_folder> [page_number]")
            return
        target = Path(sys.argv[2])
        page_num = int(sys.argv[3]) if len(sys.argv) > 3 else 1
        pdf_files = [target] if target.is_file() else sorted(target.rglob("*.pdf"))
        if not pdf_files:
            print(f"No PDF files found at: {target}")
            return
        for pdf in pdf_files:
            debug_page(pdf, page_num)
        return

    ExtractorGUI().mainloop()


if __name__ == "__main__":
    main()
