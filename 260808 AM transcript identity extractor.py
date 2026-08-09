"""
Academic Transcript Identity Extractor -- Tkinter GUI
========================================================
Extracts identity fields from PDFs that contain TWO different transcript
layouts as different pages of the same file. The two layouts are NOT
cross-matched against each other -- each is extracted independently into
its own set of columns, one output row per page:

  Layout A ("summary"): a dotted-rule course listing with a "DOB:" /
  "Student ID:" / "Print Date:" header line. The name is printed on that
  SAME line, before the "DOB:" caption (e.g. "Lastname, Firstname
  Middle DOB: ... Student ID: ... Print Date: ..."). ONLY the name is
  extracted from this layout, exactly as printed -- no reordering, no
  other fields -- into the "Name (Summary Page)" column.

  Layout B ("Etran Omed Only"): "Mr./Ms./Mrs./Dr. <Name>" followed by a
  2-3 line address (street, optional "Apt ..." line, then city/state/
  zip), with "ID Number:", "SSN:", "Birth Date:", and "Birth Name:"
  printed alongside those same rows in a column to the right. Name,
  Address, ID Number, SSN, Birth Date, and Birth Name are all extracted
  from this layout into their own columns (plus Street/Apt-Suite/City/
  State/Zip and First/Middle/Last splits of this layout's own name).

The course table itself is NOT extracted by this script (identity fields
only, per what was asked for).

GUI:
    - Source folder picker (scanned recursively for PDFs)
    - Destination folder picker (where the output XLSX goes)
    - Start Extraction button; progress prints to the console window

USAGE:
    python "260808 AM transcript identity extractor.py"
    python "260808 AM transcript identity extractor.py" --debug <pdf_or_folder> [page_number]

REQUIREMENTS:
    pip install pymupdf pandas openpyxl tqdm

SECURITY NOTE:
    These transcripts contain SSNs, birth dates, and home addresses. Run
    only on an authorized workstation, and store the output only in an
    approved location -- delete the local copy once it has been loaded
    into the authorized system of record.
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
OUTPUT_XLSX_NAME = "transcript_identity_extracted.xlsx"

HONORIFIC_RE = re.compile(r"^(Mr|Mrs|Ms|Miss|Dr)\.?\s*", re.IGNORECASE)
CITY_STATE_ZIP_RE = re.compile(r"^(?P<city>.+?),?\s*(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)\b")
APT_LINE_RE = re.compile(r"^(apt|suite|unit|ste)\b", re.IGNORECASE)

IDENTITY_LABELS_B = {
    "ID Number": [["ID Number"]],
    "SSN": [["SSN"]],
    "Birth Date": [["Birth Date"]],
    "Birth Name": [["Birth Name"]],
}
ALL_LABEL_STRINGS_B = [" ".join(v) for variants in IDENTITY_LABELS_B.values() for v in variants]

OUTPUT_COLUMNS = [
    "File Name", "Page Number",
    "Name (Summary Page)",
    "Name (As Printed)", "First Name", "Middle Name", "Last Name",
    "Address", "Street", "Apt/Suite", "City", "State", "Zip",
    "ID Number", "SSN", "Birth Date", "Birth Name", "Extraction Notes",
]


# ===========================================================================
# SHARED HELPERS (same approach as the CTAM extractor in this repo)
# ===========================================================================
def group_words_into_lines(words, y_tol=3):
    lines = {}
    for w in words:
        x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
        key = round(y0 / y_tol) * y_tol
        lines.setdefault(key, []).append((x0, x1, text))
    return [sorted(v, key=lambda t: t[0]) for _, v in sorted(lines.items())]


def find_label_spans(line, labels):
    """Case-insensitive token-sequence label matcher -- see the CTAM
    extractor for the full rationale (also tolerates a multi-word label
    landing as one fused token, e.g. "IDNumber" for "ID Number")."""
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
            if tl > 1 and tokens[i] == "".join(ltoks):
                matched_at = (pos, lbl, 1)
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
    spans = find_label_spans(line, labels)
    n = len(line)
    values = {}
    for idx, (lbl, s, e) in enumerate(spans):
        val_end = spans[idx + 1][1] if idx + 1 < len(spans) else n
        values[lbl] = line[e:val_end]
    return values, spans


def left_of_first_label(line, label_strings):
    """Everything on the line BEFORE the first of these labels starts --
    used to strip the right-hand-column identity fields (ID Number/SSN/
    Birth Date/Birth Name) off a line that also carries left-hand-column
    Name/Address text at the same page height."""
    spans = find_label_spans(line, label_strings)
    if not spans:
        return line
    first_start = min(s for _, s, _ in spans)
    return line[:first_start]


def split_name_parts(name_text):
    """Best-effort "First [Middle...] Last" split. 1 token -> First only;
    2 tokens -> First/Last, Middle blank; 3+ tokens -> first token is
    First, last token is Last, everything between is Middle."""
    tokens = name_text.split()
    if not tokens:
        return "", "", ""
    if len(tokens) == 1:
        return tokens[0], "", ""
    if len(tokens) == 2:
        return tokens[0], "", tokens[1]
    return tokens[0], " ".join(tokens[1:-1]), tokens[-1]


def split_address_parts(addr_lines):
    """addr_lines: the raw (unjoined) address line strings collected for
    one record, in reading order -- the last one is expected to be the
    city/state/zip line, an optional Apt/Suite line may sit between the
    street line and it. Returns (street, apt_suite, city, state, zip)."""
    street, apt, city, state, zip_code = "", "", "", "", ""
    remaining = list(addr_lines)
    if remaining:
        m = CITY_STATE_ZIP_RE.match(remaining[-1])
        if m:
            city = m.group("city").rstrip(",").strip()
            state = m.group("state")
            zip_code = m.group("zip")
            remaining = remaining[:-1]
    if remaining:
        street = remaining[0]
        for line in remaining[1:]:
            if APT_LINE_RE.match(line.strip()):
                apt = f"{apt} {line}".strip() if apt else line
            else:
                # doesn't look like an Apt/Suite line -- fold it into
                # Street rather than mislabeling it as Apt/Suite
                street = f"{street} {line}".strip()
    return street, apt, city, state, zip_code


# ===========================================================================
# PAGE CLASSIFICATION
# ===========================================================================
def classify_page(lines):
    text_lower = " ".join(t for line in lines for _, _, t in line).lower()
    if "id number" in text_lower and "ssn" in text_lower:
        return "B"
    if "student id" in text_lower and ("dob" in text_lower or "transfer credits" in text_lower):
        return "A"
    return None


# ===========================================================================
# LAYOUT A ("summary") -- name ONLY, exactly as printed
# ===========================================================================
def parse_layout_a(lines):
    """The name on this layout is printed as "Lastname, Firstname Middle"
    on the SAME line as the DOB/Student ID/Print Date captions (confirmed
    against a real file's masked debug output), before the "DOB:" label
    starts. Kept exactly as printed -- no reordering, no other fields
    extracted from this layout."""
    result = {"Name": ""}
    notes = []

    for line in lines:
        text = " ".join(t for _, _, t in line).lower()
        if "dob" in text and "student id" in text:
            name_words = left_of_first_label(line, ["DOB", "Student ID", "Print Date"])
            name_text = " ".join(t for _, _, t in name_words).strip()
            if name_text:
                result["Name"] = name_text
            else:
                notes.append("Layout A: no name text found before the 'DOB:' label on its line")
            break
    else:
        notes.append("Layout A: 'DOB'/'Student ID' header line not found")

    return result, notes


# ===========================================================================
# LAYOUT B ("Etran Omed Only") -- Address, ID Number, SSN, Birth Date/Name
# ===========================================================================
def parse_layout_b(lines):
    result = {"Name": "", "Address": "", "ID Number": "", "SSN": "", "Birth Date": "", "Birth Name": ""}
    notes = []

    for line in lines:
        for out_key, variants in IDENTITY_LABELS_B.items():
            if result[out_key]:
                continue
            for variant in variants:
                lbl = " ".join(variant)
                spans = find_label_spans(line, [lbl])
                if spans:
                    vals, _ = label_values(line, [lbl])
                    result[out_key] = " ".join(t for _, _, t in vals.get(lbl, []))
                    break

    name_idx = None
    for idx, line in enumerate(lines):
        left_words = left_of_first_label(line, ALL_LABEL_STRINGS_B)
        left_text = " ".join(t for _, _, t in left_words).strip()
        if HONORIFIC_RE.match(left_text):
            name_idx = idx
            result["Name"] = HONORIFIC_RE.sub("", left_text).strip()
            break

    if name_idx is None:
        notes.append("Layout B: name line (starting with Mr./Ms./Mrs./Dr.) not found")
    else:
        addr_lines = []
        found_csz = False
        for idx in range(name_idx + 1, min(name_idx + 4, len(lines))):
            left_words = left_of_first_label(lines[idx], ALL_LABEL_STRINGS_B)
            text = " ".join(t for _, _, t in left_words).strip()
            if not text:
                break
            addr_lines.append(text)
            if CITY_STATE_ZIP_RE.match(text):
                found_csz = True
                break
        if not found_csz:
            notes.append("Layout B: address block found but no city/state/zip-shaped line within it -- check Address")
        result["Address"] = ", ".join(addr_lines)
        result["_address_lines"] = addr_lines

    for k in ["ID Number", "SSN", "Birth Date", "Birth Name"]:
        if not result[k]:
            notes.append(f"Layout B: '{k}' not found/blank")

    return result, notes


# ===========================================================================
# PER-PDF DRIVER -- one row per page; Layout A and Layout B are extracted
# independently into their own columns, with no cross-matching between them.
# ===========================================================================
def process_pdf(path: Path):
    doc = fitz.open(str(path))
    rows = []
    n_a = n_b = 0

    for i, page in enumerate(doc, start=1):
        text = page.get_text()
        if len(text.strip()) < MIN_TEXT_CHARS_PER_PAGE:
            continue
        words = page.get_text("words")
        lines = group_words_into_lines(words)
        kind = classify_page(lines)
        if kind not in ("A", "B"):
            continue

        row = {col: "" for col in OUTPUT_COLUMNS if col not in ("File Name", "Page Number", "Extraction Notes")}

        if kind == "A":
            n_a += 1
            rec, notes = parse_layout_a(lines)
            row["Name (Summary Page)"] = rec.get("Name", "")
        else:
            n_b += 1
            rec, notes = parse_layout_b(lines)
            row["Name (As Printed)"] = rec.get("Name", "")
            row["Address"] = rec.get("Address", "")
            row["Street"], row["Apt/Suite"], row["City"], row["State"], row["Zip"] = \
                split_address_parts(rec.get("_address_lines", []))
            row["First Name"], row["Middle Name"], row["Last Name"] = split_name_parts(rec.get("Name", ""))
            for k in ["ID Number", "SSN", "Birth Date", "Birth Name"]:
                row[k] = rec.get(k, "")

        row["File Name"] = path.name
        row["Page Number"] = i
        row["Extraction Notes"] = "; ".join(notes)
        rows.append(row)

    doc.close()
    return rows, n_a, n_b


# ===========================================================================
# DEBUG (safe, PII-free diagnostic dump)
# ===========================================================================
KNOWN_DEBUG_LABELS = [
    "DOB", "Student ID", "Print Date", "Transfer Credits", "Hours Attempted", "Hours Passed",
    "ID Number", "SSN", "Birth Date", "Birth Name", "Etran Omed Only",
]


def mask_shape(s):
    return re.sub(r"[A-Za-z]", "X", re.sub(r"\d", "#", s))


def mask_line_except_labels(line):
    tokens = [t for _, _, t in line]
    lowered = [t.strip(":").lower() for t in tokens]
    n = len(tokens)
    is_label = [False] * n
    for lbl in sorted(KNOWN_DEBUG_LABELS, key=len, reverse=True):
        ltoks = lbl.lower().split()
        tl = len(ltoks)
        i = 0
        while i + tl <= n:
            if not any(is_label[i:i + tl]) and lowered[i:i + tl] == ltoks:
                for k in range(i, i + tl):
                    is_label[k] = True
            i += 1
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
          f"({'OK' if len(text.strip()) >= MIN_TEXT_CHARS_PER_PAGE else 'BELOW MIN -- treated as image-only'})")

    words = page.get_text("words")
    lines = group_words_into_lines(words)
    kind = classify_page(lines)
    print(f"page classified as: {kind or 'UNRECOGNIZED (neither Layout A nor Layout B markers found)'}")
    if kind == "A":
        rec, notes = parse_layout_a(lines)
        print(f"  parsed Name: {'(found)' if rec['Name'] else '(blank)'}")
        if notes:
            print(f"  notes: {'; '.join(notes)}")
    elif kind == "B":
        rec, notes = parse_layout_b(lines)
        for k in ["Name", "Address", "ID Number", "SSN", "Birth Date", "Birth Name"]:
            print(f"  parsed {k}: {'(found)' if rec.get(k) else '(blank)'}")
        if notes:
            print(f"  notes: {'; '.join(notes)}")
    print(f"lines detected: {len(lines)}")
    for i, line in enumerate(lines):
        print(f"  [{i:>3}] {mask_line_except_labels(line)}")
    doc.close()


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
    print("Academic Transcript Identity Extractor")
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
    empty_files = []
    with tqdm(pdfs, desc="Extracting", unit="pdf", ncols=100) as pbar:
        for pdf_path in pbar:
            pbar.set_postfix_str(pdf_path.name)
            rows, n_a, n_b = process_pdf(pdf_path)
            if n_b == 0:
                empty_files.append(pdf_path.name)
            for r in rows:
                if r.get("Extraction Notes"):
                    flagged += 1
            all_rows.extend(rows)
            status_callback(f"Processed {pdf_path.name} ({len(all_rows)} row(s) so far)")

    if not all_rows:
        print("No identity (Layout B / 'Etran Omed Only') pages found in any file -- nothing to write.")
        status_callback("Done. No identity pages found -- nothing written.")
        return False

    df = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS)
    df.to_excel(output_path, index=False)

    print(f"\nDone. {len(all_rows)} row(s) from {len(pdfs)} file(s) -> {output_path}")
    if flagged:
        print(f"{flagged} row(s) have a non-empty 'Extraction Notes' -- spot-check those for missed/blank fields.")
    if empty_files:
        print(f"{len(empty_files)} file(s) had no identity page at all (nothing written for them): "
              + ", ".join(empty_files))
    status_callback(f"Done. {len(all_rows)} row(s) written to {output_path.name} ({flagged} flagged for review).")
    return True


# ===========================================================================
# TKINTER GUI
# ===========================================================================
class ExtractorGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Academic Transcript Identity Extractor")
        self.geometry("640x280")
        self.resizable(False, False)
        self._running = False
        self._build_widgets()

    def _build_widgets(self):
        pad = {"padx": 12, "pady": 8}

        title = ttk.Label(self, text="Academic Transcript Identity Extractor",
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

    app = ExtractorGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
