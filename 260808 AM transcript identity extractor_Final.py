"""
Academic Transcript Identity Extractor -- Tkinter GUI
========================================================
Extracts identity fields from the "Etran Omed Only" transcript layout ONLY.

That layout prints, in a left column:

    Mr. Coby L. Bullard          <- name (honorific optional)
    1221 Winding Arbor Trail     <- street
    Apt 12-202                   <- optional Apt/Suite line
    Raleigh, NC  27606           <- city, state, ZIP

with the captioned identity fields printed in a right-hand column at the
same page heights:

    ID Number:  ...
    SSN:        ...
    Birth Date: ...
    Birth Name: ...

The older "summary" layout (dotted-rule course listing with a
"DOB:/Student ID:/Print Date:" header) is NO LONGER parsed -- only the
"Etran Omed Only" pages produce rows.

The course table itself is NOT extracted (identity fields only).

OUTPUT COLUMNS
    File Name, Page Number,
    Full Name, FN, MN, LN,
    Address, Street, Apt/Suite, City, State, ZIP,
    ID Number, SSN, Birth Date, Birth Name,
    Extraction Notes

NAME HANDLING
    - Honorific (Mr./Mrs./Ms./Miss/Dr./Prof./Rev.) is stripped from every
      name column, including "Full Name". Note the alternation is ordered
      longest-first and requires whitespace after the prefix, so "Mrs."
      no longer leaves a stray "s." at the front of the name.
    - A trailing generational suffix (Jr., Sr., II, III, IV) stays with
      the Last Name, e.g. LN = "Bullard Jr.".
    - "Lastname, Firstname Middle" is reordered to "Firstname Middle
      Lastname" when a comma is present.

ADDRESS HANDLING
    - The name row is anchored on the "ID Number:" caption (falling back
      to the honorific, then to the row above "SSN:"), then the left
      column is walked downward, SKIPPING blank and label-only rows,
      until the city/state/ZIP row is reached. This is what fixes files
      where the address came out blank: the old logic stopped at the
      first blank row and only looked three rows ahead.
    - 2-line (no Apt/Suite) and 3-line addresses are both supported, and
      an Apt/Suite printed at the end of the street line
      ("123 Main St Apt 4") is split out too.

GUI:
    - Source folder picker (scanned recursively for PDFs)
    - Destination folder picker (where the output XLSX goes)
    - Start Extraction button; progress prints to the console window

USAGE:
    python "260808 AM transcript identity extractor_Final.py"
    python "260808 AM transcript identity extractor_Final.py" --debug <pdf_or_folder> [page_number]

REQUIREMENTS:
    pip install pymupdf pandas openpyxl tqdm

SECURITY NOTE:
    These transcripts contain SSNs, birth dates and home addresses. Run
    only on an authorised workstation, save the output only to the
    approved Global Insider folder (never the desktop), and delete the
    local copy once it has been loaded into the authorised system of
    record. The output workbook is NOT encrypted.
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

# Longest-first alternation + required trailing whitespace. With the old
# "^(Mr|Mrs|Ms|...)" ordering, "Mrs. Ann" matched the "Mr" branch and left
# "s. Ann" behind.
HONORIFIC_RE = re.compile(r"^(?:Mrs|Miss|Ms|Mr|Dr|Prof|Rev)\.?\s+", re.IGNORECASE)

# Generational suffixes only -- deliberately excludes a bare "V", which is
# far more often a middle initial than a fifth-generation suffix.
SUFFIX_TOKENS = {"jr", "jr.", "sr", "sr.", "ii", "iii", "iv", "2nd", "3rd", "4th"}

CITY_STATE_ZIP_RE = re.compile(
    r"^(?P<city>.+?)\s*,?\s*(?P<state>[A-Za-z]{2})\.?\s+(?P<zip>\d{5}(?:-\d{4})?)\s*$"
)

# An Apt/Suite printed on its own line.
APT_LINE_RE = re.compile(
    r"^(?:apt|apartment|ste|suite|unit|bldg|building|fl|floor|rm|room|lot|trlr|trailer|"
    r"box|p\.?\s?o\.?\s?box|dept|department|#)\b\.?",
    re.IGNORECASE,
)

# An Apt/Suite tacked onto the end of the street line ("123 Main St Apt 4").
APT_INLINE_RE = re.compile(
    r"\s+(?P<apt>(?:apt|apartment|ste|suite|unit|bldg|building|fl|floor|rm|room|lot|trlr)\.?\s*"
    r"[#]?\s*[\w\-/]+|#\s*[\w\-/]+)\s*$",
    re.IGNORECASE,
)

# Page furniture / course-table headers that must never be read as address.
JUNK_LINE_RE = re.compile(
    r"^(?:page\s+\d+\s+of\s+\d+\b|course\b|title\b|hrs\b|grd\b|etran\b|-{3,}|_{3,}|={3,})",
    re.IGNORECASE,
)

SSN_TOKEN_FULL_RE = re.compile(r"^\d{3}-\d{2}-\d{4}$|^\d{9}$")

IDENTITY_LABELS = {
    "ID Number": "ID Number",
    "SSN": "SSN",
    "Birth Date": "Birth Date",
    "Birth Name": "Birth Name",
}
ALL_LABEL_STRINGS = list(IDENTITY_LABELS.values())

# "Birth Name" is routinely blank on this form, so a blank value there is
# not flagged; a MISSING caption still is.
FLAG_IF_BLANK = ["ID Number", "SSN", "Birth Date"]

MAX_ADDRESS_SCAN_ROWS = 12      # how far below the name row to look
MAX_ADDRESS_LINES = 4           # street + optional 2nd street/apt + city line
MAX_CONSECUTIVE_BLANKS = 3

OUTPUT_COLUMNS = [
    "File Name", "Page Number",
    "Full Name", "FN", "MN", "LN",
    "Address", "Street", "Apt/Suite", "City", "State", "ZIP",
    "ID Number", "SSN", "Birth Date", "Birth Name",
    "Extraction Notes",
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


def line_text(line):
    return " ".join(t for _, _, t in line).strip()


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


def left_of_first_label(line, label_strings=ALL_LABEL_STRINGS):
    """Everything on the row BEFORE the first right-hand-column caption --
    i.e. the left (name/address) column of a row that also carries an
    ID Number / SSN / Birth Date / Birth Name field at the same height."""
    spans = find_label_spans(line, label_strings)
    if not spans:
        return line
    first_start = min(s for _, s, _ in spans)
    return line[:first_start]


def left_text(line):
    return line_text(left_of_first_label(line))


def normalize_ssn_digits(s):
    return re.sub(r"\D", "", s or "")


# ===========================================================================
# NAME PARSING
# ===========================================================================
def strip_honorific(name_text):
    """Removes a leading Mr./Mrs./Ms./Miss/Dr./Prof./Rev., repeatedly, in
    case two got printed."""
    out = name_text.strip()
    while True:
        stripped = HONORIFIC_RE.sub("", out).strip()
        if stripped == out:
            return out
        out = stripped


def reformat_last_first_name(name_text):
    """"Lastname, Firstname Middle" -> "Firstname Middle Lastname". Only
    reorders when a comma is actually present, so it's a no-op on names
    that already print in the other order."""
    if "," not in name_text:
        return name_text
    last, _, rest = name_text.partition(",")
    return f"{rest.strip()} {last.strip()}".strip()


def clean_name_text(raw):
    """Raw left-column name cell -> canonical "Firstname Middle Lastname"
    with no honorific and no trailing punctuation."""
    txt = re.sub(r"\s+", " ", (raw or "").strip())
    # Separators only -- NOT ".", which is legitimate in "L." and "Jr.".
    txt = txt.strip(" ,;:")
    txt = strip_honorific(txt)
    txt = reformat_last_first_name(txt)
    return strip_honorific(txt).strip()  # honorific can precede the comma form


def split_name(full_name):
    """-> (FN, MN, LN). A trailing generational suffix stays with LN."""
    tokens = [t for t in full_name.split() if t]
    if not tokens:
        return "", "", ""

    suffix = ""
    if len(tokens) >= 3 and tokens[-1].lower().strip(",") in SUFFIX_TOKENS:
        suffix = tokens.pop().strip(",")

    if len(tokens) == 1:
        fn, mn, ln = tokens[0], "", ""
    elif len(tokens) == 2:
        fn, mn, ln = tokens[0], "", tokens[1]
    else:
        fn, mn, ln = tokens[0], " ".join(tokens[1:-1]), tokens[-1]

    if suffix:
        ln = f"{ln} {suffix}".strip()
    return fn, mn, ln


def looks_like_name(text):
    """Guards the name anchor against picking up a date stamp, a page
    header or an address row."""
    if not text:
        return False
    t = text.strip()
    if JUNK_LINE_RE.match(t):
        return False
    if re.match(r"^\d", t):                 # "1221 Winding Arbor Trail", "07/15/22"
        return False
    if CITY_STATE_ZIP_RE.match(t):
        return False
    return bool(re.search(r"[A-Za-z]{2}", t))


# ===========================================================================
# ADDRESS PARSING
# ===========================================================================
def split_address(addr_lines):
    """Address rows (top to bottom) -> Street / Apt-Suite / City / State /
    ZIP, plus a note if the city/state/ZIP row is missing."""
    parts = {"Street": "", "Apt/Suite": "", "City": "", "State": "", "ZIP": ""}
    notes = []
    lines = [re.sub(r"\s+", " ", l.strip()) for l in addr_lines if l and l.strip()]
    if not lines:
        return parts, ["Address: no address rows captured"]

    csz_idx = None
    for i in range(len(lines) - 1, -1, -1):
        if CITY_STATE_ZIP_RE.match(lines[i]):
            csz_idx = i
            break

    if csz_idx is None:
        notes.append("Address: no city/state/ZIP row found -- City/State/ZIP left blank, "
                     "check the full Address value")
        street_lines = lines
    else:
        m = CITY_STATE_ZIP_RE.match(lines[csz_idx])
        parts["City"] = m.group("city").strip(" ,")
        parts["State"] = m.group("state").upper()
        parts["ZIP"] = m.group("zip")
        street_lines = lines[:csz_idx]

    if not street_lines:
        notes.append("Address: city/state/ZIP found but no street row above it")
        return parts, notes

    apt_lines = [l for l in street_lines if APT_LINE_RE.match(l)]
    street_only = [l for l in street_lines if not APT_LINE_RE.match(l)]

    # A PO Box on its own is the street, not an Apt/Suite.
    if not street_only and apt_lines:
        street_only, apt_lines = apt_lines[:1], apt_lines[1:]

    street = " ".join(street_only).strip()
    apt = " ".join(apt_lines).strip()

    if not apt:
        m = APT_INLINE_RE.search(street)
        if m:
            apt = m.group("apt").strip()
            street = street[:m.start()].strip()

    parts["Street"] = street
    parts["Apt/Suite"] = apt
    return parts, notes


def collect_address_lines(lines, name_idx):
    """Walks the left column downward from the name row, SKIPPING blank and
    label-only rows, until the city/state/ZIP row (or the scan limits) are
    reached. The old version stopped at the first blank row within three
    rows of the name -- which is why some files came out with a blank
    Address."""
    addr, blanks = [], 0
    end = min(name_idx + 1 + MAX_ADDRESS_SCAN_ROWS, len(lines))
    for idx in range(name_idx + 1, end):
        text = left_text(lines[idx])
        if not text:
            blanks += 1
            if blanks >= MAX_CONSECUTIVE_BLANKS and addr:
                break
            continue
        if JUNK_LINE_RE.match(text):
            break
        blanks = 0
        addr.append(text)
        if CITY_STATE_ZIP_RE.match(text):
            break
        if len(addr) >= MAX_ADDRESS_LINES:
            break
    return addr


# ===========================================================================
# PAGE CLASSIFICATION -- "Etran Omed Only" layout only
# ===========================================================================
def classify_page(lines):
    text_lower = " ".join(line_text(l) for l in lines).lower()
    if "etran omed only" in text_lower:
        return "B"
    hits = sum(1 for lbl in ("id number", "ssn", "birth date", "birth name")
               if lbl in text_lower)
    return "B" if hits >= 2 else None


# ===========================================================================
# PAGE PARSING
# ===========================================================================
def find_name_row(lines):
    """Anchors the name row, most reliable signal first:
       1. the row carrying the "ID Number:" caption (the name prints to its
          left on this layout),
       2. any row starting with an honorific,
       3. the row directly above the "SSN:" caption.
    Returns (index, raw_left_text) or (None, "")."""
    id_idx = ssn_idx = None
    for idx, line in enumerate(lines):
        if id_idx is None and find_label_spans(line, ["ID Number"]):
            id_idx = idx
        if ssn_idx is None and find_label_spans(line, ["SSN"]):
            ssn_idx = idx

    if id_idx is not None:
        cand = left_text(lines[id_idx])
        if looks_like_name(cand):
            return id_idx, cand

    for idx, line in enumerate(lines):
        cand = left_text(line)
        if HONORIFIC_RE.match(cand) and looks_like_name(cand):
            return idx, cand

    if ssn_idx is not None and ssn_idx > 0:
        cand = left_text(lines[ssn_idx - 1])
        if looks_like_name(cand):
            return ssn_idx - 1, cand

    return None, ""


def parse_page(lines):
    result = {c: "" for c in OUTPUT_COLUMNS}
    notes = []

    # --- captioned right-hand-column fields -------------------------------
    found_labels = set()
    for line in lines:
        for out_key, lbl in IDENTITY_LABELS.items():
            if out_key in found_labels:
                continue
            if find_label_spans(line, [lbl]):
                found_labels.add(out_key)
                vals, _ = label_values(line, [lbl])
                result[out_key] = " ".join(t for _, _, t in vals.get(lbl, [])).strip()

    for k in IDENTITY_LABELS:
        if k not in found_labels:
            notes.append(f"'{k}' caption not found on this page")
        elif not result[k] and k in FLAG_IF_BLANK:
            notes.append(f"'{k}' caption found but value blank")

    # --- name --------------------------------------------------------------
    name_idx, raw_name = find_name_row(lines)
    if name_idx is None:
        notes.append("Name row not found (no 'ID Number:' caption, no honorific row) "
                     "-- Name and Address left blank")
        return result, notes

    full_name = clean_name_text(raw_name)
    if not full_name:
        notes.append("Name row located but the name cell was empty")
    result["Full Name"] = full_name
    result["FN"], result["MN"], result["LN"] = split_name(full_name)
    if full_name and not result["LN"]:
        notes.append("Name has only one word -- placed in FN, LN left blank")

    # --- address -----------------------------------------------------------
    addr_lines = collect_address_lines(lines, name_idx)
    if not addr_lines:
        notes.append("No address rows found below the name row -- Address blank")
    else:
        result["Address"] = ", ".join(addr_lines)
        parts, addr_notes = split_address(addr_lines)
        result.update(parts)
        notes.extend(addr_notes)

    return result, notes


# ===========================================================================
# PER-PDF DRIVER
# ===========================================================================
def process_pdf(path: Path):
    doc = fitz.open(str(path))
    rows = []
    image_only_pages = 0

    for i, page in enumerate(doc, start=1):
        text = page.get_text()
        if len(text.strip()) < MIN_TEXT_CHARS_PER_PAGE:
            image_only_pages += 1
            continue
        lines = group_words_into_lines(page.get_text("words"))
        if classify_page(lines) != "B":
            continue
        row, notes = parse_page(lines)
        row["File Name"] = path.name
        row["Page Number"] = i
        row["Extraction Notes"] = "; ".join(notes)
        rows.append(row)
    doc.close()

    # A multi-page "Etran Omed Only" printout repeats the same identity
    # block on continuation pages; keep the first page per person so one
    # student doesn't come out as two address rows.
    deduped, seen = [], set()
    for r in rows:
        key = (normalize_ssn_digits(r["SSN"]), r["ID Number"].strip(),
               r["Full Name"].strip().lower())
        if key != ("", "", "") and key in seen:
            continue
        seen.add(key)
        deduped.append(r)

    return deduped, len(rows) - len(deduped), image_only_pages


# ===========================================================================
# DEBUG (safe, PII-free diagnostic dump)
# ===========================================================================
KNOWN_DEBUG_LABELS = [
    "ID Number", "SSN", "Birth Date", "Birth Name", "Etran Omed Only",
    "Course", "Title", "Page",
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

    lines = group_words_into_lines(page.get_text("words"))
    kind = classify_page(lines)
    print(f"page classified as: {'Etran Omed Only (parsed)' if kind == 'B' else 'UNRECOGNISED (skipped)'}")
    if kind == "B":
        name_idx, _ = find_name_row(lines)
        print(f"  name row index: {name_idx if name_idx is not None else '(not found)'}")
        row, notes = parse_page(lines)
        for k in ["Full Name", "FN", "MN", "LN", "Address", "Street", "Apt/Suite",
                  "City", "State", "ZIP", "ID Number", "SSN", "Birth Date", "Birth Name"]:
            print(f"  parsed {k}: {'(found)' if row.get(k) else '(blank)'}")
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
    print("Academic Transcript Identity Extractor  (Etran Omed Only layout)")
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
    image_only_total = 0
    dupes_total = 0
    with tqdm(pdfs, desc="Extracting", unit="pdf", ncols=100) as pbar:
        for pdf_path in pbar:
            pbar.set_postfix_str(pdf_path.name)
            rows, dupes, image_only = process_pdf(pdf_path)
            dupes_total += dupes
            image_only_total += image_only
            if not rows:
                empty_files.append(pdf_path.name)
            for r in rows:
                if r.get("Extraction Notes"):
                    flagged += 1
            all_rows.extend(rows)
            status_callback(f"Processed {pdf_path.name} ({len(all_rows)} row(s) so far)")

    if not all_rows:
        print("No 'Etran Omed Only' pages found in any file -- nothing to write.")
        status_callback("Done. No identity pages found -- nothing written.")
        return False

    df = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS)
    df.to_excel(output_path, index=False)

    print(f"\nDone. {len(all_rows)} row(s) from {len(pdfs)} file(s) -> {output_path}")
    if flagged:
        print(f"{flagged} row(s) have a non-empty 'Extraction Notes' -- spot-check those "
              f"for missed/blank fields (especially Address).")
    if dupes_total:
        print(f"{dupes_total} duplicate identity page(s) skipped (same SSN/ID/name repeated "
              f"on a continuation page).")
    if image_only_total:
        print(f"{image_only_total} page(s) had no text layer (scanned image) and were skipped "
              f"-- those need OCR before they can be read.")
    if empty_files:
        print(f"{len(empty_files)} file(s) had no identity page at all (nothing written for them): "
              + ", ".join(empty_files))
    print("Reminder: save this workbook to the approved Global Insider folder, not the desktop, "
          "and delete local copies once loaded.")
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
