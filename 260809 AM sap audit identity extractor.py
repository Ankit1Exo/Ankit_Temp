"""
SAP Audit Identity Extractor -- Tkinter GUI
==========================================================
Extracts the student ID, SSN and Name from "Satisfactory Academic Progress
Audit Report" PDFs into an Excel workbook.

THE PAGE STRUCTURE THIS READS
    Each page repeats a four-line block of headings, and the student data
    follows the row of dashes:

        line 1    ID SSN     Name    Incl Incl  GPA   GPA  ...     <- captions
        line 2    Att                                              <- wrapped caption
        line 3    Course Name  Term/Dt  Grd Cum  Eval  Credits ... <- captions
        line 4    -----------------------------------------       <- separator
        ---->     1234567 123-45-6789 Mrs. Jane D. Smith  Academic Program: ...
                  #Excluded Remedial Credits    SAP Type: DHDHS  ...
                  DHSC-821  21FA2  A  Yes  No   3.00  12.00000  ...

    Everything after the dashes is data. A student's ID, SSN and Name are
    all on ONE line, in that order. The lines below it -- the wrapped tail
    of a long line, the "#Excluded Remedial Credits" row, and the course
    rows -- belong to that student but hold nothing this script wants.

HOW A LINE IS REBUILT
    Words are grouped into printed lines strictly by VERTICAL POSITION,
    using a tolerance derived from the page's own glyph height. PyMuPDF's
    built-in line breaking is deliberately not used: on a report this wide
    it emits text in the PDF's internal block order, which pulled words
    from the course headings and the "#Excluded Remedial Credits" row into
    the student line and produced names like "Credits Remedial".

HOW A LINE IS SPLIT
    ID    -- everything to the left of the SSN. Kept as printed.
    SSN   -- the first SSN-shaped value on the line: 123-45-6789, a masked
             XXX-XX-6789, or a running 123456789. Written out EXACTLY as
             printed -- never reformatted.
    Name  -- what follows the SSN, stopping at the "Academic Program"
             caption, at any other known caption, at the report's own
             vocabulary (Credits, Remedial, Excluded, ...), or at a token
             holding a digit -- whichever comes first.

             The heading line's "Incl" column marks where the Name column
             ends, but that edge is only WARNED about, never enforced:
             cutting names at it truncated "Samuel P. Okonkwo Jr." and, on
             pages whose heading geometry did not line up with the data,
             removed names entirely.

    The name then has any Mr./Mrs./Ms./Miss/Dr. prefix moved to its own
    column and is split three ways:

        with a middle initial   "Jane D. Smith"
            First Name = Jane      Middle = D.      Last Name = Smith
        without one             "Jane Smith"
            First Name = Jane      Middle = (blank) Last Name = Smith

    i.e. an initial, when present, is the split point: what precedes it is
    the first name, what follows it is the last name. With no initial the
    first space is the split point.

VALIDITY
    A line becomes a row only if it has an SSN AND a name that survives
    the boundary rules. Rows are then de-duplicated per page on the SSN.
    Anything a rule could not resolve is left blank with the reason in
    "Extraction Notes" -- treat any row with a note as needing a look.

GUI:
    - Source folder picker (scanned recursively for PDFs)
    - Destination folder picker (where the output XLSX goes)
    - Start Extraction button; progress prints to the console window

USAGE:
    python "260809 AM sap audit identity extractor.py"
    python "260809 AM sap audit identity extractor.py" --diagnose <pdf_or_folder>
    python "260809 AM sap audit identity extractor.py" --debug <pdf_or_folder> [page]

--diagnose reports, per page, whether the heading line and the dashes were
found, how many lines were rebuilt, and how many students came out.
--debug dumps every rebuilt line of one page. Both mask values to their
digit/letter shape (#/X) and leave only the report's own captions
readable, so output can be pasted into a ticket without exposing PII.

REQUIREMENTS:
    pip install pymupdf pandas openpyxl tqdm

SECURITY NOTE:
    These reports contain SSNs and student names, and the output workbook
    holds them in clear. Run only on an authorised workstation, save the
    XLSX only to the approved Global Insider folder (never a desktop or a
    local temp path), and delete the local copy once it has been loaded
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
OUTPUT_XLSX_NAME = "sap_audit_identity.xlsx"

# Fraction of the page's median glyph height used as the vertical
# tolerance when grouping words into a printed line. Comfortably absorbs
# the sub-point jitter inside one line while staying well under the gap
# to the next line.
LINE_TOLERANCE_RATIO = 0.4

# A word may start slightly left of its heading caption (numeric columns
# are often right-aligned), so column tests allow this much slack.
COLUMN_PAD = 4

# No \b in front: on a fused "1234567123-45-6789" the ID runs straight
# into the SSN and a word boundary would refuse to match mid-digits.
# X and * are allowed so a masked SSN is still recognised, and carried
# through exactly as printed.
SSN_DASHED_RE = re.compile(r"[0-9X*]{3}-[0-9X*]{2}-[0-9X*]{4}")

# The running form, only where whitespace or the word edge proves where it
# starts and ends -- otherwise a 9-digit run could be part of a longer
# number.
SSN_PLAIN_RE = re.compile(r"(?<!\d)(\d{9})(?!\d)")

# Requiring a period or a following capital matters: a bare
# "^(Mr|Mrs|Ms|Dr)" would also strike the first two letters off real
# surnames such as "Mroz" or "Drury".
HONORIFIC_RE = re.compile(r"^((?:Mr|Mrs|Ms|Miss|Dr|Prof)\.?)(?:\s+|(?=[A-Z]))")

INITIAL_RE = re.compile(r"^[A-Za-z]\.?$")

NAME_SUFFIXES = {"jr", "jr.", "sr", "sr.", "ii", "iii", "iv", "v"}

# A row of the course table: "DHSC-821", "DHSC 821".
COURSE_ROW_RE = re.compile(r"^[A-Za-z]{2,6}[- ]?\d{3}\b")

# Captions that can follow the name on the student line.
KNOWN_TRAILING_CAPTIONS = [
    ["academic", "program"],
    ["sap", "type"],
    ["excluded", "remedial"],
    ["program"],
    ["status"],
    ["degree"],
    ["major"],
]
KNOWN_CAPTION_LABELS = {" ".join(c) for c in KNOWN_TRAILING_CAPTIONS}
PREFERRED_BOUNDARY = "academic program"

# The report's own vocabulary, which never belongs to a student's name.
# Deliberately excludes words that are also real surnames (Page, Grade,
# Young, ...) so this can only ever cut report text, never a person.
REPORT_STOP_WORDS = {
    "credits", "credit", "remedial", "excluded", "gpa", "cmpl", "pgm", "incl", "eval",
    "grd", "cum", "att", "earn", "ssn", "dhdhs", "cred", "attempted", "sciences",
    "term/dt", "transfer", "satisfactory", "audit", "batch", "verified", "skipped",
    "section", "exists", "doctor", "health",
}

OUTPUT_COLUMNS = [
    "File Name", "Page Number",
    "ID", "SSN",
    "Prefix", "Full Name", "First Name", "Middle", "Last Name",
    "Name Boundary", "Extraction Notes",
]


# ===========================================================================
# 1. REBUILDING PRINTED LINES FROM WORD POSITIONS
# ===========================================================================
def page_lines(page):
    """Every printed line on the page, top to bottom, each as a list of
    (x0, x1, text) sorted left to right.

    Grouping is by the vertical centre of each word against a tolerance
    scaled to the page's own median glyph height, so it adapts to whatever
    point size the report was printed at. Words are compared to the
    ANCHOR of the line they might join -- the first word's centre -- not
    to the previous word, so a run of slightly drifting words cannot chain
    two printed lines together."""
    words = page.get_text("words")
    if not words:
        return []

    heights = sorted(w[3] - w[1] for w in words)
    median_height = heights[len(heights) // 2] or 1.0
    tolerance = median_height * LINE_TOLERANCE_RATIO

    items = sorted(((w[1] + w[3]) / 2.0, w[0], w[2], w[4]) for w in words)

    lines, current, anchor = [], [], None
    for centre, x0, x1, text in items:
        if anchor is None or centre - anchor <= tolerance:
            if anchor is None:
                anchor = centre
            current.append((x0, x1, text))
        else:
            lines.append(sorted(current, key=lambda t: t[0]))
            current, anchor = [(x0, x1, text)], centre
    if current:
        lines.append(sorted(current, key=lambda t: t[0]))
    return lines


def line_text(line):
    return " ".join(t for _, _, t in line)


# ===========================================================================
# 2. FINDING THE PAGE STRUCTURE
# ===========================================================================
def find_heading_columns(lines):
    """The x-positions of the captions on the heading line
    ("ID SSN Name Incl ...").

    Returns {"id", "ssn", "name", "name_end"} or None. "name_end" is the x
    of the caption AFTER Name -- normally "Incl" -- and is the right-hand
    boundary of the Name column, which is what keeps the Academic Program
    text out of the name."""
    for line in lines:
        tokens = [t.strip(":").lower() for _, _, t in line]
        for i in range(len(tokens) - 2):
            if tokens[i:i + 3] == ["id", "ssn", "name"]:
                after = line[i + 3][0] if i + 3 < len(line) else None
                return {"id": line[i][0], "ssn": line[i + 1][0],
                        "name": line[i + 2][0], "name_end": after}
    return None


def is_separator(line):
    """The row of dashes that closes the headings. Allows for the report
    wrapping it across more than one printed line."""
    text = line_text(line).replace(" ", "")
    return len(text) >= 8 and text.count("-") >= len(text) * 0.9


def data_region(lines):
    """The lines after the dashes. If a page has no dashes -- continuation
    pages sometimes don't -- every line is treated as data rather than
    dropping the page."""
    for index, line in enumerate(lines):
        if is_separator(line):
            return lines[index + 1:]
    return lines


def is_not_student_line(text):
    """Rows inside a student's block that are definitely not the student
    line: the excluded-credits row and the course table."""
    stripped = text.lstrip()
    return stripped.startswith("#") or bool(COURSE_ROW_RE.match(stripped))


# ===========================================================================
# 3. SPLITTING A STUDENT LINE
# ===========================================================================
def locate_ssn(line):
    """Find the SSN among a line's words.

    Returns (word_index, before_text, ssn_text, after_text) where `before`
    and `after` are any part of that same word either side of the SSN --
    the report can print the ID fused to the SSN. Returns None if the line
    holds no SSN."""
    for index, (_, _, text) in enumerate(line):
        match = SSN_DASHED_RE.search(text) or SSN_PLAIN_RE.search(text)
        if match:
            return index, text[:match.start()], match.group(0), text[match.end():]
    return None


def name_tokens_after(line, ssn_index, after):
    """The words to the right of the SSN, as (token, x) pairs. `after` is
    any tail of the SSN's own word, which shares that word's x."""
    tokens, xs = [], []
    if after:
        tokens.append(after)
        xs.append(line[ssn_index][0])
    for x0, _, text in line[ssn_index + 1:]:
        tokens.append(text)
        xs.append(x0)
    return tokens, xs


def overflows_name_column(xs, stop, columns):
    """True if the chosen name runs past the Name column's right edge.

    This is reported, never enforced. Cutting the name at the column edge
    was tried and was wrong twice: it truncated "Samuel P. Okonkwo Jr." to
    "... Okonkwo", and on pages whose heading geometry didn't line up with
    the data it removed the name outright. The caption and report-word
    rules do the real work; the column edge only earns a warning."""
    if not columns or not columns["name_end"] or stop == 0:
        return False
    return max(xs[:stop]) >= columns["name_end"] - COLUMN_PAD


def find_name_stop(tokens):
    """Where the name ends within `tokens`. Returns (index, reason). Every
    rule is evaluated and the EARLIEST stop wins, so a caption late on the
    line cannot pull the name past a digit that came first."""
    norm = [t.strip(":").lower() for t in tokens]
    best_index, best_reason = len(tokens), "end of line"

    def offer(index, reason):
        nonlocal best_index, best_reason
        if index < best_index:
            best_index, best_reason = index, reason

    for caption in KNOWN_TRAILING_CAPTIONS:
        label, width = " ".join(caption), len(caption)
        for i in range(len(norm) - width + 1):
            if norm[i:i + width] == caption:
                offer(i, label)
                break
        # Tight kerning sometimes fuses a two-word caption into one token.
        if width > 1:
            fused = "".join(caption)
            for i, token in enumerate(norm):
                if token == fused:
                    offer(i, label)
                    break

    for i, token in enumerate(tokens):
        if token.strip(":.,#").lower() in REPORT_STOP_WORDS:
            offer(i, f"report text ('{token}')")
            break

    for i, token in enumerate(tokens):
        if re.search(r"\d", token) or token.startswith("("):
            offer(i, "a token containing a digit or '('")
            break

    for i, token in enumerate(tokens):
        if token.endswith(":"):
            offer(i, f"an unrecognised caption ({token})")
            break

    return best_index, best_reason


def strip_prefix(name):
    """(prefix, name_without_prefix). Handles "Mrs. Jane" and the fused
    "Mrs.Jane" that tight kerning produces."""
    match = HONORIFIC_RE.match(name)
    if not match:
        return "", name.strip()
    return match.group(1), name[match.end():].strip()


def split_name(name, notes):
    """Split a prefix-free name into (first, middle, last)."""
    tokens = name.split()
    if not tokens:
        return "", "", ""

    if any("," in t for t in tokens):
        notes.append("name contains a comma -- it may be printed 'Last, First' rather than "
                     "'First M. Last'; split applied as if 'First M. Last'")

    if tokens[-1].strip(".,").lower() in NAME_SUFFIXES:
        notes.append(f"name ends in a suffix ('{tokens[-1]}') -- left attached to the last name")

    if len(tokens) == 1:
        notes.append("name is a single word -- put in First Name, Last Name left blank")
        return tokens[0], "", ""

    # An initial only counts as the split point if it is neither the first
    # token nor the last -- a trailing lone letter ("Mroz Dana K") has no
    # surname after it to split off.
    start = next((i for i in range(1, len(tokens) - 1) if INITIAL_RE.match(tokens[i])), None)

    if start is None:
        if INITIAL_RE.match(tokens[-1]):
            notes.append(f"name ends in a lone initial ('{tokens[-1]}') with nothing after it -- "
                         "treated as part of the last name, not as a middle initial")
        return tokens[0], "", " ".join(tokens[1:])

    end = start
    while end + 1 < len(tokens) - 1 and INITIAL_RE.match(tokens[end + 1]):
        end += 1

    return " ".join(tokens[:start]), " ".join(tokens[start:end + 1]), " ".join(tokens[end + 1:])


def parse_student_line(line, columns):
    """One student row from a rebuilt line, or None if the line is not a
    student line."""
    text = line_text(line)
    if is_not_student_line(text):
        return None

    located = locate_ssn(line)
    if located is None:
        return None
    ssn_index, before, ssn_text, after = located

    notes = []
    row = {col: "" for col in OUTPUT_COLUMNS}
    row["SSN"] = ssn_text

    # --- ID: everything left of the SSN, plus any digits fused in front
    #     of it within the same word.
    id_tokens = [t for _, _, t in line[:ssn_index]]
    if before:
        id_tokens.append(before)
    if not id_tokens:
        notes.append("no ID printed to the left of the SSN")
    else:
        row["ID"] = id_tokens[-1]
        if len(id_tokens) > 1:
            notes.append(f"{len(id_tokens)} tokens found left of the SSN "
                         f"('{' '.join(id_tokens)}') -- took the one nearest the SSN as the ID")

    # --- Name: everything right of the SSN, cut by the caption rules.
    #
    # The caption rules run on the UNCUT token list. Applying the column
    # edge first would hide "Academic Program" from them -- chopping
    # "Program:" off leaves a bare "Academic", which matches no caption
    # and lands in the name.
    tokens, xs = name_tokens_after(line, ssn_index, after)
    stop, reason = find_name_stop(tokens)

    prefix, full_name = strip_prefix(" ".join(tokens[:stop]))

    # Nothing survived, so what followed the SSN was report text, not a
    # person. Reject the row rather than emit a blank name beside a real
    # SSN.
    if not full_name:
        return None

    row["Name Boundary"] = reason
    row["Prefix"] = prefix
    row["Full Name"] = full_name
    if reason.startswith("report text"):
        notes.append(f"name was cut at {reason} -- the line mixed in text from another column; "
                     "verify the name against the PDF")
    elif reason.startswith("a token containing") or reason.startswith("an unrecognised caption"):
        notes.append(f"name ended at {reason} rather than the 'Academic Program' caption -- "
                     "verify the name is complete and has nothing extra")
    if overflows_name_column(xs, stop, columns):
        notes.append("name runs past the right edge of the Name column on the heading line -- "
                     "either it is simply a long name, or text from the next column was picked "
                     "up; verify against the PDF")

    row["First Name"], row["Middle"], row["Last Name"] = split_name(full_name, notes)
    row["Extraction Notes"] = "; ".join(notes)
    return row


# ===========================================================================
# 4. PER-PAGE AND PER-FILE DRIVERS
# ===========================================================================
def row_score(row):
    """Which reading of one SSN to trust. A numeric ID matters most: a
    mis-assembled row leaves a word where the ID should be."""
    boundary = row["Name Boundary"]
    rank = (3 if boundary == PREFERRED_BOUNDARY else
            2 if boundary in KNOWN_CAPTION_LABELS else
            1 if boundary == "end of line" else 0)
    note_count = len([n for n in row["Extraction Notes"].split("; ") if n])
    return (1 if row["ID"].isdigit() else 0, rank, -note_count, len(row["Full Name"]))


def dedupe_rows(rows):
    """One row per SSN per page. Keyed on the SSN and NOT the ID, because a
    mis-assembled row carries the wrong ID -- keying on it would let a good
    and a bad reading of one student both reach the workbook."""
    groups, order = {}, []
    for row in rows:
        key = (row["File Name"], row["Page Number"], row["SSN"])
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(row)

    merged = []
    for key in order:
        variants = groups[key]
        best = max(variants, key=row_score)
        if len({(v["ID"], v["Full Name"]) for v in variants}) > 1:
            best = dict(best)
            discarded = sorted(f"{v['ID'] or '(no ID)'} / {v['Full Name'] or '(no name)'}"
                               for v in variants
                               if (v["ID"], v["Full Name"]) != (best["ID"], best["Full Name"]))
            note = (f"this SSN was read more than one way; kept the best-scoring reading and "
                    f"discarded: {'; '.join(discarded)}")
            best["Extraction Notes"] = "; ".join(x for x in [best["Extraction Notes"], note] if x)
        merged.append(best)
    return merged


def read_page(page, file_name, page_num):
    """Returns (rows, heading_found, separator_found, line_count)."""
    lines = page_lines(page)
    columns = find_heading_columns(lines)
    separator_found = any(is_separator(ln) for ln in lines)

    rows = []
    for line in data_region(lines):
        try:
            row = parse_student_line(line, columns)
        except Exception as e:
            row = {col: "" for col in OUTPUT_COLUMNS}
            row["Extraction Notes"] = f"ERROR: {type(e).__name__}: {e}"
        if row is None:
            continue
        row["File Name"], row["Page Number"] = file_name, page_num
        rows.append(row)

    return dedupe_rows(rows), columns is not None, separator_found, len(lines)


def process_pdf(path: Path):
    doc = fitz.open(str(path))
    rows, image_only_pages, no_heading = [], [], 0

    for page_num, page in enumerate(doc, start=1):
        if len(page.get_text().strip()) < MIN_TEXT_CHARS_PER_PAGE:
            image_only_pages.append(page_num)
            continue
        page_rows, heading, _, _ = read_page(page, path.name, page_num)
        if not heading:
            no_heading += 1
        rows.extend(page_rows)

    doc.close()
    return rows, image_only_pages, no_heading


# ===========================================================================
# 5. MASKING (shared by --debug and --diagnose)
# ===========================================================================
KNOWN_DEBUG_LABELS = [
    "ID", "SSN", "Name", "Academic Program", "SAP Type", "Excluded Remedial Credits",
    "Course Name", "Term/Dt", "Grd", "Cum", "Eval", "Credits", "Grade Pts",
    "Report Options", "Batch ID", "Satisfactory Academic Progress Audit Report",
    "Detail of Results by Student by SAP Type", "Page", "Att", "Pgm", "Earn", "Cmpl", "GPA",
    "Incl", "Section skipped", "No Verified Grade Exists", "Doctor of Health Sciences",
]


def mask_shape(s):
    return re.sub(r"[A-Za-z]", "X", re.sub(r"\d", "#", s))


def mask_text_except_labels(text):
    tokens = text.split()
    lowered = [t.strip(":").lower() for t in tokens]
    is_label = [False] * len(tokens)
    for label in sorted(KNOWN_DEBUG_LABELS, key=len, reverse=True):
        parts = label.lower().split()
        width = len(parts)
        for i in range(len(tokens) - width + 1):
            if not any(is_label[i:i + width]) and lowered[i:i + width] == parts:
                for k in range(i, i + width):
                    is_label[k] = True
    return " ".join(t if is_label[k] else mask_shape(t) for k, t in enumerate(tokens))


# ===========================================================================
# 6. DEBUG / DIAGNOSE
# ===========================================================================
def debug_page(path: Path, page_num: int):
    doc = fitz.open(str(path))
    if page_num < 1 or page_num > len(doc):
        print(f"{path.name}: page {page_num} out of range (document has {len(doc)} page(s))")
        doc.close()
        return

    page = doc[page_num - 1]
    chars = len(page.get_text().strip())
    print(f"--- {path.name} page {page_num} ---")
    print(f"text layer: {chars} chars "
          f"({'OK' if chars >= MIN_TEXT_CHARS_PER_PAGE else 'BELOW MIN -- image-only page'})")

    lines = page_lines(page)
    columns = find_heading_columns(lines)
    separator_at = next((i for i, ln in enumerate(lines) if is_separator(ln)), None)
    print(f"heading line 'ID SSN Name': {'found' if columns else 'NOT FOUND'}"
          + (f" (Name column ends at x={columns['name_end']})"
             if columns and columns["name_end"] else ""))
    print(f"dashes separator: {'line ' + str(separator_at) if separator_at is not None else 'NOT FOUND'}")

    rows, _, _, _ = read_page(page, path.name, page_num)
    print(f"students found: {len(rows)}")
    for row in rows:
        print(f"  ID {'(found)' if row['ID'] else '(blank)'} | "
              f"SSN {'(found)' if row['SSN'] else '(blank)'} | "
              f"name {'(found)' if row['Full Name'] else '(blank)'} | "
              f"boundary: {row['Name Boundary']}")
        if row["Extraction Notes"]:
            print(f"    notes: {row['Extraction Notes']}")

    print(f"lines rebuilt: {len(lines)}")
    for i, line in enumerate(lines):
        marker = "  <-- separator" if is_separator(line) else ""
        print(f"  [{i:>3}] {mask_text_except_labels(line_text(line))}{marker}")
    doc.close()


def diagnose(target: Path):
    """Per-page account of what was found. No names, IDs or SSNs are
    printed -- only counts, captions and masked shapes -- so the output is
    safe to paste into a ticket."""
    pdfs = [target] if target.is_file() else sorted(target.rglob("*.pdf"))
    if not pdfs:
        print(f"No PDF files found at: {target}")
        return

    print(f"Diagnosing {len(pdfs)} PDF file(s). No PII is printed below.\n")
    for path in pdfs:
        doc = fitz.open(str(path))
        print(f"=== {path.name} ({len(doc)} page(s)) ===")
        total = 0

        for page_num, page in enumerate(doc, start=1):
            chars = len(page.get_text().strip())
            if chars < MIN_TEXT_CHARS_PER_PAGE:
                print(f"  page {page_num}: SKIPPED -- only {chars} chars of text. This page is a "
                      f"scanned image with no text layer and must be OCR'd first.")
                continue

            rows, heading, separator, line_count = read_page(page, path.name, page_num)
            total += len(rows)
            print(f"  page {page_num}: {len(rows)} student(s)  "
                  f"[{chars} chars, {line_count} lines rebuilt, "
                  f"heading {'YES' if heading else 'no'}, dashes {'YES' if separator else 'no'}]")

            if not rows:
                lines = data_region(page_lines(page))
                near = [ln for ln in lines
                        if locate_ssn(ln) or HONORIFIC_RE.match(line_text(ln))]
                if not near:
                    print("      no line after the dashes holds an SSN-shaped value or a "
                          "Mr./Mrs./Ms./Dr. prefix -- there is nothing here to extract")
                for line in near[:5]:
                    print(f"      near-miss: {mask_text_except_labels(line_text(line))[:150]}")

        print(f"  TOTAL: {total} student(s) in this file\n")
        doc.close()


# ===========================================================================
# 7. EXTRACTION RUNNER
# ===========================================================================
def run_extraction(source_folder, dest_folder, status_callback):
    src, dst = Path(source_folder), Path(dest_folder)
    if not src.is_dir():
        status_callback("ERROR: Source folder invalid.")
        return False
    dst.mkdir(parents=True, exist_ok=True)
    output_path = dst / OUTPUT_XLSX_NAME

    print("=" * 70)
    print("SAP Audit Identity Extractor")
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

    all_rows, no_students, image_only, headingless = [], [], [], 0
    with tqdm(pdfs, desc="Extracting", unit="pdf", ncols=100) as pbar:
        for pdf_path in pbar:
            pbar.set_postfix_str(pdf_path.name)
            rows, image_pages, no_heading = process_pdf(pdf_path)
            if not rows:
                no_students.append(pdf_path.name)
            if image_pages:
                image_only.append(f"{pdf_path.name} (page(s) {', '.join(map(str, image_pages))})")
            headingless += no_heading
            all_rows.extend(rows)
            status_callback(f"Processed {pdf_path.name} ({len(all_rows)} student(s) so far)")

    if not all_rows:
        print("No student lines found in any file -- nothing to write.")
        status_callback("Done. No student lines found -- nothing written.")
        return False

    frame = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS)
    # ID and SSN go out as text so Excel cannot strip a leading zero or
    # turn a running 9-digit SSN into scientific notation.
    for col in ["ID", "SSN"]:
        frame[col] = frame[col].astype(str)
    frame.to_excel(output_path, index=False)

    flagged = sum(1 for r in all_rows if r["Extraction Notes"])
    files_with_rows = len({r["File Name"] for r in all_rows})
    print(f"\nDone. {len(all_rows)} student row(s) from {files_with_rows} of {len(pdfs)} "
          f"file(s) -> {output_path}")

    boundaries = {}
    for r in all_rows:
        boundaries[r["Name Boundary"]] = boundaries.get(r["Name Boundary"], 0) + 1
    print("Where the name ended:")
    for reason, count in sorted(boundaries.items(), key=lambda kv: -kv[1]):
        marker = "" if reason == PREFERRED_BOUNDARY else "   <-- not the 'Academic Program' caption"
        print(f"  {count:>5}  {reason}{marker}")

    if flagged:
        print(f"{flagged} row(s) have a non-empty 'Extraction Notes' -- spot-check those.")
    if headingless:
        print(f"{headingless} page(s) had no 'ID SSN Name' heading line, so the Name column's "
              f"right edge was unknown and caption rules were used instead.")
    if image_only:
        print(f"\nSCANNED PAGES WITH NO TEXT LAYER (nothing can be read off these until they are "
              f"OCR'd):\n  {'; '.join(image_only)}")
    if no_students:
        print(f"\n{len(no_students)} file(s) produced no rows: {', '.join(no_students)}")
        print("   -> run with --diagnose <folder> to see why, per page. The output is PII-free.")

    status_callback(f"Done. {len(all_rows)} student(s) written to {output_path.name} "
                    f"({flagged} flagged for review).")
    return True


# ===========================================================================
# 8. TKINTER GUI
# ===========================================================================
class ExtractorGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SAP Audit Identity Extractor")
        self.geometry("640x300")
        self.resizable(False, False)
        self._running = False
        self._build_widgets()

    def _build_widgets(self):
        pad = {"padx": 12, "pady": 8}

        ttk.Label(self, text="SAP Audit Identity Extractor",
                  font=("Segoe UI", 14, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(self, text="Extracts ID, SSN and Name. Progress is printed to the console window.",
                  foreground="#555").grid(row=1, column=0, columnspan=3, sticky="w", padx=12)
        ttk.Label(self, text="Output contains SSNs -- save to the approved Global Insider folder only.",
                  foreground="#a33").grid(row=2, column=0, columnspan=3, sticky="w", padx=12)

        ttk.Label(self, text="Source folder (PDFs):").grid(row=3, column=0, sticky="e", **pad)
        self.src_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.src_var, width=50).grid(row=3, column=1, sticky="we", **pad)
        ttk.Button(self, text="Browse...", command=self._pick_source).grid(row=3, column=2, **pad)

        ttk.Label(self, text="Destination folder:").grid(row=4, column=0, sticky="e", **pad)
        self.dst_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.dst_var, width=50).grid(row=4, column=1, sticky="we", **pad)
        ttk.Button(self, text="Browse...", command=self._pick_destination).grid(row=4, column=2, **pad)

        self.start_btn = ttk.Button(self, text="Start Extraction", command=self._start_clicked)
        self.start_btn.grid(row=5, column=0, columnspan=3, pady=12)

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w").grid(
            row=6, column=0, columnspan=3, sticky="we", padx=12, pady=(0, 12))

        self.columnconfigure(1, weight=1)

    def _pick_source(self):
        folder = filedialog.askdirectory(title="Select folder containing SAP audit report PDFs",
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
# 9. ENTRY POINT
# ===========================================================================
def main():
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "--diagnose":
        if len(sys.argv) < 3:
            print("Usage: --diagnose <pdf_file_or_folder>")
            return
        diagnose(Path(sys.argv[2]))
        return

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
