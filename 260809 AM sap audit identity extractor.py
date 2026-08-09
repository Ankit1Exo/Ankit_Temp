"""
SAP Audit Identity Extractor -- Tkinter GUI
==========================================================
Pulls the student identity fields out of "Satisfactory Academic Progress
Audit Report" PDFs. Only the student header line is read -- the course
table below it (DHSC-821, term, grade, credits, ...) is ignored.

The report prints one line per student, with the ID, the SSN and the name
running in sequence on that single line, under these captions:

        ID SSN          Name      Incl Incl  GPA   GPA ...
    1234567 123-45-6789 Mrs. Jane D. Smith  Academic Program: DH.DHSC (2021, ...)

From that line the script takes:

  ID    -- everything printed to the left of the SSN.
  SSN   -- kept EXACTLY as printed, dashed (123-45-6789) or running
           (123456789). It is never reformatted, so the output matches
           the source document character for character.
  Name  -- everything between the SSN and the end of the name (see
           below), with any Mr./Mrs./Ms./Miss/Dr. prefix removed into its
           own "Prefix" column, then split into three columns:

             with a middle initial   "Jane D. Smith"
                 First Name = Jane      Middle = D.     Last Name = Smith
             without one             "Jane Smith"
                 First Name = Jane      Middle = (blank) Last Name = Smith

           i.e. when an initial is present it is the split point -- what
           precedes it is the first name, what follows it is the last
           name. With no initial the first space is the split point.

WHERE THE NAME ENDS
    "Academic Program:" is the preferred boundary, but it is not
    guaranteed to be on every line, so the script stops the name at
    whichever of these comes first and records which one it used in the
    "Name Boundary" column:
        1. the "Academic Program" caption          (preferred)
        2. another known caption (SAP Type, ...)
        3. a token containing a digit or "("       (flagged)
        4. any other "Something:" caption          (flagged)
        5. the end of the line
    Rules 3 and 4 mean the layout differed from the sample, so those rows
    get an Extraction Note -- check them by hand.

A field that can't be located is left blank and the reason goes in
"Extraction Notes" rather than being guessed at. Treat any row with a
note as needing a manual look.

GUI:
    - Source folder picker (scanned recursively for PDFs)
    - Destination folder picker (where the output XLSX goes)
    - Start Extraction button; progress prints to the console window

USAGE:
    python "260809 AM sap audit identity extractor.py"
    python "260809 AM sap audit identity extractor.py" --diagnose <pdf_or_folder>
    python "260809 AM sap audit identity extractor.py" --debug <pdf_or_folder> [page_number]

--diagnose answers "why did this file produce no rows?": for every page it
reports whether the page had a text layer, whether it was recognised as a
SAP page, how many student lines were found, and -- for pages that found
none -- the near-miss lines. --debug dumps every line of one page.

Both print values masked to their digit/letter shape (#/X) with only the
known report captions left readable, so a layout problem can be diagnosed
and pasted into a ticket without exposing PII.

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

# No \b in front: on a fused "1234567123-45-6789" the ID runs straight
# into the SSN, and a word boundary would refuse to match mid-digits.
# X and * are allowed so a partially masked SSN ("XXX-XX-6789") is still
# recognised and carried through as printed.
SSN_DASHED_RE = re.compile(r"[0-9X*]{3}-[0-9X*]{2}-[0-9X*]{4}")

# The undashed form, only where whitespace (or the line edge) proves where
# it starts and ends -- otherwise a 9-digit run could be the tail of a
# longer number.
SSN_PLAIN_RE = re.compile(r"(?<!\S)(\d{9})(?!\S)")

# Requiring a period OR trailing whitespace matters: a bare
# "^(Mr|Mrs|Ms|Dr)" would also strike the first two letters off real
# surnames such as "Mroz" or "Drury".
HONORIFIC_RE = re.compile(r"^((?:Mr|Mrs|Ms|Miss|Dr|Prof)\.?)(?:\s+|(?=[A-Z]))")

# A middle initial: one letter, period optional.
INITIAL_RE = re.compile(r"^[A-Za-z]\.?$")

NAME_SUFFIXES = {"jr", "jr.", "sr", "sr.", "ii", "iii", "iv", "v"}

# Captions that can follow the name on the student line. "Academic
# Program" is listed first only for readability -- the earliest match on
# the line wins regardless of order here.
KNOWN_TRAILING_CAPTIONS = [
    ["academic", "program"],
    ["sap", "type"],
    ["excluded", "remedial"],
    ["program"],
    ["status"],
    ["degree"],
    ["major"],
]

PREFERRED_BOUNDARY = "academic program"

# Any one of these on a page is enough to treat it as a SAP audit page.
SAP_PAGE_SIGNATURES = [
    "satisfactory academic progress",
    "detail of results by student",
    "excluded remedial credits",
    "academic program",
    "sap type",
    "no verified grade exists",
]

OUTPUT_COLUMNS = [
    "File Name", "Page Number",
    "ID", "SSN",
    "Prefix", "Full Name", "First Name", "Middle", "Last Name",
    "Name Boundary", "Extraction Notes",
]


# ===========================================================================
# LINE / LABEL HELPERS
# ===========================================================================
def group_words_into_lines(words, overlap_ratio=0.5):
    """words: PyMuPDF page.get_text('words') output. Returns a list of
    lines, each a list of (x0, x1, text) tuples sorted left to right.

    Words are grouped by how much their vertical extents OVERLAP, not by
    rounding y into fixed buckets. Bucketing looks simpler but silently
    tears a line in half whenever its words straddle a bucket edge: real
    text layers carry sub-point jitter, so y=100.4 and y=101.6 -- the same
    printed line -- round to 99 and 102 and become two lines. On this
    report that put the ID and SSN in one fragment and the name in
    another, and the student was dropped entirely."""
    items = sorted(((w[1], w[3], w[0], w[2], w[4]) for w in words), key=lambda t: (t[0], t[2]))
    lines = []
    for y0, y1, x0, x1, text in items:
        for line in reversed(lines):
            overlap = min(y1, line["bottom"]) - max(y0, line["top"])
            smaller = min(y1 - y0, line["bottom"] - line["top"])
            if smaller > 0 and overlap / smaller >= overlap_ratio:
                line["words"].append((x0, x1, text))
                line["top"], line["bottom"] = min(line["top"], y0), max(line["bottom"], y1)
                break
        else:
            lines.append({"top": y0, "bottom": y1, "words": [(x0, x1, text)]})

    lines.sort(key=lambda ln: ln["top"])
    return [sorted(ln["words"], key=lambda t: t[0]) for ln in lines]


def line_text(line):
    return " ".join(t for _, _, t in line)


def page_signature(lines, header_x):
    """Why this page was accepted as a SAP audit page, or None to skip it.

    Without this gate any PDF with a dashed 9-digit number on it -- a W2,
    a transcript -- would yield junk 'student' rows if it were sitting in
    the source folder. But the gate must be generous: a continuation page
    may carry neither the report title nor the caption row, so ANY of
    these markers is enough to accept it."""
    if header_x:
        return "'ID SSN Name' caption row"
    text = " ".join(line_text(line) for line in lines).lower()
    for marker in SAP_PAGE_SIGNATURES:
        if marker in text:
            return f"the text '{marker}'"
    return None


def header_columns_in_line(line):
    """If `line` IS the "ID SSN Name" caption row, return the x of each
    caption; otherwise None. Those x-positions are the last-resort
    splitter when a line has no recognisable SSN to cut on."""
    tokens = [t.strip(":").lower() for _, _, t in line]
    for i in range(len(tokens) - 2):
        if tokens[i:i + 3] == ["id", "ssn", "name"]:
            return {"id": line[i][0], "ssn": line[i + 1][0], "name": line[i + 2][0]}
    return None


def find_header_columns(lines):
    """The caption row's column x-positions, or None if the page has no
    caption row (continuation pages often don't)."""
    for line in lines:
        found = header_columns_in_line(line)
        if found:
            return found
    return None


# ===========================================================================
# NAME BOUNDARY
# ===========================================================================
def find_name_stop(tokens):
    """Where the name ends within `tokens`. Returns (stop_index, reason).
    Every rule is evaluated and the EARLIEST stop wins, so a caption late
    on the line can't pull the name past a digit that came first."""
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
        # PDF text extraction sometimes fuses a tightly kerned two-word
        # caption into one token ("AcademicProgram").
        if width > 1:
            fused = "".join(caption)
            for i, token in enumerate(norm):
                if token == fused:
                    offer(i, label)
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


# ===========================================================================
# NAME SPLITTING
# ===========================================================================
def split_name(name, notes):
    """Split a prefix-free name into (first, middle, last) per the rules
    in the module docstring."""
    tokens = name.split()
    if not tokens:
        notes.append("no name text found after the SSN")
        return "", "", ""

    if any("," in t for t in tokens):
        notes.append("name contains a comma -- it may be printed 'Last, First' rather than "
                     "'First M. Last'; split applied as if 'First M. Last'")

    if tokens[-1].strip(".,").lower() in NAME_SUFFIXES:
        notes.append(f"name ends in a suffix ('{tokens[-1]}') -- left attached to the last name")

    if len(tokens) == 1:
        notes.append("name is a single word -- put in First Name, Last Name left blank")
        return tokens[0], "", ""

    # A middle initial only counts if it isn't the first token and isn't
    # the last -- a trailing lone letter ("Mroz Dana K") has no surname
    # after it to split off, so it can't be the split point.
    start = next((i for i in range(1, len(tokens) - 1) if INITIAL_RE.match(tokens[i])), None)

    if start is None:
        if INITIAL_RE.match(tokens[-1]):
            notes.append(f"name ends in a lone initial ('{tokens[-1]}') with nothing after it -- "
                         "treated as part of the last name, not as a middle initial")
        # No initial: first token is the first name, the remainder is the
        # last name.
        return tokens[0], "", " ".join(tokens[1:])

    # Consecutive initials ("Jane D. K. Smith") are all middle.
    end = start
    while end + 1 < len(tokens) - 1 and INITIAL_RE.match(tokens[end + 1]):
        end += 1

    return " ".join(tokens[:start]), " ".join(tokens[start:end + 1]), " ".join(tokens[end + 1:])


def strip_prefix(name):
    """Returns (prefix, name_without_prefix). Handles both 'Mrs. Jane'
    and the fused 'Mrs.Jane' that tight kerning can produce."""
    match = HONORIFIC_RE.match(name)
    if not match:
        return "", name.strip()
    return match.group(1), name[match.end():].strip()


# ===========================================================================
# STUDENT LINE PARSING
# ===========================================================================
def split_id_ssn_name(line, header_x):
    """Cut a student line into (id_text, ssn_text, name_text, note).
    Returns None if the line carries no student identity at all.

    Three strategies, most reliable first. The SSN is returned verbatim --
    never reformatted -- so the output matches the printed document."""
    # The caption row sits in exactly the columns strategy 3 looks for, so
    # it would otherwise be harvested as a student named "Name".
    if header_columns_in_line(line):
        return None

    text = line_text(line)

    match = SSN_DASHED_RE.search(text)
    if match:
        return text[:match.start()].strip(), match.group(0), text[match.end():].strip(), ""

    match = SSN_PLAIN_RE.search(text)
    if match:
        return text[:match.start()].strip(), match.group(1), text[match.end():].strip(), ""

    # Nothing SSN-shaped. Only fall back to the header's column positions
    # when the line otherwise looks like a student line -- otherwise this
    # would fire on every row of the course table.
    if not header_x:
        return None

    lowered = text.lower()
    # A word sitting in the ID column and another in the Name column is
    # itself strong evidence: the course rows are indented well right of
    # the ID column, so they don't line up this way.
    in_id_column = any(abs(x0 - header_x["id"]) <= 3 for x0, _, _ in line)
    in_name_column = any(x0 >= header_x["name"] - 3 and re.search(r"[A-Za-z]", t)
                         for x0, _, t in line)
    looks_like_student = (bool(HONORIFIC_RE.match(text))
                          or "academic program" in lowered
                          or (in_id_column and in_name_column))
    if not looks_like_student:
        return None

    id_text = " ".join(t for x0, _, t in line if x0 < header_x["ssn"] - 2).strip()
    ssn_text = " ".join(t for x0, _, t in line
                        if header_x["ssn"] - 2 <= x0 < header_x["name"] - 2).strip()
    name_text = " ".join(t for x0, _, t in line if x0 >= header_x["name"] - 2).strip()

    # When no SSN is printed, the name (or its "Ms." prefix) drifts left
    # into the SSN column. Anything with no digit in it is not an SSN --
    # hand it back to the name rather than filing it as one.
    if ssn_text and not re.search(r"\d", ssn_text):
        name_text = f"{ssn_text} {name_text}".strip()
        ssn_text = ""

    note = ("no dashed or space-delimited 9-digit SSN on this line -- ID/SSN/Name split using the "
            "'ID SSN Name' header column positions instead; verify this row")
    return id_text, ssn_text, name_text, note


def parse_student_line(line, header_x):
    """One student row, or None if this line isn't a student line."""
    split = split_id_ssn_name(line, header_x)
    if split is None:
        return None

    id_text, ssn_text, name_text, note = split

    # An SSN-shaped number with neither an ID before it nor a name after
    # it isn't a student line -- it's an SSN printed somewhere else on the
    # page. Skip it rather than emitting a near-empty row. (An ID with no
    # name, or a name with no ID, IS emitted -- that's a real anomaly the
    # reviewer should see.)
    if not id_text and not name_text:
        return None

    notes = [note] if note else []

    row = {col: "" for col in OUTPUT_COLUMNS}
    row["SSN"] = ssn_text
    if not ssn_text:
        notes.append("SSN column was empty")

    # The ID is whatever sits left of the SSN. More than one token there
    # means something extra was printed in that space, so keep the token
    # nearest the SSN and say so rather than gluing them together.
    id_tokens = id_text.split()
    if not id_tokens:
        notes.append("no ID printed to the left of the SSN")
    else:
        row["ID"] = id_tokens[-1]
        if len(id_tokens) > 1:
            notes.append(f"{len(id_tokens)} tokens found left of the SSN "
                         f"('{id_text}') -- took the one nearest the SSN as the ID")

    name_tokens = name_text.split()
    stop, reason = find_name_stop(name_tokens)
    row["Name Boundary"] = reason
    if reason.startswith("a token containing") or reason.startswith("an unrecognised caption"):
        notes.append(f"name ended at {reason} rather than the 'Academic Program' caption -- "
                     "verify the name is complete and has nothing extra")

    prefix, full_name = strip_prefix(" ".join(name_tokens[:stop]))
    row["Prefix"] = prefix
    row["Full Name"] = full_name
    row["First Name"], row["Middle"], row["Last Name"] = split_name(full_name, notes)

    row["Extraction Notes"] = "; ".join(notes)
    return row


# ===========================================================================
# PER-PDF DRIVER -- one row per student line
# ===========================================================================
def process_pdf(path: Path):
    doc = fitz.open(str(path))
    rows, image_only_pages, unrecognised_pages = [], [], []

    for page_num, page in enumerate(doc, start=1):
        if len(page.get_text().strip()) < MIN_TEXT_CHARS_PER_PAGE:
            image_only_pages.append(page_num)
            continue

        lines = group_words_into_lines(page.get_text("words"))
        header_x = find_header_columns(lines)
        if page_signature(lines, header_x) is None:
            unrecognised_pages.append(page_num)
            continue

        for line in lines:
            try:
                row = parse_student_line(line, header_x)
            except Exception as e:
                row = {col: "" for col in OUTPUT_COLUMNS}
                row["Extraction Notes"] = f"ERROR: {type(e).__name__}: {e}"
            if row is None:
                continue
            row["File Name"] = path.name
            row["Page Number"] = page_num
            rows.append(row)

    doc.close()
    return rows, image_only_pages, unrecognised_pages


# ===========================================================================
# DEBUG (safe, PII-free diagnostic dump)
# ===========================================================================
KNOWN_DEBUG_LABELS = [
    "ID", "SSN", "Name", "Academic Program", "SAP Type", "Excluded Remedial Credits",
    "Course Name", "Term/Dt", "Grd", "Cum", "Eval", "Credits", "Grade Pts",
    "Report Options", "Batch ID", "Satisfactory Academic Progress Audit Report",
    "Detail of Results by Student by SAP Type", "Page", "Att", "Pgm", "Earn", "Cmpl", "GPA",
]


def mask_shape(s):
    return re.sub(r"[A-Za-z]", "X", re.sub(r"\d", "#", s))


def mask_line_except_labels(line):
    tokens = [t for _, _, t in line]
    lowered = [t.strip(":").lower() for t in tokens]
    is_label = [False] * len(tokens)
    for label in sorted(KNOWN_DEBUG_LABELS, key=len, reverse=True):
        ltoks = label.lower().split()
        width = len(ltoks)
        for i in range(len(tokens) - width + 1):
            if not any(is_label[i:i + width]) and lowered[i:i + width] == ltoks:
                for k in range(i, i + width):
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
    header_x = find_header_columns(lines)
    signature = page_signature(lines, header_x)
    print(f"'ID SSN Name' header row: {'found' if header_x else 'NOT FOUND'}")
    print(f"recognised as a SAP audit page: "
          f"{'yes, via ' + signature if signature else 'NO -- this page would be skipped'}")

    found = 0
    for i, line in enumerate(lines):
        row = parse_student_line(line, header_x)
        if row is None:
            continue
        found += 1
        print(f"  student line [{i}]")
        for key in ["ID", "SSN", "Prefix", "Full Name", "First Name", "Middle", "Last Name"]:
            print(f"    {key}: {'(found)' if row[key] else '(blank)'}")
        print(f"    name boundary: {row['Name Boundary']}")
        if row["Extraction Notes"]:
            print(f"    notes: {row['Extraction Notes']}")
    print(f"student lines detected: {found}")

    print(f"lines detected: {len(lines)}")
    for i, line in enumerate(lines):
        print(f"  [{i:>3}] {mask_line_except_labels(line)}")
    doc.close()


# ===========================================================================
# DIAGNOSE -- why did a file yield no students? (PII-free)
# ===========================================================================
def diagnose(target: Path):
    """Per-file account of what the extractor saw and where it stopped.
    Prints no names, IDs or SSNs -- only counts, captions and masked
    shapes -- so it can be pasted into a ticket safely."""
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
                print(f"  page {page_num}: SKIPPED -- only {chars} chars of text "
                      f"(image-only page; it needs OCR first)")
                continue

            lines = group_words_into_lines(page.get_text("words"))
            header_x = find_header_columns(lines)
            signature = page_signature(lines, header_x)
            if signature is None:
                print(f"  page {page_num}: SKIPPED -- not recognised as a SAP audit page "
                      f"(no 'ID SSN Name' caption row and none of: "
                      f"{', '.join(SAP_PAGE_SIGNATURES)})")
                continue

            students = [r for r in (parse_student_line(ln, header_x) for ln in lines) if r]
            total += len(students)
            print(f"  page {page_num}: {len(students)} student line(s); "
                  f"recognised via {signature}; {len(lines)} text line(s)")

            # When a page yields nothing, show what nearly matched -- that is
            # almost always enough to see what the layout is doing.
            if not students:
                near = [(i, ln) for i, ln in enumerate(lines)
                        if SSN_DASHED_RE.search(line_text(ln))
                        or SSN_PLAIN_RE.search(line_text(ln))
                        or HONORIFIC_RE.match(line_text(ln))]
                if not near:
                    print("      no line on this page contains an SSN-shaped value or a "
                          "Mr./Mrs./Ms./Dr. prefix at all")
                for i, ln in near[:5]:
                    print(f"      near-miss line [{i}]: {mask_line_except_labels(ln)}")

        print(f"  TOTAL: {total} student line(s) in this file\n")
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

    all_rows, no_students, image_only, not_sap = [], [], [], []
    with tqdm(pdfs, desc="Extracting", unit="pdf", ncols=100) as pbar:
        for pdf_path in pbar:
            pbar.set_postfix_str(pdf_path.name)
            rows, image_pages, other_pages = process_pdf(pdf_path)
            if not rows:
                no_students.append(pdf_path.name)
            if image_pages:
                image_only.append(f"{pdf_path.name} (page(s) {', '.join(map(str, image_pages))})")
            if other_pages:
                not_sap.append(f"{pdf_path.name} ({len(other_pages)} page(s))")
            all_rows.extend(rows)
            status_callback(f"Processed {pdf_path.name} ({len(all_rows)} student(s) so far)")

    if not all_rows:
        print("No student lines found in any file -- nothing to write.")
        status_callback("Done. No student lines found -- nothing written.")
        return False

    frame = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS)
    # ID and SSN go out as text so Excel can't strip a leading zero or
    # turn a running 9-digit SSN into a number in scientific notation.
    for col in ["ID", "SSN"]:
        frame[col] = frame[col].astype(str)
    frame.to_excel(output_path, index=False)

    flagged = sum(1 for r in all_rows if r["Extraction Notes"])
    unique = len({(r["ID"], r["SSN"]) for r in all_rows})
    print(f"\nDone. {len(all_rows)} student row(s) ({unique} unique ID+SSN) "
          f"from {len(pdfs)} file(s) -> {output_path}")

    boundaries = {}
    for r in all_rows:
        boundaries[r["Name Boundary"]] = boundaries.get(r["Name Boundary"], 0) + 1
    print("Name boundary used:")
    for reason, count in sorted(boundaries.items(), key=lambda kv: -kv[1]):
        marker = "" if reason == PREFERRED_BOUNDARY else "   <-- not the 'Academic Program' caption"
        print(f"  {count:>5}  {reason}{marker}")

    if flagged:
        print(f"{flagged} row(s) have a non-empty 'Extraction Notes' -- spot-check those.")
    if no_students:
        print(f"{len(no_students)} file(s) had no student line: {', '.join(no_students)}")
        print("   -> run with --diagnose <folder> to see, per page, why nothing matched "
              "(the output is PII-free).")
    if image_only:
        print(f"Image-only page(s) skipped (no text layer -- OCR them first): {'; '.join(image_only)}")
    if not_sap:
        print(f"Page(s) skipped as not a SAP audit report (no 'ID SSN Name' header and no report "
              f"title): {'; '.join(not_sap)}")

    status_callback(f"Done. {len(all_rows)} student(s) written to {output_path.name} "
                    f"({flagged} flagged for review).")
    return True


# ===========================================================================
# TKINTER GUI
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
        folder = filedialog.askdirectory(title="Select folder containing SAP audit report PDFs", mustexist=True)
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
