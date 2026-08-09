"""
SAP Audit Identity Extractor -- Tkinter GUI
==========================================================
Pulls the student identity fields out of "Satisfactory Academic Progress
Audit Report" PDFs. Only the student header line is read -- the course
table below it (DHSC-821, term, grade, credits, ...) is ignored.

HOW A STUDENT LINE IS FOUND
    Every line of every page is read, and ANY line containing a value
    shaped like an SSN is treated as a student line. Nothing else is
    required -- no report title, no "ID SSN Name" caption row, no
    "Academic Program" caption. Earlier versions gated on those and
    silently skipped whole files whose pages didn't carry them.

    Recognised SSN shapes:
        123-45-6789     dashed (the normal case)
        XXX-XX-6789     partially masked
        123456789       running, when whitespace proves where it
                        starts and ends

    Each page is read through TWO independent extraction paths -- words
    grouped by position, and PyMuPDF's own text layer -- and the results
    are merged and de-duplicated. If one path renders a page badly the
    other usually still finds the line.

WHAT IS TAKEN FROM THAT LINE
    ID    -- everything printed to the left of the SSN.
    SSN   -- kept EXACTLY as printed, dashed or running. It is never
             reformatted, so the output matches the source document
             character for character.
    Name  -- everything to the right of the SSN, up to the end of the
             name (see below), with any Mr./Mrs./Ms./Miss/Dr. prefix
             moved to its own "Prefix" column, then split into three:

               with a middle initial   "Jane D. Smith"
                   First Name = Jane      Middle = D.     Last Name = Smith
               without one             "Jane Smith"
                   First Name = Jane      Middle = (blank) Last Name = Smith

             i.e. when an initial is present it is the split point -- what
             precedes it is the first name, what follows it is the last
             name. With no initial the first space is the split point.

WHERE THE NAME ENDS
    "Academic Program:" is the preferred boundary, but it is not on every
    line, so the name stops at whichever of these comes first, and the one
    used is recorded in the "Name Boundary" column:
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

TWO EXTRA RULES THE "ANY SSN-SHAPED LINE" TRIGGER NEEDS
    - A line must have a NAME after the SSN. Without this, a captioned
      field elsewhere on the page ("SSN: 123-45-6789") becomes a row with
      every name column blank. The number of lines rejected this way is
      printed at the end, so nothing disappears silently.
    - The same student seen by both extraction paths is merged into one
      row, keyed on file + page + ID + SSN. Where the two paths read the
      NAME differently, the better-scoring reading is kept and the
      disagreement is recorded in the Extraction Notes.

TWO CONSEQUENCES OF THE "ANY SSN-SHAPED LINE" RULE
    - A student whose SSN is blank on the report will NOT be picked up,
      because there is nothing to trigger on. The run prints a WARNING
      with a count of lines that look like students but carry no SSN, so
      this gap is visible rather than silent.
    - A non-SAP PDF sitting in the source folder (a W2, a transcript) WILL
      contribute rows, since it also contains SSNs. Point the source
      folder at SAP audit reports only.

GUI:
    - Source folder picker (scanned recursively for PDFs)
    - Destination folder picker (where the output XLSX goes)
    - Start Extraction button; progress prints to the console window

USAGE:
    python "260809 AM sap audit identity extractor.py"
    python "260809 AM sap audit identity extractor.py" --diagnose <pdf_or_folder>
    python "260809 AM sap audit identity extractor.py" --debug <pdf_or_folder> [page_number]

--diagnose answers "why did this file produce no rows?": per page it
reports the size of the text layer, how many lines each extraction path
saw, and how many student lines were found. --debug dumps every line of
one page.

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

# Requiring a period OR a following capital matters: a bare
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

OUTPUT_COLUMNS = [
    "File Name", "Page Number",
    "ID", "SSN",
    "Prefix", "Full Name", "First Name", "Middle", "Last Name",
    "Name Boundary", "Extraction Notes",
]


# ===========================================================================
# READING LINES OFF A PAGE
# ===========================================================================
def group_words_into_lines(words, overlap_ratio=0.5):
    """words: PyMuPDF page.get_text('words') output. Returns a list of
    lines, each a list of (x0, x1, text) tuples sorted left to right.

    Words are grouped by how much their vertical extents OVERLAP, not by
    rounding y into fixed buckets. Bucketing looks simpler but silently
    tears a line in half whenever its words straddle a bucket edge: real
    text layers carry sub-point jitter, so y=17.4 and y=17.6 -- the same
    printed line -- can round into different buckets and become two lines.
    On this report that put the ID and SSN in one fragment and the name in
    another, blanking every ID and dropping students outright."""
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


def normalise(text):
    return " ".join(text.split()).lower()


def page_text_lines(page):
    """Every text line on the page, from two independent extraction paths,
    de-duplicated.

    The two paths fail differently: word-grouping copes with jitter and
    multi-column layouts but depends on the word geometry being sane,
    while PyMuPDF's own text layer preserves the original line breaks but
    can interleave columns. Running both and merging means a page has to
    defeat BOTH before a student is lost. De-duplication is on whitespace-
    normalised text, so a line seen by both paths yields one row, not
    two."""
    seen, out = set(), []
    for source in (
        (line_text(ln) for ln in group_words_into_lines(page.get_text("words"))),
        page.get_text("text").splitlines(),
    ):
        for raw in source:
            text = raw.strip()
            key = normalise(text)
            if key and key not in seen:
                seen.add(key)
                out.append(text)
    return out


def line_text(line):
    return " ".join(t for _, _, t in line)


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
def find_ssn(text):
    """The first SSN-shaped value on the line, as (start, end, ssn), or
    None. Dashed wins over running: a dashed match is unambiguous, while a
    bare 9-digit run needs whitespace on both sides to be trustworthy."""
    match = SSN_DASHED_RE.search(text)
    if match:
        return match.start(), match.end(), match.group(0)
    match = SSN_PLAIN_RE.search(text)
    if match:
        return match.start(1), match.end(1), match.group(1)
    return None


def parse_student_line(text):
    """One student row from a line of text, or None if the line carries no
    SSN-shaped value (and so isn't a student line)."""
    found = find_ssn(text)
    if found is None:
        return None
    start, end, ssn_text = found

    id_text, name_text = text[:start].strip(), text[end:].strip()

    # An SSN with no name after it isn't a student line -- it's an SSN
    # printed under a caption somewhere else ("SSN: 123-45-6789"). Skip it
    # rather than emit a row with every name column blank. The count of
    # lines rejected here is reported, so nothing is dropped silently.
    if not re.search(r"[A-Za-z]", name_text):
        return None

    notes = []
    row = {col: "" for col in OUTPUT_COLUMNS}
    row["SSN"] = ssn_text

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
# MERGING THE TWO EXTRACTION PATHS
# ===========================================================================
KNOWN_CAPTION_LABELS = {" ".join(c) for c in KNOWN_TRAILING_CAPTIONS}


def row_score(row):
    """How much to trust a row, for picking between two readings of the
    same student. A name that ended on a real caption beats one that ran
    to the end of the line, which in turn beats one cut short at a digit."""
    boundary = row["Name Boundary"]
    if boundary == PREFERRED_BOUNDARY:
        rank = 3
    elif boundary in KNOWN_CAPTION_LABELS:
        rank = 2
    elif boundary == "end of line":
        rank = 1
    else:
        rank = 0
    note_count = len([n for n in row["Extraction Notes"].split("; ") if n])
    return (1 if row["Full Name"] else 0, rank, -note_count, len(row["Full Name"]))


def candidate_name(text):
    """The name-looking part of a line that carries no SSN, used only to
    count how many DISTINCT students might have been missed.

    Returns "" when there is no name on the line, which is what filters
    out the bare "Academic Program: ..." fragments the second extraction
    path produces: with the caption at token 0 the name stops at index 0
    and nothing is left."""
    body = re.sub(r"^\s*[\w-]*\d[\w-]*\s+", "", text)  # drop a leading ID-ish token
    _, rest = strip_prefix(body.strip())
    tokens = rest.split()
    stop, _ = find_name_stop(tokens)
    return " ".join(tokens[:stop]).strip().lower()


def count_missed_students(candidates, found_names):
    """How many distinct SSN-less students a page appears to hold, ignoring
    those already extracted and collapsing the two extraction paths' views
    of the same person."""
    missed = set()
    for text in candidates:
        name = candidate_name(text)
        if name and not any(name in f or f in name for f in found_names):
            missed.add(name)
    return len(missed)


def dedupe_rows(rows):
    """Collapse the same student seen by both extraction paths into one
    row. The paths often break a line at different points, so the two
    readings are not textually identical even though they describe the
    same student -- de-duplicating on the text would miss them, so this
    keys on the student instead.

    Where the two readings disagree on the NAME, the better-scoring one is
    kept and the disagreement is recorded, because that is exactly the
    case a human should look at."""
    groups, order = {}, []
    for row in rows:
        key = (row["File Name"], row["Page Number"], row["ID"], row["SSN"])
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(row)

    merged = []
    for key in order:
        variants = groups[key]
        best = max(variants, key=row_score)
        readings = {v["Full Name"] for v in variants}
        if len(readings) > 1:
            best = dict(best)
            note = ("the two text-extraction paths read this name differently "
                    f"({' | '.join(sorted(readings))}) -- verify against the PDF")
            best["Extraction Notes"] = "; ".join(x for x in [best["Extraction Notes"], note] if x)
        merged.append(best)
    return merged


# ===========================================================================
# PER-PDF DRIVER -- one row per student line
# ===========================================================================
def process_pdf(path: Path):
    doc = fitz.open(str(path))
    rows, image_only_pages, nameless, ssnless = [], [], 0, 0

    for page_num, page in enumerate(doc, start=1):
        if len(page.get_text().strip()) < MIN_TEXT_CHARS_PER_PAGE:
            image_only_pages.append(page_num)
            continue

        page_rows, candidates = [], []
        for text in page_text_lines(page):
            if find_ssn(text) is None:
                # Detection keys on the SSN, so a student whose SSN is
                # blank on the report cannot be picked up. Hold on to the
                # lines that look like students anyway, so that gap can be
                # reported rather than passing silently.
                if "academic program" in text.lower() or HONORIFIC_RE.match(text):
                    candidates.append(text)
                continue
            try:
                row = parse_student_line(text)
            except Exception as e:
                row = {col: "" for col in OUTPUT_COLUMNS}
                row["Extraction Notes"] = f"ERROR: {type(e).__name__}: {e}"
            if row is None:
                nameless += 1
                continue
            row["File Name"] = path.name
            row["Page Number"] = page_num
            page_rows.append(row)

        # Most candidates are just the other extraction path's view of a
        # student already captured above -- the same line broken at a
        # different point, with the SSN in the other fragment. Counting
        # distinct names, minus those already extracted, leaves only the
        # genuinely missed students.
        found_names = {r["Full Name"].lower() for r in page_rows if r["Full Name"]}
        ssnless += count_missed_students(candidates, found_names)
        rows.extend(page_rows)

    doc.close()
    return dedupe_rows(rows), image_only_pages, nameless, ssnless


# ===========================================================================
# MASKING (shared by --debug and --diagnose)
# ===========================================================================
KNOWN_DEBUG_LABELS = [
    "ID", "SSN", "Name", "Academic Program", "SAP Type", "Excluded Remedial Credits",
    "Course Name", "Term/Dt", "Grd", "Cum", "Eval", "Credits", "Grade Pts",
    "Report Options", "Batch ID", "Satisfactory Academic Progress Audit Report",
    "Detail of Results by Student by SAP Type", "Page", "Att", "Pgm", "Earn", "Cmpl", "GPA",
    "Section skipped", "No Verified Grade Exists",
]


def mask_shape(s):
    return re.sub(r"[A-Za-z]", "X", re.sub(r"\d", "#", s))


def mask_text_except_labels(text):
    tokens = text.split()
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


# ===========================================================================
# DEBUG -- dump one page, masked
# ===========================================================================
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

    lines = page_text_lines(page)
    students = [(i, parse_student_line(t)) for i, t in enumerate(lines)]
    students = [(i, r) for i, r in students if r]
    print(f"student lines detected: {len(students)}")
    for i, row in students:
        print(f"  student line [{i}]")
        for key in ["ID", "SSN", "Prefix", "Full Name", "First Name", "Middle", "Last Name"]:
            print(f"    {key}: {'(found)' if row[key] else '(blank)'}")
        print(f"    name boundary: {row['Name Boundary']}")
        if row["Extraction Notes"]:
            print(f"    notes: {row['Extraction Notes']}")

    print(f"lines detected: {len(lines)}")
    for i, t in enumerate(lines):
        print(f"  [{i:>3}] {mask_text_except_labels(t)}")
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
                print(f"  page {page_num}: SKIPPED -- only {chars} chars of text. This page is a "
                      f"scanned image with no text layer; it must be OCR'd before anything can "
                      f"be read off it.")
                continue

            by_words = len(group_words_into_lines(page.get_text("words")))
            by_text = len(page.get_text("text").splitlines())
            lines = page_text_lines(page)

            matched = [r for r in (parse_student_line(t) for t in lines) if r]
            for row in matched:
                row["File Name"], row["Page Number"] = path.name, page_num
            students = dedupe_rows(matched)
            total += len(students)
            print(f"  page {page_num}: {len(students)} student(s) "
                  f"(from {len(matched)} matching line(s) before merge)  "
                  f"[{chars} chars; {by_words} lines by word-position, {by_text} by text layer, "
                  f"{len(lines)} merged]")

            if not students:
                near = [t for t in lines if re.search(r"\d{3}[- ]?\d{2}[- ]?\d{4}", t)
                        or HONORIFIC_RE.match(t)]
                if not near:
                    print("      no line on this page contains anything SSN-shaped or a "
                          "Mr./Mrs./Ms./Dr. prefix -- there is nothing here to extract")
                for t in near[:5]:
                    print(f"      near-miss: {mask_text_except_labels(t)}")

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

    all_rows, no_students, image_only = [], [], []
    total_nameless, total_ssnless = 0, 0
    with tqdm(pdfs, desc="Extracting", unit="pdf", ncols=100) as pbar:
        for pdf_path in pbar:
            pbar.set_postfix_str(pdf_path.name)
            rows, image_pages, nameless, ssnless = process_pdf(pdf_path)
            if not rows:
                no_students.append(pdf_path.name)
            if image_pages:
                image_only.append(f"{pdf_path.name} (page(s) {', '.join(map(str, image_pages))})")
            total_nameless += nameless
            total_ssnless += ssnless
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
    files_with_rows = len({r["File Name"] for r in all_rows})
    print(f"\nDone. {len(all_rows)} student row(s) ({unique} unique ID+SSN) "
          f"from {files_with_rows} of {len(pdfs)} file(s) -> {output_path}")

    boundaries = {}
    for r in all_rows:
        boundaries[r["Name Boundary"]] = boundaries.get(r["Name Boundary"], 0) + 1
    print("Name boundary used:")
    for reason, count in sorted(boundaries.items(), key=lambda kv: -kv[1]):
        marker = "" if reason == PREFERRED_BOUNDARY else "   <-- not the 'Academic Program' caption"
        print(f"  {count:>5}  {reason}{marker}")

    if flagged:
        print(f"{flagged} row(s) have a non-empty 'Extraction Notes' -- spot-check those.")
    if total_nameless:
        print(f"{total_nameless} line(s) held an SSN but no name after it (e.g. a captioned "
              f"'SSN: ...' field) and were not treated as students.")
    if total_ssnless:
        print(f"\nWARNING: {total_ssnless} line(s) look like student lines (a name prefix or an "
              f"'Academic Program' caption) but carry NO SSN, so they were not extracted -- "
              f"detection keys on the SSN. Check these by hand if students can have a blank SSN.")
    if image_only:
        print(f"\nSCANNED PAGES WITH NO TEXT LAYER (nothing can be read off these until they are "
              f"OCR'd):\n  {'; '.join(image_only)}")
    if no_students:
        print(f"\n{len(no_students)} file(s) produced no rows: {', '.join(no_students)}")
        print("   -> run with --diagnose <folder> to see, per page, why. The output is PII-free.")

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
