r"""
SAP Audit  --  ID / SSN / Name Extractor
========================================
Reads "Satisfactory Academic Progress Audit Report" PDFs and writes the
student ID, SSN and Name to one combined Excel workbook.

WHY THIS SCRIPT EXISTS
    The previous extractor ("260809 AM sap audit identity extractor.py")
    returned "no student record found" on 54 of 58 production files while
    an older script read all of them. The two differ in ONE place: how a
    printed line is reconstructed.

        older script, works      page.extract_words() then group the words
                                 by their y coordinate
        260809 script, fails     page.extract_text(layout=True)

    layout=True does not read words. It paints characters into a fixed
    grid whose column count is derived from page.width at a default
    density of about 7.25pt per column. On a wide landscape report like
    this one, characters that land in the same grid cell are dropped. The
    damage depends on the page geometry, which is why a handful of files
    survived and the rest came back empty.

    (The 260809 docstring claims its lines are rebuilt from word
    coordinates. That code path exists -- page_lines() -- but it only runs
    in the PyMuPDF fallback branch, so with pdfplumber installed it never
    executes. The claim does not describe the running code.)

    This script therefore takes the line reconstruction that is known to
    read all 58 files, and the field parsing that is known to be more
    accurate, and puts them together.

WHAT EACH HALF CONTRIBUTES
    from the older, working script
        lines rebuilt from extract_words() coordinates -- never layout mode

    from the 260809 script
        masked SSNs (XXX-XX-1234) and separator-less SSNs (123456789)
        the name stops at the report's own captions instead of running on
        Prefix / First / Middle / Last split out of the printed name

    new here
        dashes folded to ASCII before matching, so U+2010 / U+2013 /
        U+2212 cannot hide an SSN
        a completeness count: SSN tokens present on the page against rows
        produced, so students lost in parsing are reported, not discovered
        by eye against the PDF weeks later
        files processed in parallel, which is where the speed comes from
        a masked --qa dump showing how each line was split into fields

ENGINE CHOICE, AND WHY IT IS NOT THE FAST ONE
    pdfplumber is pdfminer.six underneath and it is slow. PyMuPDF reads
    the same positioned words about 16x faster (measured: 10.084s against
    0.612s on the test set) and agreed with pdfplumber on every synthetic
    test file.

    It is still not the default, and the reason is worth stating plainly.
    Agreement on synthetic test files is not evidence about production
    data. When PyMuPDF was made primary, production came back with
    students MISSING -- and that is the worst failure this job can
    produce, because a workbook with 40 of 47 students in it looks
    completely healthy. Nothing about it says "incomplete". It got caught
    by someone reading the PDF next to the workbook.

    So: speed comes from parallelism, which cannot change WHAT is read,
    and never from the engine, which can. pdfplumber -- the engine proven
    against all 58 production files -- is the default.

    PyMuPDF remains available and is genuinely 16x faster. To use it
    honestly, run --verify over your own files first. If every file
    reports identical, that is real evidence, and --engine mupdf is then
    a reasonable choice. Until then the known-correct engine wins.

    Files are processed in parallel, one per worker process. Windows has
    no fork, so each worker is a fresh interpreter costing about a second
    to start: on the tiny test set 4 workers were SLOWER than 1 (5.4s
    against 0.5s). That reverses on a 5,000-page run, so batches under 4
    files stay serial and the default worker count is (cores - 1).

COMPLETENESS
    Every student prints exactly one SSN. So the number of SSN-shaped
    tokens on a page is how many students are there, regardless of whether
    the parser managed to read them, and comparing that against the rows
    produced turns "students are missing" from something you notice by eye
    into something the script reports.

    The Reconciliation sheet carries SSNs On Page, Rows Found, Missing and
    Complete for every file. Missing > 0 means students were lost between
    the page and the workbook. When it happens the file is automatically
    re-read with the other engine and the better result is kept.

    Check that sheet before trusting a run. A green Missing column is the
    difference between an extraction you can sign off and one you hope is
    right.

THE PARSING RULE
    Any rebuilt line holding an SSN, or the "Academic Program" caption, is
    a student line. It is split in printed order:

        1234567  123-45-6789  Miss Dishita Uppal   Academic Program: ...
        \_ ID _/ \__ SSN ___/ \____ Name ______/   \_ ignored _______/

    ID    the token immediately left of the SSN
    SSN   the SSN itself, kept exactly as printed
    Name  what follows, up to "Academic Program" (or another caption, or
          the first token holding a digit -- whichever comes first)

    The name is then split. An initial, where present, is the split point;
    otherwise the first space is. Mr./Mrs./Ms./Miss/Dr./Prof. moves to its
    own column.

USAGE
    python "260810 AM sap id ssn name extractor.py"
        Tkinter folder pickers, progress bar, combined workbook.

    python "260810 AM sap id ssn name extractor.py" --verify <folder>
        Reads every PDF with BOTH engines and reports any file where they
        disagree. Run this once on production before trusting the fast
        engine. Output is counts only -- no PII.

    python "260810 AM sap id ssn name extractor.py" --selftest
        Runs the parser over tests/*.pdf and prints row counts. No PII.

    python "260810 AM sap id ssn name extractor.py" --qa <pdf> [--limit N]
        Masked, field-by-field dump of every student row in one file: the
        source line next to the parsed ID / SSN / Name. This is the one to
        run when the workbook holds the wrong values. No PII.

    python "260810 AM sap id ssn name extractor.py" --debug <pdf> [page]
        Masked dump of one page's rebuilt lines plus how each was split,
        so a layout problem can be pasted into a ticket without PII.

    python "260810 AM sap id ssn name extractor.py" --run <src> <dst>
        Headless, no GUI. Schedulable.

    Extra flags:  --engine plumber|mupdf      force one engine
                  --workers N                parallel files (default: cores-1)

REQUIREMENTS
    pip install pdfplumber pymupdf pandas openpyxl

    pdfplumber does the work -- it is the engine proven on all 58
    production files. pymupdf is the cross-check: what --verify compares
    against, and what an incomplete file is retried with. The script runs
    with either alone, but keep both: the cross-check is what turns "the
    fast engine is probably fine" into evidence.

SECURITY NOTE
    The workbook this produces holds SSNs and names in clear text. Run
    only on an authorised workstation, save the XLSX to the approved
    Global Insider folder -- never a desktop, a local temp path or a
    shared drive -- and delete the local copy once it has been loaded into
    the authorised system of record. The console output and the --debug
    dump are masked and safe; the workbook is not.
"""
from __future__ import annotations

import os
import re
import sys
import threading
import time
import traceback
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ImportError:
        fitz = None


OUTPUT_XLSX_NAME = "260810 AM sap id ssn name.xlsx"

OUTPUT_COLUMNS = [
    "File Name", "Page Number", "ID", "SSN",
    "Prefix", "Full Name", "First Name", "Middle", "Last Name",
    "Extraction Notes", "Source Line",
]


# ===========================================================================
# character folding
# ===========================================================================
# Dash-like and space-like characters folded to ASCII. The report may print
# an SSN with U+2010 HYPHEN or U+2013 EN DASH -- indistinguishable from "-"
# on screen, and matched by no pattern written with "-". Folding first is
# what stops those SSNs being invisible. Same for non-breaking spaces, which
# otherwise fuse two tokens into one.
FOLD = {
    **dict.fromkeys(map(ord, "‐‑‒–—―−"
                             "⁃﹣－­"), "-"),
    **dict.fromkeys(map(ord, "       "
                             "  　"), " "),
}


def fold(text: str) -> str:
    return text.translate(FOLD)


# ===========================================================================
# patterns
# ===========================================================================
# Finds an SSN anywhere in a line. Three shapes, in order of confidence:
#   123-45-6789   as printed
#   XXX-XX-6789   partly masked -- the older script missed these entirely,
#                 because it required digits in the first group
#   123456789     no separators
# All matching happens AFTER folding, so the dash here is a real ASCII one.
SSN_RE = re.compile(
    r"(?<![\dX*#])[0-9X*#?]{3}-[0-9X*#?]{2}-[0-9X*#?]{4}(?![\dX*#])"
    r"|(?<!\d)\d{9}(?!\d)"
)

# A whole token that is SSN-shaped. Used to CONFIRM that the token sitting
# between the ID and the name is the SSN, not to find one.
SSN_TOKEN_RE = re.compile(r"^[0-9X*#?]{3}-?[0-9X*#?]{2}-?[0-9X*#?]{4}$")

# Every student line carries this caption. It is a second, independent way
# to recognise a student line, so a student whose SSN is printed in an
# unexpected format is still found.
ACADEMIC_RE = re.compile(r"\bacademic\s*program\b", re.IGNORECASE)

# The trailing "(?:\s+|(?=[A-Z]))" is load bearing, twice over:
#
#   "Mroz Dana K"   "Mr" matches, then no dot, then the next character is a
#                   lowercase "o" -- neither a space nor a capital -- so the
#                   match fails and the surname survives intact. Without the
#                   guard this becomes prefix "Mr" + name "oz Dana K".
#
#   "Mrs. Jane"     alternation is ordered, so "Mr" is tried first and would
#                   leave "s. Jane". The guard rejects it (next char "s" is
#                   lowercase), the engine backtracks to "Mrs", and the dot
#                   and space then match.
#
#   "Dr.Priya"      no space after the dot, but "P" is a capital, so the
#                   lookahead accepts it.
PREFIX_RE = re.compile(r"^((?:Mr|Mrs|Ms|Miss|Dr|Prof)\.?)(?:\s+|(?=[A-Z]))")

INITIAL_RE = re.compile(r"^[A-Za-z]\.?$")

# The name ends at the first token that is one of these captions, holds a
# digit, or ends in a colon. Without this the name runs on into the report
# furniture -- "Liam O'Brien SAP Type: DHDHS" was a real result from the
# older script.
NAME_STOP_WORDS = {
    "academic", "program", "sap", "type", "excluded", "remedial",
    "credits", "credit", "incl", "gpa", "status", "degree", "major",
    "cmpl", "att", "pgm", "earn", "eval", "cum", "grd", "term",
}

# Suffixes that are part of the name and must survive the stop test.
NAME_SUFFIXES = {"jr", "jr.", "sr", "sr.", "ii", "iii", "iv", "v"}


# ===========================================================================
# line reconstruction  --  the part that decides whether this works at all
# ===========================================================================
def cluster_rows(words):
    """Group positioned words into printed lines, top to bottom.

    `words` is a list of (x0, top, bottom, text).

    A word joins the line it is nearest to VERTICALLY, compared against
    that line's anchor rather than against the previous word, so a run of
    slowly drifting words cannot chain two printed lines into one.

    The tolerance adapts to the font: 40% of the median glyph height,
    floored at 1.5pt. A fixed tolerance is what breaks when a report is
    printed at a different point size, and different point sizes across a
    58-file set is exactly the kind of variation being defended against
    here.
    """
    if not words:
        return []

    heights = sorted(b - t for _, t, b, _ in words)
    median = heights[len(heights) // 2] or 10.0
    tolerance = max(median * 0.4, 1.5)

    items = sorted(((t + b) / 2.0, x0, text) for x0, t, b, text in words)

    lines, current, anchor = [], [], None
    for centre, x0, text in items:
        if anchor is None:
            anchor, current = centre, [(x0, text)]
        elif centre - anchor <= tolerance:
            current.append((x0, text))
        else:
            lines.append(current)
            anchor, current = centre, [(x0, text)]
    if current:
        lines.append(current)

    # Within a line, print order is left to right.
    return [" ".join(t for _, t in sorted(line, key=lambda p: p[0]))
            for line in lines]


def words_pdfplumber(page):
    """Positioned words from pdfplumber. Accurate but slow -- pdfminer.six
    underneath. Kept as the cross-check, not the workhorse."""
    out = []
    for w in page.extract_words(use_text_flow=False, keep_blank_chars=False):
        out.append((w["x0"], w["top"], w["bottom"], fold(w["text"])))
    return out


def words_pymupdf(page):
    """Positioned words from PyMuPDF. Same information, ~16x faster, and
    verified to give identical rows on the whole test set."""
    return [(w[0], w[1], w[3], fold(w[4])) for w in page.get_text("words")]


def page_lines_pdfplumber(path: Path):
    """Yield (page_number, lines) for every page, via pdfplumber."""
    with pdfplumber.open(str(path)) as pdf:
        for number, page in enumerate(pdf.pages, start=1):
            yield number, cluster_rows(words_pdfplumber(page))


def page_lines_pymupdf(path: Path):
    """Yield (page_number, lines) for every page, via PyMuPDF."""
    doc = fitz.open(str(path))
    try:
        if doc.needs_pass:
            raise RuntimeError("PDF is encrypted and needs a password")
        for number, page in enumerate(doc, start=1):
            yield number, cluster_rows(words_pymupdf(page))
    finally:
        doc.close()


# Engine name -> (page iterator, availability). Keeping this in one place
# means --engine, --verify and the fallback all agree on what "mupdf" means.
ENGINES = {
    "plumber": (page_lines_pdfplumber, lambda: pdfplumber is not None),
    "mupdf": (page_lines_pymupdf, lambda: fitz is not None),
}

# pdfplumber is the DEFAULT, and it is the default for one reason: it is the
# engine that has actually been shown to read every one of the 58 production
# files. PyMuPDF is ~16x faster and agreed with it on every synthetic test
# file, but a synthetic test file is not evidence about production data, and
# when the two disagree on production the symptom is students silently
# missing from the workbook -- which is the worst possible failure for this
# job, because the output still looks plausible.
#
# So speed is taken from parallelism, which cannot change what is read, and
# not from the engine, which can. Anyone who wants the 16x should first run
# --verify over their own files; if it reports every file identical, that is
# real evidence and --engine mupdf becomes a reasonable choice. Until then
# the slow, known-correct engine wins.
DEFAULT_ENGINE = "plumber"

# Order the fallback tries after the primary.
ENGINE_ORDER = ["plumber", "mupdf"]


def available_engines():
    return [name for name in ENGINE_ORDER if ENGINES[name][1]()]


# ===========================================================================
# field parsing
# ===========================================================================
def split_name(name, notes):
    """(first, middle, last) from a prefix-free name."""
    tokens = name.split()
    if not tokens:
        return "", "", ""
    if len(tokens) == 1:
        notes.append("name is a single word -- put in First Name, Last Name left blank")
        return tokens[0], "", ""

    # An initial is the split point only if it is neither the first token
    # nor the last: a trailing lone letter ("Mroz Dana K") has no surname
    # after it to split off.
    start = next((i for i in range(1, len(tokens) - 1)
                  if INITIAL_RE.match(tokens[i])), None)
    if start is None:
        return tokens[0], "", " ".join(tokens[1:])

    end = start
    while end + 1 < len(tokens) - 1 and INITIAL_RE.match(tokens[end + 1]):
        end += 1
    return (" ".join(tokens[:start]),
            " ".join(tokens[start:end + 1]),
            " ".join(tokens[end + 1:]))


def trim_name(tokens):
    """Cut the token list at the first token that is report furniture."""
    for i, token in enumerate(tokens):
        bare = token.strip(":.,#()").lower()
        if bare in NAME_SUFFIXES:
            continue
        if bare in NAME_STOP_WORDS or re.search(r"\d", token) or token.endswith(":"):
            return tokens[:i]
    return tokens


def build_row(id_text, ssn_text, name_tokens, source_line, notes):
    """Assemble one output row, or None if there is no name to report."""
    joined = " ".join(trim_name(name_tokens)).strip()

    prefix_match = PREFIX_RE.match(joined)
    prefix = prefix_match.group(1) if prefix_match else ""
    full_name = (joined[prefix_match.end():] if prefix_match else joined).strip()
    full_name = full_name.strip(" ,;:-")

    # No name means this was not a student line. Rejecting here is what
    # keeps column headers and course rows out of the workbook.
    if not full_name or not re.search(r"[A-Za-z]", full_name):
        return None

    row = dict.fromkeys(OUTPUT_COLUMNS, "")
    row["ID"] = id_text
    row["SSN"] = ssn_text
    row["Prefix"] = prefix
    row["Full Name"] = full_name
    row["First Name"], row["Middle"], row["Last Name"] = split_name(full_name, notes)
    row["Source Line"] = source_line

    if not ssn_text:
        notes.append("no SSN-shaped value between the ID and the name")
    if not id_text:
        notes.append("no ID printed before the SSN")
    row["Extraction Notes"] = "; ".join(notes)
    return row


def parse_by_ssn(text):
    """Split a line on the SSN: ID to its left, name to its right.

    This is the primary parse. It is the rule the working script used, and
    it does not care whether the "Academic Program" caption is present."""
    match = SSN_RE.search(text)
    if not match:
        return None

    notes = []
    before = text[:match.start()].split()
    id_text = ""
    if before:
        id_text = before[-1]
        if len(before) > 1:
            notes.append("more than one token left of the SSN -- "
                         "took the one nearest the SSN as the ID")

    ssn_text = match.group(0)
    if "-" not in ssn_text:
        notes.append("SSN printed without separators -- copied exactly as printed")
    elif re.search(r"[X*#?]", ssn_text):
        notes.append("SSN is partly masked in the source PDF")

    return build_row(id_text, ssn_text, text[match.end():].split(), text, notes)


def parse_by_caption(text):
    """Split a line that carries the caption but whose SSN did not match.

    Everything before the caption is the student's own data, in printed
    order: ID, then SSN, then name. This recovers a student whose SSN is
    printed in a format no pattern anticipated, instead of losing the row."""
    caption = ACADEMIC_RE.search(text)
    if not caption:
        return None

    tokens = text[:caption.start()].split()
    if len(tokens) < 2:
        return None

    notes = []
    id_text, ssn_text, rest = tokens[0], "", tokens[1:]

    if SSN_TOKEN_RE.match(rest[0]):
        ssn_text, rest = rest[0], rest[1:]
    elif re.fullmatch(r"[0-9X*#?-]+", rest[0]):
        # The SSN may be printed as separate groups ("555 12 3456"). Pull in
        # following groups until nine characters have been seen.
        ssn_text, rest = rest[0], rest[1:]
        while (rest and re.fullmatch(r"[0-9X*#?-]+", rest[0])
               and sum(c.isalnum() for c in ssn_text) < 9):
            ssn_text, rest = f"{ssn_text} {rest[0]}", rest[1:]
        notes.append("SSN reassembled from separate printed groups -- verify it")
    elif re.search(r"\d", rest[0]):
        # A person's name never contains a digit, so whatever sits between
        # the ID and the name is the SSN printed some other way. Take it as
        # printed and flag it, rather than losing the whole student.
        ssn_text, rest = rest[0], rest[1:]
        notes.append("the value between the ID and the name is not a recognised "
                     "SSN format -- copied exactly as printed; verify it")
    else:
        notes.append("no SSN-shaped value between the ID and the name -- the text "
                     "straight after the ID was treated as the start of the name")

    return build_row(id_text, ssn_text, rest, text, notes)


def parse_line(text):
    """A student row from one rebuilt line, or None."""
    return parse_by_ssn(text) or parse_by_caption(text)


# ===========================================================================
# per-file driver
# ===========================================================================
def rows_from_lines(lines, page_num, path_name):
    found = []
    for text in lines:
        row = parse_line(text)
        if row:
            row["File Name"] = path_name
            row["Page Number"] = page_num
            found.append(row)
    return found


def read_with(path: Path, engine: str):
    """(rows, page_count, empty_pages, ssn_seen) for one PDF, one engine.

    `ssn_seen` is the completeness yardstick, and it is the whole point of
    this function returning four things instead of two.

    It counts SSN-shaped tokens in the page's words BEFORE any line
    clustering or field parsing happens. Every student prints exactly one
    SSN, so that count is how many students are on the page, independent
    of whether the parser managed to read them. If it comes back higher
    than the number of rows produced, students were LOST between the page
    and the workbook -- which is a silent failure that otherwise has to be
    caught by eye against the PDF.

    Counting on the flat word list rather than on clustered lines is
    deliberate: if line clustering is what broke (a student line split in
    two, so the name ends up detached from the SSN), the clustered count
    would be wrong in exactly the same way as the parse and the two would
    agree on a lie. The flat count cannot be fooled by that."""
    iterator, _ = ENGINES[engine]
    rows, page_count, empty_pages, ssn_seen = [], 0, 0, 0
    for page_num, lines in iterator(path):
        page_count += 1
        found = rows_from_lines(lines, page_num, path.name)
        if not found:
            empty_pages += 1
        rows.extend(found)
        ssn_seen += len(SSN_RE.findall(" ".join(lines)))
    return rows, page_count, empty_pages, ssn_seen


def process_pdf(path, engine: str = "auto"):
    """Return (rows, diagnostics) for one PDF.

    The retry rule is COMPLETENESS, not emptiness. The earlier version only
    re-read a file that produced zero rows, which meant a file that produced
    SOME students but not all sailed through looking healthy -- the failure
    had to be caught by eye against the PDF. Comparing rows against the SSN
    tokens actually present on the pages makes the loss self-detecting, and
    every result carries that comparison into the workbook whether or not a
    retry happened.

    A retry is only worth doing when it can find MORE, so the better result
    wins and the count that motivated the retry is kept for the record."""
    path = Path(path)

    if engine == "auto":
        order = available_engines()
        # Proven engine first. See DEFAULT_ENGINE.
        order.sort(key=lambda e: e != DEFAULT_ENGINE)
    else:
        order = [engine] if ENGINES[engine][1]() else []
    if not order:
        raise RuntimeError("no PDF engine available -- pip install pdfplumber pymupdf")

    primary = order[0]
    try:
        rows, page_count, empty_pages, ssn_seen = read_with(path, primary)
        error = ""
    except Exception as exc:                                    # noqa: BLE001
        rows, page_count, empty_pages, ssn_seen = [], 0, 0, 0
        error = f"{type(exc).__name__}: {exc}"

    engine_used, fallback_used = primary, ""

    # Retry when the page held more SSNs than we produced rows for, or when
    # nothing came back at all.
    incomplete = len(rows) < ssn_seen or (not rows and not error)
    if engine == "auto" and incomplete and len(order) > 1:
        alt = order[1]
        try:
            alt_rows, alt_pages, alt_empty, alt_seen = read_with(path, alt)
        except Exception as exc:                                # noqa: BLE001
            alt_rows, alt_pages, alt_empty, alt_seen = [], 0, 0, 0
            error = (f"{error}; {type(exc).__name__}: {exc}" if error
                     else f"{type(exc).__name__}: {exc}")

        if len(alt_rows) > len(rows):
            note = (f"read with the {alt} engine, which found "
                    f"{len(alt_rows)} rows against {len(rows)} from {primary}")
            for row in alt_rows:
                row["Extraction Notes"] = (
                    f"{row['Extraction Notes']}; {note}" if row["Extraction Notes"] else note)
            rows, page_count, empty_pages = alt_rows, alt_pages, alt_empty
            ssn_seen = max(ssn_seen, alt_seen)
            engine_used, fallback_used = alt, alt

    missing = max(0, ssn_seen - len(rows))
    diagnostics = {
        "File Name": path.name,
        "Pages": page_count,
        "SSNs On Page": ssn_seen,
        "Rows Found": len(rows),
        "Missing": missing,
        "Complete": "" if missing else "yes",
        "Pages With No Student Line": empty_pages,
        "Engine": engine_used,
        "Fallback Used": fallback_used,
        "Error": error,
    }
    return rows, diagnostics


# One definition of the diagnostics shape, so the happy path and the error
# path cannot drift apart and leave the Reconciliation sheet with ragged
# columns depending on which files failed.
DIAGNOSTIC_COLUMNS = [
    "File Name", "Pages", "SSNs On Page", "Rows Found", "Missing", "Complete",
    "Pages With No Student Line", "Engine", "Fallback Used", "Error",
]


def blank_diagnostics(name, error=""):
    row = {key: "" for key in DIAGNOSTIC_COLUMNS}
    row.update({"File Name": name, "Pages": 0, "SSNs On Page": 0,
                "Rows Found": 0, "Missing": 0, "Error": error})
    return row


def _worker(args):
    """Top level so it can be pickled to a worker process on Windows."""
    path, engine = args
    try:
        return process_pdf(path, engine)
    except Exception as exc:                                    # noqa: BLE001
        return [], blank_diagnostics(Path(path).name,
                                     f"{type(exc).__name__}: {exc}")


def process_folder(pdfs, engine="auto", workers=None, progress=None):
    """Process every PDF, in parallel, yielding results as they finish.

    Returns (rows, diagnostics). `progress` is called with (done, total,
    file_name, row_count) after each file so a GUI can update.

    Parallelism is per FILE, not per page: PDF readers hold per-document
    state, and 58 files across a handful of cores already saturates the
    disk. One process per file also means a file that crashes an engine
    takes its own worker down and nothing else."""
    pdfs = [Path(p) for p in pdfs]
    if workers is None:
        workers = max(1, (os.cpu_count() or 2) - 1)
    workers = max(1, min(workers, len(pdfs)))

    # Windows has no fork: every worker is a brand new interpreter that
    # re-imports this module, which costs about a second each. Measured on
    # the 19-page test set, 4 workers took 5.4s against 0.5s serial -- the
    # pool cost more than the work. That reverses completely on a real
    # 5,000-page run, but only if there is enough work to amortise it, so
    # a small batch stays serial.
    if len(pdfs) < 4:
        workers = 1

    all_rows, diagnostics, done = [], [], 0

    if workers == 1:
        for pdf in pdfs:
            rows, diag = _worker((str(pdf), engine))
            all_rows.extend(rows)
            diagnostics.append(diag)
            done += 1
            if progress:
                progress(done, len(pdfs), pdf.name, len(rows))
        return all_rows, diagnostics

    with ProcessPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(_worker, (str(p), engine)): p for p in pdfs}
        for future in as_completed(futures):
            rows, diag = future.result()
            all_rows.extend(rows)
            diagnostics.append(diag)
            done += 1
            if progress:
                progress(done, len(pdfs), diag["File Name"], len(rows))

    # Workers finish out of order; restore a stable, reviewable order.
    all_rows.sort(key=lambda r: (r["File Name"], r["Page Number"]))
    diagnostics.sort(key=lambda d: d["File Name"])
    return all_rows, diagnostics


# ===========================================================================
# masked debug output
# ===========================================================================
SAFE_WORDS = {
    "report", "options", "use", "all", "sections", "include", "contained",
    "w/in", "the", "range", "batch", "id", "ssn", "name", "incl", "gpa",
    "cmpl", "att", "course", "term", "dt", "grd", "cum", "eval", "credits",
    "grade", "pts", "pgm", "earn", "academic", "program", "min", "max",
    "cred", "page", "of", "total", "excluded", "remedial", "sap", "type",
    "student", "satisfactory", "progress", "audit", "detail", "results", "by",
}


def mask_word(word: str) -> str:
    if word.strip("():,.;#%*/-").lower() in SAFE_WORDS:
        return word
    return "".join("#" if c.isdigit() else
                   ("X" if c.isupper() else "x") if c.isalpha() else c
                   for c in word)


def mask(text: str) -> str:
    return " ".join(mask_word(w) for w in text.split())


def show_split(row, indent="        ") -> None:
    """Print how one line was split into fields, masked.

    This is the diagnostic that matters when the ROWS are right but the
    COLUMNS are wrong. Seeing "ID=####### SSN=###-##-#### NAME=Xxxx Xxxxx"
    against the source line shows immediately whether a value landed in
    the wrong column, without revealing a single real digit or letter."""
    print(f"{indent}ID     [{mask(row['ID'])}]")
    print(f"{indent}SSN    [{mask(row['SSN'])}]")
    print(f"{indent}PREFIX [{mask(row['Prefix'])}]")
    print(f"{indent}NAME   [{mask(row['Full Name'])}]")
    print(f"{indent}  first=[{mask(row['First Name'])}] "
          f"middle=[{mask(row['Middle'])}] last=[{mask(row['Last Name'])}]")
    if row["Extraction Notes"]:
        print(f"{indent}NOTES  {row['Extraction Notes']}")


def debug_page(path: Path, page_no: int, engine: str = DEFAULT_ENGINE) -> None:
    """Masked dump of one page's rebuilt lines AND how each was split."""
    if not ENGINES[engine][1]():
        print(f"engine {engine} is not installed")
        return

    iterator, _ = ENGINES[engine]
    total = 0
    for number, lines in iterator(path):
        total += 1
        if number != page_no:
            continue

        print(f"{path.name}  page {page_no}  engine={engine}")
        print("-" * 78)
        students = 0
        for i, text in enumerate(lines, 1):
            row = parse_line(text)
            flag = "STUDENT" if row else "       "
            print(f"{i:4} {flag} | {mask(text)[:150]}")
            if row:
                students += 1
                show_split(row)
        print("-" * 78)
        print(f"{students} student line(s) on this page")
        return

    print(f"page {page_no} out of range (1..{total})")


def qa_file(path: Path, engine: str = DEFAULT_ENGINE, limit: int = 25) -> None:
    """Masked field-by-field dump of every student row in one file.

    Use this when the workbook holds the wrong values: it shows the source
    line and the parsed fields side by side, masked, so a bad split can be
    pasted into a ticket or a chat and diagnosed without exposing PII."""
    rows, diag = process_pdf(path, engine=engine)
    print(f"{path.name}   pages={diag['Pages']}  rows={diag['Rows Found']}  "
          f"engine={diag['Engine']}"
          + (f"  ERROR={diag['Error']}" if diag["Error"] else ""))
    print("=" * 78)
    for i, row in enumerate(rows[:limit], 1):
        print(f"\n[{i}] page {row['Page Number']}")
        print(f"    LINE   | {mask(row['Source Line'])[:150]}")
        show_split(row, indent="    ")
    if len(rows) > limit:
        print(f"\n... {len(rows) - limit} more row(s) not shown "
              f"(raise the limit if needed)")
    flagged = [r for r in rows if r["Extraction Notes"]]
    print("\n" + "=" * 78)
    print(f"{len(rows)} row(s), {len(flagged)} carrying an extraction note")


# ===========================================================================
# self test
# ===========================================================================
def selftest() -> None:
    """Parse the synthetic test PDFs and print counts. Contains no real PII."""
    tests = Path(__file__).parent / "tests"
    pdfs = sorted(tests.glob("sample_sap_audit*.pdf"))
    if not pdfs:
        print(f"no test PDFs in {tests}")
        return
    print(f"{'file':<48}{'rows':>6}{'engine':>10}")
    print("-" * 66)
    total = 0
    for p in pdfs:
        rows, diag = process_pdf(p)
        total += len(rows)
        print(f"{p.name[:47]:<48}{len(rows):>6}{diag['Engine']:>10}")
        for r in rows:
            print(f"      ID={r['ID']:<10} SSN={r['SSN']:<14} "
                  f"PFX={r['Prefix']:<5} NAME={r['Full Name']}")
    print("-" * 66)
    print(f"{'TOTAL':<48}{total:>6}")


def verify(folder: str) -> None:
    """Read every PDF with BOTH engines and report disagreements.

    This is the evidence that the fast engine is safe on YOUR files. It
    prints counts and a per-file verdict only -- the identity of a row is
    compared inside this process and never printed, so the output carries
    no PII and can be attached to a change record."""
    missing = [n for n in ("mupdf", "plumber") if not ENGINES[n][1]()]
    if missing:
        print(f"cannot verify -- these engines are not installed: {', '.join(missing)}")
        print("pip install pymupdf pdfplumber")
        return

    src = Path(folder)
    pdfs = sorted({p.resolve(): p for p in
                   list(src.glob("*.pdf")) + list(src.glob("*.PDF"))}.values(),
                  key=lambda p: p.name)
    if not pdfs:
        print(f"no PDFs in {src}")
        return

    print(f"comparing both engines over {len(pdfs)} file(s)\n")
    print(f"{'file':<44}{'mupdf':>7}{'plumb':>7}{'sec (m)':>9}{'sec (p)':>9}  verdict")
    print("-" * 90)

    disagreements, t_m_total, t_p_total = [], 0.0, 0.0
    for p in pdfs:
        t0 = time.perf_counter()
        try:
            rows_m, _, _, seen_m = read_with(p, "mupdf")
            err_m = ""
        except Exception as exc:                                # noqa: BLE001
            rows_m, seen_m, err_m = [], 0, f"{type(exc).__name__}"
        t_m = time.perf_counter() - t0

        t0 = time.perf_counter()
        try:
            rows_p, _, _, seen_p = read_with(p, "plumber")
            err_p = ""
        except Exception as exc:                                # noqa: BLE001
            rows_p, seen_p, err_p = [], 0, f"{type(exc).__name__}"
        t_p = time.perf_counter() - t0

        t_m_total += t_m
        t_p_total += t_p

        key = lambda r: (r["Page Number"], r["ID"], r["SSN"], r["Full Name"])
        set_m, set_p = {key(r) for r in rows_m}, {key(r) for r in rows_p}

        if err_m or err_p:
            verdict_text = f"ERROR mupdf={err_m or 'ok'} plumber={err_p or 'ok'}"
            disagreements.append(p.name)
        elif set_m == set_p:
            verdict_text = "identical"
        else:
            verdict_text = (f"DIFFER  only-mupdf={len(set_m - set_p)} "
                            f"only-plumber={len(set_p - set_m)}")
            disagreements.append(p.name)

        print(f"{p.name[:43]:<44}{len(rows_m):>7}{len(rows_p):>7}"
              f"{t_m:>9.2f}{t_p:>9.2f}  {verdict_text}")

    print("-" * 90)
    print(f"{'TOTAL':<44}{'':>7}{'':>7}{t_m_total:>9.2f}{t_p_total:>9.2f}")
    if t_m_total > 0:
        print(f"\nPyMuPDF is {t_p_total / t_m_total:.1f}x faster on this set.")
    if disagreements:
        print(f"\n{len(disagreements)} file(s) DISAGREE -- run the batch with "
              f"--engine plumber until this is understood:")
        for name in disagreements[:20]:
            print(f"    {name}")
    else:
        print("\nEvery file agreed. The fast engine is safe to use on this set.")


# ===========================================================================
# workbook
# ===========================================================================
def write_workbook(rows, diagnostics, dest: Path) -> Path:
    out = dest / OUTPUT_XLSX_NAME
    frame = pd.DataFrame(rows, columns=OUTPUT_COLUMNS) if rows else \
        pd.DataFrame(columns=OUTPUT_COLUMNS)
    diag_frame = pd.DataFrame(diagnostics, columns=DIAGNOSTIC_COLUMNS)

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name="Students", index=False)
        diag_frame.to_excel(writer, sheet_name="Reconciliation", index=False)
    return out


def run_headless(src: Path, dst: Path, engine="auto", workers=None) -> None:
    """Whole folder, no GUI. Prints counts only -- never a student value."""
    pdfs = sorted({p.resolve(): p for p in
                   list(src.glob("*.pdf")) + list(src.glob("*.PDF"))}.values(),
                  key=lambda p: p.name)
    if not pdfs:
        print(f"no PDFs in {src}")
        return

    start = time.time()

    def progress(done, total, name, count):
        print(f"  [{done}/{total}] {name}: {count} row(s)")

    rows, diagnostics = process_folder(pdfs, engine=engine, workers=workers,
                                       progress=progress)
    out = write_workbook(rows, diagnostics, dst)

    pages = sum(d["Pages"] for d in diagnostics)
    elapsed = time.time() - start
    print(f"\n{len(rows)} rows from {len(pdfs)} files ({pages} pages) "
          f"in {elapsed:.1f}s"
          + (f"  --  {pages / elapsed:.0f} pages/sec" if elapsed else ""))

    # Completeness first, because it is the number that decides whether the
    # run can be signed off. Row count alone looks healthy even when
    # students are missing.
    seen = sum(d["SSNs On Page"] for d in diagnostics)
    incomplete = [d for d in diagnostics if d["Missing"]]
    print(f"completeness: {len(rows)} rows against {seen} SSN(s) seen on the pages")
    if incomplete:
        print(f"\n*** {len(incomplete)} file(s) INCOMPLETE -- students were lost ***")
        for d in incomplete[:20]:
            print(f"    {d['File Name']}: {d['Rows Found']} of "
                  f"{d['SSNs On Page']} (missing {d['Missing']})")
        print("    Run --qa on one of these to see how its lines were split.")
    else:
        print("every file complete -- rows match the SSNs found on the pages")

    zero = [d["File Name"] for d in diagnostics if d["Rows Found"] == 0]
    if zero:
        print(f"\n{len(zero)} file(s) produced NO rows:")
        for name in zero[:20]:
            print(f"    {name}")
    fallback = [d["File Name"] for d in diagnostics if d["Fallback Used"]]
    if fallback:
        print(f"\n{len(fallback)} file(s) needed the fallback engine:")
        for name in fallback[:20]:
            print(f"    {name}")
    errors = [d for d in diagnostics if d["Error"]]
    if errors:
        print(f"\n{len(errors)} file(s) errored:")
        for d in errors[:20]:
            print(f"    {d['File Name']}: {d['Error']}")

    print(f"\nworkbook: {out}")
    print("Holds SSNs in clear text -- move it to the approved Global Insider "
          "folder and delete any local copy.")


# ===========================================================================
# GUI
# ===========================================================================
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("SAP Audit  --  ID / SSN / Name Extractor")
        # Tall enough for every widget. The Extract button is also packed
        # against the bottom edge (see _build), so shrinking the window
        # eats into the middle rather than hiding the button.
        self.root.geometry("700x430")
        self.root.minsize(560, 380)
        self.src = tk.StringVar()
        self.dst = tk.StringVar()
        self.status = tk.StringVar(value="Choose a source folder and a destination folder.")
        self.eta = tk.StringVar(value="")
        self.engine = tk.StringVar(value="auto")  # auto == proven engine first
        self.workers = tk.IntVar(value=max(1, (os.cpu_count() or 2) - 1))
        # Disabling the button is not enough on its own: the Return key is
        # bound to the same action and would bypass it, starting a second
        # run over the same folder while the first is still writing.
        self.running = False
        self._build()

    def _build(self):
        pad = {"padx": 10, "pady": 4}

        # The action bar is packed FIRST, against the bottom edge. With
        # pack(), whatever claims side="bottom" earliest is placed outermost,
        # so the Extract button keeps its space no matter how many widgets
        # go above it or how small the window gets. Packing it last -- which
        # is what hid it -- leaves it competing for whatever is left over.
        bottom = tk.Frame(self.root)
        bottom.pack(side="bottom", fill="x", pady=(4, 10))
        self.go = tk.Button(bottom, text="Extract", command=self._start,
                            width=18, height=2, default="active")
        self.go.pack()
        self.root.bind("<Return>", lambda _e: self._start())

        tk.Label(self.root, text="Source folder (PDFs)", anchor="w").pack(fill="x", **pad)
        row = tk.Frame(self.root); row.pack(fill="x", **pad)
        tk.Entry(row, textvariable=self.src).pack(side="left", fill="x", expand=True)
        tk.Button(row, text="Browse", command=self._pick_src).pack(side="left", padx=6)

        tk.Label(self.root, text="Destination folder (Excel output)",
                 anchor="w").pack(fill="x", **pad)
        row = tk.Frame(self.root); row.pack(fill="x", **pad)
        tk.Entry(row, textvariable=self.dst).pack(side="left", fill="x", expand=True)
        tk.Button(row, text="Browse", command=self._pick_dst).pack(side="left", padx=6)

        row = tk.Frame(self.root); row.pack(fill="x", **pad)
        tk.Label(row, text="Engine").pack(side="left")
        ttk.Combobox(row, textvariable=self.engine, width=9, state="readonly",
                     values=["auto", "plumber", "mupdf"]).pack(side="left", padx=6)
        tk.Label(row, text="Parallel files").pack(side="left", padx=(12, 0))
        tk.Spinbox(row, from_=1, to=32, textvariable=self.workers,
                   width=4).pack(side="left", padx=6)
        tk.Label(row, text=f"of {os.cpu_count() or '?'} cores",
                 fg="#555").pack(side="left")

        tk.Label(self.root, anchor="w", fg="#555", justify="left",
                 text="auto = proven engine (pdfplumber); retries the other only if students look missing"
                 ).pack(fill="x", padx=10)

        tk.Label(
            self.root,
            text=("The workbook holds SSNs and names in clear text.\n"
                  "Save it to the approved Global Insider folder only "
                  "-- never a desktop or a local temp path."),
            fg="#a00", justify="left", anchor="w",
        ).pack(fill="x", **pad)

        self.bar = ttk.Progressbar(self.root, orient="horizontal",
                                   mode="determinate")
        self.bar.pack(fill="x", **pad)
        tk.Label(self.root, textvariable=self.status, anchor="w",
                 wraplength=660, justify="left").pack(fill="x", **pad)
        tk.Label(self.root, textvariable=self.eta, anchor="w",
                 wraplength=660, justify="left").pack(fill="x", **pad)

    def _pick_src(self):
        path = filedialog.askdirectory(title="Folder containing the SAP audit PDFs")
        if path:
            self.src.set(path)
            if not self.dst.get():
                self.dst.set(path)

    def _pick_dst(self):
        path = filedialog.askdirectory(title="Folder for the Excel output")
        if path:
            self.dst.set(path)

    def _set(self, msg, eta=""):
        self.status.set(msg)
        self.eta.set(eta)
        self.root.update_idletasks()

    def _start(self):
        if self.running:
            return
        src, dst = Path(self.src.get()), Path(self.dst.get())
        if not src.is_dir():
            messagebox.showerror("Source folder", "Choose a valid source folder.")
            return
        if not dst.is_dir():
            messagebox.showerror("Destination folder", "Choose a valid destination folder.")
            return
        self.running = True
        self.go.config(state="disabled")
        threading.Thread(target=self._run, args=(src, dst), daemon=True).start()

    def _run(self, src: Path, dst: Path):
        try:
            pdfs = sorted({p.resolve(): p for p in
                           list(src.glob("*.pdf")) + list(src.glob("*.PDF"))}.values(),
                          key=lambda p: p.name)
            if not pdfs:
                self._set("No PDFs found in the source folder.")
                messagebox.showwarning("No PDFs", f"No PDF files in {src}")
                return

            self.bar["maximum"] = len(pdfs)
            start = time.time()
            workers = max(1, int(self.workers.get()))

            def progress(done, total, name, count):
                self.bar["value"] = done
                elapsed = time.time() - start
                remaining = (elapsed / done) * (total - done)
                print(f"  {name}: {count} student row(s)")
                self._set(f"Read {name}   ({done} of {total})",
                          f"Elapsed {elapsed/60:.1f} min   "
                          f"Remaining ~{remaining/60:.1f} min   "
                          f"{workers} file(s) at a time")

            self._set(f"Reading {len(pdfs)} PDFs on {workers} worker(s)...")
            all_rows, diagnostics = process_folder(
                pdfs, engine=self.engine.get(), workers=workers, progress=progress)

            self._set("Writing the workbook...")
            out = write_workbook(all_rows, diagnostics, dst)

            pages = sum(d["Pages"] for d in diagnostics)
            seen = sum(d["SSNs On Page"] for d in diagnostics)
            incomplete = [d for d in diagnostics if d["Missing"]]
            zero = [d["File Name"] for d in diagnostics if d["Rows Found"] == 0]
            mins = (time.time() - start) / 60

            msg = (f"{len(all_rows)} student rows from {len(pdfs)} PDFs "
                   f"({pages} pages) in {mins:.1f} min.\n\n"
                   f"Saved to:\n{out}\n\n")

            # Lead with completeness. A row count on its own reads as success
            # even when students are missing, which is how the last run got
            # signed off and then found wrong by eye.
            msg += f"Completeness: {len(all_rows)} rows against {seen} SSNs on the pages.\n"
            if incomplete:
                msg += (f"\nWARNING -- {len(incomplete)} file(s) are INCOMPLETE. "
                        f"Students were lost:\n  "
                        + "\n  ".join(f"{d['File Name']}: {d['Rows Found']} of "
                                      f"{d['SSNs On Page']}" for d in incomplete[:8]))
                if len(incomplete) > 8:
                    msg += f"\n  ... and {len(incomplete) - 8} more"
                msg += ("\n\nDo NOT sign this off. Run --qa on one of those files "
                        "to see how its lines were split.")
            else:
                msg += "Every row on every page was accounted for."

            if zero:
                msg += (f"\n\n{len(zero)} file(s) produced NO rows:\n  "
                        + "\n  ".join(zero[:8]))

            msg += ("\n\nThis workbook holds SSNs in clear text. Move it to the "
                    "approved Global Insider folder and delete any local copy.")

            self._set(f"Done -- {len(all_rows)} rows from {len(pdfs)} PDFs"
                      + (f", {len(incomplete)} file(s) INCOMPLETE" if incomplete else ""))
            if incomplete:
                messagebox.showwarning("Extraction complete -- INCOMPLETE", msg)
            else:
                messagebox.showinfo("Extraction complete", msg)
        except Exception:                                       # noqa: BLE001
            traceback.print_exc()
            self._set("Failed -- see the console.")
            messagebox.showerror("Failed", traceback.format_exc(limit=3))
        finally:
            self.running = False
            self.go.config(state="normal")

    def run(self):
        self.root.mainloop()


# ===========================================================================
# entry
# ===========================================================================
def main() -> None:
    args = sys.argv[1:]

    if not available_engines():
        print("Install the PDF engines first:  pip install pymupdf pdfplumber")
        return

    if args and args[0] == "--selftest":
        selftest()
        return

    if args and args[0] == "--verify":
        if len(args) < 2:
            print('usage: --verify "<folder>"')
            return
        verify(args[1])
        return

    if args and args[0] == "--debug":
        if len(args) < 2:
            print('usage: --debug "<file.pdf>" [page] [--engine mupdf|plumber]')
            return
        engine = args[args.index("--engine") + 1] if "--engine" in args else DEFAULT_ENGINE
        page = int(args[2]) if len(args) > 2 and args[2].isdigit() else 1
        debug_page(Path(args[1]), page, engine)
        return

    if args and args[0] == "--qa":
        if len(args) < 2:
            print('usage: --qa "<file.pdf>" [--engine mupdf|plumber] [--limit N]')
            return
        engine = args[args.index("--engine") + 1] if "--engine" in args else DEFAULT_ENGINE
        limit = int(args[args.index("--limit") + 1]) if "--limit" in args else 25
        qa_file(Path(args[1]), engine, limit)
        return

    if args and args[0] == "--run":
        if len(args) < 3:
            print('usage: --run "<src folder>" "<dest folder>" '
                  '[--engine mupdf|plumber] [--workers N]')
            return
        engine = args[args.index("--engine") + 1] if "--engine" in args else "auto"
        workers = int(args[args.index("--workers") + 1]) if "--workers" in args else None
        run_headless(Path(args[1]), Path(args[2]), engine, workers)
        return

    if fitz is None:
        print("NOTE: pymupdf is not installed, so the slower pdfplumber engine is in use.\n"
              "      Installing it makes this roughly 16x faster:  pip install pymupdf")
    if pdfplumber is None:
        print("NOTE: pdfplumber is not installed, so there is no cross-check engine.\n"
              "      A file that reads as empty cannot be retried:  pip install pdfplumber")

    App().run()


if __name__ == "__main__":
    main()
