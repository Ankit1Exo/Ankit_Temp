r"""
SAP Audit  --  Student ID / SSN / Name / Address extractor
==========================================================
Reads "Satisfactory Academic Progress Audit Report" PDFs and writes four
fields per student -- Student ID, SSN, Name, Address -- into one Excel
workbook, with a reconciliation sheet that says whether anything was lost.

TWO INDEPENDENT WAYS TO READ A PAGE
    format   Words with their x/y boxes. Printed rows are rebuilt from the
             y coordinates, and the fields are cut out using the report's
             own column bands, found by locating the ID / SSN / Name
             headings and measuring where each one sits.

    text     The page's plain text lines. No geometry at all -- fields are
             found by pattern: the SSN anchors the line, the ID is the
             token to its left, the name is what follows.

    format runs first, because column geometry can keep a name apart from
    the value printed beside it even when the two collide once the page is
    flattened to a string.

    text runs when format cannot deliver:
        - the page has no heading row to measure columns from
        - the text layer has no usable word boxes
        - format found the row but left a field empty

    They are deliberately separate code paths over separate
    representations of the page. A fallback that shares the step that
    failed is not a fallback.

    Every row records which path produced it in the Method column, so a
    run can be judged rather than assumed: "format" throughout means the
    geometry held; a page of "text" means the headings were not found and
    the looser rule did the work.

WHERE THE ADDRESS COMES FROM
    The address is looked for in three places, in this order, and the row
    records which one answered:

        column       an Address heading exists, so the words sitting in
                     that band are the address
        same line    the address is printed after the name on the student
                     line, and starts where the name stops
        below        the address is printed on the line(s) under the
                     student line

    "below" is bounded: at most 3 lines, stopping at the next student, at
    a course row, or at any report caption. A line only counts as an
    address if it looks like one -- a ZIP, a ", ST ZIP" tail, a house
    number followed by a street word, or a PO box. Guessing is worse than
    a blank cell here, because a blank is visible and a wrong address is
    not.

    Which of the three applies depends on the report, so all three are
    implemented rather than one assumed. Check the Address Source column
    on the first real run: if it says "below" everywhere, that is the
    layout, and if it is blank everywhere the address is not on the page
    at all and no amount of parsing will find it.

COMPLETENESS -- READ THIS BEFORE SIGNING OFF A RUN
    Every student prints exactly one SSN, so the number of SSN-shaped
    tokens on a page is how many students are on it, whether or not the
    parser managed to read them. The Reconciliation sheet compares that
    count against the rows produced, per file.

    Missing > 0 means students were lost between the page and the
    workbook. That is the failure that matters, because a workbook
    holding 40 of 47 students looks perfectly healthy -- nothing about it
    says "incomplete". When it happens the file is re-read with the other
    engine and the better result is kept.

USAGE
    python "260810 AM sap id ssn name address extractor.py"
        Folder pickers, progress, combined workbook.

    ... --run "<src folder>" "<dest folder>"
        Headless, schedulable. Prints counts only, never a value.

    ... --qa "<file.pdf>" [--limit N]
        Masked field-by-field dump of one file: source line next to the
        parsed ID / SSN / Name / Address, and which path produced it.
        This is the one to run when the workbook holds wrong values.

    ... --debug "<file.pdf>" [page]
        Masked dump of one page -- the heading row that was found, the
        column bands measured from it, and how each line was split.

    ... --columns "<file.pdf>" [page]
        Just the measured column bands. Run this when Method says "text"
        everywhere: it shows whether the headings were found at all.

    ... --selftest
        Parses tests/*.pdf and prints counts. Synthetic data, no PII.

    Extra flags:  --engine plumber|mupdf   force one engine
                  --workers N              parallel files (default cores-1)

REQUIREMENTS
    pip install pdfplumber pymupdf pandas openpyxl

    pdfplumber is the default engine. PyMuPDF is faster and is used as the
    cross-check: it is what an incomplete file is retried with.

SECURITY NOTE
    The workbook this writes holds SSNs, names and addresses in clear
    text. Run it only on an authorised workstation, save the XLSX to the
    approved Global Insider folder -- never a desktop, a local temp path
    or a shared drive -- and delete the local copy once it has been loaded
    into the authorised system of record.

    Console output, --qa and --debug are masked and safe to paste into a
    ticket. The workbook is not.
"""
from __future__ import annotations

import os
import re
import sys
import threading
import time
import traceback
from collections import namedtuple
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


OUTPUT_XLSX_NAME = "260810 AM sap id ssn name address.xlsx"

OUTPUT_COLUMNS = [
    "File Name", "Page Number",
    "Student ID", "SSN", "Name", "Address",
    "Method", "Address Source", "Extraction Notes", "Source Line",
]

DIAGNOSTIC_COLUMNS = [
    "File Name", "Pages", "SSNs On Page", "Rows Found", "Missing", "Complete",
    "By Format", "By Text", "With Address",
    "Pages With Columns", "Pages Without Columns",
    "Engine", "Fallback Used", "Error",
]


# ===========================================================================
# character folding
# ===========================================================================
# An SSN printed with U+2010 HYPHEN or U+2013 EN DASH is indistinguishable
# from one printed with "-" on screen, and is matched by no pattern written
# with an ASCII hyphen. Folding before anything else is what stops those SSNs
# being invisible. Non-breaking spaces get the same treatment: left alone
# they fuse two tokens into one and the ID ends up glued to the SSN.
FOLD = {
    **dict.fromkeys(map(ord, "‐‑‒–—―−⁃﹣－­"), "-"),
    **dict.fromkeys(map(ord, "        "
                             "       　"), " "),
}


def fold(text: str) -> str:
    return text.translate(FOLD)


# ===========================================================================
# patterns
# ===========================================================================
# An SSN anywhere in a line, in three shapes: as printed, partly masked
# (XXX-XX-6789), and with no separators. All matching happens after folding,
# so the dash below is a real ASCII one.
SSN_RE = re.compile(
    r"(?<![\dX*#])[0-9X*#?]{3}-[0-9X*#?]{2}-[0-9X*#?]{4}(?![\dX*#])"
    r"|(?<!\d)\d{9}(?!\d)"
)

# A whole token that is SSN-shaped. Confirms that the value sitting between
# the ID and the name IS the SSN; it is not used to go looking for one.
SSN_TOKEN_RE = re.compile(r"^[0-9X*#?]{3}-?[0-9X*#?]{2}-?[0-9X*#?]{4}$")

# An ID as printed: digits, possibly with a letter prefix or a dash.
ID_TOKEN_RE = re.compile(r"^[A-Za-z]{0,3}[-]?\d[\dA-Za-z-]*$")

ACADEMIC_RE = re.compile(r"\bacademic\s*program\b", re.IGNORECASE)

# The trailing "(?:\s+|(?=[A-Z]))" is load bearing twice over:
#   "Mroz Dana K"  "Mr" matches, no dot follows, and the next character is a
#                  lowercase "o" -- neither space nor capital -- so the match
#                  fails and the surname survives. Without the guard this
#                  becomes prefix "Mr" plus name "oz Dana K".
#   "Mrs. Jane"    alternation is ordered so "Mr" is tried first, which would
#                  leave "s. Jane". The guard rejects it, the engine
#                  backtracks to "Mrs", and the dot and space then match.
PREFIX_RE = re.compile(r"^((?:Mr|Mrs|Ms|Miss|Dr|Prof)\.?)(?:\s+|(?=[A-Z]))")

# The name ends at the first token that is one of these captions, holds a
# digit, or ends in a colon. Without it the name runs on into the report
# furniture -- "Liam O'Brien SAP Type: DHDHS" is the shape of the mistake.
NAME_STOP_WORDS = {
    "academic", "program", "sap", "type", "excluded", "remedial",
    "credits", "credit", "incl", "gpa", "status", "degree", "major",
    "cmpl", "att", "pgm", "earn", "eval", "cum", "grd", "term", "dt",
    "course", "section", "skipped", "total", "page", "batch", "report",
}

# Part of a name, so they must survive the stop test above.
NAME_SUFFIXES = {"jr", "jr.", "sr", "sr.", "ii", "iii", "iv", "v"}

# Words that mean "the address starts here". They stop the name even though
# they carry no digit, which is what keeps "PO Box 41" out of a surname.
ADDRESS_START_WORDS = {
    "po", "p.o", "p.o.", "box", "apt", "apartment", "suite", "ste", "unit",
    "floor", "fl", "bldg", "building", "rm", "room", "#",
}

STREET_WORDS = {
    "st", "st.", "street", "ave", "ave.", "avenue", "rd", "rd.", "road",
    "dr", "dr.", "drive", "ln", "ln.", "lane", "blvd", "blvd.", "boulevard",
    "ct", "ct.", "court", "cir", "cir.", "circle", "way", "pl", "pl.",
    "place", "ter", "ter.", "terrace", "hwy", "highway", "pkwy", "parkway",
    "trl", "trail", "loop", "pike", "row", "run", "crossing", "xing",
    "square", "sq", "alley", "walk", "path", "plaza", "commons",
}

US_STATES = {
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID",
    "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS",
    "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK",
    "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV",
    "WI", "WY", "DC", "PR", "VI", "GU", "AS", "MP", "AA", "AE", "AP",
}

ZIP_RE = re.compile(r"\b\d{5}(?:-\d{4})?\b")
CITY_STATE_ZIP_RE = re.compile(r",\s*([A-Za-z]{2})\.?\s+\d{5}")
HOUSE_NUMBER_RE = re.compile(r"^\d+[A-Za-z]?\b")

# A course row: a course code then a term, or the report's own "Section
# skipped" text. These sit directly under a student line, which is exactly
# where an address would be, so they have to be recognised and refused.
COURSE_CODE_RE = re.compile(r"^[A-Za-z]{2,6}[- ]?\d{2,4}[A-Za-z]?$")
TERM_RE = re.compile(r"^\d{2}(?:FA|SP|SU|WI|F|S|W)\d?$", re.IGNORECASE)

MAX_ADDRESS_LINES = 3


# ===========================================================================
# line reconstruction  --  the step that decides whether any of this works
# ===========================================================================
Word = namedtuple("Word", "x0 x1 top bottom text")
Line = namedtuple("Line", "top words text")


def cluster_lines(words):
    """Group positioned words into printed lines, top to bottom.

    A word joins the line whose anchor it is nearest to VERTICALLY, compared
    against that anchor rather than against the previous word, so a run of
    slowly drifting words cannot chain two printed lines into one.

    The tolerance adapts to the font -- 40% of the median glyph height,
    floored at 1.5pt. A fixed tolerance is what breaks when the same report
    is printed at a different point size, and a 58-file set is exactly where
    that variation shows up.
    """
    if not words:
        return []

    heights = sorted(w.bottom - w.top for w in words)
    median = heights[len(heights) // 2] or 10.0
    tolerance = max(median * 0.4, 1.5)

    ordered = sorted(words, key=lambda w: ((w.top + w.bottom) / 2.0, w.x0))

    groups, current, anchor = [], [], None
    for word in ordered:
        centre = (word.top + word.bottom) / 2.0
        if anchor is None or centre - anchor <= tolerance:
            if anchor is None:
                anchor = centre
            current.append(word)
        else:
            groups.append(current)
            anchor, current = centre, [word]
    if current:
        groups.append(current)

    lines = []
    for group in groups:
        group.sort(key=lambda w: w.x0)
        lines.append(Line(top=min(w.top for w in group),
                          words=group,
                          text=" ".join(w.text for w in group)))
    return lines


def words_pdfplumber(page):
    return [Word(w["x0"], w["x1"], w["top"], w["bottom"], fold(w["text"]))
            for w in page.extract_words(use_text_flow=False,
                                        keep_blank_chars=False)]


def words_pymupdf(page):
    return [Word(w[0], w[2], w[1], w[3], fold(w[4]))
            for w in page.get_text("words")]


def pages_pdfplumber(path: Path):
    """Yield (page_number, positioned_lines, flat_text_lines) per page."""
    with pdfplumber.open(str(path)) as pdf:
        for number, page in enumerate(pdf.pages, start=1):
            flat = fold(page.extract_text() or "").splitlines()
            yield number, cluster_lines(words_pdfplumber(page)), flat


def pages_pymupdf(path: Path):
    doc = fitz.open(str(path))
    try:
        if doc.needs_pass:
            raise RuntimeError("PDF is encrypted and needs a password")
        for number, page in enumerate(doc, start=1):
            flat = fold(page.get_text("text") or "").splitlines()
            yield number, cluster_lines(words_pymupdf(page)), flat
    finally:
        doc.close()


# One place that decides what "plumber" and "mupdf" mean, so --engine, the
# fallback and --qa cannot drift apart.
ENGINES = {
    "plumber": (pages_pdfplumber, lambda: pdfplumber is not None),
    "mupdf": (pages_pymupdf, lambda: fitz is not None),
}

# pdfplumber is the default because it is the engine that has been shown to
# read this report. PyMuPDF is faster, and speed here comes from processing
# files in parallel -- which cannot change WHAT is read -- rather than from
# swapping the engine, which can.
DEFAULT_ENGINE = "plumber"
ENGINE_ORDER = ["plumber", "mupdf"]


def available_engines():
    return [name for name in ENGINE_ORDER if ENGINES[name][1]()]


# ===========================================================================
# column geometry  --  the "format" half
# ===========================================================================
Columns = namedtuple("Columns", "bands header_text")

WANTED_CAPTIONS = ("id", "ssn", "name", "address")


def find_columns(lines):
    """Measure the report's column bands from its heading row.

    Returns Columns, or None if this page has no heading row.

    The heading row must carry ID, SSN and Name together. Requiring all
    three is what stops the course heading -- "Course Name Term/Dt Grd
    Cum ..." -- being mistaken for it, since that line also contains the
    word "Name".

    A band runs from just left of its own caption to just left of the next
    caption on the heading row, whatever that next caption is. The first
    band starts at 0 rather than at its caption, because numeric columns in
    this report are often right-aligned and the value then begins to the
    LEFT of the word "ID".
    """
    for line in lines:
        seen = {}
        for word in line.words:
            key = word.text.strip(":.").lower()
            if key in WANTED_CAPTIONS and key not in seen:
                seen[key] = word
        if not {"id", "ssn", "name"} <= set(seen):
            continue
        if not seen["id"].x0 < seen["ssn"].x0 < seen["name"].x0:
            continue

        anchors = sorted(line.words, key=lambda w: w.x0)
        bands, pad = {}, 2.0
        for key, word in seen.items():
            following = [w.x0 for w in anchors if w.x0 > word.x0 + 0.5]
            left = 0.0 if word is anchors[0] else word.x0 - pad
            right = (min(following) - pad) if following else float("inf")
            bands[key] = (left, right)
        return Columns(bands=bands, header_text=line.text)
    return None


def in_band(word, band):
    left, right = band
    return left <= word.x0 < right


def near_band(word, band, unit, slack=3):
    """in_band, but tolerant of the gap between a caption and its values.

    This report does not align its captions to its data. The word "SSN" sits
    a character to the right of where the SSNs themselves begin, and the IDs
    are wider than the word "ID" and start well left of it. Tested strictly,
    a perfectly-read SSN comes back "outside its column" on nearly every row,
    and a note that fires on every row tells nobody anything.

    So membership is tested with a few characters of slack in each direction.
    That is wide enough to absorb the caption drift and still far narrower
    than the gap to the next column, so a band that is genuinely wrong -- the
    case worth reporting -- still fails the test.
    """
    left, right = band
    return left - slack * unit <= word.x0 < right + slack * unit


def char_width(line):
    """The line's typical character width, in points.

    Used as the unit for "are these two words part of the same value, or
    two different columns". Measured per line rather than assumed, because
    the same report gets printed at different point sizes.
    """
    widths = sorted((w.x1 - w.x0) / len(w.text) for w in line.words if w.text)
    return widths[len(widths) // 2] if widths else 5.0


def extend_past_band(words, last, unit):
    """Words that belong to a value but were printed past its column band.

    A value wider than its own heading overflows the band -- the ZIP at the
    end of an address is the usual casualty. Column bands cannot catch it,
    but the printed GAP can: inside one value words are a space apart, while
    the next column is separated by a run of spaces. Anything more than 2.5
    character widths away is a different column, so it is left alone.
    """
    tail = []
    for word in words:
        if word.x0 < last.x1:
            continue
        if word.x0 - last.x1 > 2.5 * unit:
            break
        if word.text.strip(":.").lower() in NAME_STOP_WORDS:
            break
        tail.append(word)
        last = word
    return tail


# ===========================================================================
# field cutting  --  shared by both halves
# ===========================================================================
def trim_name(tokens):
    """Cut a token list at the first token that is not part of a name."""
    for i, token in enumerate(tokens):
        bare = token.strip(":.,#()").lower()
        if bare in NAME_SUFFIXES:
            continue
        if (bare in NAME_STOP_WORDS
                or bare in ADDRESS_START_WORDS
                or bare in STREET_WORDS
                or re.search(r"\d", token)
                or token.endswith(":")):
            return tokens[:i]
    return tokens


def clean_name(tokens, notes):
    """A printed name from a token list, or "" if there is no name here."""
    name = " ".join(trim_name(tokens)).strip(" ,;:-")
    if not name or not re.search(r"[A-Za-z]", name):
        return ""
    if PREFIX_RE.match(name) and len(name.split()) == 1:
        # An honorific with nothing after it is not a name.
        return ""
    return name


def looks_like_address(text):
    """True when a line is address-shaped.

    Deliberately strict. A blank Address cell is visible and can be chased;
    a wrong one silently becomes someone else's address in the system of
    record, so anything that is not clearly an address is refused.
    """
    text = text.strip()
    if len(text) < 4 or not re.search(r"[A-Za-z]", text):
        return False
    if looks_like_course_row(text) or is_report_furniture(text):
        return False
    if SSN_RE.search(text):                     # that is the next student
        return False

    tokens = text.split()
    lowered = {t.strip(".,").lower() for t in tokens}

    if CITY_STATE_ZIP_RE.search(text):
        return True
    if ZIP_RE.search(text) and len(tokens) >= 2:
        return True
    if lowered & STREET_WORDS and HOUSE_NUMBER_RE.match(text):
        return True
    if lowered & STREET_WORDS and len(tokens) >= 2:
        return True
    if tokens[0].strip(".").lower() in {"po", "p.o"} or lowered & {"box"}:
        return bool(re.search(r"\d", text))
    # "Springfield, IL" with no ZIP -- a real second address line.
    match = re.search(r",\s*([A-Za-z]{2})\.?$", text)
    if match and match.group(1).upper() in US_STATES:
        return True
    return False


def looks_like_course_row(text):
    tokens = text.split()
    if not tokens:
        return False
    if "section skipped" in text.lower():
        return True
    if COURSE_CODE_RE.match(tokens[0]) and not HOUSE_NUMBER_RE.match(tokens[0]):
        return True
    return any(TERM_RE.match(t) for t in tokens[:3])


def is_report_furniture(text):
    lowered = text.lower()
    if set(text.strip()) <= set("- "):
        return True
    for caption in ("academic program", "report options", "batch id",
                    "sap type", "excluded remedial", "satisfactory academic",
                    "detail of results", "page "):
        if caption in lowered:
            return True
    tokens = [t.strip(":.,").lower() for t in text.split()]
    # A heading row is nothing but captions.
    return bool(tokens) and all(t in NAME_STOP_WORDS or not t for t in tokens)


def is_wrap_fragment(text):
    """A wrapped tail of the line above -- "00)", "---", "60.75)".

    The student line is wider than the page, so its last few characters are
    printed on their own underneath it. That fragment sits exactly where the
    address is, and it is what has to be stepped over to reach it.
    """
    stripped = text.strip()
    return bool(stripped) and len(stripped) <= 8 and not re.search(r"[A-Za-z]{2}", stripped)


def gather_address(following_texts):
    """Join the address printed under a student line, if it is there.

    Two bounds keep this honest. Only address-shaped lines are collected, and
    once collection has started the first line that is not an address ends
    it -- so a run of course rows can never be swept up as an address.

    Before the first address line, up to 2 lines may be stepped over: the
    student line's own wrapped tail and the report's captions both sit
    between the student and the address. Stepping over them is bounded and
    only ever reaches lines that pass looks_like_address anyway, and a line
    carrying an SSN is refused there, so this cannot walk into the NEXT
    student's address.
    """
    parts, skipped = [], 0
    for text in following_texts[:MAX_ADDRESS_LINES + 3]:
        text = text.strip()
        if not text:
            continue
        if looks_like_address(text):
            parts.append(text)
            if len(parts) >= MAX_ADDRESS_LINES:
                break
            continue
        if parts:
            break
        if skipped < 2 and (is_wrap_fragment(text) or is_report_furniture(text)):
            skipped += 1
            continue
        break
    return ", ".join(parts)


def split_name_and_address(tokens):
    """Split the tokens after the SSN into (name_tokens, address_text).

    The name is cut first; whatever is left over on the same line is the
    address only if it actually looks like one.
    """
    name_tokens = trim_name(tokens)
    tail = " ".join(tokens[len(name_tokens):]).strip(" ,;:-")
    if tail and looks_like_address(tail):
        return name_tokens, tail
    # An address may still be in there behind a caption -- "Uppal Academic
    # Program: ... 41 Maple St" is not a shape this report prints, so a tail
    # that fails the test is furniture and is dropped rather than guessed at.
    return name_tokens, ""


def make_row(student_id, ssn, name, address, address_source,
             method, source_line, notes):
    """Assemble one output row, or None if there is no student here."""
    if not name:
        return None

    row = dict.fromkeys(OUTPUT_COLUMNS, "")
    row["Student ID"] = student_id
    row["SSN"] = ssn
    row["Name"] = name
    row["Address"] = address
    row["Address Source"] = address_source if address else ""
    row["Method"] = method
    row["Source Line"] = source_line

    if not student_id:
        notes.append("no ID printed before the SSN")
    if not ssn:
        notes.append("no SSN-shaped value between the ID and the name")
    if not address:
        notes.append("no address found on or under the student line")
    row["Extraction Notes"] = "; ".join(notes)
    return row


# ===========================================================================
# the format half  --  fields from column bands
# ===========================================================================
def format_row(line, columns, following_texts):
    """One student row from a positioned line, using the column bands."""
    bands = columns.bands
    notes = []
    unit = char_width(line)

    # The SSN is the anchor for the whole row: it is the one field whose shape
    # is unmistakable, so it is found by pattern and then CHECKED against the
    # measured column rather than being looked for inside it. Searching inside
    # the band instead would mean a page whose bands are slightly off loses
    # every student on it.
    candidates = [w for w in line.words if SSN_TOKEN_RE.match(w.text)]
    if not candidates:
        return None
    in_column = [w for w in candidates if near_band(w, bands["ssn"], unit)]
    if in_column:
        ssn_word = in_column[0]
    else:
        ssn_word = candidates[0]
        notes.append("the SSN is nowhere near the measured SSN column -- found "
                     "by pattern instead; check the bands with --columns")
    ssn = ssn_word.text

    if "-" not in ssn and ssn.isdigit():
        notes.append("SSN printed without separators -- copied exactly as printed")
    elif re.search(r"[X*#?]", ssn):
        notes.append("SSN is partly masked in the source PDF")

    # ID: the nearest value left of the SSN. Preferring the measured ID column
    # keeps a stray label out of the field; if nothing sits there, the word
    # physically closest to the left of the SSN is the ID.
    left_of_ssn = [w for w in line.words if w.x1 <= ssn_word.x0]
    id_candidates = [w for w in left_of_ssn if near_band(w, bands["id"], unit)]
    if not id_candidates:
        id_candidates = left_of_ssn
        if id_candidates:
            notes.append("the ID is not near the measured ID column -- took the "
                         "value printed nearest left of the SSN")
    student_id = ""
    if id_candidates:
        chosen = max(id_candidates, key=lambda w: w.x0)
        student_id = chosen.text
        if not ID_TOKEN_RE.match(student_id):
            notes.append("the value left of the SSN is not ID-shaped -- "
                         "copied exactly as printed; verify it")

    after = [w for w in line.words if w.x0 >= ssn_word.x1]

    address, address_source = "", ""
    if "address" in bands:
        # An explicit Address column is the best answer available: no
        # guessing about what is or is not address-shaped.
        in_column = [w for w in after if in_band(w, bands["address"])]
        if in_column:
            in_column += extend_past_band(after, in_column[-1], unit)
            address = " ".join(w.text for w in in_column).strip(" ,;:-")
            address_source = "column"
            consumed = {id(w) for w in in_column}
            after = [w for w in after if id(w) not in consumed]

    name_tokens, same_line = split_name_and_address([w.text for w in after])
    name = clean_name(name_tokens, notes)

    if not address and same_line:
        address, address_source = same_line, "same line"
    if not address:
        below = gather_address(following_texts)
        if below:
            address, address_source = below, "below"

    return make_row(student_id, ssn, name, address, address_source,
                    "format", line.text, notes)


def caption_row(text, following_texts, method):
    """Last resort: a line carrying the report's student caption but no SSN
    that any pattern recognised.

    Every student line prints "Academic Program:", which makes the caption a
    second and independent way to know a student is on the line. This matters
    more than it looks: a student whose SSN is printed in a shape no pattern
    anticipated is not merely parsed wrongly, they are ABSENT -- and absent is
    the one outcome the reconciliation cannot catch either, because it counts
    SSNs with the same patterns that just failed.

    Everything before the caption is the student's own data in printed order:
    ID, then SSN, then name. Every field gets a loud note, because a row that
    reached this function was assembled without the one field whose shape
    could confirm it.
    """
    caption = ACADEMIC_RE.search(text)
    if not caption:
        return None
    tokens = text[:caption.start()].split()
    if len(tokens) < 2:
        return None

    notes = ["no SSN-shaped value on this line -- the student was recovered "
             "from the 'Academic Program' caption; verify every field"]
    student_id, rest = tokens[0], tokens[1:]

    ssn = ""
    if re.search(r"\d", rest[0]):
        # A person's name never carries a digit, so whatever sits between the
        # ID and the name is the SSN printed some other way. The SSN may also
        # be printed as separate groups ("555 12 4599"), so following groups
        # are pulled in until nine characters have been seen.
        ssn, rest = rest[0], rest[1:]
        while (rest and re.fullmatch(r"[0-9X*#?-]+", rest[0])
               and sum(c.isalnum() for c in ssn) < 9):
            ssn, rest = f"{ssn} {rest[0]}", rest[1:]

    name_tokens, same_line = split_name_and_address(rest)
    name = clean_name(name_tokens, notes)
    if same_line:
        address, source = same_line, "same line"
    else:
        address, source = gather_address(following_texts), "below"

    return make_row(student_id, ssn, name, address, source,
                    f"{method} (caption)", text, notes)


def rows_by_format(lines, columns):
    """Every student row on one page, read through the column bands."""
    texts = [line.text for line in lines]
    rows = []
    for i, line in enumerate(lines):
        if (is_report_furniture(line.text)
                and not SSN_RE.search(line.text)
                and not ACADEMIC_RE.search(line.text)):
            continue
        row = (format_row(line, columns, texts[i + 1:])
               or caption_row(line.text, texts[i + 1:], "format"))
        if row:
            rows.append(row)
    return rows


# ===========================================================================
# the text half  --  fields from the flat page text
# ===========================================================================
def text_row(text, following_texts):
    """One student row from a flat line of text, by pattern alone.

    The SSN is the anchor because it is the only field on the line whose
    shape is unmistakable: the ID is whatever token sits to its left, and
    the name is what follows, up to the point where the report's own
    captions begin.
    """
    match = SSN_RE.search(text)
    if not match:
        return None

    notes = []
    ssn = match.group(0)
    if "-" not in ssn:
        notes.append("SSN printed without separators -- copied exactly as printed")
    elif re.search(r"[X*#?]", ssn):
        notes.append("SSN is partly masked in the source PDF")

    before = text[:match.start()].split()
    student_id = ""
    if before:
        student_id = before[-1]
        if len(before) > 1:
            notes.append("more than one value left of the SSN -- took the one "
                         "nearest the SSN as the ID")
        if not ID_TOKEN_RE.match(student_id):
            notes.append("the value left of the SSN is not ID-shaped -- "
                         "copied exactly as printed; verify it")

    name_tokens, same_line = split_name_and_address(text[match.end():].split())
    name = clean_name(name_tokens, notes)

    address, address_source = "", ""
    if same_line:
        address, address_source = same_line, "same line"
    else:
        below = gather_address(following_texts)
        if below:
            address, address_source = below, "below"

    return make_row(student_id, ssn, name, address, address_source,
                    "text", text, notes)


def rows_by_text(flat_lines):
    rows = []
    for i, text in enumerate(flat_lines):
        row = (text_row(text, flat_lines[i + 1:])
               or caption_row(text, flat_lines[i + 1:], "text"))
        if row:
            rows.append(row)
    return rows


# ===========================================================================
# putting the two halves together
# ===========================================================================
def fill_gaps_from_text(rows, flat_lines):
    """Fill fields the format half left empty, using the text half.

    This is the field-level fallback, and it is what "if it does not come
    out by format, take it by text" means row by row rather than page by
    page. A format row with a blank ID is repaired from the pattern parse
    of the same printed line instead of shipping blank.
    """
    if not rows or not flat_lines:
        return rows

    # Match on the SSN: it is unique per student and both halves read it the
    # same way, so it joins the two parses without depending on line numbers
    # agreeing between the geometry and the flat text.
    by_ssn = {}
    for text in flat_lines:
        parsed = text_row(text, [])
        if parsed and parsed["SSN"]:
            by_ssn.setdefault(parsed["SSN"], parsed)

    for row in rows:
        other = by_ssn.get(row["SSN"])
        if not other:
            continue
        repaired = []
        for field in ("Student ID", "Name", "Address"):
            if not row[field] and other[field]:
                row[field] = other[field]
                if field == "Address":
                    row["Address Source"] = other["Address Source"]
                repaired.append(field)
        if repaired:
            row["Method"] = "format+text"
            note = ("taken from the plain text because the column bands gave "
                    "nothing for: " + ", ".join(repaired))
            row["Extraction Notes"] = (
                f"{row['Extraction Notes']}; {note}" if row["Extraction Notes"] else note)
    return rows


def rows_for_page(lines, flat_lines):
    """(rows, method_used) for one page.

    format first, text as the fallback -- and "fallback" is decided by the
    result, not by trusting the first path. If the geometry produced fewer
    students than the flat text did, the flat text wins: fewer rows means
    students missing, which is the failure this job cannot ship.
    """
    columns = find_columns(lines)

    if columns is not None:
        rows = rows_by_format(lines, columns)
        if rows:
            text_rows = rows_by_text(flat_lines)
            if len(text_rows) > len(rows):
                for row in text_rows:
                    note = (f"read from the plain text, which found "
                            f"{len(text_rows)} students against {len(rows)} "
                            f"from the columns")
                    row["Extraction Notes"] = (
                        f"{row['Extraction Notes']}; {note}"
                        if row["Extraction Notes"] else note)
                return text_rows, "text"
            return fill_gaps_from_text(rows, flat_lines), "format"

    rows = rows_by_text(flat_lines)
    if not rows and lines:
        # No heading row AND no flat text worth reading. The positioned words
        # are all that is left, so run the pattern parse over the rebuilt
        # lines rather than giving up on the page.
        rows = rows_by_text([line.text for line in lines])
        for row in rows:
            row["Method"] = "text (rebuilt lines)"
    return rows, "text"


# ===========================================================================
# per-file driver
# ===========================================================================
def read_with(path: Path, engine: str):
    """(rows, pages, ssns_on_page, pages_with_columns) for one PDF.

    ssns_on_page is the completeness yardstick and the reason this returns
    four things. It counts SSN-shaped tokens on the page BEFORE any parsing,
    so it is how many students are there regardless of whether the parser
    read them. Rows short of that count means students were lost -- a silent
    failure that otherwise has to be caught by eye against the PDF.
    """
    iterator, _ = ENGINES[engine]
    rows, pages, ssns, with_columns = [], 0, 0, 0

    for page_no, lines, flat in iterator(path):
        pages += 1
        if find_columns(lines) is not None:
            with_columns += 1

        found, _ = rows_for_page(lines, flat)
        for row in found:
            row["File Name"] = path.name
            row["Page Number"] = page_no
        rows.extend(found)

        # Count on the words, not on the flat text: if line clustering is
        # what broke, a count taken from clustered lines would be wrong in
        # the same way as the parse and the two would agree on a lie.
        joined = " ".join(w.text for line in lines for w in line.words)
        ssns += len(SSN_RE.findall(joined)) or len(SSN_RE.findall("\n".join(flat)))

    return rows, pages, ssns, with_columns


def process_pdf(path, engine: str = "auto"):
    """(rows, diagnostics) for one PDF.

    The retry rule is completeness, not emptiness: a file that produced SOME
    students but not all is the dangerous case, because it looks healthy. A
    retry only happens when it could find MORE, and the better result wins.
    """
    path = Path(path)

    if engine == "auto":
        order = available_engines()
        order.sort(key=lambda e: e != DEFAULT_ENGINE)
    else:
        order = [engine] if ENGINES[engine][1]() else []
    if not order:
        raise RuntimeError("no PDF engine available -- pip install pdfplumber pymupdf")

    primary = order[0]
    try:
        rows, pages, ssns, with_columns = read_with(path, primary)
        error = ""
    except Exception as exc:                                    # noqa: BLE001
        rows, pages, ssns, with_columns = [], 0, 0, 0
        error = f"{type(exc).__name__}: {exc}"

    engine_used, fallback_used = primary, ""

    incomplete = len(rows) < ssns or (not rows and not error)
    if engine == "auto" and incomplete and len(order) > 1:
        alt = order[1]
        try:
            alt_rows, alt_pages, alt_ssns, alt_cols = read_with(path, alt)
        except Exception as exc:                                # noqa: BLE001
            alt_rows, alt_pages, alt_ssns, alt_cols = [], 0, 0, 0
            error = (f"{error}; {type(exc).__name__}: {exc}" if error
                     else f"{type(exc).__name__}: {exc}")

        if len(alt_rows) > len(rows):
            note = (f"read with the {alt} engine, which found {len(alt_rows)} "
                    f"students against {len(rows)} from {primary}")
            for row in alt_rows:
                row["Extraction Notes"] = (
                    f"{row['Extraction Notes']}; {note}"
                    if row["Extraction Notes"] else note)
            rows, pages, with_columns = alt_rows, alt_pages, alt_cols
            ssns = max(ssns, alt_ssns)
            engine_used, fallback_used = alt, alt

    missing = max(0, ssns - len(rows))
    by_format = sum(1 for r in rows if r["Method"].startswith("format"))
    diagnostics = {
        "File Name": path.name,
        "Pages": pages,
        "SSNs On Page": ssns,
        "Rows Found": len(rows),
        "Missing": missing,
        "Complete": "" if missing else "yes",
        "By Format": by_format,
        "By Text": len(rows) - by_format,
        "With Address": sum(1 for r in rows if r["Address"]),
        "Pages With Columns": with_columns,
        "Pages Without Columns": max(0, pages - with_columns),
        "Engine": engine_used,
        "Fallback Used": fallback_used,
        "Error": error,
    }
    return rows, diagnostics


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


def find_pdfs(folder: Path):
    return sorted({p.resolve(): p for p in
                   list(folder.glob("*.pdf")) + list(folder.glob("*.PDF"))}.values(),
                  key=lambda p: p.name)


def process_folder(pdfs, engine="auto", workers=None, progress=None):
    """Every PDF, in parallel, one process per file.

    Parallelism is per file rather than per page: PDF readers hold
    per-document state, and a file that crashes an engine then takes down
    only its own worker.
    """
    pdfs = [Path(p) for p in pdfs]
    if workers is None:
        workers = max(1, (os.cpu_count() or 2) - 1)
    workers = max(1, min(workers, len(pdfs)))

    # Windows has no fork, so each worker is a fresh interpreter costing
    # about a second to start. On a handful of files the pool costs more than
    # the work it saves; that reverses on a real 5,000-page batch.
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

    all_rows.sort(key=lambda r: (r["File Name"], r["Page Number"]))
    diagnostics.sort(key=lambda d: d["File Name"])
    return all_rows, diagnostics


# ===========================================================================
# masked diagnostics
# ===========================================================================
# Captions belong to the report, not to a student, so they stay readable.
# Everything else collapses to its shape. The list is deliberately generous:
# a caption wrongly masked costs a little readability, a value wrongly left
# in the clear is a PII incident.
SAFE_WORDS = {
    "report", "options", "use", "all", "sections", "include", "contained",
    "w/in", "the", "range", "batch", "id", "ssn", "name", "address", "incl",
    "gpa", "cmpl", "att", "course", "term", "dt", "grd", "cum", "eval",
    "credits", "grade", "pts", "pgm", "earn", "academic", "program", "min",
    "max", "cred", "page", "of", "total", "excluded", "remedial", "sap",
    "type", "student", "satisfactory", "progress", "audit", "detail",
    "results", "by", "section", "skipped", "no", "yes", "exists", "verified",
}


def mask_word(word: str) -> str:
    if word.strip("():,.;#%*/-").lower() in SAFE_WORDS:
        return word
    return "".join("#" if c.isdigit() else
                   ("X" if c.isupper() else "x") if c.isalpha() else c
                   for c in word)


def mask(text: str) -> str:
    return " ".join(mask_word(w) for w in text.split())


def show_split(row, indent="    ") -> None:
    """Print how one line became fields, masked.

    This is the diagnostic that matters when the ROWS are right but the
    COLUMNS are wrong: "ID=####### SSN=###-##-#### NAME=Xxxx Xxxxx" against
    the source line shows whether a value landed in the wrong field without
    revealing a single real digit or letter.
    """
    print(f"{indent}METHOD  {row['Method']}")
    print(f"{indent}ID      [{mask(row['Student ID'])}]")
    print(f"{indent}SSN     [{mask(row['SSN'])}]")
    print(f"{indent}NAME    [{mask(row['Name'])}]")
    print(f"{indent}ADDRESS [{mask(row['Address'])}]"
          + (f"   (from {row['Address Source']})" if row["Address"] else ""))
    if row["Extraction Notes"]:
        print(f"{indent}NOTES   {row['Extraction Notes']}")


def show_columns(path: Path, page_no: int, engine: str = DEFAULT_ENGINE) -> None:
    """Print the column bands measured on one page. No PII."""
    if not ENGINES[engine][1]():
        print(f"engine {engine} is not installed")
        return

    iterator, _ = ENGINES[engine]
    for number, lines, _flat in iterator(path):
        if number != page_no:
            continue
        columns = find_columns(lines)
        print(f"{path.name}  page {page_no}  engine={engine}")
        if columns is None:
            print("\nNO heading row found on this page.")
            print("The format half cannot run here, so every row on this page")
            print("comes from the plain text. If that is true of every page,")
            print("the headings are spelled differently in these files --")
            print("compare the masked dump below against WANTED_CAPTIONS.")
            print("\nfirst 12 lines, masked:")
            for line in lines[:12]:
                print(f"    {mask(line.text)[:140]}")
            return
        print(f"\nheading row: {mask(columns.header_text)[:140]}\n")
        print(f"{'column':<10}{'from x':>10}{'to x':>10}")
        print("-" * 30)
        for key in WANTED_CAPTIONS:
            if key in columns.bands:
                left, right = columns.bands[key]
                edge = "end" if right == float("inf") else f"{right:.1f}"
                print(f"{key:<10}{left:>10.1f}{edge:>10}")
        if "address" not in columns.bands:
            print("\nNo Address column in the heading row, so the address is")
            print("looked for after the name and on the lines below it.")
        return
    print(f"page {page_no} not found in {path.name}")


def debug_page(path: Path, page_no: int, engine: str = DEFAULT_ENGINE) -> None:
    """Masked dump of one page: bands, every line, and how each was split."""
    if not ENGINES[engine][1]():
        print(f"engine {engine} is not installed")
        return

    iterator, _ = ENGINES[engine]
    total = 0
    for number, lines, flat in iterator(path):
        total += 1
        if number != page_no:
            continue

        show_columns(path, page_no, engine)
        rows, method = rows_for_page(lines, flat)
        print("\n" + "=" * 78)
        print(f"page read by: {method}    {len(rows)} student(s)")
        print("=" * 78)

        found = {r["Source Line"] for r in rows}
        for i, line in enumerate(lines, 1):
            flag = "STUDENT" if line.text in found else "       "
            print(f"{i:4} {flag} | {mask(line.text)[:140]}")
        print("-" * 78)
        for i, row in enumerate(rows, 1):
            print(f"\n[{i}]")
            show_split(row)
        return

    print(f"page {page_no} out of range (1..{total})")


def qa_file(path: Path, engine: str = DEFAULT_ENGINE, limit: int = 25) -> None:
    """Masked field-by-field dump of every student in one file.

    Run this when the workbook holds the wrong values: source line and
    parsed fields side by side, masked, so a bad split can be pasted into a
    ticket without exposing PII.
    """
    rows, diag = process_pdf(path, engine=engine)
    print(f"{path.name}   pages={diag['Pages']}  students={diag['Rows Found']}  "
          f"engine={diag['Engine']}"
          + (f"  ERROR={diag['Error']}" if diag["Error"] else ""))
    print(f"by format={diag['By Format']}  by text={diag['By Text']}  "
          f"with address={diag['With Address']}  "
          f"pages with columns={diag['Pages With Columns']} "
          f"of {diag['Pages']}")
    if diag["Missing"]:
        print(f"*** INCOMPLETE -- {diag['Missing']} student(s) on the pages "
              f"did not make it into rows ***")
    print("=" * 78)

    for i, row in enumerate(rows[:limit], 1):
        print(f"\n[{i}] page {row['Page Number']}")
        print(f"    LINE    | {mask(row['Source Line'])[:140]}")
        show_split(row)
    if len(rows) > limit:
        print(f"\n... {len(rows) - limit} more not shown (raise --limit)")

    flagged = [r for r in rows if r["Extraction Notes"]]
    no_address = [r for r in rows if not r["Address"]]
    print("\n" + "=" * 78)
    print(f"{len(rows)} student(s), {len(flagged)} carrying a note, "
          f"{len(no_address)} with no address")


def selftest() -> None:
    """Parse the synthetic test PDFs and print results. No real PII."""
    tests = Path(__file__).parent / "tests"
    pdfs = sorted(tests.glob("sample_sap*.pdf"))
    if not pdfs:
        print(f"no test PDFs in {tests}")
        return

    total = 0
    for path in pdfs:
        rows, diag = process_pdf(path)
        total += len(rows)
        flag = "" if not diag["Missing"] else f"  MISSING {diag['Missing']}"
        print(f"\n{path.name}")
        print(f"  pages={diag['Pages']}  ssns={diag['SSNs On Page']}  "
              f"rows={diag['Rows Found']}  format={diag['By Format']}  "
              f"text={diag['By Text']}  address={diag['With Address']}  "
              f"engine={diag['Engine']}{flag}")
        for row in rows:
            print(f"    ID={row['Student ID']:<9} SSN={row['SSN']:<13} "
                  f"NAME={row['Name']:<28} ADDR={row['Address']}"
                  + (f"  [{row['Address Source']}]" if row["Address"] else ""))
    print(f"\nTOTAL {total} student row(s) from {len(pdfs)} file(s)")


# ===========================================================================
# workbook
# ===========================================================================
def write_workbook(rows, diagnostics, dest: Path) -> Path:
    out = dest / OUTPUT_XLSX_NAME
    frame = (pd.DataFrame(rows, columns=OUTPUT_COLUMNS) if rows
             else pd.DataFrame(columns=OUTPUT_COLUMNS))
    diag_frame = pd.DataFrame(diagnostics, columns=DIAGNOSTIC_COLUMNS)
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name="Students", index=False)
        diag_frame.to_excel(writer, sheet_name="Reconciliation", index=False)
    return out


def summarise(rows, diagnostics):
    """The numbers that decide whether a run can be signed off."""
    return {
        "rows": len(rows),
        "pages": sum(d["Pages"] for d in diagnostics),
        "ssns": sum(d["SSNs On Page"] for d in diagnostics),
        "by_format": sum(d["By Format"] for d in diagnostics),
        "by_text": sum(d["By Text"] for d in diagnostics),
        "with_address": sum(d["With Address"] for d in diagnostics),
        "no_columns": sum(d["Pages Without Columns"] for d in diagnostics),
        "incomplete": [d for d in diagnostics if d["Missing"]],
        "zero": [d["File Name"] for d in diagnostics if d["Rows Found"] == 0],
        "fallback": [d["File Name"] for d in diagnostics if d["Fallback Used"]],
        "errors": [d for d in diagnostics if d["Error"]],
    }


def run_headless(src: Path, dst: Path, engine="auto", workers=None) -> None:
    """A whole folder, no GUI. Prints counts only, never a student value."""
    pdfs = find_pdfs(src)
    if not pdfs:
        print(f"no PDFs in {src}")
        return

    start = time.time()
    rows, diagnostics = process_folder(
        pdfs, engine=engine, workers=workers,
        progress=lambda done, total, name, n:
            print(f"  [{done}/{total}] {name}: {n} student(s)"))
    out = write_workbook(rows, diagnostics, dst)

    s = summarise(rows, diagnostics)
    elapsed = time.time() - start
    print(f"\n{s['rows']} students from {len(pdfs)} files ({s['pages']} pages) "
          f"in {elapsed:.1f}s"
          + (f"  --  {s['pages'] / elapsed:.0f} pages/sec" if elapsed else ""))
    print(f"read by columns: {s['by_format']}    by plain text: {s['by_text']}")
    print(f"with an address: {s['with_address']} of {s['rows']}")
    if s["no_columns"]:
        print(f"{s['no_columns']} page(s) had no heading row -- run --columns "
              f"on one of those files to see why")

    # Completeness first: a row count on its own reads as success even when
    # students are missing.
    print(f"completeness: {s['rows']} rows against {s['ssns']} SSN(s) on the pages")
    if s["incomplete"]:
        print(f"\n*** {len(s['incomplete'])} file(s) INCOMPLETE -- students were lost ***")
        for d in s["incomplete"][:20]:
            print(f"    {d['File Name']}: {d['Rows Found']} of "
                  f"{d['SSNs On Page']} (missing {d['Missing']})")
        print("    Run --qa on one of these to see how its lines were split.")
    else:
        print("every file complete -- rows match the SSNs found on the pages")

    for label, names in (("produced NO rows", s["zero"]),
                         ("needed the fallback engine", s["fallback"])):
        if names:
            print(f"\n{len(names)} file(s) {label}:")
            for name in names[:20]:
                print(f"    {name}")
    if s["errors"]:
        print(f"\n{len(s['errors'])} file(s) errored:")
        for d in s["errors"][:20]:
            print(f"    {d['File Name']}: {d['Error']}")

    print(f"\nworkbook: {out}")
    print("Holds SSNs, names and addresses in clear text -- move it to the "
          "approved Global Insider folder and delete any local copy.")


# ===========================================================================
# GUI
# ===========================================================================
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("SAP Audit  --  ID / SSN / Name / Address Extractor")
        self.root.geometry("720x470")
        self.root.minsize(580, 400)
        self.src = tk.StringVar()
        self.dst = tk.StringVar()
        self.status = tk.StringVar(value="Choose a source folder and a destination folder.")
        self.eta = tk.StringVar(value="")
        self.engine = tk.StringVar(value="auto")
        self.workers = tk.IntVar(value=max(1, (os.cpu_count() or 2) - 1))
        # Disabling the button is not enough on its own: Return is bound to
        # the same action and would bypass it, starting a second run over the
        # same folder while the first is still writing.
        self.running = False
        self._build()

    def _build(self):
        pad = {"padx": 10, "pady": 4}

        # Packed FIRST against the bottom edge. With pack(), whatever claims
        # side="bottom" earliest is placed outermost, so the Extract button
        # keeps its space however small the window gets.
        bottom = tk.Frame(self.root)
        bottom.pack(side="bottom", fill="x", pady=(4, 10))
        self.go = tk.Button(bottom, text="Extract", command=self._start,
                            width=18, height=2, default="active")
        self.go.pack()
        self.root.bind("<Return>", lambda _e: self._start())

        tk.Label(self.root, text="Source folder (PDFs)",
                 anchor="w").pack(fill="x", **pad)
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
                 text="Columns are read first; the plain text is the fallback. "
                      "Each row records which one produced it."
                 ).pack(fill="x", padx=10)

        tk.Label(
            self.root,
            text=("The workbook holds SSNs, names and addresses in clear text.\n"
                  "Save it to the approved Global Insider folder only "
                  "-- never a desktop or a local temp path."),
            fg="#a00", justify="left", anchor="w",
        ).pack(fill="x", **pad)

        self.bar = ttk.Progressbar(self.root, orient="horizontal", mode="determinate")
        self.bar.pack(fill="x", **pad)
        tk.Label(self.root, textvariable=self.status, anchor="w",
                 wraplength=680, justify="left").pack(fill="x", **pad)
        tk.Label(self.root, textvariable=self.eta, anchor="w",
                 wraplength=680, justify="left").pack(fill="x", **pad)

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
            pdfs = find_pdfs(src)
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
                print(f"  {name}: {count} student(s)")
                self._set(f"Read {name}   ({done} of {total})",
                          f"Elapsed {elapsed / 60:.1f} min   "
                          f"Remaining ~{remaining / 60:.1f} min   "
                          f"{workers} file(s) at a time")

            self._set(f"Reading {len(pdfs)} PDFs on {workers} worker(s)...")
            rows, diagnostics = process_folder(pdfs, engine=self.engine.get(),
                                               workers=workers, progress=progress)

            self._set("Writing the workbook...")
            out = write_workbook(rows, diagnostics, dst)

            s = summarise(rows, diagnostics)
            mins = (time.time() - start) / 60

            msg = (f"{s['rows']} students from {len(pdfs)} PDFs "
                   f"({s['pages']} pages) in {mins:.1f} min.\n\n"
                   f"Saved to:\n{out}\n\n"
                   f"Read by columns: {s['by_format']}    "
                   f"by plain text: {s['by_text']}\n"
                   f"With an address: {s['with_address']} of {s['rows']}\n\n")

            # Lead with completeness. A row count on its own reads as success
            # even when students are missing.
            msg += (f"Completeness: {s['rows']} rows against {s['ssns']} SSNs "
                    f"on the pages.\n")
            if s["incomplete"]:
                msg += (f"\nWARNING -- {len(s['incomplete'])} file(s) are "
                        f"INCOMPLETE. Students were lost:\n  "
                        + "\n  ".join(f"{d['File Name']}: {d['Rows Found']} of "
                                      f"{d['SSNs On Page']}"
                                      for d in s["incomplete"][:8]))
                if len(s["incomplete"]) > 8:
                    msg += f"\n  ... and {len(s['incomplete']) - 8} more"
                msg += ("\n\nDo NOT sign this off. Run --qa on one of those "
                        "files to see how its lines were split.")
            else:
                msg += "Every student on every page was accounted for."

            if s["zero"]:
                msg += (f"\n\n{len(s['zero'])} file(s) produced NO rows:\n  "
                        + "\n  ".join(s["zero"][:8]))
            if s["with_address"] == 0 and s["rows"]:
                msg += ("\n\nNo addresses were found in any file. Either the "
                        "report does not print one, or it is printed somewhere "
                        "this does not look -- run --debug on one page.")

            msg += ("\n\nThis workbook holds SSNs, names and addresses in clear "
                    "text. Move it to the approved Global Insider folder and "
                    "delete any local copy.")

            self._set(f"Done -- {s['rows']} students from {len(pdfs)} PDFs"
                      + (f", {len(s['incomplete'])} file(s) INCOMPLETE"
                         if s["incomplete"] else ""))
            if s["incomplete"]:
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
def flag_value(args, flag, default=None, cast=str):
    return cast(args[args.index(flag) + 1]) if flag in args else default


def main() -> None:
    args = sys.argv[1:]

    if not available_engines():
        print("Install a PDF engine first:  pip install pdfplumber pymupdf")
        return

    command = args[0] if args else ""

    if command == "--selftest":
        selftest()
        return

    if command in ("--qa", "--debug", "--columns"):
        if len(args) < 2:
            print(f'usage: {command} "<file.pdf>" '
                  f'[page] [--engine plumber|mupdf] [--limit N]')
            return
        path = Path(args[1])
        if not path.is_file():
            print(f"not a file: {path}")
            return
        engine = flag_value(args, "--engine", DEFAULT_ENGINE)
        if engine not in ENGINES:
            print(f"unknown engine {engine!r} -- use plumber or mupdf")
            return
        page = int(args[2]) if len(args) > 2 and args[2].isdigit() else 1
        if command == "--qa":
            qa_file(path, engine, flag_value(args, "--limit", 25, int))
        elif command == "--debug":
            debug_page(path, page, engine)
        else:
            show_columns(path, page, engine)
        return

    if command == "--run":
        if len(args) < 3:
            print('usage: --run "<src folder>" "<dest folder>" '
                  '[--engine plumber|mupdf] [--workers N]')
            return
        src, dst = Path(args[1]), Path(args[2])
        if not src.is_dir():
            print(f"not a folder: {src}")
            return
        if not dst.is_dir():
            print(f"not a folder: {dst}")
            return
        run_headless(src, dst, flag_value(args, "--engine", "auto"),
                     flag_value(args, "--workers", None, int))
        return

    if command:
        print(f"unknown option {command!r}")
        print(__doc__.split("USAGE")[1].split("REQUIREMENTS")[0])
        return

    if pdfplumber is None:
        print("NOTE: pdfplumber is not installed, so PyMuPDF is doing the work "
              "and there is no cross-check engine:  pip install pdfplumber")
    if fitz is None:
        print("NOTE: pymupdf is not installed, so an incomplete file cannot be "
              "retried with a second engine:  pip install pymupdf")

    App().run()


if __name__ == "__main__":
    main()
