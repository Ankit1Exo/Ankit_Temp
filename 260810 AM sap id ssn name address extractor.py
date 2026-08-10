r"""
SAP Audit  --  Student ID / SSN / Name / Address extractor
==========================================================
Reads "Satisfactory Academic Progress Audit Report" PDFs and writes four
fields per student -- Student ID, SSN, Name, Address -- into one Excel
workbook, with a reconciliation sheet that says whether anything was lost.

TWO REPORT LAYOUTS, AND WHY THAT DECIDES THE DESIGN
    Two reports need reading and they do not agree on field order:

        SAP audit report        Student ID   SSN            Name
        NSC analytical report   Student ID   Name/Address   SSN   Enrollment...

    In the NSC report the name is printed to the LEFT of the SSN. What sits
    to its right is the enrollment status -- so anything that reads the name
    from a fixed side of the SSN fills the Name column with "Half time or
    more". That failure is worth describing because of how it looks: every
    cell is populated, nothing is blank, and it passes a glance. It only
    gives itself away as four distinct values repeated down every row.

    So the side is measured, never assumed. The ID is the leftmost ID-shaped
    token and the name is whatever is left over beside it; if nothing is
    left over, the name is on the other side of the SSN. That one rule reads
    both reports, and a third one it has not seen.

TWO INDEPENDENT READERS
    text     The page's plain text lines. No geometry at all -- the SSN
             anchors the line and the fields are cut from around it.

    format   Words with their x/y boxes. Printed rows are rebuilt from the
             y coordinates and fields are cut using the report's own column
             bands, measured by finding the ID / SSN / Name headings.

    THE TEXT READER RUNS FIRST, and then the data is arranged. It needs
    nothing from the page but its characters: no heading row to find, no
    bands to measure. So it runs on every file, including the ones where the
    geometry cannot get started, and putting it first means the worst case
    for any page is the reader that always works.

    The geometry runs second and earns its place where the flat text
    genuinely cannot separate the fields. Under a combined "Name/Address"
    heading the address has its own lines but not its own place in the text:
    flattened, an address line and whatever was printed across from it
    become one string. Only position tells them apart.

    Neither reader is trusted over the other on principle -- the result
    decides. Whichever found more students wins, the other fills its blanks,
    and an address measured from a column beats one inferred from shape.
    They are separate code paths over separate representations of the page,
    because a fallback that shares the step that failed is not a fallback.

    Every row records its reader in the Method column. With the text reader
    first, "text" is the normal answer and is not a warning; "format" means
    the geometry found students the text could not.

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
def is_ssn_token(text):
    """True when a whole token is an SSN.

    This is SSN_RE anchored, and it must stay that way. An earlier version
    allowed the separators to be optional -- [0-9]{3}-?[0-9]{2}-?[0-9]{4} --
    which also matches a ZIP+4: "19901-4412" is three digits, two digits, a
    dash and four digits. The reader then found an "SSN" on an address line
    and invented a student called "Dover, DE" out of it.

    A shape loose enough to find an SSN anywhere is loose enough to find one
    where there is none, and a fabricated student is worse than a missed
    field: nothing downstream can tell it is not real.
    """
    return bool(SSN_RE.fullmatch(text))

# An ID as printed: digits, possibly with a letter prefix or a dash.
ID_TOKEN_RE = re.compile(r"^[A-Za-z]{0,3}[-]?\d[\dA-Za-z-]*$")

# A caption that only ever appears on a student's own line. Each report has
# one, and it is a second way -- independent of the SSN -- to know a student
# is on the line. The optional leading single letter absorbs the status code
# the NSC report prints in front of the words ("H Half time or more"), so the
# code does not get taken for the tail of a name.
STUDENT_CAPTION_RE = re.compile(
    r"\bacademic\s*program\b"
    r"|(?:\b[A-Za-z]\s+)?(?:\b(?:half|full)\s+time\b"
    r"|\bless\s+than\s+half\b|\bwithdrawn\b|\bgraduated\b|\bdeceased\b"
    r"|\bleave\s+of\s+absence\b)",
    re.IGNORECASE,
)

# The trailing "(?:\s+|(?=[A-Z]))" is load bearing twice over:
#   "Mroz Dana K"  "Mr" matches, no dot follows, and the next character is a
#                  lowercase "o" -- neither space nor capital -- so the match
#                  fails and the surname survives. Without the guard this
#                  becomes prefix "Mr" plus name "oz Dana K".
#   "Mrs. Jane"    alternation is ordered so "Mr" is tried first, which would
#                  leave "s. Jane". The guard rejects it, the engine
#                  backtracks to "Mrs", and the dot and space then match.
PREFIX_RE = re.compile(r"^((?:Mr|Mrs|Ms|Miss|Dr|Prof)\.?)(?:\s+|(?=[A-Z]))")

# A single letter, with or without a full stop -- a middle initial inside a
# name, but never a name in its own right.
INITIAL_RE = re.compile(r"^[A-Za-z]\.?$")

# The name ends at the first token that is one of these captions, holds a
# digit, or ends in a colon. Without it the name runs on into the report
# furniture -- "Liam O'Brien SAP Type: DHDHS" is the shape of the mistake.
NAME_STOP_WORDS = {
    # the SAP audit report's captions
    "academic", "program", "sap", "type", "excluded", "remedial",
    "credits", "credit", "incl", "gpa", "status", "degree", "major",
    "cmpl", "att", "pgm", "earn", "eval", "cum", "grd", "term", "dt",
    "course", "section", "skipped", "total", "page", "batch", "report",
    # the NSC analytical report's captions and its enrollment statuses.
    # These earn their place: this column sits directly beside the name, and
    # without them the Name field fills up with "Half time or more" -- which
    # looks like a successful extraction because every cell has something in
    # it, and only gives itself away as four values repeated over 5,000 rows.
    "enrollment", "agd", "begin", "end", "half", "full", "time", "less",
    "than", "more", "withdrawn", "graduated", "deceased", "leave",
    "absence", "nsc", "analytical", "record", "found", "name/address",
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
Columns = namedtuple("Columns", "bands header_text combined")

WANTED_CAPTIONS = ("id", "ssn", "name", "address")


def find_columns(lines):
    """Measure the report's column bands from its heading row.

    Returns Columns, or None if this page has no heading row.

    The heading row must carry ID, SSN and Name together. Requiring all
    three is what stops the SAP report's course heading -- "Course Name
    Term/Dt Grd Cum ..." -- being mistaken for it, since that line also
    contains the word "Name".

    What it must NOT do is require them in a particular ORDER. Two reports
    that both need reading disagree about it:

        Student ID   SSN   Name              the SAP audit report
        Student ID   Name/Address   SSN      the NSC analytical report

    An earlier version demanded id < ssn < name and so found no heading row
    at all on the NSC report, which threw every page onto the looser reader.
    Order is measured here and used later; it is never assumed.

    "Name/Address" is one heading over two fields, printed on separate lines
    -- the name on the student's line, the address beneath it. Both fields
    get the same band and `combined` records that they share it, because the
    address then has to be read DOWN the column rather than across the row.

    A band runs from just left of its own caption to just left of the next
    caption, whatever that next caption is. The first band starts at 0
    rather than at its caption, because a numeric column is often
    right-aligned and its values then begin to the LEFT of the heading word.
    """
    for line in lines:
        seen, combined = {}, False
        for word in line.words:
            key = word.text.strip(":.").lower()
            if key in ("name/address", "name/addr", "student"):
                if key != "student":
                    seen.setdefault("name", word)
                    seen.setdefault("address", word)
                    combined = True
                continue
            if key in WANTED_CAPTIONS:
                seen.setdefault(key, word)
        if not {"id", "ssn", "name"} <= set(seen):
            continue

        anchors = sorted(line.words, key=lambda w: w.x0)
        bands, pad = {}, 2.0
        for key, word in seen.items():
            following = [w.x0 for w in anchors if w.x0 > word.x0 + 0.5]
            left = 0.0 if word is anchors[0] else word.x0 - pad
            right = (min(following) - pad) if following else float("inf")
            bands[key] = (left, right)
        return Columns(bands=bands, header_text=line.text, combined=combined)
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
    if INITIAL_RE.match(name):
        # A lone letter is never a person. It is the NSC report's enrollment
        # status code ("H Half time or more") left behind after the words are
        # trimmed off, and without this guard it becomes the Name -- so the
        # workbook fills with one-letter names that all look plausible
        # individually.
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
    if lowered & STREET_WORDS and len(tokens) >= 2 and re.search(r"\d", text):
        # The digit is required. A street word plus any other word is far too
        # little on its own: the column heading fragment "Ct SSN" satisfied it
        # and became somebody's address. Every real address carries a number.
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


def note_ssn_shape(ssn, notes):
    if not ssn:
        return
    if "-" not in ssn and ssn.isdigit():
        notes.append("SSN printed without separators -- copied exactly as printed")
    elif re.search(r"[X*#?]", ssn):
        notes.append("SSN is partly masked in the source PDF")


def cut_fields(before, after, notes):
    """(student_id, name, same_line_address) from the tokens around the SSN.

    `before` and `after` are the tokens printed left and right of the SSN.

    The side the NAME sits on is worked out, not assumed, and that is the
    whole point of this function. The two reports disagree:

        1237906  555-11-1427  Mrs. Birtukhan Abate      name to the RIGHT
        1234506  Alvarez, Ana  555-01-2345  H Half...   name to the LEFT

    Reading a fixed side is what put "Half time or more" in the Name column
    of the NSC report: to the right of that report's SSN is the enrollment
    status, not a person.

    The rule that covers both: the ID is the leftmost ID-shaped token, and
    the name is whatever is LEFT OVER on that side once the ID is taken. If
    nothing is left over -- the SAP shape, where the SSN follows the ID
    immediately -- the name is on the right instead.

    It is self-correcting rather than clever. Anything left over that is
    report furniture is thrown away by trim_name, which empties the left
    side and sends the search to the right anyway.
    """
    student_id, rest = "", list(before)
    for i, token in enumerate(before):
        if ID_TOKEN_RE.match(token) and re.search(r"\d", token):
            student_id, rest = token, list(before[i + 1:])
            break

    left_name = trim_name(rest)
    if left_name:
        name_tokens = left_name
        # With the name on the left, everything right of the SSN is other
        # columns -- status, dates, credits. Only text that passes the strict
        # address test is taken from there, so a status never lands in Address.
        leftover = " ".join(rest[len(left_name):]).strip(" ,;:-")
        tail = " ".join(after).strip(" ,;:-")
        same_line = next((t for t in (leftover, tail)
                          if t and looks_like_address(t)), "")
    else:
        name_tokens, same_line = split_name_and_address(after)

    if before and not student_id:
        notes.append("nothing ID-shaped was printed left of the SSN")
    return student_id, clean_name(name_tokens, notes), same_line


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
def gather_address_in_band(following_lines, band, unit):
    """The address printed DOWN a column, beneath the student's line.

    This is what the geometry buys that the flat text cannot. Under a
    combined "Name/Address" heading the address occupies its own lines but
    NOT its own place in the text: flattened, an address line and whatever
    else was printed across from it become one string. Reading only the
    words that sit inside the column keeps them apart.
    """
    parts, skipped = [], 0
    for line in following_lines[:MAX_ADDRESS_LINES + 3]:
        if SSN_RE.search(line.text):            # the next student
            break
        words = [w for w in line.words if near_band(w, band, unit, slack=1)]
        text = " ".join(w.text for w in words).strip()
        if text and looks_like_address(text):
            parts.append(text)
            if len(parts) >= MAX_ADDRESS_LINES:
                break
            continue
        if parts:
            break
        if skipped < 2 and (not text or is_wrap_fragment(text)
                            or is_report_furniture(line.text)):
            skipped += 1
            continue
        break
    return ", ".join(parts)


def format_row(line, columns, following_lines):
    """One student row from a positioned line, using the column bands."""
    bands = columns.bands
    notes = []
    unit = char_width(line)

    # The SSN is the anchor for the whole row: it is the one field whose shape
    # is unmistakable, so it is found by pattern and then CHECKED against the
    # measured column rather than being looked for inside it. Searching inside
    # the band instead would mean a page whose bands are slightly off loses
    # every student on it.
    candidates = [w for w in line.words if is_ssn_token(w.text)]
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
    note_ssn_shape(ssn, notes)

    before = [w for w in line.words if w.x1 <= ssn_word.x0]
    after = [w for w in line.words if w.x0 >= ssn_word.x1]

    address, address_source = "", ""
    if "address" in bands and not columns.combined:
        # An Address column of its own is the best answer available: nothing
        # has to be guessed about what is or is not address-shaped.
        in_column = [w for w in after if in_band(w, bands["address"])]
        if in_column:
            in_column += extend_past_band(after, in_column[-1], unit)
            address = " ".join(w.text for w in in_column).strip(" ,;:-")
            address_source = "column"
            consumed = {id(w) for w in in_column}
            after = [w for w in after if id(w) not in consumed]

    student_id, name, same_line = cut_fields(
        [w.text for w in before], [w.text for w in after], notes)

    if not address and same_line:
        address, address_source = same_line, "same line"
    if not address and "address" in bands:
        below = gather_address_in_band(following_lines, bands["address"], unit)
        if below:
            address, address_source = below, "below (column)"
    if not address:
        below = gather_address([l.text for l in following_lines])
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
    ID, then SSN if one was printed, then name. Every field gets a loud note,
    because a row that reached this function was assembled without the one
    field whose shape could confirm it.
    """
    caption = STUDENT_CAPTION_RE.search(text)
    if not caption:
        return None
    tokens = text[:caption.start()].split()
    if len(tokens) < 2:
        return None

    notes = ["no SSN-shaped value on this line -- the student was recovered "
             "from the report's own student caption; verify every field"]
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
    rows = []
    for i, line in enumerate(lines):
        if (is_report_furniture(line.text)
                and not SSN_RE.search(line.text)
                and not STUDENT_CAPTION_RE.search(line.text)):
            continue
        following = lines[i + 1:]
        row = (format_row(line, columns, following)
               or caption_row(line.text, [l.text for l in following], "format"))
        if row:
            rows.append(row)
    return rows


# ===========================================================================
# the text half  --  fields from the flat page text
# ===========================================================================
def text_row(text, following_texts):
    """One student row from a flat line of text, by pattern alone.

    The SSN is the anchor because it is the only field on the line whose
    shape is unmistakable. Which side of it the ID and the name sit on is
    left to cut_fields, so this reader handles both report layouts without
    being told which one it is looking at.
    """
    match = SSN_RE.search(text)
    if not match:
        return None

    notes = []
    ssn = match.group(0)
    note_ssn_shape(ssn, notes)

    student_id, name, same_line = cut_fields(
        text[:match.start()].split(), text[match.end():].split(), notes)

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
def row_quality(rows):
    """How good a set of rows is, for choosing between two readings.

    Row COUNT alone is the wrong measure, and choosing on it is how a page
    ends up with six students named "H". A reading that finds six students
    with an ID and a name beats one that finds six SSNs with neither, so
    completed students are counted first and the raw count only breaks ties.

    The last term is total name length, and it settles the case where two
    readings find the same students but one of them lost a word. Where an
    engine's own text splits a printed line, the leading token goes missing
    -- "Mrs. Jane D. Smith" comes back as "Jane D. Smith" -- and every count
    above still ties. More surviving name text is the only thing that
    separates them, and since trim_name has already thrown out the report's
    furniture, more text here means more of the person's actual name.
    """
    complete = sum(1 for r in rows if r["Student ID"] and r["Name"])
    return (complete,
            sum(1 for r in rows if r["Address"]),
            len(rows),
            sum(len(r["Name"]) for r in rows))


def merge_rows(primary, secondary, label):
    """Fill blanks in `primary` from `secondary`, matching on SSN then ID.

    This is the field-level fallback: whichever reader produced a row, a
    field it left empty is taken from the other one rather than shipping
    blank. Matching on the SSN joins the two without depending on line
    numbers agreeing between the geometry and the flat text; the ID is the
    fallback key, for the rows where no SSN was printed at all.

    Only blanks are filled. A field both readers answered keeps the primary
    reader's value, so the result stays explainable -- one reader per row,
    named in the Method column, plus whatever it had to borrow.
    """
    other = "format" if label == "text" else "text"
    index = {}
    for row in secondary:
        for key in (row["SSN"], row["Student ID"]):
            if key:
                index.setdefault(key, row)

    for row in primary:
        donor = index.get(row["SSN"]) or index.get(row["Student ID"])
        if not donor:
            continue
        filled = []
        for field in ("Student ID", "SSN", "Name", "Address"):
            if not row[field] and donor[field]:
                row[field] = donor[field]
                if field == "Address":
                    row["Address Source"] = donor["Address Source"]
                filled.append(field)
        if filled:
            row["Method"] = f"{label}+{other}"
            note = (f"taken from the {other} reader, which the {label} reader "
                    f"left empty: " + ", ".join(filled))
            row["Extraction Notes"] = (
                f"{row['Extraction Notes']}; {note}" if row["Extraction Notes"]
                else note)
            # A borrowed field is no longer missing.
            row["Extraction Notes"] = row["Extraction Notes"].replace(
                "no address found on or under the student line; ", "")

        # One exception to "blanks only". When both readers found an address
        # and they disagree, the column reader wins: it knows WHERE the
        # address column is, while the text reader has to infer from what an
        # address looks like. On a combined "Name/Address" column the flat
        # text can hand back an address line fused with whatever was printed
        # across from it -- still address-shaped, still wrong.
        if (donor["Address"] and "column" in donor["Address Source"]
                and "column" not in row["Address Source"]
                and donor["Address"] != row["Address"]):
            row["Address"] = donor["Address"]
            row["Address Source"] = donor["Address Source"]
            note = ("address taken from the measured column rather than the "
                    "plain text, which read it differently")
            row["Extraction Notes"] = (
                f"{row['Extraction Notes']}; {note}" if row["Extraction Notes"]
                else note)
    return primary


def rows_for_page(lines, flat_lines):
    """(rows, reader_used) for one page.

    THE TEXT READER GOES FIRST, and then the data is arranged.

    The reason is that the text reader needs nothing from the page except its
    characters. It has no heading row to find and no column bands to measure,
    so it works on every file, including the ones where the geometry cannot
    get started. Putting it first means the worst case for any page is the
    reader that always runs, rather than the reader that sometimes cannot.

    The geometry runs second, and it earns its place on the pages where the
    flat text genuinely cannot separate the fields: under a combined
    "Name/Address" heading the address is told apart from the columns printed
    across from it only by WHERE IT SITS.

    Neither reader is trusted over the other on principle -- the result
    decides. Whichever found more students wins and the other fills its
    blanks, because fewer rows means students missing, and that is the one
    failure this job cannot ship.
    """
    # The text reader gets TWO text sources, because "the plain text of the
    # page" is not one thing. Asked for a page's text, pdfplumber returns one
    # string per printed line, while PyMuPDF returns it block by block -- and
    # on a columnar report that puts a student's ID, name and SSN on three
    # separate output lines. The reader then finds an SSN with nothing beside
    # it and reads the status code as the name.
    #
    # So both arrangements are read and the better one is kept: the engine's
    # own text, and lines rebuilt from the word coordinates. Whichever yields
    # more complete students wins, which makes the text reader independent of
    # how a particular engine chose to serialise the page.
    candidates = [rows_by_text(flat_lines)]
    if lines:
        rebuilt = rows_by_text([line.text for line in lines])
        for row in rebuilt:
            row["Method"] = "text (rebuilt lines)"
        candidates.append(rebuilt)
    text_rows = max(candidates, key=row_quality)

    columns = find_columns(lines)
    format_rows = rows_by_format(lines, columns) if columns is not None else []

    if row_quality(format_rows) > row_quality(text_rows):
        note = (f"read from the column layout, which read {len(format_rows)} "
                f"students against {len(text_rows)} from the plain text")
        for row in format_rows:
            row["Extraction Notes"] = (
                f"{row['Extraction Notes']}; {note}" if row["Extraction Notes"]
                else note)
        return merge_rows(format_rows, text_rows, "format"), "format"

    if text_rows:
        return merge_rows(text_rows, format_rows, "text"), "text"
    return format_rows, "format" if format_rows else "text"


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
    pdfs = sorted(list(tests.glob("sample_sap*.pdf"))
                  + list(tests.glob("sample_nsc*.pdf")))
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
    # The text reader runs first, so it normally reads everything and a zero
    # here is the expected result, not a problem. A non-zero count means the
    # column layout found students the plain text could not.
    print(f"read by the text reader: {s['by_text']}    "
          f"by the column layout: {s['by_format']}")
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
                 text="The plain text is read first, then the column layout "
                      "arranges what it could not separate."
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
                   f"Read by the text reader: {s['by_text']}    "
                   f"by the column layout: {s['by_format']}\n"
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
