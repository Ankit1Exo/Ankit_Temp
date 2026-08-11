r"""
TIAA participant reports  --  Part ID / Name / Date of Birth extractor
======================================================================
Reads searchable TIAA-CREF participant PDFs and writes Part ID (SSN),
Name and Date of Birth to one combined Excel workbook.

TWO LAYOUTS, DETECTED PER PAGE
    A  TABLE      e.g. "Small Cash Out Monitoring Report"
       One participant per printed row under a column header:

         PART ID       NAME                DIV/LOC  STATUS      DATE OF   DATE OF ...
                                                                BIRTH     HIRE
         123-45-6789   ACKERMAN, KENTON C  0001     Terminated  01/15/63  03/01/90

    B  LABEL      e.g. "Participant Contribution Investment Allocation Report"
       One participant per block of "LABEL: value" pairs:

         NAME: ABBEY, BORIS S      DATE OF BIRTH: 04/12/1973   DATE OF TERMINATION:
         PART ID: 987654321        DATE OF HIRE: 06/01/2015    DATE OF DEATH:
         STATUS: Active            DATE OF ENTRY: 07/01/2015   DATE OF REHIRE:

    Detection is per PAGE, not per file, so a single PDF may hold both.
    A page that is neither (a detail or continuation page with no
    participant identity on it) simply produces no rows.

WHY BLOCKS ARE GROUPED BY POSITION, NOT BY CAPTION
    These reports OMIT a caption entirely when its value is empty. A
    participant with no name on file prints no "NAME:" at all -- not
    "NAME:" followed by blank. Same for PART ID and the rest.

    So a participant block cannot be delimited by watching for a caption
    to repeat. The caption that would have marked the boundary is exactly
    the one that is missing, and the next participant's name gets read
    backwards onto the previous participant's Part ID:

         printed                        read as
         PART ID: ...4444               ...4444  CLIFTON, BERA S   <- wrong
         DATE OF BIRTH: 03/04/1971               person
         NAME: CLIFTON, BERA S          (row lost entirely)
         DATE OF BIRTH: 05/06/1972

    Vertical position is never omitted. Blocks are therefore grouped by
    vertical gap and each is parsed in isolation, so a missing caption
    leaves a blank field and cannot leak a value across a boundary. A
    repeated caption inside one block still splits it, which covers two
    blocks printed with no gap between them.

    selftest() asserts this on a synthetic page built from the case above.
    Raising BLOCK_GAP_HEIGHTS until the grouping collapses reproduces the
    original fault, which is how that test was checked to have teeth.

A BLANK PART ID KEEPS THE PARTICIPANT
    A row or block with no Part ID is still written out, with the cell
    empty, the reason in Notes and a count in Reconciliation. Dropping it
    would be a silent loss. In the table layout such a row is only
    accepted if it still looks like a participant -- a name plus a date or
    a status -- otherwise report titles and page footers would become rows.

WHY THE DATE COLUMN NEEDS GEOMETRY
    The table layout prints four date columns -- BIRTH, HIRE, REHIRE,
    TERMINATION -- and any of them may be blank on any given row. Reading
    "the first date on the line" is therefore wrong the moment a birth
    date is missing: the hire date slides left into its place and is
    silently written out as a date of birth. Nothing about the resulting
    workbook looks broken.

    So the column header band is read first, clustered into columns by x
    position, and every date token on a data row is assigned to whichever
    column caption it sits under. Only the token under BIRTH is kept.

    Header captions stack across two printed lines ("DATE OF" above
    "BIRTH"), so the band is the header line plus the line above it, and
    the two are merged by horizontal overlap.

    Continuation pages that repeat no header reuse the last column map
    seen in the same file. If no header has been seen at all, the parser
    falls back to pattern order (SSN, then name up to the first date or
    status word, then the leftmost date) and every row it produces is
    flagged in the Notes column, because that fallback is exactly the
    guess described above.

LINE RECONSTRUCTION
    Lines are rebuilt from extract_words() coordinates and clustered by y.
    page.extract_text(layout=True) is NOT used anywhere: it paints
    characters into a fixed grid sized from page.width, and on a wide
    landscape report characters landing in the same cell are dropped.
    That failure is page-geometry dependent, so it shows up on some files
    and not others -- the worst kind of extraction bug to inherit.

ENGINES
    pdfplumber is the default. PyMuPDF reads the same positioned words
    far faster and is available as --engine mupdf, but speed here comes
    from processing files in parallel, which cannot change WHAT is read,
    rather than from swapping the engine, which can. Run --verify over
    your own files first if you intend to switch.

COMPLETENESS
    The Reconciliation sheet counts, per file, how many Part IDs were
    sitting on the pages against how many rows were written. A file is
    Complete only when

        rows written == Part IDs on page + rows with no Part ID printed
        no row is missing a name
        no row had dates on it but none under the BIRTH column

    A blank date of birth does NOT flag the file -- the field is often
    genuinely empty, and a flag that fires on every file is a flag people
    learn to ignore. A date that failed to land under the BIRTH column
    does flag it, because that is a possible misread rather than an empty
    field. Both counts are on the sheet either way.

    Check that sheet before signing off a run. A workbook holding 40 of 47
    participants looks completely healthy otherwise -- there is nothing in
    it that says "incomplete".

HANDLING OF PERSONAL DATA
    Every row of the output is name + SSN + date of birth, which is a
    re-identification file. Write it to a controlled location, not a
    desktop or a personal drive, and give it a retention and disposal
    date. Nothing here should be pasted into chat tools or e-mail.

    The console, the QA dump and the log NEVER print an unmasked SSN.
    --mask writes the workbook itself with masked Part IDs
    (***-**-6789) for copies that need to circulate more widely.

USAGE
    python "260811 AM tiaa participant id name dob extractor.py"
        Tkinter folder picker, progress bar, combined workbook.

    python "260811 AM tiaa participant id name dob extractor.py" \
        --src <folder> --out <file.xlsx> [--mask] [--engine pdfplumber|mupdf]
        Headless run, no window.

    ... --selftest                Build a synthetic PDF (fabricated data,
                                  no real PII) and assert the parse.
    ... --qa <file.pdf>           Masked dump of how each line was split.
    ... --lines <file.pdf> [--page N]
                                  Masked geometry dump: every rebuilt line,
                                  each word's x span and the column it was
                                  read as. Letters print as A and digits as
                                  9, so the output can be shared to diagnose
                                  a misread without moving any PII.
    ... --verify <folder>         Read every PDF with both engines and
                                  report disagreements. Counts only.

FIELDS WRITTEN
    Part ID (SSN, normalised to 123-45-6789), Name as printed plus a
    Last / First / Middle / Suffix split, Date of Birth and Date of Death.
    Each date and the Part ID also keep an "As Printed" column, so nothing
    is lost to normalisation.
"""
from __future__ import annotations

import argparse
import os
import re
import sys
import traceback
import unicodedata
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

try:
    import pdfplumber
except ImportError:  # pragma: no cover - environment dependent
    pdfplumber = None

try:
    import pymupdf as fitz
except ImportError:  # pragma: no cover - environment dependent
    try:
        import fitz
    except ImportError:
        fitz = None

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# configuration
# ---------------------------------------------------------------------------

OUTPUT_XLSX_NAME = "260811 AM tiaa participant id name dob.xlsx"

DEFAULT_ENGINE = "pdfplumber"

OUTPUT_COLUMNS = [
    "File Name",
    "Page",
    "Layout",
    "Part ID (SSN)",
    "Part ID As Printed",
    "Name As Printed",
    "Last Name",
    "First Name",
    "Middle Name",
    "Suffix",
    "Date Of Birth",
    "Date Of Birth As Printed",
    "Date Of Death",
    "Date Of Death As Printed",
    "Notes",
]

RECON_COLUMNS = [
    "File Name",
    "Pages",
    "Layouts Seen",
    "Part IDs On Page",
    "Rows Written",
    "Rows Missing Name",
    "Rows Missing DOB",
    "Rows With Suspicious DOB",
    "Rows Without Part ID",
    "Rows With Non-SSN Part ID",
    "Complete",
    "Error",
]

# Written into Notes when a row carried dates but none of them sat under the
# DATE OF BIRTH column. Distinct from a row that printed no dates at all: the
# first is a possible misread, the second is just an empty field.
SUSPICIOUS_DOB_NOTE = "no date sat under the DATE OF BIRTH column"
NO_PART_ID_NOTE = "no Part ID printed"

# horizontal gap, in points, that separates one header caption from the next.
# Captions inside one column ("DATE OF HIRE") sit a few points apart; adjacent
# columns on these reports are 20pt or more apart.
HEADER_COLUMN_GAP = 14.0

# how far right of the Name caption a printed name may run before it is taken
# to belong to the next column. Set from the next caption's left edge.
NAME_STOP_PAD = 2.0

MIN_FILES_FOR_PARALLEL = 4


# ---------------------------------------------------------------------------
# text normalisation
# ---------------------------------------------------------------------------

# Every dash-like codepoint a PDF may carry, folded to ASCII "-" before any
# matching happens. A U+2010 in "123‐45‐6789" otherwise hides an SSN
# from every pattern below.
_DASHES = "‐‑‒–—―−﹘﹣－­"
_SPACES = "       "
_FOLD_MAP = {ord(c): "-" for c in _DASHES}
_FOLD_MAP.update({ord(c): " " for c in _SPACES})


def fold(text: str) -> str:
    """Normalise unicode punctuation and whitespace to ASCII equivalents."""
    if not text:
        return ""
    return unicodedata.normalize("NFKC", str(text)).translate(_FOLD_MAP)


def squeeze(text: str) -> str:
    """Collapse runs of whitespace to single spaces."""
    return re.sub(r"\s+", " ", fold(text)).strip()


SSN_TOKEN_RE = re.compile(
    r"^(?:"
    r"\d{3}-\d{2}-\d{4}"          # 123-45-6789
    r"|[Xx*]{3}-[Xx*]{2}-\d{4}"   # XXX-XX-6789
    r"|\d{9}"                     # 123456789
    r"|[Xx*]{5}\d{4}"             # XXXXX6789
    r")$"
)

SSN_ANY_RE = re.compile(
    r"(?<![\d-])(?:"
    r"\d{3}-\d{2}-\d{4}"
    r"|[Xx*]{3}-[Xx*]{2}-\d{4}"
    r"|\d{9}"
    r"|[Xx*]{5}\d{4}"
    r")(?![\d-])"
)

DATE_TOKEN_RE = re.compile(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{2}|\d{4})$")
DATE_ANY_RE = re.compile(r"(?<!\d)(\d{1,2})[/-](\d{1,2})[/-](\d{2}|\d{4})(?!\d)")

STATUS_WORDS = {
    "ACTIVE", "TERMINATED", "TERMINATE", "RETIRED", "DECEASED", "INACTIVE",
    "LEAVE", "DISABLED", "SUSPENDED", "REHIRED", "ELIGIBLE", "INELIGIBLE",
}

NAME_SUFFIXES = {
    "JR", "SR", "II", "III", "IV", "V", "VI", "MD", "PHD", "DDS", "DVM",
    "ESQ", "CPA", "RN",
}

# Plausibility window for a date of birth. Anything outside it is still
# written out, but the row carries a note -- a wrong DOB that looks
# reasonable is far more dangerous than one that is obviously junk.
DOB_MIN_YEAR = 1890
DOB_MAX_YEAR = 2030


def norm_ssn(raw: str) -> str:
    """Return an SSN as 123-45-6789 (or XXX-XX-6789) regardless of input form."""
    compact = re.sub(r"[^0-9Xx*]", "", fold(raw))
    if len(compact) == 9:
        compact = compact.upper().replace("*", "X")
        return f"{compact[0:3]}-{compact[3:5]}-{compact[5:9]}"
    return squeeze(raw).upper()


def mask_ssn(value: str) -> str:
    """***-**-6789 -- safe for console, logs and circulated copies."""
    digits = re.sub(r"\D", "", fold(value))
    if len(digits) >= 4:
        return f"***-**-{digits[-4:]}"
    return "***-**-****" if value else ""


def norm_date(raw: str):
    """Return (normalised MM/DD/YYYY, note). Empty string if unparseable."""
    m = DATE_ANY_RE.search(fold(raw))
    if not m:
        return "", ""
    month, day, year = int(m.group(1)), int(m.group(2)), m.group(3)
    if len(year) == 2:
        # A two-digit year on a date of birth is a 20th century year far more
        # often than not; 30 is the usual pivot on these reports.
        year = 1900 + int(year) if int(year) > 30 else 2000 + int(year)
    else:
        year = int(year)
    if not (1 <= month <= 12) or not (1 <= day <= 31):
        return "", "date of birth does not parse"
    note = ""
    if not (DOB_MIN_YEAR <= year <= DOB_MAX_YEAR):
        note = f"date of birth year {year} outside plausible range"
    return f"{month:02d}/{day:02d}/{year}", note


def split_name(printed: str):
    """Split "LAST, FIRST MIDDLE" into parts. Returns (last, first, middle,
    suffix, note)."""
    text = squeeze(printed).strip(" ,")
    if not text:
        # No note: whichever parser produced the blank has already said so,
        # and in more specific terms than this function could.
        return "", "", "", "", ""

    note = ""
    if "," in text:
        last_part, _, rest = text.partition(",")
    else:
        # No comma. The report format is Last, First -- so this is either a
        # single-word name or something the column split got wrong. Assume
        # "FIRST MIDDLE LAST" and flag it, rather than guessing silently.
        tokens = text.split()
        if len(tokens) == 1:
            return tokens[0], "", "", "", "no comma in name; treated as surname only"
        last_part, rest = tokens[-1], " ".join(tokens[:-1])
        note = "no comma in name; word order assumed First Middle Last"

    last_tokens = last_part.split()
    rest_tokens = rest.split()

    suffix = ""
    for bucket in (rest_tokens, last_tokens):
        while bucket and bucket[-1].strip(".").upper() in NAME_SUFFIXES:
            suffix = bucket.pop().strip(".").upper()

    last = " ".join(last_tokens)
    first = rest_tokens[0] if rest_tokens else ""
    middle = " ".join(rest_tokens[1:]) if len(rest_tokens) > 1 else ""
    if not last:
        note = (note + "; " if note else "") + "surname missing"
    return last, first, middle, suffix, note


def looks_like_part_id(text: str) -> bool:
    """True for anything that could be a printed Part ID.

    Deliberately looser than the SSN patterns: a row whose Part ID is not
    SSN-shaped is still a participant, and dropping it would be a silent
    loss. Loose enough to catch those, tight enough to reject page footers
    and running totals.
    """
    compact = re.sub(r"[^0-9A-Za-z*-]", "", fold(text))
    if len(compact) < 4:
        return False
    return len(re.findall(r"[0-9Xx*]", compact)) >= 4


# ---------------------------------------------------------------------------
# word extraction and line rebuilding
# ---------------------------------------------------------------------------

def words_pdfplumber(page):
    out = []
    for w in page.extract_words(use_text_flow=False, keep_blank_chars=False):
        out.append({
            "text": fold(w["text"]),
            "x0": float(w["x0"]),
            "x1": float(w["x1"]),
            "top": float(w["top"]),
            "bottom": float(w["bottom"]),
        })
    return out


def words_pymupdf(page):
    out = []
    for x0, y0, x1, y1, text, *_ in page.get_text("words"):
        out.append({
            "text": fold(text),
            "x0": float(x0),
            "x1": float(x1),
            "top": float(y0),
            "bottom": float(y1),
        })
    return out


def cluster_rows(words, y_tol=None):
    """Group words into printed lines by vertical centre.

    Tolerance defaults to half the median glyph height, so it adapts to the
    point size instead of assuming one.
    """
    if not words:
        return []
    heights = sorted(w["bottom"] - w["top"] for w in words)
    median = heights[len(heights) // 2] or 8.0
    tol = y_tol if y_tol is not None else max(1.5, median * 0.5)

    ordered = sorted(words, key=lambda w: ((w["top"] + w["bottom"]) / 2.0, w["x0"]))
    rows, current, centre = [], [], None
    for w in ordered:
        c = (w["top"] + w["bottom"]) / 2.0
        if current and abs(c - centre) > tol:
            rows.append(sorted(current, key=lambda x: x["x0"]))
            current = []
        current.append(w)
        centre = sum((x["top"] + x["bottom"]) / 2.0 for x in current) / len(current)
    if current:
        rows.append(sorted(current, key=lambda x: x["x0"]))
    return rows


def row_text(row) -> str:
    return " ".join(w["text"] for w in row)


def centre_x(word) -> float:
    return (word["x0"] + word["x1"]) / 2.0


def read_pages(path: Path, engine: str):
    """Yield (page_number, words) for every page. One place per engine."""
    if engine == "mupdf":
        if fitz is None:
            raise RuntimeError("PyMuPDF is not installed")
        with fitz.open(str(path)) as doc:
            for i, page in enumerate(doc, start=1):
                yield i, words_pymupdf(page)
    else:
        if pdfplumber is None:
            raise RuntimeError("pdfplumber is not installed")
        with pdfplumber.open(str(path)) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                yield i, words_pdfplumber(page)


def available_engines():
    found = []
    if pdfplumber is not None:
        found.append("pdfplumber")
    if fitz is not None:
        found.append("mupdf")
    return found


# ---------------------------------------------------------------------------
# layout A -- the column table
# ---------------------------------------------------------------------------

# Checked in order. REHIRE must be tested before HIRE, since "DATE OF REHIRE"
# contains "HIRE" and would otherwise be labelled as the hire column.
COLUMN_KEYWORDS = [
    ("part_id", ("PARTID", "PART ID", "PARTICIPANT ID")),
    ("name", ("NAME",)),
    ("div_loc", ("DIV/LOC", "DIVLOC", "DIV / LOC", "DIV")),
    ("status", ("STATUS",)),
    ("dob", ("BIRTH",)),
    ("death", ("DEATH",)),
    ("rehire", ("REHIRE",)),
    ("hire", ("HIRE",)),
    ("term", ("TERMINATION",)),
    ("balance", ("BALANCE",)),
]

DATE_COLUMN_KEYS = ("dob", "death", "hire", "rehire", "term")


def find_header_band(rows):
    """Return (header words, index of the header row) or (None, -1).

    The caption block spans two printed lines on these reports -- "DATE OF"
    sits above "BIRTH" -- so the line above the PART ID line is pulled in
    whenever it carries caption text.
    """
    for i, row in enumerate(rows):
        upper = [w["text"].upper() for w in row]
        joined = " ".join(upper)
        has_part_id = ("PART" in upper and "ID" in upper) or "PART ID" in joined
        if not (has_part_id and "NAME" in upper):
            continue
        band = list(row)
        if i > 0:
            above_text = " ".join(w["text"].upper() for w in rows[i - 1])
            if "DATE" in above_text or "VESTED" in above_text or "TOTAL" in above_text:
                band += rows[i - 1]
        return band, i
    return None, -1


def cluster_columns(band, gap=HEADER_COLUMN_GAP):
    """Merge header words into columns by horizontal overlap / proximity."""
    clusters = []
    for w in sorted(band, key=lambda x: x["x0"]):
        if clusters and w["x0"] <= clusters[-1]["x1"] + gap:
            c = clusters[-1]
            c["x0"] = min(c["x0"], w["x0"])
            c["x1"] = max(c["x1"], w["x1"])
            c["words"].append(w)
        else:
            clusters.append({"x0": w["x0"], "x1": w["x1"], "words": [w]})
    for c in clusters:
        # Caption words are ordered left-to-right then top-to-bottom so
        # "DATE OF" + "BIRTH" reads as one caption.
        ordered = sorted(c["words"], key=lambda w: (w["top"], w["x0"]))
        c["text"] = " ".join(w["text"] for w in ordered).upper()
        c["centre"] = (c["x0"] + c["x1"]) / 2.0
    return clusters


def build_column_map(rows):
    """Return (column map, header row index). Map is key -> {x0,x1,centre}."""
    band, index = find_header_band(rows)
    if band is None:
        return None, -1

    colmap = {}
    for cluster in cluster_columns(band):
        text = cluster["text"]
        flat = text.replace(" ", "")
        for key, keywords in COLUMN_KEYWORDS:
            if key in colmap:
                continue
            if any(kw in text or kw.replace(" ", "") in flat for kw in keywords):
                colmap[key] = {
                    "x0": cluster["x0"],
                    "x1": cluster["x1"],
                    "centre": cluster["centre"],
                }
                break
    if "part_id" not in colmap or "name" not in colmap:
        return None, -1
    return colmap, index


def part_id_right_edge(colmap) -> float:
    """Boundary between the Part ID column and the Name column."""
    return (colmap["part_id"]["x1"] + colmap["name"]["x0"]) / 2.0


def name_right_edge(colmap) -> float:
    """Where a printed name must stop.

    The next caption's LEFT edge, not the midpoint to it: names are left
    aligned and run much wider than the four-letter "NAME" caption, so a
    midpoint would truncate "ACKERMAN, KENTON CLYDE".
    """
    name_x0 = colmap["name"]["x0"]
    candidates = [
        c["x0"] for key, c in colmap.items()
        if key != "name" and c["x0"] > name_x0
    ]
    return (min(candidates) - NAME_STOP_PAD) if candidates else float("inf")


def nearest_date_column(colmap, x):
    """Which date column a token at centre x sits under."""
    anchors = [(key, colmap[key]["centre"]) for key in DATE_COLUMN_KEYS if key in colmap]
    if not anchors:
        return None
    return min(anchors, key=lambda kv: abs(kv[1] - x))[0]


def parse_table_page(rows, colmap, header_index, file_name, page_num, inherited):
    """Extract one row per participant from a column-table page."""
    out, candidates = [], 0
    right_of_part_id = part_id_right_edge(colmap)
    name_stop = name_right_edge(colmap)
    part_id_left = colmap["part_id"]["x0"] - 6.0

    for row in rows[header_index + 1:]:
        id_words = [
            w for w in row
            if w["x0"] >= part_id_left and centre_x(w) <= right_of_part_id
        ]
        id_word = next((w for w in id_words if looks_like_part_id(w["text"])), None)

        notes = []
        if inherited:
            notes.append("column layout inherited from an earlier page")

        raw_id = id_word["text"] if id_word else ""
        if id_word is None:
            # The report leaves the cell empty rather than printing a
            # placeholder, so the participant has to be recognised from the
            # rest of the row instead.
            notes.append(NO_PART_ID_NOTE + " on this row")
        elif not SSN_TOKEN_RE.match(raw_id):
            notes.append("Part ID is not SSN-shaped")

        name_left = id_word["x1"] if id_word else colmap["name"]["x0"] - 6.0
        name_words = [
            w for w in row
            if w["x0"] > name_left and centre_x(w) < name_stop
        ]
        # Second guard, independent of geometry: a name never contains a
        # date, an SSN or a status word, so anything from there rightwards
        # belongs to a later column even if the caption positions mislead.
        kept = []
        for w in name_words:
            t = w["text"]
            if DATE_TOKEN_RE.match(t) or SSN_TOKEN_RE.match(t):
                break
            if t.strip(".,").upper() in STATUS_WORDS:
                break
            kept.append(t)
        printed_name = squeeze(" ".join(kept))

        dob_raw, death_raw = "", ""
        date_words = [w for w in row if DATE_TOKEN_RE.match(w["text"])]
        for w in date_words:
            column = nearest_date_column(colmap, centre_x(w))
            if column == "dob" and not dob_raw:
                dob_raw = w["text"]
            elif column == "death" and not death_raw:
                death_raw = w["text"]

        if id_word is None:
            # With no Part ID to anchor on, only accept the row if it still
            # looks like a participant. Otherwise report titles, period
            # headings and page footers would all become rows.
            has_status = any(
                w["text"].strip(".,").upper() in STATUS_WORDS for w in row
            )
            if not (printed_name and (date_words or has_status)):
                continue
        else:
            candidates += 1

        if not printed_name:
            notes.append("no name found in the name column")
        if not dob_raw and date_words:
            notes.append(SUSPICIOUS_DOB_NOTE)

        out.append(make_row(
            file_name, page_num, "table (inherited)" if inherited else "table",
            raw_id, printed_name, dob_raw, notes, death_raw,
        ))
    return out, candidates


# ---------------------------------------------------------------------------
# layout B -- the labelled block
# ---------------------------------------------------------------------------

# Every caption these reports print. Matched longest-first so "DATE OF BIRTH"
# is never read as "DATE OF" plus stray text. Add new captions here; the
# parser needs no other change.
LABELS = {
    "AFTER-TAX DEFERRAL % OR $": "after_tax",
    "PRE-TAX DEFERRAL % OR $": "pre_tax",
    "SOCIAL SECURITY NUMBER": "part_id",
    "DATE OF TERMINATION": "term",
    "PARTICIPANT NAME": "name",
    "DATE OF REHIRE": "rehire",
    "DATE OF BIRTH": "dob",
    "DATE OF DEATH": "death",
    "DATE OF ENTRY": "entry",
    "DATE OF HIRE": "hire",
    "PARTICIPANT ID": "part_id",
    "EMPLOYEE ID": "employee_id",
    "PLAN NUMBER": "plan",
    "DIV / LOC": "div_loc",
    "PART ID": "part_id",
    "DIV/LOC": "div_loc",
    "STATUS": "status",
    "NAME": "name",
    "SSN": "part_id",
}

_LABEL_ALTERNATION = "|".join(
    re.escape(label) for label in sorted(LABELS, key=len, reverse=True)
)
LABEL_RE = re.compile(r"(?<![A-Z])(" + _LABEL_ALTERNATION + r")\s*:", re.IGNORECASE)

LABEL_LAYOUT_RE = re.compile(r"(?:PART\s*ID|PARTICIPANT\s*ID|SSN)\s*:", re.IGNORECASE)


def label_pairs(line: str):
    """Split "A: x   B: y   C:" into [(key, value, match_start), ...]."""
    matches = list(LABEL_RE.finditer(line))
    pairs = []
    for i, m in enumerate(matches):
        end = matches[i + 1].start() if i + 1 < len(matches) else len(line)
        key = LABELS[m.group(1).upper()]
        pairs.append((key, line[m.end():end].strip(" \t:"), m.start()))
    return pairs


# Labels that belong to a participant's identity block. Anything else on the
# page (deferral percentages, fund allocations) is read but not used to decide
# where one participant ends and the next begins.
PARTICIPANT_LABEL_KEYS = {
    "part_id", "name", "dob", "death", "status", "hire", "rehire", "term",
    "entry", "div_loc", "employee_id", "plan",
}

# Words that turn "NAME:" into somebody else's caption. The allocation detail
# below a participant block prints fund and plan names with the same word, and
# reading one of those as a participant name would invent a record.
NAME_PREFIX_BLOCK = {
    "FUND", "PLAN", "INVESTMENT", "OPTION", "SOURCE", "CONTRACT", "EMPLOYER",
    "COMPANY", "ACCOUNT", "PRODUCT", "VENDOR", "SUBACCOUNT", "PORTFOLIO",
    "FIRST", "LAST", "MIDDLE",
}

# How many glyph heights of blank space end a participant block. Lines inside
# one block are printed about 1.5 heights apart; the gap to the next block, or
# down to the allocation detail, is far larger.
BLOCK_GAP_HEIGHTS = 2.5


def name_is_participant(line: str, position: int) -> bool:
    """False when "NAME:" is the tail of a compound caption like "FUND NAME:"."""
    before = line[:position].rstrip()
    if not before:
        return True
    return before.split()[-1].strip(":").upper() not in NAME_PREFIX_BLOCK


def label_blocks(rows):
    """Group label-bearing lines into participant blocks by vertical gap.

    This is the whole reason the parser does not lose track of who owns which
    value. These reports omit a caption entirely when its value is empty -- a
    participant with no name on file prints no "NAME:" at all, rather than
    "NAME:" followed by blank. So a block cannot be delimited by watching for
    a caption to repeat: the caption that would have marked the boundary is
    exactly the one that is missing, and the next participant's name gets
    read back onto the previous participant's Part ID.

    Vertical position is not omitted. Grouping on it means a missing caption
    leaves a blank field and nothing more.
    """
    heights, entries = [], []
    for row in rows:
        line = squeeze(row_text(row))
        if not line:
            continue
        heights.append(sum(w["bottom"] - w["top"] for w in row) / len(row))
        pairs = label_pairs(line)
        if pairs:
            entries.append({
                "y": sum((w["top"] + w["bottom"]) / 2 for w in row) / len(row),
                "pairs": pairs,
                "line": line,
            })
    if not entries:
        return []

    median_height = sorted(heights)[len(heights) // 2] or 8.0
    gap_limit = median_height * BLOCK_GAP_HEIGHTS

    blocks, current, previous_y = [], [], None
    for entry in entries:
        if previous_y is not None and (entry["y"] - previous_y) > gap_limit:
            blocks.append(current)
            current = []
        current.append(entry)
        previous_y = entry["y"]
    blocks.append(current)
    return blocks


def records_from_block(block):
    """One block usually holds one participant. A repeated caption inside it
    means two blocks were printed with no gap between them, so split there."""
    records, record = [], {}
    for entry in block:
        for key, value, position in entry["pairs"]:
            if key not in PARTICIPANT_LABEL_KEYS:
                continue
            if key == "name" and not name_is_participant(entry["line"], position):
                continue
            if value and record.get(key) and value != record[key]:
                records.append(record)
                record = {}
            # Blank values are still stored: a caption that was printed with
            # nothing after it is evidence this is a participant block.
            if value or key not in record:
                record[key] = value
    if record:
        records.append(record)
    return records


def parse_label_page(rows, file_name, page_num):
    """Extract participant records from a "LABEL: value" page."""
    out, candidates = [], 0
    for block in label_blocks(rows):
        for record in records_from_block(block):
            part_id = record.get("part_id", "")
            name = record.get("name", "")
            # A lone caption is not a participant. Two or more identity
            # captions in one block is.
            if not part_id and not (name and len(record) >= 2):
                continue
            notes = []
            if part_id:
                candidates += 1
            else:
                notes.append(NO_PART_ID_NOTE + " in this block")
            if not name:
                notes.append("no name printed in this block")
            out.append(make_row(
                file_name, page_num, "label", part_id, name,
                record.get("dob", ""), notes, record.get("death", ""),
            ))
    return out, candidates


# ---------------------------------------------------------------------------
# fallback -- no column header has ever been seen for this file
# ---------------------------------------------------------------------------

def parse_pattern_rows(rows, file_name, page_num):
    """Last resort: read a table row by pattern order alone.

    Every row produced here is flagged, because with no header the birth
    date is only "the leftmost date", and that is wrong on any row whose
    birth date is blank.
    """
    out, candidates = [], 0
    for row in rows:
        texts = [w["text"] for w in row]
        idx = next((i for i, t in enumerate(texts) if SSN_TOKEN_RE.match(t)), None)
        if idx is None:
            continue
        candidates += 1
        name_tokens = []
        dob_raw = ""
        for t in texts[idx + 1:]:
            if DATE_TOKEN_RE.match(t):
                dob_raw = t
                break
            if t.strip(".,").upper() in STATUS_WORDS or SSN_TOKEN_RE.match(t):
                continue
            if any(ch.isdigit() for ch in t):
                continue
            name_tokens.append(t)
        out.append(make_row(
            file_name, page_num, "table (no header)", texts[idx],
            squeeze(" ".join(name_tokens)), dob_raw,
            ["no column header found; date of birth taken as the leftmost date"],
        ))
    return out, candidates


# ---------------------------------------------------------------------------
# row assembly
# ---------------------------------------------------------------------------

def make_row(file_name, page_num, layout, raw_id, printed_name, raw_dob, notes,
             raw_death=""):
    notes = list(notes)
    last, first, middle, suffix, name_note = split_name(printed_name)
    if name_note:
        notes.append(name_note)

    dob, dob_note = norm_date(raw_dob)
    if raw_dob and not dob and not dob_note:
        notes.append("date of birth does not parse")
    elif dob_note:
        notes.append(dob_note)
    elif not raw_dob and SUSPICIOUS_DOB_NOTE not in notes:
        notes.append("no date of birth printed")

    # A blank date of death is the normal case, so it is never noted. Only a
    # value that will not parse is worth reporting.
    death, death_note = norm_date(raw_death)
    if raw_death and not death:
        notes.append(death_note or "date of death does not parse")

    return {
        "File Name": file_name,
        "Page": page_num,
        "Layout": layout,
        "Part ID (SSN)": norm_ssn(raw_id),
        "Part ID As Printed": squeeze(raw_id),
        "Name As Printed": squeeze(printed_name),
        "Last Name": last,
        "First Name": first,
        "Middle Name": middle,
        "Suffix": suffix,
        "Date Of Birth": dob,
        "Date Of Birth As Printed": squeeze(raw_dob),
        "Date Of Death": death,
        "Date Of Death As Printed": squeeze(raw_death),
        "Notes": "; ".join(n for n in notes if n),
    }


# ---------------------------------------------------------------------------
# per-file processing
# ---------------------------------------------------------------------------

def detect_layout(rows):
    """"label", "table" or None, decided from the page's own content."""
    for row in rows:
        if LABEL_LAYOUT_RE.search(row_text(row)):
            return "label"
        if re.match(r"^\s*NAME\s*:", row_text(row), re.IGNORECASE):
            return "label"
    if find_header_band(rows)[0] is not None:
        return "table"
    return None


def process_pdf(path, engine: str = DEFAULT_ENGINE):
    """Read one PDF. Returns (rows, diagnostics)."""
    path = Path(path)
    rows_out = []
    diag = {
        "File Name": path.name,
        "Pages": 0,
        "Layouts Seen": "",
        "Part IDs On Page": 0,
        "Rows Written": 0,
        "Rows Missing Name": 0,
        "Rows Missing DOB": 0,
        "Rows With Suspicious DOB": 0,
        "Rows Without Part ID": 0,
        "Rows With Non-SSN Part ID": 0,
        "Complete": "",
        "Error": "",
    }
    layouts = []
    last_colmap = None

    try:
        for page_num, words in read_pages(path, engine):
            diag["Pages"] += 1
            rows = cluster_rows(words)
            if not rows:
                continue

            layout = detect_layout(rows)
            if layout == "label":
                page_rows, candidates = parse_label_page(rows, path.name, page_num)
            elif layout == "table":
                colmap, header_index = build_column_map(rows)
                if colmap:
                    last_colmap = colmap
                    page_rows, candidates = parse_table_page(
                        rows, colmap, header_index, path.name, page_num, False
                    )
                else:
                    page_rows, candidates = parse_pattern_rows(rows, path.name, page_num)
            elif last_colmap is not None and any(
                SSN_TOKEN_RE.match(w["text"]) for w in words
            ):
                # A continuation page: participant rows, header printed only
                # on the first page of the run.
                page_rows, candidates = parse_table_page(
                    rows, last_colmap, -1, path.name, page_num, True
                )
                layout = "table (inherited)"
            else:
                continue

            if page_rows or candidates:
                layouts.append(layout)
            rows_out.extend(page_rows)
            diag["Part IDs On Page"] += candidates
    except Exception as exc:  # pragma: no cover - depends on the file
        diag["Error"] = f"{type(exc).__name__}: {exc}"

    diag["Rows Written"] = len(rows_out)
    diag["Rows Missing Name"] = sum(1 for r in rows_out if not r["Name As Printed"])
    diag["Rows Missing DOB"] = sum(1 for r in rows_out if not r["Date Of Birth"])
    diag["Rows With Suspicious DOB"] = sum(
        1 for r in rows_out if SUSPICIOUS_DOB_NOTE in r["Notes"]
    )
    diag["Rows Without Part ID"] = sum(
        1 for r in rows_out if not r["Part ID (SSN)"]
    )
    diag["Rows With Non-SSN Part ID"] = sum(
        1 for r in rows_out
        if r["Part ID (SSN)"]
        and not re.match(r"^[\dX]{3}-[\dX]{2}-\d{4}$", r["Part ID (SSN)"])
    )
    diag["Layouts Seen"] = ", ".join(sorted(set(layouts))) or "none"
    # Every Part ID on the page became a row, plus the rows that legitimately
    # had no Part ID printed. A blank date of birth is normal and does not
    # flag the file; a date that failed to land under the BIRTH column does.
    diag["Complete"] = "YES" if (
        not diag["Error"]
        and diag["Rows Written"] == diag["Part IDs On Page"] + diag["Rows Without Part ID"]
        and diag["Rows Missing Name"] == 0
        and diag["Rows With Suspicious DOB"] == 0
    ) else "CHECK"
    return rows_out, diag


def _worker(payload):
    """Top-level so Windows spawn can pickle it."""
    path, engine = payload
    return process_pdf(path, engine)


def process_folder(pdfs, engine=DEFAULT_ENGINE, workers=None, progress=None):
    """Read every PDF. Returns (rows, diagnostics)."""
    pdfs = list(pdfs)
    rows, diags = [], []

    if workers is None:
        workers = max(1, (os.cpu_count() or 2) - 1)
    # Windows has no fork, so each worker is a fresh interpreter costing about
    # a second to start. On a handful of files that start-up cost exceeds the
    # work itself.
    if len(pdfs) < MIN_FILES_FOR_PARALLEL or workers == 1:
        for i, path in enumerate(pdfs, start=1):
            r, d = process_pdf(path, engine)
            rows.extend(r)
            diags.append(d)
            if progress:
                progress(i, len(pdfs), Path(path).name)
        return rows, diags

    with ProcessPoolExecutor(max_workers=workers) as pool:
        payloads = [(str(p), engine) for p in pdfs]
        for i, (r, d) in enumerate(pool.map(_worker, payloads), start=1):
            rows.extend(r)
            diags.append(d)
            if progress:
                progress(i, len(pdfs), d["File Name"])
    return rows, diags


# ---------------------------------------------------------------------------
# workbook
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True)
FLAG_FILL = PatternFill("solid", fgColor="FCE4E4")


def _write_sheet(ws, columns, records, text_columns=()):
    ws.append(columns)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    text_indexes = {columns.index(c) + 1 for c in text_columns if c in columns}
    for record in records:
        ws.append([record.get(c, "") for c in columns])
        for idx in text_indexes:
            # Excel eats leading zeros and reformats 123-45-6789 unless the
            # cell is explicitly text.
            ws.cell(row=ws.max_row, column=idx).number_format = "@"

    widths = [len(c) for c in columns]
    for record in records:
        for i, c in enumerate(columns):
            widths[i] = max(widths[i], min(60, len(str(record.get(c, "")))))
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width + 2
    ws.freeze_panes = "A2"


def write_workbook(rows, diagnostics, dest: Path, masked=False) -> Path:
    dest = Path(dest)
    dest.parent.mkdir(parents=True, exist_ok=True)

    if masked:
        rows = [dict(r) for r in rows]
        for r in rows:
            r["Part ID (SSN)"] = mask_ssn(r["Part ID (SSN)"])
            r["Part ID As Printed"] = mask_ssn(r["Part ID As Printed"])

    wb = Workbook()
    _write_sheet(
        wb.active, OUTPUT_COLUMNS, rows,
        text_columns=("Part ID (SSN)", "Part ID As Printed",
                      "Date Of Birth", "Date Of Birth As Printed",
                      "Date Of Death", "Date Of Death As Printed"),
    )
    wb.active.title = "Data"

    recon = wb.create_sheet("Reconciliation")
    _write_sheet(recon, RECON_COLUMNS, diagnostics)
    complete_col = RECON_COLUMNS.index("Complete") + 1
    for row_idx in range(2, recon.max_row + 1):
        if recon.cell(row=row_idx, column=complete_col).value != "YES":
            for col_idx in range(1, len(RECON_COLUMNS) + 1):
                recon.cell(row=row_idx, column=col_idx).fill = FLAG_FILL

    wb.save(dest)
    return dest


# ---------------------------------------------------------------------------
# QA and verification -- masked output only
# ---------------------------------------------------------------------------

def qa_file(path, engine=DEFAULT_ENGINE, limit=25):
    """Masked dump of how each participant line was split."""
    rows, diag = process_pdf(path, engine)
    print(f"\n{Path(path).name}  --  {diag['Pages']} page(s), "
          f"layout(s): {diag['Layouts Seen']}")
    print(f"  Part IDs on page {diag['Part IDs On Page']}, "
          f"rows written {diag['Rows Written']}, complete: {diag['Complete']}")
    if diag["Error"]:
        print(f"  ERROR  {diag['Error']}")
    for r in rows[:limit]:
        print(f"    p{r['Page']:<4} {mask_ssn(r['Part ID (SSN)']):<12} "
              f"{r['Last Name']:<18} {r['First Name']:<12} "
              f"{r['Middle Name']:<8} {r['Suffix']:<5} {r['Date Of Birth']}")
        if r["Notes"]:
            print(f"          note: {r['Notes']}")
    if len(rows) > limit:
        print(f"    ... {len(rows) - limit} more row(s) not shown")


def _shape(text: str) -> str:
    """Reduce a word to its shape: letters to A, digits to 9.

    "CLARK,LISA" becomes "AAAAA,AAAA" and "177-72-5725" becomes
    "999-99-9999". Enough to diagnose a column or line-grouping fault,
    nothing that identifies anybody.
    """
    return re.sub(r"[A-Za-z]", "A", re.sub(r"\d", "9", fold(text)))


def which_column(colmap, word):
    """The column a word would be read as belonging to."""
    if not colmap:
        return "?"
    x = centre_x(word)
    if x <= part_id_right_edge(colmap):
        return "part_id"
    if x < name_right_edge(colmap):
        return "name"
    if DATE_TOKEN_RE.match(word["text"]):
        return nearest_date_column(colmap, x) or "?"
    return min(colmap.items(), key=lambda kv: abs(kv[1]["centre"] - x))[0]


def dump_lines(path, engine=DEFAULT_ENGINE, page_no=None, limit=80):
    """Masked geometry dump: every rebuilt line, its y, and each word's x
    span and column. Diagnoses a misread without moving any PII."""
    path = Path(path)
    print(f"{path.name}  --  shapes only, no participant data")
    colmap = None
    for page_num, words in read_pages(path, engine):
        if page_no and page_num != page_no:
            continue
        rows = cluster_rows(words)
        found, header_index = build_column_map(rows)
        if found:
            colmap = found
        print(f"\n--- page {page_num}   layout={detect_layout(rows)}   "
              f"header row={header_index}   lines={len(rows)}")
        if colmap:
            ordered = sorted(colmap.items(), key=lambda kv: kv[1]["x0"])
            print("    columns  " + "  ".join(
                f"{k}[{c['x0']:.0f}-{c['x1']:.0f}]" for k, c in ordered))
            print(f"    part id ends {part_id_right_edge(colmap):.0f}, "
                  f"name ends {name_right_edge(colmap):.0f}")
        for row in rows[:limit]:
            y = sum((w["top"] + w["bottom"]) / 2 for w in row) / len(row)
            body = " ".join(
                f"[{w['x0']:.0f}-{w['x1']:.0f} {which_column(colmap, w)}]"
                f"{_shape(w['text'])}"
                for w in row
            )
            print(f"  y={y:7.1f}  {body}")
        if len(rows) > limit:
            print(f"  ... {len(rows) - limit} more line(s) not shown")


def verify(folder, limit=None):
    """Read every PDF with both engines and report disagreements.

    Counts only -- no participant data is printed.
    """
    engines = available_engines()
    if len(engines) < 2:
        print("Both pdfplumber and PyMuPDF must be installed to verify.")
        return 1
    pdfs = sorted(Path(folder).rglob("*.pdf"))
    if limit:
        pdfs = pdfs[:limit]
    if not pdfs:
        print(f"No PDFs under {folder}")
        return 1

    mismatches = 0
    for path in pdfs:
        a_rows, _ = process_pdf(path, "pdfplumber")
        b_rows, _ = process_pdf(path, "mupdf")
        key = lambda rs: sorted(
            (r["Part ID (SSN)"], r["Name As Printed"], r["Date Of Birth"]) for r in rs
        )
        same = key(a_rows) == key(b_rows)
        mismatches += 0 if same else 1
        status = "identical" if same else "DIFFERS"
        print(f"  {status:<10} {path.name}  "
              f"pdfplumber={len(a_rows)}  mupdf={len(b_rows)}")
    print(f"\n{len(pdfs)} file(s), {mismatches} disagreement(s).")
    if mismatches == 0:
        print("Both engines agree on every file. --engine mupdf is safe here.")
    else:
        print("Engines disagree. Keep the default engine.")
    return 0


# ---------------------------------------------------------------------------
# selftest -- synthetic PDF, fabricated data, no real PII
# ---------------------------------------------------------------------------

def _build_test_pdf(dest: Path) -> Path:
    """Write a 3-page synthetic PDF. Every value below is made up."""
    if fitz is None:
        raise RuntimeError("PyMuPDF is required to build the selftest PDF")
    doc = fitz.open()
    put = lambda page, x, y, text: page.insert_text(
        (x, y), text, fontsize=8, fontname="cour"
    )

    # --- page 1: column table, header present --------------------------------
    p1 = doc.new_page(width=792, height=612)
    put(p1, 40, 60, "Small Cash Out Monitoring Report")
    put(p1, 355, 100, "DATE OF")
    put(p1, 585, 100, "DATE OF")
    put(p1, 660, 100, "VESTED TOTAL")
    for x, caption in [
        (40, "PART ID"), (110, "NAME"), (250, "DIV/LOC"), (305, "STATUS"),
        (360, "BIRTH"), (410, "DATE OF HIRE"), (490, "DATE OF REHIRE"),
        (580, "TERMINATION"), (655, "ACCOUNT BALANCE"),
    ]:
        put(p1, x, 110, caption)

    rows = [
        # part id, name, div, status, dob, hire, rehire, term, balance
        ("123-45-6789", "ACKERMAN, KENTON CLYDE", "0001", "Terminated",
         "01/15/1963", "03/01/1990", "", "06/30/2015", "$1,234.18"),
        # no birth date printed -- the hire date must NOT be read as the DOB
        ("234-56-7890", "CHILDERS,MORGAN N", "0002", "Terminated",
         "", "07/12/1988", "", "01/31/2015", "$88.15"),
        # running nine digits, and a suffix on the surname
        ("345678901", "SMITH JR, JOHN A", "0003", "Active",
         "12/25/1977", "05/02/2001", "05/02/2009", "", "$4,010.06"),
    ]
    y = 130
    for pid, name, div, status, dob, hire, rehire, term, bal in rows:
        put(p1, 40, y, pid)
        put(p1, 110, y, name)
        put(p1, 250, y, div)
        put(p1, 305, y, status)
        if dob:
            put(p1, 358, y, dob)
        if hire:
            put(p1, 415, y, hire)
        if rehire:
            put(p1, 495, y, rehire)
        if term:
            put(p1, 582, y, term)
        put(p1, 660, y, bal)
        y += 14
    put(p1, 40, y + 20, "Page 1 of 3")

    # --- page 2: labelled block ---------------------------------------------
    p2 = doc.new_page(width=792, height=612)
    put(p2, 40, 60, "Participant Contribution Investment Allocation Report")
    put(p2, 40, 100, "NAME: ABBEY, BORIS S")
    put(p2, 250, 100, "DATE OF BIRTH: 04/12/1973")
    put(p2, 430, 100, "DATE OF TERMINATION:")
    put(p2, 620, 100, "DIV/LOC:")
    put(p2, 40, 112, "PART ID: 987654321")
    put(p2, 250, 112, "DATE OF HIRE: 06/01/2015")
    put(p2, 430, 112, "DATE OF DEATH: 03/22/2024")
    put(p2, 620, 112, "PRE-TAX DEFERRAL % OR $:")
    put(p2, 40, 124, "STATUS: Active")
    put(p2, 250, 124, "DATE OF ENTRY: 07/01/2015")
    put(p2, 430, 124, "DATE OF REHIRE:")
    put(p2, 620, 124, "AFTER-TAX DEFERRAL % OR $:")
    # a fund name below the block, which must not start a second record
    put(p2, 40, 160, "FUND NAME: TIAA TRADITIONAL")

    # --- page 3: table continuation, no header printed -----------------------
    p3 = doc.new_page(width=792, height=612)
    put(p3, 40, 60, "Small Cash Out Monitoring Report (continued)")
    put(p3, 40, 130, "456-78-9012")
    put(p3, 110, 130, "O'LEARY, MARY JANE")
    put(p3, 250, 130, "0004")
    put(p3, 305, 130, "Retired")
    put(p3, 358, 130, "07/04/1955")
    put(p3, 415, 130, "09/09/1979")
    put(p3, 582, 130, "12/31/2014")

    # --- page 4: table carrying BOTH a birth and a death column --------------
    # The two captions sit side by side, so this is where a date-of-death
    # would be misread as a date of birth if the columns were not anchored.
    p4 = doc.new_page(width=792, height=612)
    put(p4, 40, 60, "Small Cash Out Monitoring Report -- with date of death")
    put(p4, 305, 100, "DATE OF")
    put(p4, 380, 100, "DATE OF")
    put(p4, 455, 100, "DATE OF")
    for x, caption in [
        (40, "PART ID"), (110, "NAME"), (250, "STATUS"),
        (310, "BIRTH"), (385, "DEATH"), (450, "TERMINATION"),
    ]:
        put(p4, x, 110, caption)
    # a participant with a date of death, and one without
    put(p4, 40, 130, "567-89-0123")
    put(p4, 110, 130, "DENNING, JAMES DARYL")
    put(p4, 250, 130, "Deceased")
    put(p4, 303, 130, "02/09/1962")
    put(p4, 378, 130, "11/30/2019")
    put(p4, 452, 130, "11/30/2019")
    put(p4, 40, 144, "678-90-1234")
    put(p4, 110, 144, "DAVIS, CATHRYN R")
    put(p4, 250, 144, "Active")
    put(p4, 303, 144, "08/21/1960")
    # a participant whose PART ID cell is empty: must still be written out,
    # flagged, rather than dropped
    put(p4, 110, 158, "BAYLOCK, CINDY KING")
    put(p4, 250, 158, "Active")
    put(p4, 303, 158, "04/05/1959")

    # --- page 5: label layout with captions OMITTED, not blank ---------------
    # This is the case that mis-assigned names. The report prints no caption
    # at all when a value is empty, so block 2 has no "NAME:" line and block 3
    # has no "PART ID:" line. Block 3's name must NOT land on block 2's
    # Part ID.
    p5 = doc.new_page(width=792, height=612)
    put(p5, 40, 60, "Participant Contribution Investment Allocation Report")

    put(p5, 40, 100, "NAME: CLARK, CHRISTOPHER R")
    put(p5, 300, 100, "DATE OF BIRTH: 01/02/1970")
    put(p5, 40, 112, "PART ID: 111223333")
    put(p5, 300, 112, "DATE OF HIRE: 01/01/2000")

    # no NAME caption anywhere in this block
    put(p5, 40, 170, "PART ID: 222334444")
    put(p5, 300, 170, "DATE OF BIRTH: 03/04/1971")
    put(p5, 40, 182, "STATUS: Active")
    put(p5, 300, 182, "DATE OF HIRE: 02/02/2001")

    # no PART ID caption anywhere in this block
    put(p5, 40, 240, "NAME: CLIFTON, BERA S")
    put(p5, 300, 240, "DATE OF BIRTH: 05/06/1972")
    put(p5, 40, 252, "STATUS: Active")
    put(p5, 300, 252, "DATE OF HIRE: 03/03/2002")

    doc.save(str(dest))
    doc.close()
    return dest


def selftest() -> int:
    import tempfile

    if fitz is None:
        print("SKIP  PyMuPDF is not installed, cannot build the test PDF.")
        return 0

    failures = []

    def check(label, got, want):
        if got != want:
            failures.append(f"{label}: got {got!r}, want {want!r}")

    # pure functions first
    check("norm_ssn dashed", norm_ssn("123-45-6789"), "123-45-6789")
    check("norm_ssn running", norm_ssn("123456789"), "123-45-6789")
    check("norm_ssn unicode dash", norm_ssn("123‐45‐6789"), "123-45-6789")
    check("norm_ssn masked", norm_ssn("XXX-XX-6789"), "XXX-XX-6789")
    check("norm_date 4-digit", norm_date("01/15/1963")[0], "01/15/1963")
    check("norm_date 2-digit", norm_date("1/5/63")[0], "01/05/1963")
    check("mask", mask_ssn("123-45-6789"), "***-**-6789")
    check("split comma", split_name("ACKERMAN, KENTON CLYDE")[:4],
          ("ACKERMAN", "KENTON", "CLYDE", ""))
    check("split no space", split_name("CHILDERS,MORGAN N")[:4],
          ("CHILDERS", "MORGAN", "N", ""))
    check("split suffix", split_name("SMITH JR, JOHN A")[:4],
          ("SMITH", "JOHN", "A", "JR"))

    with tempfile.TemporaryDirectory() as tmp:
        pdf = _build_test_pdf(Path(tmp) / "synthetic.pdf")
        for engine in available_engines():
            rows, diag = process_pdf(pdf, engine)
            tag = f"[{engine}]"
            by_id = {r["Part ID (SSN)"]: r for r in rows}

            check(f"{tag} row count", len(rows), 11)
            check(f"{tag} error", diag["Error"], "")

            r = by_id.get("123-45-6789", {})
            check(f"{tag} table name", r.get("Name As Printed"),
                  "ACKERMAN, KENTON CLYDE")
            check(f"{tag} table dob", r.get("Date Of Birth"), "01/15/1963")

            # the row with no birth date must NOT pick up the hire date
            r = by_id.get("234-56-7890", {})
            check(f"{tag} blank dob", r.get("Date Of Birth"), "")
            check(f"{tag} blank dob name", r.get("Name As Printed"), "CHILDERS,MORGAN N")

            r = by_id.get("345-67-8901", {})
            check(f"{tag} running ssn dob", r.get("Date Of Birth"), "12/25/1977")
            check(f"{tag} running ssn suffix", r.get("Suffix"), "JR")

            r = by_id.get("987-65-4321", {})
            check(f"{tag} label layout", r.get("Layout"), "label")
            check(f"{tag} label name", r.get("Name As Printed"), "ABBEY, BORIS S")
            check(f"{tag} label dob", r.get("Date Of Birth"), "04/12/1973")

            check(f"{tag} label death", r.get("Date Of Death"), "03/22/2024")

            r = by_id.get("456-78-9012", {})
            check(f"{tag} inherited dob", r.get("Date Of Birth"), "07/04/1955")
            check(f"{tag} inherited layout", r.get("Layout"), "table (inherited)")

            # birth and death columns side by side must not be confused
            r = by_id.get("567-89-0123", {})
            check(f"{tag} table dob beside death", r.get("Date Of Birth"), "02/09/1962")
            check(f"{tag} table death", r.get("Date Of Death"), "11/30/2019")

            r = by_id.get("678-90-1234", {})
            check(f"{tag} dob with no death", r.get("Date Of Birth"), "08/21/1960")
            check(f"{tag} blank death", r.get("Date Of Death"), "")

            # a table row whose PART ID cell was empty: kept, and flagged
            by_name = {r["Name As Printed"]: r for r in rows}
            r = by_name.get("BAYLOCK, CINDY KING", {})
            check(f"{tag} blank part id kept", r.get("Date Of Birth"), "04/05/1959")
            check(f"{tag} blank part id blank", r.get("Part ID (SSN)"), "")
            check(f"{tag} blank part id flagged",
                  NO_PART_ID_NOTE in r.get("Notes", ""), True)

            # --- the omitted-caption case that mis-assigned names -----------
            r = by_id.get("111-22-3333", {})
            check(f"{tag} block 1 name", r.get("Name As Printed"),
                  "CLARK, CHRISTOPHER R")
            check(f"{tag} block 1 dob", r.get("Date Of Birth"), "01/02/1970")

            # block 2 printed no NAME caption at all. Its name must be blank,
            # NOT block 3's name.
            r = by_id.get("222-33-4444", {})
            check(f"{tag} block 2 name stays blank", r.get("Name As Printed"), "")
            check(f"{tag} block 2 dob", r.get("Date Of Birth"), "03/04/1971")

            # block 3 printed no PART ID caption. Its name must survive with a
            # blank Part ID, not be swallowed by block 2.
            r = by_name.get("CLIFTON, BERA S", {})
            check(f"{tag} block 3 kept", r.get("Date Of Birth"), "05/06/1972")
            check(f"{tag} block 3 part id blank", r.get("Part ID (SSN)"), "")
            check(f"{tag} block 3 flagged",
                  NO_PART_ID_NOTE in r.get("Notes", ""), True)

    if failures:
        print("SELFTEST FAILED")
        for f in failures:
            print("  " + f)
        return 1
    print(f"SELFTEST PASSED  (engines: {', '.join(available_engines())})")
    return 0


# ---------------------------------------------------------------------------
# headless run
# ---------------------------------------------------------------------------

def find_pdfs(folder):
    return sorted(Path(folder).rglob("*.pdf"))


def recon_summary(d) -> str:
    """One line per flagged file. No participant data, counts only."""
    return (
        f"{d['File Name']}: ids={d['Part IDs On Page']} rows={d['Rows Written']} "
        f"no-part-id={d['Rows Without Part ID']} "
        f"no-name={d['Rows Missing Name']} "
        f"no-dob={d['Rows Missing DOB']} "
        f"dob-suspect={d['Rows With Suspicious DOB']} {d['Error']}".rstrip()
    )


def run_headless(src, out, engine=DEFAULT_ENGINE, workers=None, masked=False) -> int:
    pdfs = find_pdfs(src)
    if not pdfs:
        print(f"No PDFs found under {src}")
        return 1

    def progress(done, total, name):
        print(f"  [{done}/{total}] {name}")

    print(f"Reading {len(pdfs)} PDF(s) with {engine} ...")
    rows, diags = process_folder(pdfs, engine, workers, progress)
    dest = write_workbook(rows, diags, Path(out), masked=masked)

    flagged = [d for d in diags if d["Complete"] != "YES"]
    print(f"\n{len(rows)} row(s) written to {dest}")
    if masked:
        print("Part IDs are MASKED in this workbook.")
    if flagged:
        print(f"{len(flagged)} file(s) need checking on the Reconciliation sheet:")
        for d in flagged:
            print("  " + recon_summary(d))
    else:
        print("Reconciliation: every file complete.")
    print("\nThis workbook holds names, SSNs and dates of birth. Save it to the "
          "appropriate Global Insider folder with restricted access -- not to a "
          "desktop or personal drive.")
    return 0


# ---------------------------------------------------------------------------
# Tkinter front end
# ---------------------------------------------------------------------------

class App:
    def __init__(self, root):
        import tkinter as tk
        from tkinter import ttk

        self.tk = tk
        self.root = root
        root.title("TIAA participant Part ID / Name / DOB extractor")
        root.geometry("860x560")

        frame = ttk.Frame(root, padding=12)
        frame.pack(fill="both", expand=True)

        self.src = tk.StringVar()
        self.out = tk.StringVar()
        self.engine = tk.StringVar(value=DEFAULT_ENGINE)
        self.masked = tk.BooleanVar(value=False)

        ttk.Label(frame, text="Folder of PDFs").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.src, width=78).grid(row=0, column=1, padx=6)
        ttk.Button(frame, text="Browse", command=self.pick_src).grid(row=0, column=2)

        ttk.Label(frame, text="Output workbook").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(frame, textvariable=self.out, width=78).grid(row=1, column=1, padx=6)
        ttk.Button(frame, text="Save as", command=self.pick_out).grid(row=1, column=2)

        options = ttk.Frame(frame)
        options.grid(row=2, column=1, sticky="w", pady=4)
        ttk.Label(options, text="Engine").pack(side="left")
        ttk.Combobox(
            options, textvariable=self.engine, width=12, state="readonly",
            values=available_engines() or [DEFAULT_ENGINE],
        ).pack(side="left", padx=6)
        ttk.Checkbutton(
            options, text="Mask Part IDs in the workbook (***-**-6789)",
            variable=self.masked,
        ).pack(side="left", padx=16)

        self.run_button = ttk.Button(frame, text="Extract", command=self.run)
        self.run_button.grid(row=3, column=1, sticky="w", pady=8)

        self.progress = ttk.Progressbar(frame, length=640, mode="determinate")
        self.progress.grid(row=4, column=0, columnspan=3, sticky="we", pady=4)

        self.status = ttk.Label(frame, text="Ready.")
        self.status.grid(row=5, column=0, columnspan=3, sticky="w")

        from tkinter import scrolledtext
        self.log_box = scrolledtext.ScrolledText(frame, height=20, width=104)
        self.log_box.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=8)
        frame.rowconfigure(6, weight=1)
        frame.columnconfigure(1, weight=1)

        self.log("Output contains names, SSNs and dates of birth. Save it to a "
                 "restricted Global Insider folder.")
        self.log("No SSN is ever written to this log.")

    def log(self, message):
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.root.update_idletasks()

    def pick_src(self):
        from tkinter import filedialog
        path = filedialog.askdirectory(title="Select the folder holding the PDFs")
        if path:
            self.src.set(path)
            if not self.out.get():
                self.out.set(str(Path(path) / OUTPUT_XLSX_NAME))

    def pick_out(self):
        from tkinter import filedialog
        path = filedialog.asksaveasfilename(
            title="Save the workbook as", defaultextension=".xlsx",
            initialfile=OUTPUT_XLSX_NAME,
            filetypes=[("Excel workbook", "*.xlsx")],
        )
        if path:
            self.out.set(path)

    def run(self):
        from tkinter import messagebox
        src, out = self.src.get().strip(), self.out.get().strip()
        if not src or not Path(src).is_dir():
            messagebox.showerror("Folder", "Pick a folder of PDFs first.")
            return
        if not out:
            out = str(Path(src) / OUTPUT_XLSX_NAME)
            self.out.set(out)

        pdfs = find_pdfs(src)
        if not pdfs:
            messagebox.showerror("Folder", f"No PDFs found under {src}")
            return

        self.run_button.state(["disabled"])
        self.progress["maximum"] = len(pdfs)
        self.progress["value"] = 0
        self.log(f"\nReading {len(pdfs)} PDF(s) with {self.engine.get()} ...")

        def progress(done, total, name):
            self.progress["value"] = done
            self.status.config(text=f"[{done}/{total}] {name}")
            self.root.update_idletasks()

        try:
            rows, diags = process_folder(pdfs, self.engine.get(), None, progress)
            dest = write_workbook(rows, diags, Path(out), masked=self.masked.get())
        except Exception as exc:
            self.log(f"FAILED  {type(exc).__name__}: {exc}")
            self.log(traceback.format_exc())
            messagebox.showerror("Extraction failed", str(exc))
            self.run_button.state(["!disabled"])
            return

        flagged = [d for d in diags if d["Complete"] != "YES"]
        self.log(f"{len(rows)} row(s) written to {dest}")
        if self.masked.get():
            self.log("Part IDs are MASKED in this workbook.")
        if flagged:
            self.log(f"{len(flagged)} file(s) need checking on the "
                     f"Reconciliation sheet:")
            for d in flagged:
                self.log("  " + recon_summary(d))
        else:
            self.log("Reconciliation: every file complete.")
        self.status.config(text="Done.")
        self.run_button.state(["!disabled"])
        messagebox.showinfo("Done", f"{len(rows)} row(s) written.\n\n{dest}")


def launch_gui() -> int:
    try:
        import tkinter as tk
    except ImportError:
        print("Tkinter is unavailable. Use --src and --out for a headless run.")
        return 1
    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


# ---------------------------------------------------------------------------

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract Part ID, Name and Date of Birth from TIAA "
                    "participant PDFs.",
    )
    parser.add_argument("--src", help="folder of PDFs (headless run)")
    parser.add_argument("--out", help="output .xlsx path")
    parser.add_argument("--engine", choices=["pdfplumber", "mupdf"],
                        default=DEFAULT_ENGINE)
    parser.add_argument("--workers", type=int, default=None)
    parser.add_argument("--mask", action="store_true",
                        help="write masked Part IDs (***-**-6789)")
    parser.add_argument("--selftest", action="store_true")
    parser.add_argument("--qa", metavar="PDF", help="masked dump of one file")
    parser.add_argument("--lines", metavar="PDF",
                        help="masked geometry dump: rebuilt lines, x spans "
                             "and column assignment (shapes only, no PII)")
    parser.add_argument("--page", type=int, default=None,
                        help="restrict --lines to one page")
    parser.add_argument("--verify", metavar="FOLDER",
                        help="compare both engines over a folder")
    args = parser.parse_args()

    if not available_engines():
        print("Install pdfplumber or PyMuPDF first:  pip install pdfplumber")
        return 1

    if args.selftest:
        return selftest()
    if args.qa:
        qa_file(args.qa, args.engine)
        return 0
    if args.lines:
        dump_lines(args.lines, args.engine, args.page)
        return 0
    if args.verify:
        return verify(args.verify)
    if args.src:
        out = args.out or str(Path(args.src) / OUTPUT_XLSX_NAME)
        return run_headless(args.src, out, args.engine, args.workers, args.mask)
    return launch_gui()


if __name__ == "__main__":
    sys.exit(main())
