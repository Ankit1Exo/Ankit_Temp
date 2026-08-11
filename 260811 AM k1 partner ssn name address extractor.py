r"""
Schedule K-1 (Form 1065)  --  Partner SSN/TIN, Name and Address extractor
=========================================================================
Reads searchable Schedule K-1 (Form 1065) PDFs and writes one CSV row per
partner: Name, First/MI/Last, Address, TIN, TIN Type, Entity Type.

WHAT IS READ, AND FROM WHERE
    Part II "Information About the Partner":

        E  Partner's SSN or TIN (Do not use TIN of a disregarded entity. ...)
           111-11-1111                                   -> TIN, TIN Type
        F  Name, address, city, state, and ZIP code for partner entered in E.
           FIRST M LAST                                  -> Name / First / MI / Last
           122 SAMPLE DR                                 -> Address
           SAMPLE TOWN            PA 19444               -> Address
        I1 What type of entity is this partner?  INDIVIDUAL  -> Entity Type

    Part I box A ("Partnership's employer identification number") is also
    read, but only as a de-duplication discriminator -- see WHY THE
    PARTNERSHIP EIN IS CARRIED below. Drop it from CSV_COLUMNS if unwanted.

THE BOX LETTER IS NEVER USED AS AN ANCHOR
    On these documents the letter printed before each heading ("E", "F",
    "H1", "I1") is not a structured element of its own -- selecting it in a
    viewer also selects unrelated text elsewhere on the page, because the
    text layer's blocks/lines span both printed columns. So the letter
    carries no reliable position or grouping, and it can also differ between
    form years (older 1065 K-1s label the identifying number differently).

    Every anchor here is therefore the HEADING TEXT, matched anywhere in a
    line rather than at its start, and lines are rebuilt from word boxes by
    this script (group_words_into_lines) instead of trusting the text
    layer's own line/block structure.

WHY THE PART II COLUMN HAS TO BE ISOLATED BY GEOMETRY
    A K-1 prints Part II (partner identity) in the left column and Part III
    (boxes 1-20, income amounts) in the right column, at overlapping
    vertical positions. Any line rebuilt purely by y therefore mixes the
    two, and the partner's address silently acquires words from the income
    boxes:

        printed (two columns)                 read as one line
        122 SAMPLE DR        7  Royalties     122 SAMPLE DR 7 Royalties
        SAMPLE TOWN  PA 19444   8  Net ...    SAMPLE TOWN PA 19444 8 Net ...

    detect_column_gutter() finds the boundary from word geometry alone: it
    marks which horizontal bands of the page contain any word at all and
    takes the widest empty band in the middle of the page. That is the
    printed gutter between the two columns, and it does not depend on the
    text layer's grouping, on the box letters, or on a hard-coded fraction.

    Two fallbacks follow, in order, each logged so the run is auditable:
      1. the right-hand extent of the Part II headings themselves, cut at
         the first wide horizontal gap in the heading's own line (that gap
         IS the gutter, on the one line where a heading is present);
      2. --column-fallback-fraction (default 0.50 of page width).
    --column-right-fraction overrides all three for a stubborn layout.

ONE RECORD PER "Partner's SSN or TIN" HEADING
    Records are split on the heading's own vertical position, so a page
    holding one K-1 and a page holding several are handled the same way and
    no per-page layout is assumed. A page with no Part II heading at all
    (K-1 page 2, the codes/instructions page) produces no rows.

TIN: SSN AND EIN ARE BOTH KEPT, AND LABELLED
    Box E holds a 3-2-4 SSN for an individual partner but a 2-7 EIN when the
    partner is a trust, LLC or corporation. Both are extracted and TIN Type
    records which was found, so nothing is silently dropped:

        SSN                  111-11-1111
        EIN                  11-1111111
        SSN (masked)         XXX-XX-1111   client copies often print these
        EIN (masked)         XX-XXX1111
        MASKED (no digits)   TIN blank, nothing recoverable was printed
        ... (no separator)   9 digits run together; 3-2-4 vs 2-7 is then
                             ambiguous, so box I1's entity type decides and
                             the guess is logged
        ... (OCR corrected)  a digit was recovered from an OCR look-alike

WHY THE FIRST/MI/LAST SPLIT IS BEST-EFFORT, AND WHEN IT IS SKIPPED
    Unlike W-2 box e, K-1 box F has no "Last name" sub-heading to give a
    column position to split on -- the whole name is one printed line. So:
      - "LAST, FIRST M" (comma form) splits exactly;
      - otherwise a trailing single letter is taken as the middle initial
        and the last token as the surname, which mis-splits a multi-word
        surname ("VEGA TAPIA", "DE LA CRUZ") and is logged as a guess;
      - an entity partner ("... LLC", "... TRUST", or box I1 saying
        anything other than INDIVIDUAL) is NOT split at all -- First/MI/Last
        are left blank rather than filled with pieces of a company name;
      - a joint name ("A DOE & B DOE") is likewise left unsplit.
    The verbatim box F name is always in the Name column, so a reviewer can
    correct any split without reopening the PDF.

WHY THE PARTNERSHIP EIN IS CARRIED
    One PDF can hold K-1s from several partnerships for the same partner.
    De-duplicating on TIN alone would collapse those into one row and lose a
    real K-1, so the de-duplication key is Partnership EIN + TIN (falling
    back to Partnership EIN + Name + Address when no TIN was found). Copies
    of the SAME K-1 (client copy, file copy) still collapse to one row.

OCR TOLERANCES
    These PDFs are usually made searchable by ABBYY rather than carrying
    native text, so: normalize_text() folds curly quotes/dashes and stray
    whitespace; heading patterns allow a missing apostrophe ("Partners SSN");
    and if no strictly numeric TIN is found next to the heading, digits are
    recovered from common OCR look-alikes (O/0, I/1, S/5 ...) within that
    small window only -- never page-wide, so unrelated text cannot become a
    false TIN.

HANDLING NOTE
    The output CSV contains taxpayer identifiers in clear text. Write it to
    an access-controlled location (the appropriate Global Insider folder),
    not to a desktop or a source-control working tree.

Usage:
    Run it with no arguments -- from IDLE with F5, or by double-clicking -- and
    it asks for the PDFs and for where to save the CSV.

    Or from a command line:
      python "260811 AM k1 partner ssn name address extractor.py" <pdf_or_folder> [-o out.csv]
      python "260811 AM k1 partner ssn name address extractor.py" <pdf> --debug [--debug-page N]

Requires:
    pip install pymupdf tqdm
"""

import re
import csv
import argparse
import unicodedata
from pathlib import Path

import pymupdf as fitz
from tqdm import tqdm

MIN_TEXT_CHARS_PER_PAGE = 20

# --- column isolation tuning (all fractions of page width) -------------------
# Bin width for the gutter scan. 1pt on a 612pt page: fine enough to resolve a
# printed gutter, coarse enough that a hairline gap inside a word run doesn't
# register as one.
GUTTER_BIN_POINTS = 1.0
# Only look for the gutter in the middle of the page -- a K-1's two columns
# meet near the centre, and the page margins are empty bands too (they would
# otherwise win the "widest empty band" contest).
GUTTER_SEARCH_LO_FRACTION = 0.30
GUTTER_SEARCH_HI_FRACTION = 0.70
# An empty band narrower than this is word spacing, not a column gutter.
GUTTER_MIN_WIDTH_FRACTION = 0.010
# Fallback 1: pad added to the right-hand extent of the Part II headings.
COLUMN_PAD_FRACTION = 0.04
# Fallback 1: a horizontal gap this wide inside a heading's own line is the
# gutter, so the heading text ends there. Normal inter-word spacing is far
# smaller; the gap from a left-column heading to right-column text is far
# larger.
COLUMN_GAP_FRACTION = 0.03
# Fallback 2: used when neither the gutter nor a heading extent is available.
COLUMN_FALLBACK_FRACTION = 0.50
# A computed boundary left of this is implausibly narrow (it would cut the
# partner's own address off) -- reject it and fall back.
MIN_COLUMN_FRACTION = 0.25

CSV_COLUMNS = ["Document ID", "Page", "Name", "First Name", "MI", "Last Name",
               "Address", "TIN", "TIN Type", "Entity Type", "Partnership EIN"]

OUTPUT_CSV_NAME = "k1_partner_extracted.csv"

# Set when the script was started with no arguments (IDLE F5, or a
# double-click). The console window is then kept open at the end so the summary
# and any warnings can actually be read before it closes.
LAUNCHED_INTERACTIVELY = False

# ".?" (not ".") for the apostrophe so a dropped or misread apostrophe -- common
# when ABBYY OCRs a scanned page -- still matches ("Partners SSN or TIN").
# Matched with search(), never anchored to the start of a line, because the
# box letter before the heading is not reliably part of the same line.
TIN_CAPTION_RE = re.compile(
    r"Partner.?s\s+(?:SSN\s*(?:or|/)\s*TIN|identifying\s+number)", re.IGNORECASE)

# Current wording is "Name, address, city, state, and ZIP code for partner
# entered in E."; older form years print "Partner's name, address, city,
# state, and ZIP code". Both are accepted.
NAME_CAPTION_RE = re.compile(
    r"Name,?\s*address,?\s*city,?\s*(?:state,?\s*)?and\s*ZIP\s*code\s*for\s*partner"
    r"|Partner.?s\s+name,?\s*address,?\s*city", re.IGNORECASE)

# Part I box B is "Partnership's name, address, city, state, and ZIP code" --
# the partnership's own address, which must never be recorded as the
# partner's. Any line matching this is excluded from NAME_CAPTION_RE hits.
PARTNERSHIP_NAME_CAPTION_RE = re.compile(r"Partnership.?s\s+name,?\s*address", re.IGNORECASE)

PARTNERSHIP_EIN_CAPTION_RE = re.compile(
    r"Partnership.?s\s+employer\s+identification\s+number", re.IGNORECASE)

# "(?!ship)" so Part I's "Information About the Partnership" heading is not
# mistaken for Part II's "Information About the Partner".
PART_II_HEADING_RE = re.compile(r"Information\s+About\s+the\s+Partner(?!ship)", re.IGNORECASE)

ENTITY_TYPE_CAPTION_RE = re.compile(r"What\s+type\s+of\s+entity\s+is\s+this\s+partner", re.IGNORECASE)

# Headings printed after box F, which end the name/address block. Matched with
# search() (not anchored) because the box letter may or may not share the line.
# NOTE: deliberately never applied to the FIRST line of the block -- that line
# is the partner's name, and a partner really can be called
# "XYZ GENERAL PARTNER LLC", which would otherwise terminate the block before
# any address was collected.
STOP_LABEL_RE = re.compile(
    r"General\s+partner|Limited\s+partner|member-manager|other\s+LLC\s+member"
    r"|Domestic\s+partner|Foreign\s+partner|disregarded\s+entity"
    r"|What\s+type\s+of\s+entity|retirement\s+plan"
    r"|Partner.?s\s+share\s+of|Partner.?s\s+capital\s+account"
    r"|Nonrecourse|Recourse\s+financing|^Part\s+I{1,3}\b|^Beginning\b|^Ending\b",
    re.IGNORECASE)

# --- TIN value shapes -------------------------------------------------------
# A separator is REQUIRED here so that 3-2-4 (SSN) and 2-7 (EIN) stay
# distinguishable; the separator-less 9-digit case is handled separately and
# labelled, because 9 bare digits are genuinely ambiguous between the two.
SSN_SEP_RE = re.compile(r"\b(\d{3})[-\s](\d{2})[-\s](\d{4})\b")
EIN_SEP_RE = re.compile(r"\b(\d{2})[-\s](\d{7})\b")
NINE_DIGIT_RE = re.compile(r"\b(\d{9})\b")
# Tax software client copies commonly print the TIN partly masked.
MASKED_SSN_RE = re.compile(r"[X*]{3}[-\s]?[X*]{2}[-\s]?(\d{4})\b", re.IGNORECASE)
MASKED_EIN_RE = re.compile(r"[X*]{2}[-\s]?[X*]{3}[-\s]?(\d{4})\b", re.IGNORECASE)
FULLY_MASKED_RE = re.compile(r"[X*]{2,3}[-\s]?[X*]{2,3}[-\s]?[X*]{4}", re.IGNORECASE)

CITY_STATE_ZIP_RE = re.compile(r"^(?P<city>.+?),?\s+(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)\b")

# Letters ABBYY commonly confuses with digits on a noisy scan. Applied only
# inside the few lines beside an already-found TIN heading -- never page-wide --
# so it cannot turn unrelated text into a false TIN.
OCR_DIGIT_FIX = str.maketrans({
    "O": "0", "o": "0", "Q": "0", "D": "0",
    "I": "1", "l": "1", "i": "1", "|": "1",
    "Z": "2", "z": "2",
    "S": "5", "s": "5",
    "G": "6", "b": "6",
    "T": "7",
    "B": "8",
    "g": "9", "q": "9",
})

# Tokens that mark box F's name line as an organisation rather than a person,
# in which case no First/MI/Last split is attempted.
ENTITY_HINT_RE = re.compile(
    r"\b(?:LLC|L\.L\.C|INC|INCORPORATED|CORP|CORPORATION|CO|COMPANY|LP|L\.P|LLP|LLLP"
    r"|PLLC|PARTNERSHIP|PARTNERS|TRUST|TRUSTEE|ESTATE|FOUNDATION|FUND|HOLDINGS"
    r"|VENTURE|VENTURES|ASSOCIATES|GROUP|PROPERTIES|ENTERPRISES|BANK|IRA|SEP|KEOGH"
    r"|PENSION|CUSTODIAN|LTD|PC|PA)\b\.?", re.IGNORECASE)

JOINT_NAME_RE = re.compile(r"&|\bAND\b", re.IGNORECASE)

# Marks a box F line as the start of the street address rather than a
# continuation of the name, so a name printed over two lines is not mistaken
# for the first line of the address.
STREET_HINT_RE = re.compile(
    r"^\d|\bP\.?\s*O\.?\s*BOX\b"
    r"|\b(?:ST|STREET|AVE|AVENUE|RD|ROAD|DR|DRIVE|LN|LANE|BLVD|BOULEVARD|CT|COURT"
    r"|WAY|PL|PLACE|TER|TERRACE|CIR|CIRCLE|HWY|HIGHWAY|PKWY|PARKWAY|PIKE|TRL|TRAIL"
    r"|SUITE|STE|APT|UNIT|FLOOR|FL|RM|ROOM|BOX|PMB)\b\.?", re.IGNORECASE)

# How far below the TIN heading its value may sit. The window is additionally
# cut short at box F's heading, so a blank box E cannot reach down and pick up
# the partner's ZIP code or street number instead.
TIN_WINDOW_LINES = 4
# How far below box F's heading the name/address block may extend.
ADDRESS_BLOCK_MAX_LINES = 9


def mask_shape(s):
    """Reduce a line to its digit/letter shape so a layout can be shown for
    troubleshooting without exposing the values themselves."""
    return re.sub(r"[A-Za-z]", "X", re.sub(r"\d", "#", s))


def normalize_text(s):
    """Fold ABBYY's curly quotes/dashes and odd whitespace to plain ASCII
    equivalents before any regex runs against OCR'd text."""
    s = unicodedata.normalize("NFKC", s)
    s = s.translate({0x2018: "'", 0x2019: "'", 0x00B4: "'", 0x0060: "'",
                     0x201C: '"', 0x201D: '"', 0x00A0: " ",
                     0x2013: "-", 0x2014: "-"})
    return re.sub(r"[ \t]+", " ", s).strip()


def group_words_into_lines(words, y_tol=3):
    """Rebuild printed lines from word boxes, clustering by top edge.

    Done here rather than taken from the PDF's own line/block structure
    because on these documents that structure spans both printed columns --
    which is the same reason the box letter before a heading cannot be
    selected on its own. Each returned line is a list of (x0, x1, text)
    sorted left to right.
    """
    items = sorted(((w[1], w[0], w[2], w[4]) for w in words), key=lambda t: (t[0], t[1]))
    lines, current, current_y = [], [], None
    for y0, x0, x1, text in items:
        if current_y is not None and abs(y0 - current_y) > y_tol:
            lines.append(sorted(current, key=lambda t: t[0]))
            current, current_y = [], None
        if current_y is None:
            current_y = y0
        current.append((x0, x1, text))
    if current:
        lines.append(sorted(current, key=lambda t: t[0]))
    return lines


def to_plain_lines(lines):
    return [normalize_text(" ".join(t for _, _, t in ln)) for ln in lines]


def detect_column_gutter(words, page_width):
    """Find the x boundary between Part II (left) and Part III (right) from
    word geometry alone: mark every horizontal bin that contains any word, then
    take the widest unmarked run in the middle of the page. That run is the
    printed gutter, and its midpoint separates the columns.

    Returns None when no wide-enough empty band exists there -- e.g. a page
    that isn't the two-column K-1 grid, or one whose address text runs right up
    to the income boxes.
    """
    if page_width <= 0:
        return None
    bins = int(page_width / GUTTER_BIN_POINTS) + 1
    occupied = bytearray(bins)
    for w in words:
        if not str(w[4]).strip():
            continue
        lo = max(0, int(w[0] / GUTTER_BIN_POINTS))
        hi = min(bins - 1, int(w[2] / GUTTER_BIN_POINTS))
        for b in range(lo, hi + 1):
            occupied[b] = 1

    search_lo = int(page_width * GUTTER_SEARCH_LO_FRACTION / GUTTER_BIN_POINTS)
    search_hi = min(bins, int(page_width * GUTTER_SEARCH_HI_FRACTION / GUTTER_BIN_POINTS))
    min_width = page_width * GUTTER_MIN_WIDTH_FRACTION

    best_lo = best_hi = None
    run_lo = None
    for b in range(search_lo, search_hi):
        if not occupied[b]:
            if run_lo is None:
                run_lo = b
        else:
            if run_lo is not None:
                if best_lo is None or (b - run_lo) > (best_hi - best_lo):
                    best_lo, best_hi = run_lo, b
                run_lo = None
    if run_lo is not None and (best_lo is None or (search_hi - run_lo) > (best_hi - best_lo)):
        best_lo, best_hi = run_lo, search_hi

    if best_lo is None:
        return None
    width = (best_hi - best_lo) * GUTTER_BIN_POINTS
    if width < min_width:
        return None
    return (best_lo + best_hi) / 2 * GUTTER_BIN_POINTS


def caption_extent(line_words, page_width):
    """Right-hand x of a heading's own text within its rebuilt line, cutting at
    the first horizontal gap wide enough to be the column gutter.

    Needed because a heading's line, rebuilt by vertical position, can also
    contain right-column income-box text sitting at the same height -- so the
    line's own right edge is not the heading's right edge.
    """
    if not line_words:
        return None
    gap = page_width * COLUMN_GAP_FRACTION
    x1 = line_words[0][1]
    for prev, cur in zip(line_words, line_words[1:]):
        if cur[0] - prev[1] > gap:
            break
        x1 = cur[1]
    return x1


def is_part_ii_anchor(text):
    if PARTNERSHIP_NAME_CAPTION_RE.search(text):
        return False
    return bool(PART_II_HEADING_RE.search(text)
                or TIN_CAPTION_RE.search(text)
                or NAME_CAPTION_RE.search(text))


def compute_left_column_bound(lines, plain_lines, page_width, words, override, page_num):
    """Decide where the Part II column ends. Returns (bound, how) or
    (None, reason) when the page carries no Part II heading to work from."""
    anchor_extents = [caption_extent(lw, page_width)
                      for lw, text in zip(lines, plain_lines) if is_part_ii_anchor(text)]
    anchor_extents = [x for x in anchor_extents if x is not None]
    if not anchor_extents:
        return None, "no Part II heading found on this page"

    if override is not None:
        return page_width * override, f"--column-right-fraction {override}"

    gutter = detect_column_gutter(words, page_width)
    if gutter is not None and gutter > max(anchor_extents):
        return gutter, "column gutter detected from word geometry"
    if gutter is not None:
        tqdm.write(f"  page {page_num}: gutter found at x={gutter:.0f} but a Part II heading extends past it "
                   f"(x={max(anchor_extents):.0f}) -- using the heading extent instead")

    bound = max(anchor_extents) + page_width * COLUMN_PAD_FRACTION
    if bound < page_width * MIN_COLUMN_FRACTION:
        tqdm.write(f"  page {page_num}: Part II heading extent gives an implausibly narrow column "
                   f"(x={bound:.0f}, under {MIN_COLUMN_FRACTION:.0%} of page width) -- falling back to "
                   f"{COLUMN_FALLBACK_FRACTION:.0%} of page width; check with --debug and set "
                   f"--column-right-fraction if the address looks truncated")
        return page_width * COLUMN_FALLBACK_FRACTION, "fallback fraction (heading extent too narrow)"
    return bound, "Part II heading extent"


def normalize_ssn(match):
    return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"


def normalize_ein(match):
    return f"{match.group(1)}-{match.group(2)}"


def extract_tin(window_text, entity_type, page_num, label):
    """Read box E's value out of the few lines beside its heading and say what
    kind of identifier it is. Order matters: separator-bearing shapes are tried
    first, because they are the only ones where 3-2-4 (SSN) and 2-7 (EIN) can be
    told apart with certainty."""
    m = SSN_SEP_RE.search(window_text)
    if m:
        return normalize_ssn(m), "SSN"
    m = EIN_SEP_RE.search(window_text)
    if m:
        return normalize_ein(m), "EIN"

    m = MASKED_SSN_RE.search(window_text)
    if m:
        tqdm.write(f"  page {page_num} ({label}): box E holds a masked SSN -- only the last 4 digits were "
                   f"printed on this copy")
        return f"XXX-XX-{m.group(1)}", "SSN (masked)"
    m = MASKED_EIN_RE.search(window_text)
    if m:
        tqdm.write(f"  page {page_num} ({label}): box E holds a masked EIN -- only the last 4 digits were "
                   f"printed on this copy")
        return f"XX-XXX{m.group(1)}", "EIN (masked)"
    if FULLY_MASKED_RE.search(window_text):
        tqdm.write(f"  page {page_num} ({label}): box E is fully masked -- no digits were printed, TIN left blank")
        return "", "MASKED (no digits)"

    m = NINE_DIGIT_RE.search(window_text)
    if m:
        digits = m.group(1)
        et = entity_type.upper()
        if "INDIVIDUAL" in et:
            tqdm.write(f"  page {page_num} ({label}): box E printed 9 digits with no separator; box I1 says "
                       f"INDIVIDUAL, so read as an SSN")
            return f"{digits[:3]}-{digits[3:5]}-{digits[5:]}", "SSN (no separator)"
        if et:
            tqdm.write(f"  page {page_num} ({label}): box E printed 9 digits with no separator; box I1 says "
                       f"'{entity_type}', so read as an EIN")
            return f"{digits[:2]}-{digits[2:]}", "EIN (no separator)"
        tqdm.write(f"  page {page_num} ({label}): box E printed 9 digits with no separator and box I1 gave no "
                   f"entity type -- 3-2-4 vs 2-7 cannot be determined, recorded as an SSN; spot-check this record")
        return f"{digits[:3]}-{digits[3:5]}-{digits[5:]}", "SSN (no separator, unverified)"

    fixed = window_text.translate(OCR_DIGIT_FIX)
    m = SSN_SEP_RE.search(fixed)
    if m:
        tqdm.write(f"  page {page_num} ({label}): SSN recovered via OCR digit-lookalike correction (a letter "
                   f"ABBYY likely misread as a digit was substituted back)")
        return normalize_ssn(m), "SSN (OCR corrected)"
    m = EIN_SEP_RE.search(fixed)
    if m:
        tqdm.write(f"  page {page_num} ({label}): EIN recovered via OCR digit-lookalike correction")
        return normalize_ein(m), "EIN (OCR corrected)"
    return "", ""


def split_first_mi(first_region):
    tokens = first_region.split()
    if len(tokens) >= 2 and re.fullmatch(r"[A-Za-z]\.?", tokens[-1]):
        return " ".join(tokens[:-1]), tokens[-1].rstrip(".")
    return first_region, ""


def split_name_fallback(name):
    """Peel a trailing single-letter token off as a middle initial, then treat
    the last remaining token as the surname. Box F gives no column position to
    split on, so this is a guess and mis-splits multi-word surnames
    ("VEGA TAPIA", "DE LA CRUZ"); callers log it so the row can be checked."""
    tokens = name.split()
    if not tokens:
        return "", "", ""
    if len(tokens) >= 3 and re.fullmatch(r"[A-Za-z]\.?", tokens[-2]):
        return " ".join(tokens[:-2]), tokens[-2].rstrip("."), tokens[-1]
    if len(tokens) >= 2:
        return " ".join(tokens[:-1]), "", tokens[-1]
    return tokens[0], "", ""


def looks_like_entity(name, entity_type, tin_type):
    if ENTITY_HINT_RE.search(name):
        return True
    et = entity_type.upper()
    if et and "INDIVIDUAL" not in et:
        return True
    return tin_type.startswith("EIN")


def split_partner_name(name, entity_type, tin_type, page_num, label):
    """Split box F's name line into First/MI/Last where that is meaningful, and
    leave it blank where it is not (an organisation, or two people)."""
    if not name:
        return "", "", ""
    if looks_like_entity(name, entity_type, tin_type):
        tqdm.write(f"  page {page_num} ({label}): partner looks like an entity, not an individual "
                   f"(entity type '{entity_type or 'not printed'}', TIN type '{tin_type or 'none'}') -- "
                   f"First/MI/Last left blank, full name kept in the Name column")
        return "", "", ""
    if JOINT_NAME_RE.search(name):
        tqdm.write(f"  page {page_num} ({label}): box F names more than one person -- First/MI/Last left blank, "
                   f"full name kept in the Name column")
        return "", "", ""
    if "," in name:
        last, _, rest = name.partition(",")
        first, mi = split_first_mi(rest.strip())
        return first, mi, last.strip()
    tqdm.write(f"  page {page_num} ({label}): box F has no 'Last name' sub-heading and no comma to split on -- "
               f"falling back to a whitespace-based guess for first/MI/last (may mis-split a multi-word "
               f"surname); spot-check this record")
    return split_name_fallback(name)


def find_entity_type(region_plain):
    """Box I1's value ("INDIVIDUAL", "S CORPORATION", "TRUST" ...), printed
    after the question on the same line, or on the line below it."""
    for i, line in enumerate(region_plain):
        if not ENTITY_TYPE_CAPTION_RE.search(line):
            continue
        _, sep, tail = line.partition("?")
        if sep and tail.strip():
            return tail.strip()
        for nxt in region_plain[i + 1:i + 3]:
            candidate = nxt.strip()
            if candidate and not STOP_LABEL_RE.search(candidate):
                return candidate
        return ""
    return ""


def find_partnership_ein(plain_lines, upto, page_num):
    """Part I box A. Read as a de-duplication discriminator, so that one PDF
    holding K-1s from several partnerships for the same partner does not
    collapse to a single row."""
    head = plain_lines[:upto] if upto else plain_lines
    for i, line in enumerate(head):
        if PARTNERSHIP_EIN_CAPTION_RE.search(line):
            window = " ".join(head[i:i + 3])
            m = EIN_SEP_RE.search(window) or EIN_SEP_RE.search(window.translate(OCR_DIGIT_FIX))
            if m:
                return normalize_ein(m)
            break
    # No heading (or none with a value beside it). Part I sits above Part II, so
    # the first EIN-shaped value before the partner section is the partnership's.
    for line in head:
        m = EIN_SEP_RE.search(line)
        if m:
            return normalize_ein(m)
    tqdm.write(f"  page {page_num}: no partnership EIN (Part I box A) found -- de-duplication for this page "
               f"falls back to TIN/name alone")
    return ""


def find_name_address(region_plain, page_num, label):
    """Box F: the name line, then street line(s), then city/state/ZIP."""
    for i, line in enumerate(region_plain):
        if not NAME_CAPTION_RE.search(line) or PARTNERSHIP_NAME_CAPTION_RE.search(line):
            continue

        scan = list(range(i + 1, min(i + 1 + ADDRESS_BLOCK_MAX_LINES, len(region_plain))))
        block = []
        for idx in scan:
            candidate = region_plain[idx].strip()
            if not candidate:
                continue
            # STOP_LABEL_RE is not applied to the first collected line: that is
            # the partner's name, and a partner can be called
            # "XYZ GENERAL PARTNER LLC".
            if block and STOP_LABEL_RE.search(candidate):
                break
            if not block and TIN_CAPTION_RE.search(candidate):
                break
            block.append(candidate)

        if not block:
            shape = " | ".join(mask_shape(region_plain[idx]) for idx in scan)
            tqdm.write(f"  page {page_num} ({label}): box F heading found but nothing was collected below it "
                       f"(line shapes below the heading, safe to share): {shape}")
            return "", ""

        # A name printed over two lines must not have its second line eaten as
        # the street. Never consume the last line as part of the name -- at
        # least one line has to remain for the address.
        name_parts = [block[0]]
        j = 1
        while j < len(block) - 1 and not STREET_HINT_RE.search(block[j]) \
                and not CITY_STATE_ZIP_RE.match(block[j]):
            name_parts.append(block[j])
            j += 1
        if len(name_parts) > 1:
            tqdm.write(f"  page {page_num} ({label}): box F's name appears to run over {len(name_parts)} lines -- "
                       f"joined them into the Name column; spot-check this record")
        name = " ".join(name_parts)
        addr_lines = block[j:]

        city_idx = next((k for k, l in enumerate(addr_lines) if CITY_STATE_ZIP_RE.match(l)), None)
        if city_idx is not None:
            m = CITY_STATE_ZIP_RE.match(addr_lines[city_idx])
            city = m.group("city").rstrip(",")
            street = " ".join(addr_lines[:city_idx])
            tail = f"{city}, {m.group('state')} {m.group('zip')}"
            address = f"{street}, {tail}" if street else tail
        else:
            address = " ".join(addr_lines)
            if addr_lines:
                shape = " | ".join(mask_shape(l) for l in addr_lines)
                tqdm.write(f"  page {page_num} ({label}): no city/state/ZIP-shaped line in box F (a foreign "
                           f"address, or the lines ran together) -- kept them verbatim; line shapes "
                           f"(safe to share): {shape}")
        return name, address

    tqdm.write(f"  page {page_num} ({label}): box F heading ('Name, address, city, state, and ZIP code for "
               f"partner entered in E') not found")
    return "", ""


def extract_partner(region_plain, page_num, label):
    entity_type = find_entity_type(region_plain)

    tin, tin_type = "", ""
    for i, line in enumerate(region_plain):
        if not TIN_CAPTION_RE.search(line):
            continue
        end = min(i + TIN_WINDOW_LINES, len(region_plain))
        # Cut the window short at box F, so a blank box E cannot reach down and
        # read the partner's street number or ZIP code as a TIN.
        for j in range(i + 1, end):
            if NAME_CAPTION_RE.search(region_plain[j]):
                end = j
                break
        window_lines = region_plain[i:end]
        tin, tin_type = extract_tin(" ".join(window_lines), entity_type, page_num, label)
        if not tin and not tin_type:
            shape = " | ".join(l if TIN_CAPTION_RE.search(l) else mask_shape(l) for l in window_lines)
            tqdm.write(f"  page {page_num} ({label}): box E heading found but no TIN-shaped value beside it "
                       f"-- nearby line shapes (digits/letters masked, safe to share): {shape}")
        break
    else:
        tqdm.write(f"  page {page_num} ({label}): box E heading (\"Partner's SSN or TIN\") not found in this "
                   f"record")

    name, address = find_name_address(region_plain, page_num, label)
    first, mi, last = split_partner_name(name, entity_type, tin_type, page_num, label)

    return {"Page": page_num, "Name": name, "First Name": first, "MI": mi, "Last Name": last,
            "Address": address, "TIN": tin, "TIN Type": tin_type, "Entity Type": entity_type}


def process_page(page, page_num, override_fraction):
    text = page.get_text()
    if len(text.strip()) < MIN_TEXT_CHARS_PER_PAGE:
        tqdm.write(f"  page {page_num}: only {len(text.strip())} chars of text -- scanned/image-only page, "
                   f"skipping (this script expects a real text layer; OCR it first if needed)")
        return []

    words = page.get_text("words")
    page_width = page.rect.width
    full_lines = group_words_into_lines(words)
    full_plain = to_plain_lines(full_lines)

    bound, how = compute_left_column_bound(full_lines, full_plain, page_width, words,
                                           override_fraction, page_num)
    if bound is None:
        tqdm.write(f"  page {page_num}: {how} -- not a K-1 partner page (page 2 / codes page?), skipping")
        return []

    column_words = [w for w in words if (w[0] + w[2]) / 2 <= bound]
    lines = group_words_into_lines(column_words)
    plain = to_plain_lines(lines)

    starts = [i for i, l in enumerate(plain) if TIN_CAPTION_RE.search(l)]
    anchor = "box E heading"
    if not starts:
        starts = [i for i, l in enumerate(plain)
                  if NAME_CAPTION_RE.search(l) and not PARTNERSHIP_NAME_CAPTION_RE.search(l)]
        anchor = "box F heading (no box E heading on the page)"
    if not starts:
        tqdm.write(f"  page {page_num}: Part II is present but neither box E nor box F heading survived inside "
                   f"the partner column (boundary x={bound:.0f} via {how}) -- check with --debug")
        return []

    if len(starts) > 1:
        tqdm.write(f"  page {page_num}: {len(starts)} partner records on this page (split on {anchor})")

    partnership_ein = find_partnership_ein(plain, starts[0], page_num)

    records = []
    for n, start in enumerate(starts):
        end = starts[n + 1] if n + 1 < len(starts) else len(plain)
        label = "Single" if len(starts) == 1 else f"Partner{n + 1}"
        rec = extract_partner(plain[start:end], page_num, label)
        rec["Partnership EIN"] = partnership_ein
        records.append(rec)
    return records


def process_pdf(path: Path, override_fraction):
    doc = fitz.open(path)
    records = []
    for i, page in enumerate(doc, start=1):
        records.extend(process_page(page, i, override_fraction))
    doc.close()
    return records


def dedupe_records(records):
    """Collapse repeated copies of the same K-1 (client copy, file copy) to one
    row, keeping the first occurrence's Page.

    The key includes the partnership EIN because one PDF can hold K-1s from
    several partnerships for the same partner -- keying on TIN alone would
    delete a real K-1. Scoping to Document ID keeps two separately combined
    documents from collapsing into each other.
    """
    seen, deduped = set(), []
    for rec in records:
        doc_id = rec.get("Document ID", "")
        pship = rec.get("Partnership EIN", "")
        tin = rec.get("TIN", "")
        name = rec.get("Name", "")
        address = rec.get("Address", "")
        if tin:
            key = (doc_id, pship, "tin", tin)
        elif name or address:
            key = (doc_id, pship, "name_addr", name, address)
        else:
            # Nothing usable was extracted -- keep every such row rather than
            # collapsing unrelated blank rows into one and hiding the failures.
            deduped.append(rec)
            continue
        if key in seen:
            continue
        seen.add(key)
        deduped.append(rec)
    return deduped


def write_csv(records, output_path):
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        for rec in records:
            writer.writerow({col: rec.get(col, "") for col in CSV_COLUMNS})


def debug_page(path: Path, page_num: int, override_fraction):
    """Print a masked, line-by-line view of one page so a failed extraction can
    be diagnosed without exposing PII: heading lines print in full (they are
    only form labels), every other line is reduced to its digit/letter shape by
    mask_shape() -- safe to paste back for troubleshooting."""
    doc = fitz.open(path)
    if page_num < 1 or page_num > len(doc):
        print(f"{path.name}: page {page_num} out of range (document has {len(doc)} page(s))")
        doc.close()
        return
    page = doc[page_num - 1]
    text = normalize_text(page.get_text())
    words = page.get_text("words")
    page_width = page.rect.width

    print(f"--- {path.name} page {page_num} ---")
    print(f"page width: {page_width:.0f}pt   text layer: {len(text.strip())} chars "
          f"({'OK' if len(text.strip()) >= MIN_TEXT_CHARS_PER_PAGE else 'BELOW MIN -- treated as image-only'})")
    print(f"box E headings: {len(TIN_CAPTION_RE.findall(text))}   "
          f"box F headings: {len(NAME_CAPTION_RE.findall(text))}   "
          f"Part II headings: {len(PART_II_HEADING_RE.findall(text))}")

    full_lines = group_words_into_lines(words)
    full_plain = to_plain_lines(full_lines)
    gutter = detect_column_gutter(words, page_width)
    print(f"gutter scan: {'x=%.0f (%.0f%% of width)' % (gutter, 100 * gutter / page_width) if gutter else 'none found'}")
    bound, how = compute_left_column_bound(full_lines, full_plain, page_width, words,
                                           override_fraction, page_num)
    if bound is None:
        print(f"partner column: NOT ESTABLISHED -- {how}")
        doc.close()
        return
    print(f"partner column: x <= {bound:.0f} ({100 * bound / page_width:.0f}% of width) via {how}")

    def dump(title, lines_plain):
        print(f"\n{title} ({len(lines_plain)} lines)")
        for i, line in enumerate(lines_plain):
            if TIN_CAPTION_RE.search(line):
                tag, shown = "  <-- box E heading", line
            elif NAME_CAPTION_RE.search(line) and not PARTNERSHIP_NAME_CAPTION_RE.search(line):
                tag, shown = "  <-- box F heading", line
            elif PARTNERSHIP_NAME_CAPTION_RE.search(line) or PARTNERSHIP_EIN_CAPTION_RE.search(line):
                tag, shown = "  <-- Part I (partnership) heading", line
            elif ENTITY_TYPE_CAPTION_RE.search(line):
                tag, shown = "  <-- box I1 heading", line
            elif STOP_LABEL_RE.search(line):
                tag, shown = "  <-- stop label", line
            else:
                tag, shown = "", mask_shape(line)
            print(f"  [{i:>3}] {shown}{tag}")

    dump("WHOLE PAGE (both columns mixed together -- this is the problem)", full_plain)
    column_words = [w for w in words if (w[0] + w[2]) / 2 <= bound]
    dump("PARTNER COLUMN ONLY (what the extractor actually reads)", to_plain_lines(group_words_into_lines(column_words)))
    doc.close()


def pick_input_and_output():
    """Ask for the PDFs and the output path when the script was started with no
    arguments, rather than failing with an argparse usage error -- these scripts
    are normally launched from IDLE with F5, where there is no command line to
    pass a path on. Returns (input, output), either of which is None if the
    dialog was cancelled or tkinter is unavailable."""
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except ImportError:
        print("tkinter is not available, so the file pickers cannot open.")
        return None, None

    root = tk.Tk()
    root.withdraw()
    try:
        whole_folder = messagebox.askyesno(
            "Schedule K-1 extractor",
            "Process a whole FOLDER of Schedule K-1 PDFs?\n\n"
            "Yes  -  pick a folder\n"
            "No   -  pick a single PDF file")
        if whole_folder:
            src = filedialog.askdirectory(title="Select the folder holding the Schedule K-1 PDFs")
        else:
            src = filedialog.askopenfilename(title="Select a Schedule K-1 PDF",
                                             filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if not src:
            return None, None

        initial_dir = Path(src) if whole_folder else Path(src).parent
        out = filedialog.asksaveasfilename(
            title="Save the CSV as - use an access-controlled folder, it holds taxpayer identifiers",
            defaultextension=".csv", initialfile=OUTPUT_CSV_NAME, initialdir=str(initial_dir),
            filetypes=[("CSV files", "*.csv")])
        if not out:
            return None, None
        return src, out
    finally:
        root.destroy()


def main():
    global LAUNCHED_INTERACTIVELY

    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("input", nargs="?",
                        help="Schedule K-1 PDF file, or folder of them. Omit it to be asked for it "
                             "with a file picker (which is what happens when this script is run "
                             "from IDLE with F5, or double-clicked)")
    parser.add_argument("-o", "--output", default=OUTPUT_CSV_NAME,
                        help="Combined (all files) output CSV path. Write it to an access-controlled "
                             "location, not a desktop or a source-control working tree "
                             f"(default: {OUTPUT_CSV_NAME})")
    parser.add_argument("--individual-dir", default=None,
                        help="Directory for the individual per-PDF CSVs, named <pdf_stem>_k1_extracted.csv "
                             "(default: same folder as each input PDF)")
    parser.add_argument("--column-right-fraction", type=float, default=None, metavar="F",
                        help="Override the partner-column right boundary as a fraction of page width "
                             "(e.g. 0.5). Only needed when --debug shows the automatic gutter detection "
                             "cutting the address off or letting the income boxes in")
    parser.add_argument("--debug", action="store_true",
                        help="Print a masked line-by-line layout of one page per input file and exit -- "
                             "heading lines print in full, everything else is masked to digit/letter shape "
                             "(safe to paste back for troubleshooting). Use --debug-page to pick the page")
    parser.add_argument("--debug-page", type=int, default=1, metavar="N",
                        help="Page number to debug when --debug is set (default: 1)")
    args = parser.parse_args()

    if args.input is None:
        LAUNCHED_INTERACTIVELY = True
        picked_input, picked_output = pick_input_and_output()
        if not picked_input:
            print("Nothing selected -- nothing to do.\n\n"
                  "From a command line you can also pass the path directly:\n"
                  f'  python "{Path(__file__).name}" <pdf_or_folder> [-o out.csv]\n'
                  f'  python "{Path(__file__).name}" <pdf> --debug')
            return
        args.input, args.output = picked_input, picked_output

    input_path = Path(args.input)
    pdf_files = [input_path] if input_path.is_file() else sorted(input_path.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found at: {input_path}")
        return

    if args.debug:
        for pdf in pdf_files:
            debug_page(pdf, args.debug_page, args.column_right_fraction)
        return

    individual_dir = Path(args.individual_dir) if args.individual_dir else None
    if individual_dir:
        individual_dir.mkdir(parents=True, exist_ok=True)

    all_records = []
    with tqdm(pdf_files, desc="Processing K-1 PDFs", unit="pdf") as pbar:
        for pdf in pbar:
            pbar.set_postfix_str(pdf.name)
            records = process_pdf(pdf, args.column_right_fraction)
            for rec in records:
                rec["Document ID"] = pdf.stem

            if not records:
                tqdm.write(f"  {pdf.name}: no partner records extracted")
                continue

            file_deduped = dedupe_records(records)
            removed = len(records) - len(file_deduped)
            if removed:
                tqdm.write(f"  {pdf.name}: removed {removed} duplicate record(s) within this file "
                           f"(same partnership EIN + TIN -- repeated copies of one K-1)")

            individual_path = (individual_dir or pdf.parent) / f"{pdf.stem}_k1_extracted.csv"
            write_csv(file_deduped, individual_path)
            tqdm.write(f"  {pdf.name}: -> {individual_path} ({len(file_deduped)} record(s))")
            all_records.extend(file_deduped)

    if not all_records:
        print("No records extracted from any file -- combined CSV not written.")
        return

    combined = dedupe_records(all_records)
    removed = len(all_records) - len(combined)
    if removed:
        print(f"\nRemoved {removed} duplicate record(s) across files "
              f"(same Document ID + partnership EIN + TIN)")

    write_csv(combined, Path(args.output))
    print(f"\nWrote combined file: {len(combined)} record(s) -> {args.output}")
    print("Reminder: this CSV holds taxpayer identifiers in clear text -- store it in the appropriate "
          "access-controlled Global Insider folder and do not commit it to source control.")


if __name__ == "__main__":
    try:
        main()
    finally:
        # Started by double-click or F5: hold the window open so the summary and
        # any per-page warnings can be read before it disappears.
        if LAUNCHED_INTERACTIVELY:
            try:
                input("\nPress Enter to close...")
            except EOFError:
                pass
