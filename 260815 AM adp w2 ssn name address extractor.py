r"""
ADP "W-2 and Earnings Summary"  --  Employee SSN, Name and Address extractor
============================================================================
Reads searchable ADP W-2 statements and writes one CSV row per employee:
Name, First/MI/Last, Address, SSN.

WHY THE EXISTING W-2 SCRIPT RETURNS A BLANK SSN ON THIS FORMAT
    w2_ssn_name_address_csv.py anchors the SSN on

        Employee's social security number   /   Employee's SSN

    but ADP labels that box

        a  Employee's SSA number

    "SSA number" matches neither alternative, so the caption is never found
    and the SSN comes back empty. The name/address caption ("e/f Employee's
    name, address, and ZIP code") does still match, which is why only the
    SSN went missing. This script matches all three wordings.

THE PAGE HAS TWO PANES, AND ONLY THE LEFT ONE MAY BE READ
    An ADP statement prints the W-2 itself on the left and ADP's Earnings
    Summary on the right, at overlapping heights. Rebuilding lines by
    vertical position glues the two together, so the employee's street line
    picks up the summary's figures:

        printed (two panes)                  read as one line
        123 SAMPLE ST     Gross Pay ...      123 SAMPLE ST Gross Pay 104,485.46
        SAMPLE CITY CA 92336   Plus GTL      SAMPLE CITY CA 92336 Plus GTL ...

    The right pane also repeats the employee's name, address and SSN under
    "3. Employee W-4 Profile". That copy is deliberately NOT read: it carries
    no box caption of its own, and its "Social Security Number:" label has no
    "Employee's" prefix, so neither anchor here can reach it.

    detect_gutters() finds the pane boundaries from word geometry alone --
    it marks which horizontal bands hold any word and keeps the empty ones
    wide enough to be printed gutters. Each caption is then confined to the
    pane bounded by the nearest gutter on either side of the caption itself,
    so a boundary can never fall inside the block being read.

    Occupancy is measured over a vertical band around the caption rather than
    the whole page (CAPTION_Y_BAND_*): the tear-off copies across the bottom
    of the page have their own gutters, which the full-page text above would
    otherwise paint over.

THE SSN VALUE IS TIED TO ITS CAPTION BY X-POSITION
    ADP prints two boxes side by side:

        b Employer's FED ID number | a Employee's SSA number
        45-1234567                 | 123-45-6789

    Both values land on one rebuilt line, so "the number under the caption"
    has to mean the one starting at or right of where "Employee's SSA number"
    itself starts -- otherwise the employer's FED ID is read as the SSN.
    caption_spans_x() maps the matched caption text back to the words that
    carry it, which gives that x-position.

THE NAME/ADDRESS BLOCK IS POSITIONAL
    Box e/f prints, in order: the name on line 1, then the street, then an
    optional APT/Suite line, then CITY STATE ZIP. Line 1 is always taken as
    the name; the remaining lines are the address, with the CITY STATE ZIP
    line located by shape so an extra APT/Suite line does not shift anything.
    A verbatim Name column is written alongside First/MI/Last because this
    layout has no "Last name" sub-caption to anchor a column split on, so the
    split is a best-effort guess -- see split_employee_name().

ONE PDF HOLDS MANY EMPLOYEES
    De-duplication is therefore keyed on Document ID + SSN, never on SSN
    alone across files: different employees keep their own rows, while the
    repeated tear-off copies of one employee's W-2 (Copy B / Copy C / Copy 2,
    printed across the bottom of the same page) collapse to a single row.
    Where those bottom copies share a pane with each other, their addresses
    can pick up the neighbouring copy's text -- the row kept is the first
    one found, which is the clean top-left statement, and the count collapsed
    is logged.

OCR TOLERANCES
    normalize_text() folds curly quotes/dashes and stray whitespace; the
    caption patterns allow a missing apostrophe ("Employees SSA number"); and
    if no strictly numeric SSN is found beside the caption, digits are
    recovered from common OCR look-alikes (O/0, I/1, S/5 ...) within that
    small window only -- never page-wide.

HANDLING NOTE
    The output CSV contains SSNs and home addresses in clear text. Write it
    to an access-controlled location (the appropriate Global Insider folder),
    not to a desktop or a source-control working tree.

Usage:
    Run it with no arguments -- from IDLE with F5, or by double-clicking --
    and it asks for the PDFs and for where to save the CSV.

    Or from a command line:
      python "260815 AM adp w2 ssn name address extractor.py" <pdf_or_folder> [-o out.csv]
      python "260815 AM adp w2 ssn name address extractor.py" <pdf> --debug [--debug-page N]

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

CSV_COLUMNS = ["Document ID", "Page", "Name", "First Name", "MI", "Last Name", "Address", "SSN"]

OUTPUT_CSV_NAME = "adp_w2_extracted.csv"

# Set when the script was started with no arguments (IDLE F5, or a
# double-click), so the console window is held open at the end.
LAUNCHED_INTERACTIVELY = False

# --- pane detection ---------------------------------------------------------
GUTTER_BIN_POINTS = 1.0
# An empty vertical band narrower than this is word spacing, not a printed
# gutter between panes.
GUTTER_MIN_WIDTH_FRACTION = 0.010
# Page margins are empty bands too; ignore them.
GUTTER_SEARCH_LO_FRACTION = 0.03
GUTTER_SEARCH_HI_FRACTION = 0.97
# Occupancy is measured over this vertical band around the caption, not the
# whole page -- see the module docstring.
CAPTION_Y_BAND_ABOVE = 120.0
CAPTION_Y_BAND_BELOW = 180.0
# How far left of the SSN caption's own start a value may sit and still count
# as being in its sub-column.
SSN_SUBCOLUMN_PAD_FRACTION = 0.015

NAME_BLOCK_MAX_LINES = 8
SSN_WINDOW_LINES = 4
LINE_Y_TOLERANCE = 3

# ".?" (not ".") for the apostrophe so a dropped or misread apostrophe still
# matches ("Employees SSA number"). The "Employee.?s" prefix is required, which
# is what keeps the Earnings Summary pane's bare "Social Security Number:"
# label (under "3. Employee W-4 Profile") from being picked up.
SSN_CAPTION_RE = re.compile(
    r"Employee.?s\s+(?:SSA\s+number|social\s+security\s+number|SSN)\b", re.IGNORECASE)

# ADP prints "e/f Employee's name, address, and ZIP code"; the plain W-2
# prints "e Employee's first name and initial". Both are accepted, and the
# comma/"address," are optional so a slightly different ADP revision still
# matches. "Employer's ..." cannot match: "Employee" requires an 'e' where
# "Employer" has an 'r'.
NAME_CAPTION_RE = re.compile(
    r"Employee.?s\s+name,?\s*(?:address,?\s*)?and\s+ZIP\s+code"
    r"|Employee.?s\s+first\s+name\s+and\s+initial", re.IGNORECASE)

# On the plain W-2 layout the address is introduced by its own caption below
# box e. That line is skipped over inside the block rather than treated as a
# stop label -- stopping there would end the block right before the address it
# introduces.
ADDRESS_CAPTION_RE = re.compile(r"Employee.?s\s+address", re.IGNORECASE)

# Recognised only so the debug dump can label it, and to be certain it is never
# mistaken for the employee's block.
EMPLOYER_NAME_CAPTION_RE = re.compile(r"Employer.?s\s+name,?\s*address", re.IGNORECASE)

# What comes after box e/f. Never applied to the block's first line -- that is
# the employee's name, and a surname can collide with almost any word.
STOP_LABEL_RE = re.compile(
    r"Employer.?s\s+FED\s+ID|Employee.?s\s+SSA|Employer.?s\s+name|Employer.?s\s+state\s+ID"
    r"|Employer\s+identification|Wages,\s*tips|Federal\s+income\s+tax"
    r"|Social\s+security\s+(?:wages|tax|tips)|Medicare\s+wages|Allocated\s+tips"
    r"|Control\s+number|Verification\s+Code|Dependent\s+care|Nonqualified\s+plans"
    r"|See\s+instructions\s+for\s+box\s+12|Statutory\s+employee|Retirement\s+plan"
    r"|State\s+wages|Local\s+wages|Local\s+income\s+tax|Locality\s+name|Batch\s*#"
    r"|Wage\s*(?:&|and)\s*Tax\s+Statement|Copy\s+[A-Z0-9]\b|Department\s+of\s+the\s+Treasury"
    r"|Employee\s+Reference\s+Copy|Earnings\s+Summary|Employee\s+W-?4\s+Profile"
    r"|^Form\s*W-?2\b|OMB\s*No",
    re.IGNORECASE)

# Dashes required, so an employer FED ID (2-7) can never satisfy it.
SSN_STRICT_RE = re.compile(r"\b(\d{3})-(\d{2})-(\d{4})\b")
# Separator optional, for text layers that dropped the dashes. Only used after
# EIN-shaped tokens have been removed, since 9 bare digits also fit 2-7.
SSN_LOOSE_RE = re.compile(r"\b(\d{3})[-\s]?(\d{2})[-\s]?(\d{4})\b")
EIN_VALUE_RE = re.compile(r"\b\d{2}-\d{7}\b")

CITY_STATE_ZIP_RE = re.compile(r"^(?P<city>.+?),?\s+(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)\b")

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

ENTITY_HINT_RE = re.compile(r"\b(?:LLC|INC|CORP|CO|COMPANY|LP|LLP|TRUST|ESTATE|LTD)\b\.?",
                            re.IGNORECASE)


def mask_shape(s):
    """Reduce a line to its digit/letter shape so a layout can be shared for
    troubleshooting without exposing the values themselves."""
    return re.sub(r"[A-Za-z]", "X", re.sub(r"\d", "#", s))


def normalize_text(s):
    s = unicodedata.normalize("NFKC", s)
    s = s.translate({0x2018: "'", 0x2019: "'", 0x00B4: "'", 0x0060: "'",
                     0x201C: '"', 0x201D: '"', 0x00A0: " ",
                     0x2013: "-", 0x2014: "-"})
    return re.sub(r"[ \t]+", " ", s).strip()


def group_words_into_lines(words, y_tol=LINE_Y_TOLERANCE):
    """Rebuild printed lines from word boxes, clustering by top edge. Returns
    [(y, [(x0, x1, text), ...]), ...] sorted top to bottom, each line's words
    left to right. Done here rather than taken from the PDF's own line/block
    structure because that structure spans both panes."""
    items = sorted(((w[1], w[0], w[2], w[4]) for w in words), key=lambda t: (t[0], t[1]))
    lines, current, current_y = [], [], None
    for y0, x0, x1, text in items:
        if current_y is not None and abs(y0 - current_y) > y_tol:
            lines.append((current_y, sorted(current, key=lambda t: t[0])))
            current, current_y = [], None
        if current_y is None:
            current_y = y0
        current.append((x0, x1, text))
    if current:
        lines.append((current_y, sorted(current, key=lambda t: t[0])))
    return lines


def line_text(line_words):
    return normalize_text(" ".join(t for _, _, t in line_words))


def caption_spans_x(line_words, caption_re):
    """Map EVERY match of a caption back to the words that carry it, returning
    a list of (x0, x1).

    Two reasons this has to find every match, not just the first:
      - a rebuilt line can hold several boxes' captions -- "b Employer's FED ID
        number a Employee's SSA number" is one line, and only the second half's
        x-position identifies the SSN's sub-column;
      - W-2s printed side by side (the tear-off copies across the bottom of an
        ADP page, or a 2-up page of two different employees) put their captions
        at the same height, so one rebuilt line carries the same caption twice.
        Taking only the first would silently drop the second employee.
    """
    parts, offsets, pos = [], [], 0
    for x0, x1, text in line_words:
        token = normalize_text(text)
        offsets.append((pos, pos + len(token), x0, x1))
        parts.append(token)
        pos += len(token) + 1

    spans = []
    for m in caption_re.finditer(" ".join(parts)):
        lo = hi = None
        for start, end, x0, x1 in offsets:
            if end > m.start() and start < m.end():
                lo = x0 if lo is None else min(lo, x0)
                hi = x1 if hi is None else max(hi, x1)
        if lo is not None:
            spans.append((lo, hi))
    return spans


def detect_gutters(words, page_width, y_lo, y_hi, min_width_fraction):
    """Find the printed vertical gutters between panes from word geometry
    alone: mark every horizontal bin holding a word within the vertical band
    y_lo..y_hi, then return the midpoint of each empty run wide enough to be a
    gutter rather than word spacing."""
    if page_width <= 0:
        return []
    bins = int(page_width / GUTTER_BIN_POINTS) + 1
    occupied = bytearray(bins)
    for w in words:
        if not str(w[4]).strip():
            continue
        if w[3] < y_lo or w[1] > y_hi:
            continue
        lo = max(0, int(w[0] / GUTTER_BIN_POINTS))
        hi = min(bins - 1, int(w[2] / GUTTER_BIN_POINTS))
        for b in range(lo, hi + 1):
            occupied[b] = 1

    search_lo = int(page_width * GUTTER_SEARCH_LO_FRACTION / GUTTER_BIN_POINTS)
    search_hi = min(bins, int(page_width * GUTTER_SEARCH_HI_FRACTION / GUTTER_BIN_POINTS))
    min_bins = max(1, int(page_width * min_width_fraction / GUTTER_BIN_POINTS))

    gutters, run_lo = [], None
    for b in range(search_lo, search_hi + 1):
        empty = b < search_hi and not occupied[b]
        if empty:
            if run_lo is None:
                run_lo = b
        elif run_lo is not None:
            if b - run_lo >= min_bins:
                gutters.append((run_lo + b) / 2 * GUTTER_BIN_POINTS)
            run_lo = None
    return gutters


def pane_bounds(x0, x1, gutters, page_width):
    """The pane holding a caption: bounded by the nearest gutter outside the
    caption's own span on each side, so a boundary can never fall inside the
    block being read."""
    left = max([g for g in gutters if g <= x0], default=0.0)
    right = min([g for g in gutters if g >= x1], default=float(page_width))
    return left, right


def words_in_pane(words, left, right):
    return [w for w in words if left <= (w[0] + w[2]) / 2 <= right]


def siblings_in_pane(name_hits, y, left, right):
    """How many box e/f captions sit at this same height inside this pane. More
    than one means that many W-2s are printed side by side there -- the gutter
    between them was painted over by whatever spans them further up the page."""
    return sum(1 for hy, hx0, _ in name_hits
               if abs(hy - y) <= LINE_Y_TOLERANCE * 2 and left <= hx0 <= right)


def block_x_window(cap_x0, cap_x1, left, right, page_width, siblings, page_num, label):
    """The x range to read a caption's block from: the whole pane normally, but
    only the caption's own width when W-2s are printed side by side in it --
    otherwise the neighbouring copy's name and street lines are read into this
    one. Decided by counting the captions at that height rather than by
    comparing widths, so a single W-2 in a roomy pane is left alone."""
    if siblings <= 1:
        return left, right
    pad = page_width * SSN_SUBCOLUMN_PAD_FRACTION
    tqdm.write(f"  page {page_num} ({label}): {siblings} W-2s are printed side by side in this pane "
               f"(x {left:.0f}-{right:.0f}) -- reading this one's name/address from its own caption width only")
    return cap_x0 - pad, cap_x1 + pad


def nearest_line_index(lines, y):
    best, best_d = None, None
    for i, (ly, _) in enumerate(lines):
        d = abs(ly - y)
        if best_d is None or d < best_d:
            best, best_d = i, d
    return best if best_d is not None and best_d <= LINE_Y_TOLERANCE * 2 else None


def normalize_ssn(match):
    return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"


def find_ssn_in(text):
    """Strict (dashed) first, then separator-less once EIN-shaped tokens have
    been removed, then the same two again after OCR digit look-alikes are mapped
    back. Returns (ssn, how) or (None, None)."""
    m = SSN_STRICT_RE.search(text)
    if m:
        return normalize_ssn(m), "dashed"
    stripped = EIN_VALUE_RE.sub(" ", text)
    m = SSN_LOOSE_RE.search(stripped)
    if m:
        return normalize_ssn(m), "no separator"
    fixed = text.translate(OCR_DIGIT_FIX)
    m = SSN_STRICT_RE.search(fixed)
    if m:
        return normalize_ssn(m), "OCR corrected"
    m = SSN_LOOSE_RE.search(EIN_VALUE_RE.sub(" ", fixed))
    if m:
        return normalize_ssn(m), "OCR corrected, no separator"
    return None, None


def find_ssn(pane_lines, caption_idx, ssn_x0, page_width, page_num, label):
    """Read the SSN from the lines at and just below its caption, restricted to
    the caption's own sub-column so the neighbouring "Employer's FED ID number"
    value cannot be picked up instead."""
    pad = page_width * SSN_SUBCOLUMN_PAD_FRACTION
    end = min(caption_idx + SSN_WINDOW_LINES, len(pane_lines))
    restricted, full = [], []
    for k in range(caption_idx, end):
        _, lw = pane_lines[k]
        if k > caption_idx and NAME_CAPTION_RE.search(line_text(lw)):
            break
        restricted.append(normalize_text(
            " ".join(t for x0, x1, t in lw if (x0 + x1) / 2 >= ssn_x0 - pad)))
        full.append(line_text(lw))

    ssn, how = find_ssn_in(" ".join(restricted))
    if ssn:
        if how != "dashed":
            tqdm.write(f"  page {page_num} ({label}): SSN read via {how} tolerance -- spot-check this record")
        return ssn

    # Nothing in the caption's own sub-column. Retry across the full width of
    # those lines, but dashed-only, so the employer's FED ID still cannot win.
    m = SSN_STRICT_RE.search(" ".join(full))
    if m:
        tqdm.write(f"  page {page_num} ({label}): no SSN in the 'Employee's SSA number' sub-column, but a dashed "
                   f"SSN was found elsewhere on those lines -- used it; the box may be positioned differently on "
                   f"this document, check with --debug")
        return normalize_ssn(m)

    shape = " | ".join(l if SSN_CAPTION_RE.search(l) else mask_shape(l) for l in full)
    tqdm.write(f"  page {page_num} ({label}): SSN caption found but no SSN-shaped value beside it -- nearby line "
               f"shapes (digits/letters masked, safe to share): {shape}")
    return ""


def split_first_mi(first_region):
    tokens = first_region.split()
    if len(tokens) >= 2 and re.fullmatch(r"[A-Za-z]\.?", tokens[-1]):
        return " ".join(tokens[:-1]), tokens[-1].rstrip(".")
    return first_region, ""


def split_employee_name(name, page_num, label):
    """Best-effort First/MI/Last. This layout prints the whole name on one line
    with no "Last name" sub-caption to give a column position to split on, so:
    a comma form ("LAST, FIRST M") splits exactly; otherwise a trailing single
    letter is taken as the middle initial and the last token as the surname,
    which mis-splits a multi-word surname ("VEGA TAPIA", "DE LA CRUZ") and is
    logged. The verbatim Name column always keeps the printed line."""
    if not name:
        return "", "", ""
    if ENTITY_HINT_RE.search(name):
        tqdm.write(f"  page {page_num} ({label}): box e/f line 1 looks like an organisation, not a person -- "
                   f"First/MI/Last left blank, full text kept in the Name column")
        return "", "", ""
    if "," in name:
        last, _, rest = name.partition(",")
        first, mi = split_first_mi(rest.strip())
        return first, mi, last.strip()
    tokens = name.split()
    if len(tokens) >= 3 and re.fullmatch(r"[A-Za-z]\.?", tokens[1]):
        # "RYAN P ADAMS" -- the initial sits between the given and family name.
        return tokens[0], tokens[1].rstrip("."), " ".join(tokens[2:])
    if len(tokens) >= 2:
        tqdm.write(f"  page {page_num} ({label}): no middle initial and no comma in box e/f line 1 -- treating the "
                   f"last word as the surname (may mis-split a multi-word surname); spot-check this record")
        return " ".join(tokens[:-1]), "", tokens[-1]
    return tokens[0] if tokens else "", "", ""


def find_name_address(pane_lines, caption_idx, page_num, label):
    """Box e/f: line 1 is the name, then the street, then an optional APT/Suite
    line, then CITY STATE ZIP."""
    block = []
    for k in range(caption_idx + 1, min(caption_idx + 1 + NAME_BLOCK_MAX_LINES, len(pane_lines))):
        candidate = line_text(pane_lines[k][1])
        if not candidate:
            continue
        if ADDRESS_CAPTION_RE.search(candidate) and not block:
            # Plain-W-2 layout: box f's own caption sits between the name and
            # the address it introduces. Step over it.
            continue
        if block and STOP_LABEL_RE.search(candidate):
            break
        block.append(candidate)

    if not block:
        tqdm.write(f"  page {page_num} ({label}): box e/f caption found but nothing was collected below it")
        return "", ""

    name = block[0]
    addr_lines = block[1:]
    if not addr_lines:
        tqdm.write(f"  page {page_num} ({label}): box e/f gave a name but no address lines below it")
        return name, ""

    city_idx = next((k for k, l in enumerate(addr_lines) if CITY_STATE_ZIP_RE.match(l)), None)
    if city_idx is None:
        shape = " | ".join(mask_shape(l) for l in addr_lines)
        tqdm.write(f"  page {page_num} ({label}): no CITY STATE ZIP-shaped line in box e/f (a foreign address, or "
                   f"the lines ran together) -- kept the lines verbatim; shapes (safe to share): {shape}")
        return name, " ".join(addr_lines)

    m = CITY_STATE_ZIP_RE.match(addr_lines[city_idx])
    city = m.group("city").rstrip(",")
    street = " ".join(addr_lines[:city_idx])
    tail = f"{city}, {m.group('state')} {m.group('zip')}"
    return name, (f"{street}, {tail}" if street else tail)


def process_page(page, page_num, min_gutter_fraction):
    text = page.get_text()
    if len(text.strip()) < MIN_TEXT_CHARS_PER_PAGE:
        tqdm.write(f"  page {page_num}: only {len(text.strip())} chars of text -- scanned/image-only page, "
                   f"skipping (this script expects a real text layer; OCR it first if needed)")
        return []

    words = page.get_text("words")
    page_width = page.rect.width
    full_lines = group_words_into_lines(words)

    # No employer-caption guard is needed here: "Employer's" cannot match
    # NAME_CAPTION_RE ("Employee" requires an 'e' where "Employer" has an 'r'),
    # and skipping the whole line would drop an employee caption printed beside
    # an employer one.
    name_hits, ssn_hits = [], []
    for y, lw in full_lines:
        name_hits.extend((y, x0, x1) for x0, x1 in caption_spans_x(lw, NAME_CAPTION_RE))
        ssn_hits.extend((y, x0, x1) for x0, x1 in caption_spans_x(lw, SSN_CAPTION_RE))
    name_hits.sort(key=lambda h: (h[0], h[1]))
    ssn_hits.sort(key=lambda h: (h[0], h[1]))

    if not name_hits and not ssn_hits:
        tqdm.write(f"  page {page_num}: no 'Employee's name/address' or 'Employee's SSA number' caption -- not an "
                   f"employee page, skipping")
        return []

    used_ssn = set()
    records = []

    for n, (y, x0, x1) in enumerate(name_hits):
        label = "Single" if len(name_hits) == 1 else f"Employee{n + 1}"
        gutters = detect_gutters(words, page_width, y - CAPTION_Y_BAND_ABOVE,
                                 y + CAPTION_Y_BAND_BELOW, min_gutter_fraction)
        left, right = pane_bounds(x0, x1, gutters, page_width)
        # The SSN is read from the whole pane (its caption sits in a different
        # sub-column of the same box row); the name/address block is read from
        # the caption's own width when the pane holds more than one W-2.
        pane_lines = group_words_into_lines(words_in_pane(words, left, right))
        nlo, nhi = block_x_window(x0, x1, left, right, page_width,
                                  siblings_in_pane(name_hits, y, left, right), page_num, label)
        name_lines = (pane_lines if (nlo, nhi) == (left, right)
                      else group_words_into_lines(words_in_pane(words, nlo, nhi)))
        idx = nearest_line_index(name_lines, y)
        if idx is None:
            tqdm.write(f"  page {page_num} ({label}): the box e/f caption could not be relocated inside its own "
                       f"pane (x {left:.0f}-{right:.0f}) -- skipping this record")
            continue

        name, address = find_name_address(name_lines, idx, page_num, label)

        # Pair with the SSN caption in the same pane, nearest by vertical
        # distance -- the tear-off copies repeat both captions down the page.
        # Nearest by vertical distance, then by horizontal distance -- side by
        # side copies in one pane put their SSN captions at the same height, so
        # x is what tells this employee's from the neighbouring one's.
        candidates = [(abs(sy - y), abs(sx0 - x0), k) for k, (sy, sx0, _) in enumerate(ssn_hits)
                      if k not in used_ssn and left <= sx0 <= right]
        ssn = ""
        if candidates:
            _, _, k = min(candidates)
            used_ssn.add(k)
            sy, sx0, _ = ssn_hits[k]
            s_idx = nearest_line_index(pane_lines, sy)
            if s_idx is None:
                tqdm.write(f"  page {page_num} ({label}): the SSN caption could not be relocated inside the pane "
                           f"-- SSN left blank")
            else:
                ssn = find_ssn(pane_lines, s_idx, sx0, page_width, page_num, label)
        else:
            tqdm.write(f"  page {page_num} ({label}): no \"Employee's SSA number\" caption in this employee's pane "
                       f"-- SSN left blank")

        first, mi, last = split_employee_name(name, page_num, label)
        records.append({"Page": page_num, "Name": name, "First Name": first, "MI": mi,
                        "Last Name": last, "Address": address, "SSN": ssn})

    # An SSN caption with no name/address caption in its pane still identifies
    # an employee; emit it rather than dropping it silently.
    for k, (y, x0, x1) in enumerate(ssn_hits):
        if k in used_ssn:
            continue
        label = f"SSN-only{k + 1}"
        gutters = detect_gutters(words, page_width, y - CAPTION_Y_BAND_ABOVE,
                                 y + CAPTION_Y_BAND_BELOW, min_gutter_fraction)
        left, right = pane_bounds(x0, x1, gutters, page_width)
        pane_lines = group_words_into_lines(words_in_pane(words, left, right))
        idx = nearest_line_index(pane_lines, y)
        if idx is None:
            continue
        ssn = find_ssn(pane_lines, idx, x0, page_width, page_num, label)
        if not ssn:
            continue
        tqdm.write(f"  page {page_num} ({label}): an SSN was found with no name/address caption in its pane -- "
                   f"row written with the name and address blank")
        records.append({"Page": page_num, "Name": "", "First Name": "", "MI": "", "Last Name": "",
                        "Address": "", "SSN": ssn})

    return records


def process_pdf(path: Path, min_gutter_fraction):
    doc = fitz.open(path)
    records = []
    for i, page in enumerate(doc, start=1):
        records.extend(process_page(page, i, min_gutter_fraction))
    doc.close()
    return records


def dedupe_records(records):
    """One PDF holds many employees, and each employee's W-2 is printed several
    times on a page (the reference copy plus the tear-off Copy B / C / 2). Key
    on Document ID + SSN so repeats of one employee collapse while different
    employees keep their own rows, and never on SSN alone across documents.
    The first occurrence is kept -- that is the top-left reference copy, whose
    pane is cleanly isolated."""
    seen, deduped = set(), []
    for rec in records:
        doc_id = rec.get("Document ID", "")
        ssn = rec.get("SSN", "")
        name = rec.get("Name", "")
        address = rec.get("Address", "")
        if ssn:
            key = (doc_id, "ssn", ssn)
        elif name or address:
            key = (doc_id, "name_addr", name, address)
        else:
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


def debug_page(path: Path, page_num: int, min_gutter_fraction):
    """Print a masked, pane-by-pane view of one page so a failed extraction can
    be diagnosed without exposing PII: caption lines print in full (they are
    only form labels), every other line is reduced to its digit/letter shape --
    safe to paste back for troubleshooting."""
    doc = fitz.open(path)
    if page_num < 1 or page_num > len(doc):
        print(f"{path.name}: page {page_num} out of range (document has {len(doc)} page(s))")
        doc.close()
        return
    page = doc[page_num - 1]
    words = page.get_text("words")
    page_width = page.rect.width
    text = normalize_text(page.get_text())
    full_lines = group_words_into_lines(words)

    print(f"--- {path.name} page {page_num} ---")
    print(f"page width: {page_width:.0f}pt   text layer: {len(text.strip())} chars "
          f"({'OK' if len(text.strip()) >= MIN_TEXT_CHARS_PER_PAGE else 'BELOW MIN -- image-only'})")
    print(f"'Employee's SSA number'-style captions: {len(SSN_CAPTION_RE.findall(text))}   "
          f"box e/f captions: {len(NAME_CAPTION_RE.findall(text))}")

    def dump(title, lines):
        print(f"\n{title} ({len(lines)} lines)")
        for i, (y, lw) in enumerate(lines):
            line = line_text(lw)
            if SSN_CAPTION_RE.search(line):
                tag, shown = "  <-- SSN caption", line
            elif NAME_CAPTION_RE.search(line) and not EMPLOYER_NAME_CAPTION_RE.search(line):
                tag, shown = "  <-- box e/f caption", line
            elif EMPLOYER_NAME_CAPTION_RE.search(line):
                tag, shown = "  <-- employer caption (ignored)", line
            elif STOP_LABEL_RE.search(line):
                tag, shown = "  <-- stop label", line
            else:
                tag, shown = "", mask_shape(line)
            print(f"  [{i:>3}] y={y:>6.1f}  {shown}{tag}")

    dump("WHOLE PAGE (both panes mixed together -- this is the problem)", full_lines)

    name_hits = []
    for y, lw in full_lines:
        name_hits.extend((y, x0, x1) for x0, x1 in caption_spans_x(lw, NAME_CAPTION_RE))

    for y, x0, x1 in name_hits:
        gutters = detect_gutters(words, page_width, y - CAPTION_Y_BAND_ABOVE,
                                 y + CAPTION_Y_BAND_BELOW, min_gutter_fraction)
        left, right = pane_bounds(x0, x1, gutters, page_width)
        siblings = siblings_in_pane(name_hits, y, left, right)
        pad = page_width * SSN_SUBCOLUMN_PAD_FRACTION
        nlo, nhi = (left, right) if siblings <= 1 else (x0 - pad, x1 + pad)
        print(f"\ncaption at y={y:.1f} x={x0:.0f}-{x1:.0f}")
        print(f"  gutters near it: {', '.join(f'{g:.0f}' for g in gutters) or 'none'}")
        print(f"  pane: x {left:.0f} .. {right:.0f} "
              f"({100 * left / page_width:.0f}% .. {100 * right / page_width:.0f}% of width)")
        if siblings > 1:
            print(f"  {siblings} W-2s side by side at this height -- name/address narrowed to "
                  f"x {nlo:.0f} .. {nhi:.0f}")
        dump("  EMPLOYEE BLOCK ONLY (what the extractor actually reads)",
             group_words_into_lines(words_in_pane(words, nlo, nhi)))
    doc.close()


def pick_input_and_output():
    """Ask for the PDFs and the output path when the script was started with no
    arguments, rather than failing with an argparse usage error -- these scripts
    are normally launched from IDLE with F5, where there is no command line."""
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
            "ADP W-2 extractor",
            "Process a whole FOLDER of ADP W-2 PDFs?\n\n"
            "Yes  -  pick a folder\n"
            "No   -  pick a single PDF file")
        if whole_folder:
            src = filedialog.askdirectory(title="Select the folder holding the ADP W-2 PDFs")
        else:
            src = filedialog.askopenfilename(title="Select an ADP W-2 PDF",
                                             filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if not src:
            return None, None

        initial_dir = Path(src) if whole_folder else Path(src).parent
        out = filedialog.asksaveasfilename(
            title="Save the CSV as - use an access-controlled folder, it holds SSNs",
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
                        help="ADP W-2 PDF file, or folder of them. Omit it to be asked for it with a "
                             "file picker (which is what happens when this script is run from IDLE "
                             "with F5, or double-clicked)")
    parser.add_argument("-o", "--output", default=OUTPUT_CSV_NAME,
                        help="Combined (all files) output CSV path. Write it to an access-controlled "
                             f"location, not a desktop or a working tree (default: {OUTPUT_CSV_NAME})")
    parser.add_argument("--individual-dir", default=None,
                        help="Directory for the individual per-PDF CSVs, named <pdf_stem>_adp_w2.csv "
                             "(default: same folder as each input PDF)")
    parser.add_argument("--min-gutter-width", type=float, default=GUTTER_MIN_WIDTH_FRACTION,
                        metavar="F",
                        help="Narrowest empty vertical band, as a fraction of page width, still counted "
                             "as a gutter between panes. Raise it if unrelated columns are being merged, "
                             f"lower it if a pane is being cut in half (default: {GUTTER_MIN_WIDTH_FRACTION})")
    parser.add_argument("--debug", action="store_true",
                        help="Print a masked pane-by-pane layout of one page per input file and exit -- "
                             "caption lines print in full, everything else is masked to digit/letter "
                             "shape (safe to paste back for troubleshooting)")
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
            debug_page(pdf, args.debug_page, args.min_gutter_width)
        return

    individual_dir = Path(args.individual_dir) if args.individual_dir else None
    if individual_dir:
        individual_dir.mkdir(parents=True, exist_ok=True)

    all_records = []
    with tqdm(pdf_files, desc="Processing ADP W-2 PDFs", unit="pdf") as pbar:
        for pdf in pbar:
            pbar.set_postfix_str(pdf.name)
            records = process_pdf(pdf, args.min_gutter_width)
            for rec in records:
                rec["Document ID"] = pdf.stem

            if not records:
                tqdm.write(f"  {pdf.name}: no employee records extracted")
                continue

            file_deduped = dedupe_records(records)
            removed = len(records) - len(file_deduped)
            if removed:
                tqdm.write(f"  {pdf.name}: removed {removed} duplicate record(s) within this file "
                           f"(same SSN -- the repeated tear-off copies of one employee's W-2)")

            individual_path = (individual_dir or pdf.parent) / f"{pdf.stem}_adp_w2.csv"
            write_csv(file_deduped, individual_path)
            tqdm.write(f"  {pdf.name}: -> {individual_path} ({len(file_deduped)} record(s))")
            all_records.extend(file_deduped)

    if not all_records:
        print("No records extracted from any file -- combined CSV not written.")
        return

    combined = dedupe_records(all_records)
    removed = len(all_records) - len(combined)
    if removed:
        print(f"\nRemoved {removed} duplicate record(s) across files (same Document ID + SSN)")

    write_csv(combined, Path(args.output))
    print(f"\nWrote combined file: {len(combined)} record(s) -> {args.output}")
    print("Reminder: this CSV holds SSNs and home addresses in clear text -- store it in the appropriate "
          "access-controlled Global Insider folder and do not commit it to source control.")


if __name__ == "__main__":
    try:
        main()
    finally:
        if LAUNCHED_INTERACTIVELY:
            try:
                input("\nPress Enter to close...")
            except EOFError:
                pass
