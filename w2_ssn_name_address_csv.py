"""
Extract Employee First Name, MI, Last Name, Address, and SSN from Form W-2
PDFs and write a single CSV: First Name, MI, Last Name, Address, SSN.

Layout handling (same approach as extract_w2.py):
  - SSN is read from the "a Employee's SSN" / "a Employee's social security
    number" box and validated/normalized to XXX-XX-XXXX.
  - Name is read from box "e" ("Employee's first name and initial / Last
    name"); when that caption's own line carries a "Last name" sub-label,
    its x-position is used to split the name value below into First/MI vs
    Last by column position (so a multi-word last name like "VEGA TAPIA"
    isn't mistaken for a first/last boundary). Address is read from box "f"
    ("Employee's address and zip code"), which prints below box e as its own
    caption line -- that line is skipped over rather than treated as a stop
    label, so the street/city/state/zip lines after it still get collected.
  - A page can hold 1, 2, or 4 employees (side by side, stacked, or a 2x2
    grid) -- however many "SSN" captions land on each distinct row band
    decides the split, so no layout has to be assumed up front.
  - Some PDFs also carry a second, unrelated page format: a state payroll/
    DE9C-style continuation page listing many employees per page (SSN +
    name, then a wage-only line) instead of the W-2 box grid. Its own
    caption text is printed on some documents and omitted on others, so
    these pages are recognized by the repeating shape of the record
    (see detect_format2()), and a record from one is only kept if its SSN
    wasn't already found on a W-2 page in the same document -- this format
    carries no address, so a kept record's Address is blank.

These PDFs are OCR'd (made searchable) via ABBYY rather than carrying native
text, so two extra tolerances are built in:
  - normalize_text() folds curly quotes/dashes and stray whitespace ABBYY
    tends to introduce, and the caption regexes allow an optional (rather
    than required) apostrophe character so a dropped apostrophe ("Employees
    SSN") still matches.
  - find_ssn() falls back to correcting single-character digit/letter OCR
    confusions (O<->0, I/l<->1, S<->5, etc.) within the SSN value window if
    the strict digit-only pattern finds nothing there.

Usage:
    python w2_ssn_name_address_csv.py <pdf_file_or_folder> [-o output.csv] [--split-fraction 0.5]

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
COLUMN_SPLIT_FRACTION = 0.5

CSV_COLUMNS = ["Document ID", "Page", "First Name", "MI", "Last Name", "Address", "SSN"]

# ".?" (not ".") for the apostrophe so a dropped/misread apostrophe -- common
# when ABBYY OCRs a scanned page -- still matches ("Employees SSN").
SSN_CAPTION_RE = re.compile(r"Employee.?s\s+(?:social security number|SSN)\b", re.IGNORECASE)
NAME_CAPTION_RE = re.compile(r"Employee.?s\s+(?:first name and initial|name,\s*address)", re.IGNORECASE)

STOP_LABEL_RE = re.compile(
    r"^(?:\d{1,2}\s+State\b|Employer.s state ID|Local income tax|Locality name|"
    r"Form\s*W-?2\b|Wage\s*(?:&|and)\s*Tax Statement|Copy\s+[A-Z0-9]|"
    r"For (?:Official|Privacy)|Department of the Treasury|20\d{2}$)",
    re.IGNORECASE,
)

# Box "f" ("Employee's address and zip code") prints as its own caption line
# directly below box e, with the actual street/city/state/zip lines coming
# after it -- skip over this caption rather than treating it as a stop
# label (stopping there would end the block right before the address
# content it introduces).
ADDRESS_CAPTION_RE = re.compile(r"^f\s+Employee.?s\s+address", re.IGNORECASE)

# Sub-labels that can appear on box e's own caption line alongside
# "Employee's first name and initial" -- their x-position anchors the
# column split of the name value line below into First/MI vs Last.
LAST_NAME_LABEL_RE = re.compile(r"^last$", re.IGNORECASE)
SUFFIX_LABEL_RE = re.compile(r"^suff", re.IGNORECASE)

# Accepts dash, space, or no separator (9 digits run together) so a valid
# SSN isn't missed just because the PDF's text layer dropped the dashes;
# normalize_ssn() below reformats whatever is found to XXX-XX-XXXX.
SSN_VALUE_RE = re.compile(r"\b(\d{3})[-\s]?(\d{2})[-\s]?(\d{4})\b")

CITY_STATE_ZIP_RE = re.compile(r"^(?P<city>.+?),?\s+(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(?:-\d{4})?)\b")

# "b Employer ID no. (EIN)" value, e.g. "12-3456789" -- 2-7 grouping, distinct
# from the SSN's 3-2-4 grouping. Used only as an anchor in the shape-based
# fallback below, to tell the employer's address block (box c, which always
# prints directly after the EIN) apart from the employee's (box e, which
# always prints later, after box d).
EIN_VALUE_RE = re.compile(r"\b\d{2}-\d{7}\b")

# A name/address block starting within this many lines of the EIN is almost
# certainly box c (Employer's name/address) riding right below box b, not
# box e (Employee's) -- box e always comes later, after box c and box d.
EMPLOYER_BLOCK_MAX_GAP_LINES = 3

# Letters ABBYY commonly confuses with digits when the source image is a bit
# noisy (thin strokes, low DPI, etc.). Only applied as a fallback within the
# few lines around an already-found SSN caption -- not to the whole page --
# so it can't turn unrelated text elsewhere into a false SSN match.
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

# Format 2: a state payroll/DE9C-style continuation page that lists many
# employees (SSN + name, then a wage-only line) instead of the W-2 box
# grid. Its own caption text ("D. SOCIAL SECURITY NUMBER", "E. EMPLOYEE
# NAME ...") is printed per record on some documents and omitted on
# others, so pages are recognized by the repeating shape of the record
# (see detect_format2()) rather than by that caption being present.
FORMAT2_NAME_CAPTION_RE = re.compile(r"E\.\s*EMPLOYEE\s+NAME", re.IGNORECASE)

# A line that's nothing but two or more space-separated digit groups, with
# no letters and no comma/decimal currency formatting -- this is what a
# Format-2 wage line looks like ("134 133 11"), and what tells a Format-2
# record apart from a W-2 box's dollar amounts even without any caption.
WAGE_ONLY_LINE_RE = re.compile(r"[\d ]+")


def mask_shape(s):
    return re.sub(r"[A-Za-z]", "X", re.sub(r"\d", "#", s))


def normalize_text(s):
    """Fold ABBYY's curly quotes/dashes and odd whitespace to plain ASCII
    equivalents before any regex runs against OCR'd text."""
    s = unicodedata.normalize("NFKC", s)
    s = s.translate({0x2018: "'", 0x2019: "'", 0x00B4: "'", 0x0060: "'",
                      0x201C: '"', 0x201D: '"', 0x00A0: " ",
                      0x2013: "-", 0x2014: "-"})
    return re.sub(r"[ \t]+", " ", s).strip()


def normalize_ssn(match):
    return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"


def group_words_into_lines(words, y_tol=3):
    lines = {}
    for w in words:
        x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
        key = round(y0 / y_tol) * y_tol
        lines.setdefault(key, []).append((x0, x1, text))
    return [sorted(v, key=lambda t: t[0]) for _, v in sorted(lines.items())]


def find_name_column_bounds(caption_words):
    """Locate the x-position of box e's 'Last name' (and 'Suff.', if
    present) sub-captions on its own caption line, so the name value line
    below it can be split into First/MI vs Last by column position --
    rather than guessing from whitespace in the value itself, which can't
    tell a two-word last name ("VEGA TAPIA") from a first/last column gap."""
    last_x = None
    suffix_x = None
    for idx, (x0, x1, text) in enumerate(caption_words):
        token = text.strip(".,")
        if LAST_NAME_LABEL_RE.match(token) and idx + 1 < len(caption_words) \
                and caption_words[idx + 1][2].lower().startswith("name"):
            last_x = x0
        elif SUFFIX_LABEL_RE.match(token):
            suffix_x = x0
    return last_x, suffix_x


def split_name_by_columns(value_words, last_x, suffix_x):
    first_words, last_words = [], []
    for x0, x1, text in value_words:
        center = (x0 + x1) / 2
        if center < last_x:
            first_words.append(text)
        elif suffix_x is None or center < suffix_x:
            last_words.append(text)
    return " ".join(first_words), " ".join(last_words)


def split_first_mi(first_region):
    tokens = first_region.split()
    if len(tokens) >= 2 and re.fullmatch(r"[A-Za-z]\.?", tokens[-1]):
        return " ".join(tokens[:-1]), tokens[-1].rstrip(".")
    return first_region, ""


def split_name_fallback(name_line):
    """Used when there's no column position to split on (box e's caption has
    no separate 'Last name' sub-label, or this is a Format-2 record with no
    'E. EMPLOYEE NAME' caption above it) -- peels a trailing single-letter
    token off as a middle initial, then treats the last remaining token as
    the last name. Imperfect for multi-word last names ("VEGA TAPIA", "DE LA
    CRUZ") since there's no column position to go by; callers log a warning
    so the record can be spot-checked."""
    tokens = name_line.split()
    if not tokens:
        return "", "", ""
    if len(tokens) >= 3 and re.fullmatch(r"[A-Za-z]\.?", tokens[-2]):
        return " ".join(tokens[:-2]), tokens[-2].rstrip("."), tokens[-1]
    if len(tokens) >= 2:
        return " ".join(tokens[:-1]), "", tokens[-1]
    return tokens[0], "", ""


def find_ssn(plain_lines, page_num, column_label):
    for i, line in enumerate(plain_lines):
        if SSN_CAPTION_RE.search(line):
            lo, hi = max(0, i - 2), min(len(plain_lines), i + 3)
            window_lines = plain_lines[lo:hi]
            window_text = " ".join(window_lines)
            m = SSN_VALUE_RE.search(window_text)
            if m:
                return normalize_ssn(m)
            # Strict digit match failed -- retry after mapping OCR digit-lookalike
            # letters (O/I/S/etc.) back to digits, in case ABBYY misread one.
            m = SSN_VALUE_RE.search(window_text.translate(OCR_DIGIT_FIX))
            if m:
                tqdm.write(f"  page {page_num} ({column_label}): SSN recovered via OCR digit-lookalike "
                      f"correction (a letter ABBYY likely misread as a digit was substituted back)")
                return normalize_ssn(m)
            shape = " | ".join(l if SSN_CAPTION_RE.search(l) else mask_shape(l) for l in window_lines)
            tqdm.write(f"  page {page_num} ({column_label}): SSN caption found but no SSN-shaped value nearby "
                  f"-- nearby line shapes (digits/letters masked, safe to share): {shape}")
            return ""
    tqdm.write(f"  page {page_num} ({column_label}): SSN caption not found")
    return ""


def find_name_address(plain_lines, lines, page_num, column_label):
    for i, line in enumerate(plain_lines):
        if NAME_CAPTION_RE.search(line):
            last_x, suffix_x = find_name_column_bounds(lines[i])
            raw_idx = list(range(i + 1, min(i + 9, len(plain_lines))))
            block, block_idx = [], []
            for idx in raw_idx:
                candidate = plain_lines[idx].strip()
                if not candidate:
                    continue
                if ADDRESS_CAPTION_RE.search(candidate):
                    continue
                if STOP_LABEL_RE.search(candidate):
                    break
                block.append(candidate)
                block_idx.append(idx)
            if not block:
                shape = " | ".join(mask_shape(plain_lines[idx]) for idx in raw_idx)
                tqdm.write(f"  page {page_num} ({column_label}): name/address caption found but block below it is "
                      f"empty (line shapes below caption, safe to share): {shape}")
                return "", "", "", ""
            if last_x is not None:
                first_region, last = split_name_by_columns(lines[block_idx[0]], last_x, suffix_x)
                first, mi = split_first_mi(first_region)
            else:
                tqdm.write(f"  page {page_num} ({column_label}): box e caption has no separate 'Last name' "
                      f"sub-label to anchor a column split on -- falling back to a whitespace-based guess for "
                      f"first/MI/last (may mis-split multi-word last names); spot-check this record")
                first, mi, last = split_name_fallback(block[0])
            city_idx = next((idx for idx, l in enumerate(block) if idx > 0 and CITY_STATE_ZIP_RE.match(l)), None)
            if city_idx is not None:
                m = CITY_STATE_ZIP_RE.match(block[city_idx])
                city, state, zip_code = m.group("city").rstrip(","), m.group("state"), m.group("zip")
                street = " ".join(block[1:city_idx])
                address = f"{street}, {city}, {state} {zip_code}" if street else f"{city}, {state} {zip_code}"
            else:
                street = " ".join(block[1:])
                address = street
                if len(block) > 1:
                    shape = " | ".join(mask_shape(l) for l in block[1:])
                    tqdm.write(f"  page {page_num} ({column_label}): could not find a city/state/zip-shaped line "
                          f"in the address block -- line shapes (safe to share): {shape}")
            return first, mi, last, address
    tqdm.write(f"  page {page_num} ({column_label}): name/address caption not found")
    return "", "", "", ""


def extract_employee(plain_lines, lines, page_num, column_label):
    ssn = find_ssn(plain_lines, page_num, column_label)
    first, mi, last, address = find_name_address(plain_lines, lines, page_num, column_label)
    return {"Page": page_num, "First Name": first, "MI": mi, "Last Name": last, "Address": address, "SSN": ssn}


def is_letterish(line):
    """True for lines that could be part of a name/street -- used to tell a
    wage-amount/EIN/SSN line (zero real letters) apart from an address line
    when there's no caption to anchor on. Deliberately NOT "letters outnumber
    digits": a real street line like "123 N 45TH ST." has as many digits as
    letters (house number, direction, ordinal) and would otherwise be
    mistaken for a boundary line, collapsing the whole address block."""
    letters = sum(c.isalpha() for c in line)
    return letters >= 2


def find_ssn_by_shape(plain_lines, page_num, column_label):
    """Caption-less fallback: find an SSN-shaped value directly, without an
    'Employee's SSN' caption to anchor on (used when that caption isn't in
    the OCR text at all -- e.g. it's part of a background template image
    ABBYY didn't recognize as text, while the payroll-filled values are)."""
    for line in plain_lines:
        m = SSN_VALUE_RE.search(line)
        if m:
            return normalize_ssn(m)
        m = SSN_VALUE_RE.search(line.translate(OCR_DIGIT_FIX))
        if m:
            tqdm.write(f"  page {page_num} ({column_label}): [shape-fallback] SSN recovered via OCR "
                  f"digit-lookalike correction")
            return normalize_ssn(m)
    tqdm.write(f"  page {page_num} ({column_label}): [shape-fallback] no SSN-shaped value found in this cell")
    return ""


def find_name_address_by_shape(plain_lines, page_num, column_label):
    """Caption-less fallback for name/address. A W-2 box stack always prints
    box c (Employer's name/address/ZIP) directly under box b (EIN), then box
    d (Control number), then box e (Employee's name/address/ZIP) -- so when
    there's no caption to anchor on, collect every name+street+city/state/zip
    -shaped block and use the EIN's position to tell an employer-adjacent
    block apart from the employee's, rather than assuming "the only block
    found" or "the last block" is automatically the employee's (either one
    can silently return the employer's info instead if the employee block
    wasn't detected for some reason)."""
    csz_indices = [i for i, l in enumerate(plain_lines) if CITY_STATE_ZIP_RE.match(l)]
    if not csz_indices:
        tqdm.write(f"  page {page_num} ({column_label}): [shape-fallback] no city/state/zip-shaped line found "
              f"-- cannot locate name/address")
        return "", "", "", ""

    blocks = []
    for csz_idx in csz_indices:
        block = [csz_idx]
        j = csz_idx - 1
        noise_skipped = 0
        while j >= 0 and len(block) < 4:
            line = plain_lines[j].strip()
            if not line or CITY_STATE_ZIP_RE.match(line):
                break
            if not is_letterish(line):
                # A stray 1-2 char token (checkbox mark, OCR artifact) between
                # the street and city/state/zip line -- e.g. a lone "X" --
                # shouldn't be treated as a box boundary. Skip a couple of
                # these without giving up the scan; anything longer is more
                # likely a real non-address line (wage amount, etc.) and IS
                # a genuine boundary.
                if len(line) <= 2 and noise_skipped < 2:
                    noise_skipped += 1
                    j -= 1
                    continue
                break
            block.insert(0, j)
            j -= 1
        if len(block) >= 2:
            blocks.append(block)

    if not blocks:
        tqdm.write(f"  page {page_num} ({column_label}): [shape-fallback] city/state/zip-shaped line(s) found "
              f"but no preceding name/street line(s)")
        return "", "", "", ""

    ein_idx = next((i for i, l in enumerate(plain_lines) if EIN_VALUE_RE.search(l)), None)

    if ein_idx is None:
        # No EIN anchor to disambiguate with. Only safe to act when there's
        # more than one block (last = employee, per box order); a lone block
        # is genuinely ambiguous -- guessing it's the employee's is exactly
        # the bug this replaced (it can just as easily be the employer's).
        if len(blocks) > 1:
            tqdm.write(f"  page {page_num} ({column_label}): [shape-fallback] no EIN found to anchor on, but "
                  f"{len(blocks)} name/address blocks were -- using the last one (box 'e' prints after box "
                  f"'c') as the employee's; spot-check this record")
            chosen = blocks[-1]
        else:
            tqdm.write(f"  page {page_num} ({column_label}): [shape-fallback] only one name/address block found "
                  f"and no EIN to confirm whether it's box 'c' (employer) or box 'e' (employee) -- skipping "
                  f"rather than risk recording the employer's info as the employee's")
            return "", "", "", ""
    else:
        employee_like = [b for b in blocks if b[0] - ein_idx > EMPLOYER_BLOCK_MAX_GAP_LINES]
        if not employee_like:
            tqdm.write(f"  page {page_num} ({column_label}): [shape-fallback] every name/address block found sits "
                  f"within {EMPLOYER_BLOCK_MAX_GAP_LINES} lines of the EIN -- that's box 'c' (employer's), "
                  f"not box 'e'; employee name/address not found in this cell")
            return "", "", "", ""
        if len(employee_like) > 1:
            tqdm.write(f"  page {page_num} ({column_label}): [shape-fallback] {len(employee_like)} candidate "
                  f"employee blocks found past the EIN -- using the last one; spot-check this record")
        chosen = employee_like[-1]

    name = plain_lines[chosen[0]].strip()
    first, mi, last = split_name_fallback(name)
    csz_line = plain_lines[chosen[-1]].strip()
    m = CITY_STATE_ZIP_RE.match(csz_line)
    city, state, zip_code = m.group("city").rstrip(","), m.group("state"), m.group("zip")
    street = " ".join(plain_lines[i].strip() for i in chosen[1:-1])
    address = f"{street}, {city}, {state} {zip_code}" if street else f"{city}, {state} {zip_code}"
    return first, mi, last, address


def extract_employee_by_shape(plain_lines, page_num, column_label):
    ssn = find_ssn_by_shape(plain_lines, page_num, column_label)
    first, mi, last, address = find_name_address_by_shape(plain_lines, page_num, column_label)
    return {"Page": page_num, "First Name": first, "MI": mi, "Last Name": last, "Address": address, "SSN": ssn}


def caption_line_groups(words, caption_re, y_tol=3):
    groups = {}
    for w in words:
        x0, y0, _, _, text = w[0], w[1], w[2], w[3], w[4]
        key = round(y0 / y_tol) * y_tol
        groups.setdefault(key, []).append((x0, text))
    result = []
    for key, items in sorted(groups.items()):
        joined = normalize_text(" ".join(t for _, t in sorted(items, key=lambda t: t[0])))
        count = len(caption_re.findall(joined))
        if count:
            result.append((key, count))
    return result


def build_grid_cells(page, words, groups, split_fraction, page_num):
    row_ys = [y for y, _ in groups]
    row_counts = [c for _, c in groups]
    num_rows = len(row_ys)

    if num_rows <= 1:
        row_bounds = [(float("-inf"), float("inf"))]
    else:
        splits = [(row_ys[i] + row_ys[i + 1]) / 2 for i in range(num_rows - 1)]
        edges = [float("-inf")] + splits + [float("inf")]
        row_bounds = [(edges[i], edges[i + 1]) for i in range(num_rows)]

    if num_rows == 1:
        row_label = lambda idx: ""
    elif num_rows == 2:
        row_label = lambda idx: ("Top", "Bottom")[idx]
    else:
        row_label = lambda idx: f"Row{idx + 1}"

    split_x = page.rect.width * split_fraction
    cells = []
    for idx, (y_lo, y_hi) in enumerate(row_bounds):
        row_words = words if num_rows <= 1 else [w for w in words if y_lo <= (w[1] + w[3]) / 2 < y_hi]
        count = row_counts[idx] if idx < len(row_counts) else 1
        prefix = row_label(idx)

        if count >= 2:
            if count > 2:
                tqdm.write(f"  page {page_num}: row {idx + 1} has {count} SSN captions (expected 1 or 2) -- only "
                      f"the first 2 columns will be split out")
            cells.append((prefix + "Left", [w for w in row_words if (w[0] + w[2]) / 2 < split_x]))
            cells.append((prefix + "Right", [w for w in row_words if (w[0] + w[2]) / 2 >= split_x]))
        else:
            cells.append((prefix if prefix else "Single", row_words))

    return cells


def is_format2_wage_line(line):
    """A Format-2 wage line is nothing but two or more space-separated
    digit groups -- no letters, no comma/decimal currency formatting (that's
    what a W-2 box amount looks like instead)."""
    return bool(WAGE_ONLY_LINE_RE.fullmatch(line)) and len(line.split()) >= 2


def detect_format2(plain_lines):
    """Recognize a Format-2 (state payroll/DE9C-style list) page by the
    repeating shape of its records -- an SSN-shaped line followed shortly by
    a wage-only line -- since this format's own caption text is printed on
    some documents and omitted on others, so it can't be relied on as the
    page-level signal. Two or more such records confirm it (one match alone
    could be a coincidental digit run)."""
    hits = 0
    for i, line in enumerate(plain_lines):
        if is_format2_wage_line(line):
            continue
        if not (SSN_VALUE_RE.search(line) or SSN_VALUE_RE.search(line.translate(OCR_DIGIT_FIX))):
            continue
        if any(is_format2_wage_line(plain_lines[j]) for j in range(i + 1, min(i + 4, len(plain_lines)))):
            hits += 1
            if hits >= 2:
                return True
    return False


def split_ssn_and_name_words(line_words):
    """The SSN sits in the leading word-token(s) of a Format-2 record line
    (all digits/dashes), directly followed by the employee name -- split on
    that boundary rather than assuming the SSN is exactly one token, since
    OCR can emit it as "123-45-6789" in one token or "123 45 6789" split
    across three."""
    i = 0
    while i < len(line_words) and re.fullmatch(r"[\d\-]+", line_words[i][2]):
        i += 1
    return line_words[:i], line_words[i:]


def find_format2_name_columns(caption_words):
    """Locate the x-position of the 'M.I.' and 'Last name' sub-captions on
    box E's caption line ("E. EMPLOYEE NAME (FIRST NAME) (M.I.) (LAST
    NAME)"), when that line is present, to split the name value below it by
    column position."""
    mi_x = None
    last_x = None
    for x0, x1, text in caption_words:
        token = text.strip("().,").upper()
        if last_x is None and token.startswith("LAST"):
            last_x = x0
        elif mi_x is None and re.fullmatch(r"M\.?I\.?", token):
            mi_x = x0
    return mi_x, last_x


def split_name_by_columns3(name_words, mi_x, last_x):
    first_words, mi_words, last_words = [], [], []
    for x0, x1, text in name_words:
        center = (x0 + x1) / 2
        if last_x is not None and center >= last_x:
            last_words.append(text)
        elif mi_x is not None and center >= mi_x:
            mi_words.append(text)
        else:
            first_words.append(text)
    return " ".join(first_words), " ".join(mi_words), " ".join(last_words)


def process_format2_page(page_plain_lines, page_lines, page_num):
    """Parse a Format-2 (state payroll/DE9C-style) page: each employee is an
    SSN + name on one line, optionally preceded by a 'D. SOCIAL SECURITY
    NUMBER / E. EMPLOYEE NAME ...' caption line and followed by a wage-only
    line -- no address exists on this page for any employee, so Address is
    always blank. Records are flagged _format2 so process_pdf can merge them
    against Format-1 (W-2) records by SSN before deciding what to keep."""
    records = []
    for idx, line in enumerate(page_plain_lines):
        if is_format2_wage_line(line):
            continue
        if not (SSN_VALUE_RE.search(line) or SSN_VALUE_RE.search(line.translate(OCR_DIGIT_FIX))):
            continue

        ssn_words, name_words = split_ssn_and_name_words(page_lines[idx])
        if not ssn_words or not name_words:
            continue
        ssn_text = "".join(t for _, _, t in ssn_words)
        m = SSN_VALUE_RE.search(ssn_text) or SSN_VALUE_RE.search(ssn_text.translate(OCR_DIGIT_FIX))
        if not m:
            continue
        ssn = normalize_ssn(m)

        mi_x = last_x = None
        if idx > 0 and FORMAT2_NAME_CAPTION_RE.search(page_plain_lines[idx - 1]):
            mi_x, last_x = find_format2_name_columns(page_lines[idx - 1])
        if mi_x is not None or last_x is not None:
            first, mi, last = split_name_by_columns3(name_words, mi_x, last_x)
            if not mi:
                # The M.I. column position doesn't always cleanly separate a
                # tightly-kerned initial from the first name -- fall back to
                # peeling a trailing single-letter token off of "first" so a
                # slightly-off column boundary doesn't just drop the initial.
                first, mi = split_first_mi(first)
        else:
            tqdm.write(f"  page {page_num}: [format-2] no 'E. EMPLOYEE NAME' column caption above this record -- "
                  f"falling back to a whitespace-based guess for first/MI/last; spot-check this record")
            first, mi, last = split_name_fallback(" ".join(t for _, _, t in name_words))

        records.append({"Page": page_num, "First Name": first, "MI": mi, "Last Name": last,
                         "Address": "", "SSN": ssn, "_format2": True})

    if not records:
        tqdm.write(f"  page {page_num}: [format-2] page matched the list-format shape but no SSN+name rows could "
              f"be parsed")
    return records


def process_page(page, page_num, split_fraction):
    text = page.get_text()
    if len(text.strip()) < MIN_TEXT_CHARS_PER_PAGE:
        tqdm.write(f"  page {page_num}: only {len(text.strip())} chars of text -- scanned/image-only page, skipping "
              f"(this script expects a real text layer; OCR it first if needed)")
        return []

    words = page.get_text("words")
    page_lines = group_words_into_lines(words)
    page_plain_lines = [normalize_text(" ".join(t for _, _, t in ln)) for ln in page_lines]

    norm_text = normalize_text(text)
    caption_count = len(SSN_CAPTION_RE.findall(norm_text))

    if caption_count == 0 and detect_format2(page_plain_lines):
        tqdm.write(f"  page {page_num}: matches the Format-2 (state payroll list) record shape -- parsing as an "
              f"employee list instead of a W-2 box grid")
        return process_format2_page(page_plain_lines, page_lines, page_num)

    shape_fallback = False
    marker_re = SSN_CAPTION_RE
    marker_count = caption_count

    if caption_count == 0:
        # The caption itself isn't in the OCR text at all -- happens when the
        # box labels are part of a background template image ABBYY didn't
        # recognize as text, while the payroll-filled values on top of it
        # were. If SSN-shaped values are on the page anyway, fall back to
        # locating employees by value shape/position instead of caption text.
        value_count = len(SSN_VALUE_RE.findall(norm_text))
        if value_count == 0:
            tqdm.write(f"  page {page_num}: no 'Employee's SSN' caption and no SSN-shaped value found -- not a "
                  f"W-2 employee page, skipping")
            return []
        tqdm.write(f"  page {page_num}: no 'Employee's SSN' caption in the OCR text (likely baked into a "
              f"background image) -- falling back to locating employees by SSN value shape/position")
        shape_fallback = True
        marker_re = SSN_VALUE_RE
        marker_count = value_count

    records = []

    if marker_count == 1:
        cells = [("Single", words)]
    else:
        groups = caption_line_groups(words, marker_re)
        if not groups:
            tqdm.write(f"  page {page_num}: {marker_count} SSN {'values' if shape_fallback else 'captions'} found "
                  f"on the page but none could be grouped into row bands -- falling back to treating the "
                  f"whole page as one employee")
            cells = [("Single", words)]
        else:
            cells = build_grid_cells(page, words, groups, split_fraction, page_num)

    for column_label, column_words in cells:
        lines = group_words_into_lines(column_words)
        plain_lines = [normalize_text(" ".join(t for _, _, t in ln)) for ln in lines]

        if shape_fallback:
            has_ssn_here = any(SSN_VALUE_RE.search(l) or SSN_VALUE_RE.search(l.translate(OCR_DIGIT_FIX))
                                for l in plain_lines)
            if not has_ssn_here:
                tqdm.write(f"  page {page_num} ({column_label}): [shape-fallback] no SSN-shaped value in this "
                      f"cell -- split-fraction may be off for this file, try --split-fraction")
                continue
            records.append(extract_employee_by_shape(plain_lines, page_num, column_label))
        else:
            if not any(SSN_CAPTION_RE.search(l) for l in plain_lines):
                tqdm.write(f"  page {page_num} ({column_label}): no SSN caption on this cell -- split-fraction "
                      f"may be off for this file, try --split-fraction")
                continue
            records.append(extract_employee(plain_lines, lines, page_num, column_label))

    return records


def process_pdf(path: Path, split_fraction):
    doc = fitz.open(path)
    records = []
    for i, page in enumerate(doc, start=1):
        records.extend(process_page(page, i, split_fraction))
    doc.close()

    format1_records, format2_records = [], []
    for rec in records:
        (format2_records if rec.pop("_format2", False) else format1_records).append(rec)

    known_ssns = {r["SSN"] for r in format1_records if r.get("SSN")}
    kept_format2 = []
    skipped = 0
    for rec in format2_records:
        if rec.get("SSN") and rec["SSN"] in known_ssns:
            skipped += 1
            continue
        kept_format2.append(rec)
    if format2_records:
        tqdm.write(f"  {path.name}: {len(format2_records)} format-2 record(s) found, {skipped} already covered "
              f"by a W-2 record (same SSN) -- keeping {len(kept_format2)}")

    return format1_records + kept_format2


def dedupe_records(records):
    """Each employee typically appears on more than one identical W-2 copy
    (Copy B, Copy C, Copy 2, ...) within the same PDF, producing one record
    per copy with the same First/MI/Last Name/Address/SSN but a different
    Page -- collapse those down to a single row per employee, keeping the
    first occurrence's Page. Uniqueness is scoped to Document ID + SSN
    (falling back to Document ID + First/MI/Last Name/Address when SSN
    extraction failed) -- NOT to SSN alone across the whole combined file,
    so two different documents that happen to share an SSN/name/address are
    never collapsed into one row just because they were combined into the
    same output."""
    seen = set()
    deduped = []
    for rec in records:
        doc_id = rec.get("Document ID", "")
        ssn = rec.get("SSN", "")
        first = rec.get("First Name", "")
        mi = rec.get("MI", "")
        last = rec.get("Last Name", "")
        address = rec.get("Address", "")
        if ssn:
            key = (doc_id, "ssn", ssn)
        elif first or last or address:
            key = (doc_id, "name_addr", first, mi, last, address)
        else:
            # Nothing usable was extracted -- keep every such record rather
            # than collapsing unrelated blank rows into one.
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


def debug_page(path: Path, page_num: int):
    """Print a masked, line-by-line view of one page's OCR text layer so a
    failed extraction can be diagnosed without exposing PII: caption lines
    print in full (they're just form labels), every other line is reduced to
    its digit/letter shape via mask_shape() -- safe to paste back here."""
    doc = fitz.open(path)
    if page_num < 1 or page_num > len(doc):
        print(f"{path.name}: page {page_num} out of range (document has {len(doc)} page(s))")
        doc.close()
        return
    page = doc[page_num - 1]

    text = normalize_text(page.get_text())
    print(f"--- {path.name} page {page_num} ---")
    print(f"text layer: {len(text.strip())} chars "
          f"({'OK' if len(text.strip()) >= MIN_TEXT_CHARS_PER_PAGE else 'BELOW MIN_TEXT_CHARS_PER_PAGE -- treated as image-only'})")
    ssn_caption_count = len(SSN_CAPTION_RE.findall(text))
    name_caption_count = len(NAME_CAPTION_RE.findall(text))
    ssn_value_count = len(SSN_VALUE_RE.findall(text))
    print(f"SSN captions found on page: {ssn_caption_count}")
    print(f"Name/address captions found on page: {name_caption_count}")
    print(f"SSN-shaped values found on page: {ssn_value_count}")
    if ssn_caption_count == 0 and ssn_value_count > 0:
        print("  -> no caption text in the OCR layer; extraction will use the shape-based fallback "
              "(SSN located by value shape, name/address by position -- box 'e' assumed to be the LAST "
              "name/street/city-state-zip-shaped block in each cell)")

    words = page.get_text("words")
    lines = group_words_into_lines(words)
    plain_lines = [normalize_text(" ".join(t for _, _, t in ln)) for ln in lines]
    print(f"lines detected: {len(plain_lines)}")
    for i, line in enumerate(plain_lines):
        if SSN_CAPTION_RE.search(line):
            tag = "  <-- SSN caption"
            shown = line
        elif NAME_CAPTION_RE.search(line):
            tag = "  <-- name/address caption"
            shown = line
        elif STOP_LABEL_RE.search(line):
            tag = "  <-- stop label"
            shown = line
        else:
            tag = ""
            shown = mask_shape(line)
        print(f"  [{i:>3}] {shown}{tag}")
    doc.close()


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("input", help="PDF file or folder of PDFs")
    parser.add_argument("-o", "--output", default="w2_extracted.csv",
                         help="Combined (all files) output CSV path (default: w2_extracted.csv)")
    parser.add_argument("--individual-dir", default=None,
                         help="Directory for the individual per-PDF CSVs, named <pdf_stem>_extracted.csv "
                              "(default: same folder as each input PDF)")
    parser.add_argument("--split-fraction", type=float, default=COLUMN_SPLIT_FRACTION,
                         help="Fraction of page width where a two-up page is cut into left/right halves "
                              f"(default {COLUMN_SPLIT_FRACTION})")
    parser.add_argument("--debug", action="store_true",
                         help="Print a masked line-by-line layout of one page per input file and exit -- "
                              "caption lines print in full, everything else is masked to digit/letter shape "
                              "(safe to paste back for troubleshooting). Use --debug-page to pick the page.")
    parser.add_argument("--debug-page", type=int, default=1, metavar="N",
                         help="Page number to debug when --debug is set (default: 1)")
    args = parser.parse_args()

    input_path = Path(args.input)
    pdf_files = [input_path] if input_path.is_file() else sorted(input_path.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found at: {input_path}")
        return

    if args.debug:
        for pdf in pdf_files:
            debug_page(pdf, args.debug_page)
        return

    individual_dir = Path(args.individual_dir) if args.individual_dir else None
    if individual_dir:
        individual_dir.mkdir(parents=True, exist_ok=True)

    all_records = []
    with tqdm(pdf_files, desc="Processing PDFs", unit="pdf") as pbar:
        for pdf in pbar:
            pbar.set_postfix_str(pdf.name)
            records = process_pdf(pdf, args.split_fraction)
            for rec in records:
                rec["Document ID"] = pdf.stem

            if not records:
                tqdm.write(f"  {pdf.name}: no records extracted")
                continue

            file_deduped = dedupe_records(records)
            file_removed = len(records) - len(file_deduped)
            if file_removed:
                tqdm.write(f"  {pdf.name}: removed {file_removed} duplicate record(s) within this file")

            individual_path = (individual_dir or pdf.parent) / f"{pdf.stem}_extracted.csv"
            write_csv(file_deduped, individual_path)
            tqdm.write(f"  {pdf.name}: -> {individual_path} ({len(file_deduped)} record(s))")

            all_records.extend(file_deduped)

    if not all_records:
        print("No records extracted from any file -- combined CSV not written.")
        return

    combined_deduped = dedupe_records(all_records)
    combined_removed = len(all_records) - len(combined_deduped)
    if combined_removed:
        print(f"\nRemoved {combined_removed} duplicate record(s) across files (same Document ID + SSN)")

    output_path = Path(args.output)
    write_csv(combined_deduped, output_path)
    print(f"\nWrote combined file: {len(combined_deduped)} record(s) -> {output_path}")


if __name__ == "__main__":
    main()
