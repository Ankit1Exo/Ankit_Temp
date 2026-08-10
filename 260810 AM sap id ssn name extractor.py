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
        a second reading of any page that yields nothing, using the other
        text engine, so an engine-specific quirk costs a page and not a file
        a per-file count in the console and a reconciliation sheet, so a
        file that returns zero rows is visible instead of silent

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

    python "260810 AM sap id ssn name extractor.py" --selftest
        Runs the parser over tests/*.pdf and prints row counts. No PII.

    python "260810 AM sap id ssn name extractor.py" --debug <pdf> [page]
        Masked dump of one page's rebuilt lines, with coordinates, so a
        layout problem can be pasted into a ticket without exposing PII.

REQUIREMENTS
    pip install pdfplumber pymupdf pandas openpyxl

    pdfplumber is required. pymupdf is optional but recommended -- it is
    only used as the second opinion on a page that reads as empty.

SECURITY NOTE
    The workbook this produces holds SSNs and names in clear text. Run
    only on an authorised workstation, save the XLSX to the approved
    Global Insider folder -- never a desktop, a local temp path or a
    shared drive -- and delete the local copy once it has been loaded into
    the authorised system of record. The console output and the --debug
    dump are masked and safe; the workbook is not.
"""
from __future__ import annotations

import re
import sys
import threading
import time
import traceback
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
    """Positioned words from pdfplumber. This is the engine that reads all
    58 files, so it is the primary."""
    out = []
    for w in page.extract_words(use_text_flow=False, keep_blank_chars=False):
        out.append((w["x0"], w["top"], w["bottom"], fold(w["text"])))
    return out


def words_pymupdf(page):
    """Positioned words from PyMuPDF. Second opinion only."""
    return [(w[0], w[1], w[3], fold(w[4])) for w in page.get_text("words")]


def page_lines_pdfplumber(path: Path):
    """Yield (page_number, lines) for every page, via pdfplumber."""
    with pdfplumber.open(str(path)) as pdf:
        for number, page in enumerate(pdf.pages, start=1):
            yield number, cluster_rows(words_pdfplumber(page))


def page_lines_pymupdf(path: Path, only_pages=None):
    """Yield (page_number, lines) via PyMuPDF, optionally for a subset."""
    doc = fitz.open(str(path))
    try:
        if doc.needs_pass:
            return
        for number, page in enumerate(doc, start=1):
            if only_pages is not None and number not in only_pages:
                continue
            yield number, cluster_rows(words_pymupdf(page))
    finally:
        doc.close()


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


def process_pdf(path: Path):
    """Return (rows, diagnostics) for one PDF.

    A page that yields nothing is read a SECOND time with the other engine
    before it is written off. Doing this per page rather than per file means
    an engine-specific quirk costs one page, not the whole document."""
    rows, empty_pages, page_count = [], [], 0
    engine_used = "pdfplumber"

    if pdfplumber is not None:
        for page_num, lines in page_lines_pdfplumber(path):
            page_count += 1
            found = rows_from_lines(lines, page_num, path.name)
            if not found:
                empty_pages.append(page_num)
            rows.extend(found)
    elif fitz is not None:
        engine_used = "PyMuPDF"
        for page_num, lines in page_lines_pymupdf(path):
            page_count += 1
            found = rows_from_lines(lines, page_num, path.name)
            if not found:
                empty_pages.append(page_num)
            rows.extend(found)
    else:
        raise RuntimeError("neither pdfplumber nor pymupdf is installed")

    # Second opinion on the pages that came back empty.
    recovered_rows, recovered_pages = 0, set()
    if empty_pages and pdfplumber is not None and fitz is not None:
        for page_num, lines in page_lines_pymupdf(path, only_pages=set(empty_pages)):
            found = rows_from_lines(lines, page_num, path.name)
            if not found:
                continue
            for row in found:
                note = "read with the PyMuPDF engine after pdfplumber found nothing on this page"
                row["Extraction Notes"] = (
                    f"{row['Extraction Notes']}; {note}" if row["Extraction Notes"] else note)
            recovered_rows += len(found)
            recovered_pages.add(page_num)
            rows.extend(found)

    # Most pages in this report are course detail with no student line on
    # them at all, so a page with no rows is normal. The number is here to
    # be compared ACROSS files: a file where it equals the page count is a
    # file that read as empty, and that is the failure being guarded against.
    diagnostics = {
        "File Name": path.name,
        "Pages": page_count,
        "Rows Found": len(rows),
        "Pages With No Rows": len(set(empty_pages) - recovered_pages),
        "Rows Recovered By Second Engine": recovered_rows,
        "Primary Engine": engine_used,
    }
    return rows, diagnostics


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


def debug_page(path: Path, page_no: int) -> None:
    """Masked dump of one page's rebuilt lines, with coordinates."""
    if pdfplumber is None:
        print("pdfplumber is required for --debug.  pip install pdfplumber")
        return
    with pdfplumber.open(str(path)) as pdf:
        if not 1 <= page_no <= len(pdf.pages):
            print(f"page {page_no} out of range (1..{len(pdf.pages)})")
            return
        page = pdf.pages[page_no - 1]
        print(f"{path.name}  page {page_no} of {len(pdf.pages)}  "
              f"size={page.width:.0f}x{page.height:.0f}  "
              f"rotation={getattr(page, 'rotation', 0)}")
        print("-" * 78)
        lines = cluster_rows(words_pdfplumber(page))
        for i, text in enumerate(lines, 1):
            row = parse_line(text)
            flag = "STUDENT" if row else "       "
            print(f"{i:4} {flag} | {mask(text)[:160]}")
        print("-" * 78)
        print(f"{sum(1 for t in lines if parse_line(t))} student line(s) on this page")


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
    print(f"{'file':<48}{'rows':>6}{'recovered':>11}")
    print("-" * 66)
    total = 0
    for p in pdfs:
        rows, diag = process_pdf(p)
        total += len(rows)
        print(f"{p.name[:47]:<48}{len(rows):>6}{diag['Rows Recovered By Second Engine']:>11}")
        for r in rows:
            print(f"      ID={r['ID']:<10} SSN={r['SSN']:<14} "
                  f"PFX={r['Prefix']:<5} NAME={r['Full Name']}")
    print("-" * 66)
    print(f"{'TOTAL':<48}{total:>6}")


# ===========================================================================
# workbook
# ===========================================================================
def write_workbook(rows, diagnostics, dest: Path) -> Path:
    out = dest / OUTPUT_XLSX_NAME
    frame = pd.DataFrame(rows, columns=OUTPUT_COLUMNS) if rows else \
        pd.DataFrame(columns=OUTPUT_COLUMNS)
    diag_frame = pd.DataFrame(diagnostics)

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        frame.to_excel(writer, sheet_name="Students", index=False)
        diag_frame.to_excel(writer, sheet_name="Reconciliation", index=False)
    return out


# ===========================================================================
# GUI
# ===========================================================================
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("SAP Audit  --  ID / SSN / Name Extractor")
        self.root.geometry("640x300")
        self.src = tk.StringVar()
        self.dst = tk.StringVar()
        self.status = tk.StringVar(value="Choose a source folder and a destination folder.")
        self.eta = tk.StringVar(value="")
        self._build()

    def _build(self):
        pad = {"padx": 10, "pady": 4}

        tk.Label(self.root, text="Source folder (PDFs)", anchor="w").pack(fill="x", **pad)
        row = tk.Frame(self.root); row.pack(fill="x", **pad)
        tk.Entry(row, textvariable=self.src).pack(side="left", fill="x", expand=True)
        tk.Button(row, text="Browse", command=self._pick_src).pack(side="left", padx=6)

        tk.Label(self.root, text="Destination folder (Excel output)",
                 anchor="w").pack(fill="x", **pad)
        row = tk.Frame(self.root); row.pack(fill="x", **pad)
        tk.Entry(row, textvariable=self.dst).pack(side="left", fill="x", expand=True)
        tk.Button(row, text="Browse", command=self._pick_dst).pack(side="left", padx=6)

        tk.Label(
            self.root,
            text=("The workbook holds SSNs and names in clear text.\n"
                  "Save it to the approved Global Insider folder only "
                  "-- never a desktop or a local temp path."),
            fg="#a00", justify="left", anchor="w",
        ).pack(fill="x", **pad)

        self.bar = ttk.Progressbar(self.root, orient="horizontal",
                                   mode="determinate", length=600)
        self.bar.pack(**pad)
        tk.Label(self.root, textvariable=self.status, anchor="w").pack(fill="x", **pad)
        tk.Label(self.root, textvariable=self.eta, anchor="w").pack(fill="x", **pad)
        self.go = tk.Button(self.root, text="Extract", command=self._start, width=18)
        self.go.pack(pady=8)

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
        src, dst = Path(self.src.get()), Path(self.dst.get())
        if not src.is_dir():
            messagebox.showerror("Source folder", "Choose a valid source folder.")
            return
        if not dst.is_dir():
            messagebox.showerror("Destination folder", "Choose a valid destination folder.")
            return
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
            all_rows, diagnostics = [], []
            start = time.time()

            for i, pdf in enumerate(pdfs, start=1):
                self._set(f"Reading {pdf.name}   ({i} of {len(pdfs)})")
                try:
                    rows, diag = process_pdf(pdf)
                except Exception as exc:                        # noqa: BLE001
                    rows = []
                    diag = {"File Name": pdf.name, "Pages": 0, "Rows Found": 0,
                            "Pages With No Rows": 0,
                            "Rows Recovered By Second Engine": 0,
                            "Primary Engine": f"FAILED: {type(exc).__name__}: {exc}"}
                all_rows.extend(rows)
                diagnostics.append(diag)
                print(f"  {pdf.name}: {len(rows)} student row(s)")

                self.bar["value"] = i
                elapsed = time.time() - start
                remaining = (elapsed / i) * (len(pdfs) - i)
                self._set(f"Reading {pdf.name}   ({i} of {len(pdfs)})",
                          f"Elapsed {elapsed/60:.1f} min   Remaining ~{remaining/60:.1f} min")

            self._set("Writing the workbook...")
            out = write_workbook(all_rows, diagnostics, dst)

            zero = [d["File Name"] for d in diagnostics if d["Rows Found"] == 0]
            msg = (f"{len(all_rows)} student rows from {len(pdfs)} PDFs.\n\n"
                   f"Saved to:\n{out}\n\n")
            if zero:
                msg += (f"{len(zero)} file(s) produced NO rows -- see the "
                        f"Reconciliation sheet:\n  " + "\n  ".join(zero[:10]))
                if len(zero) > 10:
                    msg += f"\n  ... and {len(zero) - 10} more"
            else:
                msg += "Every file produced at least one row."
            msg += ("\n\nThis workbook holds SSNs in clear text. Move it to the "
                    "approved Global Insider folder and delete any local copy.")

            self._set(f"Done -- {len(all_rows)} rows from {len(pdfs)} PDFs.")
            messagebox.showinfo("Extraction complete", msg)
        except Exception:                                       # noqa: BLE001
            traceback.print_exc()
            self._set("Failed -- see the console.")
            messagebox.showerror("Failed", traceback.format_exc(limit=3))
        finally:
            self.go.config(state="normal")

    def run(self):
        self.root.mainloop()


# ===========================================================================
# entry
# ===========================================================================
def main() -> None:
    args = sys.argv[1:]

    if pdfplumber is None and fitz is None:
        print("Install the PDF engines first:  pip install pdfplumber pymupdf")
        return

    if args and args[0] == "--selftest":
        selftest()
        return

    if args and args[0] == "--debug":
        if len(args) < 2:
            print('usage: --debug "<file.pdf>" [page]')
            return
        debug_page(Path(args[1]), int(args[2]) if len(args) > 2 else 1)
        return

    if pdfplumber is None:
        print("WARNING: pdfplumber is not installed, so the PyMuPDF engine is in use.\n"
              "         pdfplumber is the engine that reads these reports reliably.\n"
              "         Install it with:  pip install pdfplumber")

    App().run()


if __name__ == "__main__":
    main()
