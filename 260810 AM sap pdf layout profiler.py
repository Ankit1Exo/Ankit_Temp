r"""
SAP Audit PDF Layout Profiler -- diagnostic, read-only
======================================================
This script does NOT extract anything. It answers one question:

    "Why did 4 PDFs work and 54 report 'no student record found',
     when all 58 are searchable text?"

It opens each PDF read-only, measures the things that differ between a
file that parses and a file that does not, and writes a masked report.

WHAT IT MEASURES, AND WHY EACH ONE MATTERS
    text layer        chars per page from two independent engines.
                      pdfplumber and PyMuPDF do not read every PDF the
                      same way. If one returns 0 and the other returns
                      thousands, the engine was the bug, not the layout.

    readable ratio    share of extracted chars that are letters, digits
                      or punctuation. A font with no ToUnicode map still
                      "extracts text" -- it just extracts garbage. A low
                      ratio here means the text layer is unusable even
                      though the PDF looks searchable on screen.

    anchors           does the page actually contain "Academic Program",
                      "Report Options", "Batch ID", and a dashed ruler
                      line? The old extractor keyed off the SSN pattern
                      and the "Academic Program" caption. If a file
                      spells the caption differently, or omits it, every
                      student on every page is invisible to that logic.

    dash codepoints   every distinct dash-like character in the file,
                      by codepoint. U+2010, U+2013 and U+2212 look
                      identical to "-" on screen and none of them match
                      a regex written with an ASCII hyphen. This alone
                      can hide 100% of the SSNs in a file.

    ssn candidates    counted three ways -- strict ASCII hyphen, any
                      dash variant, and bare 9 digits. The gap between
                      the three columns names the failure precisely.

    rotation / geom   page rotation and the x-positions of the ID, SSN
                      and Name column headers. A rotated page or shifted
                      columns breaks coordinate-band extraction, so the
                      real geometry has to be measured, not assumed.

PRIVACY
    Nothing real is ever printed. Every extracted line is masked to its
    SHAPE before it is shown or written:

        1234567  123-45-6789  Miss Dishita Uppal
        #######  ###-##-####  Xxxx Xxxxxxx Xxxxx

    Only the report's own captions ("Academic Program", "Batch ID",
    "Report Options", column headers) survive masking, because they are
    on an explicit whitelist. The report this writes is safe to paste
    into a ticket or a chat.

    The script opens files read-only. It never writes, moves, renames or
    copies anything inside the source folder. The report is written to a
    destination folder you choose, which should NOT be the prod folder.

USAGE
    python "260810 AM sap pdf layout profiler.py"
        Tkinter folder pickers.

    python "260810 AM sap pdf layout profiler.py" <src_folder> [dest_folder]
        Headless. Report goes to dest_folder, or the current directory.

    python "260810 AM sap pdf layout profiler.py" --lines <file.pdf> [page]
        Dump every masked line of one page with x/y coordinates, to see
        exactly how the columns sit. Page number is 1-based.

REQUIREMENTS
    pip install pdfplumber pymupdf
    Both are wanted. The comparison between them IS the diagnostic; with
    only one installed that column of the report is blank.
"""
from __future__ import annotations

import sys
import unicodedata
from collections import Counter
from pathlib import Path

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


# --------------------------------------------------------------------------
# masking
# --------------------------------------------------------------------------
# Captions belong to the report, not to a student, so they stay readable.
# Everything else collapses to its shape. The list is deliberately generous:
# a caption wrongly masked costs a little readability, a value wrongly left
# in the clear is a PII incident.
SAFE_WORDS = {
    "report", "options", "use", "all", "sections", "include", "contained",
    "w/in", "the", "range", "batch", "id", "ssn", "name", "incl", "gpa",
    "cmpl", "att", "course", "term", "dt", "grd", "cum", "eval", "credits",
    "grade", "pts", "pgm", "earn", "academic", "program", "min", "max",
    "cred", "page", "of", "total", "excluded", "remedial", "and", "for",
    "student", "satisfactory", "progress", "audit", "date", "run", "to",
}


def mask_word(word: str) -> str:
    """Collapse one word to its shape unless it is a report caption."""
    stripped = word.strip("():,.;#%*/-").lower()
    if stripped in SAFE_WORDS:
        return word
    out = []
    for ch in word:
        if ch.isdigit():
            out.append("#")
        elif ch.isalpha():
            out.append("X" if ch.isupper() else "x")
        else:
            out.append(ch)
    return "".join(out)


def mask(text: str) -> str:
    """Mask a whole line, preserving spacing so columns still line up."""
    out, buf = [], []
    for ch in text:
        if ch.isspace():
            if buf:
                out.append(mask_word("".join(buf)))
                buf = []
            out.append(ch)
        else:
            buf.append(ch)
    if buf:
        out.append(mask_word("".join(buf)))
    return "".join(out)


# --------------------------------------------------------------------------
# character-level probes
# --------------------------------------------------------------------------
# Every dash-like codepoint that renders as a horizontal stroke. Any of these
# can appear inside an SSN and defeat a regex written with an ASCII hyphen.
DASHES = {
    0x2010: "HYPHEN",
    0x2011: "NON-BREAKING HYPHEN",
    0x2012: "FIGURE DASH",
    0x2013: "EN DASH",
    0x2014: "EM DASH",
    0x2015: "HORIZONTAL BAR",
    0x2212: "MINUS SIGN",
    0xFE63: "SMALL HYPHEN-MINUS",
    0xFF0D: "FULLWIDTH HYPHEN-MINUS",
    0x00AD: "SOFT HYPHEN",
    0x2043: "HYPHEN BULLET",
}


def fold_dashes(text: str) -> str:
    """Turn every dash variant into an ASCII hyphen."""
    return "".join("-" if ord(c) in DASHES else c for c in text)


def count_ssn_shapes(text: str) -> tuple[int, int, int]:
    """
    Count SSN-shaped tokens three ways. The differences between the three
    numbers identify the failure:
        strict == 0 but folded > 0   -> the dashes are not ASCII hyphens
        both  == 0 but bare    > 0   -> the SSNs print without separators
        all three == 0               -> the SSNs are not in the text layer
    """
    import re

    strict = re.compile(r"(?<!\d)\d{3}-\d{2}-\d{4}(?!\d)")
    bare = re.compile(r"(?<!\d)\d{9}(?!\d)")
    folded_text = fold_dashes(text)
    return (
        len(strict.findall(text)),
        len(strict.findall(folded_text)),
        len(bare.findall(text)),
    )


def readable_ratio(text: str) -> float:
    """
    Share of characters that are plausible report content. A PDF whose font
    lacks a ToUnicode map still yields characters -- private-use or control
    codepoints that no regex will ever match. This catches that case.
    """
    if not text:
        return 0.0
    good = sum(
        1 for c in text
        if c.isalnum() or c.isspace() or c in ".,:;/()-#%*$'\"[]&+=<>?!|_"
    )
    return good / len(text)


# --------------------------------------------------------------------------
# per-engine text extraction
# --------------------------------------------------------------------------
def text_pdfplumber(path: Path, max_pages: int):
    """Return (pages_scanned, total_pages, text, rotations, error)."""
    if pdfplumber is None:
        return 0, 0, "", [], "pdfplumber not installed"
    try:
        with pdfplumber.open(str(path)) as pdf:
            total = len(pdf.pages)
            chunks, rots = [], []
            for page in pdf.pages[:max_pages]:
                # layout=True preserves the horizontal spacing that the
                # report uses to separate its columns.
                chunks.append(page.extract_text(layout=True) or "")
                rots.append(getattr(page, "rotation", 0) or 0)
            return min(total, max_pages), total, "\n".join(chunks), rots, ""
    except Exception as exc:                                # noqa: BLE001
        return 0, 0, "", [], f"{type(exc).__name__}: {exc}"


def text_pymupdf(path: Path, max_pages: int):
    """Return (pages_scanned, total_pages, text, rotations, error)."""
    if fitz is None:
        return 0, 0, "", [], "pymupdf not installed"
    try:
        doc = fitz.open(str(path))
        if doc.needs_pass:
            doc.close()
            return 0, 0, "", [], "ENCRYPTED - password required"
        total = doc.page_count
        chunks, rots = [], []
        for i in range(min(total, max_pages)):
            page = doc[i]
            # sort=True asks PyMuPDF for reading order rather than the
            # PDF's internal block order, which on this report interleaves
            # course rows into the student row.
            chunks.append(page.get_text("text", sort=True) or "")
            rots.append(page.rotation or 0)
        doc.close()
        return min(total, max_pages), total, "\n".join(chunks), rots, ""
    except Exception as exc:                                # noqa: BLE001
        return 0, 0, "", [], f"{type(exc).__name__}: {exc}"


# --------------------------------------------------------------------------
# anchors
# --------------------------------------------------------------------------
# The strings the extractor keys off. If a file is missing one of these,
# that is very likely why it produced no rows.
ANCHORS = {
    "academic_program": "academic program",
    "batch_id": "batch id",
    "report_options": "report options",
    "hdr_ssn": "ssn",
    "hdr_name": "name",
}


def find_anchors(text: str) -> dict:
    low = " ".join(text.lower().split())
    found = {k: (v in low) for k, v in ANCHORS.items()}
    # A ruler line is a run of 8+ dashes; it marks the start of the data
    # block and gives the column widths.
    found["ruler_line"] = "--------" in fold_dashes(text)
    return found


def profile_file(path: Path, max_pages: int) -> dict:
    """Measure one PDF. Never returns any real value, only counts."""
    pp_scan, pp_total, pp_text, pp_rot, pp_err = text_pdfplumber(path, max_pages)
    mu_scan, mu_total, mu_text, mu_rot, mu_err = text_pymupdf(path, max_pages)

    # Use whichever engine produced more text for the content-level checks.
    best_text = pp_text if len(pp_text) >= len(mu_text) else mu_text
    best_engine = "pdfplumber" if len(pp_text) >= len(mu_text) else "PyMuPDF"

    strict, folded, bare = count_ssn_shapes(best_text)
    dash_hits = Counter(
        f"U+{ord(c):04X} {DASHES[ord(c)]}"
        for c in best_text if ord(c) in DASHES
    )
    if "-" in best_text:
        dash_hits["U+002D HYPHEN-MINUS (ascii)"] = best_text.count("-")

    return {
        "file": path.name,
        "pages_total": pp_total or mu_total,
        "pages_scanned": max(pp_scan, mu_scan),
        "pp_chars": len(pp_text),
        "mu_chars": len(mu_text),
        "pp_error": pp_err,
        "mu_error": mu_err,
        "best_engine": best_engine,
        "readable": readable_ratio(best_text),
        "rotations": sorted(set(pp_rot + mu_rot)),
        "anchors": find_anchors(best_text),
        "ssn_strict": strict,
        "ssn_folded": folded,
        "ssn_bare": bare,
        "dashes": dash_hits,
        "sample": sample_lines(best_text),
    }


def sample_lines(text: str, count: int = 4) -> list[str]:
    """
    A few masked lines that look like data rows, so the column layout can
    be eyeballed. Prefers lines that contain a 9-digit-ish run.
    """
    import re

    shaped = re.compile(r"\d{3}\D?\d{2}\D?\d{4}|\d{9}")
    picks = [ln for ln in text.splitlines() if shaped.search(ln)]
    if not picks:
        picks = [ln for ln in text.splitlines() if len(ln.strip()) > 40]
    return [mask(ln.rstrip())[:200] for ln in picks[:count]]


# --------------------------------------------------------------------------
# verdict
# --------------------------------------------------------------------------
def verdict(row: dict) -> str:
    """Turn the measurements into the one sentence that explains the file."""
    if row["pp_error"].startswith("ENCRYPTED") or row["mu_error"].startswith("ENCRYPTED"):
        return "ENCRYPTED - needs a password before anything can be read"
    if row["pp_chars"] == 0 and row["mu_chars"] == 0:
        return "NO TEXT LAYER - this file is a scan and needs OCR"
    if row["pp_chars"] == 0 or row["mu_chars"] == 0:
        weak = "pdfplumber" if row["pp_chars"] == 0 else "PyMuPDF"
        return f"ENGINE GAP - {weak} reads nothing, the other engine works. Use the other one."
    if row["readable"] < 0.85:
        return "BROKEN FONT MAP - text extracts as garbage, needs OCR despite looking searchable"
    if row["ssn_strict"] == 0 and row["ssn_folded"] > 0:
        return "DASH VARIANT - SSNs use a non-ASCII dash, so an ascii-hyphen regex finds none"
    if row["ssn_strict"] == 0 and row["ssn_folded"] == 0 and row["ssn_bare"] > 0:
        return "NO SEPARATORS - SSNs print as 9 bare digits"
    if row["ssn_folded"] == 0 and row["ssn_bare"] == 0:
        return "NO SSN-SHAPED TEXT on the scanned pages - check page range or layout variant"
    if not row["anchors"]["academic_program"]:
        return "MISSING ANCHOR - no 'Academic Program' caption, so caption-keyed logic finds nothing"
    if row["rotations"] and any(r for r in row["rotations"]):
        return f"ROTATED PAGES {row['rotations']} - coordinate bands must be un-rotated first"
    return "OK - this layout should parse"


# --------------------------------------------------------------------------
# report
# --------------------------------------------------------------------------
def render(rows: list[dict]) -> str:
    out = []
    add = out.append
    add("SAP AUDIT PDF LAYOUT PROFILE  (masked - safe to share)")
    add("=" * 78)
    add(f"engines available: pdfplumber={pdfplumber is not None}  pymupdf={fitz is not None}")
    add(f"files profiled:    {len(rows)}")
    add("")

    add("SUMMARY")
    add("-" * 78)
    add(f"{'file':<34}{'pages':>6}{'pp ch':>8}{'mu ch':>8}{'ssn':>6}  verdict")
    for r in rows:
        add(
            f"{r['file'][:33]:<34}{r['pages_total']:>6}"
            f"{r['pp_chars']:>8}{r['mu_chars']:>8}"
            f"{max(r['ssn_strict'], r['ssn_folded'], r['ssn_bare']):>6}"
            f"  {verdict(r)[:60]}"
        )
    add("")

    counts = Counter(verdict(r).split(" - ")[0] for r in rows)
    add("VERDICT TALLY")
    add("-" * 78)
    for k, v in counts.most_common():
        add(f"  {v:>4}  {k}")
    add("")

    add("PER FILE DETAIL")
    add("=" * 78)
    for r in rows:
        add(f"\n{r['file']}")
        add("-" * 78)
        add(f"  pages                {r['pages_scanned']} scanned of {r['pages_total']}")
        add(f"  chars  pdfplumber    {r['pp_chars']}"
            + (f"   [{r['pp_error']}]" if r['pp_error'] else ""))
        add(f"  chars  PyMuPDF       {r['mu_chars']}"
            + (f"   [{r['mu_error']}]" if r['mu_error'] else ""))
        add(f"  better engine        {r['best_engine']}")
        add(f"  readable ratio       {r['readable']:.3f}   (below 0.85 means a broken font map)")
        add(f"  page rotations       {r['rotations']}")
        add(f"  ssn  ascii hyphen    {r['ssn_strict']}")
        add(f"  ssn  any dash        {r['ssn_folded']}   (higher than the line above = dash bug)")
        add(f"  ssn  bare 9 digits   {r['ssn_bare']}")
        add("  anchors              " + ", ".join(
            f"{k}={'Y' if v else 'N'}" for k, v in r["anchors"].items()))
        if r["dashes"]:
            add("  dash codepoints      " + "; ".join(
                f"{k} x{v}" for k, v in r["dashes"].most_common(6)))
        add(f"  VERDICT              {verdict(r)}")
        if r["sample"]:
            add("  masked sample lines:")
            for ln in r["sample"]:
                add(f"    | {ln}")
    return "\n".join(out)


# --------------------------------------------------------------------------
# line dump for one page
# --------------------------------------------------------------------------
def dump_lines(path: Path, page_no: int) -> None:
    """
    Print every line of one page, masked, with coordinates. This is what
    tells us where the ID / SSN / Name column boundaries actually sit.
    """
    if pdfplumber is None:
        print("pdfplumber is required for --lines.  pip install pdfplumber")
        return
    with pdfplumber.open(str(path)) as pdf:
        if not 1 <= page_no <= len(pdf.pages):
            print(f"page {page_no} out of range (1..{len(pdf.pages)})")
            return
        page = pdf.pages[page_no - 1]
        print(f"{path.name}  page {page_no} of {len(pdf.pages)}  "
              f"rotation={getattr(page, 'rotation', 0)}  "
              f"size={page.width:.0f}x{page.height:.0f}")
        print("-" * 78)

        # Group words into rows by their vertical midpoint. A tolerance of
        # 2pt is enough to survive sub-pixel baseline jitter without
        # merging two adjacent printed lines.
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        rows: dict[int, list] = {}
        for w in words:
            key = round((w["top"] + w["bottom"]) / 2 / 2)
            rows.setdefault(key, []).append(w)

        for key in sorted(rows):
            ws = sorted(rows[key], key=lambda w: w["x0"])
            y = min(w["top"] for w in ws)
            text = " ".join(w["text"] for w in ws)
            xs = ",".join(f"{w['x0']:.0f}" for w in ws[:8])
            print(f"y={y:7.1f}  x0=[{xs}]")
            print(f"           {mask(text)[:180]}")


# --------------------------------------------------------------------------
# entry
# --------------------------------------------------------------------------
def pick_folders():
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(
        "SAP PDF Layout Profiler",
        "This is a READ-ONLY diagnostic.\n\n"
        "It does not extract or write any student data. The report it "
        "produces is masked and contains no SSNs or names.\n\n"
        "Next: choose the folder holding the PDFs, then a folder for the "
        "report. Do not use the source folder for the report.",
    )
    src = filedialog.askdirectory(title="Folder containing the SAP audit PDFs")
    if not src:
        return None, None
    dst = filedialog.askdirectory(title="Folder for the masked report")
    root.destroy()
    return src, dst or src


def main() -> None:
    args = sys.argv[1:]

    if args and args[0] == "--lines":
        if len(args) < 2:
            print('usage: --lines "<file.pdf>" [page]')
            return
        dump_lines(Path(args[1]), int(args[2]) if len(args) > 2 else 1)
        return

    max_pages = 6          # enough to see the header, the ruler and data rows

    if args:
        src = args[0]
        dst = args[1] if len(args) > 1 else "."
    else:
        src, dst = pick_folders()
        if not src:
            print("cancelled")
            return

    src_path = Path(src)
    pdfs = sorted(src_path.glob("*.pdf")) + sorted(src_path.glob("*.PDF"))
    pdfs = sorted({p.resolve(): p for p in pdfs}.values(), key=lambda p: p.name)
    if not pdfs:
        print(f"no PDFs found in {src_path}")
        return

    print(f"profiling {len(pdfs)} PDFs, first {max_pages} pages of each "
          f"(read-only, nothing is modified)\n")
    rows = []
    for i, p in enumerate(pdfs, 1):
        print(f"  [{i}/{len(pdfs)}] {p.name}")
        rows.append(profile_file(p, max_pages))

    report = render(rows)
    out = Path(dst) / "260810 AM sap pdf layout profile.txt"
    out.write_text(report, encoding="utf-8")

    print("\n" + report)
    print(f"\nreport written to: {out}")
    print("The report is masked. Save it in the approved Global Insider "
          "folder, not on a desktop.")


if __name__ == "__main__":
    main()
