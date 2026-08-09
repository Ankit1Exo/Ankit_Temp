"""Generate a synthetic (fabricated, non-real) SAP Audit Report test PDF
that reproduces the REAL-WORLD conditions the first test PDF missed.

Every word is placed individually at a jittered y, the way a real text
layer comes out -- the first test PDF put every word of a line at an
identical y, which hid the line-grouping bug entirely.

    page 1 -- normal caption row, but every line carries sub-point y
              jitter deliberately straddling a 3pt bucket edge. This is
              the regression test for the bug where a student line was
              torn into an "ID + SSN" fragment and a separate "name"
              fragment, and the student was silently dropped.
        1. wide run of spaces between the name and "Academic Program:"
        2. masked SSN ("XXX-XX-4569")
        3. very wide gap, no middle initial

    page 2 -- a CONTINUATION page: no "ID SSN Name" caption row and no
              report title, only "Academic Program:" to identify it. The
              earlier gate skipped pages like this outright.
        4. Mrs. prefix + middle initial

    page 3 -- caption row present, but a student with NO SSN printed at
              all, so the row can only be found by column position.
        5. no SSN, detected via the ID and Name column x-positions

Used only to validate "260809 AM sap audit identity extractor.py" -- the
IDs, SSNs and names below are all made up. No real PII.
"""
from pathlib import Path

import pymupdf as fitz

doc = fitz.open()
SIZE = 8
FONT = "cour"
CHAR_W = SIZE * 0.6  # Courier advance width
LEFT_X = 24

# Sub-point vertical jitter, cycled word by word.
JITTER = [0.0, 1.2, 0.4, 1.6, 0.2, 1.4]


def measure_ascent():
    """Gap between the baseline passed to insert_text and the bbox top
    that get_text('words') reports back. Measured rather than hardcoded --
    it depends on the font and size."""
    probe = fitz.open()
    page = probe.new_page(width=200, height=200)
    page.insert_text((10, 100), "X", fontsize=SIZE, fontname=FONT)
    top = page.get_text("words")[0][1]
    probe.close()
    return 100 - top


ASCENT = measure_ascent()


def snap(y):
    """Shift the baseline so the reported word tops land at 3k+1, and
    therefore y0+JITTER spans 3k+1 .. 3k+2.6 -- straddling 3k+1.5.

    This is what makes the test bite. A round(y/3)*3 grouper puts its
    boundaries at y = 3k+1.5, so a line whose words happen to sit wholly
    inside one bucket survives bucketing by luck -- which is why the first
    attempt at this test passed under BOTH groupers and proved nothing.
    Snapping against the measured word top guarantees the words fall on
    both sides of a boundary, which is exactly the case that tore a
    student line in two."""
    return 3 * round((y - ASCENT - 1) / 3) + 1 + ASCENT


def put_line(page, y, text, jitter=True):
    """Place each word at its monospace column, with per-word y jitter, so
    runs of spaces survive as real horizontal gaps."""
    y = snap(y)
    col, i = 0, 0
    for token in text.split(" "):
        if token:
            dy = JITTER[i % len(JITTER)] if jitter else 0.0
            page.insert_text((LEFT_X + col * CHAR_W, y + dy), token,
                             fontsize=SIZE, fontname=FONT)
            i += 1
        col += len(token) + 1


def report_header(page, page_no):
    put_line(page, 24, f"Aug 03 2023     Satisfactory Academic Progress Audit Report      Page {page_no}")
    put_line(page, 36, "10:30AM              Detail of Results by Student by SAP Type")
    put_line(page, 48, "     Report Options:  Use All Sections   Include Sections: Contained W/in the Range")


def caption_row(page, y):
    put_line(page, y, "        ID SSN         Name              Incl Incl  GPA      GPA")
    put_line(page, y + 12, "   Course Name   Term/Dt  Grd Cum  Eval  Credits Grade Pts   %")
    return y + 26


def student(page, y, line, skipped_section=False):
    put_line(page, y, line)
    put_line(page, y + 12,
             "#Excluded Remedial Credits     SAP Type: DHDHS   Doctor of Health Sciences SAP")
    put_line(page, y + 24, "   DHSC-821   21FA2  A  Yes  No     3.00   12.00000     3.00   12.00000")
    if skipped_section:
        put_line(page, y + 36, "   DHSC-831   22SP2       No   Section skipped - No Verified Grade Exists")
        return y + 52
    return y + 40


AP = "Academic Program: DH.DHSC (2021, min pgm cred = 54.00, max pgm cred = 81.00)"

# --------------------------------------------------------------- page 1
page = doc.new_page(width=900, height=520)
report_header(page, 1)
y = caption_row(page, 84)
y = student(page, y, f"  1234567 555-12-4567 Mrs. Jane D. Smith          {AP}")
y = student(page, y, f"  1234569 XXX-XX-4569 Mr. Robert A. Chen      {AP}", skipped_section=True)
y = student(page, y, f"  1234568 555-12-4568 Anna Perez Lopez                        {AP}")

# --------------------------------------------------------------- page 2
# Continuation page: no title block, no caption row.
page = doc.new_page(width=900, height=520)
y = 40
y = student(page, y, f"  1234571 555-12-4571 Mrs. Priya N. Raman        {AP}")

# --------------------------------------------------------------- page 3
# Caption row present; one student has no SSN printed at all.
page = doc.new_page(width=900, height=520)
report_header(page, 3)
y = caption_row(page, 84)
y = student(page, y, f"  1234576             Ms. Grace T. Oyelaran      {AP}")

out_path = str(Path(__file__).parent / "sample_sap_audit_hard_synthetic.pdf")
doc.save(out_path)
doc.close()
print("wrote", out_path)
