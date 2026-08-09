"""Generate a synthetic (fabricated, non-real) SAP Audit Report test PDF
that reproduces the "same SSN appears twice -- once with a name, once
blank" symptom.

The cause is a text-layer BLOCK SPLIT. The report leaves a wide gap
between the SSN column and the Name column, and PDF generators routinely
emit the two sides of such a gap as separate text objects. PyMuPDF's raw
text layer then reports them as two separate lines:

    "1234567 555-12-4567"                         <- SSN, no name  -> blank name row
    "Mrs. Jane D. Smith  Academic Program: ..."   <- name, no SSN  -> ignored

while the word-position path, which groups by geometry, still sees the
whole line. So the same student is read two different ways, and one of
the readings has no name on it.

This file also reproduces the line WRAP visible in the real report, where
a long student line overflows and its tail ("00)") starts at column 0 of
the next line, and the "Section skipped - No Verified Grade Exists" row.

Used only to validate "260809 AM sap audit identity extractor.py" -- the
IDs, SSNs and names below are all made up. No real PII.
"""
from pathlib import Path

import pymupdf as fitz

doc = fitz.open()
SIZE = 8
FONT = "cour"
CHAR_W = SIZE * 0.6
LEFT_X = 24


def at(col):
    return LEFT_X + col * CHAR_W


def put(page, col, y, text):
    page.insert_text((at(col), y), text, fontsize=SIZE, fontname=FONT)


AP = "Academic Program: DH.DHSC (2021, min pgm cred = 54.00, max pgm cred = 81"

page = doc.new_page(width=960, height=420)

put(page, 0, 24, "Aug 03 2023            Satisfactory Academic Progress Audit Report")
put(page, 118, 24, "Page")
put(page, 0, 33, "1")
put(page, 0, 42, "10:30AM              Detail of Results by Student by SAP Type")
put(page, 0, 54, "     Report Options:  Use All Sections   Include Sections: Contained W/in the Range")
put(page, 0, 66, "                     Batch ID: SAPC_BMSEVERIN0325_37695_20304")
put(page, 0, 78, "Att")
put(page, 8, 86, "ID SSN          Name      Incl Incl GPA     GPA")
put(page, 0, 98, "Pgm")
put(page, 3, 106, "Course Name   Term/Dt  Grd Cum  Eval  Credits Grade Pts   %")
put(page, 0, 118, "-" * 110)

# --- the student line, deliberately emitted as TWO separate text objects
#     with a wide gap between them: ID+SSN on the left, name+program on
#     the right. This is what splits the text layer.
put(page, 2, 134, "1234567 555-12-4567")
put(page, 30, 134, f"Mrs. Jane D. Smith    {AP}")
put(page, 0, 146, "00)")  # the wrapped tail, at column 0 of the next line

put(page, 0, 158, "#Excluded Remedial Credits     SAP Type: DHDHS   Doctor of Health Sciences SAP")
put(page, 4, 170, "DHSC-821   21FA2  A  Yes  No     3.00   12.00000     3.00   12.00000")
put(page, 4, 182, "DHSC-831   22SP2       No    Section skipped - No Verified Grade Exists")
put(page, 4, 194, "DHSC-833   21FA2  A  Yes  No     3.00   12.00000     3.00   12.00000")

# --- a second student, same split, to confirm the fix is not a one-off
put(page, 2, 214, "1234568 555-12-4568")
put(page, 30, 214, f"Anna Perez Lopez      {AP}")
put(page, 0, 226, "00)")
put(page, 0, 238, "#Excluded Remedial Credits     SAP Type: DHDHS   Doctor of Health Sciences SAP")
put(page, 4, 250, "DHSC-827   22SU2  A  Yes  No     3.00   12.00000     3.00   12.00000")

out_path = str(Path(__file__).parent / "sample_sap_audit_split_synthetic.pdf")
doc.save(out_path)
doc.close()
print("wrote", out_path)
