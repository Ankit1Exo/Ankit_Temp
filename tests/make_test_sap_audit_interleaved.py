"""Generate a synthetic (fabricated, non-real) SAP Audit Report test PDF
that reproduces the INTERLEAVED-COLUMN garbage seen in production:

    ID column        SSN            Full Name
    Birtukhan        xxx-xx-1427    Credits Remedial      <- garbage
    1237906          xxx-xx-1427    Birtukhan Abate       <- correct

Both rows carry the SAME SSN, so the workbook showed each student twice --
once correctly and once with the student's first name sitting in the ID
column and the report's own words ("Credits Remedial") as the name.

The cause is the raw text-layer extraction path. On a wide columnar report
PyMuPDF's line detection can pull text from the "#Excluded Remedial
Credits" row in beside the student line, producing a line that has an
SSN in it and report vocabulary either side. The word-position path,
which groups strictly by geometry, reads the same page correctly.

This file lays the text out so the two paths disagree that way: the
student line is emitted as separate left/right text objects, and the
"#Excluded Remedial Credits" row is placed tightly under it.

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


def put(page, col, y, text):
    page.insert_text((LEFT_X + col * CHAR_W, y), text, fontsize=SIZE, fontname=FONT)


AP = "Academic Program: DH.DHSC (2021, min pgm cred = 54.00, max pgm cred = 81"

page = doc.new_page(width=980, height=340)

put(page, 0, 24, "Aug 03 2023            Satisfactory Academic Progress Audit Report")
put(page, 0, 36, "10:30AM              Detail of Results by Student by SAP Type")
put(page, 0, 48, "                     Batch ID: SAPC_BMSEVERIN0325_37695_20304")
put(page, 0, 60, "Att")
put(page, 8, 68, "ID SSN          Name      Incl Incl GPA     GPA")
put(page, 0, 80, "Pgm")
put(page, 3, 88, "Course Name   Term/Dt  Grd Cum  Eval  Credits Grade Pts   %")

# --- student 1: left/right split, with the #Excluded row tight underneath
put(page, 2, 108, "1237906 555-11-1427")
put(page, 30, 108, f"Mrs. Birtukhan Abate   {AP}")
put(page, 0, 117, "00)")
put(page, 0, 126, "#Excluded Remedial Credits     SAP Type: DHDHS   Doctor of Health Sciences SAP")
put(page, 4, 138, "DHSC-821   21FA2  A  Yes  No     3.00   12.00000     3.00   12.00000")
put(page, 4, 150, "DHSC-831   22SP2       No    Section skipped - No Verified Grade Exists")

# --- student 2
put(page, 2, 175, "1235956 622-33-4573")
put(page, 30, 175, f"Markeece Alexander     {AP}")
put(page, 0, 184, "00)")
put(page, 0, 193, "#Excluded Remedial Credits     SAP Type: DHDHS   Doctor of Health Sciences SAP")
put(page, 4, 205, "DHSC-827   22SU2  A  Yes  No     3.00   12.00000     3.00   12.00000")

out_path = str(Path(__file__).parent / "sample_sap_audit_interleaved_synthetic.pdf")
doc.save(out_path)
doc.close()
print("wrote", out_path)
