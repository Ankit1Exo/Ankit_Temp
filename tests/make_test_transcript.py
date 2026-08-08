"""Generate a synthetic (fabricated, non-real) 2-page transcript test PDF:
page 1 = the dotted-rule "summary" layout (Layout A, full middle name,
SSN present but uncaptioned), page 2 = the "Etran Omed Only" layout
(Layout B, middle-initial-only name, address, and captioned ID Number/
SSN/Birth Date/Birth Name) for the SAME fabricated person, matched by
SSN. Used only to validate "260808 AM transcript identity extractor.py"
-- no real PII."""
from pathlib import Path

import pymupdf as fitz

doc = fitz.open()
size = 10

# ---- Page 1: Layout A (summary) ----
p1 = doc.new_page(width=650, height=500)


def put1(x, y, text, sz=size):
    p1.insert_text((x, y), text, fontsize=sz)


put1(30, 30, "XGRA    mclambm")
put1(30, 55, "Jamie Michael Testerson")  # full middle name -- this is what should win
put1(30, 80, "DOB: 14 Dec 1995    Student ID: 1138266    Print Date: 15 Jul 2022")
put1(30, 105, "TRANSFER CREDITS      Hours Attempted   0.0      Hours Passed   0.0")
put1(30, 125, "555-12-4567")  # SSN, uncaptioned, shape-based match target
put1(30, 150, "20FAD12 2020 Fall D01, DO-2, PA AY: 2020      ATT   CPT   PTS")
put1(40, 170, "MPAP 504  MCO1  CLINICAL MEDICINE I     C   3.0  3.0  6.0")

# ---- Page 2: Layout B (Etran Omed Only) ----
p2 = doc.new_page(width=800, height=500)


def put2(x, y, text, sz=size):
    p2.insert_text((x, y), text, fontsize=sz)


put2(30, 30, "07/15/22")
put2(300, 30, "Etran Omed Only")
put2(600, 30, "Page 1 of 2")

put2(30, 60, "Mr.Jamie M. Testerson")
put2(430, 60, "ID Number: 1138266")

put2(30, 80, "42 Rideout Ct")
put2(430, 80, "SSN: 555-12-4567")

put2(30, 100, "Apt 12-202")
put2(430, 100, "Birth Date: 12/14/95")

put2(30, 120, "Fort Leonard Wood, MO 65473")
put2(430, 120, "Birth Name:")

put2(30, 140, "1138266")

put2(30, 170, "Course        Title                  Grd  Hrs Att  Hrs Cmpt  Hrs Gpa  Grade Points  Course Dates")
put2(30, 190, "MPAP  504  CLINICAL MEDICINE   C   3.00   3.00   3.00   6.00000   07/28/20-12/18/20")

out_path = str(Path(__file__).parent / "sample_transcript_synthetic.pdf")
doc.save(out_path)
doc.close()
print("wrote", out_path)
