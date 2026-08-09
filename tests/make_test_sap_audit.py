"""Generate a synthetic (fabricated, non-real) 2-page SAP Audit Report test
PDF covering the awkward student-line cases:

    page 1 -- dashed SSNs, and the "Academic Program:" caption present
        1. Mrs. prefix + middle initial      -> Jane / D. / Smith
        2. no prefix, no initial             -> Anna / (blank) / Perez Lopez
        3. two middle initials               -> Robert / A. B. / Chen
        4. surname starting "Mr" (Mroz) with a trailing lone letter, which
           must NOT be stripped as a prefix nor read as a middle initial
        5. Dr. prefix fused to the name ("Dr.Priya")

    page 2 -- no "Academic Program:" caption anywhere, so the name boundary
        has to fall back
        6. running (undashed) 9-digit SSN, name is last thing on the line
        7. dashed SSN followed by a different caption ("SAP Type:")
        8. a single-word name
        9. a name with a "Jr." suffix

Used only to validate "260809 AM sap audit identity extractor.py" -- the
IDs, SSNs and names below are all made up. No real PII.
"""
from pathlib import Path

import pymupdf as fitz

doc = fitz.open()
SIZE = 8
LEFT_X = 24
FONT = "cour"  # the real report is monospaced


def new_page(page_no):
    page = doc.new_page(width=820, height=460)
    put(page, LEFT_X, 24, "Aug 03 2023                Satisfactory Academic Progress Audit Report")
    put(page, 700, 24, f"Page {page_no}")
    put(page, LEFT_X, 36, "10:30AM              Detail of Results by Student by SAP Type")
    put(page, LEFT_X, 48, "     Report Options:  Use All Sections   Include Sections: Contained W/in the Range")
    put(page, LEFT_X, 60, "                     Batch ID: SAPC_BMSEVERIN0325_37695_20304")
    return page


def put(page, x, y, text):
    page.insert_text((x, y), text, fontsize=SIZE, fontname=FONT)


def header_row(page, y):
    """The caption row the extractor reads column positions from."""
    put(page, LEFT_X, y, "        ID SSN         Name              Incl Incl  GPA      GPA")
    put(page, LEFT_X, y + 12, "   Course Name   Term/Dt  Grd Cum  Eval  Credits Grade Pts   %")
    put(page, LEFT_X, y + 24, "-" * 100)
    return y + 36


def student(page, y, id_ssn_name, trailing):
    put(page, LEFT_X, y, f"  {id_ssn_name}  {trailing}".rstrip())
    put(page, LEFT_X, y + 12,
        "#Excluded Remedial Credits     SAP Type: DHDHS   Doctor of Health Sciences SAP")
    # A couple of course rows, so the page isn't unrealistically bare and we
    # can confirm they are NOT mistaken for student lines.
    put(page, LEFT_X + 24, y + 24, "DHSC-821   21FA2  A  Yes  No     3.00   12.00000     3.00   12.00000")
    put(page, LEFT_X + 24, y + 36, "DHSC-833   21FA2  A  Yes  No     3.00   12.00000     3.00   12.00000")
    return y + 54


ACADEMIC = "Academic Program: DH.DHSC (2021, min pgm cred = 54.00, max pgm cred = 81.00)"

# ---------------------------------------------------------------- page 1
page = new_page(1)
y = header_row(page, 84)
y = student(page, y, "1234567 555-12-4567 Mrs. Jane D. Smith", ACADEMIC)
y = student(page, y, "1234568 555-12-4568 Anna Perez Lopez", ACADEMIC)
y = student(page, y, "1234569 555-12-4569 Mr. Robert A. B. Chen", ACADEMIC)
y = student(page, y, "1234570 555-12-4570 Mroz Dana K", ACADEMIC)
y = student(page, y, "1234571 555-12-4571 Dr.Priya N. Raman", ACADEMIC)

# ---------------------------------------------------------------- page 2
page = new_page(2)
y = header_row(page, 84)
# running 9-digit SSN, nothing after the name at all
y = student(page, y, "1234572 555124572 Ms. Grace T. Oyelaran", "")
# a different caption closes the name
y = student(page, y, "1234573 555-12-4573 Liam O'Brien", "SAP Type: DHDHS")
# single-word name
y = student(page, y, "1234574 555-12-4574 Cher", "")
# suffix
y = student(page, y, "1234575 555-12-4575 Mr. Samuel P. Okonkwo Jr.", "")

out_path = str(Path(__file__).parent / "sample_sap_audit_synthetic.pdf")
doc.save(out_path)
doc.close()
print("wrote", out_path)
