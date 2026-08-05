"""Generate a synthetic (fabricated, non-real) W-2 test PDF with two pages:
page 1 = 2-line address (no apt/suite), page 2 = 3-line address (with apt/suite).
Used only to validate w2_ssn_name_address_csv.py extraction logic -- no real PII."""
from pathlib import Path

import pymupdf as fitz

doc = fitz.open()

def add_page(street_lines, ssn, first, mi, last, suffix=""):
    page = doc.new_page(width=612, height=792)
    font = "helv"
    size = 10

    def put(x, y, text, sz=size):
        page.insert_text((x, y), text, fontsize=sz, fontname=font)

    put(50, 50, "a Employee's social security number")
    put(300, 50, "OMB No. 1545-0008")
    put(50, 70, ssn)

    put(50, 100, "b Employer identification number (EIN)")
    put(50, 120, "98-7654321")

    put(50, 150, "c Employer's name, address, and ZIP code")
    put(50, 170, "SAMPLE TEST EMPLOYER LLC")
    put(50, 190, "100 MAIN ST")
    put(50, 210, "ANYTOWN FL 00000")

    put(50, 240, "d Control number")
    put(50, 260, "000123 TEST")

    # Box e caption with Last name / Suff sub-labels -- placed per-word so the
    # x-positions anchor a real column split.
    ey = 290
    put(50, ey, "e Employee's first name and initial")
    put(300, ey, "Last name")
    put(420, ey, "Suff.")

    # Value line below box e caption
    vy = 310
    put(50, vy, first)
    put(50 + 7 * (len(first) + 1), vy, mi)
    put(300, vy, last)
    if suffix:
        put(420, vy, suffix)

    put(50, 340, "f Employee's address and zip code")
    y = 360
    for line in street_lines:
        put(50, y, line)
        y += 20

    doc_row_y = y + 10
    put(50, doc_row_y, "15 State  Employer's state ID number")
    put(50, doc_row_y + 20, "FL")

    return page

# Page 1: 2-line address, no apt/suite
add_page(
    street_lines=["456 OAK AVENUE", "SPRINGFIELD, IL 62704"],
    ssn="123-45-6789",
    first="JOHN", mi="A", last="SAMPLE",
)

# Page 2: 3-line address, with apt/suite
add_page(
    street_lines=["789 ELM STREET", "APT 2B", "RIVERSIDE, CA 92501"],
    ssn="321-54-9876",
    first="JANE", mi="B", last="TESTCASE",
)

out_path = str(Path(__file__).parent / "sample_w2_synthetic.pdf")
doc.save(out_path)
doc.close()
print("wrote", out_path)
