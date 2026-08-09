"""Generate a synthetic (fabricated, non-real) 4-page transcript test PDF
covering both page formats and the awkward cases:

    page 1 -- Format B, with an Apt line and a spaced "Mr. " honorific
    page 2 -- Format A, name + Student ID on the DOB caption line
    page 3 -- Format B, NO apt line, and a surname beginning "Mr" ("Mroz")
              which must NOT have its first two letters stripped as an
              honorific
    page 4 -- Format B, honorific fused to the name ("Mrs.Dana"), a
              "Suite" line rather than "Apt", and a filled Birth Name

Used only to validate "260809 AM transcript extractor.py" -- no real PII.
"""
from pathlib import Path

import pymupdf as fitz

doc = fitz.open()
SIZE = 10
LEFT_X = 30
LABEL_X = 440


def new_page():
    return doc.new_page(width=800, height=400)


def put(page, x, y, text):
    page.insert_text((x, y), text, fontsize=SIZE)


def format_b_page(name, street, middle_line, city_line, id_number, ssn, birth_date,
                  birth_name="", trailing_id=""):
    page = new_page()
    put(page, LEFT_X, 30, "07/15/22")
    put(page, 300, 30, "Etran Omed Only")
    put(page, 620, 30, "Page 1 of 2")

    rows = [(name, f"ID Number: {id_number}"), (street, f"SSN: {ssn}")]
    if middle_line:
        rows.append((middle_line, f"Birth Date: {birth_date}"))
        rows.append((city_line, f"Birth Name: {birth_name}".rstrip()))
    else:
        rows.append((city_line, f"Birth Date: {birth_date}"))
        rows.append(("", f"Birth Name: {birth_name}".rstrip()))

    y = 60
    for left, right in rows:
        if left:
            put(page, LEFT_X, y, left)
        put(page, LABEL_X, y, right)
        y += 18

    if trailing_id:
        put(page, LEFT_X + 20, y, trailing_id)

    # A little of the course table, so the page isn't unrealistically bare.
    put(page, LEFT_X, y + 30, "Course    Title                Grd R   Hrs Att  Hrs Cmpt")
    put(page, LEFT_X, y + 48, "MPAP  504  CLINICAL MEDICINE   C       3.00     3.00")
    return page


def format_a_page(name, student_id):
    page = new_page()
    put(page, LEFT_X, 30, "XGRA        mclambm")
    put(page, LEFT_X, 50,
        f"{name}    DOB: 14 Dec XXXX    Student ID: {student_id}    Print Date: 15 Jul 2022")
    put(page, LEFT_X, 68, "TRANSFER CREDITS      Hours Attempted    0.0     Hours Passed    0.0")
    put(page, LEFT_X + 20, 96, "20FAD12 2020 Fall D01, DO-2, PA AY: 2020      ATT   CPT   PTS")
    put(page, LEFT_X + 40, 114, "MPAP 504  MCO1  CLINICAL MEDICINE I    C   3.0  3.0  6.0")
    return page


# page 1 -- spaced honorific, Apt line present, stray ID printed under the address
format_b_page(
    name="Mr. Jamie M. Tester",
    street="1221 Example Street",
    middle_line="Apt 12-202",
    city_line="Raleigh, NC  27606",
    id_number="1234567", ssn="555-12-4567", birth_date="12/14/95",
    trailing_id="1234567",
)

# page 2 -- Format A
format_a_page(name="Tester, Jamie Michael Lee", student_id="1234567")

# page 3 -- no honorific at all, surname starts with "Mr", no Apt line
format_b_page(
    name="Mroz Dana K",
    street="9 Example Road",
    middle_line="",
    city_line="Durham, NC  27701",
    id_number="7654321", ssn="555-99-1234", birth_date="01/02/90",
)

# page 4 -- fused honorific, "Suite" instead of "Apt", Birth Name filled in
format_b_page(
    name="Mrs.Dana K. Example",
    street="5 Test Boulevard",
    middle_line="Suite 400",
    city_line="Cary, NC  27511",
    id_number="1112223", ssn="555-33-2222", birth_date="03/04/88",
    birth_name="Sample, Dana",
)

out_path = str(Path(__file__).parent / "sample_transcript_v2_synthetic.pdf")
doc.save(out_path)
doc.close()
print("wrote", out_path)
