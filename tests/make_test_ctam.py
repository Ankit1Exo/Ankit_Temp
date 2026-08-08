"""Generate a synthetic (fabricated, non-real) CTAM registration-form test
PDF with two course rows on the page, so the field-anchored/column-bucketed
extraction logic and the multi-course semicolon-combining can be validated.
Used only to validate "260808 AM ctam registration form extractor.py" --
no real PII."""
from pathlib import Path

import pymupdf as fitz

doc = fitz.open()
page = doc.new_page(width=1050, height=500)
size = 10


def put(x, y, text):
    page.insert_text((x, y), text, fontsize=size)


put(50, 40, "UNDERGRADUATE REGISTRATION FORM -CTAM ONLY")
put(50, 60, "Campbell University - Ft. Bragg and Pope AFB Campuses")

put(50, 100, "Social Security Number")
put(230, 100, "123456789")

put(50, 125, "Last Nam")
put(115, 125, "Testerson")
put(300, 125, "First Name")
put(380, 125, "Jamie")
put(460, 125, "MI")
put(490, 125, "Q")

put(50, 150, "Other Names")

put(50, 175, "Address")
put(110, 175, "42 Rideout Ct")
put(260, 175, "Fort Leonard Wood")
put(430, 175, "MO")
put(470, 175, "65473")

put(50, 200, "Current Telephone")
put(180, 200, "(347)768-7108")
put(350, 200, "E-mail Address")
put(450, 200, "jamie.q.testerson.mil@mail.mil")

headers = [
    ("Term Code", 50), ("Action Requested", 140), ("Section Number", 280),
    ("Subject Code", 380), ("Catalog Number", 470), ("Course Title", 570),
    ("Course Credits", 680), ("Student Class Cost", 790), ("Total Cost of Class", 910),
]
HEADER_Y = 240
for text, x in headers:
    put(x, HEADER_Y, text)

put(40, 260, "[]")  # row-select checkbox glyph -- must be ignored, not a data column

row1 = [
    (55, "680"), (145, "E"), (285, "OL70"), (385, "CRIM"), (475, "232"),
    (575, "Intro to"), (685, "3"), (795, "75"), (915, "$825.00"),
]
for x, text in row1:
    put(x, 260, text)

row2 = [
    (55, "681"), (145, "E"), (285, "OL71"), (385, "MATH"), (475, "101"),
    (575, "Calculus I"), (685, "4"), (795, "100"), (915, "$900.00"),
]
for x, text in row2:
    put(x, 280, text)

put(50, 310, "Total Cost of Courses")
put(230, 310, "$1,725.00")
put(50, 330, "Total Cost to Student")
put(230, 330, "$175.00")
put(50, 350, "Total TA")
put(230, 350, "$1,550.00")

out_path = str(Path(__file__).parent / "sample_ctam_synthetic.pdf")
doc.save(out_path)
doc.close()
print("wrote", out_path)
