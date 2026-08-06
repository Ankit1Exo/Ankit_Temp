"""Generate a synthetic (fabricated, non-real) single-employee "Copy D" W-2
test PDF that mirrors the box layout the user shared a screenshot of: box e
(Employee's name/address) on the left, boxes 9-14 / 12a-12d on the right, at
matching row heights, with box 15-20 (state) spanning the full width at the
bottom. Used only to validate w2_ssn_name_address_csv.py extraction logic
against this specific layout -- no real PII."""
from pathlib import Path

import pymupdf as fitz

doc = fitz.open()
page = doc.new_page(width=1300, height=790)
font = "helv"
size = 11

def put(x, y, text, sz=size):
    page.insert_text((x, y), text, fontsize=sz, fontname=font)

# Box a / OMB row
put(330, 40, "a Employee's social security number")
put(650, 40, "OMB No. 1545-0008  PZB   330000   001251")
put(330, 60, "416-17-2201")

# Box b
put(80, 80, "b Employer identification number (EIN)")
put(80, 100, "98-7654321")
put(730, 80, "1 Wages, tips, other compensation")
put(1010, 80, "2 Federal income tax withheld")
put(730, 100, "41047.55")
put(1010, 100, "5652.74")

# Box c
put(80, 130, "c Employer's name, address, and ZIP code")
put(80, 150, "SAMPLE TEST EMPLOYER LLC")
put(80, 175, "5904 SAMPLE OAKS PKY #D")
put(80, 195, "TAMPA FL 33610")
put(730, 130, "3 Social security wages")
put(1010, 130, "4 Social security tax withheld")
put(730, 150, "42726.44")
put(1010, 150, "2649.04")
put(730, 175, "5 Medicare wages and tips")
put(1010, 175, "6 Medicare tax withheld")
put(730, 195, "42726.44")
put(1010, 195, "619.53")
put(730, 220, "7 Social security tips")
put(1010, 220, "8 Allocated tips")

# Box d
put(80, 280, "d Control number")
put(80, 300, "001251 ATLA/PZB")

# Row: box e caption | box 9 | box 10 -- SAME y as the name/address block starts
ey = 330
put(80, ey, "e Employee's first name and initial")
put(330, ey, "Last name")
put(450, ey, "Suff.")
put(730, ey, "9")
put(1010, ey, "10 Dependent care benefits")

# Row: employee name value | box 11 | box 12a
vy = 355
put(80, vy, "NADINE")
put(330, vy, "ABERNETHY")
put(730, vy, "11 Nonqualified plans")
put(1010, vy, "12a See instructions for box 12")
put(1010, vy + 22, "D          1709.09")

# Row: street line 1 | box 13 checkboxes | box 12b
sy1 = 385
put(80, sy1, "875 117TH TERR. N.")
put(730, sy1, "13 Statutory   Retirement   Third-party")
put(730, sy1 + 15, "   employee    plan  X     sick pay")
put(1010, sy1 + 15, "12b")
put(1050, sy1 + 15, "W          211.38")

# Row: street line 2 (apt) | box 14
sy2 = 415
put(80, sy2, "APT# 7")
put(1010, sy2 - 15, "14 Other")
put(1010, sy2 + 10, "12c")

# Row: city/state/zip | box 12d
sy3 = 445
put(80, sy3, "ST. PETERSBURG,FL 33716")
put(1010, sy3, "12d")

put(80, 480, "f Employee's address and ZIP code")

# Box 15-20, full width, bottom
by = 520
put(80, by, "15 State   Employer's state ID number")
put(430, by, "16 State wages, tips, etc.")
put(650, by, "17 State income tax")
put(870, by, "18 Local wages, tips, etc.")
put(1080, by, "19 Local income tax")
put(1220, by, "20 Locality name")
put(80, by + 20, "FL")

out_path = str(Path(__file__).parent / "sample_w2_copyd_synthetic.pdf")
doc.save(out_path)
doc.close()
print("wrote", out_path)
