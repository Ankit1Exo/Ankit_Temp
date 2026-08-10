"""Generate a synthetic SAP Audit Report test PDF for the ID / SSN / Name /
Address extractor.

Every ID, SSN, name and address below is FABRICATED. No real PII.

Four pages, one per layout the extractor has to survive:

    page 1  address on the line UNDER the student line, with course rows
            underneath it that must not be swept into the address
    page 2  address on the SAME line, after the name -- including "PO Box",
            which carries no digit at its start and so has to stop the name
            by keyword rather than by digit
    page 3  an explicit Address column in the heading row
    page 4  NO heading row at all, so the column half cannot run and the
            plain-text fallback has to produce the students on its own

Also covered, because each has caused a wrong value before:
    - a surname beginning "Mr" (Mroz), which must not lose letters to the
      honorific test
    - a masked SSN (XXX-XX-nnnn) and one printed with no separators
    - "Academic Program:" trailing the name, which must not reach the name
    - a course grade letter ("A") that must not become a middle initial
    - per-word y jitter, as a real text layer carries
"""
from pathlib import Path

import pymupdf as fitz

SIZE = 8
FONT = "cour"
CHAR_W = SIZE * 0.6
LEFT_X = 24
JITTER = [0.0, 0.9, 0.3, 1.1, 0.2, 0.7]

HEAD = "        ID SSN          Name              Incl Incl  GPA      GPA   Cmpl"
HEAD_ADDR = ("        ID SSN          Name              Address"
             "                        Incl  GPA   Cmpl")
COURSE_HEAD = ("   Course Name   Term/Dt  Grd Cum  Eval  Credits Grade Pts"
               "   % Pgm % Earn")
AP = "Academic Program: MS.CLNR (2025, min pgm cred = 38.00, max pgm cred = 60"
EXCLUDED = ("#Excluded Remedial Credits     SAP Type: DHDHS   Doctor of Health"
            " Sciences SAP")

doc = fitz.open()


def put(page, y, text, jitter=True):
    """Place each word at its monospace column, with per-word y jitter, so
    runs of spaces survive as real horizontal gaps in the text layer."""
    col, n = 0, 0
    for token in text.split(" "):
        if token:
            dy = JITTER[n % len(JITTER)] if jitter else 0.0
            page.insert_text((LEFT_X + col * CHAR_W, y + dy), token,
                             fontsize=SIZE, fontname=FONT)
            n += 1
        col += len(token) + 1
    return y


def page_frame(page, page_no, heading=HEAD):
    put(page, 24, "Aug 10 2026        Satisfactory Academic Progress Audit Report")
    put(page, 33, f"Page {page_no}")
    put(page, 42, "10:30AM              Detail of Results by Student by SAP Type")
    put(page, 54, "     Report Options:  Use All Sections   Include Sections:"
                  " Contained W/in the Range")
    put(page, 66, "                     Batch ID: SAPC_BMSEVERIN0325_51962_21031")
    put(page, 78, "Att")                    # wrapped caption
    if heading:
        put(page, 86, heading)              # the row the columns come from
    put(page, 98, "Pgm")                    # wrapped caption
    put(page, 106, COURSE_HEAD)
    put(page, 118, "-" * 118)
    put(page, 127, "---")


def courses(page, y):
    put(page, y, "   CLNR-821   26FA2  A  Yes  No     3.00   12.00000"
                 "     3.00   12.00000")
    put(page, y + 10, "   CLNR-831   26SP2       No    Section skipped"
                      " - No Verified Grade Exists")


# --------------------------------------------------------------------------
# page 1  --  address on the line below the student line
# --------------------------------------------------------------------------
students_below = [
    ("1237906 555-11-1427 Mrs. Birtukhan D. Abate      " + AP,
     "        41 Maple Street Apt 3B",
     "        Springfield, IL 62704"),
    ("1235956 622-33-4573 Markeece Alexander           " + AP,
     "        1180 West Oak Ridge Road",
     "        Dover, DE 19901-4412"),
    ("1234570 555-12-4570 Mroz Dana K                   " + AP,
     "        7 Hillcrest Court",
     "        Tulsa, OK 74104"),
]

page = doc.new_page(width=1000, height=150 + 74 * len(students_below))
page_frame(page, 1)
y = 140
for line, addr1, addr2 in students_below:
    put(page, y, "  " + line)
    put(page, y + 10, "00)")                # the student line's wrapped tail
    put(page, y + 20, addr1)
    put(page, y + 30, addr2)
    put(page, y + 42, EXCLUDED)
    courses(page, y + 52)
    y += 74

# --------------------------------------------------------------------------
# page 2  --  address on the same line, after the name
# --------------------------------------------------------------------------
students_same = [
    "1234569 XXX-XX-4569 Mr. Robert A. B. Chen  88 Sycamore Ave, Reno, NV 89501",
    "1234572 555124572 Ms. Grace T. Oyelaran  PO Box 417, Ithaca, NY 14850",
    "1234575 555-12-4575 Dr.Samuel P. Okonkwo Jr.  2200 Bellevue Pkwy Suite 14,"
    " Wilmington, DE 19809",
]

page = doc.new_page(width=1100, height=150 + 62 * len(students_same))
page_frame(page, 2)
y = 140
for line in students_same:
    put(page, y, "  " + line)
    put(page, y + 10, EXCLUDED)
    courses(page, y + 20)
    y += 62

# --------------------------------------------------------------------------
# page 3  --  an explicit Address column in the heading row
# --------------------------------------------------------------------------
students_column = [
    "1234580 555-12-4580 Liam O'Brien          19 Kingsway Terrace, Boise, ID 83702",
    "1234581 555-12-4581 Aiko Tanaka-Reyes      506 North 4th St, Fargo, ND 58102",
]

page = doc.new_page(width=1100, height=150 + 62 * len(students_column))
page_frame(page, 3, heading=HEAD_ADDR)
y = 140
for line in students_column:
    put(page, y, "  " + line)
    put(page, y + 10, EXCLUDED)
    courses(page, y + 20)
    y += 62

# --------------------------------------------------------------------------
# page 4  --  no heading row, so the text fallback has to carry the page
# --------------------------------------------------------------------------
students_no_head = [
    ("1234590 555-12-4590 Priya Ramanathan             " + AP,
     "        903 Cedar Hollow Lane",
     "        Cary, NC 27513"),
    ("1234591 555-12-4591 Kwame Osei-Bonsu              " + AP,
     "        62 Harbour View Road",
     "        Portland, ME 04101"),
]

page = doc.new_page(width=1000, height=150 + 74 * len(students_no_head))
page_frame(page, 4, heading=None)
y = 140
for line, addr1, addr2 in students_no_head:
    put(page, y, "  " + line)
    put(page, y + 20, addr1)
    put(page, y + 30, addr2)
    put(page, y + 42, EXCLUDED)
    courses(page, y + 52)
    y += 74

# --------------------------------------------------------------------------
# page 5  --  an SSN printed as separate groups, which no pattern matches.
#             The student must still be recovered, from the "Academic Program"
#             caption, rather than silently vanishing.
# --------------------------------------------------------------------------
students_odd_ssn = [
    ("1234599 555 12 4599 Nadia Haddad-Ferreira      " + AP,
     "        14 Riverbend Drive",
     "        Athens, GA 30601"),
]

page = doc.new_page(width=1000, height=150 + 74 * len(students_odd_ssn))
page_frame(page, 5)
y = 140
for line, addr1, addr2 in students_odd_ssn:
    put(page, y, "  " + line)
    put(page, y + 10, "00)")
    put(page, y + 20, addr1)
    put(page, y + 30, addr2)
    put(page, y + 42, EXCLUDED)
    courses(page, y + 52)
    y += 74

out_path = str(Path(__file__).parent / "sample_sap_address_synthetic.pdf")
doc.save(out_path)
doc.close()
print("wrote", out_path)
