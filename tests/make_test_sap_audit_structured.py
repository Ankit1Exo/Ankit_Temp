"""Generate a synthetic (fabricated, non-real) SAP Audit Report test PDF
laid out exactly as the real report is:

    line 1    ID SSN     Name    Incl Incl  GPA  GPA ...      <- captions
    line 2    Att                                             <- wrapped caption
    line 3    Course Name  Term/Dt  Grd Cum  Eval  Credits    <- captions
    line 4    ------------------------------------------      <- separator
    then      the student line: ID, SSN and Name in sequence
              its wrapped tail ("00)")
              #Excluded Remedial Credits    SAP Type: DHDHS ...
              the course rows, including a "Section skipped" row

Covers the cases that have caused trouble:
    - the student line runs past the Name column into "Academic Program:",
      which must NOT reach the name
    - "#Excluded Remedial Credits" and the course headings sit close by and
      must never contribute words to a name ("Credits Remedial")
    - a course row's grade letter ("A") must not be read as a middle initial
    - names with and without a middle initial, with and without a prefix
    - a surname starting "Mr" (Mroz), which must not be stripped as an
      honorific
    - a masked SSN, and a running (undashed) SSN
    - per-word y jitter, as real text layers carry

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
JITTER = [0.0, 0.9, 0.3, 1.1, 0.2, 0.7]

HEAD = "        ID SSN          Name              Incl Incl  GPA      GPA       GPA  GPA Cmpl"
COURSE_HEAD = "   Course Name   Term/Dt  Grd Cum  Eval  Credits Grade Pts   % Pgm % Earn Credits"
AP = "Academic Program: DH.DHSC (2021, min pgm cred = 54.00, max pgm cred = 81"
EXCLUDED = "#Excluded Remedial Credits     SAP Type: DHDHS   Doctor of Health Sciences SAP"


def put(page, y, text, jitter=True):
    """Place each word at its monospace column, with per-word y jitter, so
    runs of spaces survive as real horizontal gaps."""
    col, n = 0, 0
    for token in text.split(" "):
        if token:
            dy = JITTER[n % len(JITTER)] if jitter else 0.0
            page.insert_text((LEFT_X + col * CHAR_W, y + dy), token,
                             fontsize=SIZE, fontname=FONT)
            n += 1
        col += len(token) + 1


def report_page(page_no, students):
    page = doc.new_page(width=1000, height=140 + 78 * len(students))
    put(page, 24, f"Aug 03 2023        Satisfactory Academic Progress Audit Report        Page")
    put(page, 33, str(page_no))
    put(page, 42, "10:30AM              Detail of Results by Student by SAP Type")
    put(page, 54, "     Report Options:  Use All Sections   Include Sections: Contained W/in the Range")
    put(page, 66, "                     Batch ID: SAPC_BMSEVERIN0325_37695_20304")

    put(page, 78, "Att")                 # line 2 of the heading block (wrapped)
    put(page, 86, HEAD)                  # line 1 -- the captions we key on
    put(page, 98, "Pgm")                 # wrapped caption
    put(page, 106, COURSE_HEAD)          # line 3
    put(page, 118, "-" * 120)            # line 4 -- the separator
    put(page, 127, "---")                # its wrapped tail

    y = 140
    for line in students:
        put(page, y, f"  {line}")
        put(page, y + 12, "00)")                                   # wrapped tail
        put(page, y + 24, EXCLUDED)
        put(page, y + 36, "   DHSC-821   21FA2  A  Yes  No     3.00   12.00000     3.00   12.00000")
        put(page, y + 48, "   DHSC-831   22SP2       No    Section skipped - No Verified Grade Exists")
        put(page, y + 60, "   DHSC-833   21FA2  A  Yes  No     3.00   12.00000     3.00   12.00000")
        y += 78
    return page


report_page(1, [
    f"1237906 555-11-1427 Mrs. Birtukhan D. Abate      {AP}",
    f"1235956 622-33-4573 Markeece Alexander           {AP}",
    f"1234570 555-12-4570 Mroz Dana K                  {AP}",
])
report_page(2, [
    f"1234569 XXX-XX-4569 Mr. Robert A. B. Chen        {AP}",
    f"1234572 555124572 Ms. Grace T. Oyelaran          {AP}",
    f"1234575 555-12-4575 Dr.Samuel P. Okonkwo Jr.     {AP}",
])

out_path = str(Path(__file__).parent / "sample_sap_audit_structured_synthetic.pdf")
doc.save(out_path)
doc.close()
print("wrote", out_path)
