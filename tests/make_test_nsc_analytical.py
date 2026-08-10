"""Generate a synthetic "NSC Analytical Report" test PDF.

Every student ID, SSN, name and address below is FABRICATED. No real PII.

This report is laid out differently from the SAP audit report, and the
difference is the whole reason this fixture exists:

    Student ID  Name/Address              SSN         Enrollment Status  ...
    1234506     Alvarez, Ana              555-01-2345 H Half time or more
                118 Sanders Road Apt 81
                Beulaville, NC 28518

Three things here break a parser written for the SAP report:

    1. the NAME is printed to the LEFT of the SSN, not the right. To the
       right is the enrollment status -- so a parser that always reads the
       name from the right of the SSN fills the Name column with "Half time
       or more". It looks like a successful run: every cell has something in
       it. It only gives itself away as four distinct values repeated down
       thousands of rows.

    2. "Name/Address" is ONE heading over TWO fields, so the address has to
       be read down the column rather than across the row.

    3. the heading order is ID, Name, SSN -- so any parser that requires
       ID < SSN < Name finds no heading row at all and never measures the
       columns.

Also covered:
    - a student with NO SSN printed, who must still be recovered from the
      enrollment status caption rather than vanishing
    - a single-line address, and one with an apartment number
    - a name with a suffix, and one with a hyphen
    - per-word y jitter, as a real text layer carries
"""
from pathlib import Path

import pymupdf as fitz

SIZE = 8
FONT = "cour"
CHAR_W = SIZE * 0.6
LEFT_X = 30
JITTER = [0.0, 0.9, 0.3, 1.1, 0.2, 0.7]

HEAD = ("Student ID Name/Address              SSN        Enrollment Status"
        "      Status DT      AGD Begin Term  End Term")
RULE = ("---------- ------------------------- ---------- -----------------"
        "----- ---------- -------- ---------- ----------")

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


# Each entry is (student line, address lines...). The student line is written
# to match the heading above it column for column.
STUDENTS = [
    ("1234506    Alvarez, Ana              555-01-2345 H Half time or more"
     "                08/19/17 07/29/15   12/18/15",
     ["           118 Sanders Road Apt 81",
      "           Beulaville, NC 28518"]),
    ("1234508    Devries, Allison          555-01-2346 L Less than half time"
     "   12/31/16          07/29/15   12/18/15",
     ["           205 Green Height Road",
      "           Jacksonville, NC 28546"]),
    ("1234517    Richards, Robert          555-01-2347 H Half time or more"
     "                10/13/19 07/29/15   12/18/15",
     ["           64 Twelve Oaks Circle",
      "           Linden, NC 28356"]),
    ("1234521    Williams, Drew M          555-01-2348 L Less than half time"
     "   12/31/16          07/29/15   12/18/15",
     ["           9 Tarver St",
      "           Cameron, NC 28326"]),
    ("1234533    Okonkwo-Baptiste, Ruth Jr 555-01-2349 F Full time"
     "             12/31/18          07/29/15   12/18/15",
     ["           PO Box 2214, Sanford, NC 27330"]),
    # No SSN printed at all. The student must still come through, recovered
    # from the enrollment status, rather than disappearing from the workbook.
    ("1234540    Petrov, Yelena                        W Withdrawn"
     "            12/31/16          07/29/15   12/18/15",
     ["           77 Old Mill Trail",
      "           Spring Lake, NC 28390"]),
]


def report_page(page_no, students):
    page = doc.new_page(width=1080, height=140 + 34 * len(students))
    put(page, 24, "Oct 30 2015                    NSC Analytical Report")
    put(page, 33, f"Page {page_no}")
    put(page, 42, "16:17")
    put(page, 60, HEAD)
    put(page, 70, RULE)

    y = 84
    for line, address_lines in students:
        put(page, y, line)
        for i, address in enumerate(address_lines, start=1):
            put(page, y + 10 * i, address)
        y += 34
    return page


report_page(375, STUDENTS[:4])
report_page(376, STUDENTS[4:])

out_path = str(Path(__file__).parent / "sample_nsc_analytical_synthetic.pdf")
doc.save(out_path)
doc.close()
print("wrote", out_path)
