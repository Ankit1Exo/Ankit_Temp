"""Generate a synthetic (fabricated, non-real) ADP "W-2 and Earnings Summary"
test PDF used only to validate "260815 AM adp w2 ssn name address extractor.py".
No real PII: every SSN, name and address below is invented.

This reproduces the two conditions that make the real layout hard:

1. TWO PANES AT OVERLAPPING HEIGHTS. The W-2 prints on the left, ADP's
   Earnings Summary on the right. Right-pane text is placed at exactly the y
   of the employee's name, street and city lines, so a line rebuilt by
   vertical position contains both:

       printed (two panes)                 read as one line
       123 SAMPLE ST      Gross Pay ...    123 SAMPLE ST Gross Pay 104,485.46

   The right pane also repeats the employee's name and address and carries
   "Social Security Number: 900-00-0002" under "3. Employee W-4 Profile" --
   a DELIBERATE DECOY. It is a different number from the real one, so if the
   pane isolation regresses, the wrong SSN appears in the output.

2. TWO BOXES SHARING ONE ROW. "b Employer's FED ID number" and
   "a Employee's SSA number" print side by side with their values on one
   line, so the extractor must pick the value under the SSA caption by
   x-position rather than taking the first number it sees. The FED ID is
   45-1234567 -- if that appears in the SSN column, that pairing broke.

Pages:
    1  employee A, 2-line address, "FIRST M LAST" name, plus two tear-off
       copies across the bottom repeating the same employee (must collapse
       to one row)
    2  employee B, 3-line address with an APT line, "LAST, FIRST M" name
    3  employee C, name and address present but the SSA box left blank
       (must still produce a row, with the SSN column empty)
    4  TWO DIFFERENT employees printed side by side at the same height, with a
       banner spanning both so no gutter separates them. Both must appear, each
       with its own SSN and address -- taking only the first caption on the
       rebuilt line, or reading the block across the whole merged pane, loses
       or corrupts one of them.

Usage:
    python tests/make_test_adp_w2.py
"""
from pathlib import Path

import pymupdf as fitz

PAGE_W, PAGE_H = 792, 612          # landscape letter, as ADP prints these
LEFT_X = 30                        # left (W-2) pane
LEFT_MAX_X = 258
RIGHT_X = 280                      # right (Earnings Summary) pane
CAPTION_SIZE = 5.5
VALUE_SIZE = 8
FONT = "helv"

EMPLOYER_FED_ID = "45-1234567"     # 2-7; must never land in the SSN column


def put(page, x, y, text, size=CAPTION_SIZE):
    page.insert_text((x, y), text, fontsize=size, fontname=FONT)


def put_left(page, x, y, text, size=CAPTION_SIZE):
    """Place left-pane text, asserting it stays inside the W-2 pane -- otherwise
    the synthetic page would not have the printed gutter the real one has, and
    the test would be validating the wrong geometry."""
    width = fitz.get_text_length(text, fontname=FONT, fontsize=size)
    assert x + width <= LEFT_MAX_X, (
        f"left-pane text overflows the W-2 pane ({x + width:.0f} > {LEFT_MAX_X}): {text!r}")
    put(page, x, y, text, size)


def w2_block(page, x, y, ssn, name, street_lines, city_line, put_fn=None, width_limit=None):
    """One W-2's identity boxes: employer name/address, then box e/f, then the
    'b Employer's FED ID number | a Employee's SSA number' row."""
    place = put_fn or (lambda *a, **k: put(page, *a, **k))

    place(x, y, "c  Employer's name, address, and ZIP code")
    place(x + 4, y + 12, "SAMPLE EMPLOYER INC", VALUE_SIZE)
    place(x + 4, y + 24, "1180 SAMPLE ROAD", VALUE_SIZE)
    place(x + 4, y + 36, "SAMPLE JUNCTION IN 47546", VALUE_SIZE)
    place(x, y + 52, "Batch  #02816")

    place(x, y + 68, "e/f  Employee's name, address, and ZIP code")
    yy = y + 82
    place(x + 4, yy, name, VALUE_SIZE)
    for line in street_lines:
        yy += 13
        place(x + 4, yy, line, VALUE_SIZE)
    yy += 13
    place(x + 4, yy, city_line, VALUE_SIZE)

    # The two boxes that share a row. The SSA caption starts well to the right
    # of the FED ID caption, which is the only thing distinguishing their
    # values on the single rebuilt line below.
    row = yy + 20
    place(x, row, "b  Employer's FED ID number")
    place(x + 120, row, "a  Employee's SSA number")
    place(x + 2, row + 13, EMPLOYER_FED_ID, VALUE_SIZE)
    if ssn:
        place(x + 122, row + 13, ssn, VALUE_SIZE)

    place(x, row + 30, "1  Wages, tips, other comp.")
    place(x + 120, row + 30, "2  Federal income tax withheld")
    place(x + 2, row + 43, "96724.07", VALUE_SIZE)
    place(x + 122, row + 43, "15530.34", VALUE_SIZE)
    return row + 55


def earnings_summary(page, name, city_line, decoy_ssn):
    """The right pane, at y values chosen to collide with the left pane's
    name/street/city lines."""
    put(page, RIGHT_X + 120, 40, "2018 W-2 and EARNINGS SUMMARY", 12)
    put(page, RIGHT_X, 62,
        "This blue Earnings Summary section is included with your W-2 to help describe portions in more detail.")
    put(page, RIGHT_X, 74,
        "The reverse side includes general information that you may also find helpful.")
    put(page, RIGHT_X, 92,
        "1. The following information reflects your final 2018 pay stub plus any adjustments submitted.")

    put(page, RIGHT_X + 10, 110, "Gross Pay", VALUE_SIZE)
    put(page, RIGHT_X + 120, 110, "104485.46", VALUE_SIZE)
    put(page, RIGHT_X + 220, 110, "Social Security Tax Withheld", VALUE_SIZE)
    put(page, RIGHT_X + 380, 110, "6275.13", VALUE_SIZE)

    put(page, RIGHT_X, 130,
        "2. Your Gross Pay was adjusted as follows to produce your W-2 Statement.")
    # Rows placed at the same heights as the employee's name/street/city lines.
    for k, ry in enumerate((152, 165, 178, 191, 204)):
        put(page, RIGHT_X + 10, ry, ["Gross Pay", "Plus GTL (C-Box 12)", "Less 401(k) (D-Box 12)",
                                     "Less Other Cafe 125", "Reported W-2 Wages"][k], VALUE_SIZE)
        for col, cx in enumerate((150, 250, 350, 450)):
            put(page, RIGHT_X + cx, ry, ["104485.46", "108.03", "4487.68", "3243.96",
                                         "96724.07"][k], VALUE_SIZE)

    put(page, RIGHT_X, 226, "3. Employee W-4 Profile.  To change your Employee W-4 Profile Information,")
    put(page, RIGHT_X + 10, 246, name, VALUE_SIZE)
    put(page, RIGHT_X + 10, 259, city_line, VALUE_SIZE)
    # Decoy: the summary's own copy of the SSN, a DIFFERENT number, with no
    # "Employee's" prefix on its label.
    put(page, RIGHT_X + 260, 246, "Social Security Number:", VALUE_SIZE)
    put(page, RIGHT_X + 380, 246, decoy_ssn, VALUE_SIZE)
    put(page, RIGHT_X + 260, 259, "Taxable Marital Status:  SINGLE", VALUE_SIZE)


def add_page(doc, ssn, name, street_lines, city_line, decoy_ssn, bottom_copies=0):
    page = doc.new_page(width=PAGE_W, height=PAGE_H)
    put_left(page, LEFT_X, 40, "Employee Reference Copy")
    put_left(page, LEFT_X, 54, "W-2  Wage and Tax Statement  2018", 8)
    put_left(page, LEFT_X, 68, "d  Control number")
    put_left(page, LEFT_X + 4, 80, "000323 CLEV/4E1 000201", VALUE_SIZE)

    w2_block(page, LEFT_X, 96, ssn, name, street_lines, city_line,
             put_fn=lambda x, y, t, s=CAPTION_SIZE: put_left(page, x, y, t, s))

    earnings_summary(page, name, city_line, decoy_ssn)

    # Tear-off copies across the bottom, repeating the SAME employee. The
    # middle and right ones share a pane (the summary text above spans their
    # gutter), which is exactly the case de-duplication has to absorb.
    for k in range(bottom_copies):
        x = LEFT_X + k * 250
        put(page, x, 400, f"Copy {'BC2'[k]}  For employee's records.")
        w2_block(page, x, 412, ssn, name, street_lines, city_line)

    return page


def main():
    doc = fitz.open()

    add_page(doc,
             ssn="111-11-1111",
             name="RIVERA P JORDAN",
             street_lines=["122 SAMPLE DR"],
             city_line="SAMPLE TOWN CA 92336",
             decoy_ssn="900-00-0002",
             bottom_copies=2)

    add_page(doc,
             ssn="222-22-2222",
             name="OKONKWO, ADAEZE M",
             street_lines=["78 SAMPLE LANE", "APT 4B"],
             city_line="SAMPLE BOROUGH NJ 07030",
             decoy_ssn="900-00-0003")

    add_page(doc,
             ssn="",                       # SSA box left blank
             name="ALVES MARIA",
             street_lines=["9 SAMPLE CT"],
             city_line="SAMPLE VILLAGE NY 10001",
             decoy_ssn="900-00-0004")

    # Page 4: two different employees side by side, with a banner above that
    # spans both so the gutter between them is painted over.
    page = doc.new_page(width=PAGE_W, height=PAGE_H)
    put(page, LEFT_X, 40,
        "2018 W-2 and Earnings Summary - continuation sheet - multiple employees per page", 8)
    w2_block(page, LEFT_X, 60, "333-33-3333", "PARK MIN-JUN",
             ["9 SAMPLE CT"], "SAMPLE VILLAGE NY 10001")
    w2_block(page, LEFT_X + 250, 60, "444-44-4444", "SILVA, JOAO P",
             ["44 SAMPLE WAY", "STE 300"], "SAMPLE CITY FL 33101")

    out = Path(__file__).with_name("sample_adp_w2_synthetic.pdf")
    doc.save(out)
    doc.close()
    print(f"wrote {out}")
    print("expected: 5 rows -- 111-11-1111, 222-22-2222, one with a blank SSN, "
          "333-33-3333 and 444-44-4444.")
    print(f"never expected in the SSN column: {EMPLOYER_FED_ID} (employer FED ID) "
          f"or 900-00-000x (Earnings Summary decoys).")


if __name__ == "__main__":
    main()
