"""Generate a synthetic (fabricated, non-real) Schedule K-1 (Form 1065) test PDF
used only to validate "260811 AM k1 partner ssn name address extractor.py".
No real PII: every TIN, name and address below is invented.

The point of this file is to reproduce the one condition that silently corrupts
a K-1 extraction -- Part II (partner identity, left column) and Part III
(boxes 1-20, right column) printed at OVERLAPPING vertical positions, so any
line rebuilt purely by y contains words from both columns:

    printed (two columns)                    read as one line
    JORDAN B RIVERA        7  Royalties      JORDAN B RIVERA 7 Royalties
    122 SAMPLE DR          8  Net short...   122 SAMPLE DR 8 Net short...

Right-column text is therefore placed at exactly the y of the partner's name,
street and city lines, and page 1 additionally carries a DELIBERATE DECOY:
a second, fabricated SSN-shaped value in the right column at the y of the
partner's name line. If the extractor's column isolation ever regresses, the
decoy lands in the Name column and the expected output no longer matches.

Pages:
    1  individual partner, 2-line address, "FIRST M LAST" name, dashed SSN,
       plus the right-column decoy SSN
    2  LLC partner, EIN in box E, 3-line address (suite line), entity name
       that must NOT be split into First/MI/Last
    3  individual partner, MASKED SSN (client-copy style "XXX-XX-1111") and a
       "LAST, FIRST M" comma-form name
    4  byte-for-byte repeat of page 1 (the "file copy") -- must de-duplicate
       away, leaving 3 rows
    5  the K-1 codes/instructions page -- no Part II heading, must yield no rows

Usage:
    python tests/make_test_k1.py
"""
from pathlib import Path

import pymupdf as fitz

PAGE_W, PAGE_H = 612, 792
LEFT_X = 36
LEFT_MAX_X = 300      # printed right edge of the Part II column
RIGHT_X = 318         # printed left edge of the Part III boxes column
CAPTION_SIZE = 6
VALUE_SIZE = 9
RIGHT_SIZE = 6.5
FONT = "helv"

# Right-column (Part III) lines, at y values chosen to collide with the Part II
# rows on the left. See the module docstring.
RIGHT_COLUMN = [
    (40, "Part III  Partner's Share of Current Year Income, Deductions, Credits"),
    (220, "4c  Total guaranteed payments"),
    (234, "5   Interest income                        1,250"),
    (249, "6a  Ordinary dividends"),
    (263, "7   Royalties"),
    (277, "8   Net short-term capital gain (loss)"),
    (291, "9a  Net long-term capital gain (loss)"),
    (305, "9b  Collectibles (28%) gain (loss)"),
    (319, "10  Net section 1231 gain (loss)"),
    (333, "11  Other income (loss)"),
    (347, "13  Other deductions"),
    (361, "20  Other information            Z*  STMT"),
]


def put(page, x, y, text, size):
    page.insert_text((x, y), text, fontsize=size, fontname=FONT)


def put_left(page, x, y, text, size):
    """Place left-column text, asserting it stays inside the Part II column --
    otherwise the synthetic page would not have the printed gutter that a real
    K-1 has, and the test would be validating the wrong geometry."""
    width = fitz.get_text_length(text, fontname=FONT, fontsize=size)
    assert x + width <= LEFT_MAX_X, (
        f"left-column text overflows the Part II column ({x + width:.0f} > {LEFT_MAX_X}): {text!r}")
    put(page, x, y, text, size)


def add_k1_page(doc, tin_text, name, street_lines, city_line, entity_type, decoy_ssn=None):
    page = doc.new_page(width=PAGE_W, height=PAGE_H)

    put_left(page, LEFT_X, 40, "Schedule K-1 (Form 1065)  2024", CAPTION_SIZE)

    # --- Part I: the PARTNERSHIP. Its name/address must never be recorded as
    # the partner's, and its EIN is what de-duplication keys on.
    put_left(page, LEFT_X, 60, "Part I   Information About the Partnership", CAPTION_SIZE)
    put_left(page, LEFT_X, 75, "A  Partnership's employer identification number", CAPTION_SIZE)
    put_left(page, LEFT_X, 88, "98-7654321", VALUE_SIZE)
    put_left(page, LEFT_X, 103, "B  Partnership's name, address, city, state, and ZIP code", CAPTION_SIZE)
    put_left(page, LEFT_X, 116, "SAMPLE PARTNERSHIP HOLDINGS LLC", VALUE_SIZE)
    put_left(page, LEFT_X, 129, "1074 SAMPLE AVE NE", VALUE_SIZE)
    put_left(page, LEFT_X, 142, "SAMPLE CITY GA 30307", VALUE_SIZE)
    put_left(page, LEFT_X, 157, "C  IRS Center where partnership filed return:", CAPTION_SIZE)
    put_left(page, LEFT_X, 170, "E-FILE", VALUE_SIZE)
    put_left(page, LEFT_X, 185, "D  Check if this is a publicly traded partnership (PTP)", CAPTION_SIZE)

    # --- Part II: the PARTNER. The box letters are placed as separate text
    # runs from the headings, matching documents where the letter is not a
    # structured element of the heading it precedes.
    put_left(page, LEFT_X, 205, "Part II   Information About the Partner", CAPTION_SIZE)
    put_left(page, LEFT_X, 220, "E", CAPTION_SIZE)
    put_left(page, LEFT_X + 10, 220,
             "Partner's SSN or TIN (Do not use TIN of a disregarded entity. See instructions.)",
             CAPTION_SIZE)
    put_left(page, LEFT_X + 10, 234, tin_text, VALUE_SIZE)

    put_left(page, LEFT_X, 249, "F", CAPTION_SIZE)
    put_left(page, LEFT_X + 10, 249,
             "Name, address, city, state, and ZIP code for partner entered in E. See instructions.",
             CAPTION_SIZE)

    y = 263
    put_left(page, LEFT_X + 10, y, name, VALUE_SIZE)
    for line in street_lines:
        y += 14
        put_left(page, LEFT_X + 10, y, line, VALUE_SIZE)
    y += 14
    put_left(page, LEFT_X + 10, y, city_line, VALUE_SIZE)

    y += 18
    put_left(page, LEFT_X, y, "G  General partner or LLC member-manager", CAPTION_SIZE)
    y += 12
    put_left(page, LEFT_X, y, "   Limited partner or other LLC member   X", CAPTION_SIZE)
    y += 14
    put_left(page, LEFT_X, y, "H1  Domestic partner   X       Foreign partner", CAPTION_SIZE)
    y += 14
    put_left(page, LEFT_X, y, "H2  If the partner is a disregarded entity (DE), enter the partner's:",
             CAPTION_SIZE)
    y += 12
    put_left(page, LEFT_X + 20, y, "TIN", CAPTION_SIZE)
    y += 12
    put_left(page, LEFT_X + 20, y, "Name", CAPTION_SIZE)
    y += 14
    put_left(page, LEFT_X, y, f"I1  What type of entity is this partner?   {entity_type}", CAPTION_SIZE)
    y += 14
    put_left(page, LEFT_X, y, "I2  If this partner is a retirement plan (IRA/SEP/Keogh/etc.), check here",
             CAPTION_SIZE)
    y += 14
    put_left(page, LEFT_X, y, "J  Partner's share of profit, loss, and capital (see instructions):",
             CAPTION_SIZE)
    y += 12
    put_left(page, LEFT_X + 20, y, "Beginning              Ending", CAPTION_SIZE)
    y += 12
    put_left(page, LEFT_X + 20, y, "17.094000              17.094000", CAPTION_SIZE)

    for ry, text in RIGHT_COLUMN:
        put(page, RIGHT_X, ry, text, RIGHT_SIZE)

    # Deliberate decoy: an SSN-shaped value in the RIGHT column at the y of the
    # partner's name line. Reachable only if column isolation fails.
    if decoy_ssn:
        put(page, RIGHT_X + 150, 263, decoy_ssn, VALUE_SIZE)

    return page


def add_codes_page(doc):
    """The K-1's second page: box codes and instructions, no Part II heading."""
    page = doc.new_page(width=PAGE_W, height=PAGE_H)
    put(page, LEFT_X, 50, "This list identifies the codes used on Schedule K-1 for all partners",
        CAPTION_SIZE)
    rows = [
        "1.  Ordinary business income (loss)          Determine whether the income (loss) is",
        "2.  Net rental real estate income (loss)     passive or nonpassive and enter on your",
        "3.  Other net rental income (loss)           return as follows:",
        "4a. Guaranteed payment services             Schedule E (Form 1040), line 28, column (k)",
        "5.  Interest income                          Form 1040 or 1040-SR, line 2b",
        "6a. Ordinary dividends                       Form 1040 or 1040-SR, line 3b",
        "13. Other deductions                         See the Partner's Instructions",
        "20. Other information                        Code Z  Section 199A information",
    ]
    y = 70
    for row in rows:
        put(page, LEFT_X, y, row, CAPTION_SIZE)
        y += 14
    return page


def main():
    doc = fitz.open()

    # Page 1 -- individual, dashed SSN, 2-line address, whitespace-split name,
    # plus the right-column decoy SSN.
    page1_args = dict(
        tin_text="111-11-1111",
        name="JORDAN B RIVERA",
        street_lines=["122 SAMPLE DR"],
        city_line="SAMPLE TOWN            PA 19444",
        entity_type="INDIVIDUAL",
        decoy_ssn="900-00-0001",
    )
    add_k1_page(doc, **page1_args)

    # Page 2 -- entity partner: EIN in box E, 3-line address, and a name that
    # must be left unsplit.
    add_k1_page(
        doc,
        tin_text="11-1111111",
        name="SAMPLE VENTURE I LLC",
        street_lines=["500 SAMPLE BLVD", "SUITE 1200"],
        city_line="SAMPLE CITY            TX 75201",
        entity_type="LIMITED LIABILITY COMPANY",
    )

    # Page 3 -- masked TIN and comma-form name.
    add_k1_page(
        doc,
        tin_text="XXX-XX-1111",
        name="OKONKWO, ADAEZE M",
        street_lines=["78 SAMPLE LANE APT 4B"],
        city_line="SAMPLE BOROUGH         NJ 07030",
        entity_type="INDIVIDUAL",
    )

    # Page 4 -- the file copy of page 1: must de-duplicate away.
    add_k1_page(doc, **page1_args)

    # Page 5 -- codes page: must yield no rows.
    add_codes_page(doc)

    out = Path(__file__).with_name("sample_k1_synthetic.pdf")
    doc.save(out)
    doc.close()
    print(f"wrote {out}")
    print("expected: 3 rows after de-duplication (pages 1, 2, 3); page 4 collapses into page 1; "
          "page 5 yields none")


if __name__ == "__main__":
    main()
