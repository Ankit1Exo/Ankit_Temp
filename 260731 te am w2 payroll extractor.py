"""
260731 te am w2 payroll extractor.py

Extracts employee data from two kinds of text-based (selectable) PDF layouts and
writes the results to a single Excel file.

FORMAT 1 - "W-2 style" (2 forms printed side-by-side per page, 4-up W-2 layout)
    Extracts: SSN, First Name, Middle Initial, Last Name, Street, City, State, Zip

FORMAT 2/3 - "Payroll list style" (one or more employees listed per page, with or
    without column headers). Only SSN and Name are on these pages.
    Extracts: SSN, First Name, Middle Initial, Last Name  (address left blank)

USAGE
    pip install pdfplumber openpyxl
    python "260731 te am w2 payroll extractor.py" file1.pdf file2.pdf -o output.xlsx

NOTES / ASSUMPTIONS (please review and tell me if any of these don't match reality
once you run this against a real file - I built this from sample text you gave me,
not an actual PDF, so the exact spatial layout logic (grouping words into lines,
left/right half split) may need small tweaks):

  - Format 1 is detected by the presence of "W-2" / "Wage and Tax Statement" text
    on the page. Each such page is split into a LEFT half and a RIGHT half at the
    horizontal midpoint of the page, and each half is parsed as one employee record.
  - Within a half, words are grouped into "lines" by clustering on their vertical
    (top) coordinate. This reconstructs the visual line order regardless of the
    raw order the PDF stores text runs in.
  - The SSN for Format 1 is the first XXX-XX-XXXX (or masked) pattern found on/near
    the "Employee's social security number" label, distinguished from the Employer
    ID (which is XX-XXXXXXX format).
  - The Name+Address block for Format 1 is taken as the line(s) between the
    "...first name and initial...Last name" label and the "...address and ZIP
    code" label. First non-empty line = name. Last non-empty line (matching a
    CITY STATE ZIP pattern) = city/state/zip. Anything in between = street
    (joined with a space if there are 2 street lines, e.g. Street + Apt/Suite).
  - Format 2/3 detection & extraction does NOT rely on headers at all (since
    Format 3 has none). Instead every line on non-W-2 pages is scanned for the
    pattern: SSN (digits separated by spaces or dashes) followed by a name.
    This naturally works whether a header row is present or not.
  - Name splitting rule (applies to both formats):
        "John J Doe"   -> First=John, MI=J,  Last=Doe
        "John Doe Jr"  -> First=John, MI='', Last=Doe Jr   (2nd word isn't a
                           single letter, so it's treated as part of Last)
  - SSN is written to Excel as raw digits only (no dashes/spaces).

If your real PDFs turn out to extract in a different order/layout than assumed
here, send me one redacted/dummy-data sample PDF (not real SSNs) and I'll adjust
the parsing logic to match exactly.
"""

import argparse
import re
import sys
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font

# ---------------------------------------------------------------------------
# Regex patterns
# ---------------------------------------------------------------------------
SSN_DASHED_RE = re.compile(r'\b(\d{3})[-\s](\d{2})[-\s](\d{4})\b')
EIN_RE = re.compile(r'\b\d{2}-\d{7}\b')
CITY_STATE_ZIP_RE = re.compile(r'^(?P<city>.+?)\s+(?P<state>[A-Z]{2})\s+(?P<zip>\d{5}(-\d{4})?)$')

W2_KEYWORDS = ("wage and tax statement", "form w-2", "w-2")
SSN_LABEL_RE = re.compile(r"employee'?s social security number", re.I)
NAME_LABEL_RE = re.compile(r"first name and initial", re.I)
ADDRESS_LABEL_RE = re.compile(r"employee'?s address", re.I)


# ---------------------------------------------------------------------------
# Line reconstruction helpers
# ---------------------------------------------------------------------------
def words_to_lines(words, y_tolerance=3):
    """Group pdfplumber word dicts into visual lines based on 'top' coordinate."""
    if not words:
        return []
    words = sorted(words, key=lambda w: (round(w["top"] / y_tolerance), w["x0"]))
    lines = []
    current_top = None
    current_words = []
    for w in words:
        if current_top is None or abs(w["top"] - current_top) > y_tolerance:
            if current_words:
                lines.append(current_words)
            current_words = [w]
            current_top = w["top"]
        else:
            current_words.append(w)
    if current_words:
        lines.append(current_words)

    result = []
    for line_words in lines:
        line_words = sorted(line_words, key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in line_words)
        top = min(w["top"] for w in line_words)
        result.append({"text": text, "top": top, "words": line_words})
    return result


def split_page_halves(page):
    """Split a page's words into left-half and right-half word lists."""
    mid_x = page.width / 2
    words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    left = [w for w in words if w["x0"] < mid_x]
    right = [w for w in words if w["x0"] >= mid_x]
    return left, right


# ---------------------------------------------------------------------------
# Name splitting
# ---------------------------------------------------------------------------
def split_name(name_str):
    """
    'John J Doe'  -> ('John', 'J', 'Doe')
    'John Doe Jr' -> ('John', '', 'Doe Jr')
    """
    parts = name_str.split()
    if not parts:
        return "", "", ""
    if len(parts) == 1:
        return parts[0], "", ""
    second = parts[1].strip(".")
    if len(parts) >= 3 and len(second) == 1 and second.isalpha():
        first = parts[0]
        mi = second
        last = " ".join(parts[2:])
    else:
        first = parts[0]
        mi = ""
        last = " ".join(parts[1:])
    return first, mi, last


def clean_ssn(raw):
    digits = re.sub(r"\D", "", raw)
    return digits


# ---------------------------------------------------------------------------
# Format 1 (W-2 style) parsing
# ---------------------------------------------------------------------------
def parse_format1_half(lines):
    """lines: list of {'text', 'top', 'words'} for one half of a W-2 page."""
    record = {"ssn": "", "first": "", "mi": "", "last": "",
              "street": "", "city": "", "state": "", "zip": ""}

    ssn_idx = next((i for i, l in enumerate(lines) if SSN_LABEL_RE.search(l["text"])), None)
    if ssn_idx is not None:
        # look on the label line itself, then following lines, skipping any EIN match
        for i in range(ssn_idx, min(ssn_idx + 3, len(lines))):
            candidates = [m for m in SSN_DASHED_RE.finditer(lines[i]["text"])
                          if not EIN_RE.search(m.group(0))]
            if candidates:
                m = candidates[0]
                record["ssn"] = clean_ssn(m.group(0))
                break

    name_idx = next((i for i, l in enumerate(lines) if NAME_LABEL_RE.search(l["text"])), None)
    addr_idx = next((i for i, l in enumerate(lines) if ADDRESS_LABEL_RE.search(l["text"])), None)

    if name_idx is not None:
        end = addr_idx if addr_idx is not None and addr_idx > name_idx else len(lines)
        block_lines = [l["text"].strip() for l in lines[name_idx + 1:end] if l["text"].strip()]
        # drop the label line's own leftover text if the label and value share a line
        block_lines = [t for t in block_lines if not NAME_LABEL_RE.search(t) and not ADDRESS_LABEL_RE.search(t)]

        if block_lines:
            name_line = block_lines[0]
            record["first"], record["mi"], record["last"] = split_name(name_line)

            remaining = block_lines[1:]
            # last line matching CITY STATE ZIP is the city/state/zip line
            csz_i = None
            for i in range(len(remaining) - 1, -1, -1):
                if CITY_STATE_ZIP_RE.match(remaining[i].strip()):
                    csz_i = i
                    break
            if csz_i is not None:
                m = CITY_STATE_ZIP_RE.match(remaining[csz_i].strip())
                record["city"] = m.group("city").strip()
                record["state"] = m.group("state")
                record["zip"] = m.group("zip")
                street_lines = remaining[:csz_i]
            else:
                street_lines = remaining
            record["street"] = " ".join(s.strip() for s in street_lines if s.strip())

    return record


def is_w2_page(page_text):
    text_lower = page_text.lower()
    return any(kw in text_lower for kw in W2_KEYWORDS)


def process_format1_page(page):
    records = []
    for half_words in split_page_halves(page):
        lines = words_to_lines(half_words)
        if not lines:
            continue
        rec = parse_format1_half(lines)
        if rec["ssn"] or rec["last"]:
            records.append(rec)
    return records


# ---------------------------------------------------------------------------
# Format 2/3 (payroll list) parsing - no header dependency
# ---------------------------------------------------------------------------
FORMAT23_LINE_RE = re.compile(
    r'\b(\d{3})[-\s](\d{2})[-\s](\d{4})\b\s+(?P<name>[A-Z][A-Za-z.\-]*(?:\s+[A-Z][A-Za-z.\-]*){1,3})'
)


def process_format23_page(page):
    records = []
    text = page.extract_text() or ""
    for line in text.split("\n"):
        line = line.strip()
        if not line:
            continue
        m = FORMAT23_LINE_RE.search(line)
        if not m:
            continue
        ssn = clean_ssn(f"{m.group(1)}{m.group(2)}{m.group(3)}")
        name = m.group("name").strip()
        first, mi, last = split_name(name)
        records.append({"ssn": ssn, "first": first, "mi": mi, "last": last,
                         "street": "", "city": "", "state": "", "zip": ""})
    return records


# ---------------------------------------------------------------------------
# Main extraction
# ---------------------------------------------------------------------------
def extract_pdf(path):
    all_records = []
    with pdfplumber.open(path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            page_text = page.extract_text() or ""
            if is_w2_page(page_text):
                recs = process_format1_page(page)
            else:
                recs = process_format23_page(page)
            for r in recs:
                r["source_file"] = Path(path).name
                r["page"] = page_num
                all_records.append(r)
    return all_records


def write_excel(records, out_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Employees"
    headers = ["Source File", "Page", "SSN", "First Name", "MI", "Last Name",
               "Street", "City", "State", "Zip"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for r in records:
        ws.append([
            r.get("source_file", ""), r.get("page", ""), r.get("ssn", ""),
            r.get("first", ""), r.get("mi", ""), r.get("last", ""),
            r.get("street", ""), r.get("city", ""), r.get("state", ""), r.get("zip", ""),
        ])

    # basic column widths
    widths = [22, 6, 12, 14, 4, 16, 24, 16, 6, 10]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    wb.save(out_path)


def main():
    parser = argparse.ArgumentParser(description="Extract employee SSN/Name/Address from PDFs to Excel.")
    parser.add_argument("pdfs", nargs="+", help="Path(s) to input PDF file(s)")
    parser.add_argument("-o", "--output", default="employee_extract.xlsx", help="Output Excel file path")
    args = parser.parse_args()

    all_records = []
    for pdf_path in args.pdfs:
        print(f"Processing {pdf_path} ...")
        recs = extract_pdf(pdf_path)
        print(f"  -> {len(recs)} record(s) found")
        all_records.extend(recs)

    write_excel(all_records, args.output)
    print(f"Done. Wrote {len(all_records)} record(s) to {args.output}")


if __name__ == "__main__":
    main()
