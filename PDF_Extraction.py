"""
ADP Master Control PII Extractor
---------------------------------
Extracts employee records from ADP Master Control PDF reports.

SECURITY / COMPLIANCE NOTES:
  * Output contains SSN, DOB, account numbers, routing numbers.
  * Store output ONLY in an encrypted, access-controlled location.
  * Do NOT email, upload to chat tools, or place on shared drives without authorization.
  * Comply with ISO 27001 / SOC 2 / GLBA / applicable state privacy law.
  * Delete intermediate files securely when finished.

Requirements:
    pip install pdfplumber pandas

Usage:
    python extract_adp_pii.py file1.pdf file2.pdf -o output.xlsx
"""

import argparse
import re
from pathlib import Path
from typing import Optional

import pdfplumber
import pandas as pd

# ---------------------------------------------------------------------------
# Column X-coordinate boundaries for ADP Master Control layout.
# These are approximate; tune by inspecting one page with pdfplumber if needed.
# ---------------------------------------------------------------------------
COLUMN_BOUNDS = {
    "PERSONNEL":       (0,   270),
    "PAY":             (270, 460),
    "TAX_STATUS":      (460, 660),
    "SCHEDULED":       (660, 920),
    "ACCUMULATIONS":   (920, 1300),
}

# Regex patterns
RE_FILE        = re.compile(r"File:\s*(\d+)")
RE_HIRE        = re.compile(r"Hire:\s*(\d{2}/\d{2}/\d{4})")
RE_BIRTH       = re.compile(r"Birth:\s*([\d/Xx\-\*]+)")
RE_SSN         = re.compile(r"SSN:\s*([\w\-\*Xx ]+?)(?:\n|$)")
RE_CONTINUED   = re.compile(r"\(continued\)", re.IGNORECASE)
RE_ACCT        = re.compile(r"Acct\s*#:\s*([A-Za-z0-9\*Xx\-]+)")
RE_TRAN        = re.compile(r"Tran/ABA:\s*([A-Za-z0-9\*Xx\-]+)")
RE_NAME_HEADER = re.compile(r"^([A-Z][A-Z\-\' ]+,\s*[A-Z][A-Za-z\-\' ]+(?:\s+[A-Z])?)\s*$")


def get_column_text(page, col_name: str) -> str:
    """Crop a page to one column and return text in reading order."""
    x0, x1 = COLUMN_BOUNDS[col_name]
    # Clamp to page width
    x1 = min(x1, page.width)
    cropped = page.crop((x0, 0, x1, page.height))
    return cropped.extract_text(x_tolerance=2, y_tolerance=3) or ""


def split_into_employee_blocks(personnel_text: str) -> list[dict]:
    """
    Split the PERSONNEL column text into per-employee blocks.

    An employee block starts at a name line (LASTNAME,FIRSTNAME ...) and
    continues until the next name line. We also capture whether the block
    is marked '(continued)'.
    """
    lines = [ln.rstrip() for ln in personnel_text.splitlines()]
    blocks = []
    current = None

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # Detect a new employee name header.
        # Heuristic: ALL-CAPS-ish "LAST,FIRST" pattern; ignore section headers.
        if (
            "," in stripped
            and stripped.split(",")[0].isupper()
            and len(stripped.split(",")[0]) >= 2
            and "ADDRESS" not in stripped.upper()
            and "PERSONNEL" not in stripped.upper()
        ):
            if current:
                blocks.append(current)
            current = {"name": stripped, "raw_lines": [stripped], "continued": False}
            continue

        if current is None:
            continue

        current["raw_lines"].append(stripped)
        if RE_CONTINUED.search(stripped):
            current["continued"] = True

    if current:
        blocks.append(current)

    return blocks


def parse_personnel_block(block: dict) -> dict:
    """Extract fields from a single employee's PERSONNEL block."""
    text = "\n".join(block["raw_lines"])
    record = {
        "Name": block["name"],
        "Continued": block["continued"],
        "File_Number": None,
        "SSN": None,
        "DOB": None,
        "Hire_Date": None,
        "Mailing_Address": None,
        "Home_Address": None,
    }

    m = RE_FILE.search(text);  record["File_Number"] = m.group(1) if m else None
    m = RE_BIRTH.search(text); record["DOB"]         = m.group(1) if m else None
    m = RE_HIRE.search(text);  record["Hire_Date"]   = m.group(1) if m else None
    m = RE_SSN.search(text);   record["SSN"]         = m.group(1).strip() if m else None

    # Addresses: collect the lines that appear AFTER "Mailing Address" and
    # AFTER "Home Address" markers, up to the next labeled section.
    record["Mailing_Address"] = _extract_address(block["raw_lines"], "Mailing Address")
    record["Home_Address"]    = _extract_address(block["raw_lines"], "Home Address")

    return record


def _extract_address(lines: list[str], marker: str) -> Optional[str]:
    """Pull the 1-3 lines immediately following an address marker line."""
    collected = []
    capture = False
    stop_tokens = ("File:", "Dept:", "SSN:", "Cost:", "Status:", "Sex:",
                   "Mailing Address", "Home Address", "Dates", "Hire:", "Term:")

    for ln in lines:
        if marker in ln:
            capture = True
            continue
        if capture:
            if not ln.strip():
                if collected:
                    break
                continue
            if any(tok in ln for tok in stop_tokens):
                break
            collected.append(ln.strip())
            if len(collected) >= 3:   # safety cap
                break

    return ", ".join(collected) if collected else None


def parse_accumulations_block(text: str) -> list[dict]:
    """
    Pull Acct # and Tran/ABA pairs from the ACCUMULATIONS column.
    A single employee may have multiple direct deposit accounts.
    """
    accts = RE_ACCT.findall(text)
    trans = RE_TRAN.findall(text)
    pairs = []
    for i in range(max(len(accts), len(trans))):
        pairs.append({
            "Account_Number": accts[i] if i < len(accts) else None,
            "Routing_Number": trans[i] if i < len(trans) else None,
        })
    return pairs or [{"Account_Number": None, "Routing_Number": None}]


def process_pdf(pdf_path: Path) -> list[dict]:
    """Process one PDF and return a list of employee records."""
    records = []
    print(f"  Opening {pdf_path.name} ...")

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            personnel_text     = get_column_text(page, "PERSONNEL")
            accumulations_text = get_column_text(page, "ACCUMULATIONS")

            employee_blocks = split_into_employee_blocks(personnel_text)
            if not employee_blocks:
                continue

            # Direct-deposit info is per-page; we associate it with the first
            # non-continued employee on the page when possible. For pages with
            # multiple employees this is a best-effort mapping — review output.
            deposit_pairs = parse_accumulations_block(accumulations_text)

            for i, block in enumerate(employee_blocks):
                rec = parse_personnel_block(block)
                rec["Source_File"] = pdf_path.name
                rec["Page"]        = page_idx

                # Attach deposit info to the first block on the page only,
                # to avoid duplicating across multiple employees on same page.
                if i == 0 and deposit_pairs:
                    for dp in deposit_pairs:
                        merged = {**rec, **dp}
                        records.append(merged)
                else:
                    records.append({**rec,
                                    "Account_Number": None,
                                    "Routing_Number": None})

            if page_idx % 50 == 0:
                print(f"    ...processed page {page_idx}")

    return records


def main():
    ap = argparse.ArgumentParser(description="Extract PII from ADP Master Control PDFs.")
    ap.add_argument("pdfs", nargs="+", help="Path(s) to PDF file(s)")
    ap.add_argument("-o", "--output", default="adp_extracted.xlsx",
                    help="Output XLSX file (default: adp_extracted.xlsx)")
    args = ap.parse_args()

    print("=" * 60)
    print("ADP PII EXTRACTION  --  HANDLE OUTPUT AS CONFIDENTIAL")
    print("=" * 60)

    all_records = []
    for p in args.pdfs:
        path = Path(p)
        if not path.exists():
            print(f"  SKIP (not found): {p}")
            continue
        all_records.extend(process_pdf(path))

    if not all_records:
        print("No records extracted.")
        return

    # Column order with Source_File first per request
    cols = ["Source_File", "Page", "Continued", "File_Number", "Name",
            "DOB", "SSN", "Hire_Date",
            "Mailing_Address", "Home_Address",
            "Account_Number", "Routing_Number"]
    df = pd.DataFrame(all_records)
    df = df.reindex(columns=cols)

    out = Path(args.output)
    if out.suffix.lower() == ".csv":
        df.to_csv(out, index=False)
    else:
        df.to_excel(out, index=False)

    print(f"\n  Extracted {len(df)} rows -> {out.resolve()}")
    print("  REMINDER: Output contains sensitive PII. Store in encrypted, access-controlled location.")


if __name__ == "__main__":
    main()
