"""
extract_file_status.py
-----------------------
Opens a file-picker dialog so you can select your PDF, then extracts
'File:' numbers and 'Status:' values for every employee in a Personnel
Tax Jurisdictional Status report.

Results are printed to the console AND saved as a CSV next to this script.

Requirements (install once):
    pip install pdfplumber pdf2image pytesseract pillow
    # Linux also needs:  sudo apt-get install tesseract-ocr poppler-utils
    # macOS also needs:  brew install tesseract poppler

Run:
    python extract_file_status.py
"""

import os
import re
import csv
import sys


# ── file picker ───────────────────────────────────────────────────────────────

def pick_pdf() -> str:
    """
    Open a native file-chooser dialog filtered to PDF files.
    Falls back to a console prompt if no display is available.
    """
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()                        # hide the blank Tk window
        root.attributes("-topmost", True)      # bring dialog to front

        path = filedialog.askopenfilename(
            title="Select the Personnel Tax PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        root.destroy()

        if not path:
            sys.exit("No file selected. Exiting.")
        return path

    except Exception:
        # Headless / no display → fall back to typed input
        path = input("Enter the full path to your PDF file: ").strip()
        if not os.path.exists(path):
            sys.exit(f"File not found: {path}")
        return path


# ── PDF text extraction ───────────────────────────────────────────────────────

def extract_text_from_pdf(pdf_path: str) -> str:
    """
    Try direct text extraction first (fast, works for digital PDFs).
    If the pages come back blank, fall back to OCR (for scanned PDFs).
    """
    import pdfplumber

    text_pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text_pages.append(page.extract_text() or "")

    full_text = "\n".join(text_pages)

    # If almost no text was extracted the PDF is likely a scanned image
    if len(full_text.strip()) < 50:
        print("  → No text layer found. Running OCR (this may take a moment)...")
        full_text = ocr_pdf(pdf_path)

    return full_text


def ocr_pdf(pdf_path: str) -> str:
    """Convert each PDF page to an image and run Tesseract OCR on it."""
    import pytesseract
    from pdf2image import convert_from_path
    from PIL import ImageEnhance, ImageFilter

    pages = convert_from_path(pdf_path, dpi=300)
    all_text = []

    for i, img in enumerate(pages, start=1):
        print(f"  OCR page {i}/{len(pages)} ...")
        # Upscale + sharpen for better accuracy
        img = img.resize((img.width * 2, img.height * 2), img.LANCZOS)
        img = ImageEnhance.Contrast(img).enhance(2.0)
        img = img.filter(ImageFilter.SHARPEN)
        all_text.append(pytesseract.image_to_string(img, config="--psm 6"))

    return "\n".join(all_text)


# ── value normalisation ───────────────────────────────────────────────────────

def normalise_status(raw: str) -> str:
    r = re.sub(r"[^A-Za-z]", "", raw).upper()
    if re.search(r"TERMIN|TERMIAZ|TERHBI|FERT|TEED", r):
        return "TERMINATED"
    if re.search(r"ACTIV|SAUTER", r):
        return "ACTIVE"
    if r in ("GENDER", "MALE", "FEMALE", ""):
        return "REDACTED"
    return r


def normalise_name(raw: str) -> str:
    cleaned = re.sub(r"^[\[\]*\s\"']+", "", raw).strip()
    return cleaned if re.search(r"[A-Z][a-z]{2,}|[A-Z]{4,}", cleaned) else "REDACTED"


def normalise_file(raw: str) -> str:
    """
    A real file number is purely numeric or starts with a digit.
    OCR noise from redacted fields (e.g. 'ase', 'Beer?') is flagged REDACTED.
    """
    cleaned = re.sub(r"^[^A-Za-z0-9]+", "", raw).rstrip("-,").strip()
    if re.fullmatch(r"\d+", cleaned):              # all digits
        return cleaned
    if re.fullmatch(r"\d[A-Za-z0-9\-]{3,}", cleaned):  # starts with digit
        return cleaned
    return "REDACTED"


# ── regex patterns ────────────────────────────────────────────────────────────

# Handles OCR typo "Siatu" and Tesseract smart-quote artefact (U+201C)
STATUS_RE = re.compile(
    u"(?:Status|Siatu)[r:\\s]+[\u201c\u201d\"']?([A-Z]+)",
    re.IGNORECASE,
)
FILE_RE   = re.compile(r"\bFile[:\s]+(\S+)", re.IGNORECASE)
NAME_RE   = re.compile(r"^[\[\s*\"']*(.+?)\s+(?:Status|Siatu)", re.IGNORECASE)
SKIP_FILE = re.compile(r"^(SSN|DEPT|On|Data|Qualified)$", re.IGNORECASE)


# ── record parser ─────────────────────────────────────────────────────────────

def extract_records(text: str) -> list:
    """
    Parse OCR/extracted text into employee records.

    Layout per employee block:
        Line N  :  LASTNAME,FIRSTNAME ...  Status: <VALUE>  Gender: ...
        Line N+1:  File: <NUMBER>  SSN: On File  DEPT: ...
    """
    lines   = text.splitlines()
    records = []

    for i, line in enumerate(lines):
        sm = STATUS_RE.search(line)
        if not sm:
            continue
        nm = NAME_RE.match(line.strip())
        if not nm:
            continue  # skip page-title line "Jurisdictional Status"

        status = normalise_status(sm.group(1))
        name   = normalise_name(nm.group(1))

        # Look ahead up to 3 lines for "File: <number>"
        file_num = "NOT FOUND"
        for j in range(i + 1, min(i + 4, len(lines))):
            nxt = lines[j].strip()
            if not nxt:
                continue
            fm = FILE_RE.search(nxt)
            if fm:
                cand = fm.group(1).rstrip(",.").strip()
                if not SKIP_FILE.match(cand):
                    file_num = normalise_file(cand)
                break

        records.append({"Name": name, "File": file_num, "Status": status})

    return records


# ── output helpers ────────────────────────────────────────────────────────────

def print_table(records: list) -> None:
    if not records:
        print("\nNo records found. Check the PDF or OCR quality.")
        return

    wn = max(4, max(len(r["Name"])   for r in records)) + 2
    wf = max(4, max(len(r["File"])   for r in records)) + 2
    ws = max(6, max(len(r["Status"]) for r in records)) + 2
    sep = "-" * (wn + wf + ws + 6)

    print(f"\n{sep}")
    print(f"  {'Name':<{wn}}  {'File':<{wf}}  {'Status':<{ws}}")
    print(sep)
    for r in records:
        print(f"  {r['Name']:<{wn}}  {r['File']:<{wf}}  {r['Status']:<{ws}}")
    print(sep)
    print(f"  Total records: {len(records)}\n")


def save_csv(records: list, out_path: str) -> None:
    with open(out_path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["Name", "File", "Status"])
        writer.writeheader()
        writer.writerows(records)
    print(f"CSV saved: {out_path}")


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    print("=" * 55)
    print("  Personnel Tax — File & Status Extractor")
    print("=" * 55)

    # Open file picker
    pdf_path = pick_pdf()
    print(f"\nSelected: {pdf_path}")

    # Extract text (direct or OCR)
    print("Extracting text from PDF...")
    text = extract_text_from_pdf(pdf_path)

    # Parse records
    all_records = extract_records(text)

    # De-duplicate (same employee appearing on multiple pages)
    seen, unique = set(), []
    for r in all_records:
        k = (r["Name"], r["File"], r["Status"])
        if k not in seen:
            seen.add(k)
            unique.append(r)

    # Display
    print_table(unique)

    # Save CSV in the same folder as this script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_path   = os.path.join(script_dir, "extracted_results.csv")
    save_csv(unique, csv_path)


if __name__ == "__main__":
    main()
