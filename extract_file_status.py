"""
extract_file_status.py
-----------------------
Extracts 'File:' numbers and 'Status:' values from a Personnel Tax
Jurisdictional Status document (image or multi-page PDF) using OCR.

Requirements:
    pip install pytesseract pillow pdf2image
    sudo apt-get install tesseract-ocr poppler-utils   (Linux)

Usage:
    python extract_file_status.py                         # default path
    python extract_file_status.py path/to/your/file.png  # custom image
    python extract_file_status.py path/to/your/doc.pdf   # PDF input

Notes:
    File numbers / names physically redacted in the source document are
    reported as "REDACTED". Status values with OCR noise are normalised
    to ACTIVE or TERMINATED.
"""

import re
import sys
import os
import csv

try:
    import pytesseract
    from PIL import Image, ImageEnhance, ImageFilter
except ImportError:
    sys.exit("Missing deps. Run: pip install pytesseract pillow")


# ── image helpers ─────────────────────────────────────────────────────────────

def load_images(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        try:
            from pdf2image import convert_from_path
            return convert_from_path(path, dpi=300)
        except ImportError:
            sys.exit("PDF support requires: pip install pdf2image")
    return [Image.open(path)]


def preprocess(img):
    """Upscale + contrast + sharpen for higher OCR accuracy."""
    img = img.resize((img.width * 2, img.height * 2), Image.LANCZOS)
    img = ImageEnhance.Contrast(img).enhance(2.0)
    img = img.filter(ImageFilter.SHARPEN)
    return img


def run_ocr(img):
    return pytesseract.image_to_string(preprocess(img), config="--psm 6")


# ── value normalisation ───────────────────────────────────────────────────────

def normalise_status(raw):
    r = re.sub(r"[^A-Za-z]", "", raw).upper()
    if re.search(r"TERMIN|TERMIAZ|TERHBI|FERT|TEED", r):
        return "TERMINATED"
    if re.search(r"ACTIV|SAUTER", r):
        return "ACTIVE"
    # "Gender" captured when status value itself is redacted/missing in source
    if r in ("GENDER", "MALE", "FEMALE", "UNKNOWN", ""):
        return "REDACTED"
    return r


def normalise_name(raw):
    cleaned = re.sub(r"^[\[\]*\s\"']+", "", raw).strip()
    # Must contain a recognisable surname: a word of 4+ letters starting with uppercase
    return cleaned if re.search(r"[A-Z][a-z]{3,}|[A-Z]{4,}", cleaned) else "REDACTED"


def normalise_file(raw):
    """
    Accept only tokens that look like real file numbers:
    - All digits (e.g. 12345)
    - Alphanumeric starting with a digit (e.g. 4A2B1)
    - 4 or more characters, no trailing dash
    Everything else is OCR noise from a redacted/struck-through field.
    """
    cleaned = re.sub(r"^[^A-Za-z0-9]+", "", raw).strip().rstrip("-")
    # Real file numbers typically start with a digit
    if re.fullmatch(r"\d[A-Za-z0-9]{3,}", cleaned):
        return cleaned
    # Or are purely numeric
    if re.fullmatch(r"\d{4,}", cleaned):
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


# ── core parser ───────────────────────────────────────────────────────────────

def extract_records(text):
    """
    Parse OCR output into [{Name, File, Status}] dicts.

    Document layout per employee:
        Line N  : LASTNAME,FIRSTNAME ...  Status: <VALUE>  Gender: ...
        Line N+1: File: <NUMBER>  SSN: On File  DEPT: ...
    """
    lines = text.splitlines()
    records = []

    for i, line in enumerate(lines):
        sm = STATUS_RE.search(line)
        if not sm:
            continue
        nm = NAME_RE.match(line.strip())
        if not nm:
            continue   # skip page title "Jurisdictional Status"

        status = normalise_status(sm.group(1))
        name   = normalise_name(nm.group(1))

        # Look ahead for "File: <number>" on the next non-blank line
        file_num = "NOT FOUND"
        for j in range(i + 1, min(i + 4, len(lines))):
            nxt = lines[j].strip()
            if not nxt:
                continue
            fm = FILE_RE.search(nxt)
            if fm:
                cand = fm.group(1).rstrip(",").strip()
                if not SKIP_FILE.match(cand):
                    file_num = normalise_file(cand)
                break

        records.append({"Name": name, "File": file_num, "Status": status})

    return records


# ── output ────────────────────────────────────────────────────────────────────

def print_table(records):
    if not records:
        print("\nNo records found. Check image quality or OCR output.")
        return
    wn = max(4, max(len(r["Name"])   for r in records)) + 2
    wf = max(4, max(len(r["File"])   for r in records)) + 2
    ws = max(6, max(len(r["Status"]) for r in records)) + 2
    sep = "-" * (wn + wf + ws + 6)
    print("\n" + sep)
    print(f"  {'Name':<{wn}}  {'File':<{wf}}  {'Status':<{ws}}")
    print(sep)
    for r in records:
        print(f"  {r['Name']:<{wn}}  {r['File']:<{wf}}  {r['Status']:<{ws}}")
    print(sep)
    print(f"  Total records: {len(records)}\n")


def save_csv(records, out_path):
    with open(out_path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["Name", "File", "Status"])
        writer.writeheader()
        writer.writerows(records)
    print(f"CSV saved: {out_path}")


# ── main ─────────────────────────────────────────────────────────────────────

def main():
    path = (
        sys.argv[1]
        if len(sys.argv) > 1
        else "/mnt/user-data/uploads/1782711811309_image.png"
    )
    if not os.path.exists(path):
        sys.exit(f"File not found: {path}")

    print(f"Processing: {path}")
    all_records = []
    for page_num, img in enumerate(load_images(path), start=1):
        if page_num > 1:
            print(f"  Page {page_num} ...")
        all_records.extend(extract_records(run_ocr(img)))

    # De-duplicate across pages (use all three fields so two different
    # REDACTED employees with the same File aren't collapsed into one)
    seen, unique = set(), []
    for r in all_records:
        k = (r["Name"], r["File"], r["Status"])
        if k not in seen:
            seen.add(k)
            unique.append(r)

    print_table(unique)

    csv_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "extracted_results.csv"
    )
    save_csv(unique, csv_path)


if __name__ == "__main__":
    main()
