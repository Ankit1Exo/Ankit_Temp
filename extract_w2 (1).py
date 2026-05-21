"""
W-2 Employee Data Extractor (Fast / Bulk)
==========================================
Extracts Employee Name, Street Address, City/State/ZIP, and SSN from W-2 PDFs.
Optimized for large batches (15K+ pages):
  - GUI folder picker
  - Parallel processing across CPU cores
  - Live progress bar
  - Streaming CSV writes (crash-safe, low memory)
  - Auto-resume: skips files already in the output CSV

USAGE:
    python extract_w2.py
    (a folder picker dialog will open)

REQUIREMENTS:
    pip install pdfplumber pandas openpyxl tqdm

SECURITY NOTE:
    W-2s contain SSNs and PII. Run ONLY on an authorized workstation.
    Restrict file permissions on the output (chmod 600 on Linux/Mac).
    Delete outputs once data is loaded into your authorized HR system.
"""

from __future__ import annotations
import os
import re
import sys
import csv
import time
import argparse
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed
from multiprocessing import freeze_support

try:
    import pdfplumber
except ImportError:
    sys.exit("Missing dependency. Run: pip install pdfplumber pandas openpyxl tqdm")

try:
    from tqdm import tqdm
except ImportError:
    sys.exit("Missing dependency. Run: pip install tqdm")

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
DEFAULT_WORKERS = max(1, (os.cpu_count() or 4) - 1)  # leave 1 core for OS
CSV_FIELDS = [
    "source_file", "page", "ssn",
    "employee_name", "street_address", "city_state_zip",
]

SSN_PATTERN = re.compile(r"\b(\d{3}-\d{2}-\d{4})\b")
CITY_STATE_ZIP_PATTERN = re.compile(
    r"^.+,?\s+[A-Z]{2}\s+\d{5}(?:-\d{4})?\s*$"
)

# ---------------------------------------------------------------------------
# FOLDER PICKER
# ---------------------------------------------------------------------------
def pick_folder_gui(title: str = "Select folder containing W-2 PDFs"):
    """Open a native folder picker dialog. Returns folder path or None."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        return None
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title=title, mustexist=True)
    root.destroy()
    return folder or None


def pick_output_file_gui(default_name: str = "w2_output.csv"):
    """Open a Save As dialog for the output CSV."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        return None
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.asksaveasfilename(
        title="Save extracted data as...",
        defaultextension=".csv",
        initialfile=default_name,
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
    )
    root.destroy()
    return path or None


# ---------------------------------------------------------------------------
# EXTRACTION (per-page; runs in worker processes)
# ---------------------------------------------------------------------------
def _extract_employee_block(page):
    """Word-coordinate-based extraction of the Employee box (left half)."""
    try:
        words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
    except Exception:
        return []
    if not words:
        return []

    page_width = page.width
    left_max_x = page_width * 0.55

    anchor_top = None
    for w in words:
        if w["text"].strip().lower() == "employee's" and w["x0"] < left_max_x:
            anchor_top = w["top"]
            break
    if anchor_top is None:
        return []

    end_top = None
    end_labels = {"15", "State", "Employer's"}
    for w in sorted(words, key=lambda w: w["top"]):
        if w["top"] <= anchor_top + 5:
            continue
        if w["text"].strip() in end_labels and w["x0"] < left_max_x:
            end_top = w["top"]
            break
    if end_top is None:
        end_top = anchor_top + 200

    block_words = []
    for w in words:
        if w["top"] <= anchor_top + 8:
            continue
        if w["top"] >= end_top - 2:
            continue
        if w["x0"] >= left_max_x:
            continue
        if w["text"].strip().lower() == "suff.":
            continue
        block_words.append(w)
    if not block_words:
        return []

    block_words.sort(key=lambda w: (w["top"], w["x0"]))
    lines = []
    LINE_TOL = 4
    for w in block_words:
        if not lines:
            lines.append([w]); continue
        if abs(w["top"] - lines[-1][0]["top"]) <= LINE_TOL:
            lines[-1].append(w)
        else:
            lines.append([w])

    text_lines = []
    for line in lines:
        line.sort(key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in line).strip()
        if text and len(text) > 1:
            text_lines.append(text)
    return text_lines


def _parse_employee_lines(lines):
    name = street = csz = ""
    if not lines:
        return {"name": "", "street": "", "city_state_zip": ""}

    csz_idx = -1
    for i, ln in enumerate(lines):
        if CITY_STATE_ZIP_PATTERN.match(ln):
            csz_idx = i
            csz = ln
            break

    if csz_idx == -1:
        if len(lines) >= 1: name   = lines[0]
        if len(lines) >= 2: street = lines[1]
        if len(lines) >= 3: csz    = lines[2]
    elif csz_idx == 0:
        pass
    elif csz_idx == 1:
        name = lines[0]
    else:
        name = lines[0]
        street = " ".join(lines[1:csz_idx]).strip()

    return {"name": name.strip(), "street": street.strip(), "city_state_zip": csz.strip()}


def process_pdf(pdf_path_str):
    """
    Worker function. Returns (filename, records, error_message).
    Runs in a separate process.
    """
    pdf_path = Path(pdf_path_str)
    records = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            seen_keys = set()
            for page_num, page in enumerate(pdf.pages, start=1):
                # Cheap text scan first — skip pages with no SSN pattern
                try:
                    full_text = page.extract_text() or ""
                except Exception:
                    full_text = ""
                ssns = SSN_PATTERN.findall(full_text)
                if not ssns:
                    # Not a W-2 page (or scanned). Skip the expensive word call.
                    continue

                emp_lines = _extract_employee_block(page)
                parsed = _parse_employee_lines(emp_lines)
                key = (parsed["name"], parsed["street"], parsed["city_state_zip"])
                if key == ("", "", ""):
                    continue
                if key in seen_keys:
                    continue
                seen_keys.add(key)

                records.append({
                    "source_file": pdf_path.name,
                    "page": page_num,
                    "ssn": ssns[0],
                    "employee_name": parsed["name"],
                    "street_address": parsed["street"],
                    "city_state_zip": parsed["city_state_zip"],
                })
        return (pdf_path.name, records, "")
    except Exception as e:
        return (pdf_path.name, [], "{}: {}".format(type(e).__name__, e))


# ---------------------------------------------------------------------------
# RESUME SUPPORT
# ---------------------------------------------------------------------------
def load_already_processed(csv_path):
    """Return set of source_file values already in the CSV."""
    if not csv_path.exists():
        return set()
    done = set()
    try:
        with csv_path.open("r", encoding="utf-8", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if row.get("source_file"):
                    done.add(row["source_file"])
    except Exception:
        pass
    return done


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Extract employee data from W-2 PDFs.")
    parser.add_argument("--input",   "-i", help="Folder of PDFs (skips GUI picker)")
    parser.add_argument("--output",  "-o", help="Output CSV path (skips GUI picker)")
    parser.add_argument("--workers", "-w", type=int, default=DEFAULT_WORKERS,
                        help="Parallel workers (default: {})".format(DEFAULT_WORKERS))
    parser.add_argument("--no-resume", action="store_true",
                        help="Don't skip files already in output CSV")
    parser.add_argument("--no-xlsx", action="store_true",
                        help="Don't generate Excel file at the end")
    args = parser.parse_args()

    # Resolve input folder
    if args.input:
        input_folder = Path(args.input)
    else:
        print("Opening folder picker...")
        folder = pick_folder_gui()
        if not folder:
            print("No folder selected. Exiting.")
            return 1
        input_folder = Path(folder)

    if not input_folder.exists() or not input_folder.is_dir():
        print("Not a valid folder: {}".format(input_folder))
        return 1

    # Resolve output file
    if args.output:
        output_csv = Path(args.output)
    else:
        print("Opening save dialog...")
        out = pick_output_file_gui()
        if not out:
            output_csv = input_folder / "w2_output.csv"
            print("No output selected — defaulting to {}".format(output_csv))
        else:
            output_csv = Path(out)

    # Scan for PDFs (recursively, so subfolders are included)
    print("\nScanning {} for PDFs (recursive)...".format(input_folder))
    pdfs = sorted(input_folder.rglob("*.pdf"))
    if not pdfs:
        print("No PDFs found.")
        return 1
    print("Found {} PDF file(s).".format(len(pdfs)))

    # Resume: skip files already done
    already_done = set()
    if not args.no_resume:
        already_done = load_already_processed(output_csv)
        if already_done:
            print("Resume mode: {} file(s) already in output — will skip.".format(len(already_done)))
            pdfs = [p for p in pdfs if p.name not in already_done]
            if not pdfs:
                print("Nothing new to process.")
                return 0

    # Open CSV in append mode (or write mode with header if new)
    csv_has_data = output_csv.exists() and bool(already_done)
    output_csv.parent.mkdir(parents=True, exist_ok=True)
    mode = "a" if csv_has_data else "w"
    csv_file = output_csv.open(mode, encoding="utf-8", newline="")
    writer = csv.DictWriter(csv_file, fieldnames=CSV_FIELDS)
    if not csv_has_data:
        writer.writeheader()
        csv_file.flush()

    # Process in parallel
    workers = max(1, args.workers)
    print("\nProcessing {} PDF(s) with {} worker(s)...\n".format(len(pdfs), workers))

    start = time.time()
    total_records = 0
    error_log = []
    flush_every = 50  # flush to disk every N completed files

    with ProcessPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(process_pdf, str(p)): p for p in pdfs}
        with tqdm(total=len(futures), unit="file", desc="Extracting") as pbar:
            done_count = 0
            for fut in as_completed(futures):
                fname, records, err = fut.result()
                if err:
                    error_log.append("{}: {}".format(fname, err))
                for r in records:
                    writer.writerow(r)
                total_records += len(records)
                done_count += 1
                if done_count % flush_every == 0:
                    csv_file.flush()
                pbar.update(1)
                pbar.set_postfix(records=total_records, errors=len(error_log))

    csv_file.flush()
    csv_file.close()
    elapsed = time.time() - start

    print("\nDone in {:.1f} min ({:.0f}s)".format(elapsed / 60, elapsed))
    print("Records extracted this run: {}".format(total_records))
    print("CSV: {}".format(output_csv.resolve()))

    # Write error log alongside CSV
    if error_log:
        err_path = output_csv.with_suffix(".errors.log")
        with err_path.open("w", encoding="utf-8") as f:
            f.write("\n".join(error_log))
        print("Errors: {} (see {})".format(len(error_log), err_path.name))

    # Optional XLSX
    if not args.no_xlsx:
        try:
            import pandas as pd
            xlsx_path = output_csv.with_suffix(".xlsx")
            print("\nWriting Excel file: {} ...".format(xlsx_path.name))
            df = pd.read_csv(output_csv, dtype=str)
            df.to_excel(xlsx_path, index=False)
            print("Excel: {}".format(xlsx_path.resolve()))
        except Exception as e:
            print("(Excel export skipped: {})".format(e))

    return 0


if __name__ == "__main__":
    freeze_support()  # Windows-safe
    sys.exit(main())
