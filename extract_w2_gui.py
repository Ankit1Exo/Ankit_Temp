"""
W-2 Employee Data Extractor — Tkinter GUI
==========================================
Extracts Employee Name, Street Address, City/State/ZIP, and SSN from W-2 PDFs.

GUI:
    - Source folder picker  (folder containing W-2 PDFs, recursive)
    - Destination folder picker (where CSV / XLSX / error log go)
    - Start Extraction button
    - Status label

Progress is printed to the prompt/console window (tqdm-style):
    Extracting:  42%|████▏     | 6300/15000 [12:34<17:21, 8.34page/s, records=6280, errors=2]

USAGE:
    python extract_w2_gui.py

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
import threading
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed
from multiprocessing import freeze_support

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

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

OUTPUT_CSV_NAME  = "w2_output.csv"
OUTPUT_XLSX_NAME = "w2_output.xlsx"
OUTPUT_ERR_NAME  = "w2_output.errors.log"


# ===========================================================================
# EXTRACTION LOGIC (unchanged — kept module-level so workers can import)
# ===========================================================================
def _extract_employee_block(page):
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
        if w["top"] <= anchor_top + 8: continue
        if w["top"] >= end_top - 2: continue
        if w["x0"] >= left_max_x: continue
        if w["text"].strip().lower() == "suff.": continue
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
            csz_idx = i; csz = ln; break

    if csz_idx == -1:
        if len(lines) >= 1: name   = lines[0]
        if len(lines) >= 2: street = lines[1]
        if len(lines) >= 3: csz    = lines[2]
    elif csz_idx == 1:
        name = lines[0]
    elif csz_idx > 1:
        name = lines[0]
        street = " ".join(lines[1:csz_idx]).strip()

    return {"name": name.strip(), "street": street.strip(), "city_state_zip": csz.strip()}


def count_pages(pdf_path_str):
    try:
        with pdfplumber.open(pdf_path_str) as pdf:
            return (pdf_path_str, len(pdf.pages), "")
    except Exception as e:
        return (pdf_path_str, 0, "{}: {}".format(type(e).__name__, e))


def process_pdf_batch(pdf_path_str):
    """Worker: process one PDF, return (filename, pages_done, records, error)."""
    pdf_path = Path(pdf_path_str)
    records = []
    pages_done = 0
    try:
        with pdfplumber.open(pdf_path) as pdf:
            seen_keys = set()
            for page_num, page in enumerate(pdf.pages, start=1):
                pages_done += 1
                try:
                    full_text = page.extract_text() or ""
                except Exception:
                    full_text = ""
                ssns = SSN_PATTERN.findall(full_text)
                if not ssns:
                    continue
                emp_lines = _extract_employee_block(page)
                parsed = _parse_employee_lines(emp_lines)
                key = (parsed["name"], parsed["street"], parsed["city_state_zip"])
                if key == ("", "", "") or key in seen_keys:
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
        return (pdf_path.name, pages_done, records, "")
    except Exception as e:
        return (pdf_path.name, pages_done, [], "{}: {}".format(type(e).__name__, e))


def load_already_processed(csv_path):
    if not csv_path.exists():
        return set()
    done = set()
    try:
        with csv_path.open("r", encoding="utf-8", newline="") as f:
            for row in csv.DictReader(f):
                if row.get("source_file"):
                    done.add(row["source_file"])
    except Exception:
        pass
    return done


# ===========================================================================
# EXTRACTION RUNNER (called from GUI thread)
# ===========================================================================
def run_extraction(source_folder, dest_folder, workers,
                   resume, generate_xlsx, status_callback):
    """
    Runs the extraction. Prints tqdm progress to console.
    status_callback(message) — called to update the GUI status label.
    """
    src = Path(source_folder)
    dst = Path(dest_folder)
    if not src.is_dir():
        status_callback("ERROR: Source folder invalid.")
        return False
    dst.mkdir(parents=True, exist_ok=True)

    output_csv  = dst / OUTPUT_CSV_NAME
    output_xlsx = dst / OUTPUT_XLSX_NAME
    output_err  = dst / OUTPUT_ERR_NAME

    print("=" * 70)
    print("W-2 Extractor")
    print("Source:      {}".format(src))
    print("Destination: {}".format(dst))
    print("Workers:     {}".format(workers))
    print("=" * 70)

    status_callback("Scanning for PDFs...")
    print("\nScanning {} for PDFs (recursive)...".format(src))
    pdfs = sorted(src.rglob("*.pdf"))
    if not pdfs:
        print("No PDFs found.")
        status_callback("No PDFs found in source folder.")
        return False
    print("Found {} PDF file(s).".format(len(pdfs)))
    status_callback("Found {} PDF(s). Pre-counting pages...".format(len(pdfs)))

    # Resume
    already_done = set()
    if resume:
        already_done = load_already_processed(output_csv)
        if already_done:
            print("Resume mode: {} file(s) already in output — will skip.".format(len(already_done)))
            pdfs = [p for p in pdfs if p.name not in already_done]
            if not pdfs:
                print("Nothing new to process.")
                status_callback("Nothing new to process — all files already done.")
                return True

    # Pre-count pages for accurate progress bar
    print("\nCounting pages...")
    total_pages = 0
    with tqdm(total=len(pdfs), unit="file", desc="Counting", ncols=100) as cbar:
        for p in pdfs:
            _, npages, _ = count_pages(str(p))
            total_pages += npages
            cbar.update(1)
    print("Total pages to process: {}".format(total_pages))
    status_callback("Extracting from {} pages across {} file(s)...".format(total_pages, len(pdfs)))

    # Open CSV
    csv_has_data = output_csv.exists() and bool(already_done)
    csv_file = output_csv.open("a" if csv_has_data else "w", encoding="utf-8", newline="")
    writer = csv.DictWriter(csv_file, fieldnames=CSV_FIELDS)
    if not csv_has_data:
        writer.writeheader()
        csv_file.flush()

    start = time.time()
    total_records = 0
    error_log = []
    files_done = 0

    print("\nProcessing {} PDF(s) with {} worker(s)...\n".format(len(pdfs), workers))
    pbar = tqdm(total=total_pages, unit="page", desc="Extracting", ncols=100)

    try:
        if workers == 1:
            # Single-process path
            for p in pdfs:
                fname, pages_done, records, err = process_pdf_batch(str(p))
                if err: error_log.append("{}: {}".format(fname, err))
                for r in records:
                    writer.writerow(r); total_records += 1
                files_done += 1
                pbar.update(pages_done)
                pbar.set_postfix(records=total_records, errors=len(error_log),
                                 files="{}/{}".format(files_done, len(pdfs)))
                status_callback("Processed {}/{} file(s) | {} record(s)".format(
                    files_done, len(pdfs), total_records))
                if files_done % 25 == 0:
                    csv_file.flush()
        else:
            with ProcessPoolExecutor(max_workers=workers) as executor:
                futures = {executor.submit(process_pdf_batch, str(p)): p for p in pdfs}
                for fut in as_completed(futures):
                    fname, pages_done, records, err = fut.result()
                    if err: error_log.append("{}: {}".format(fname, err))
                    for r in records:
                        writer.writerow(r); total_records += 1
                    files_done += 1
                    pbar.update(pages_done)
                    pbar.set_postfix(records=total_records, errors=len(error_log),
                                     files="{}/{}".format(files_done, len(pdfs)))
                    status_callback("Processed {}/{} file(s) | {} record(s)".format(
                        files_done, len(pdfs), total_records))
                    if files_done % 25 == 0:
                        csv_file.flush()
    finally:
        pbar.close()
        csv_file.flush()
        csv_file.close()

    elapsed = time.time() - start
    print("\nDone in {:.1f} min ({:.0f}s)".format(elapsed/60, elapsed))
    print("Records extracted this run: {}".format(total_records))
    print("CSV: {}".format(output_csv.resolve()))

    if error_log:
        with output_err.open("w", encoding="utf-8") as f:
            f.write("\n".join(error_log))
        print("Errors: {} (see {})".format(len(error_log), output_err.name))

    if generate_xlsx:
        try:
            import pandas as pd
            print("\nWriting Excel: {} ...".format(output_xlsx.name))
            df = pd.read_csv(output_csv, dtype=str)
            df.to_excel(output_xlsx, index=False)
            print("Excel: {}".format(output_xlsx.resolve()))
        except Exception as e:
            print("(Excel export skipped: {})".format(e))

    status_callback("Done. {} record(s) extracted in {:.1f} min. Output in: {}".format(
        total_records, elapsed/60, dst))
    return True


# ===========================================================================
# TKINTER GUI
# ===========================================================================
class ExtractorGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("W-2 Extractor")
        self.geometry("640x340")
        self.resizable(False, False)

        self._building = False  # prevents double-clicks while running
        self._build_widgets()

    def _build_widgets(self):
        pad = {"padx": 12, "pady": 8}

        # Title
        title = ttk.Label(self, text="W-2 Employee Data Extractor",
                          font=("Segoe UI", 14, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w", **pad)

        subtitle = ttk.Label(
            self,
            text="Progress will be displayed in the console window.",
            foreground="#555")
        subtitle.grid(row=1, column=0, columnspan=3, sticky="w", padx=12)

        # Source folder
        ttk.Label(self, text="Source folder:").grid(row=2, column=0, sticky="e", **pad)
        self.src_var = tk.StringVar()
        self.src_entry = ttk.Entry(self, textvariable=self.src_var, width=55)
        self.src_entry.grid(row=2, column=1, sticky="we", **pad)
        ttk.Button(self, text="Browse...", command=self._pick_source).grid(
            row=2, column=2, **pad)

        # Destination folder
        ttk.Label(self, text="Destination folder:").grid(row=3, column=0, sticky="e", **pad)
        self.dst_var = tk.StringVar()
        self.dst_entry = ttk.Entry(self, textvariable=self.dst_var, width=55)
        self.dst_entry.grid(row=3, column=1, sticky="we", **pad)
        ttk.Button(self, text="Browse...", command=self._pick_destination).grid(
            row=3, column=2, **pad)

        # Options
        opts = ttk.LabelFrame(self, text="Options")
        opts.grid(row=4, column=0, columnspan=3, sticky="we", padx=12, pady=8)

        ttk.Label(opts, text="Workers:").grid(row=0, column=0, sticky="e", padx=8, pady=6)
        self.workers_var = tk.IntVar(value=DEFAULT_WORKERS)
        ttk.Spinbox(opts, from_=1, to=max(1, (os.cpu_count() or 4)),
                    textvariable=self.workers_var, width=5).grid(
            row=0, column=1, sticky="w", padx=4, pady=6)

        self.resume_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Resume (skip files already in output CSV)",
                        variable=self.resume_var).grid(
            row=0, column=2, sticky="w", padx=20, pady=6)

        self.xlsx_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Also generate Excel (.xlsx)",
                        variable=self.xlsx_var).grid(
            row=1, column=2, sticky="w", padx=20, pady=6)

        # Start button
        self.start_btn = ttk.Button(self, text="Start Extraction",
                                    command=self._start_clicked)
        self.start_btn.grid(row=5, column=0, columnspan=3, pady=10)

        # Status label
        self.status_var = tk.StringVar(value="Ready.")
        status_lbl = ttk.Label(self, textvariable=self.status_var,
                               relief="sunken", anchor="w")
        status_lbl.grid(row=6, column=0, columnspan=3, sticky="we",
                        padx=12, pady=(0, 12))

        self.columnconfigure(1, weight=1)

    def _pick_source(self):
        folder = filedialog.askdirectory(title="Select source folder (contains W-2 PDFs)",
                                          mustexist=True)
        if folder:
            self.src_var.set(folder)
            # Auto-suggest destination if empty
            if not self.dst_var.get():
                self.dst_var.set(str(Path(folder).parent / "w2_extracted"))

    def _pick_destination(self):
        folder = filedialog.askdirectory(title="Select destination folder (for output)",
                                          mustexist=False)
        if folder:
            self.dst_var.set(folder)

    def _set_status(self, msg):
        # Thread-safe status update
        self.after(0, lambda: self.status_var.set(msg))

    def _start_clicked(self):
        if self._building:
            return

        src = self.src_var.get().strip()
        dst = self.dst_var.get().strip()

        if not src:
            messagebox.showerror("Missing source", "Please select a source folder.")
            return
        if not dst:
            messagebox.showerror("Missing destination", "Please select a destination folder.")
            return
        if not Path(src).is_dir():
            messagebox.showerror("Invalid source", "Source folder does not exist:\n{}".format(src))
            return

        # Same-folder warning (not blocking)
        if Path(src).resolve() == Path(dst).resolve():
            if not messagebox.askyesno(
                "Same folder",
                "Source and destination are the same folder. Continue?"):
                return

        workers = self.workers_var.get()
        resume = self.resume_var.get()
        xlsx = self.xlsx_var.get()

        self._building = True
        self.start_btn.config(state="disabled", text="Running...")
        self._set_status("Starting...")

        # Run in background thread so GUI stays responsive
        t = threading.Thread(
            target=self._run_in_thread,
            args=(src, dst, workers, resume, xlsx),
            daemon=True,
        )
        t.start()

    def _run_in_thread(self, src, dst, workers, resume, xlsx):
        try:
            run_extraction(src, dst, workers, resume, xlsx, self._set_status)
        except Exception as e:
            print("\nFATAL ERROR: {}".format(e))
            self._set_status("Error: {}".format(e))
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.after(0, self._finished)

    def _finished(self):
        self._building = False
        self.start_btn.config(state="normal", text="Start Extraction")


# ===========================================================================
# ENTRY POINT
# ===========================================================================
def main():
    app = ExtractorGUI()
    app.mainloop()


if __name__ == "__main__":
    freeze_support()  # Windows-safe for ProcessPoolExecutor
    main()
