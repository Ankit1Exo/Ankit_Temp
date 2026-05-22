"""
W-2 Employee Data Extractor (FINAL — tuned to ADP layout)
==========================================================
Extracts ONLY these 4 fields from each employee on each page:
  - employee_name        (from 'e/f Employee's name, address and ZIP code' line 1)
  - street_address       (line 2)
  - city_state_zip       (line 3)
  - ssn                  (from 'Employee's SSA number')

Layout assumptions (verified via diagnostic):
  - 'e/f' block sits at top-left of page (x < 200, top < 250)
  - Block label line contains the text 'e/f' followed by 'name'
  - Below the label, 3 lines: NAME / STREET / CITY,ST ZIP
  - SSN appears ~30-50pt below the block on the line with EIN
  - Multiple employees per page are supported (rare per user)
  - Bottom-row SSNs (top > 450) are employer copies — IGNORED

GUI:
  - Source folder picker (recursive PDF scan)
  - Destination folder picker (output goes here)
  - Live per-page progress bar in console window

USAGE:
    python extract_w2_gui.py

REQUIREMENTS:
    pip install pdfplumber pandas openpyxl tqdm

SECURITY:
    Output contains SSNs and PII. Restrict access. Delete after loading
    into your authorized HR system.
"""

from __future__ import annotations
import os
import re
import sys
import csv
import time
import threading
import multiprocessing as mp
from pathlib import Path
from multiprocessing import freeze_support, Process, Queue

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
DEFAULT_WORKERS = max(1, (os.cpu_count() or 4) - 1)

CSV_FIELDS = [
    "source_file", "page", "ssn",
    "employee_name", "street_address", "city_state_zip",
]

# SSN: matches XXX-XX-XXXX (but NOT EIN format XX-XXXXXXX which has dash at pos 3)
SSN_PATTERN = re.compile(r"^\d{3}-\d{2}-\d{4}$")
SSN_IN_LINE = re.compile(r"\b(\d{3}-\d{2}-\d{4})\b")
EIN_PATTERN = re.compile(r"^\d{2}-\d{7}$")

# City, ST ZIP pattern
CITY_STATE_ZIP_PATTERN = re.compile(
    r"^.+,?\s+[A-Z]{2}\s+\d{5}(?:-\d{4})?\s*$"
)

# Layout constants (from diagnostic on real PDFs)
# The 'e/f' block label sits at top < 250 and x0 < 50
EF_LABEL_MAX_Y = 280
EF_LABEL_MAX_X = 60
# SSN within ~50pt below the block label
SSN_OFFSET_FROM_LABEL_MIN = 30
SSN_OFFSET_FROM_LABEL_MAX = 80
# Ignore the bottom-row SSNs (employer copies)
BOTTOM_ROW_Y_MIN = 450

OUTPUT_CSV_NAME  = "w2_output.csv"
OUTPUT_XLSX_NAME = "w2_output.xlsx"
OUTPUT_ERR_NAME  = "w2_output.errors.log"


# ===========================================================================
# EXTRACTION
# ===========================================================================
def _find_ef_blocks(words):
    """
    Find all 'e/f' label positions on a page.
    Returns list of dicts with 'top' and 'x0' for each label found.
    """
    blocks = []
    # The label appears as a single fused word like 'e/f' or as separate tokens.
    # Look for any word containing 'e/f' near the start of a line at the top of page.
    for w in words:
        if w["top"] > EF_LABEL_MAX_Y:
            continue
        if w["x0"] > EF_LABEL_MAX_X:
            continue
        text_lower = w["text"].lower()
        # The diagnostic showed: 'e/f' appears as its own token at x=10
        if text_lower == "e/f" or text_lower.startswith("e/f"):
            blocks.append({"top": w["top"], "x0": w["x0"]})
    return blocks


def _extract_employee_from_block(words, label_top, label_x0):
    """
    Given the position of an 'e/f' label, extract the 3 lines below it
    (name / street / city-state-zip) and the SSN that appears further below.

    Returns dict with name/street/csz/ssn or None if extraction failed.
    """
    # Lines under the label: between label_top+2 and label_top+30 (roughly 3 lines tall)
    # We want words with similar x0 (within ~50pt) and inside the left column
    name_lines_words = [
        w for w in words
        if label_top + 2 < w["top"] < label_top + 35
        and w["x0"] < 200          # left column
        and w["x0"] > label_x0 - 5  # not too far left
    ]
    # Group into visual lines by 'top'
    name_lines_words.sort(key=lambda w: (w["top"], w["x0"]))
    lines = []
    LINE_TOL = 4
    for w in name_lines_words:
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

    if len(text_lines) < 2:
        return None  # not enough data

    # Find SSN in the area below the block
    # SSN line is typically ~40-60pt below the label
    ssn_candidates = []
    for w in words:
        if label_top + SSN_OFFSET_FROM_LABEL_MIN < w["top"] < label_top + SSN_OFFSET_FROM_LABEL_MAX:
            if w["x0"] < 200:  # left column
                m = SSN_IN_LINE.search(w["text"])
                if m:
                    # Make sure it's really an SSN (XXX-XX-XXXX) not EIN (XX-XXXXXXX)
                    candidate = m.group(1)
                    if SSN_PATTERN.match(candidate):
                        ssn_candidates.append(candidate)

    # Sometimes the SSN is concatenated with the EIN on the same line — search
    # the full text of that line region too
    if not ssn_candidates:
        region_words = [
            w for w in words
            if label_top + SSN_OFFSET_FROM_LABEL_MIN < w["top"] < label_top + SSN_OFFSET_FROM_LABEL_MAX
            and w["x0"] < 250
        ]
        line_text = " ".join(w["text"] for w in sorted(region_words, key=lambda w: w["x0"]))
        for m in SSN_IN_LINE.finditer(line_text):
            cand = m.group(1)
            if SSN_PATTERN.match(cand):
                ssn_candidates.append(cand)

    ssn = ssn_candidates[0] if ssn_candidates else ""

    # Parse the 3 lines: name / street / city,state,zip
    name = street = csz = ""

    # Identify city/state/zip line by pattern, work backward
    csz_idx = -1
    for i, ln in enumerate(text_lines):
        if CITY_STATE_ZIP_PATTERN.match(ln):
            csz_idx = i
            csz = ln
            break

    if csz_idx == -1:
        # Positional fallback
        if len(text_lines) >= 1: name   = text_lines[0]
        if len(text_lines) >= 2: street = text_lines[1]
        if len(text_lines) >= 3: csz    = text_lines[2]
    elif csz_idx == 1:
        name = text_lines[0]
    elif csz_idx >= 2:
        name = text_lines[0]
        street = " ".join(text_lines[1:csz_idx]).strip()

    return {
        "name":           name.strip(),
        "street":         street.strip(),
        "city_state_zip": csz.strip(),
        "ssn":            ssn,
    }


def count_pages(pdf_path_str):
    try:
        with pdfplumber.open(pdf_path_str) as pdf:
            return len(pdf.pages)
    except Exception:
        return 0


def worker_extract(pdf_path_str, out_queue):
    """
    Worker process: extract employee data from one PDF.
    Sends progress and records to out_queue.
    """
    pdf_path = Path(pdf_path_str)
    pages_done = 0
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                try:
                    words = page.extract_words(use_text_flow=False,
                                               keep_blank_chars=False)
                except Exception:
                    words = []

                if words:
                    # Find all e/f blocks on the page (usually 1, sometimes more)
                    blocks = _find_ef_blocks(words)

                    seen_ssns_on_page = set()
                    seen_keys_on_page = set()
                    for block in blocks:
                        result = _extract_employee_from_block(
                            words, block["top"], block["x0"])
                        if not result:
                            continue
                        # Dedup within page (3 side-by-side copies = same SSN)
                        if result["ssn"] and result["ssn"] in seen_ssns_on_page:
                            continue
                        # Also dedup by name+address in case SSN is missing
                        key = (result["name"], result["street"],
                               result["city_state_zip"])
                        if key == ("", "", ""):
                            continue
                        if not result["ssn"] and key in seen_keys_on_page:
                            continue
                        if result["ssn"]:
                            seen_ssns_on_page.add(result["ssn"])
                        seen_keys_on_page.add(key)

                        record = {
                            "source_file": pdf_path.name,
                            "page": page_num,
                            "ssn": result["ssn"],
                            "employee_name": result["name"],
                            "street_address": result["street"],
                            "city_state_zip": result["city_state_zip"],
                        }
                        out_queue.put(("record", record))

                pages_done += 1
                out_queue.put(("page", pdf_path.name, page_num))

        out_queue.put(("done", pdf_path.name, pages_done, ""))
    except Exception as e:
        out_queue.put(("done", pdf_path.name, pages_done,
                       "{}: {}".format(type(e).__name__, e)))


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
# DRIVER — manages workers and progress
# ===========================================================================
def run_extraction(source_folder, dest_folder, max_workers,
                   resume, generate_xlsx, status_callback):
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
    print("W-2 Extractor (ADP layout, live per-page progress)")
    print("Source:      {}".format(src))
    print("Destination: {}".format(dst))
    print("Max workers: {}".format(max_workers))
    print("=" * 70)

    status_callback("Scanning for PDFs...")
    print("\nScanning {} for PDFs (recursive)...".format(src))
    pdfs = sorted(src.rglob("*.pdf"))
    if not pdfs:
        print("No PDFs found.")
        status_callback("No PDFs found in source folder.")
        return False
    print("Found {} PDF file(s).".format(len(pdfs)))

    # Resume
    already_done = set()
    if resume:
        already_done = load_already_processed(output_csv)
        if already_done:
            print("Resume: skipping {} file(s) already in output.".format(len(already_done)))
            pdfs = [p for p in pdfs if p.name not in already_done]
            if not pdfs:
                print("Nothing new to process.")
                status_callback("Nothing new to process.")
                return True

    # Pre-count pages
    status_callback("Counting pages in {} file(s)...".format(len(pdfs)))
    print("\nCounting pages...")
    total_pages = 0
    with tqdm(total=len(pdfs), unit="file", desc="Counting", ncols=100) as cbar:
        for p in pdfs:
            total_pages += count_pages(str(p))
            cbar.update(1)
    print("Total pages to process: {}".format(total_pages))
    if total_pages == 0:
        status_callback("No readable pages.")
        return False

    # Open CSV
    csv_has_data = output_csv.exists() and bool(already_done)
    csv_file = output_csv.open("a" if csv_has_data else "w",
                               encoding="utf-8", newline="")
    writer = csv.DictWriter(csv_file, fieldnames=CSV_FIELDS)
    if not csv_has_data:
        writer.writeheader()
        csv_file.flush()

    workers = min(max_workers, len(pdfs))
    status_callback("Extracting {} pages across {} file(s) with {} worker(s)...".format(
        total_pages, len(pdfs), workers))
    print("\nProcessing with {} worker(s)...\n".format(workers))

    out_queue: Queue = mp.Queue()
    pdfs_iter = iter(pdfs)
    active = {}
    error_log = []
    total_records = 0
    files_done = 0
    start = time.time()
    last_flush = time.time()

    def launch_next():
        try:
            p = next(pdfs_iter)
        except StopIteration:
            return False
        proc = Process(target=worker_extract, args=(str(p), out_queue), daemon=True)
        proc.start()
        active[proc.pid] = (proc, p)
        return True

    for _ in range(workers):
        if not launch_next():
            break

    pbar = tqdm(total=total_pages, unit="page", desc="Extracting", ncols=100,
                smoothing=0.1)

    try:
        while active:
            try:
                msg = out_queue.get(timeout=1.0)
            except Exception:
                pbar.refresh()
                continue

            kind = msg[0]
            if kind == "page":
                _, fname, page_num = msg
                pbar.update(1)
                pbar.set_postfix(file=fname[:24], page=page_num,
                                 records=total_records,
                                 files="{}/{}".format(files_done, len(pdfs)),
                                 errors=len(error_log))
            elif kind == "record":
                _, record = msg
                writer.writerow(record)
                total_records += 1
                if time.time() - last_flush > 2.0:
                    csv_file.flush()
                    last_flush = time.time()
            elif kind == "done":
                _, fname, _, err = msg
                if err: error_log.append("{}: {}".format(fname, err))
                files_done += 1
                finished_pid = None
                for pid, (proc, p) in list(active.items()):
                    if p.name == fname:
                        proc.join(timeout=2)
                        finished_pid = pid
                        break
                if finished_pid is not None:
                    del active[finished_pid]
                launch_next()
                status_callback("Files: {}/{} | Records: {} | Errors: {}".format(
                    files_done, len(pdfs), total_records, len(error_log)))
                csv_file.flush()
                last_flush = time.time()
    finally:
        pbar.close()
        for pid, (proc, _) in list(active.items()):
            if proc.is_alive():
                proc.terminate()
            proc.join(timeout=2)
        csv_file.flush()
        csv_file.close()

    elapsed = time.time() - start
    print("\nDone in {:.1f} min ({:.0f}s)".format(elapsed/60, elapsed))
    print("Records extracted: {}".format(total_records))
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

    status_callback("Done. {} record(s) in {:.1f} min. Output: {}".format(
        total_records, elapsed/60, dst))
    return True


# ===========================================================================
# TKINTER GUI
# ===========================================================================
class ExtractorGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("W-2 Extractor")
        self.geometry("680x360")
        self.resizable(False, False)
        self._running = False
        self._build_widgets()

    def _build_widgets(self):
        pad = {"padx": 12, "pady": 8}

        ttk.Label(self, text="W-2 Employee Data Extractor",
                  font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(self,
            text="Live progress shows in the console window (per-page bar).",
            foreground="#555").grid(row=1, column=0, columnspan=3,
                                    sticky="w", padx=12)

        ttk.Label(self, text="Source folder:").grid(row=2, column=0, sticky="e", **pad)
        self.src_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.src_var, width=60).grid(
            row=2, column=1, sticky="we", **pad)
        ttk.Button(self, text="Browse...", command=self._pick_source).grid(
            row=2, column=2, **pad)

        ttk.Label(self, text="Destination folder:").grid(row=3, column=0, sticky="e", **pad)
        self.dst_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.dst_var, width=60).grid(
            row=3, column=1, sticky="we", **pad)
        ttk.Button(self, text="Browse...", command=self._pick_destination).grid(
            row=3, column=2, **pad)

        opts = ttk.LabelFrame(self, text="Options")
        opts.grid(row=4, column=0, columnspan=3, sticky="we", padx=12, pady=8)

        ttk.Label(opts, text="Workers:").grid(row=0, column=0, sticky="e", padx=8, pady=6)
        self.workers_var = tk.IntVar(value=DEFAULT_WORKERS)
        ttk.Spinbox(opts, from_=1, to=max(1, (os.cpu_count() or 4)),
                    textvariable=self.workers_var, width=5).grid(
            row=0, column=1, sticky="w", padx=4, pady=6)

        self.resume_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Resume (skip files already in CSV)",
                        variable=self.resume_var).grid(
            row=0, column=2, sticky="w", padx=20, pady=6)

        self.xlsx_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Also generate Excel (.xlsx)",
                        variable=self.xlsx_var).grid(
            row=1, column=2, sticky="w", padx=20, pady=6)

        self.start_btn = ttk.Button(self, text="Start Extraction",
                                    command=self._start_clicked)
        self.start_btn.grid(row=5, column=0, columnspan=3, pady=10)

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(self, textvariable=self.status_var, relief="sunken",
                  anchor="w").grid(row=6, column=0, columnspan=3,
                                   sticky="we", padx=12, pady=(0, 12))
        self.columnconfigure(1, weight=1)

    def _pick_source(self):
        folder = filedialog.askdirectory(title="Select source folder (W-2 PDFs)",
                                          mustexist=True)
        if folder:
            self.src_var.set(folder)
            if not self.dst_var.get():
                self.dst_var.set(str(Path(folder).parent / "w2_extracted"))

    def _pick_destination(self):
        folder = filedialog.askdirectory(title="Select destination folder",
                                          mustexist=False)
        if folder:
            self.dst_var.set(folder)

    def _set_status(self, msg):
        self.after(0, lambda: self.status_var.set(msg))

    def _start_clicked(self):
        if self._running:
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
            messagebox.showerror("Invalid source", "Source folder does not exist.")
            return

        self._running = True
        self.start_btn.config(state="disabled", text="Running...")
        self._set_status("Starting...")

        t = threading.Thread(
            target=self._run_in_thread,
            args=(src, dst, self.workers_var.get(),
                  self.resume_var.get(), self.xlsx_var.get()),
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
        self._running = False
        self.start_btn.config(state="normal", text="Start Extraction")


def main():
    app = ExtractorGUI()
    app.mainloop()


if __name__ == "__main__":
    freeze_support()
    main()
