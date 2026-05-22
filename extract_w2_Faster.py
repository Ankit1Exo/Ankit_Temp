"""
W-2 Employee Data Extractor (FINAL — pypdfium2 edition)
========================================================
Uses pypdfium2 (Chrome's PDFium engine) for fast PDF text extraction.
Significantly faster and lower memory than pdfplumber.

Extracts ONLY:
  - employee_name        (from 'e/f Employee's name, address and ZIP code' line 1)
  - street_address       (line 2)
  - city_state_zip       (line 3)
  - ssn                  (Employee's SSA number)

GUI:
  - Source folder picker
  - Destination folder picker
  - Live per-page progress bar in console window

USAGE:
    python extract_w2_gui.py

REQUIREMENTS:
    pip install pypdfium2 pandas openpyxl tqdm

(pdfplumber is no longer needed.)

SECURITY:
    Output contains SSNs and PII. Restrict access. Delete after
    loading into your authorized HR system.
"""

from __future__ import annotations
import os
import re
import sys
import csv
import gc
import time
import threading
import multiprocessing as mp
from pathlib import Path
from multiprocessing import freeze_support, Process, Queue

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import pypdfium2 as pdfium
except ImportError:
    sys.exit("Missing dependency. Run: pip install pypdfium2 pandas openpyxl tqdm")

try:
    from tqdm import tqdm
except ImportError:
    sys.exit("Missing dependency. Run: pip install tqdm")

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
DEFAULT_WORKERS = min(4, max(1, (os.cpu_count() or 4) - 1))
CHUNK_SIZE = 100            # pages per worker chunk
PROGRESS_BATCH = 10
QUEUE_MAX_SIZE = 1000
GC_EVERY_N_PAGES = 50

CSV_FIELDS = [
    "source_file", "page", "ssn",
    "employee_name", "street_address", "city_state_zip",
]

SSN_PATTERN = re.compile(r"^\d{3}-\d{2}-\d{4}$")
SSN_IN_LINE = re.compile(r"\b(\d{3}-\d{2}-\d{4})\b")
CITY_STATE_ZIP_PATTERN = re.compile(
    r"^.+,?\s+[A-Z]{2}\s+\d{5}(?:-\d{4})?\s*$"
)

# Layout constants
EF_LABEL_MAX_Y = 280
EF_LABEL_MAX_X = 60
SSN_OFFSET_MIN = 30
SSN_OFFSET_MAX = 80

OUTPUT_CSV_NAME  = "w2_output.csv"
OUTPUT_XLSX_NAME = "w2_output.xlsx"
OUTPUT_ERR_NAME  = "w2_output.errors.log"


# ===========================================================================
# WORD EXTRACTION (pypdfium2)
# ===========================================================================
def extract_words_from_page(page):
    """
    Extract words from a pypdfium2 PdfPage. Returns list of dicts with
    {text, x0, x1, top, bottom} in pdfplumber's top-down coordinate system.
    """
    page_height = page.get_height()
    textpage = page.get_textpage()
    n_chars = textpage.count_chars()

    words = []
    cur_chars = []
    cur_left = cur_right = cur_top_pdf = cur_bottom_pdf = None

    for i in range(n_chars):
        ch = textpage.get_text_range(i, 1)
        if not ch:
            continue
        # End of word on whitespace
        if ch.isspace() or ord(ch[0]) < 32:
            if cur_chars and cur_left is not None:
                text = "".join(cur_chars)
                if text.strip():
                    words.append({
                        "text": text,
                        "x0": cur_left,
                        "x1": cur_right,
                        "top": page_height - cur_top_pdf,
                        "bottom": page_height - cur_bottom_pdf,
                    })
            cur_chars = []
            cur_left = cur_right = cur_top_pdf = cur_bottom_pdf = None
            continue

        try:
            left, bottom, right, top = textpage.get_charbox(i, loose=False)
        except Exception:
            continue

        # New word if vertical jump is large
        if cur_top_pdf is not None and abs(top - cur_top_pdf) > 6:
            text = "".join(cur_chars)
            if text.strip():
                words.append({
                    "text": text,
                    "x0": cur_left,
                    "x1": cur_right,
                    "top": page_height - cur_top_pdf,
                    "bottom": page_height - cur_bottom_pdf,
                })
            cur_chars = []
            cur_left = cur_right = cur_top_pdf = cur_bottom_pdf = None

        cur_chars.append(ch)
        if cur_left is None or left < cur_left: cur_left = left
        if cur_right is None or right > cur_right: cur_right = right
        if cur_top_pdf is None or top > cur_top_pdf: cur_top_pdf = top
        if cur_bottom_pdf is None or bottom < cur_bottom_pdf: cur_bottom_pdf = bottom

    # Final flush
    if cur_chars and cur_left is not None:
        text = "".join(cur_chars)
        if text.strip():
            words.append({
                "text": text,
                "x0": cur_left,
                "x1": cur_right,
                "top": page_height - cur_top_pdf,
                "bottom": page_height - cur_bottom_pdf,
            })

    textpage.close()
    return words


# ===========================================================================
# EMPLOYEE DATA EXTRACTION (same logic as before)
# ===========================================================================
def _find_ef_blocks(words):
    blocks = []
    for w in words:
        if w["top"] > EF_LABEL_MAX_Y: continue
        if w["x0"] > EF_LABEL_MAX_X:  continue
        t = w["text"].lower()
        if t == "e/f" or t.startswith("e/f"):
            blocks.append({"top": w["top"], "x0": w["x0"]})
    return blocks


def _extract_employee_from_block(words, label_top, label_x0):
    name_lines_words = [
        w for w in words
        if label_top + 2 < w["top"] < label_top + 35
        and w["x0"] < 200 and w["x0"] > label_x0 - 5
    ]
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
        return None

    ssn_candidates = []
    for w in words:
        if label_top + SSN_OFFSET_MIN < w["top"] < label_top + SSN_OFFSET_MAX:
            if w["x0"] < 200:
                m = SSN_IN_LINE.search(w["text"])
                if m and SSN_PATTERN.match(m.group(1)):
                    ssn_candidates.append(m.group(1))
    if not ssn_candidates:
        region = [w for w in words
                  if label_top + SSN_OFFSET_MIN < w["top"] < label_top + SSN_OFFSET_MAX
                  and w["x0"] < 250]
        line_text = " ".join(w["text"] for w in sorted(region, key=lambda w: w["x0"]))
        for m in SSN_IN_LINE.finditer(line_text):
            if SSN_PATTERN.match(m.group(1)):
                ssn_candidates.append(m.group(1))
    ssn = ssn_candidates[0] if ssn_candidates else ""

    name = street = csz = ""
    csz_idx = -1
    for i, ln in enumerate(text_lines):
        if CITY_STATE_ZIP_PATTERN.match(ln):
            csz_idx = i; csz = ln; break

    if csz_idx == -1:
        if len(text_lines) >= 1: name   = text_lines[0]
        if len(text_lines) >= 2: street = text_lines[1]
        if len(text_lines) >= 3: csz    = text_lines[2]
    elif csz_idx == 1:
        name = text_lines[0]
    elif csz_idx >= 2:
        name = text_lines[0]
        street = " ".join(text_lines[1:csz_idx]).strip()

    return {"name": name.strip(), "street": street.strip(),
            "city_state_zip": csz.strip(), "ssn": ssn}


# ===========================================================================
# WORKER
# ===========================================================================
def worker_chunk(pdf_path_str, start_page, end_page, out_queue):
    """
    Process pages [start_page, end_page) of a PDF.
    """
    pdf_path = Path(pdf_path_str)
    pages_in_batch = 0
    pages_since_gc = 0
    pdf = None
    try:
        pdf = pdfium.PdfDocument(str(pdf_path))
        n_pages_in_pdf = len(pdf)
        for page_idx in range(start_page, min(end_page, n_pages_in_pdf)):
            page = pdf[page_idx]
            page_num = page_idx + 1
            try:
                words = extract_words_from_page(page)
            except Exception:
                words = []

            if words:
                blocks = _find_ef_blocks(words)
                seen_ssns = set()
                seen_keys = set()
                for block in blocks:
                    r = _extract_employee_from_block(
                        words, block["top"], block["x0"])
                    if not r: continue
                    if r["ssn"] and r["ssn"] in seen_ssns: continue
                    key = (r["name"], r["street"], r["city_state_zip"])
                    if key == ("", "", ""): continue
                    if not r["ssn"] and key in seen_keys: continue
                    if r["ssn"]: seen_ssns.add(r["ssn"])
                    seen_keys.add(key)
                    out_queue.put(("record", {
                        "source_file":    pdf_path.name,
                        "page":           page_num,
                        "ssn":            r["ssn"],
                        "employee_name":  r["name"],
                        "street_address": r["street"],
                        "city_state_zip": r["city_state_zip"],
                    }))

            # Release the page object
            try:
                page.close()
            except Exception:
                pass
            del words

            pages_in_batch += 1
            pages_since_gc += 1

            if pages_in_batch >= PROGRESS_BATCH:
                out_queue.put(("pages", pdf_path.name, pages_in_batch))
                pages_in_batch = 0

            if pages_since_gc >= GC_EVERY_N_PAGES:
                gc.collect()
                pages_since_gc = 0

        if pages_in_batch > 0:
            out_queue.put(("pages", pdf_path.name, pages_in_batch))

        out_queue.put(("chunk_done", pdf_path.name, start_page, end_page, ""))
    except Exception as e:
        if pages_in_batch > 0:
            out_queue.put(("pages", pdf_path.name, pages_in_batch))
        out_queue.put(("chunk_done", pdf_path.name, start_page, end_page,
                       "{}: {}".format(type(e).__name__, e)))
    finally:
        if pdf is not None:
            try: pdf.close()
            except Exception: pass
        gc.collect()


def count_pages(pdf_path_str):
    try:
        pdf = pdfium.PdfDocument(str(pdf_path_str))
        n = len(pdf)
        pdf.close()
        gc.collect()
        return n
    except Exception:
        return 0


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
# DRIVER
# ===========================================================================
def run_extraction(source_folder, dest_folder, max_workers,
                   resume, generate_xlsx, status_callback):
    src = Path(source_folder)
    dst = Path(dest_folder)
    if not src.is_dir():
        status_callback("ERROR: Source folder invalid."); return False
    dst.mkdir(parents=True, exist_ok=True)

    output_csv  = dst / OUTPUT_CSV_NAME
    output_xlsx = dst / OUTPUT_XLSX_NAME
    output_err  = dst / OUTPUT_ERR_NAME

    print("=" * 70)
    print("W-2 Extractor (pypdfium2, chunked parallelism)")
    print("Source:      {}".format(src))
    print("Destination: {}".format(dst))
    print("Max workers: {}  |  Chunk size: {} pages".format(max_workers, CHUNK_SIZE))
    print("=" * 70)

    status_callback("Scanning for PDFs...")
    print("\nScanning {} for PDFs (recursive)...".format(src))
    pdfs = sorted(src.rglob("*.pdf"))
    if not pdfs:
        print("No PDFs found."); status_callback("No PDFs found."); return False
    print("Found {} PDF file(s).".format(len(pdfs)))

    already_done = set()
    if resume:
        already_done = load_already_processed(output_csv)
        if already_done:
            print("Resume: skipping {} file(s) already in output.".format(len(already_done)))
            pdfs = [p for p in pdfs if p.name not in already_done]
            if not pdfs:
                print("Nothing new."); status_callback("Nothing new."); return True

    status_callback("Counting pages in {} file(s)...".format(len(pdfs)))
    print("\nCounting pages...")
    page_counts = {}
    total_pages = 0
    with tqdm(total=len(pdfs), unit="file", desc="Counting", ncols=100) as cbar:
        for p in pdfs:
            n = count_pages(str(p))
            page_counts[str(p)] = n
            total_pages += n
            cbar.update(1)
    print("Total pages: {}".format(total_pages))
    if total_pages == 0:
        status_callback("No readable pages."); return False

    chunks = []
    for p in pdfs:
        n = page_counts[str(p)]
        for start in range(0, n, CHUNK_SIZE):
            end = min(start + CHUNK_SIZE, n)
            chunks.append((str(p), start, end))
    print("Total chunks: {} (~{} pages each)".format(len(chunks), CHUNK_SIZE))

    workers = min(max_workers, len(chunks))
    status_callback("Extracting {} pages in {} chunks with {} worker(s)...".format(
        total_pages, len(chunks), workers))
    print("\nProcessing with {} worker(s)...\n".format(workers))

    csv_has_data = output_csv.exists() and bool(already_done)
    csv_file = output_csv.open("a" if csv_has_data else "w",
                               encoding="utf-8", newline="")
    writer = csv.DictWriter(csv_file, fieldnames=CSV_FIELDS)
    if not csv_has_data:
        writer.writeheader(); csv_file.flush()

    out_queue: Queue = mp.Queue(maxsize=QUEUE_MAX_SIZE)
    chunks_iter = iter(chunks)
    active = {}
    error_log = []
    total_records = 0
    chunks_done = 0
    total_chunks = len(chunks)
    start = time.time()
    last_flush = time.time()

    def launch_next():
        try:
            pdf_str, s, e = next(chunks_iter)
        except StopIteration:
            return False
        proc = Process(target=worker_chunk,
                       args=(pdf_str, s, e, out_queue), daemon=True)
        proc.start()
        active[proc.pid] = proc
        return True

    for _ in range(workers):
        if not launch_next(): break

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
            if kind == "pages":
                _, fname, count = msg
                pbar.update(count)
                pbar.set_postfix(records=total_records,
                                 chunks="{}/{}".format(chunks_done, total_chunks),
                                 errors=len(error_log))
            elif kind == "record":
                _, record = msg
                writer.writerow(record)
                total_records += 1
                if time.time() - last_flush > 2.0:
                    csv_file.flush()
                    last_flush = time.time()
            elif kind == "chunk_done":
                _, fname, s, e, err = msg
                if err:
                    error_log.append("{} [pages {}-{}]: {}".format(fname, s, e, err))
                chunks_done += 1
                for pid, proc in list(active.items()):
                    if not proc.is_alive():
                        proc.join(timeout=1)
                        del active[pid]
                launch_next()
                status_callback("Chunks: {}/{} | Records: {} | Errors: {}".format(
                    chunks_done, total_chunks, total_records, len(error_log)))
                csv_file.flush()
                last_flush = time.time()
    finally:
        pbar.close()
        for pid, proc in list(active.items()):
            if proc.is_alive(): proc.terminate()
            proc.join(timeout=2)
        csv_file.flush(); csv_file.close()

    elapsed = time.time() - start
    pps = total_pages / elapsed if elapsed > 0 else 0
    print("\nDone in {:.1f} min ({:.0f}s) — avg {:.1f} pages/sec".format(
        elapsed/60, elapsed, pps))
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
            print("(Excel skipped: {})".format(e))

    status_callback("Done. {} record(s) in {:.1f} min ({:.1f} pages/sec).".format(
        total_records, elapsed/60, pps))
    return True


# ===========================================================================
# TKINTER GUI
# ===========================================================================
class ExtractorGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("W-2 Extractor (pypdfium2)")
        self.geometry("720x420")
        self.resizable(False, False)
        self._running = False
        self._build_widgets()

    def _build_widgets(self):
        pad = {"padx": 12, "pady": 8}

        ttk.Label(self, text="W-2 Employee Data Extractor",
                  font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(self,
            text="Fast extraction using PDFium engine. Live progress in console.",
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

        ttk.Label(opts, text="Workers (parallel processes):").grid(
            row=0, column=0, sticky="e", padx=8, pady=6)
        self.workers_var = tk.IntVar(value=DEFAULT_WORKERS)
        ttk.Spinbox(opts, from_=1, to=max(1, (os.cpu_count() or 4)),
                    textvariable=self.workers_var, width=5).grid(
            row=0, column=1, sticky="w", padx=4, pady=6)
        ttk.Label(opts,
            text="(default: {}. Lower if system slows; raise if you have plenty of RAM)".format(
                DEFAULT_WORKERS),
            foreground="#888").grid(row=0, column=2, sticky="w", padx=4)

        self.resume_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Resume (skip files already in CSV)",
                        variable=self.resume_var).grid(
            row=1, column=0, columnspan=3, sticky="w", padx=8, pady=6)

        self.xlsx_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Also generate Excel (.xlsx)",
                        variable=self.xlsx_var).grid(
            row=2, column=0, columnspan=3, sticky="w", padx=8, pady=6)

        ttk.Label(self,
            text="ℹ Test on 1 PDF first to confirm output. Keep Task Manager open "
                 "to watch RAM — pypdfium2 uses much less than pdfplumber.",
            foreground="#005a8c",
            wraplength=680, justify="left").grid(
            row=5, column=0, columnspan=3, sticky="w", padx=12, pady=4)

        self.start_btn = ttk.Button(self, text="Start Extraction",
                                    command=self._start_clicked)
        self.start_btn.grid(row=6, column=0, columnspan=3, pady=10)

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(self, textvariable=self.status_var, relief="sunken",
                  anchor="w").grid(row=7, column=0, columnspan=3,
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
        if self._running: return
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
