"""
Extract per-student summary records from academic transcript/SAP report PDFs
into Excel.

Written blind against a masked screenshot -- column labels/positions are a
best-effort reconstruction, NOT verified against a real file. Run with
DEBUG_MODE = True first and pick ONE real file in the dialog to check the
printed output before trusting the full batch. The output keeps raw debug
columns (Raw_Header_Line / Raw_Trailer_Line) specifically so you can
eyeball each row against the source PDF and catch misparses early.

Each output row is ONE student (summary only -- individual course lines are
not extracted). A record is bounded by:
  - start: the line carrying that student's ID + SSN + Name (+ Academic
    Program / cumulative GPA figure), which follows the two-row column
    header and dashed separator
  - end: the trailer line carrying "#Excluded Remedial Credits" / "SAP
    Type: ..."
If a file has more than one student block (e.g. a batch report), each block
becomes its own row.

NOTE -- full SSNs are kept unmasked in the output workbook per explicit
request. That means the .xlsx this script produces is a plain-text store of
PII: keep it in an approved/secured location only, not on a shared drive,
desktop, or emailed around.

Install deps once:
    pip install pdfplumber pandas openpyxl
"""

import re
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

import pdfplumber
import pandas as pd

# ---------------------------------------------------------------------------
# CONFIG -- edit these before running
# ---------------------------------------------------------------------------

DEBUG_MODE = False  # <-- set True only to sanity-check parsing against a single PDF you pick

ROW_TOP_TOLERANCE = 3  # points; words within this many points of top are treated as the same row

SSN_RE = re.compile(r"\b\d{3}-(?:\d{2}|[Xx\*]{2})-\d{4}\b")
ACADEMIC_PROGRAM_RE = re.compile(r"Academic Program\s*:\s*(?P<program>.+?)\s*=\s*(?P<gpa>[\d.]+)\)")
SAP_TRAILER_RE = re.compile(
    r"#\s*Excluded Remedial Credits.*?SAP Type\s*:\s*(?P<sap_type>\S+)\s*(?P<sap_desc>.*)$"
)
EXCLUDED_REMEDIAL_RE = re.compile(r"#\s*Excluded Remedial Credits")


# ---------------------------------------------------------------------------
# Word-coordinate extraction (robust to PDF-to-text spacing quirks)
# ---------------------------------------------------------------------------

def extract_rows(pdf_path):
    """Return every row (across all pages) as {'text': str, 'page': int}, rows
    reconstructed from word x/y coordinates rather than pdfplumber's raw text
    extraction, since fixed-width reports often collapse/expand whitespace
    inconsistently between the header and data rows."""
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            words.sort(key=lambda w: (round(w["top"] / ROW_TOP_TOLERANCE), w["x0"]))
            current_top = None
            current = []
            for w in words:
                if current_top is None or abs(w["top"] - current_top) > ROW_TOP_TOLERANCE:
                    if current:
                        rows.append(_finalize_row(current, page.page_number))
                    current = [w]
                    current_top = w["top"]
                else:
                    current.append(w)
            if current:
                rows.append(_finalize_row(current, page.page_number))
    return rows


def _finalize_row(words, page_number):
    words = sorted(words, key=lambda w: w["x0"])
    return {"text": " ".join(w["text"] for w in words), "page": page_number}


# ---------------------------------------------------------------------------
# Record parsing
# ---------------------------------------------------------------------------

def chunk_records(rows):
    """Split rows into per-student chunks, starting each chunk at the row
    carrying that student's SSN."""
    start_indices = [i for i, r in enumerate(rows) if SSN_RE.search(r["text"])]
    chunks = []
    for n, start in enumerate(start_indices):
        end = start_indices[n + 1] if n + 1 < len(start_indices) else len(rows)
        chunks.append(rows[start:end])
    return chunks


def parse_record_chunk(chunk_rows, source_name):
    header_row = chunk_rows[0]
    header_text = header_row["text"]

    ssn_match = SSN_RE.search(header_text)
    ssn = ssn_match.group(0) if ssn_match else ""

    id_token = ""
    if ssn_match:
        before = header_text[: ssn_match.start()].strip().split()
        id_token = before[-1] if before else ""

    program = ""
    gpa_summary = ""
    name = ""
    prog_match = ACADEMIC_PROGRAM_RE.search(header_text)
    if prog_match:
        program = prog_match.group("program").strip()
        gpa_summary = prog_match.group("gpa").strip()
        if ssn_match:
            name = header_text[ssn_match.end(): prog_match.start()].strip()
    elif ssn_match:
        name = header_text[ssn_match.end():].strip()

    trailer_row = next((r for r in chunk_rows if EXCLUDED_REMEDIAL_RE.search(r["text"])), None)
    sap_type, sap_desc, excluded_remedial = "", "", False
    trailer_text = ""
    if trailer_row is not None:
        trailer_text = trailer_row["text"]
        excluded_remedial = True
        sap_match = SAP_TRAILER_RE.search(trailer_text)
        if sap_match:
            sap_type = sap_match.group("sap_type").strip()
            sap_desc = sap_match.group("sap_desc").strip()

    notes = []
    if not ssn_match:
        notes.append("No SSN pattern matched on header line")
    if not prog_match:
        notes.append("Academic Program / GPA figure not matched")
    if trailer_row is None:
        notes.append("Trailer line (#Excluded Remedial Credits / SAP Type) not found in block")

    return {
        # Column order matters: A=File Name, B=Page Number, C=SSN, D=Name, per spec.
        "File Name": source_name,
        "Page Number": header_row["page"],
        "SSN": ssn,
        "Name": name,
        "ID": id_token,
        "Academic_Program": program,
        "GPA_Summary": gpa_summary,
        "SAP_Type": sap_type,
        "SAP_Description": sap_desc,
        "Excluded_Remedial_Credits": excluded_remedial,
        "Notes": "; ".join(notes),
        "Raw_Header_Line": header_text,
        "Raw_Trailer_Line": trailer_text,
    }


def records_from_pdf(pdf_path):
    rows = extract_rows(pdf_path)
    chunks = chunk_records(rows)
    return [parse_record_chunk(chunk, pdf_path.name) for chunk in chunks]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def prompt_for_folders():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    source = filedialog.askdirectory(title="Select folder containing source PDF files")
    if not source:
        messagebox.showwarning("Cancelled", "No source folder selected -- exiting.")
        root.destroy()
        return None, None

    destination = filedialog.askdirectory(title="Select destination folder for the Excel output")
    root.destroy()
    if not destination:
        print("No destination folder selected -- exiting.")
        return None, None

    return Path(source), Path(destination)


def prompt_for_debug_file():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title="Select a PDF file to debug", filetypes=[("PDF files", "*.pdf")]
    )
    root.destroy()
    return Path(path) if path else None


def format_duration(seconds):
    seconds = int(max(seconds, 0))
    minutes, seconds = divmod(seconds, 60)
    return f"{minutes:02d}:{seconds:02d}"


def process_with_progress(pdf_files, output_dir):
    root = tk.Tk()
    root.title("Extracting Academic Transcripts")
    root.geometry("420x120")
    root.attributes("-topmost", True)

    status_var = tk.StringVar(value="Starting...")
    tk.Label(root, textvariable=status_var).pack(pady=(15, 5))

    progress = ttk.Progressbar(root, orient="horizontal", length=380, mode="determinate", maximum=len(pdf_files))
    progress.pack(pady=5)

    eta_var = tk.StringVar(value="")
    tk.Label(root, textvariable=eta_var).pack(pady=(0, 10))

    start_time = time.time()
    all_rows = []

    for i, pdf_path in enumerate(pdf_files, start=1):
        status_var.set(f"Processing {pdf_path.name}  ({i} of {len(pdf_files)})")
        root.update()

        rows = records_from_pdf(pdf_path)
        df = pd.DataFrame(rows)
        out_name = f"{pdf_path.stem}_extracted.xlsx"
        df.to_excel(output_dir / out_name, index=False)
        all_rows.extend(rows)
        print(f"Processed {pdf_path.name}: {len(rows)} student record(s) -> {out_name}")

        progress["value"] = i
        elapsed = time.time() - start_time
        remaining = (elapsed / i) * (len(pdf_files) - i)
        eta_var.set(f"Elapsed: {format_duration(elapsed)}   ETA: {format_duration(remaining)}")
        root.update()

    status_var.set("Writing combined workbook...")
    eta_var.set("")
    root.update()

    combined_df = pd.DataFrame(all_rows)
    combined_name = "combined_extracted.xlsx"
    combined_df.to_excel(output_dir / combined_name, index=False)
    print(f"Combined workbook -> {combined_name} ({len(all_rows)} rows)")

    root.destroy()
    return all_rows, combined_name


def main():
    if DEBUG_MODE:
        debug_file = prompt_for_debug_file()
        if debug_file is None:
            print("No file selected -- exiting.")
            return
        for row in extract_rows(debug_file):
            print(row["text"])
        print("---- parsed records ----")
        for r in records_from_pdf(debug_file):
            print(r)
        return

    source_dir, output_dir = prompt_for_folders()
    if source_dir is None:
        return
    output_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = sorted(source_dir.glob("*.pdf"))
    if not pdf_files:
        print(f"No PDF files found in {source_dir}")
        return

    process_with_progress(pdf_files, output_dir)


if __name__ == "__main__":
    main()
