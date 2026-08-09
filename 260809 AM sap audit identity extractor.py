r"""
SAP Audit Identity Extractor -- Tkinter GUI
==========================================================
Reads "Satisfactory Academic Progress Audit Report" PDFs and writes the
student ID, SSN and Name to an Excel workbook.

THE WHOLE RULE
    Read every line of every page. Any line holding an SSN is a student
    line, and it is split in printed order:

        1234567  123-45-6789  Miss Dishita Uppal   Academic Program: ...
        \_ ID _/ \__ SSN ___/ \____ Name ______/   \_ ignored _______/

        ID    everything to the left of the SSN
        SSN   the SSN itself, kept exactly as printed
        Name  what follows, up to "Academic Program" (or another caption,
              or the first number -- whichever comes first)

    The Name is then split:

        with a middle initial   "Dishita D. Uppal"
            First = Dishita     Middle = D.       Last = Uppal
        without one             "Dishita Uppal"
            First = Dishita     Middle = (blank)  Last = Uppal

    An initial, when present, is the split point: what precedes it is the
    first name, what follows it is the last name. With no initial, the
    first space is the split point. Any Mr./Mrs./Ms./Miss/Dr. prefix is
    moved to its own column first.

TWO THINGS THAT ARE NOT OBVIOUS, AND WHY THE CODE DOES THEM
    1. Lines are rebuilt from word COORDINATES, not from PyMuPDF's own
       line breaking. PyMuPDF emits text in the PDF's internal block
       order, which on this report mixes the course headings and the
       "#Excluded Remedial Credits" row into the student line.

    2. Every dash is folded to an ASCII hyphen before matching. This
       report's SSNs may be printed with U+2010, U+2013 or U+2212 -- all
       identical on screen, none of them matched by a regex written with
       "-". This silently hid most SSNs.

    Everything else here is the rule above, written out.

USAGE
    python "260809 AM sap audit identity extractor.py"
    python "260809 AM sap audit identity extractor.py" --debug <pdf> [page]

    --debug prints every rebuilt line of a page with the values masked to
    their shape (# for a digit, X for a letter) and only the report's own
    captions left readable, so a layout problem can be diagnosed and
    pasted into a ticket without exposing PII.

REQUIREMENTS
    pip install pymupdf pandas openpyxl tqdm

SECURITY NOTE
    These reports contain SSNs and student names, and the output workbook
    holds them in clear. Run only on an authorised workstation, save the
    XLSX only to the approved Global Insider folder (never a desktop or a
    local temp path), and delete the local copy once it has been loaded
    into the authorised system of record.
"""
from __future__ import annotations

import re
import threading
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pymupdf as fitz
import pandas as pd
from tqdm import tqdm

OUTPUT_XLSX_NAME = "sap_audit_identity.xlsx"

OUTPUT_COLUMNS = [
    "File Name", "Page Number", "ID", "SSN",
    "Prefix", "Full Name", "First Name", "Middle", "Last Name", "Extraction Notes",
]

# 123-45-6789, the masked XXX-XX-6789, or a running 123456789.
SSN_RE = re.compile(r"[0-9X*#?]{3}-[0-9X*#?]{2}-[0-9X*#?]{4}|(?<!\d)\d{9}(?!\d)")

# A whole token that is SSN-shaped, allowing spaces or nothing where the
# dashes would be. Used to decide whether the token sitting between the ID
# and the name really is the SSN, rather than to find one.
SSN_TOKEN_RE = re.compile(r"^[0-9X*#?]{3}[-]?[0-9X*#?]{2}[-]?[0-9X*#?]{4}$")

# Every student line carries this caption, which makes it a far more
# dependable anchor than the SSN: it says both "this is a student line"
# and "the name stops here", without depending on the SSN being printed in
# any particular format.
ACADEMIC_RE = re.compile(r"\bacademic\s*program\b", re.IGNORECASE)

# Requiring a period or a following capital matters: a bare "^(Mr|Mrs|Ms)"
# would also strike the first two letters off surnames like "Mroz".
PREFIX_RE = re.compile(r"^((?:Mr|Mrs|Ms|Miss|Dr|Prof)\.?)(?:\s+|(?=[A-Z]))")

INITIAL_RE = re.compile(r"^[A-Za-z]\.?$")

# The name ends at the first token that is one of these captions, holds a
# digit, or ends in a colon.
NAME_STOP_WORDS = {"academic", "program", "sap", "type", "excluded", "remedial",
                   "credits", "credit", "incl", "gpa", "status", "degree", "major"}

# Dash-like and space-like characters folded to ASCII before matching.
FOLD = {
    **dict.fromkeys(map(ord, "‐‑‒–—―−⁃﹘﹣－­"), "-"),
    **dict.fromkeys(map(ord, "     　"), " "),
}


def page_lines(page):
    """Every printed line on the page, top to bottom, as plain text.

    Words are grouped by the vertical centre of their bounding box, with a
    tolerance of 40% of the page's median glyph height -- enough to absorb
    the sub-point jitter inside one printed line, well under the gap to
    the next one. Each word is compared to the ANCHOR of the line it might
    join, not to its predecessor, so a run of drifting words cannot chain
    two printed lines into one."""
    words = page.get_text("words")
    if not words:
        return []

    heights = sorted(w[3] - w[1] for w in words)
    tolerance = (heights[len(heights) // 2] or 1.0) * 0.4

    items = sorted(((w[1] + w[3]) / 2.0, w[0], w[4].translate(FOLD)) for w in words)

    lines, current, anchor = [], [], None
    for centre, x0, text in items:
        if anchor is None or centre - anchor <= tolerance:
            anchor = centre if anchor is None else anchor
            current.append((x0, text))
        else:
            lines.append(current)
            current, anchor = [(x0, text)], centre
    if current:
        lines.append(current)

    return [" ".join(t for _, t in sorted(line)) for line in lines]


def split_name(name, notes):
    """(first, middle, last) from a prefix-free name."""
    tokens = name.split()
    if not tokens:
        return "", "", ""
    if len(tokens) == 1:
        notes.append("name is a single word -- put in First Name, Last Name left blank")
        return tokens[0], "", ""

    # An initial is the split point only if it is neither the first token
    # nor the last: a trailing lone letter ("Mroz Dana K") has no surname
    # after it to split off.
    start = next((i for i in range(1, len(tokens) - 1) if INITIAL_RE.match(tokens[i])), None)
    if start is None:
        return tokens[0], "", " ".join(tokens[1:])

    end = start
    while end + 1 < len(tokens) - 1 and INITIAL_RE.match(tokens[end + 1]):
        end += 1
    return " ".join(tokens[:start]), " ".join(tokens[start:end + 1]), " ".join(tokens[end + 1:])


def build_row(id_text, ssn_text, name_tokens, notes):
    """Assemble a row from the three pieces, or None if there is no name."""
    # Trim anything past the name: a caption, a number, or "Something:".
    stop = len(name_tokens)
    for i, token in enumerate(name_tokens):
        if (token.strip(":.,#").lower() in NAME_STOP_WORDS
                or re.search(r"\d", token) or token.endswith(":")):
            stop = i
            break

    joined = " ".join(name_tokens[:stop])
    prefix_match = PREFIX_RE.match(joined)
    prefix = prefix_match.group(1) if prefix_match else ""
    full_name = (joined[prefix_match.end():] if prefix_match else joined).strip()
    if not full_name:
        return None

    row = dict.fromkeys(OUTPUT_COLUMNS, "")
    row["ID"], row["SSN"], row["Prefix"], row["Full Name"] = id_text, ssn_text, prefix, full_name
    row["First Name"], row["Middle"], row["Last Name"] = split_name(full_name, notes)
    if not ssn_text:
        notes.append("no SSN-shaped value between the ID and the name")
    if not id_text:
        notes.append("no ID printed before the SSN")
    row["Extraction Notes"] = "; ".join(notes)
    return row


def parse_by_caption(text):
    """Split a line that carries the "Academic Program" caption.

    Everything before the caption is the student's own data, in printed
    order: ID, then SSN, then Name. Anchoring on the caption rather than
    on the SSN means the line is still read correctly when the SSN is
    printed in a format the SSN pattern does not recognise, or is missing
    altogether."""
    caption = ACADEMIC_RE.search(text)
    if not caption:
        return None

    tokens = text[:caption.start()].split()
    if len(tokens) < 2:
        return None

    notes = []
    id_text, ssn_text, rest = tokens[0], "", tokens[1:]

    # The ID and SSN can be printed hard against each other, so check for
    # an SSN inside the first token before looking at the second.
    inside = SSN_RE.search(id_text)
    if inside:
        ssn_text = inside.group(0)
        id_text = id_text[:inside.start()]
    elif SSN_TOKEN_RE.match(rest[0]):
        ssn_text, rest = rest[0], rest[1:]
    elif re.fullmatch(r"[0-9X*#?-]+", rest[0]):
        # The SSN may be printed as separate groups ("555 12 3456"). Pull
        # in following groups until nine digits have been seen, keeping the
        # spacing as printed.
        ssn_text, rest = rest[0], rest[1:]
        while (rest and re.fullmatch(r"[0-9X*#?-]+", rest[0])
               and sum(c.isdigit() for c in ssn_text) < 9):
            ssn_text, rest = f"{ssn_text} {rest[0]}", rest[1:]
    elif re.search(r"\d", rest[0]):
        # Not the shape expected, but a person's name never contains a
        # digit, so whatever sits between the ID and the name is the SSN
        # printed some other way. Take it exactly as printed and flag it,
        # rather than losing the whole student over its format.
        ssn_text, rest = rest[0], rest[1:]
        notes.append("the value between the ID and the name is not a recognised SSN format -- "
                     "copied exactly as printed; verify it")
    else:
        notes.append("no SSN-shaped value between the ID and the name -- the text straight after "
                     "the ID was treated as the start of the name")

    return build_row(id_text, ssn_text, rest, notes)


def parse_by_ssn(text):
    """Split a line that has no "Academic Program" caption, using the SSN
    itself as the divider: ID to its left, name to its right."""
    match = SSN_RE.search(text)
    if not match:
        return None

    before = text[:match.start()].split()
    notes = []
    id_text = ""
    if before:
        id_text = before[-1]
        if len(before) > 1:
            notes.append(f"more than one token left of the SSN ('{' '.join(before)}') -- "
                         "took the one nearest the SSN as the ID")
    return build_row(id_text, match.group(0), text[match.end():].split(), notes)


def parse_line(text):
    """A student row from one line, or None if it is not a student line.

    The caption is tried first because it is the stronger signal; the SSN
    is the fallback for any student line that does not carry it."""
    return parse_by_caption(text) or parse_by_ssn(text)


def process_pdf(path: Path):
    """(rows, pages_with_no_text_layer)."""
    doc = fitz.open(str(path))
    rows, image_only = [], []
    for page_num, page in enumerate(doc, start=1):
        if len(page.get_text().strip()) < 20:
            image_only.append(page_num)
            continue
        for text in page_lines(page):
            row = parse_line(text)
            if row:
                row["File Name"], row["Page Number"] = path.name, page_num
                rows.append(row)
    doc.close()
    return rows, image_only


# ===========================================================================
# DEBUG -- masked dump of one page
# ===========================================================================
CAPTIONS = ["ID", "SSN", "Name", "Academic Program", "SAP Type", "Excluded Remedial Credits",
            "Course Name", "Term/Dt", "Grd", "Cum", "Eval", "Credits", "Grade Pts", "Incl",
            "Report Options", "Batch ID", "Page", "Att", "Pgm", "Earn", "Cmpl", "GPA",
            "Satisfactory Academic Progress Audit Report", "Section skipped"]


def mask(text):
    tokens = text.split()
    lowered = [t.strip(":").lower() for t in tokens]
    keep = [False] * len(tokens)
    for caption in sorted(CAPTIONS, key=len, reverse=True):
        parts = caption.lower().split()
        for i in range(len(tokens) - len(parts) + 1):
            if not any(keep[i:i + len(parts)]) and lowered[i:i + len(parts)] == parts:
                for k in range(i, i + len(parts)):
                    keep[k] = True
    return " ".join(t if keep[i] else re.sub(r"[A-Za-z]", "X", re.sub(r"\d", "#", t))
                    for i, t in enumerate(tokens))


def write_diagnostic(source_folder, dest_folder, max_files=3, max_pages=2):
    """Write a masked, PII-free dump of the first few pages to a text file.

    Exists because the command-line --debug was never going to get run by
    someone working in the GUI. When extraction comes back short, this is
    the file to send on: it shows exactly what the parser sees, with every
    value reduced to its shape."""
    src, dst = Path(source_folder), Path(dest_folder)
    dst.mkdir(parents=True, exist_ok=True)
    report_path = dst / "sap_audit_diagnostic.txt"

    pdfs = sorted(src.rglob("*.pdf"))[:max_files]
    out = ["SAP Audit Identity Extractor -- diagnostic",
           "Every value below is masked: # = a digit, X = a letter.",
           "Only the report's own captions are left readable. Safe to share.",
           f"(first {max_pages} page(s) of the first {max_files} file(s))", ""]

    if not pdfs:
        out.append(f"No PDF files found in {src}")

    for path in pdfs:
        doc = fitz.open(str(path))
        out.append(f"=== {path.name} ({len(doc)} pages) ===")
        for page_num, page in enumerate(doc, start=1):
            if page_num > max_pages:
                break
            chars = len(page.get_text().strip())
            lines = page_lines(page)
            out.append(f"-- page {page_num}: {chars} chars of text, {len(lines)} lines rebuilt"
                       + ("" if chars >= 20 else "   <-- SCANNED IMAGE, no text layer; "
                                                 "needs OCR before anything can be read"))
            for i, text in enumerate(lines):
                row = parse_line(text)
                flag = ("   <-- STUDENT" if row else
                        "   <-- has an SSN, no name after it" if SSN_RE.search(text) else "")
                out.append(f"  [{i:>3}] {mask(text)}{flag}")
            out.append(f"-- students found on page {page_num}: "
                       f"{sum(1 for t in lines if parse_line(t))}")
            out.append("")
        doc.close()

    report_path.write_text("\n".join(out), encoding="utf-8")
    return report_path


def debug_page(path: Path, page_num: int):
    doc = fitz.open(str(path))
    if not 1 <= page_num <= len(doc):
        print(f"{path.name}: page {page_num} out of range ({len(doc)} page(s))")
        doc.close()
        return

    page = doc[page_num - 1]
    chars = len(page.get_text().strip())
    lines = page_lines(page)
    print(f"--- {path.name} page {page_num} ---")
    print(f"text layer: {chars} chars"
          + ("" if chars >= 20 else "  <-- scanned image, no text; needs OCR first"))
    print(f"lines rebuilt: {len(lines)}")

    found = 0
    for i, text in enumerate(lines):
        row = parse_line(text)
        flag = ""
        if row:
            found += 1
            flag = "  <-- STUDENT"
        elif SSN_RE.search(text):
            flag = "  <-- has an SSN but no name after it"
        print(f"  [{i:>3}] {mask(text)}{flag}")
    print(f"students found: {found}")
    doc.close()


# ===========================================================================
# RUNNER
# ===========================================================================
def run_extraction(source_folder, dest_folder, status_callback):
    src, dst = Path(source_folder), Path(dest_folder)
    if not src.is_dir():
        status_callback("ERROR: Source folder invalid.")
        return False
    dst.mkdir(parents=True, exist_ok=True)
    output_path = dst / OUTPUT_XLSX_NAME

    print("=" * 70)
    print(f"SAP Audit Identity Extractor\nSource:      {src}\nDestination: {dst}")
    print("=" * 70)

    pdfs = sorted(src.rglob("*.pdf"))
    if not pdfs:
        print(f"No PDF files found in {src}")
        status_callback("No PDFs found in source folder.")
        return False
    print(f"Found {len(pdfs)} PDF file(s).")

    all_rows, image_only, empty_files = [], [], []
    with tqdm(pdfs, desc="Extracting", unit="pdf", ncols=100) as pbar:
        for pdf_path in pbar:
            pbar.set_postfix_str(pdf_path.name)
            rows, image_pages = process_pdf(pdf_path)
            if not rows:
                empty_files.append(pdf_path.name)
            if image_pages:
                image_only.append(f"{pdf_path.name} (page(s) {', '.join(map(str, image_pages))})")
            all_rows.extend(rows)
            status_callback(f"{pdf_path.name}: {len(all_rows)} student(s) so far")

    if not all_rows:
        print("No student lines found -- nothing to write.")
        status_callback("Done. No student lines found.")
        return False

    frame = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS)
    # As text, so Excel cannot drop a leading zero or turn a running
    # 9-digit SSN into scientific notation.
    for col in ("ID", "SSN"):
        frame[col] = frame[col].astype(str)
    frame.to_excel(output_path, index=False)

    flagged = sum(1 for r in all_rows if r["Extraction Notes"])
    files_with_rows = len({r["File Name"] for r in all_rows})
    print(f"\nDone. {len(all_rows)} student(s) from {files_with_rows} of {len(pdfs)} file(s) "
          f"-> {output_path}")
    if flagged:
        print(f"{flagged} row(s) have an Extraction Note -- spot-check those.")
    if image_only:
        print(f"Scanned pages with no text layer (OCR them first): {'; '.join(image_only)}")
    if empty_files:
        print(f"{len(empty_files)} file(s) produced nothing: {', '.join(empty_files)}")
        print('   -> run:  --debug "<that file>" 1   to see what the page actually looks like.')

    status_callback(f"Done. {len(all_rows)} student(s) -> {output_path.name}")
    return True


# ===========================================================================
# GUI
# ===========================================================================
class ExtractorGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SAP Audit Identity Extractor")
        self.geometry("640x290")
        self.resizable(False, False)
        self._running = False
        self._build()

    def _build(self):
        pad = {"padx": 12, "pady": 8}
        ttk.Label(self, text="SAP Audit Identity Extractor",
                  font=("Segoe UI", 14, "bold")).grid(row=0, column=0, columnspan=3,
                                                      sticky="w", **pad)
        ttk.Label(self, text="Extracts ID, SSN and Name. Progress prints to the console.",
                  foreground="#555").grid(row=1, column=0, columnspan=3, sticky="w", padx=12)
        ttk.Label(self, text="Output contains SSNs -- save to the approved Global Insider folder only.",
                  foreground="#a33").grid(row=2, column=0, columnspan=3, sticky="w", padx=12)

        ttk.Label(self, text="Source folder (PDFs):").grid(row=3, column=0, sticky="e", **pad)
        self.src_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.src_var, width=50).grid(row=3, column=1, sticky="we", **pad)
        ttk.Button(self, text="Browse...", command=self._pick_src).grid(row=3, column=2, **pad)

        ttk.Label(self, text="Destination folder:").grid(row=4, column=0, sticky="e", **pad)
        self.dst_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.dst_var, width=50).grid(row=4, column=1, sticky="we", **pad)
        ttk.Button(self, text="Browse...", command=self._pick_dst).grid(row=4, column=2, **pad)

        buttons = ttk.Frame(self)
        buttons.grid(row=5, column=0, columnspan=3, pady=12)
        self.start_btn = ttk.Button(buttons, text="Start Extraction", command=self._start)
        self.start_btn.pack(side="left", padx=6)
        ttk.Button(buttons, text="Save Diagnostic (masked)",
                   command=self._diagnostic).pack(side="left", padx=6)

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w").grid(
            row=6, column=0, columnspan=3, sticky="we", padx=12, pady=(0, 12))
        self.columnconfigure(1, weight=1)

    def _pick_src(self):
        folder = filedialog.askdirectory(title="Select folder containing the report PDFs",
                                         mustexist=True)
        if folder:
            self.src_var.set(folder)
            if not self.dst_var.get():
                self.dst_var.set(folder)

    def _pick_dst(self):
        folder = filedialog.askdirectory(title="Select destination folder for the XLSX")
        if folder:
            self.dst_var.set(folder)

    def _status(self, msg):
        self.after(0, lambda: self.status_var.set(msg))

    def _start(self):
        if self._running:
            return
        src, dst = self.src_var.get().strip(), self.dst_var.get().strip()
        if not src or not Path(src).is_dir():
            messagebox.showerror("Missing source", "Please select a valid source folder.")
            return
        if not dst:
            messagebox.showerror("Missing destination", "Please select a destination folder.")
            return
        self._running = True
        self.start_btn.config(state="disabled", text="Running...")
        threading.Thread(target=self._run, args=(src, dst), daemon=True).start()

    def _diagnostic(self):
        src = self.src_var.get().strip()
        dst = self.dst_var.get().strip() or src
        if not src or not Path(src).is_dir():
            messagebox.showerror("Missing source", "Please select a valid source folder first.")
            return
        try:
            path = write_diagnostic(src, dst)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        self._status(f"Diagnostic written to {path.name}")
        messagebox.showinfo(
            "Diagnostic saved",
            f"Written to:\n{path}\n\nEvery value is masked (# = digit, X = letter) and only the "
            f"report's own captions are readable, so this file is safe to share.\n\n"
            f"If extraction came back short, send this file on.")

    def _run(self, src, dst):
        try:
            run_extraction(src, dst, self._status)
        except Exception as e:
            print(f"\nFATAL ERROR: {e}")
            self._status(f"Error: {e}")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.after(0, self._done)

    def _done(self):
        self._running = False
        self.start_btn.config(state="normal", text="Start Extraction")


def main():
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "--debug":
        if len(sys.argv) < 3:
            print("Usage: --debug <pdf_file_or_folder> [page_number]")
            return
        target = Path(sys.argv[2])
        page_num = int(sys.argv[3]) if len(sys.argv) > 3 else 1
        for pdf in ([target] if target.is_file() else sorted(target.rglob("*.pdf"))):
            debug_page(pdf, page_num)
        return
    ExtractorGUI().mainloop()


if __name__ == "__main__":
    main()
