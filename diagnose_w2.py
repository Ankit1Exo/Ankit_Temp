"""
W-2 PDF Diagnostic Tool
========================
Inspects what pdfplumber can actually extract from your W-2 PDF.
SSNs are auto-redacted in the output.

USAGE (run from a prompt/terminal):
    python diagnose_w2.py

A folder picker opens. Pick ONE of your W-2 PDFs.
You'll be asked which page to inspect (default: first).

The script will print:
  1. PDF info (size, page count, image-based check)
  2. Page raw text (SSNs redacted as XXX-XX-XXXX)
  3. Words near the 'Employee' anchor with coordinates
  4. The 'employee block' the current extractor would pick

IMPORTANT: Names and addresses are NOT redacted automatically.
Before sharing the output, please redact employee names + addresses
manually, or just paste the structural parts (coordinates, labels).
"""

from __future__ import annotations
import re
import sys
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    sys.exit("Missing dependency. Run: pip install pdfplumber")

import tkinter as tk
from tkinter import filedialog, simpledialog

SSN_PATTERN = re.compile(r"\b\d{3}-\d{2}-\d{4}\b")

# Show extra detail
SHOW_ALL_WORDS = False  # set True to dump every word with coordinates


def pick_pdf():
    root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
    p = filedialog.askopenfilename(
        title="Pick ONE W-2 PDF to diagnose",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
    )
    root.destroy()
    return p or None


def ask_page_number(max_pages):
    root = tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
    n = simpledialog.askinteger(
        "Page",
        "Which page to inspect? (1-{})".format(max_pages),
        initialvalue=1, minvalue=1, maxvalue=max_pages,
    )
    root.destroy()
    return n or 1


def redact(text):
    """Hide SSNs in any text we print."""
    return SSN_PATTERN.sub("XXX-XX-XXXX", text)


def main():
    print("Opening PDF picker...")
    pdf_path = pick_pdf()
    if not pdf_path:
        print("No file selected."); return
    pdf_path = Path(pdf_path)
    print("File: {}".format(pdf_path.name))
    print("Size: {:.1f} MB".format(pdf_path.stat().st_size / 1024 / 1024))

    print("\n[1/4] Opening PDF...")
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        print("    Total pages: {}".format(total))

        page_num = ask_page_number(total)
        page = pdf.pages[page_num - 1]
        print("\n[2/4] Page {} info".format(page_num))
        print("    Width  x Height: {} x {}".format(page.width, page.height))
        print("    Rotation: {}".format(page.rotation))

        # Check if scanned (image-based)
        images = page.images
        print("    Embedded images on this page: {}".format(len(images)))

        # ----------------------------------------------------------------
        # 2. Raw text
        # ----------------------------------------------------------------
        print("\n[3/4] Raw extracted text (SSNs auto-redacted)")
        print("-" * 70)
        try:
            text = page.extract_text() or ""
        except Exception as e:
            text = ""
            print("    ERROR extracting text: {}".format(e))

        if not text.strip():
            print("    (NO TEXT FOUND — this PDF is likely a scanned image)")
            print("    Image-based PDFs need OCR. Tell me and I can adjust.")
        else:
            redacted = redact(text)
            for line in redacted.splitlines():
                print("    " + line)
        print("-" * 70)

        # ----------------------------------------------------------------
        # 3. Words near 'Employee' anchors with coordinates
        # ----------------------------------------------------------------
        print("\n[4/4] Looking for 'Employee' anchor positions on this page")
        try:
            words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        except Exception as e:
            print("    ERROR extracting words: {}".format(e))
            words = []

        print("    Total words on page: {}".format(len(words)))

        # Find anchor positions
        anchors = []
        for w in words:
            t = w["text"].strip().lower()
            if t in ("employee's", "employee", "employees"):
                anchors.append(w)
        print("    'Employee'-like words found: {}".format(len(anchors)))
        for i, a in enumerate(anchors, 1):
            print("       [{}] text='{}' x0={:.0f} top={:.0f}".format(
                i, a["text"], a["x0"], a["top"]))

        # For each anchor, show the next 25 words below it (within ~150pt vertically,
        # within left half of page width)
        print("\n    --- Words below each 'Employee' anchor (current extractor's view) ---")
        for i, a in enumerate(anchors, 1):
            anchor_top = a["top"]
            anchor_left = a["x0"]
            # Window: 8pt below anchor, 150pt vertical extent, within ~250pt to the right
            window_words = [
                w for w in words
                if anchor_top + 8 < w["top"] < anchor_top + 150
                and anchor_left - 5 < w["x0"] < anchor_left + 250
            ]
            window_words.sort(key=lambda w: (w["top"], w["x0"]))
            print("\n    Anchor [{}] at x={:.0f} top={:.0f}:".format(
                i, anchor_left, anchor_top))
            # Group into lines by top coordinate
            lines = []
            for w in window_words:
                if not lines:
                    lines.append([w]); continue
                if abs(w["top"] - lines[-1][0]["top"]) <= 4:
                    lines[-1].append(w)
                else:
                    lines.append([w])
            for line_idx, line in enumerate(lines[:6], 1):
                line.sort(key=lambda w: w["x0"])
                text_line = " ".join(w["text"] for w in line)
                print("       line {}: top={:.0f}  '{}'".format(
                    line_idx, line[0]["top"], redact(text_line)))

        # SSN locations on page
        print("\n    --- SSN-pattern matches on this page ---")
        ssn_words = []
        for w in words:
            if SSN_PATTERN.match(w["text"].strip()):
                ssn_words.append(w)
        print("    SSN-shaped tokens found: {}".format(len(ssn_words)))
        for w in ssn_words:
            print("       'XXX-XX-XXXX' at x0={:.0f} top={:.0f}".format(w["x0"], w["top"]))

        if SHOW_ALL_WORDS:
            print("\n    --- ALL WORDS (verbose) ---")
            for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
                print("       top={:.0f} x0={:.0f}  '{}'".format(
                    w["top"], w["x0"], redact(w["text"])))

    print("\nDiagnostic complete.")
    print("Please redact employee names + addresses before sharing this output.")


if __name__ == "__main__":
    main()
