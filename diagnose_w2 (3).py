"""
W-2 PDF Diagnostic Tool (v2)
=============================
For non-standard W-2 layouts (ADP-style with 3 copies per page).
Anchors on SSN positions instead of the word 'Employee'.
SSNs are auto-redacted in output.

USAGE:
    python diagnose_w2.py
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

# Anchor phrases we'll search for
ANCHOR_PHRASES = [
    "name, address",      # e.g. "Employee's name, address, and ZIP code"
    "name, address,",
    "SSA number",
    "social security number",
    "Employee's",
    "Employees",
    "e/f",
    "Employee's name",
]


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
        "Page", "Which page? (1-{})".format(max_pages),
        initialvalue=1, minvalue=1, maxvalue=max_pages,
    )
    root.destroy()
    return n or 1


def redact(text):
    return SSN_PATTERN.sub("XXX-XX-XXXX", text)


def words_in_box(words, x_min, x_max, y_min, y_max):
    return [w for w in words
            if x_min <= w["x0"] <= x_max and y_min <= w["top"] <= y_max]


def group_into_lines(words, line_tol=4):
    if not words:
        return []
    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines = [[words[0]]]
    for w in words[1:]:
        if abs(w["top"] - lines[-1][0]["top"]) <= line_tol:
            lines[-1].append(w)
        else:
            lines.append([w])
    for line in lines:
        line.sort(key=lambda w: w["x0"])
    return lines


def main():
    print("Opening PDF picker...")
    pdf_path = pick_pdf()
    if not pdf_path:
        print("No file selected."); return
    pdf_path = Path(pdf_path)
    print("File: {}".format(pdf_path.name))
    print("Size: {:.1f} MB".format(pdf_path.stat().st_size / 1024 / 1024))

    print("\n[1/5] Opening PDF...")
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        print("    Total pages: {}".format(total))

        page_num = ask_page_number(total)
        page = pdf.pages[page_num - 1]
        print("\n[2/5] Page {} info".format(page_num))
        print("    Width  x Height: {} x {}".format(page.width, page.height))

        try:
            words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
        except Exception as e:
            print("    ERROR: {}".format(e)); return
        print("    Total words: {}".format(len(words)))

        # ----------------------------------------------------------------
        # 3. Find ALL SSN positions
        # ----------------------------------------------------------------
        print("\n[3/5] SSN locations on page (these will be our anchors)")
        ssn_positions = []
        for w in words:
            if SSN_PATTERN.match(w["text"].strip()):
                ssn_positions.append(w)
        ssn_positions.sort(key=lambda w: (w["top"], w["x0"]))
        print("    Found {} SSN-shaped tokens:".format(len(ssn_positions)))
        for i, w in enumerate(ssn_positions, 1):
            print("       [{}] x0={:.0f} top={:.0f} (value redacted)".format(
                i, w["x0"], w["top"]))

        # ----------------------------------------------------------------
        # 4. Search for known anchor phrases
        # ----------------------------------------------------------------
        print("\n[4/5] Searching for anchor phrases in raw text")
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""
        for phrase in ANCHOR_PHRASES:
            count = text.lower().count(phrase.lower())
            print("    '{}': {} occurrence(s)".format(phrase, count))

        # ----------------------------------------------------------------
        # 5. For EACH SSN, dump nearby words in a box around it
        # ----------------------------------------------------------------
        print("\n[5/5] Words around each SSN (looking for name/address)")
        print("=" * 70)
        for i, ssn in enumerate(ssn_positions, 1):
            print("\nSSN #{} at x0={:.0f} top={:.0f}".format(i, ssn["x0"], ssn["top"]))
            print("-" * 70)
            # Define a search box: same column-ish, extending above and below
            # Box width: about 240pt wide centered on the SSN x
            x_min = ssn["x0"] - 130
            x_max = ssn["x0"] + 130
            # Vertical: 60pt above to 100pt below the SSN
            y_min = ssn["top"] - 60
            y_max = ssn["top"] + 100
            nearby = words_in_box(words, x_min, x_max, y_min, y_max)
            lines = group_into_lines(nearby)
            print("  Box: x=[{:.0f},{:.0f}] y=[{:.0f},{:.0f}], {} lines".format(
                x_min, x_max, y_min, y_max, len(lines)))
            for line in lines:
                text_line = " ".join(w["text"] for w in line)
                marker = ""
                # mark the SSN line for orientation
                if any(SSN_PATTERN.match(w["text"]) for w in line):
                    marker = " <-- SSN LINE"
                print("    top={:.0f} x0={:.0f}  '{}'{}".format(
                    line[0]["top"], line[0]["x0"], redact(text_line), marker))

        print("\n" + "=" * 70)
        print("Diagnostic complete.")
        print("Please redact employee names + addresses before pasting output.")


if __name__ == "__main__":
    main()
