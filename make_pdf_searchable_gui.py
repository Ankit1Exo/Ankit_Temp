"""
Make Scanned PDF(s) Searchable — with file/folder picker GUI
--------------------------------------------------------------
Works for ONE file or MANY files in one go. On launch, a small window
lets you choose:
    - Pick File(s)   -> select one or more individual PDFs
    - Pick Folder    -> select a folder; every .pdf inside it is processed
Then it asks you to choose an OUTPUT folder where the searchable PDFs
will be saved (original filenames + "_searchable.pdf").

WHAT IT DOES (per file):
1. Renders each page to an image (pdf2image / poppler).
2. Runs Tesseract OCR to get each word's text + position.
3. Rebuilds the page: original image on top + an invisible text layer
   underneath, positioned to match the image exactly.
4. Saves a new PDF that looks identical but is fully searchable/selectable.

------------------------------------------------------------------
SETUP (do this once)
------------------------------------------------------------------
1. Install Python packages:
       pip install pytesseract pdf2image reportlab pypdf pillow

2. Install Tesseract OCR (the actual OCR engine, not a Python package):
       Windows: https://github.com/UB-Mannheim/tesseract/wiki  (installer)
                after installing, note the install path, e.g.
                C:\\Program Files\\Tesseract-OCR\\tesseract.exe
       Mac:     brew install tesseract

3. Install Poppler (required by pdf2image to render PDF pages):
       Windows: download from https://github.com/oschwartz10612/poppler-windows/releases
                unzip it somewhere, e.g. C:\\poppler
                then set POPPLER_PATH below to the "Library\\bin" folder inside it,
                e.g. C:\\poppler\\Library\\bin
       Mac:     brew install poppler   (no extra config needed)

4. If Windows and tesseract.exe is not on your PATH, set TESSERACT_PATH below.

Then just run:
       python make_pdf_searchable_gui.py
------------------------------------------------------------------
"""

import io
import os
import sys
import traceback

# ============ EDIT THESE TWO LINES ON WINDOWS IF NEEDED ============
TESSERACT_PATH = None   # e.g. r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPPLER_PATH = None     # e.g. r"C:\poppler\Library\bin"
# =====================================================================

import tkinter as tk
from tkinter import filedialog, messagebox

import pytesseract
from pdf2image import convert_from_path
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from pypdf import PdfReader, PdfWriter

if TESSERACT_PATH:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

DPI = 300  # higher = better OCR accuracy, slower, bigger files


# ---------------------------------------------------------------------------
# Core OCR / PDF-building logic (per page, per file)
# ---------------------------------------------------------------------------

def ocr_page_to_pdf_page(image: Image.Image) -> bytes:
    """Given a page image, OCR it and return a single-page PDF (bytes) with
    the image on top and an invisible, correctly-positioned text layer."""
    width_px, height_px = image.size
    width_pt = width_px * 72.0 / DPI
    height_pt = height_px * 72.0 / DPI

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(width_pt, height_pt))
    c.drawImage(ImageReader(image), 0, 0, width=width_pt, height=height_pt)

    ocr_data = pytesseract.image_to_data(image, output_type=pytesseract.Output.DICT)

    text_object = c.beginText()
    text_object.setTextRenderMode(3)  # invisible but selectable/searchable

    for i in range(len(ocr_data["text"])):
        word = ocr_data["text"][i].strip()
        if not word:
            continue

        conf_raw = ocr_data["conf"][i]
        try:
            conf = int(float(conf_raw))
        except (ValueError, TypeError):
            conf = -1
        if conf < 30:
            continue

        x_px, y_px = ocr_data["left"][i], ocr_data["top"][i]
        w_px, h_px = ocr_data["width"][i], ocr_data["height"][i]

        x_pt = x_px * 72.0 / DPI
        y_pt = height_pt - (y_px + h_px) * 72.0 / DPI
        box_width_pt = w_px * 72.0 / DPI
        box_height_pt = h_px * 72.0 / DPI
        if box_height_pt <= 0:
            continue

        font_size = box_height_pt * 0.9
        c.setFont("Helvetica", font_size)

        text_width = c.stringWidth(word, "Helvetica", font_size)
        h_scale = (box_width_pt / text_width) * 100.0 if text_width > 0 else 100.0

        text_object.setTextOrigin(x_pt, y_pt)
        text_object.setHorizScale(h_scale)
        text_object.textOut(word)

    c.drawText(text_object)
    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.read()


def make_searchable(input_path: str, output_path: str, log=print):
    convert_kwargs = {"dpi": DPI}
    if POPPLER_PATH:
        convert_kwargs["poppler_path"] = POPPLER_PATH

    log(f"  Rendering pages of '{os.path.basename(input_path)}'...")
    images = convert_from_path(input_path, **convert_kwargs)
    log(f"  -> {len(images)} page(s)")

    writer = PdfWriter()
    for i, image in enumerate(images, start=1):
        log(f"  OCR page {i}/{len(images)}...")
        page_bytes = ocr_page_to_pdf_page(image)
        reader = PdfReader(io.BytesIO(page_bytes))
        writer.add_page(reader.pages[0])

    with open(output_path, "wb") as f:
        writer.write(f)
    log(f"  Saved -> {output_path}")


def process_files(file_paths, output_folder, log=print):
    """Process a list of PDF file paths, saving results into output_folder."""
    os.makedirs(output_folder, exist_ok=True)
    results = []
    for path in file_paths:
        base = os.path.splitext(os.path.basename(path))[0]
        out_path = os.path.join(output_folder, f"{base}_searchable.pdf")
        try:
            log(f"\nProcessing: {path}")
            make_searchable(path, out_path, log=log)
            results.append((path, out_path, None))
        except Exception as e:
            log(f"  ERROR: {e}")
            results.append((path, None, str(e)))
    return results


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

def check_dependencies():
    """Give a clear, friendly message if Tesseract or Poppler aren't set up,
    instead of a cryptic crash."""
    problems = []

    try:
        pytesseract.get_tesseract_version()
    except Exception:
        problems.append(
            "Tesseract OCR engine was not found.\n"
            "Install it (see the SETUP section at the top of this script),\n"
            "and if on Windows, set TESSERACT_PATH near the top of this file."
        )

    # Poppler is only checked implicitly when converting a PDF, so we do a
    # lightweight probe by trying to import and call pdfinfo via pdf2image.
    try:
        from pdf2image.pdf2image import pdfinfo_from_path  # noqa
    except Exception:
        pass  # not critical; real check happens on first conversion

    return problems


def pick_mode():
    """Small window: choose between picking files or a whole folder."""
    choice = {"mode": None}

    root = tk.Tk()
    root.title("Make PDF Searchable")
    root.geometry("360x160")
    root.resizable(False, False)

    tk.Label(root, text="What would you like to process?", font=("Segoe UI", 11)).pack(pady=15)

    def choose_files():
        choice["mode"] = "files"
        root.destroy()

    def choose_folder():
        choice["mode"] = "folder"
        root.destroy()

    tk.Button(root, text="Pick File(s)", width=20, command=choose_files).pack(pady=5)
    tk.Button(root, text="Pick Folder (all PDFs inside)", width=25, command=choose_folder).pack(pady=5)

    root.mainloop()
    return choice["mode"]


def run_gui():
    problems = check_dependencies()
    if problems:
        # Still let the user proceed, but warn clearly.
        warning_root = tk.Tk()
        warning_root.withdraw()
        messagebox.showwarning("Setup check", "\n\n".join(problems))
        warning_root.destroy()

    mode = pick_mode()
    if mode is None:
        return  # user closed the window

    root = tk.Tk()
    root.withdraw()  # hide the empty main window; we only need dialogs

    if mode == "files":
        file_paths = filedialog.askopenfilenames(
            title="Select PDF file(s)",
            filetypes=[("PDF files", "*.pdf")],
        )
        file_paths = list(file_paths)
    else:
        folder = filedialog.askdirectory(title="Select a folder containing PDFs")
        if not folder:
            file_paths = []
        else:
            file_paths = [
                os.path.join(folder, f)
                for f in sorted(os.listdir(folder))
                if f.lower().endswith(".pdf")
            ]

    if not file_paths:
        messagebox.showinfo("Nothing selected", "No PDF files were selected. Exiting.")
        return

    output_folder = filedialog.askdirectory(title="Choose an OUTPUT folder for the searchable PDFs")
    if not output_folder:
        messagebox.showinfo("Cancelled", "No output folder chosen. Exiting.")
        return

    print(f"\n{len(file_paths)} file(s) selected. Output folder: {output_folder}\n")

    results = process_files(file_paths, output_folder, log=print)

    succeeded = [r for r in results if r[2] is None]
    failed = [r for r in results if r[2] is not None]

    summary = f"Done!\n\nSucceeded: {len(succeeded)}\nFailed: {len(failed)}"
    if failed:
        summary += "\n\nFailed files:\n" + "\n".join(
            f"- {os.path.basename(p)}: {err}" for p, _, err in failed
        )

    messagebox.showinfo("Finished", summary)
    print("\n" + summary)


if __name__ == "__main__":
    try:
        run_gui()
    except Exception:
        # Fall back to printing the full traceback so setup issues are easy to diagnose
        traceback.print_exc()
        input("\nAn error occurred. Press Enter to close...")
