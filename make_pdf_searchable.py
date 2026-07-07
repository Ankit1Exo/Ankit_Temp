"""
Make a scanned PDF searchable
------------------------------
Takes a scanned (image-only) PDF and produces a new PDF that looks
identical but has an invisible, selectable/searchable text layer
placed exactly over the words Tesseract OCR detects.

How it works:
1. Each page of the input PDF is rendered to an image (pdf2image / poppler).
2. Tesseract OCR extracts each word's text + bounding box (pytesseract).
3. A new PDF page is built: the original page image on top, and an
   invisible text layer (rendering mode 3 = invisible) underneath it,
   with each word placed at its correct position and scaled to fit
   its bounding box so text selection lines up with the image.
4. Pages are combined into the final searchable PDF.

Usage:
    python make_pdf_searchable.py input_scanned.pdf output_searchable.pdf

Requirements (all already used here):
    pip install pytesseract pdf2image reportlab pypdf pillow
    Also requires the tesseract-ocr binary and poppler-utils (pdftoppm)
    to be installed on the system.
"""

import sys
import io
import os

import pytesseract
from pdf2image import convert_from_path
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from pypdf import PdfReader, PdfWriter

# Higher DPI = better OCR accuracy but slower processing and larger files
DPI = 300


def ocr_page_to_pdf_page(image: Image.Image) -> bytes:
    """
    Given a page image, run OCR and build a single-page PDF (as bytes)
    that has the image on top and an invisible text layer beneath it.
    """
    width_px, height_px = image.size

    # Point size in PDF units (72 points per inch)
    width_pt = width_px * 72.0 / DPI
    height_pt = height_px * 72.0 / DPI

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(width_pt, height_pt))

    # 1. Draw the original scanned image covering the full page
    c.drawImage(ImageReader(image), 0, 0, width=width_pt, height=height_pt)

    # 2. Run OCR to get word-level text + bounding boxes
    ocr_data = pytesseract.image_to_data(
        image, output_type=pytesseract.Output.DICT
    )

    text_object = c.beginText()
    text_object.setTextRenderMode(3)  # 3 = invisible text (still selectable/searchable)

    n_boxes = len(ocr_data["text"])
    for i in range(n_boxes):
        word = ocr_data["text"][i].strip()
        if not word:
            continue

        conf = int(float(ocr_data["conf"][i])) if ocr_data["conf"][i] not in ("-1", "") else -1
        if conf < 30:  # skip very low-confidence junk to keep the layer clean
            continue

        x_px = ocr_data["left"][i]
        y_px = ocr_data["top"][i]
        w_px = ocr_data["width"][i]
        h_px = ocr_data["height"][i]

        # Convert pixel coords (origin top-left) to PDF points (origin bottom-left)
        x_pt = x_px * 72.0 / DPI
        y_pt = height_pt - (y_px + h_px) * 72.0 / DPI
        box_width_pt = w_px * 72.0 / DPI
        box_height_pt = h_px * 72.0 / DPI

        if box_height_pt <= 0:
            continue

        font_size = box_height_pt * 0.9  # slightly smaller than box looks more natural
        c.setFont("Helvetica", font_size)

        # Scale horizontally so the invisible word spans the same width as the
        # detected box, keeping text-selection alignment close to the image.
        text_width = c.stringWidth(word, "Helvetica", font_size)
        if text_width > 0:
            h_scale = (box_width_pt / text_width) * 100.0
        else:
            h_scale = 100.0

        text_object.setTextOrigin(x_pt, y_pt)
        text_object.setHorizScale(h_scale)
        text_object.textOut(word)

    c.drawText(text_object)
    c.showPage()
    c.save()

    buffer.seek(0)
    return buffer.read()


def make_searchable(input_path: str, output_path: str):
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    print(f"Converting pages of '{input_path}' to images at {DPI} DPI...")
    images = convert_from_path(input_path, dpi=DPI)
    print(f"  -> {len(images)} page(s) found")

    writer = PdfWriter()

    for i, image in enumerate(images, start=1):
        print(f"OCR-ing page {i}/{len(images)}...")
        page_pdf_bytes = ocr_page_to_pdf_page(image)
        page_reader = PdfReader(io.BytesIO(page_pdf_bytes))
        writer.add_page(page_reader.pages[0])

    with open(output_path, "wb") as f:
        writer.write(f)

    print(f"\nDone! Searchable PDF saved to: {output_path}")


def main():
    if len(sys.argv) != 3:
        print("Usage: python make_pdf_searchable.py <input_scanned.pdf> <output_searchable.pdf>")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]
    make_searchable(input_path, output_path)


if __name__ == "__main__":
    main()
