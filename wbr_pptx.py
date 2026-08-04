"""
wbr_pptx.py — WBR PowerPoint export.

Strategy: render the generated PDF page to a high-resolution PNG via PyMuPDF,
then embed it as a full-slide image in a PPTX file.  This guarantees the PPTX
is a pixel-faithful match to the PDF — same fonts, same chart lines, same table
formatting.  The slide dimensions mirror the PDF (US Letter landscape, 11" × 8.5").

Usage:
    from wbr_pptx import pdf_to_pptx
    pptx_bytes = pdf_to_pptx(pdf_bytes, dpi=200)
"""
from __future__ import annotations
import io
from typing import Optional


def pdf_to_pptx(pdf_bytes: bytes, dpi: int = 200) -> bytes:
    """
    Convert the first page of a PDF to a PPTX slide.

    Parameters
    ----------
    pdf_bytes : raw PDF bytes (output of generate_standard_wbr)
    dpi       : render resolution — 200 gives crisp display; 300 for print quality

    Returns
    -------
    Raw .pptx bytes ready for st.download_button
    """
    import fitz  # PyMuPDF — already in requirements
    from pptx import Presentation
    from pptx.util import Inches, Pt

    # ── 1. Render PDF page to PNG ─────────────────────────────────────────────
    scale  = dpi / 72.0            # PDF native unit = 72 dpi
    matrix = fitz.Matrix(scale, scale)

    doc  = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc[0]
    pix  = page.get_pixmap(matrix=matrix, alpha=False)
    png_bytes = io.BytesIO(pix.tobytes("png"))
    doc.close()

    # ── 2. Create PPTX — same physical dimensions as the PDF page ────────────
    # PDF is landscape letter: 792 × 612 pts  =  11" × 8.5"
    prs = Presentation()
    prs.slide_width  = Inches(11.0)
    prs.slide_height = Inches(8.5)

    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

    # Full-slide image — fills the entire slide with no margins
    slide.shapes.add_picture(png_bytes, 0, 0, prs.slide_width, prs.slide_height)

    # ── 3. Return bytes ───────────────────────────────────────────────────────
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()
