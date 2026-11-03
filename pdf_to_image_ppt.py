"""Convert an image-heavy PDF into a PPTX — one slide per page, image as full-slide background.

Works best for PDFs that are essentially exported slide decks: each page has a
full-page image (16:9) with text overlaid. No AI, no text generation — just
extract the embedded images and place them on slides.
"""

from __future__ import annotations

import io
import sys
from pathlib import Path

from PIL import Image
from pypdf import PdfReader
from pptx import Presentation
from pptx.util import Emu, Inches, Pt


def extract_page_image(pdf_path: str | Path) -> list[bytes | None]:
    """Return a list of PNG/JPEG bytes, one per page (None if no usable image)."""
    reader = PdfReader(str(pdf_path))
    results: list[bytes | None] = []
    for page in reader.pages:
        try:
            images = list(page.images)
        except Exception:
            images = []
        if not images:
            results.append(None)
            continue
        # Pick the largest image on the page (the full-page background)
        best_raw = None
        best_area = 0
        for img in images:
            try:
                im = Image.open(io.BytesIO(img.data))
                area = im.size[0] * im.size[1]
                if area > best_area:
                    best_area = area
                    best_raw = img.data
            except Exception:
                continue
        if best_raw is None:
            results.append(None)
            continue
        # Re-encode to ensure PPTX compatibility
        prepared = _prepare_image(best_raw)
        results.append(prepared)
    return results


def _prepare_image(raw: bytes, max_dim: int = 2000) -> bytes | None:
    """Downscale if needed and return PNG/JPEG bytes."""
    try:
        im = Image.open(io.BytesIO(raw))
        im.load()
    except Exception:
        return None
    w, h = im.size
    if w < 100 or h < 100:
        return None
    # Downscale if very large (keeps PPTX file size reasonable)
    if w > max_dim or h > max_dim:
        im.thumbnail((max_dim, max_dim))
    buf = io.BytesIO()
    if im.mode in ("RGBA", "LA", "P"):
        im.convert("RGBA").save(buf, "PNG")
    else:
        im.convert("RGB").save(buf, "JPEG", quality=90)
    return buf.getvalue()


def _add_picture_full_slide(slide, data: bytes, slide_w: Emu, slide_h: Emu):
    """Place an image to cover the entire slide, preserving aspect ratio (centered)."""
    with Image.open(io.BytesIO(data)) as im:
        w, h = im.size
    # Scale to cover the slide while keeping aspect ratio
    scale_w = slide_w / w
    scale_h = slide_h / h
    scale = max(scale_w, scale_h)  # cover mode — fills the slide
    pw, ph = int(w * scale), int(h * scale)
    left = int((slide_w - pw) / 2)
    top = int((slide_h - ph) / 2)
    slide.shapes.add_picture(io.BytesIO(data), left, top, width=pw, height=ph)


def convert_pdf_to_pptx(
    pdf_path: str | Path,
    output_path: str | Path | None = None,
    title: str = "",
) -> Path:
    """Convert a PDF to a PPTX file. Returns the output path."""
    pdf_path = Path(pdf_path)
    if output_path is None:
        output_path = pdf_path.with_suffix(".pptx")
    output_path = Path(output_path)

    print(f"Reading PDF: {pdf_path.name}")
    page_images = extract_page_image(pdf_path)

    prs = Presentation()
    # 16:9 widescreen
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide_w = prs.slide_width
    slide_h = prs.slide_height

    total = len(page_images)
    added = 0
    for i, img_data in enumerate(page_images, start=1):
        if img_data is None:
            # Page has no extractable image — add a blank slide with page number
            slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
            txbox = slide.shapes.add_textbox(Inches(5), Inches(3), Inches(3), Inches(1))
            txbox.text_frame.text = f"Page {i}"
            print(f"  Page {i}/{total}: no image (blank slide)")
            continue

        slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
        _add_picture_full_slide(slide, img_data, slide_w, slide_h)
        added += 1
        print(f"  Page {i}/{total}: image added ({added} slides total)")

    prs.save(str(output_path))
    size_mb = output_path.stat().st_size / (1024 * 1024)
    print(f"\nDone! {added} slides written to {output_path.name} ({size_mb:.1f} MB)")
    return output_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        # Default to the PDF in the project directory
        pdf = Path(__file__).parent / "Environmental Studies – Units 1 to 4_2860.pdf"
        if not pdf.exists():
            print("Usage: python pdf_to_image_ppt.py <input.pdf> [output.pptx]")
            sys.exit(1)
    else:
        pdf = Path(sys.argv[1])

    out = sys.argv[2] if len(sys.argv) > 2 else None
    convert_pdf_to_pptx(pdf, out)
