"""
Khanna Travels and Holidays - Passport Document Normalization Pipeline
Provides intelligent preprocessing, orientation correction, EXIF transpose,
margin trimming, and separate front/back page handling for passport PDFs and images.
"""

import os
import gc
from typing import Dict, Any, List, Tuple, Optional
from PIL import Image, ImageOps, ImageChops


def auto_trim_margins(img: Image.Image, padding: int = 15) -> Image.Image:
    """
    Detects and crops excessive white or black scanner margins,
    retaining a clean, readable passport frame with a safe border.
    """
    try:
        gray = img.convert("L")
        bg = Image.new(gray.mode, gray.size, gray.getpixel((0, 0)))
        diff = ImageChops.difference(gray, bg)
        diff = ImageChops.add(diff, diff, 2.0, -100)
        bbox = diff.getbbox()
        if bbox:
            w, h = img.size
            x0, y0, x1, y1 = bbox
            if (x0 > w * 0.04 or y0 > h * 0.04 or x1 < w * 0.96 or y1 < h * 0.96):
                crop_w = x1 - x0
                crop_h = y1 - y0
                if crop_w > w * 0.4 and crop_h > h * 0.4:
                    pad_x0 = max(0, x0 - padding)
                    pad_y0 = max(0, y0 - padding)
                    pad_x1 = min(w, x1 + padding)
                    pad_y1 = min(h, y1 + padding)
                    return img.crop((pad_x0, pad_y0, pad_x1, pad_y1))
    except Exception as e:
        print(f"[Normalizer] Auto-trim notice: {e}")
    return img


def normalize_passport_image(img: Image.Image, max_dim: int = 1800) -> Image.Image:
    """
    Normalizes a PIL image:
    1. Corrects smartphone EXIF orientation tag.
    2. Resizes large images (> 1800px) using high-quality Lanczos resampling.
    3. Crops excessive scanning margins.
    """
    try:
        img = ImageOps.exif_transpose(img)
    except Exception as e:
        print(f"[Normalizer] EXIF transpose notice: {e}")

    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")

    w, h = img.size
    if max(w, h) > max_dim:
        scale = max_dim / float(max(w, h))
        new_w = int(w * scale)
        new_h = int(h * scale)
        img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)

    img = auto_trim_margins(img)
    return img


def normalize_image_buffer(raw_bytes: bytes, max_dim: int = 1800) -> Tuple[bytes, str]:
    """
    Takes raw image bytes, applies EXIF orientation, size normalization,
    and border margin trimming, returning (normalized_bytes, mime_type).
    """
    import io
    img = Image.open(io.BytesIO(raw_bytes))
    normalized = normalize_passport_image(img, max_dim=max_dim)
    
    out_buf = io.BytesIO()
    normalized.save(out_buf, format="JPEG", quality=90, optimize=True)
    return out_buf.getvalue(), "image/jpeg"


def normalize_pdf_document(pdf_path: str, output_dir: str, max_pages: int = 4) -> Dict[str, Any]:
    """
    Renders and normalizes PDF pages using pypdfium2:
    - Page 1: Passport Front (Biographical)
    - Page 2: Passport Back (Address and Family)
    - Saves preview_page_{i}.png for each page.
    - Memory-safe: scale=1.5 and explicit garbage collection.
    """
    import pypdfium2 as pdfium

    os.makedirs(output_dir, exist_ok=True)
    doc = pdfium.PdfDocument(pdf_path)
    total_pages = len(doc)
    pages_meta: List[Dict[str, Any]] = []

    pages_to_render = min(total_pages, max_pages)
    for i in range(pages_to_render):
        page_num = i + 1
        if page_num == 1:
            label = "Passport Front (Biographical)"
        elif page_num == 2:
            label = "Passport Back (Address and Family)"
        else:
            label = f"Page {page_num}"

        raw_page_img = doc[i].render(scale=1.5).to_pil()
        normalized_img = normalize_passport_image(raw_page_img)

        out_path = os.path.join(output_dir, f"preview_page_{page_num}.png")
        normalized_img.save(out_path, "PNG", optimize=True)

        pages_meta.append({
            "page": page_num,
            "label": label,
            "width": normalized_img.width,
            "height": normalized_img.height,
            "url": f"preview_page_{page_num}.png"
        })

    doc.close()
    gc.collect()

    return {
        "totalPages": total_pages,
        "renderedPages": pages_to_render,
        "pages": pages_meta
    }
