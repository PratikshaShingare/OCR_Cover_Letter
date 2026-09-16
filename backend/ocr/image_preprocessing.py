"""
Khanna Travels & Holidays — Passport Image Preprocessing Engine
Renders high-resolution PDF pages (300 DPI equivalent) and optimizes
image regions (contrast enhancement, sharpening, MRZ cropping, rotation checks)
for optical character recognition.
"""

import io
from typing import List, Tuple, Optional
from PIL import Image, ImageEnhance, ImageFilter


def render_pdf_to_images(pdf_bytes: bytes, scale: float = 2.0, max_pages: int = 4) -> List[Image.Image]:
    """
    Renders PDF pages to high-resolution PIL images using pypdfium2.
    scale=2.0 produces crisp ~150-200 DPI, ensuring sharp text and clear MRZ glyphs with fast processing.
    For booklets with > 3 pages, renders the first two pages and the final page (family & address).
    """

    images = []
    try:
        import pypdfium2 as pdfium
        doc = pdfium.PdfDocument(io.BytesIO(pdf_bytes))
        total = len(doc)
        if total <= 3:
            indices = list(range(total))
        else:
            indices = [0, 1, total - 1]
        for i in indices:
            img = doc[i].render(scale=scale).to_pil()
            images.append(img)
    except Exception as e:
        print(f"pypdfium2 rendering notice: {e}")
    return images


def load_image_bytes(image_bytes: bytes) -> Optional[Image.Image]:
    """Loads raw image bytes into a PIL Image."""
    try:
        img = Image.open(io.BytesIO(image_bytes))
        return img.convert("RGB")
    except Exception as e:
        print(f"PIL load image notice: {e}")
        return None


def preprocess_mrz_crop(image: Image.Image) -> Image.Image:
    """
    Crops the bottom 28% of a passport biographical page (where MRZ is situated)
    and enhances contrast and edge sharpness for optimal OCR accuracy.
    """
    w, h = image.size
    # MRZ occupies the bottom 25-30% of standard ICAO Doc 9303 passports
    top_y = int(h * 0.70)
    crop = image.crop((0, top_y, w, h))

    # Convert to grayscale for thresholding & noise reduction
    gray = crop.convert("L")

    # Boost contrast (1.8x)
    enhancer = ImageEnhance.Contrast(gray)
    contrasted = enhancer.enhance(1.8)

    # Sharpen text edges
    sharpened = contrasted.filter(ImageFilter.SHARPEN)

    return sharpened


def preprocess_full_page(image: Image.Image) -> Image.Image:
    """
    Applies gentle contrast enhancement and sharpening to full passport page
    without distorting visual layout.
    """
    enhancer = ImageEnhance.Contrast(image)
    contrasted = enhancer.enhance(1.3)
    return contrasted.filter(ImageFilter.SHARPEN)


def get_rotation_variants(image: Image.Image) -> List[Tuple[int, Image.Image]]:
    """
    Returns image in 0°, 90°, and 270° orientations to handle sideways phone uploads.
    """
    variants = [(0, image)]
    # If height > width (portrait booklet), test horizontal orientation (landscape)
    w, h = image.size
    if h > w:
        variants.append((90, image.rotate(90, expand=True)))
        variants.append((270, image.rotate(270, expand=True)))
    return variants
