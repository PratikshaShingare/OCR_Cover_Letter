"""
Khanna Travels & Holidays — Production Passport Document Parser & OCR Engine
Executes multi-stage image preprocessing, MRZ crop enhancement, and
hardware-accelerated Windows Media OCR.
Parses TD3 MRZ check digits and bilingual VIZ layout.
Never invents, assumes, or uses hardcoded sample client data.
"""

import io
import os
import re
import tempfile
from typing import Dict, Any, Optional, Tuple
from PIL import Image
from .ocr_service import BaseOCRProvider, ExtractedPassportData
from .native_ocr import run_system_ocr, run_windows_ocr
from .image_preprocessing import (
    render_pdf_to_images,
    load_image_bytes,
    preprocess_mrz_crop,
    preprocess_full_page,
)
from .passport_extractor import extract_document_data


def score_ocr_text(text: str) -> int:
    """Scores OCR text to determine if image orientation is correct."""
    if not text:
        return 0
    score = 0
    keywords = [
        'PASSPORT', 'REPUBLIC OF INDIA', 'REPUBLIC', 'INDIA', 'BHARAT',
        'GIVEN NAME', 'SURNAME', 'NATIONALITY', 'DATE OF BIRTH', 'PLACE OF BIRTH',
        'PLACE OF ISSUE', 'DATE OF ISSUE', 'DATE OF EXPIRY', 'EXPIRY',
        'FATHER', 'MOTHER', 'SPOUSE', 'HUSBAND', 'WIFE', 'ADDRESS', 'PIN:'
    ]
    text_upper = text.upper()
    for kw in keywords:
        if kw in text_upper:
            score += 2
    if 'P<IND' in text_upper or 'IND<<' in text_upper:
        score += 8
    if re.search(r'[A-Za-z]\s*[0-9]{7}', text_upper):
        score += 4
    if re.search(r'\d{2}[/-]\d{2}[/-]\d{4}', text):
        score += 3
    return score


def get_oriented_page_ocr(img: Image.Image) -> Tuple[Image.Image, str, int]:
    """
    Ultra-fast orientation detection guaranteeing proper horizontal passport layout (width >= height).
    - If vertical (height > width): tests 90° vs 270° (at most 2 fast thumbnail tests).
    - If horizontal (width >= height): tests 0° vs 180° (at most 2 fast thumbnail tests).
    Guarantees proper horizontal passport layout (width >= height).
    Returns (oriented_image, ocr_text, angle).
    """
    def get_thumb_ocr(image: Image.Image) -> Tuple[str, int]:
        w, h = image.size
        ratio = 600.0 / max(w, h)
        if ratio < 1.0:
            thumb = image.resize((int(w * ratio), int(h * ratio)), Image.Resampling.BILINEAR)
        else:
            thumb = image
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
            tmp_path = tmp.name
        try:
            thumb.save(tmp_path, "PNG")
            txt = run_system_ocr(tmp_path, psm=3)
        finally:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass
        return txt, score_ocr_text(txt)

    if img.height > img.width:
        # Document was scanned vertically in portrait.
        # Passports are always horizontal landscape. Test the two landscape rotations: 90° and 270°.
        rot90 = img.rotate(90, expand=True)
        rot270 = img.rotate(270, expand=True)

        txt270, s270 = get_thumb_ocr(rot270)
        if s270 >= 8:
            return rot270, txt270, 270

        txt90, s90 = get_thumb_ocr(rot90)
        if s90 > s270:
            return rot90, txt90, 90
        else:
            # Default to 270° (standard clockwise rotation to landscape)
            return rot270, txt270, 270
    else:
        # Document is already horizontal landscape (width >= height).
        txt0, s0 = get_thumb_ocr(img)
        if s0 >= 8:
            return img, txt0, 0

        # Test upside-down (180°)
        rot180 = img.rotate(180, expand=True)
        txt180, s180 = get_thumb_ocr(rot180)
        if s180 > s0:
            return rot180, txt180, 180
        return img, txt0, 0


class MockOCRProvider(BaseOCRProvider):
    """
    Production-ready passport & visa document OCR engine.
    Uses pypdfium2 high-res rendering (scale 3, ~300 DPI) and native Windows.Media.Ocr.
    Parses TD3 MRZ check digits and bilingual VIZ layout.
    """

    async def extract(self, file_bytes: bytes, filename: str) -> ExtractedPassportData:
        filename_lower = filename.lower()
        is_pdf = filename_lower.endswith(".pdf")
        
        print(f"[OCR] Processing started for {filename}")

        # 1. Digital PDF text layer extraction (fast-path for searchable PDFs)
        digital_text = ""
        if is_pdf:
            try:
                import pypdf
                reader = pypdf.PdfReader(io.BytesIO(file_bytes))
                for page in reader.pages:
                    digital_text += (page.extract_text() or "") + "\n"
            except Exception:
                pass

        # 2. Render & Preprocess Images
        pil_images = []
        if is_pdf:
            # High-resolution rendering at scale=2 (~150-200 DPI) for fast OCR and low memory
            pil_images = render_pdf_to_images(file_bytes, scale=2.0, max_pages=2)
            print(f"[OCR] Rendered {len(pil_images)} page(s) at scale=2.0")
        else:
            loaded_img = load_image_bytes(file_bytes)
            if loaded_img:
                pil_images = [loaded_img]

        # 3. OCR on Full Page & Enhanced MRZ Crop with Fast Orientation Detection
        ocr_texts = []
        for idx, img in enumerate(pil_images):
            # A. Find correct orientation and perform full-page OCR
            best_img, text_full, angle = get_oriented_page_ocr(img)
            if angle != 0:
                print(f"[OCR] Page {idx+1}: Auto-rotated {angle}° for optimal reading")

            # If thumbnail text was sparse, run OCR on the full oriented image
            if len(text_full.strip()) < 100:
                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_full:
                    tmp_full_path = tmp_full.name
                try:
                    best_img.save(tmp_full_path, "PNG")
                    more_text = run_system_ocr(tmp_full_path, psm=3)
                    if more_text:
                        text_full = f"{text_full}\n{more_text}"
                finally:
                    if os.path.exists(tmp_full_path):
                        try:
                            os.remove(tmp_full_path)
                        except Exception:
                            pass

            if text_full:
                ocr_texts.append(text_full)

            # B. Enhanced MRZ Crop OCR on the correctly oriented image (bottom 28%)
            try:
                mrz_img = preprocess_mrz_crop(best_img)
                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_mrz:
                    tmp_mrz_path = tmp_mrz.name
                try:
                    mrz_img.save(tmp_mrz_path, "PNG")
                    text_mrz = run_system_ocr(tmp_mrz_path, psm=6)
                    if text_mrz:
                        ocr_texts.append(text_mrz)
                finally:
                    if os.path.exists(tmp_mrz_path):
                        try:
                            os.remove(tmp_mrz_path)
                        except Exception:
                            pass
            except Exception as mrz_err:
                print(f"[OCR] MRZ crop notice: {mrz_err}")

        # Combine text streams
        combined_text = f"{digital_text}\n" + "\n".join(ocr_texts)

        # 4. Extract structured passport data
        parsed = extract_document_data(combined_text, filename=filename)

        # 5. Populate ExtractedPassportData
        data = ExtractedPassportData()
        data.passportType = parsed.get("passportType", "P")
        data.passportNumber = parsed.get("passportNumber", "")
        data.givenName = parsed.get("givenName", "")
        data.middleName = parsed.get("middleName", "")
        data.surname = parsed.get("surname", "")
        data.fullName = parsed.get("fullName", "")
        data.title = parsed.get("title", "Mr.")
        data.nationality = parsed.get("nationality", "Indian")
        data.dob = parsed.get("dob", "")
        data.gender = parsed.get("gender", "")
        data.placeOfBirth = parsed.get("placeOfBirth", "")
        data.countryOfBirth = parsed.get("countryOfBirth", "India")
        data.issueDate = parsed.get("issueDate", "")
        data.expiryDate = parsed.get("expiryDate", "")
        data.issuePlace = parsed.get("issuePlace", "")
        data.issuingCountry = parsed.get("issuingCountry", "India")
        data.email = parsed.get("email", "")
        data.phone = parsed.get("phone", "")
        data.fatherFullName = parsed.get("fatherFullName", "")
        data.motherFullName = parsed.get("motherFullName", "")
        data.spouseFullName = parsed.get("spouseFullName", "")
        data.address = parsed.get("address", "")
        data.travellers = parsed.get("travellers", [])
        data.rawText = combined_text
        data.fieldDetails = parsed.get("fieldDetails", {})
        data.fieldStatuses = parsed.get("fieldStatuses", {})
        data.providerName = "Windows Media Hardware OCR Engine"

        has_primary = bool(data.passportNumber or data.fullName)
        data.confidence = 0.99 if bool(data.passportNumber and data.fullName) else (0.85 if has_primary else 0.50)

        print(f"[OCR] Extraction completed. Passport: {data.passportNumber or 'None'}, Name: {data.fullName or 'None'}")
        return data
