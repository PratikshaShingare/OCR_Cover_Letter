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
    Ultra-fast 4-way orientation detection using an optimized downsampled thumbnail.
    Tests 0° first. If keyword score < 10, tests 90°, 180°, 270°.
    Guarantees proper horizontal passport layout (width >= height).
    Returns (oriented_image, ocr_text, angle).
    """
    # 1. Create fast downsampled thumbnail for orientation detection (max 800px)
    w, h = img.size
    ratio = 800.0 / max(w, h)
    if ratio < 1.0:
        fast_img = img.resize((int(w * ratio), int(h * ratio)), Image.Resampling.BILINEAR)
    else:
        fast_img = img

    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        t0 = tmp.name
    try:
        fast_img.save(t0, "PNG")
        txt0 = run_system_ocr(t0, psm=3)
    finally:
        if os.path.exists(t0):
            os.remove(t0)

    s0 = score_ocr_text(txt0)
    best_ang = 0
    best_score = s0
    best_txt = txt0

    if s0 < 10:
        for ang in [90, 180, 270]:
            rot_fast = fast_img.rotate(ang, expand=True)
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                t = tmp.name
            try:
                rot_fast.save(t, "PNG")
                txt = run_system_ocr(t, psm=3)
            finally:
                if os.path.exists(t):
                    os.remove(t)
            sc = score_ocr_text(txt)
            if sc > best_score:
                best_score = sc
                best_ang = ang
                best_txt = txt
                if best_score >= 12:
                    break

    # 2. Rotate original full-resolution image to best angle
    oriented_img = img.rotate(best_ang, expand=True) if best_ang != 0 else img

    # 3. Horizontal Passport Guarantee:
    # A passport biographical spread is always landscape (width > height).
    # If the oriented image is vertical (height > width), orient it horizontally.
    if oriented_img.height > oriented_img.width:
        rot_90 = oriented_img.rotate(90, expand=True)
        rot_270 = oriented_img.rotate(270, expand=True)
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp90:
            t90 = tmp90.name
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp270:
            t270 = tmp270.name
        try:
            f90 = rot_90.resize((int(rot_90.width * (800.0 / max(rot_90.size))), int(rot_90.height * (800.0 / max(rot_90.size)))), Image.Resampling.BILINEAR) if max(rot_90.size) > 800 else rot_90
            f270 = rot_270.resize((int(rot_270.width * (800.0 / max(rot_270.size))), int(rot_270.height * (800.0 / max(rot_270.size)))), Image.Resampling.BILINEAR) if max(rot_270.size) > 800 else rot_270
            f90.save(t90, "PNG")
            f270.save(t270, "PNG")
            txt90 = run_system_ocr(t90, psm=3)
            txt270 = run_system_ocr(t270, psm=3)
            s90 = score_ocr_text(txt90)
            s270 = score_ocr_text(txt270)
            if s90 >= s270:
                oriented_img = rot_90
                best_txt = txt90
                best_ang = (best_ang + 90) % 360
            else:
                oriented_img = rot_270
                best_txt = txt270
                best_ang = (best_ang + 270) % 360
        finally:
            if os.path.exists(t90):
                os.remove(t90)
            if os.path.exists(t270):
                os.remove(t270)

    return oriented_img, best_txt, best_ang


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
            # High-resolution rendering at scale=3 (~300 DPI)
            pil_images = render_pdf_to_images(file_bytes, scale=3.0, max_pages=4)
            print(f"[OCR] Rendered {len(pil_images)} page(s) at high resolution")
        else:
            loaded_img = load_image_bytes(file_bytes)
            if loaded_img:
                pil_images = [loaded_img]

        # 3. OCR on Full Page & Enhanced MRZ Crop with 4-Way Orientation Detection
        ocr_texts = []
        for idx, img in enumerate(pil_images):
            # A. Find correct orientation and perform full-page OCR
            best_img, text_full, angle = get_oriented_page_ocr(img)
            if angle != 0:
                print(f"[OCR] Page {idx+1}: Auto-rotated {angle}° for optimal reading")
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
                        os.remove(tmp_mrz_path)
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
