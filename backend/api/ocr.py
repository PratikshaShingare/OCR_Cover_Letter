"""
Khanna Travels & Holidays — OCR API Endpoints
Handles passport upload and data extraction.
"""

from fastapi import APIRouter, UploadFile, File, HTTPException
from ..ocr.ocr_service import get_ocr_service, ExtractedPassportData

router = APIRouter(prefix="/api/ocr", tags=["OCR"])


@router.post("/upload-passport", response_model=ExtractedPassportData)
async def upload_passport(file: UploadFile = File(...)):
    """
    Uploads a passport copy (PDF, JPG, PNG) and extracts structured data.
    Never uses fake or sample client records.
    """
    valid_extensions = (".jpg", ".jpeg", ".png", ".pdf")
    filename = file.filename or "passport.pdf"
    if not any(filename.lower().endswith(ext) for ext in valid_extensions):
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported file format '{filename}'. Allowed formats: PDF, JPG, JPEG, PNG."
        )
        
    try:
        content = await file.read()
        ocr_service = get_ocr_service()
        extracted_data = await ocr_service.extract(content, filename)
        return extracted_data
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"OCR Extraction failed: {str(e)}")
