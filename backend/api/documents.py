"""
Khanna Travels & Holidays — Document Generation API
Generates, previews, and exports Cover Letters as DOCX and PDF.
"""

from fastapi import APIRouter, HTTPException, Body
from fastapi.responses import FileResponse
import tempfile
import os
from typing import Dict, Any, Optional

from ..models.master_schema import MasterApplicationData
from ..document_generator.cover_letter_engine import CoverLetterEngine
from ..document_generator.docx_generator import build_docx_cover_letter
from ..pdf_generator.pdf_builder import build_pdf_cover_letter
from ..services.storage import save_application

router = APIRouter(prefix="/api/documents", tags=["Documents"])


@router.post("/generate-cover-letter")
async def generate_cover_letter(app_data: MasterApplicationData):
    """
    Generates country-specific cover letter from Master Client Data
    using Master Europe Base + Country Overrides.
    """
    try:
        engine = CoverLetterEngine()
        doc_result = engine.generate_document(app_data)
        
        # Cache generated HTML in app_data and save to DB
        app_data.generatedDocumentHtml = doc_result.get("html")
        app_data.isDocumentManuallyEdited = False
        save_application(app_data)
        
        return {
            "success": True,
            "document": doc_result,
            "applicationId": app_data.applicationId
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Document generation failed: {str(e)}")


@router.post("/download-docx")
async def download_docx(payload: Dict[str, Any] = Body(...)):
    """Generates and downloads genuine Microsoft Word .docx file."""
    try:
        # Determine if payload is MasterApplicationData or document structure
        if "personal" in payload and "passport" in payload:
            app_data = MasterApplicationData.model_validate(payload)
            engine = CoverLetterEngine()
            doc_data = engine.generate_document(app_data)
            applicant_name = app_data.personal.fullName or "Applicant"
            country = app_data.selectedCountry or "Visa"
        elif "document" in payload:
            doc_data = payload["document"]
            applicant_name = payload.get("applicantName", "Applicant")
            country = payload.get("country", "Visa")
        else:
            doc_data = payload
            applicant_name = "Applicant"
            country = "Visa"
            
        temp_dir = tempfile.mkdtemp()
        safe_name = "".join(c for c in applicant_name if c.isalnum() or c in (' ', '-', '_')).strip()
        filename = f"Cover_Letter_{country}_{safe_name.replace(' ', '_')}.docx"
        file_path = os.path.join(temp_dir, filename)
        
        build_docx_cover_letter(doc_data, file_path, app_data=app_data)
        
        return FileResponse(
            path=file_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate DOCX: {str(e)}")


@router.post("/download-pdf")
async def download_pdf(payload: Dict[str, Any] = Body(...)):
    """Generates and downloads vector-faithful PDF file."""
    try:
        if "personal" in payload and "passport" in payload:
            app_data = MasterApplicationData.model_validate(payload)
            engine = CoverLetterEngine()
            doc_data = engine.generate_document(app_data)
            applicant_name = app_data.personal.fullName or "Applicant"
            country = app_data.selectedCountry or "Visa"
        elif "document" in payload:
            doc_data = payload["document"]
            applicant_name = payload.get("applicantName", "Applicant")
            country = payload.get("country", "Visa")
        else:
            doc_data = payload
            applicant_name = "Applicant"
            country = "Visa"
            
        temp_dir = tempfile.mkdtemp()
        safe_name = "".join(c for c in applicant_name if c.isalnum() or c in (' ', '-', '_')).strip()
        filename = f"Cover_Letter_{country}_{safe_name.replace(' ', '_')}.pdf"
        file_path = os.path.join(temp_dir, filename)
        
        build_pdf_cover_letter(doc_data, file_path)
        
        return FileResponse(
            path=file_path,
            filename=filename,
            media_type="application/pdf"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate PDF: {str(e)}")
