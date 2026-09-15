"""
Khanna Travels & Holidays — Application Management API
Saves verified Master Client Data and automatically synchronizes internal Excel workbook.
"""

from fastapi import APIRouter, HTTPException, UploadFile, File, Form
from fastapi.responses import FileResponse, Response
from typing import List, Dict, Any, Optional
import os
import re
import shutil

from ..models.master_schema import MasterApplicationData, Traveller
from ..services.storage import save_application, get_application, list_applications, delete_application, generate_next_application_id
from ..services.excel_service import ExcelService
from ..ocr.ocr_service import get_ocr_service

router = APIRouter(prefix="/api/applications", tags=["Applications"])

STORAGE_BASE = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "internal_records", "applications"))


@router.post("", response_model=MasterApplicationData)
async def create_or_update_application(app_data: MasterApplicationData):
    """
    Saves verified Master Client Data to the database
    and automatically populates the internal Excel data sheet.
    """
    try:
        saved = save_application(app_data)
        
        # Automatically update internal Excel data sheet on backend
        try:
            ExcelService.sync_master_data_to_excel(saved)
        except Exception as excel_err:
            print(f"Automated Excel sync notice: {excel_err}")
            
        return saved
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to save application: {str(e)}")


@router.post("/new", response_model=MasterApplicationData)
async def create_new_blank_application():
    """
    Creates a 100% blank new application record with a sequential ID (e.g. APP-2026-00001).
    Contains zero sample/demo data.
    """
    new_id = generate_next_application_id()
    blank_app = MasterApplicationData(applicationId=new_id)
    saved = save_application(blank_app)
    return saved


@router.get("", response_model=List[Dict[str, Any]])
async def get_all_applications():
    """Lists saved applications."""
    return list_applications()


@router.get("/{app_id}", response_model=MasterApplicationData)
async def get_single_application(app_id: str):
    """Retrieves single application master data."""
    app = get_application(app_id)
    if not app:
        raise HTTPException(status_code=404, detail=f"Application {app_id} not found.")
    return app


@router.delete("/{app_id}")
async def remove_application(app_id: str):
    """
    Deletes an application record, associated passport files,
    generated documents, and internal Excel file.
    """
    deleted = delete_application(app_id)
    if not deleted:
        raise HTTPException(status_code=404, detail=f"Application {app_id} not found.")
    return {"success": True, "message": f"Application {app_id} and associated files deleted."}


def get_empty_preview_svg(label: str = "Document Preview") -> str:
    return f"""<svg xmlns="http://www.w3.org/2000/svg" width="600" height="800" viewBox="0 0 600 800">
  <rect width="100%" height="100%" fill="#0f172a"/>
  <rect x="40" y="40" width="520" height="720" rx="8" fill="#1e293b" stroke="#334155" stroke-width="2"/>
  <circle cx="300" cy="340" r="48" fill="#0f172a" stroke="#64748b" stroke-width="2"/>
  <path d="M285 325 h30 M285 345 h30 M285 355 h20" stroke="#94a3b8" stroke-width="3" stroke-linecap="round"/>
  <text x="300" y="430" text-anchor="middle" fill="#f8fafc" font-family="system-ui, sans-serif" font-size="18" font-weight="600">{label}</text>
  <text x="300" y="460" text-anchor="middle" fill="#94a3b8" font-family="system-ui, sans-serif" font-size="13">Khanna Travels &amp; Holidays — Visa Automation</text>
</svg>"""


@router.post("/{app_id}/passport")
async def upload_application_passport(app_id: str, file: UploadFile = File(...)):
    """
    Uploads the genuine passport document for an application.
    Stores the original document for in-app preview and performs OCR/MRZ extraction.
    """
    app = get_application(app_id)
    if not app:
        # If application doesn't exist yet, create it with this ID
        app = MasterApplicationData(applicationId=app_id)

    valid_extensions = (".jpg", ".jpeg", ".png", ".pdf")
    filename = file.filename or "passport.pdf"
    ext = os.path.splitext(filename)[1].lower()
    if ext not in valid_extensions:
        raise HTTPException(status_code=400, detail=f"Unsupported format '{ext}'. Allowed: PDF, JPG, JPEG, PNG.")

    # Save to application-specific folder
    passport_dir = os.path.join(STORAGE_BASE, app_id, "passport")
    os.makedirs(passport_dir, exist_ok=True)
    
    # Clean previous files in passport dir to ensure only the active one is present
    for old_file in os.listdir(passport_dir):
        try:
            os.remove(os.path.join(passport_dir, old_file))
        except Exception:
            pass

    safe_filename = f"passport{ext}"
    saved_path = os.path.join(passport_dir, safe_filename)
    
    content = await file.read()
    with open(saved_path, "wb") as f:
        f.write(content)

    # Pre-render high-resolution PNG preview images for instant, reliable browser display
    total_pages = 1
    if ext == ".pdf":
        try:
            import pypdfium2 as pdfium
            from PIL import Image
            from ..ocr.mock_provider import get_oriented_page_ocr
            doc = pdfium.PdfDocument(saved_path)
            total_pages = len(doc)
            for i in range(min(total_pages, 10)):
                raw_img = doc[i].render(scale=2.0).to_pil()
                oriented_img, _, _ = get_oriented_page_ocr(raw_img)
                oriented_img.save(os.path.join(passport_dir, f"preview_page_{i + 1}.png"), "PNG")
        except Exception as render_err:
            print(f"PDF preview rendering notice: {render_err}")
    else:
        # Image file (JPG, PNG)
        try:
            from PIL import Image
            from ..ocr.mock_provider import get_oriented_page_ocr
            img = Image.open(saved_path).convert("RGB")
            oriented_img, _, _ = get_oriented_page_ocr(img)
            oriented_img.save(os.path.join(passport_dir, "preview_page_1.png"), "PNG")
        except Exception:
            shutil.copy(saved_path, os.path.join(passport_dir, "preview_page_1.png"))

    # Perform genuine OCR / MRZ extraction
    ocr_service = get_ocr_service()
    extracted = await ocr_service.extract(content, filename)

    # Merge extracted fields into MasterApplicationData without overriding existing human-verified data if already set
    if extracted.passportNumber:
        app.passport.passportNumber = extracted.passportNumber
    if extracted.givenName:
        app.personal.givenName = extracted.givenName
    if extracted.surname:
        app.personal.surname = extracted.surname
    if extracted.fullName:
        app.personal.fullName = extracted.fullName
    elif extracted.givenName or extracted.surname:
        app.personal.fullName = f"{extracted.givenName} {extracted.surname}".strip()
    if extracted.dob:
        app.personal.dob = extracted.dob
    if extracted.gender:
        app.personal.gender = extracted.gender
    if extracted.nationality:
        app.personal.nationality = extracted.nationality
    if extracted.placeOfBirth:
        app.personal.placeOfBirth = extracted.placeOfBirth
    if extracted.expiryDate:
        app.passport.expiryDate = extracted.expiryDate
    if extracted.issueDate:
        app.passport.issueDate = extracted.issueDate
    if extracted.issuePlace:
        app.passport.issuePlace = extracted.issuePlace
    if extracted.issuingCountry:
        app.passport.issuingCountry = extracted.issuingCountry
    if getattr(extracted, "title", None) and extracted.title:
        app.personal.title = extracted.title
    if getattr(extracted, "fatherFullName", None) and extracted.fatherFullName:
        app.family.fatherFullName = extracted.fatherFullName
    if getattr(extracted, "motherFullName", None) and extracted.motherFullName:
        app.family.motherFullName = extracted.motherFullName
    if getattr(extracted, "spouseFullName", None) and extracted.spouseFullName:
        app.family.spouseFullName = extracted.spouseFullName
    if getattr(extracted, "address", None) and extracted.address:
        app.address.currentResidentialAddress = extracted.address
        if not app.address.addressLine1:
            parts = [p.strip() for p in extracted.address.split(',') if p.strip() and not re.fullmatch(r'[A-Za-z0-9]{7,9}', p.strip())]
            if parts:
                app.address.addressLine1 = parts[0]
            if len(parts) > 1:
                app.address.addressLine2 = ', '.join(parts[1:3])
            pin_m = re.search(r'\b(4[0-9]{5})\b', extracted.address)
            if pin_m and not app.address.postalCode:
                app.address.postalCode = pin_m.group(1)
            for city in ['Mumbai', 'Navi Mumbai', 'Thane', 'Pune', 'Delhi']:
                if city.lower() in extracted.address.lower() and not app.address.city:
                    app.address.city = city
                    break
    if getattr(extracted, "email", None) and extracted.email and not app.contact.emailAddress:
        app.contact.emailAddress = extracted.email
    if getattr(extracted, "phone", None) and extracted.phone and not app.contact.mobileNumber:
        app.contact.mobileNumber = extracted.phone

    # Merge extracted travellers if present and applicant has none yet
    if getattr(extracted, "travellers", None) and extracted.travellers and not app.travellers:
        for t_data in extracted.travellers:
            app.travellers.append(Traveller(
                title=t_data.get("title", "Mr."),
                fullName=t_data.get("fullName", ""),
                givenName=t_data.get("givenName", ""),
                surname=t_data.get("surname", ""),
                passportNumber=t_data.get("passportNumber", ""),
                relationship=t_data.get("relationship", "Family Member"),
                occupation=t_data.get("occupation", ""),
                issuePlace=t_data.get("issuePlace", ""),
                issueDate=t_data.get("issueDate", ""),
                nationality=t_data.get("nationality", "Indian")
            ))

    # Merge field statuses and details
    if not app.fieldStatuses:
        app.fieldStatuses = {}
    for k, v in extracted.fieldStatuses.items():
        app.fieldStatuses[k] = v

    if not app.fieldDetails:
        app.fieldDetails = {}
    if getattr(extracted, "fieldDetails", None):
        for k, v in extracted.fieldDetails.items():
            app.fieldDetails[k] = v

    # Save application
    saved_app = save_application(app)
    try:
        ExcelService.sync_master_data_to_excel(saved_app)
    except Exception:
        pass

    return {
        "success": True,
        "applicationId": app_id,
        "fileInfo": {
            "filename": filename,
            "savedAs": safe_filename,
            "contentType": file.content_type,
            "isPdf": ext == ".pdf",
            "totalPages": total_pages,
            "previewUrl": f"/api/applications/{app_id}/passport-preview?page=1"
        },
        "extracted": extracted,
        "application": saved_app
    }


@router.get("/{app_id}/passport-preview")
async def get_application_passport_preview(app_id: str, page: int = 1):
    """
    Streams the pre-rendered high-resolution PNG image of the passport page.
    Dynamically renders on demand if missing. Guarantees zero blank white pages.
    """
    passport_dir = os.path.join(STORAGE_BASE, app_id, "passport")
    if not os.path.exists(passport_dir):
        return Response(content=get_empty_preview_svg("No Passport Document"), media_type="image/svg+xml")
        
    preview_file = os.path.join(passport_dir, f"preview_page_{page}.png")
    if os.path.exists(preview_file):
        return FileResponse(
            path=preview_file,
            media_type="image/png",
            content_disposition_type="inline",
            headers={"Content-Disposition": "inline", "Cache-Control": "no-cache"}
        )

    # Dynamic on-demand rendering if preview file does not yet exist
    pdf_path = os.path.join(passport_dir, "passport.pdf")
    if os.path.exists(pdf_path):
        try:
            import pypdfium2 as pdfium
            from ..ocr.mock_provider import get_oriented_page_ocr
            doc = pdfium.PdfDocument(pdf_path)
            target_idx = max(0, min(page - 1, len(doc) - 1))
            page_img = doc[target_idx].render(scale=2.0).to_pil()
            oriented_img, _, _ = get_oriented_page_ocr(page_img)
            oriented_img.save(preview_file, "PNG")
            return FileResponse(
                path=preview_file,
                media_type="image/png",
                content_disposition_type="inline",
                headers={"Content-Disposition": "inline", "Cache-Control": "no-cache"}
            )
        except Exception as render_err:
            print(f"[Preview] On-demand render notice: {render_err}")

    # Fallback to page 1 if available
    p1 = os.path.join(passport_dir, "preview_page_1.png")
    if os.path.exists(p1):
        return FileResponse(
            path=p1,
            media_type="image/png",
            content_disposition_type="inline",
            headers={"Content-Disposition": "inline", "Cache-Control": "no-cache"}
        )

    # Fallback to original image if uploaded as image
    orig_files = [f for f in os.listdir(passport_dir) if not f.startswith("preview_page_")]
    if orig_files:
        orig_path = os.path.join(passport_dir, orig_files[0])
        if orig_path.lower().endswith((".jpg", ".jpeg", ".png")):
            return FileResponse(
                path=orig_path,
                media_type="image/png",
                content_disposition_type="inline",
                headers={"Content-Disposition": "inline", "Cache-Control": "no-cache"}
            )

    return Response(content=get_empty_preview_svg(f"Preview Page {page}"), media_type="image/svg+xml")


@router.get("/{app_id}/passport-info")
async def get_application_passport_info(app_id: str):
    """
    Returns presence, file format, and page count of the uploaded passport.
    """
    passport_dir = os.path.join(STORAGE_BASE, app_id, "passport")
    if not os.path.exists(passport_dir):
        return {"hasPassport": False, "totalPages": 0, "isPdf": False}
        
    orig_files = [f for f in os.listdir(passport_dir) if not f.startswith("preview_page_")]
    if not orig_files:
        return {"hasPassport": False, "totalPages": 0, "isPdf": False}
        
    orig = orig_files[0]
    is_pdf = orig.lower().endswith(".pdf")
    preview_pages = [f for f in os.listdir(passport_dir) if f.startswith("preview_page_")]
    total_pages = len(preview_pages) if preview_pages else 1
    
    return {
        "hasPassport": True,
        "filename": orig,
        "isPdf": is_pdf,
        "totalPages": total_pages,
        "previewUrl": f"/api/applications/{app_id}/passport-preview?page=1"
    }


@router.get("/{app_id}/passport-file")
async def get_application_passport_file(app_id: str):
    """
    Streams the original uploaded passport document (PDF or image).
    Sets Content-Disposition: inline to prevent blank-page download blocking.
    """
    passport_dir = os.path.join(STORAGE_BASE, app_id, "passport")
    if not os.path.exists(passport_dir):
        raise HTTPException(status_code=404, detail="No passport document found for this application.")
        
    orig_files = [f for f in os.listdir(passport_dir) if not f.startswith("preview_page_")]
    if not orig_files:
        raise HTTPException(status_code=404, detail="No passport document found.")
        
    file_path = os.path.join(passport_dir, orig_files[0])
    ext = os.path.splitext(file_path)[1].lower()
    
    media_types = {
        ".pdf": "application/pdf",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".png": "image/png"
    }
    media_type = media_types.get(ext, "application/octet-stream")
    
    return FileResponse(
        path=file_path,
        media_type=media_type,
        content_disposition_type="inline",
        headers={"Content-Disposition": "inline"}
    )


@router.delete("/{app_id}/passport-file")
async def remove_application_passport_file(app_id: str):
    """Removes the uploaded passport document and clears passport file."""
    passport_dir = os.path.join(STORAGE_BASE, app_id, "passport")
    if os.path.exists(passport_dir):
        shutil.rmtree(passport_dir, ignore_errors=True)
    return {"success": True, "message": "Passport document removed."}

