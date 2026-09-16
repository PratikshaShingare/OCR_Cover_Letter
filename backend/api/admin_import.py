"""
Khanna Travels & Holidays — Admin Bulk Data Ingestion API
Supports bulk import of client applications from .xlsx, .xls, and .csv files
with automated column mapping detection, interactive field overrides,
duplicate detection/handling, and persistent audit history logging.
"""

import os
import csv
import io
import uuid
import json
from typing import Dict, Any, List, Optional
from datetime import datetime
from pydantic import BaseModel
from fastapi import APIRouter, HTTPException, UploadFile, File

from ..models.master_schema import MasterApplicationData
from ..services.storage import (
    save_application,
    list_applications,
    get_application,
    record_import_history,
    list_import_history,
    generate_next_application_id,
    get_connection
)
from ..services.excel_service import ExcelService

router = APIRouter(prefix="/api/admin/import", tags=["Admin Bulk Import"])

TEMP_IMPORT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "internal_records", "temp_imports"))
os.makedirs(TEMP_IMPORT_DIR, exist_ok=True)

TARGET_FIELDS = [
    {"key": "passport.passportNumber", "label": "Passport Number", "required": True},
    {"key": "personal.fullName", "label": "Full Name", "required": False},
    {"key": "personal.givenName", "label": "Given Name", "required": False},
    {"key": "personal.surname", "label": "Surname", "required": False},
    {"key": "personal.dob", "label": "Date of Birth (YYYY-MM-DD)", "required": False},
    {"key": "personal.gender", "label": "Gender", "required": False},
    {"key": "personal.nationality", "label": "Nationality", "required": False},
    {"key": "passport.issueDate", "label": "Passport Issue Date", "required": False},
    {"key": "passport.expiryDate", "label": "Passport Expiry Date", "required": False},
    {"key": "passport.issuePlace", "label": "Place of Issue", "required": False},
    {"key": "contact.mobileNumber", "label": "Mobile Phone Number", "required": False},
    {"key": "contact.emailAddress", "label": "Email Address", "required": False},
    {"key": "address.addressLine1", "label": "Address Line 1", "required": False},
    {"key": "address.city", "label": "City", "required": False},
    {"key": "address.state", "label": "State", "required": False},
    {"key": "address.postalCode", "label": "Postal Code / PIN", "required": False},
    {"key": "family.fatherFullName", "label": "Father's Name", "required": False},
    {"key": "family.spouseFullName", "label": "Spouse's Name", "required": False},
    {"key": "selectedCountry", "label": "Destination Country", "required": False},
    {"key": "travel.visaType", "label": "Visa Type", "required": False},
    {"key": "employment.employerName", "label": "Employer Name", "required": False},
    {"key": "employment.jobTitle", "label": "Job Title / Role", "required": False}
]

COLUMN_HEURISTICS = {
    "passport.passportNumber": ["passport", "passport number", "passport no", "passport_no", "passportnum", "passport_number"],
    "personal.fullName": ["name", "full name", "applicant name", "passenger name", "client name", "traveller name"],
    "personal.givenName": ["first name", "given name", "first_name", "fname"],
    "personal.surname": ["last name", "surname", "last_name", "lname"],
    "personal.dob": ["dob", "date of birth", "birth date", "birth_date", "birthdate"],
    "personal.gender": ["gender", "sex"],
    "personal.nationality": ["nationality", "citizenship", "country of citizenship"],
    "passport.issueDate": ["issue date", "passport issue date", "date of issue", "issued_on"],
    "passport.expiryDate": ["expiry date", "passport expiry date", "date of expiry", "valid_until", "expires_on"],
    "passport.issuePlace": ["place of issue", "issue place", "issued at"],
    "contact.mobileNumber": ["mobile", "phone", "contact", "mobile number", "phone number", "cell"],
    "contact.emailAddress": ["email", "email address", "email id", "email_id"],
    "address.addressLine1": ["address", "residential address", "address line 1", "street"],
    "address.city": ["city", "town"],
    "address.state": ["state", "province"],
    "address.postalCode": ["pin", "pincode", "postal code", "zip", "zip code"],
    "family.fatherFullName": ["father", "father name", "father's name"],
    "family.spouseFullName": ["spouse", "spouse name", "spouse's name", "husband name", "wife name"],
    "selectedCountry": ["destination", "country", "destination country", "target country"],
    "travel.visaType": ["visa", "visa type", "visa_type"],
    "employment.employerName": ["employer", "company", "organization", "employer name"],
    "employment.jobTitle": ["job title", "designation", "role", "occupation"]
}


def guess_field_mapping(headers: List[str]) -> Dict[str, str]:
    """Guesses which table header maps to each schema field using normalized heuristics."""
    mappings = {}
    for h in headers:
        clean_h = h.strip().lower().replace("_", " ").replace("-", " ")
        for field_key, aliases in COLUMN_HEURISTICS.items():
            if field_key in mappings.values():
                continue
            for alias in aliases:
                if alias == clean_h or alias in clean_h:
                    mappings[h] = field_key
                    break
            if h in mappings:
                break
    return mappings


class ImportExecuteRequest(BaseModel):
    tempId: str
    mappings: Dict[str, str]
    duplicateStrategy: str = "update"  # "skip" | "update" | "create_new"


@router.post("/upload")
async def upload_bulk_file(file: UploadFile = File(...)):
    """
    Uploads .xlsx, .xls, or .csv file for admin bulk data ingestion.
    Parses headers and first 5 sample rows, suggesting intelligent mappings.
    """
    filename = file.filename or "bulk_import.csv"
    ext = os.path.splitext(filename)[1].lower()
    if ext not in [".xlsx", ".xls", ".csv"]:
        raise HTTPException(status_code=400, detail="Only .xlsx, .xls, and .csv files are supported.")

    content = await file.read()
    headers: List[str] = []
    rows: List[Dict[str, str]] = []

    if ext == ".csv":
        try:
            text = content.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = content.decode("latin-1")
        reader = csv.DictReader(io.StringIO(text))
        headers = reader.fieldnames or []
        for r in reader:
            rows.append({k: str(v or "").strip() for k, v in r.items() if k})
    else:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))
        if not all_rows:
            raise HTTPException(status_code=400, detail="The uploaded Excel sheet is empty.")
        raw_headers = all_rows[0]
        headers = [str(h or f"Column_{i+1}").strip() for i, h in enumerate(raw_headers) if h is not None]
        for row in all_rows[1:]:
            if any(cell is not None and str(cell).strip() for cell in row):
                row_dict = {}
                for i, h in enumerate(headers):
                    val = row[i] if i < len(row) else ""
                    if isinstance(val, datetime):
                        row_dict[h] = val.strftime("%Y-%m-%d")
                    else:
                        row_dict[h] = str(val or "").strip()
                rows.append(row_dict)

    if not headers or not rows:
        raise HTTPException(status_code=400, detail="No readable rows found in the uploaded file.")

    temp_id = f"TMP-{uuid.uuid4().hex[:12]}"
    temp_file = os.path.join(TEMP_IMPORT_DIR, f"{temp_id}.json")
    with open(temp_file, "w", encoding="utf-8") as f:
        json.dump({"filename": filename, "headers": headers, "rows": rows}, f)

    detected_mappings = guess_field_mapping(headers)

    return {
        "success": True,
        "tempId": temp_id,
        "filename": filename,
        "headers": headers,
        "totalRows": len(rows),
        "sampleRows": rows[:5],
        "detectedMappings": detected_mappings,
        "targetFields": TARGET_FIELDS
    }


@router.post("/execute")
async def execute_bulk_import(req: ImportExecuteRequest):
    """
    Executes the bulk data import using verified column mappings and duplicate strategy.
    """
    temp_file = os.path.join(TEMP_IMPORT_DIR, f"{req.tempId}.json")
    if not os.path.exists(temp_file):
        raise HTTPException(status_code=404, detail="Upload session expired or invalid. Please re-upload.")

    with open(temp_file, "r", encoding="utf-8") as f:
        data = json.load(f)

    rows = data.get("rows", [])
    filename = data.get("filename", "Import")
    mappings = req.mappings  # header_name -> target_field_key

    # Reverse lookup map: target_field_key -> header_name
    field_to_col = {v: k for k, v in mappings.items() if v}

    imported_count = 0
    updated_count = 0
    skipped_count = 0
    errors_count = 0

    conn = get_connection()
    cursor = conn.cursor()

    for idx, row in enumerate(rows):
        try:
            # Extract row values according to mappings
            row_vals = {}
            for field_key, col_name in field_to_col.items():
                row_vals[field_key] = row.get(col_name, "").strip()

            p_num = row_vals.get("passport.passportNumber", "").strip().upper()
            if not p_num and not row_vals.get("personal.fullName"):
                skipped_count += 1
                continue

            # Check if an existing application exists for this passport number
            existing_app_id = None
            if p_num:
                cursor.execute("SELECT id FROM applications WHERE passport_number = ? LIMIT 1", (p_num,))
                found = cursor.fetchone()
                if found:
                    existing_app_id = found["id"]

            if existing_app_id:
                if req.duplicateStrategy == "skip":
                    skipped_count += 1
                    continue
                elif req.duplicateStrategy == "update":
                    app = get_application(existing_app_id) or MasterApplicationData(applicationId=existing_app_id)
                    updated_count += 1
                else:  # "create_new"
                    app = MasterApplicationData(applicationId=generate_next_application_id())
                    imported_count += 1
            else:
                app = MasterApplicationData(applicationId=generate_next_application_id())
                imported_count += 1

            # Populate fields
            if p_num:
                app.passport.passportNumber = p_num
            if row_vals.get("personal.fullName"):
                app.personal.fullName = row_vals["personal.fullName"]
            if row_vals.get("personal.givenName"):
                app.personal.givenName = row_vals["personal.givenName"]
            if row_vals.get("personal.surname"):
                app.personal.surname = row_vals["personal.surname"]
            if row_vals.get("personal.dob"):
                app.personal.dob = row_vals["personal.dob"]
            if row_vals.get("personal.gender"):
                app.personal.gender = row_vals["personal.gender"]
            if row_vals.get("personal.nationality"):
                app.personal.nationality = row_vals["personal.nationality"]
            if row_vals.get("passport.issueDate"):
                app.passport.issueDate = row_vals["passport.issueDate"]
            if row_vals.get("passport.expiryDate"):
                app.passport.expiryDate = row_vals["passport.expiryDate"]
            if row_vals.get("passport.issuePlace"):
                app.passport.issuePlace = row_vals["passport.issuePlace"]
            if row_vals.get("contact.mobileNumber"):
                app.contact.mobileNumber = row_vals["contact.mobileNumber"]
            if row_vals.get("contact.emailAddress"):
                app.contact.emailAddress = row_vals["contact.emailAddress"]
            if row_vals.get("address.addressLine1"):
                app.address.addressLine1 = row_vals["address.addressLine1"]
            if row_vals.get("address.city"):
                app.address.city = row_vals["address.city"]
            if row_vals.get("address.state"):
                app.address.state = row_vals["address.state"]
            if row_vals.get("address.postalCode"):
                app.address.postalCode = row_vals["address.postalCode"]
            if row_vals.get("family.fatherFullName"):
                app.family.fatherFullName = row_vals["family.fatherFullName"]
            if row_vals.get("family.spouseFullName"):
                app.family.spouseFullName = row_vals["family.spouseFullName"]
            if row_vals.get("selectedCountry"):
                app.selectedCountry = row_vals["selectedCountry"]
            if row_vals.get("travel.visaType"):
                app.travel.visaType = row_vals["travel.visaType"]
            if row_vals.get("employment.employerName"):
                app.employment.employerName = row_vals["employment.employerName"]
            if row_vals.get("employment.jobTitle"):
                app.employment.jobTitle = row_vals["employment.jobTitle"]

            app.applicationStatus = "Imported"
            saved_app = save_application(app)
            try:
                ExcelService.sync_master_data_to_excel(saved_app)
            except Exception:
                pass

        except Exception as row_err:
            print(f"[Bulk Import] Error row {idx}: {row_err}")
            errors_count += 1

    conn.close()

    # Log to import history
    hist_id = record_import_history(
        filename=filename,
        imported_count=imported_count,
        updated_count=updated_count,
        skipped_count=skipped_count,
        errors_count=errors_count,
        status="Completed" if errors_count == 0 else "Completed with Warnings"
    )

    # Clean up temp file
    try:
        os.remove(temp_file)
    except Exception:
        pass

    return {
        "success": True,
        "historyId": hist_id,
        "importedCount": imported_count,
        "updatedCount": updated_count,
        "skippedCount": skipped_count,
        "errorsCount": errors_count,
        "message": f"Bulk import complete: {imported_count} imported, {updated_count} updated, {skipped_count} skipped, {errors_count} errors."
    }


@router.get("/history")
async def get_import_history():
    """Returns past import runs."""
    return list_import_history()
