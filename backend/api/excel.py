"""
Khanna Travels & Holidays — Excel Generation & Import API
Builds professional 5-sheet client workbook and diffs uploaded workbooks.
"""

from fastapi import APIRouter, HTTPException, UploadFile, File, Form
from fastapi.responses import FileResponse
import tempfile
import os
import json

from ..models.master_schema import MasterApplicationData
from ..services.excel_service import ExcelService
from ..excel_generator.excel_builder import parse_and_diff_excel

router = APIRouter(prefix="/api/excel", tags=["Excel"])


@router.post("/generate")
async def generate_client_excel(app_data: MasterApplicationData):
    """Syncs to and returns the single Master Excel workbook."""
    try:
        excel_path = ExcelService.sync_master_data_to_excel(app_data)
        return FileResponse(
            path=excel_path,
            filename="Khanna_Travels_Client_Master.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate Master Excel workbook: {str(e)}")


@router.get("/master")
async def download_master_excel():
    """Serves the single centralized Master Excel workbook."""
    excel_path = ExcelService.ensure_master_excel_initialized()
    return FileResponse(
        path=excel_path,
        filename="Khanna_Travels_Client_Master.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@router.post("/import-diff")
async def import_and_diff(
    file: UploadFile = File(...),
    currentDataJson: str = Form(...)
):
    """
    Parses an edited Excel file and detects differences against current master data.
    Returns comparison preview without blindly overwriting verified master data.
    """
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx Excel files are supported.")
        
    try:
        file_bytes = await file.read()
        current_data = MasterApplicationData.model_validate_json(currentDataJson)
        
        diff_result = parse_and_diff_excel(file_bytes, current_data)
        return diff_result
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to inspect Excel file: {str(e)}")
