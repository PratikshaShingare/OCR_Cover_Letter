"""
Khanna Travels & Holidays — SQLite Application Storage Service
Persists verified master application data, tracks versions and audit timestamps.
"""

import sqlite3
import json
import os
import shutil
from datetime import datetime
from typing import List, Optional, Dict, Any
from ..models.master_schema import MasterApplicationData


DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "khanna_visa.db")


def get_connection():
    os.makedirs(os.path.dirname(os.path.abspath(DB_PATH)), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS applications (
        id TEXT PRIMARY KEY,
        applicant_name TEXT,
        passport_number TEXT,
        destination_country TEXT,
        visa_type TEXT,
        status TEXT,
        created_at TEXT,
        updated_at TEXT,
        data_json TEXT
    )
    """)
    conn.commit()
    conn.close()


def generate_next_application_id() -> str:
    """Generates sequential clean ID like APP-2026-00001."""
    init_db()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id FROM applications WHERE id LIKE 'APP-____-_____'")
    rows = cursor.fetchall()
    conn.close()
    year = datetime.now().year
    max_seq = 0
    for r in rows:
        app_id = r["id"]
        parts = app_id.split("-")
        if len(parts) == 3 and parts[2].isdigit():
            max_seq = max(max_seq, int(parts[2]))
    return f"APP-{year}-{max_seq + 1:05d}"


def save_application(app_data: MasterApplicationData) -> MasterApplicationData:
    if not app_data.applicationId:
        app_data.applicationId = generate_next_application_id()
        
    init_db()
    conn = get_connection()
    cursor = conn.cursor()
    
    full_name = app_data.personal.fullName or f"{app_data.personal.givenName} {app_data.personal.surname}".strip()
    pass_num = app_data.passport.passportNumber
    country = app_data.selectedCountry
    visa_type = app_data.travel.visaType
    status = app_data.applicationStatus
    now_iso = app_data.updatedAt
    
    data_json = app_data.model_dump_json()
    
    cursor.execute("""
    INSERT INTO applications (id, applicant_name, passport_number, destination_country, visa_type, status, created_at, updated_at, data_json)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET
        applicant_name = excluded.applicant_name,
        passport_number = excluded.passport_number,
        destination_country = excluded.destination_country,
        visa_type = excluded.visa_type,
        status = excluded.status,
        updated_at = excluded.updated_at,
        data_json = excluded.data_json
    """, (app_data.applicationId, full_name, pass_num, country, visa_type, status, app_data.createdAt, now_iso, data_json))
    
    conn.commit()
    conn.close()
    return app_data


def get_application(app_id: str) -> Optional[MasterApplicationData]:
    init_db()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT data_json FROM applications WHERE id = ?", (app_id,))
    row = cursor.fetchone()
    conn.close()
    
    if row:
        return MasterApplicationData.model_validate_json(row["data_json"])
    return None


def list_applications() -> List[Dict[str, Any]]:
    init_db()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, applicant_name, passport_number, destination_country, visa_type, status, created_at, updated_at FROM applications ORDER BY updated_at DESC")
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def delete_application(app_id: str) -> bool:
    """Deletes application record and its associated files."""
    init_db()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM applications WHERE id = ?", (app_id,))
    deleted = cursor.rowcount > 0
    conn.commit()
    conn.close()
    
    if deleted:
        # 1. Remove application storage directory (passport files, generated assets)
        app_dir = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "internal_records", "applications", app_id))
        if os.path.exists(app_dir):
            shutil.rmtree(app_dir, ignore_errors=True)
            
        # 2. Update Master Excel workbook
        try:
            from .excel_service import ExcelService
            ExcelService.remove_application_from_master_excel(app_id)
        except Exception:
            pass
            
        # Legacy per-client Excel cleanup (if present)
        excel_path = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "internal_records", "excels", f"Client_Data_{app_id}.xlsx"))
        if os.path.exists(excel_path):
            try:
                os.remove(excel_path)
            except Exception:
                pass
    return deleted


def clear_all_applications():
    """Purges all application records and associated files."""
    init_db()
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM applications")
    conn.commit()
    conn.close()
    
    apps_dir = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "internal_records", "applications"))
    if os.path.exists(apps_dir):
        shutil.rmtree(apps_dir, ignore_errors=True)
        
    excels_dir = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "internal_records", "excels"))
    if os.path.exists(excels_dir):
        shutil.rmtree(excels_dir, ignore_errors=True)

    try:
        from .excel_service import ExcelService
        ExcelService.reset_master_excel()
    except Exception:
        pass


