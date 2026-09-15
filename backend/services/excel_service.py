"""
Khanna Travels & Holidays — Single Master Client Excel Service
Maintains exactly ONE master Excel workbook and ONE worksheet on the server.
Each application occupies exactly ONE row, identified by its unique Application ID.
When an existing application is updated, its existing row is updated in place.
"""

import os
from typing import Optional, List, Tuple, Any, Dict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from ..models.master_schema import MasterApplicationData

MASTER_DATA_DIR = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "data")
)
MASTER_EXCEL_PATH = os.path.join(MASTER_DATA_DIR, "Khanna_Travels_Client_Master.xlsx")
SHEET_NAME = "Client Data"

# Branding & Styling
FONT_HEADER = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FILL_HEADER = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
FONT_DATA = Font(name="Calibri", size=11, color="0F172A")
FILL_ZEBRA = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
FILL_WHITE = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

THIN_BORDER_SIDE = Side(border_style="thin", color="CBD5E1")
BORDER_CELL = Border(
    left=THIN_BORDER_SIDE, right=THIN_BORDER_SIDE, top=THIN_BORDER_SIDE, bottom=THIN_BORDER_SIDE
)

# Comprehensive column definitions mapped to MasterApplicationData fields
EXCEL_COLUMNS: List[Tuple[str, Any]] = [
    ("Application ID", lambda a: a.applicationId or ""),
    ("Status", lambda a: a.applicationStatus or "Draft"),
    ("Full Name", lambda a: a.personal.fullName or ""),
    ("Title", lambda a: a.personal.title or ""),
    ("Given Name", lambda a: a.personal.givenName or ""),
    ("Middle Name", lambda a: a.personal.middleName or ""),
    ("Surname", lambda a: a.personal.surname or ""),
    ("DOB", lambda a: a.personal.dob or ""),
    ("Gender", lambda a: a.personal.gender or ""),
    ("Place of Birth", lambda a: a.personal.placeOfBirth or ""),
    ("Country of Birth", lambda a: a.personal.countryOfBirth or ""),
    ("Nationality", lambda a: a.personal.nationality or ""),
    ("Passport Number", lambda a: a.passport.passportNumber or ""),
    ("Passport Type", lambda a: a.passport.passportType or "P"),
    ("Passport Issue Date", lambda a: a.passport.issueDate or ""),
    ("Passport Expiry Date", lambda a: a.passport.expiryDate or ""),
    ("Passport Issue Place", lambda a: a.passport.issuePlace or ""),
    ("Issuing Country", lambda a: a.passport.issuingCountry or ""),
    ("Issuing Authority", lambda a: a.passport.issuingAuthority or ""),
    ("Address", lambda a: a.address.addressLine1 or a.address.currentResidentialAddress or ""),
    ("Address Line 2", lambda a: a.address.addressLine2 or ""),
    ("City", lambda a: a.address.city or ""),
    ("State", lambda a: a.address.state or ""),
    ("Country", lambda a: a.address.country or ""),
    ("PIN", lambda a: a.address.postalCode or ""),
    ("Permanent Address", lambda a: a.address.permanentAddress or ""),
    ("Father Name", lambda a: a.family.fatherFullName or ""),
    ("Mother Name", lambda a: a.family.motherFullName or ""),
    ("Spouse Name", lambda a: a.family.spouseFullName or ""),
    ("Spouse Passport Number", lambda a: a.family.spousePassportNumber or ""),
    ("Phone", lambda a: a.contact.mobileNumber or ""),
    ("Email", lambda a: a.contact.emailAddress or ""),
    ("Emergency Contact", lambda a: a.contact.emergencyContact or ""),
    ("Occupation", lambda a: a.employment.employmentStatus or ""),
    ("Employer", lambda a: a.employment.employerName or ""),
    ("Job Title", lambda a: a.employment.jobTitle or ""),
    ("Annual Income", lambda a: a.employment.annualIncome or ""),
    ("Destination", lambda a: a.travel.destinationCountry or ""),
    ("All Destinations", lambda a: ", ".join(a.travel.destinationCountries) if a.travel.destinationCountries else (a.travel.destinationCountry or "")),
    ("Visa Type", lambda a: a.travel.visaType or ""),
    ("Purpose", lambda a: a.travel.purposeOfTravel or ""),
    ("Travel Start Date", lambda a: a.travel.travelStartDate or ""),
    ("Travel End Date", lambda a: a.travel.travelEndDate or ""),
    ("Number of Days", lambda a: a.travel.numberOfDays or 0),
    ("Number of Nights", lambda a: a.travel.numberOfNights or 0),
    ("Flight Number", lambda a: a.travel.flightNumber or ""),
    ("Departure Airport", lambda a: a.travel.departureAirport or ""),
    ("Arrival Airport", lambda a: a.travel.arrivalAirport or ""),
    ("Hotel", lambda a: ", ".join(h.hotelName for h in a.hotels if h.hotelName) if a.hotels else ""),
    ("Accommodation Info", lambda a: "; ".join(f"{h.hotelName} ({h.city}, {h.country}, In: {h.checkInDate}, Out: {h.checkOutDate})" for h in a.hotels if h.hotelName) if a.hotels else ""),
    ("Trip Sponsor", lambda a: a.financial.tripSponsor or ""),
    ("Sponsor Name", lambda a: a.financial.sponsorName or ""),
    ("Bank Statement Available", lambda a: "Yes" if a.financial.bankStatementAvailable else "No"),
    ("ITR Available", lambda a: "Yes" if a.financial.itrAvailable else "No"),
    ("Salary Slips Available", lambda a: "Yes" if a.financial.salarySlipsAvailable else "No"),
    ("Employment Letter Available", lambda a: "Yes" if a.financial.employmentLetterAvailable else "No"),
    ("Financial Info", lambda a: a.financial.notes or (f"Sponsor: {a.financial.tripSponsor}" if a.financial.tripSponsor else "")),
    ("Travellers Count", lambda a: len(a.travellers)),
    ("Additional Travellers", lambda a: "; ".join(f"{t.fullName} ({t.relationship}, Pass: {t.passportNumber})" for t in a.travellers if t.fullName) if a.travellers else ""),
    ("Selected Template", lambda a: a.selectedTemplateId or "standard"),
    ("Additional Information", lambda a: a.additional.content or ""),
    ("Created At", lambda a: a.createdAt or ""),
    ("Updated At", lambda a: a.updatedAt or "")
]


def _format_header_row(ws):
    """Sets row 1 headers with dark blue branding and borders."""
    headers = [col[0] for col in EXCEL_COLUMNS]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
        cell.border = BORDER_CELL
    ws.row_dimensions[1].height = 28


def _autofit_columns(ws, max_cols: int):
    """Adjusts column widths dynamically with reasonable min and max limits."""
    for col in range(1, max_cols + 1):
        col_letter = get_column_letter(col)
        max_len = 0
        for cell in ws[col_letter]:
            if cell.value is not None:
                val_str = str(cell.value)
                line_max = max(len(l) for l in val_str.split("\n"))
                max_len = max(max_len, line_max)
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)


def _get_or_create_workbook() -> Tuple[openpyxl.Workbook, Any]:
    """
    Opens the single master workbook or creates it if missing.
    Guarantees exactly ONE worksheet named 'Client Data'.
    """
    os.makedirs(MASTER_DATA_DIR, exist_ok=True)
    if os.path.exists(MASTER_EXCEL_PATH):
        try:
            wb = openpyxl.load_workbook(MASTER_EXCEL_PATH)
        except Exception:
            wb = openpyxl.Workbook()
        
        if SHEET_NAME in wb.sheetnames:
            ws = wb[SHEET_NAME]
        else:
            ws = wb.create_sheet(SHEET_NAME, 0)
        
        # Enforce ONE worksheet only
        for name in list(wb.sheetnames):
            if name != SHEET_NAME:
                del wb[name]
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
    
    # Always ensure row 1 headers are intact
    _format_header_row(ws)
    return wb, ws


class ExcelService:
    @staticmethod
    def sync_master_data_to_excel(app_data: MasterApplicationData) -> str:
        """
        Automatically writes or updates verified master client data in the
        ONE single master Excel workbook on the server.
        Updates the existing row if applicationId is found; appends a new row if not.
        """
        app_id = (app_data.applicationId or "").strip()
        if not app_id:
            return MASTER_EXCEL_PATH
        
        wb, ws = _get_or_create_workbook()
        
        # Locate existing row by Application ID in Column A (starting from row 2)
        target_row = None
        for r in range(2, ws.max_row + 1):
            cell_val = ws.cell(row=r, column=1).value
            if cell_val is not None and str(cell_val).strip() == app_id:
                target_row = r
                break
        
        if target_row is None:
            # Genuinely new application: append a new row
            target_row = ws.max_row + 1
        
        is_even = (target_row % 2 == 0)
        row_fill = FILL_WHITE if is_even else FILL_ZEBRA
        
        # Center align specific columns (ID, dates, codes, status)
        center_cols = {1, 2, 8, 9, 12, 13, 15, 16, 22, 25, 42, 43, 44, 45, 53, 54, 55, 56, 58, 61, 62, 63}
        
        for col_idx, (_, extractor) in enumerate(EXCEL_COLUMNS, 1):
            try:
                val = extractor(app_data)
            except Exception:
                val = ""
            cell = ws.cell(row=target_row, column=col_idx, value=val)
            cell.font = FONT_DATA
            cell.fill = row_fill
            cell.border = BORDER_CELL
            cell.alignment = Alignment(
                horizontal="center" if col_idx in center_cols else "left",
                vertical="center"
            )
        
        ws.row_dimensions[target_row].height = 22
        _autofit_columns(ws, len(EXCEL_COLUMNS))
        
        wb.save(MASTER_EXCEL_PATH)
        wb.close()
        return MASTER_EXCEL_PATH

    @staticmethod
    def remove_application_from_master_excel(app_id: str) -> bool:
        """Removes the application row from the master workbook if deleted."""
        if not os.path.exists(MASTER_EXCEL_PATH):
            return False
        app_id = (app_id or "").strip()
        if not app_id:
            return False
            
        wb, ws = _get_or_create_workbook()
        target_row = None
        for r in range(2, ws.max_row + 1):
            cell_val = ws.cell(row=r, column=1).value
            if cell_val is not None and str(cell_val).strip() == app_id:
                target_row = r
                break
                
        if target_row is not None:
            ws.delete_rows(target_row)
            wb.save(MASTER_EXCEL_PATH)
            wb.close()
            return True
            
        wb.close()
        return False

    @staticmethod
    def reset_master_excel():
        """Resets the master workbook to empty headers."""
        if os.path.exists(MASTER_EXCEL_PATH):
            try:
                os.remove(MASTER_EXCEL_PATH)
            except Exception:
                pass
        _get_or_create_workbook()

    @staticmethod
    def ensure_master_excel_initialized() -> str:
        """Guarantees the master workbook exists with the header row."""
        wb, _ = _get_or_create_workbook()
        wb.save(MASTER_EXCEL_PATH)
        wb.close()
        return MASTER_EXCEL_PATH

    @staticmethod
    def get_master_excel_path() -> str:
        """Returns the centralized path to the single master Excel workbook."""
        return MASTER_EXCEL_PATH

    @staticmethod
    def get_internal_excel_path(app_id: Optional[str] = None) -> Optional[str]:
        """Backward-compatible helper returning the single master Excel workbook path."""
        if os.path.exists(MASTER_EXCEL_PATH):
            return MASTER_EXCEL_PATH
        return None
