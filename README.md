# Khanna Travels & Holidays — Visa Document Automation System

Production-ready visa document automation system engineered for the **Khanna Travels & Holidays** executive visa-processing operations team.

The system ingests passport copies (scans, camera photos, multi-page PDFs), performs intelligent OCR and MRZ extraction with 4-way auto-orientation, allows visa executives to verify and edit structured data with field-level confidence indicators, automatically synchronizes a centralized single-sheet Master Excel client record (`.xlsx`), generates country-specific visa cover letters adhering strictly to approved company templates (Europe/Standard, Japan, Singapore), provides an in-browser rich document editor, and exports genuine Microsoft Word (`.docx`) and vector-rendered (`.pdf`) cover letters.

---

## 1. End-to-End Application Workflow

```
Passport Upload (PDF / PNG / JPG)
               ↓
OCR & 4-Way Auto-Orientation (0°, 90°, 180°, 270°)
               ↓
ICAO 9303 TD3 MRZ Parser + VIZ Regex Extraction
               ↓
Master Client Data Model (Central Source of Truth)
               ↓
Executive Verification & Data Editing (Field Badges)
               ↓
Travel, Itinerary & Accommodation Details
               ↓
ONE Centralized Master Excel Workbook (Row Upsert by Application ID)
               ↓
Country Template Selection (Standard Europe / Japan / Singapore)
               ↓
Cover Letter Generation Engine (Exact Layout Matching)
               ↓
In-Browser Executive Rich Document Editor
               ↓
DOCX & PDF Final Document Export
```

* **Centralized Source of Truth**: The Master Client Data model acts as the single source of truth across all features. Any correction made during verification automatically propagates to generated documents and the centralized Master Excel record.
* **Server-Side Data Storage**: Master application records are stored server-side in the SQLite database and consolidated into a single master Excel file on the server.

---

## 2. Master Excel Storage Architecture

The system enforces a strict single-workbook, single-worksheet operational architecture:

* **File Location**: `data/Khanna_Travels_Client_Master.xlsx` (server-side application storage).
* **Worksheet Name**: `Client Data` (exactly ONE worksheet).
* **Row-per-Application**:
  - **Row 1**: Formatted executive header row (Calibri, Navy blue fill `#1E3A8A`, white bold text, cell borders, autofitted columns).
  - **Row 2+**: Each visa application occupies exactly one row.
* **Row-Level Upsert by Application ID**:
  - Uses `Application ID` (e.g. `APP-2026-00001`) in Column A as the unique key.
  - When an existing application is updated, its existing row is updated in place (preventing duplicate rows).
  - Genuinely new applications are appended as new rows.
  - When an application is deleted, its row is removed from the worksheet.
* **Master Client Data Fields Included**:
  - Application ID, Status, Full Name, Title, Given Name, Middle Name, Surname, DOB, Gender, Place of Birth, Country of Birth, Nationality, Passport Number, Passport Type, Passport Issue Date, Passport Expiry Date, Passport Issue Place, Issuing Country, Issuing Authority, Address, Address Line 2, City, State, Country, PIN, Permanent Address, Father Name, Mother Name, Spouse Name, Spouse Passport Number, Phone, Email, Emergency Contact, Occupation, Employer, Job Title, Annual Income, Destination, All Destinations, Visa Type, Purpose, Travel Start Date, Travel End Date, Number of Days, Number of Nights, Flight Number, Departure Airport, Arrival Airport, Hotel, Accommodation Info, Trip Sponsor, Sponsor Name, Bank Statement Available, ITR Available, Salary Slips Available, Employment Letter Available, Financial Info, Travellers Count, Additional Travellers, Selected Template, Additional Information, Created At, Updated At.

---

## 3. Reference Cover Letter Templates

The system includes and conforms to the approved Khanna Travels & Holidays cover-letter templates:

| Country Format | Recipient Consulate | Approved Template Files | Structural Features |
| :--- | :--- | :--- | :--- |
| **Europe / Standard (France)** | The Consulate General of France, Mumbai | `Europe_covering_letter_template_clean.docx`<br>`Europe_covering_letter_template_clean.pdf` | Narrative corporate employment, Schengen itinerary breakdown with nights per country, return commitment. |
| **Japan** | Consulate General of Japan, Mumbai | `Japan_covering_letter_template_clean.docx`<br>`Japan_covering_letter_template_clean.pdf` | Formatted **3-Column Hotel Table** (`NAME \| DATE \| CONTACT NO.`), family member school/grade info, holiday purpose wording. |
| **Singapore** | Consulate General of the Republic of Singapore, Mumbai | `Singapore_covering_letter_template_clean.docx`<br>`Singapore_covering_letter_template_clean.pdf` | Formatted **5-Column Passenger Table** (`Sr no. \| Passengers Name \| Passport No \| Relation \| Occupation`), underlined subject, paragraph hotel stay, explicit **Multiple Entries** visa request. |

---

## 4. Technology Stack

* **Backend**: Python 3.10+ / 3.14, FastAPI, Uvicorn, Pydantic v2
* **OCR & Document Processing**:
  * Auto-Orientation (0°, 90°, 180°, 270°)
  * ICAO 9303 TD3 MRZ parser
  * VIZ regex extractor
  * `pypdfium2` PDF rasterizer
* **Document & Spreadsheet Generation**:
  * `openpyxl`: Single centralized Master Excel workbook (`data/Khanna_Travels_Client_Master.xlsx`) with ONE worksheet (`Client Data`) and row-level upsert by Application ID
  * `python-docx`: Native Microsoft Word DOCX generator adhering strictly to approved templates
  * `reportlab`: Platypus vector PDF generation matching DOCX typography
* **Database**: SQLite (`khanna_visa.db`) initialized automatically on startup with sequential IDs (`APP-YYYY-NNNNN`)
* **Frontend**: HTML5, CSS3 (Executive styling, responsive design), Vanilla JavaScript (Modular ES modules)

---

## 5. Getting Started & Running Instructions

### Prerequisites
* Python 3.10 or higher
* Git

### Step 1: Install Dependencies
```bash
# Install backend dependencies
python -m pip install -r backend/requirements.txt
```

### Step 2: Run the Application
Start the FastAPI server (serves the application API and executive frontend portal):

```bash
python -m uvicorn backend.main:app --host 127.0.0.1 --port 8000 --reload
```

Open your browser at:
```
http://127.0.0.1:8000
```

---

## 6. Running Automated Tests

Run the comprehensive unit and integration test suite:

```bash
python -m unittest discover tests
```

*Tests verify: Master Schema validation, Base Europe template assembly, Japan hotel table structure, Singapore passenger table structure, Single Master Excel workbook and worksheet upsert, python-docx document generation, ReportLab PDF generation, and all FastAPI endpoints.*

---

## 7. Security & Data Protection

* **Zero Client Data in Version Control**: Uploaded passport scans, processed preview images, generated cover letters, and client Excel files are stored on the server and are strictly excluded from version control via `.gitignore`.
* **Master Excel Exclusion**: The runtime master workbook `data/Khanna_Travels_Client_Master.xlsx` is created dynamically at runtime and is never committed to GitHub.
* **No Hardcoded Secrets**: Secrets and SQLite database files (`*.db`) are excluded by `.gitignore`.

