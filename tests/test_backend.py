"""
Automated Test Suite for Khanna Travels & Holidays Visa System
Verifies:
1. Every new application starts with completely blank client fields.
2. Zero hardcoded sample client data exists.
3. France Base Template generates accurately from current application data only.
4. Japan Country Override generates 3-column hotel table from current application data only.
5. Singapore Country Override generates 5-column passenger table and underlined subject.
6. Automated internal Excel generation populates 5 sheets directly from verified master data.
7. Genuine Word DOCX generation.
8. ReportLab PDF generation.
"""

import os
import sys
import unittest
import tempfile
import openpyxl
from docx import Document

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from backend.models.master_schema import (
    MasterApplicationData, ApplicantPersonal, ApplicantPassport, ApplicantAddress,
    ApplicantContact, ApplicantEmployment, Traveller, TravelDetails, Hotel, FinancialInfo
)
from backend.document_generator.cover_letter_engine import CoverLetterEngine
from backend.document_generator.docx_generator import build_docx_cover_letter
from backend.pdf_generator.pdf_builder import build_pdf_cover_letter
from backend.services.excel_service import ExcelService


class TestKhannaVisaSystem(unittest.TestCase):
    
    def setUp(self):
        self.engine = CoverLetterEngine()

    def test_01_new_application_starts_100_percent_blank(self):
        """Verify that every newly created application record has empty client fields."""
        app = MasterApplicationData(applicationId="APP-TEST-BLANK")
        self.assertEqual(app.personal.fullName, "")
        self.assertEqual(app.personal.givenName, "")
        self.assertEqual(app.personal.surname, "")
        self.assertEqual(app.personal.dob, "")
        self.assertEqual(app.passport.passportNumber, "")
        self.assertEqual(app.passport.issueDate, "")
        self.assertEqual(app.passport.expiryDate, "")
        self.assertEqual(app.address.city, "")
        self.assertEqual(app.contact.mobileNumber, "")
        self.assertEqual(app.contact.emailAddress, "")
        self.assertEqual(app.employment.employerName, "")
        self.assertEqual(len(app.travellers), 0)
        self.assertEqual(len(app.hotels), 0)
        self.assertEqual(app.travel.travelStartDate, "")
        self.assertEqual(app.travel.travelEndDate, "")
        self.assertEqual(app.travel.numberOfDays, 0)
        self.assertEqual(app.travel.numberOfNights, 0)

    def test_02_france_cover_letter_without_sample_data(self):
        """Verify France base template generation uses only current application data."""
        app = MasterApplicationData(
            selectedCountry="France",
            personal=ApplicantPersonal(fullName="Mr. Amit Verma", title="Mr.", nationality="Indian"),
            passport=ApplicantPassport(passportNumber="Z9876543", issuePlace="Delhi", issueDate="2023-05-10"),
            address=ApplicantAddress(city="Delhi", country="India"),
            employment=ApplicantEmployment(employmentStatus="Employed", employerName="Global Tech Solutions", jobTitle="Tech Lead"),
            travellers=[
                Traveller(
                    title="Mrs.", fullName="Mrs. Neha Verma", passportNumber="Y1234567",
                    relationship="Spouse", occupation="Architect"
                )
            ],
            travel=TravelDetails(
                destinationCountry="France",
                countriesToBeVisited="France",
                travelStartDate="2026-10-01",
                travelEndDate="2026-10-15",
                numberOfNights=14,
                purposeOfTravel="Tourism"
            ),
            financial=FinancialInfo(tripSponsor="Self-funded")
        )
        
        doc_result = self.engine.generate_document(app)
        all_text = " ".join(b.get("content", "") for b in doc_result["blocks"] if b.get("type") == "paragraph")
        
        # Verify current data is present
        self.assertIn("Amit Verma", all_text)
        self.assertIn("Z9876543", all_text)
        self.assertIn("Global Tech Solutions", all_text)
        self.assertIn("Neha Verma", all_text)
        
        # Strict verification: NEVER contain sample client names or employers from previous demo
        self.assertNotIn("Shirish", all_text)
        self.assertNotIn("Ghosal", all_text)
        self.assertNotIn("The Trade Desk", all_text)
        self.assertNotIn("LinkedIn India", all_text)
        self.assertNotIn("VisionSpring", all_text)

    def test_03_japan_hotel_table_uses_current_data_only(self):
        """Verify Japan hotel table renders only entered hotels."""
        app = MasterApplicationData(
            selectedCountry="Japan",
            personal=ApplicantPersonal(fullName="Mr. Rajesh Kumar", title="Mr.", nationality="Indian"),
            passport=ApplicantPassport(passportNumber="K7654321", issuePlace="Bengaluru"),
            employment=ApplicantEmployment(employerName="Apex Consulting"),
            travel=TravelDetails(destinationCountry="Japan", purposeOfTravel="Holiday"),
            hotels=[
                Hotel(hotelName="Tokyo Bay Hotel", city="Tokyo", checkInDate="2026-06-01", checkOutDate="2026-06-05", contactNumber="+81 3-0000-0000")
            ]
        )
        
        doc_result = self.engine.generate_document(app)
        table_blocks = [b for b in doc_result["blocks"] if b.get("type") == "table"]
        self.assertEqual(len(table_blocks), 1)
        self.assertEqual(table_blocks[0]["headers"], ["NAME", "DATE", "CONTACT NO."])
        self.assertEqual(len(table_blocks[0]["rows"]), 1)
        self.assertIn("Tokyo Bay Hotel", table_blocks[0]["rows"][0][0])
        
        # Verify sample hotel names do NOT appear
        self.assertNotIn("RIHGA Royal", str(doc_result))
        self.assertNotIn("Mitsui Garden", str(doc_result))

    def test_04_singapore_passenger_table_uses_current_data_only(self):
        """Verify Singapore passenger table renders current passengers and underlined subject."""
        app = MasterApplicationData(
            selectedCountry="Singapore",
            personal=ApplicantPersonal(fullName="Dr. Vikram Patel", title="Dr.", nationality="Indian"),
            passport=ApplicantPassport(passportNumber="P4567890"),
            employment=ApplicantEmployment(jobTitle="Surgeon"),
            travellers=[
                Traveller(title="Mrs.", fullName="Mrs. Ananya Patel", passportNumber="P4567891", relationship="Wife", occupation="Professor")
            ],
            travel=TravelDetails(destinationCountry="Singapore", entryType="Multiple", numberOfEntries="Multiple Entries")
        )
        
        doc_result = self.engine.generate_document(app)
        self.assertTrue(doc_result["subject"]["underline"])
        
        table_blocks = [b for b in doc_result["blocks"] if b.get("type") == "table"]
        self.assertEqual(len(table_blocks), 1)
        self.assertEqual(len(table_blocks[0]["rows"]), 2) # Self + Wife
        self.assertEqual(table_blocks[0]["rows"][0][1], "Dr. Vikram Patel")
        self.assertEqual(table_blocks[0]["rows"][1][1], "Mrs. Ananya Patel")
        
        # Verify sample client names do NOT appear
        self.assertNotIn("Nimish", str(doc_result))
        self.assertNotIn("Dhara", str(doc_result))

    def test_05_automated_internal_excel_generation(self):
        """Verify ExcelService maintains ONE master workbook, ONE worksheet ('Client Data'), and updates rows."""
        ExcelService.reset_master_excel()
        app = MasterApplicationData(
            applicationId="APP-TEST-EXCEL-INTERNAL",
            selectedCountry="France",
            personal=ApplicantPersonal(fullName="Mr. Suresh Rao", nationality="Indian"),
            passport=ApplicantPassport(passportNumber="R1122334")
        )
        
        generated_path = ExcelService.sync_master_data_to_excel(app)
        self.assertTrue(os.path.exists(generated_path))
        self.assertTrue(generated_path.endswith("Khanna_Travels_Client_Master.xlsx"))
        
        wb = openpyxl.load_workbook(generated_path)
        self.assertEqual(len(wb.sheetnames), 1)
        self.assertEqual(wb.sheetnames[0], "Client Data")
        
        ws = wb["Client Data"]
        # Row 1 is header, Row 2 is APP-TEST-EXCEL-INTERNAL
        self.assertEqual(ws.max_row, 2)
        self.assertEqual(ws.cell(row=2, column=1).value, "APP-TEST-EXCEL-INTERNAL")
        self.assertEqual(ws.cell(row=2, column=3).value, "Mr. Suresh Rao")
        self.assertEqual(ws.cell(row=2, column=13).value, "R1122334")

        # Verify editing the same application updates in place (no duplicate rows)
        app.personal.fullName = "Mr. Suresh Rao Updated"
        ExcelService.sync_master_data_to_excel(app)
        wb2 = openpyxl.load_workbook(generated_path)
        ws2 = wb2["Client Data"]
        self.assertEqual(ws2.max_row, 2)
        self.assertEqual(ws2.cell(row=2, column=3).value, "Mr. Suresh Rao Updated")
        wb.close()
        wb2.close()

    def test_06_docx_generation(self):
        """Verify DOCX generation with python-docx."""
        app = MasterApplicationData(
            selectedCountry="Japan",
            personal=ApplicantPersonal(fullName="Mr. Test Applicant"),
            passport=ApplicantPassport(passportNumber="T0000001")
        )
        doc_data = self.engine.generate_document(app)
        
        temp_file = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        temp_file.close()
        try:
            build_docx_cover_letter(doc_data, temp_file.name)
            self.assertTrue(os.path.exists(temp_file.name))
            self.assertGreater(os.path.getsize(temp_file.name), 1000)
        finally:
            if os.path.exists(temp_file.name):
                os.remove(temp_file.name)

    def test_07_pdf_generation(self):
        """Verify PDF generation with ReportLab."""
        app = MasterApplicationData(
            selectedCountry="Singapore",
            personal=ApplicantPersonal(fullName="Dr. Test User"),
            passport=ApplicantPassport(passportNumber="T0000002")
        )
        doc_data = self.engine.generate_document(app)
        
        temp_file = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        temp_file.close()
        try:
            build_pdf_cover_letter(doc_data, temp_file.name)
            self.assertTrue(os.path.exists(temp_file.name))
            self.assertGreater(os.path.getsize(temp_file.name), 1000)
        finally:
            if os.path.exists(temp_file.name):
                os.remove(temp_file.name)


    def test_08_authoritative_docx_europe_template_fidelity(self):
        """Verify Europe cover letter DOCX clones clean template and replaces all placeholders."""
        app = MasterApplicationData(
            selectedCountry="France",
            selectedTemplateId="standard",
            personal=ApplicantPersonal(fullName="Rahul Sharma", title="Mr.", nationality="Indian", gender="Male"),
            passport=ApplicantPassport(passportNumber="Z9876543", issuePlace="Mumbai", issueDate="2020-01-15"),
            travel=TravelDetails(travelStartDate="2026-10-01", travelEndDate="2026-10-15", numberOfNights=14, purposeOfTravel="Tourism"),
            financial=FinancialInfo(tripSponsor="Self-funded"),
            employment=ApplicantEmployment(jobTitle="Software Engineer", employerName="Tech Solutions Ltd"),
            contact=ApplicantContact(mobileNumber="+91 9876543210", emailAddress="rahul@example.com"),
            travellers=[
                Traveller(fullName="Priya Sharma", passportNumber="Z1234567", relationship="Wife", occupation="Homemaker", issuePlace="Mumbai", issueDate="2021-02-10")
            ]
        )
        temp_file = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        temp_file.close()
        try:
            build_docx_cover_letter({}, temp_file.name, app_data=app)
            doc = Document(temp_file.name)
            doc_text = "\n".join(p.text for p in doc.paragraphs)
            self.assertIn("Rahul Sharma", doc_text)
            self.assertIn("Z9876543", doc_text)
            self.assertIn("The Consulate General of France", doc_text)
            self.assertIn("Tech Solutions Ltd", doc_text)
            self.assertIn("Priya Sharma", doc_text)
            # Ensure no brackets remain
            import re
            brackets = re.findall(r'\[[A-Za-z0-9 /,\.\-–—\?]+\]', doc_text)
            self.assertEqual(len(brackets), 0, f"Residual brackets found: {brackets}")
        finally:
            if os.path.exists(temp_file.name):
                os.remove(temp_file.name)

    def test_09_authoritative_docx_japan_template_fidelity(self):
        """Verify Japan cover letter DOCX populates dynamic hotel table and family details."""
        app = MasterApplicationData(
            selectedCountry="Japan",
            selectedTemplateId="japan",
            personal=ApplicantPersonal(fullName="Aarav Mehta", title="Mr.", nationality="Indian", gender="Male"),
            passport=ApplicantPassport(passportNumber="J1234567", issuePlace="Delhi", issueDate="2022-03-10"),
            travel=TravelDetails(travelStartDate="2026-11-01", travelEndDate="2026-11-10", numberOfNights=9, purposeOfTravel="Holiday"),
            financial=FinancialInfo(tripSponsor="Self-funded"),
            employment=ApplicantEmployment(employerName="Mehta Enterprises", jobTitle="Director"),
            contact=ApplicantContact(mobileNumber="+91 9876500000", emailAddress="aarav@mehta.com"),
            travellers=[
                Traveller(fullName="Kavita Mehta", passportNumber="J7654321", relationship="Wife", occupation="Architect", issuePlace="Delhi", issueDate="2022-03-10")
            ],
            hotels=[
                Hotel(hotelName="Park Hyatt Tokyo", city="Tokyo", checkInDate="2026-11-01", checkOutDate="2026-11-06", contactNumber="+81 3-5322-1234"),
                Hotel(hotelName="Kyoto Hotel Okura", city="Kyoto", checkInDate="2026-11-06", checkOutDate="2026-11-10", contactNumber="+81 75-211-5111")
            ]
        )
        temp_file = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        temp_file.close()
        try:
            build_docx_cover_letter({}, temp_file.name, app_data=app)
            doc = Document(temp_file.name)
            self.assertEqual(len(doc.tables), 1)
            tbl = doc.tables[0]
            self.assertEqual(len(tbl.rows), 3) # Header + 2 hotels
            self.assertIn("Park Hyatt Tokyo", tbl.rows[1].cells[0].text)
            self.assertIn("Kyoto Hotel Okura", tbl.rows[2].cells[0].text)
            doc_text = "\n".join(p.text for p in doc.paragraphs)
            self.assertIn("Consulate General of Japan", doc_text)
            self.assertIn("Kavita Mehta", doc_text)
            # Ensure no demo hotel data leaks
            self.assertNotIn("RIHGA Royal", doc_text)
            self.assertNotIn("Mitsui Garden", doc_text)
        finally:
            if os.path.exists(temp_file.name):
                os.remove(temp_file.name)

    def test_10_authoritative_docx_singapore_template_fidelity(self):
        """Verify Singapore cover letter DOCX populates 5-column passenger table and removes leftover draft markers."""
        app = MasterApplicationData(
            selectedCountry="Singapore",
            selectedTemplateId="singapore",
            personal=ApplicantPersonal(fullName="Deepak Chopra", title="Mr.", nationality="Indian", gender="Male"),
            passport=ApplicantPassport(passportNumber="S9988776", issuePlace="Bengaluru", issueDate="2021-08-12"),
            travel=TravelDetails(travelStartDate="2026-12-20", travelEndDate="2026-12-28", numberOfNights=8, purposeOfTravel="Tourism"),
            financial=FinancialInfo(tripSponsor="Self-funded"),
            employment=ApplicantEmployment(jobTitle="Principal Architect", employerName="Chopra & Associates"),
            contact=ApplicantContact(mobileNumber="+91 9900011223", emailAddress="deepak@chopra.org"),
            travellers=[
                Traveller(fullName="Suman Chopra", passportNumber="S1122334", relationship="Wife", occupation="Manager")
            ],
            hotels=[
                Hotel(hotelName="Marina Bay Sands", address="10 Bayfront Avenue", city="Singapore", contactNumber="+65 6688 8868")
            ]
        )
        temp_file = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        temp_file.close()
        try:
            build_docx_cover_letter({}, temp_file.name, app_data=app)
            doc = Document(temp_file.name)
            self.assertEqual(len(doc.tables), 1)
            tbl = doc.tables[0]
            self.assertEqual(len(tbl.rows), 3) # Header + 2 passengers
            self.assertEqual(tbl.rows[1].cells[1].text, "Deepak Chopra")
            self.assertEqual(tbl.rows[1].cells[3].text, "Self")
            self.assertEqual(tbl.rows[2].cells[1].text, "Suman Chopra")
            self.assertEqual(tbl.rows[2].cells[3].text, "Wife")
            doc_text = "\n".join(p.text for p in doc.paragraphs)
            self.assertIn("The Consulate General of the Republic of Singapore", doc_text)
            self.assertNotIn("Passanger 1 Name", doc_text)
            self.assertNotIn("P1’s Email ID", doc_text)
        finally:
            if os.path.exists(temp_file.name):
                os.remove(temp_file.name)

    def test_11_anti_leak_guard_strict_enforcement(self):
        """Verify that any forbidden demo string raises an immediate ValueError."""
        from backend.document_generator.docx_generator import check_anti_leak_guard
        # Legitimate text passes
        check_anti_leak_guard("Legitimate text for client Amit Patel visiting France", None)
        # Forbidden text raises
        with self.assertRaises(ValueError):
            check_anti_leak_guard("Booking confirmed at RIHGA Royal hotel", None)


if __name__ == "__main__":
    unittest.main()

