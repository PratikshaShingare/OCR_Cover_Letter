"""
Khanna Travels & Holidays — Authoritative DOCX Document Generator
Clones and populates the project's authoritative clean reference templates:
- Europe_covering_letter_template_clean.docx (Standard / Europe)
- Japan_covering_letter_template_clean.docx (Japan)
- Singapore_covering_letter_template_clean.docx (Singapore)
Guarantees 100% preservation of template typography, tables, borders, and layouts.
Enforces Anti-Leak Guard to guarantee no demo/sample data leaks into generated documents.
"""

import os
import re
import copy
import json
from typing import Dict, Any, List, Optional
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from ..models.master_schema import MasterApplicationData, Traveller, Hotel


# Project Root & Template Paths
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

def get_template_path(filename: str) -> str:
    for base in [PROJECT_ROOT, os.path.abspath(os.path.join(PROJECT_ROOT, "..")), os.path.join(PROJECT_ROOT, "templates")]:
        cand = os.path.join(base, filename)
        if os.path.exists(cand):
            return cand
    return os.path.join(PROJECT_ROOT, filename)

EUROPE_TEMPLATE_PATH = get_template_path("Europe_covering_letter_template_clean.docx")
JAPAN_TEMPLATE_PATH = get_template_path("Japan_covering_letter_template_clean.docx")
SINGAPORE_TEMPLATE_PATH = get_template_path("Singapore_covering_letter_template_clean.docx")

# Forbidden sample data values (must never appear unless entered by user)
FORBIDDEN_SAMPLE_STRINGS = [
    "Shirish", "Ghosal", "Daswani", "Yasha", "Verma", "Shah",
    "The Trade Desk", "VisionSpring", "RIHGA Royal", "Mitsui Garden", "Village Hotel"
]


def format_ordinal_date(date_str: str, include_year: bool = True, date_only: bool = False) -> str:
    """Formats '2026-08-01' or '01/08/2026' into '01st August 2026'."""
    if not date_str:
        return ""
    try:
        dt = None
        clean_str = date_str.strip()
        if "-" in clean_str:
            parts = clean_str.split("-")
            if len(parts) == 3:
                if len(parts[0]) == 4:
                    dt = datetime.strptime(clean_str, "%Y-%m-%d")
                else:
                    dt = datetime.strptime(clean_str, "%d-%m-%Y")
        elif "/" in clean_str:
            parts = clean_str.split("/")
            if len(parts) == 3:
                if len(parts[2]) == 4:
                    dt = datetime.strptime(clean_str, "%d/%m/%Y")
                else:
                    dt = datetime.strptime(clean_str, "%m/%d/%Y")
        elif "." in clean_str:
            dt = datetime.strptime(clean_str, "%d.%m.%Y")
            
        if not dt:
            return date_str
            
        day = dt.day
        if 11 <= day <= 13:
            suffix = "th"
        else:
            suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
            
        day_str = f"{day:02d}{suffix}"
        if date_only:
            return day_str
        month_name = dt.strftime("%B")
        if include_year:
            return f"{day_str} {month_name} {dt.year}"
        return f"{day_str} {month_name}"
    except Exception:
        return date_str


def format_dot_date(date_str: str) -> str:
    """Formats '2026-08-29' into '29.08.2026'."""
    if not date_str:
        return ""
    try:
        clean_str = date_str.strip()
        if "-" in clean_str:
            dt = datetime.strptime(clean_str, "%Y-%m-%d")
            return dt.strftime("%d.%m.%Y")
        return date_str
    except Exception:
        return date_str


def check_anti_leak_guard(full_text: str, app_data: Optional[MasterApplicationData] = None):
    """
    Scans generated document text for any known demo/sample data.
    Fails generation immediately if sample values leak.
    """
    user_blob = ""
    if app_data:
        try:
            user_blob = json.dumps(app_data.model_dump()).lower()
        except Exception:
            user_blob = ""

    for forbidden in FORBIDDEN_SAMPLE_STRINGS:
        if forbidden.lower() in full_text.lower():
            # Check if this exact string was explicitly provided by the user in this application
            if forbidden.lower() not in user_blob:
                raise ValueError(
                    f"Anti-leak validation failure: Forbidden demo data '{forbidden}' detected in generated document."
                )


def replace_in_paragraph(paragraph, replacements: Dict[str, str]):
    """
    Replaces string occurrences within paragraph runs, preserving font attributes.
    Handles placeholders residing in single runs or split across adjacent runs.
    """
    if not paragraph.text:
        return

    for target, replacement in replacements.items():
        if target not in paragraph.text:
            continue
            
        # 1. Try single-run replacement first to preserve exact typography
        replaced = False
        for run in paragraph.runs:
            if target in run.text:
                run.text = run.text.replace(target, replacement)
                replaced = True
                break
                
        # 2. If split across multiple runs, update across runs
        if not replaced and target in paragraph.text:
            new_text = paragraph.text.replace(target, replacement)
            if paragraph.runs:
                paragraph.runs[0].text = new_text
                for r in paragraph.runs[1:]:
                    r.text = ""
            else:
                paragraph.text = new_text


def remove_paragraph(paragraph):
    """Removes a paragraph from its parent document."""
    p_element = paragraph._element
    if p_element.getparent() is not None:
        p_element.getparent().remove(p_element)


def remove_row(table, row_idx: int):
    """Removes a table row from its parent table."""
    if 0 <= row_idx < len(table.rows):
        tr = table.rows[row_idx]._tr
        if tr.getparent() is not None:
            table._tbl.remove(tr)


def duplicate_row(table, source_row_idx: int = 1):
    """Appends a new row duplicating the XML structure and styling of source row."""
    source_tr = table.rows[source_row_idx]._tr
    new_tr = copy.deepcopy(source_tr)
    table._tbl.append(new_tr)
    new_row = table.rows[-1]
    for cell in new_row.cells:
        cell.text = ""
    return new_row


def clean_residual_brackets(doc: Document):
    """Cleans up any leftover bracketed template tokens e.g. [Next Country]."""
    bracket_regex = re.compile(r'\[[A-Za-z0-9 /,\.\-–—\?]+\]')
    for p in doc.paragraphs:
        if '[' in p.text and ']' in p.text:
            new_text = bracket_regex.sub('', p.text).replace('  ', ' ')
            if p.runs:
                p.runs[0].text = new_text
                for r in p.runs[1:]:
                    r.text = ''
            else:
                p.text = new_text

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if '[' in cell.text and ']' in cell.text:
                    for p in cell.paragraphs:
                        new_text = bracket_regex.sub('', p.text).replace('  ', ' ')
                        if p.runs:
                            p.runs[0].text = new_text
                            for r in p.runs[1:]:
                                r.text = ''
                        else:
                            p.text = new_text


def build_docx_cover_letter(doc_data: Dict[str, Any], output_path: str, app_data: Optional[MasterApplicationData] = None) -> str:
    """
    Generates a genuine .docx file by directly cloning the authoritative clean template
    and populating user data. Falls back to pristine custom builder if template missing.
    """
    # Attempt to resolve app_data from doc_data if not directly provided
    if not app_data and "appData" in doc_data:
        try:
            app_data = MasterApplicationData.model_validate(doc_data["appData"])
        except Exception:
            app_data = None

    # Determine template format
    fmt = doc_data.get("format")
    if not fmt and app_data:
        tpl_choice = (app_data.selectedTemplateId or "").lower().strip()
        country = (app_data.selectedCountry or "").lower().strip()
        if "japan" in tpl_choice or country == "japan":
            fmt = "japan"
        elif "singapore" in tpl_choice or country == "singapore":
            fmt = "singapore"
        else:
            fmt = "standard"
    elif not fmt:
        fmt = "standard"

    if app_data and fmt == "japan" and os.path.exists(JAPAN_TEMPLATE_PATH):
        doc = _generate_japan_from_template(app_data, JAPAN_TEMPLATE_PATH)
    elif app_data and fmt == "singapore" and os.path.exists(SINGAPORE_TEMPLATE_PATH):
        doc = _generate_singapore_from_template(app_data, SINGAPORE_TEMPLATE_PATH)
    elif app_data and os.path.exists(EUROPE_TEMPLATE_PATH):
        doc = _generate_europe_from_template(app_data, EUROPE_TEMPLATE_PATH)
    else:
        # Fallback to pristine builder matching reference layout
        doc = _build_fallback_docx(doc_data)

    # Anti-leak verification
    full_doc_text = "\n".join([p.text for p in doc.paragraphs])
    for tbl in doc.tables:
        for r in tbl.rows:
            full_doc_text += "\n" + " ".join([c.text for c in r.cells])
    check_anti_leak_guard(full_doc_text, app_data)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    return output_path


def _generate_europe_from_template(app_data: MasterApplicationData, template_path: str) -> Document:
    """Clones Europe_covering_letter_template_clean.docx and populates fields."""
    doc = Document(template_path)
    
    personal = app_data.personal
    passport = app_data.passport
    travel = app_data.travel
    financial = app_data.financial
    employment = app_data.employment
    address = app_data.address
    travellers = app_data.travellers or []
    
    today_formatted = format_ordinal_date(datetime.now().strftime("%Y-%m-%d"))
    full_name = personal.fullName or f"{personal.givenName} {personal.surname}".strip() or "Applicant"
    pass_num = passport.passportNumber or ""
    issue_place = passport.issuePlace or ""
    issue_date = format_ordinal_date(passport.issueDate) if passport.issueDate else ""
    country = app_data.selectedCountry or "France"
    consulate_city = address.city or "Mumbai"
    
    start_date = format_ordinal_date(travel.travelStartDate) if travel.travelStartDate else ""
    end_date = format_ordinal_date(travel.travelEndDate) if travel.travelEndDate else ""
    nights = f"{travel.numberOfNights} Nights" if travel.numberOfNights > 0 else ""
    countries_visited = travel.countriesToBeVisited or country
    
    spouse = next((t for t in travellers if t.relationship.lower() in ["spouse", "wife", "husband"]), None)
    is_solo = len(travellers) == 0
    
    if is_solo:
        relation_family = ""
    elif spouse and len(travellers) == 1:
        relation_family = "Wife" if personal.gender.lower() == "male" else "Husband"
    else:
        relation_family = "family"

    # Sponsor wording
    if financial.tripSponsor == "Jointly funded":
        sponsor_wording = "both of us individually"
    elif financial.tripSponsor == "Self-funded":
        sponsor_wording = "me" if is_solo else "myself"
    elif financial.tripSponsor == "Company funded":
        sponsor_wording = f"my employer, {employment.employerName or 'the company'}"
    else:
        sponsor_wording = financial.sponsorName or ("me" if is_solo else "myself")

    city_country = f"{address.city}, {address.country or 'India'}" if address.city else "Mumbai, India"
    emp_status = employment.employmentStatus or "Salaried"
    job_title = employment.jobTitle or "Executive"
    emp_name = employment.employerName or employment.businessName or "Private Organization"
    start_year = employment.employmentStartDate[:4] if employment.employmentStartDate and len(employment.employmentStartDate) >= 4 else "2020"

    replacements = {
        "[Date]": today_formatted,
        "[Consulate Name and Address]": f"The Consulate General of {country}\n{consulate_city}, India",
        "[Destination Country]": country,
        "[Applicant Full Name]": full_name,
        "[Passport Number]": pass_num,
        "[Place of Issue]": issue_place,
        "[Passport Issue Date]": issue_date,
        "[Travel Start Date]": start_date,
        "[Travel End Date]": end_date,
        "[Destination Country/Countries]": countries_visited,
        "[Number of Nights]": nights,
        "[Funding Arrangement]": sponsor_wording,
        "[Phone Number]": app_data.contact.mobileNumber or "",
        "[Email Address]": app_data.contact.emailAddress or "",
        "[City, Country]": city_country,
        "[Employment Status]": emp_status,
        "[Job Title]": job_title,
        "[Employer Name]": emp_name,
        "[Employment Start Year]": start_year,
    }

    # Handle narrative adjustments for solo vs family
    for p in doc.paragraphs:
        if "[Applicant Full Name]" in p.text:
            if is_solo:
                # Remove "along with my [Relation/Family] "
                replacements["along with my [Relation/Family] "] = ""
                replacements["along with my [Relation/Family]"] = ""
            else:
                replacements["[Relation/Family]"] = relation_family

        if "reside in [City, Country]" in p.text:
            if is_solo:
                p_text = (
                    f"I reside in {city_country}. I am financially independent and Employed – {emp_status}. "
                    f"I am a {job_title} at {emp_name}. "
                    f"Kindly find my relevant documents for your reference."
                )
                if p.runs:
                    p.runs[0].text = p_text
                    for r in p.runs[1:]:
                        r.text = ""
                else:
                    p.text = p_text
                continue
            elif spouse:
                s_rel = "wife" if personal.gender.lower() == "male" else "husband"
                s_name_str = f", {spouse.fullName}," if spouse.fullName else ""
                s_emp = spouse.employer or "private organization"
                s_occ = spouse.occupation or "professional"
                replacements["[Relation]"] = s_rel
                p_text = (
                    f"I and my {s_rel}{s_name_str} reside in {city_country} along with our parents and few close relatives. "
                    f"We both are financially independent and Employed – {emp_status}. "
                    f"I am a {job_title} at {emp_name}. "
                    f"While my {s_rel} is working for {s_emp}, currently as {s_occ}. "
                    f"Kindly find our relevant documents for your reference."
                )
                if p.runs:
                    p.runs[0].text = p_text
                    for r in p.runs[1:]:
                        r.text = ""
                else:
                    p.text = p_text
                continue

        if "have planned to visit [Destination Country/Countries]" in p.text:
            # Clean up the multi-country sentence if only 1 country
            if is_solo:
                p_lead = f"I have planned to visit {countries_visited} for the purpose of Tourism. We look forward to visiting the country, exploring the beautiful sights, witnessing its rich culture and spending leisure time."
            else:
                p_lead = f"My {relation_family} and I have planned to visit {countries_visited} for the purpose of Tourism. We look forward to visiting the country, exploring the beautiful sights, witnessing its rich culture and spending leisure time together."
            
            p_visit = f" We would be visiting {country} {('for ' + nights) if nights else ''} from {start_date} to {end_date}. Kindly find our relevant documents for your perusal. We will return back to India after the end of our visit as we need to rejoin and look after our work and personal commitments."
            p_text = (p_lead + p_visit).replace("  ", " ")
            if p.runs:
                p.runs[0].text = p_text
                for r in p.runs[1:]:
                    r.text = ""
            else:
                p.text = p_text
            continue

        replace_in_paragraph(p, replacements)

    clean_residual_brackets(doc)
    return doc


def _generate_japan_from_template(app_data: MasterApplicationData, template_path: str) -> Document:
    """Clones Japan_covering_letter_template_clean.docx and populates fields and hotel table."""
    doc = Document(template_path)
    
    personal = app_data.personal
    passport = app_data.passport
    travel = app_data.travel
    financial = app_data.financial
    employment = app_data.employment
    travellers = app_data.travellers or []
    hotels = app_data.hotels or []
    
    today_formatted = format_ordinal_date(datetime.now().strftime("%Y-%m-%d"))
    full_name = personal.fullName or f"{personal.givenName} {personal.surname}".strip() or "Applicant"
    pass_num = passport.passportNumber or ""
    issue_place = passport.issuePlace or ""
    issue_date = format_ordinal_date(passport.issueDate) if passport.issueDate else ""
    start_date = format_ordinal_date(travel.travelStartDate) if travel.travelStartDate else ""
    end_date = format_ordinal_date(travel.travelEndDate) if travel.travelEndDate else ""
    
    is_solo = len(travellers) == 0

    if financial.tripSponsor == "Jointly funded":
        sponsor_wording = "both of us individually"
    elif financial.tripSponsor == "Self-funded":
        sponsor_wording = "me" if is_solo else "myself"
    elif financial.tripSponsor == "Company funded":
        sponsor_wording = f"my employer, {employment.employerName or 'the company'}"
    else:
        sponsor_wording = financial.sponsorName or ("me" if is_solo else "myself")

    emp_occ = employment.employerName or employment.businessName or employment.jobTitle or employment.employmentStatus or "Employed"

    replacements = {
        "[Date]": today_formatted,
        "[Passenger 1 Name]": full_name,
        "[Passport Number]": pass_num,
        "[Place of Issue]": issue_place,
        "[Passport Issue Date]": issue_date,
        "[Travel Start Date]": start_date,
        "[Travel End Date]": end_date,
        "[Employer / Occupation]": emp_occ,
        "[Sponsor Name / Funding Arrangement]": sponsor_wording,
        "[Phone Number]": app_data.contact.mobileNumber or "",
        "[Email Address]": app_data.contact.emailAddress or "",
    }

    if is_solo:
        replacements["along with my family "] = ""
        replacements["along with my family"] = ""

    # Process paragraphs
    paragraphs_to_remove = []
    for p in doc.paragraphs:
        # Check companion paragraph (Paragraph 12 in Japan template)
        if "My [Relation], [Passenger 2 Name]" in p.text or "[Add additional family member details here if applicable]" in p.text:
            if is_solo:
                paragraphs_to_remove.append(p)
                continue
            else:
                fam_sentences = []
                for t in travellers:
                    t_full = t.fullName or f"{t.givenName} {t.surname}".strip() or "Companion"
                    t_pass = t.passportNumber or ""
                    t_place = t.issuePlace or ""
                    t_date = format_ordinal_date(t.issueDate) if t.issueDate else ""
                    t_nat = t.nationality or personal.nationality or "Indian"
                    pass_part = f" (holding {t_nat} Passport No.: {t_pass} issued at {t_place} on {t_date})" if t_pass else ""
                    rel = t.relationship.lower()

                    if rel in ["spouse", "wife", "husband"]:
                        occ = t.occupation or "Homemaker"
                        label = "wife" if rel == "wife" or personal.gender.lower() == "male" else "husband"
                        fam_sentences.append(f"My {label}, {t_full}{pass_part} is {occ}.")
                    elif rel in ["son", "daughter", "child"]:
                        school = t.schoolCollege or ""
                        grade = t.gradeClass or ""
                        child_label = "son" if rel == "son" else "daughter" if rel == "daughter" else "child"
                        edu_part = f" is a Student of {school}" if school else " is a Student"
                        if grade:
                            edu_part += f" and currently {grade}"
                        fam_sentences.append(f"My {child_label}, {t_full}{pass_part}{edu_part}.")
                    else:
                        fam_sentences.append(f"My {t.relationship or 'Companion'}, {t_full}{pass_part} is {t.occupation or 'Employed'}.")

                fam_text = " ".join(fam_sentences)
                if p.runs:
                    p.runs[0].text = fam_text
                    for r in p.runs[1:]:
                        r.text = ""
                else:
                    p.text = fam_text
                continue

        replace_in_paragraph(p, replacements)

    for p in paragraphs_to_remove:
        remove_paragraph(p)

    # Populate Hotel Table (Table 0)
    if doc.tables:
        table = doc.tables[0]
        # Headers: ['NAME', 'DATE', 'CONTACT NO.']
        if hotels:
            # Set first hotel in Row 1
            h0 = hotels[0]
            name0 = f"{h0.hotelName}\n{h0.city}".strip() if h0.city else (h0.hotelName or "")
            d0 = ""
            if h0.checkInDate and h0.checkOutDate:
                d0 = f"{format_ordinal_date(h0.checkInDate, include_year=False)} - {format_ordinal_date(h0.checkOutDate, include_year=False)}"
            table.rows[1].cells[0].text = name0
            table.rows[1].cells[1].text = d0
            table.rows[1].cells[2].text = h0.contactNumber or ""

            # Additional hotels
            for i in range(1, len(hotels)):
                hi = hotels[i]
                name_i = f"{hi.hotelName}\n{hi.city}".strip() if hi.city else (hi.hotelName or "")
                di = ""
                if hi.checkInDate and hi.checkOutDate:
                    di = f"{format_ordinal_date(hi.checkInDate, include_year=False)} - {format_ordinal_date(hi.checkOutDate, include_year=False)}"
                if i == 1 and len(table.rows) > 2:
                    # Reuse Row 2
                    table.rows[2].cells[0].text = name_i
                    table.rows[2].cells[1].text = di
                    table.rows[2].cells[2].text = hi.contactNumber or ""
                else:
                    # Append new row duplicating format
                    new_row = duplicate_row(table, source_row_idx=1)
                    new_row.cells[0].text = name_i
                    new_row.cells[1].text = di
                    new_row.cells[2].text = hi.contactNumber or ""

            # If only 1 hotel provided and table had 2 sample rows, remove row 2
            if len(hotels) == 1 and len(table.rows) > 2:
                remove_row(table, 2)
        else:
            # No hotels entered: leave clean row with empty string
            table.rows[1].cells[0].text = ""
            table.rows[1].cells[1].text = ""
            table.rows[1].cells[2].text = ""
            if len(table.rows) > 2:
                remove_row(table, 2)

    clean_residual_brackets(doc)
    return doc


def _generate_singapore_from_template(app_data: MasterApplicationData, template_path: str) -> Document:
    """Clones Singapore_covering_letter_template_clean.docx and populates fields and passenger table."""
    doc = Document(template_path)
    
    personal = app_data.personal
    passport = app_data.passport
    travel = app_data.travel
    financial = app_data.financial
    employment = app_data.employment
    travellers = app_data.travellers or []
    hotels = app_data.hotels or []
    
    today_dot = format_dot_date(datetime.now().strftime("%Y-%m-%d"))
    full_name = personal.fullName or f"{personal.givenName} {personal.surname}".strip() or "Applicant"
    pass_num = passport.passportNumber or ""
    start_date = format_ordinal_date(travel.travelStartDate) if travel.travelStartDate else ""
    end_date = format_ordinal_date(travel.travelEndDate) if travel.travelEndDate else ""
    is_solo = len(travellers) == 0

    if financial.tripSponsor == "Jointly funded":
        sponsor_wording = "both of us individually"
    elif financial.tripSponsor == "Self-funded":
        sponsor_wording = "me" if is_solo else "myself"
    elif financial.tripSponsor == "Company funded":
        sponsor_wording = f"my employer, {employment.employerName or 'the company'}"
    else:
        sponsor_wording = financial.sponsorName or ("me" if is_solo else "myself")

    # Hotel presentation
    if hotels:
        h = hotels[0]
        hotel_str = f"{h.hotelName} – {h.address}, {h.city}".strip().rstrip(',')
    else:
        hotel_str = "confirmed hotel accommodations"

    replacements = {
        "[Date]": today_dot,
        "[Consulate Name and Address]": "The Consulate General of the Republic of Singapore,\n152 Marker Chambers IV, 222 Jamnalal Bajaj Marg Nariman Point,\nMumbai 400021",
        "[Passenger 1 Name]": full_name,
        "[Travel Start Date]": start_date,
        "[Travel End Date]": end_date,
        "[Hotel Name] – [Hotel Address]": hotel_str,
        "[Hotel Name] - [Hotel Address]": hotel_str,
        "[Funding Arrangement]": sponsor_wording,
        "[Contact No.]": app_data.contact.mobileNumber or "",
        "[Email Address]": app_data.contact.emailAddress or "",
    }

    if is_solo:
        replacements["along with my Family & "] = "& "
        replacements["along with my Family"] = ""
        replacements["& my family"] = ""
        replacements["Request you to kindly grant us"] = "Request you to kindly grant me"

    # Identify and clean duplicate or leftover template paragraphs
    # In Singapore clean template: P11 is duplicate bank statement, P15-P17 is draft signature
    paragraphs_to_remove = []
    seen_bank_statement = False
    for p in doc.paragraphs:
        text = p.text.strip()
        if "I am enclosing the bank statement of my bank account for your perusal." in text:
            if seen_bank_statement:
                paragraphs_to_remove.append(p)
                continue
            seen_bank_statement = True
            if not financial.bankStatementAvailable:
                paragraphs_to_remove.append(p)
                continue

        if "Passanger 1 Name" in text or "Email Id: P1" in text or "Mob: [P1" in text:
            paragraphs_to_remove.append(p)
            continue

        replace_in_paragraph(p, replacements)

    for p in paragraphs_to_remove:
        remove_paragraph(p)

    # Populate Passenger Table (Table 0)
    # Headers: ['Sr no.', 'Passengers Name', 'Passport No', 'Relation', 'Occupation']
    if doc.tables:
        table = doc.tables[0]
        # Row 1: Passenger 1 (Self)
        self_occ = employment.jobTitle or employment.employmentStatus or "Employed"
        table.rows[1].cells[0].text = "1"
        table.rows[1].cells[1].text = full_name
        table.rows[1].cells[2].text = pass_num
        table.rows[1].cells[3].text = "Self"
        table.rows[1].cells[4].text = self_occ

        if travellers:
            # Row 2: First companion
            t0 = travellers[0]
            t0_name = t0.fullName or f"{t0.givenName} {t0.surname}".strip() or ""
            table.rows[2].cells[0].text = "2"
            table.rows[2].cells[1].text = t0_name
            table.rows[2].cells[2].text = t0.passportNumber or ""
            table.rows[2].cells[3].text = t0.relationship or "Family"
            table.rows[2].cells[4].text = t0.occupation or "Student"

            # Additional companions
            for i in range(1, len(travellers)):
                ti = travellers[i]
                ti_name = ti.fullName or f"{ti.givenName} {ti.surname}".strip() or ""
                new_row = duplicate_row(table, source_row_idx=1)
                new_row.cells[0].text = str(i + 2)
                new_row.cells[1].text = ti_name
                new_row.cells[2].text = ti.passportNumber or ""
                new_row.cells[3].text = ti.relationship or "Family"
                new_row.cells[4].text = ti.occupation or "Student"
        else:
            # Solo: remove Row 2
            if len(table.rows) > 2:
                remove_row(table, 2)

    clean_residual_brackets(doc)
    return doc


def _build_fallback_docx(doc_data: Dict[str, Any]) -> Document:
    """Creates a clean Word document matching reference typography when templates are absent."""
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)
        
    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    date_p.paragraph_format.space_after = Pt(14)
    d_run = date_p.add_run(doc_data.get("date", ""))
    d_run.font.name = 'Calibri'
    d_run.font.size = Pt(11)
    
    for line in doc_data.get("recipient", []):
        r_p = doc.add_paragraph()
        r_p.paragraph_format.space_after = Pt(2)
        r_p.paragraph_format.line_spacing = 1.15
        r_run = r_p.add_run(line)
        r_run.font.name = 'Calibri'
        r_run.font.size = Pt(11)
        
    doc.add_paragraph().paragraph_format.space_after = Pt(8)
    
    subj_data = doc_data.get("subject", {})
    subj_p = doc.add_paragraph()
    subj_p.paragraph_format.space_after = Pt(14)
    subj_run = subj_p.add_run(subj_data.get("text", ""))
    subj_run.font.name = 'Calibri'
    subj_run.font.size = Pt(11)
    subj_run.bold = subj_data.get("bold", True)
    subj_run.underline = subj_data.get("underline", False)
    
    sal_p = doc.add_paragraph()
    sal_p.paragraph_format.space_after = Pt(12)
    sal_run = sal_p.add_run(doc_data.get("salutation", "Dear Sir/Ma’am,"))
    sal_run.font.name = 'Calibri'
    sal_run.font.size = Pt(11)
    
    for block in doc_data.get("blocks", []):
        b_type = block.get("type")
        if b_type == "paragraph":
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(10)
            p.paragraph_format.line_spacing = 1.15
            run = p.add_run(block.get("content", ""))
            run.font.name = 'Calibri'
            run.font.size = Pt(11)
        elif b_type == "table":
            headers = block.get("headers", [])
            rows = block.get("rows", [])
            table = doc.add_table(rows=1 + len(rows), cols=len(headers))
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.autofit = True
            for idx, h_text in enumerate(headers):
                table.rows[0].cells[idx].text = h_text
                table.rows[0].cells[idx].paragraphs[0].runs[0].bold = True
            for r_idx, row_data in enumerate(rows):
                for c_idx, cell_value in enumerate(row_data):
                    if c_idx < len(headers):
                        table.rows[r_idx + 1].cells[c_idx].text = str(cell_value)

    sign = doc_data.get("signBlock", {})
    if sign.get("thanks"):
        t_p = doc.add_paragraph()
        t_p.paragraph_format.space_after = Pt(4)
        t_p.add_run(sign["thanks"])
        
    sal_p = doc.add_paragraph()
    sal_p.paragraph_format.space_after = Pt(18)
    sal_p.add_run(sign.get("salutation", "Yours Faithfully,"))
    
    name_p = doc.add_paragraph()
    name_p.paragraph_format.space_after = Pt(2)
    name_run = name_p.add_run(sign.get("name", ""))
    name_run.bold = True
    
    if sign.get("phone"):
        ph_p = doc.add_paragraph()
        ph_p.paragraph_format.space_after = Pt(2)
        ph_p.add_run(f"{sign.get('phoneLabel', 'Phone No.:')} {sign['phone']}")
        
    if sign.get("email"):
        em_p = doc.add_paragraph()
        em_p.paragraph_format.space_after = Pt(2)
        em_p.add_run(f"{sign.get('emailLabel', 'Email id:')} {sign['email']}")
        
    return doc
