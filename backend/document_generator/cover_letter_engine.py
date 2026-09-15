"""
Khanna Travels & Holidays — Cover Letter Template Engine
Merges Verified Master Client Data with Country Templates (Base + Country Overrides).
Strictly uses current application data. Never injects sample client names, employers, or hotels.
"""

import json
import os
from typing import Dict, Any, List, Optional
from datetime import datetime
from ..models.master_schema import MasterApplicationData, Traveller, Hotel
from .docx_generator import check_anti_leak_guard


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


class CoverLetterEngine:
    def __init__(self, templates_dir: Optional[str] = None):
        if not templates_dir:
            curr_dir = os.path.dirname(os.path.abspath(__file__))
            self.templates_dir = os.path.abspath(os.path.join(curr_dir, "..", "..", "templates"))
        else:
            self.templates_dir = templates_dir
            
    def load_template_config(self, country_or_format: str) -> Dict[str, Any]:
        """Loads base template and merges country override config."""
        base_path = os.path.join(self.templates_dir, "base", "tourism", "template.json")
        with open(base_path, "r", encoding="utf-8") as f:
            base_data = json.load(f)
            
        target = (country_or_format or "").lower().strip()
        if "japan" in target:
            target_key = "japan"
        elif "singapore" in target:
            target_key = "singapore"
        else:
            target_key = None
            
        if target_key:
            override_file = os.path.join(self.templates_dir, "country_overrides", target_key, "tourism.json")
            if os.path.exists(override_file):
                with open(override_file, "r", encoding="utf-8") as f:
                    override_data = json.load(f)
                return self._deep_merge(base_data, override_data)
        return base_data

    def _deep_merge(self, base: dict, override: dict) -> dict:
        result = dict(base)
        for k, v in override.items():
            if k in result and isinstance(result[k], dict) and isinstance(v, dict):
                result[k] = self._deep_merge(result[k], v)
            else:
                result[k] = v
        return result

    def generate_document(self, app_data: MasterApplicationData) -> Dict[str, Any]:
        """
        Generates structured blocks and rich HTML for cover letter
        strictly using the current verified master data and selected format (Standard, Japan, Singapore).
        """
        tpl_choice = (app_data.selectedTemplateId or "").lower().strip()
        country = app_data.selectedCountry or "France"
        
        # Decide format: Standard, Japan, Singapore
        if "japan" in tpl_choice or (tpl_choice in ["standard", "france_tourism", ""] and country.lower() == "japan"):
            fmt = "japan"
        elif "singapore" in tpl_choice or (tpl_choice in ["standard", "france_tourism", ""] and country.lower() == "singapore"):
            fmt = "singapore"
        else:
            fmt = "standard"

        template = self.load_template_config(fmt)
        config = template.get("config", {})
        
        personal = app_data.personal
        passport = app_data.passport
        travel = app_data.travel
        financial = app_data.financial
        employment = app_data.employment
        travellers = app_data.travellers or []
        hotels = app_data.hotels or []
        
        # Date header
        today_iso = datetime.now().strftime("%Y-%m-%d")
        date_cfg = config.get("dateHeader", {})
        date_format = date_cfg.get("format", "ordinal")
        date_prefix = date_cfg.get("prefix", "Date- ")
        
        if date_format == "dot":
            formatted_today = date_prefix + format_dot_date(today_iso)
        else:
            formatted_today = date_prefix + format_ordinal_date(today_iso)
            
        # Applicant Information
        full_name = personal.fullName or f"{personal.title} {personal.givenName} {personal.surname}".strip()
        if not full_name:
            full_name = "Applicant"
            
        pass_num = passport.passportNumber or ""
        issue_place = passport.issuePlace or ""
        issue_date_fmt = format_ordinal_date(passport.issueDate) if passport.issueDate else ""
        nationality = personal.nationality or "Indian"

        # Companions description
        spouse = next((t for t in travellers if t.relationship.lower() == "spouse" or t.relationship.lower() == "wife" or t.relationship.lower() == "husband"), None)
        has_children = any(t.relationship.lower() in ["son", "daughter", "child"] for t in travellers)
        
        if fmt == "japan":
            companions_text = "along with my family" if len(travellers) > 0 else ""
        elif fmt == "singapore":
            companions_text = "along with my Family" if len(travellers) > 0 else ""
        else:
            if spouse and not has_children and len(travellers) == 1:
                rel_word = "Wife" if personal.gender.lower() == "male" else "Husband"
                companions_text = f"along with my {rel_word}"
            elif len(travellers) > 0:
                companions_text = "along with my family"
            else:
                companions_text = ""

        # Trip Dates
        start_fmt = format_ordinal_date(travel.travelStartDate) if travel.travelStartDate else ""
        end_fmt = format_ordinal_date(travel.travelEndDate) if travel.travelEndDate else ""
        
        # Consular Recipient
        recipient_lines = []
        for line in config.get("recipient", []):
            formatted_line = line.replace("{country}", country).replace("{consulateCity}", app_data.address.city or "Mumbai")
            recipient_lines.append(formatted_line)
            
        # Subject
        subj_cfg = config.get("subject", {})
        if fmt == "singapore":
            subject_text = "Sub: Request for issue of Tourism Visa for me & my family" if len(travellers) > 0 else "Sub: Request for issue of Tourism Visa for me"
        else:
            subj_text_template = subj_cfg.get("text", "")
            subject_text = subj_cfg.get("prefix", "Subject: ") + subj_text_template.replace("{country}", country).replace("{purpose}", travel.purposeOfTravel or "Tourism")
        subject_bold = subj_cfg.get("bold", True)
        subject_underline = subj_cfg.get("underline", False)
        
        salutation = config.get("salutation", "Dear Sir/Ma’am,")
        
        # Financial Sponsor
        if financial.tripSponsor == "Jointly funded":
            sponsor_wording = "both of us individually"
        elif financial.tripSponsor == "Self-funded":
            sponsor_wording = "me"
        elif financial.tripSponsor == "Company funded":
            emp_name = employment.employerName or "the company"
            sponsor_wording = f"my employer, {emp_name}"
        else:
            sponsor_wording = financial.sponsorName or "myself"

        blocks: List[Dict[str, Any]] = []
        
        # -------------------------------------------------------------
        # Paragraph 1: Applicant Introduction & Trip Purpose
        # -------------------------------------------------------------
        pass_snippet = f" (holding {nationality} Passport No.: {pass_num} issued at {issue_place} on {issue_date_fmt})" if pass_num else ""
        dates_snippet = f" from {start_fmt} to {end_fmt}" if (start_fmt and end_fmt) else ""
        purpose = travel.purposeOfTravel or ("Holiday" if fmt == "japan" else "Tourism")

        if fmt == "singapore":
            date_part = f" & would like to travel on {start_fmt} and will stay until {end_fmt}" if (start_fmt and end_fmt) else ""
            if companions_text:
                p1_text = f"I, {full_name} planning to visit {country} for {purpose} purposes {companions_text}{date_part}.".replace("  ", " ")
            else:
                p1_text = f"I, {full_name} planning to visit {country} for {purpose} purposes{date_part}.".replace("  ", " ")
        elif fmt == "japan":
            comp_part = f" {companions_text}" if companions_text else ""
            p1_text = f"I, {full_name}{pass_snippet} would be travelling to your esteemed country{comp_part}{dates_snippet} for the purpose of {purpose}.".replace("  ", " ")
        else:
            comp_part = f" {companions_text}" if companions_text else ""
            p1_text = f"I, {full_name}{pass_snippet} would like to visit your admired country{comp_part}{dates_snippet} for the purpose of {purpose}.".replace("  ", " ")
        
        blocks.append({"type": "paragraph", "content": p1_text.strip()})

        # -------------------------------------------------------------
        # Paragraph 2: Employment, Family & Companions
        # -------------------------------------------------------------
        if fmt == "japan":
            if employment.employerName or employment.businessName:
                emp_name = employment.employerName or employment.businessName
                blocks.append({"type": "paragraph", "content": f"I am Employed – {emp_name}."})
            elif employment.jobTitle:
                blocks.append({"type": "paragraph", "content": f"I am employed as {employment.jobTitle}."})

            fam_sentences = []
            for t in travellers:
                t_full = t.fullName or f"{t.givenName} {t.surname}".strip() or "Companion"
                t_pass = t.passportNumber or ""
                t_place = t.issuePlace or ""
                t_date = format_ordinal_date(t.issueDate) if t.issueDate else ""
                t_nat = t.nationality or nationality
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
                    edu_part = f" is a Student of {school}" if school else ""
                    if grade:
                        edu_part += f" and currently {grade}"
                    fam_sentences.append(f"My {child_label} is {t_full}{pass_part}{edu_part}.")
                else:
                    fam_sentences.append(f"Accompanying traveller {t_full}{pass_part} is {t.occupation or 'Employed'}.")

            if fam_sentences:
                blocks.append({"type": "paragraph", "content": " ".join(fam_sentences)})

        elif fmt == "singapore":
            # Formatted Passenger Table
            blocks.append({"type": "paragraph", "content": "The names of the passengers are:"})
            headers = ["Sr no.", "Passengers Name", "Passport No", "Relation", "Occupation"]
            rows = []
            rows.append(["1", full_name, pass_num, "Self", employment.jobTitle or employment.employmentStatus or "Employed"])
            for idx, t in enumerate(travellers, start=2):
                t_name = t.fullName or f"{t.givenName} {t.surname}".strip() or "Passenger"
                rows.append([str(idx), t_name, t.passportNumber or "", t.relationship or "Family", t.occupation or "Student"])
            blocks.append({"type": "table", "headers": headers, "rows": rows, "style": "singapore_passenger"})

        else:
            # Base Europe narrative
            city = app_data.address.city or ""
            state = app_data.address.country or ""
            location_phrase = f"reside in {city}, {state}" if (city and state) else f"reside in {city or state}" if (city or state) else "reside in India"

            if spouse:
                s_name = f"{spouse.title} {spouse.fullName}".strip() if spouse.fullName and not spouse.fullName.startswith(spouse.title) else (spouse.fullName or "Spouse")
                s_pass = spouse.passportNumber or ""
                s_place = spouse.issuePlace or ""
                s_date = format_ordinal_date(spouse.issueDate) if spouse.issueDate else ""
                s_nat = spouse.nationality or nationality
                s_pass_part = f" (holding {s_nat} Passport No.: {s_pass} issued at {s_place} on {s_date})" if s_pass else ""
                
                spouse_occ = ""
                if spouse.employer and spouse.occupation:
                    spouse_occ = f"While my spouse is working for {spouse.employer}, currently as {spouse.occupation}."
                elif spouse.occupation:
                    spouse_occ = f"My spouse is {spouse.occupation}."

                emp_detail = ""
                if employment.employerName and employment.jobTitle:
                    emp_detail = f"I am a {employment.jobTitle} at {employment.employerName}."
                elif employment.employerName:
                    emp_detail = f"I am employed at {employment.employerName}."

                p2_text = (
                    f"I and my spouse, {s_name}{s_pass_part} {location_phrase}. "
                    f"We both are financially independent. {emp_detail} {spouse_occ} "
                    f"Kindly find our relevant documents for your reference."
                ).replace("  ", " ").strip()
            else:
                emp_detail = ""
                if employment.employerName and employment.jobTitle:
                    emp_detail = f"I am working as {employment.jobTitle} at {employment.employerName}."
                elif employment.employerName:
                    emp_detail = f"I am employed at {employment.employerName}."

                p2_text = (
                    f"I {location_phrase}. I am financially independent and {emp_detail} "
                    f"Kindly find my relevant documents for your reference."
                ).replace("  ", " ").strip()

            blocks.append({"type": "paragraph", "content": p2_text})

        # -------------------------------------------------------------
        # Paragraph 3: Itinerary / Accommodation
        # -------------------------------------------------------------
        if fmt == "japan":
            if hotels:
                blocks.append({"type": "paragraph", "content": "We will be staying at the below hotels:"})
                headers = ["NAME", "DATE", "CONTACT NO."]
                rows = []
                for h in hotels:
                    date_range = ""
                    if h.checkInDate and h.checkOutDate:
                        date_range = f"{format_ordinal_date(h.checkInDate, include_year=False)} - {format_ordinal_date(h.checkOutDate, include_year=False)}"
                    hotel_name_loc = f"{h.hotelName}\n{h.city}".strip() if h.city else h.hotelName
                    rows.append([hotel_name_loc or "Hotel Stay", date_range or "Confirmed", h.contactNumber or ""])
                blocks.append({"type": "table", "headers": headers, "rows": rows, "style": "japan_hotel"})

        elif fmt == "singapore":
            if hotels:
                h = hotels[0]
                addr_part = f" – {h.address}, {h.city}".strip() if (h.address or h.city) else ""
                blocks.append({"type": "paragraph", "content": f"We will be staying at {h.hotelName}{addr_part}."})

        else:
            # Base Europe / Schengen Itinerary Narrative
            countries_visited = travel.countriesToBeVisited or country
            nights = travel.numberOfNights
            duration_str = f"for {nights} Nights" if nights > 0 else ""
            dates_str = f"from {start_fmt} to {end_fmt}" if (start_fmt and end_fmt) else ""

            p3_text = (
                f"We have planned to visit {countries_visited} for the purpose of {travel.purposeOfTravel or 'Tourism'}. "
                f"We look forward to visiting the country, exploring the sights, witnessing its culture and spending leisure time together. "
                f"We would be visiting {countries_visited} {duration_str} {dates_str}. "
                f"Kindly find our relevant documents for your perusal. We will return back to India after the end of our visit as we need to rejoin and look after our work and personal commitments."
            ).replace("  ", " ").strip()
            blocks.append({"type": "paragraph", "content": p3_text})

        # -------------------------------------------------------------
        # Financial Responsibility & Enclosures
        # -------------------------------------------------------------
        if fmt == "japan":
            blocks.append({"type": "paragraph", "content": f"All our travel expenses will be borne by {sponsor_wording}."})
            blocks.append({"type": "paragraph", "content": "Kindly find the attachment to our visa application form along with other relevant documents for the issuance of our visa."})
        elif fmt == "singapore":
            blocks.append({"type": "paragraph", "content": f"All our travel expenses will be borne by {sponsor_wording}."})
            if financial.bankStatementAvailable:
                blocks.append({"type": "paragraph", "content": "I am enclosing the bank statement of my bank account for your perusal."})
            blocks.append({"type": "paragraph", "content": "Request you to kindly grant us the Singapore visa multiple entries." if travel.entryType.lower() == 'multiple' or 'multiple' in travel.numberOfEntries.lower() else "Request you to kindly grant us the necessary visa."})
        else:
            p_fin = (
                f"All the expenses for travel to your country including Visa Fees, Flight bookings, Hotel booking, "
                f"Travel Insurance and other travel and personal expenses for us will be borne by {sponsor_wording}. "
                f"Kindly find our financial documents for your reference."
            )
            blocks.append({"type": "paragraph", "content": p_fin})
            blocks.append({"type": "paragraph", "content": "Enclosed please find our visa application form duly filled in and signed along with valid passport and other relevant documents to support our visa application."})
            blocks.append({"type": "paragraph", "content": "I request that you kindly grant us the necessary visa."})

        # Additional information if user chose to include it
        if app_data.additional.includeInCoverLetter and app_data.additional.content.strip():
            blocks.append({"type": "paragraph", "content": app_data.additional.content.strip()})

        # Sign-off Block
        closing_sal = config.get("closing", {}).get("salutation", "Yours Faithfully,")
        closing_thanks = config.get("closing", {}).get("thanks", "Thanking You!")
        
        sign_block = {
            "thanks": closing_thanks,
            "salutation": closing_sal,
            "name": full_name,
            "phone": app_data.contact.mobileNumber or "",
            "email": app_data.contact.emailAddress or "",
            "phoneLabel": "Mob:" if fmt == "singapore" else "Phone No.:",
            "emailLabel": "Email Id:" if fmt == "singapore" else "Email id:"
        }

        html = self._render_html(formatted_today, recipient_lines, subject_text, subject_bold, subject_underline, salutation, blocks, sign_block)

        # Anti-leak validation on generated output
        check_anti_leak_guard(html, app_data)

        return {
            "format": fmt,
            "appData": app_data.model_dump(),
            "date": formatted_today,
            "recipient": recipient_lines,
            "subject": {
                "text": subject_text,
                "bold": subject_bold,
                "underline": subject_underline
            },
            "salutation": salutation,
            "blocks": blocks,
            "signBlock": sign_block,
            "html": html
        }

    def _render_html(self, date_str, recipient, subject, bold, underline, salutation, blocks, sign_block) -> str:
        """Constructs semantic HTML document for the WYSIWYG editor."""
        lines = []
        lines.append(f'<div class="cover-letter-doc" style="font-family: Calibri, Arial, sans-serif; font-size: 11pt; line-height: 1.5; color: #1a1a1a; max-width: 800px; margin: 0 auto; background: #fff; padding: 40px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border-radius: 4px;">')
        
        # Date
        lines.append(f'<div style="text-align: right; margin-bottom: 24px; font-weight: 500;">{date_str}</div>')
        
        # Recipient
        lines.append('<div style="margin-bottom: 20px;">')
        for r in recipient:
            lines.append(f'<div>{r}</div>')
        lines.append('</div>')
        
        # Subject
        subj_style = f"font-weight: {'bold' if bold else 'normal'}; text-decoration: {'underline' if underline else 'none'}; margin-bottom: 20px;"
        lines.append(f'<div style="{subj_style}">{subject}</div>')
        
        # Salutation
        lines.append(f'<div style="margin-bottom: 16px;">{salutation}</div>')
        
        # Blocks
        for b in blocks:
            if b["type"] == "paragraph":
                lines.append(f'<p style="margin-bottom: 14px; text-align: justify;">{b["content"]}</p>')
            elif b["type"] == "table":
                lines.append('<table style="width: 100%; border-collapse: collapse; margin: 16px 0; font-size: 10.5pt; border: 1px solid #333;">')
                lines.append('<thead><tr style="background-color: #f8f9fa;">')
                for h in b["headers"]:
                    lines.append(f'<th style="border: 1px solid #333; padding: 8px 10px; text-align: left; font-weight: bold;">{h}</th>')
                lines.append('</tr></thead><tbody>')
                for row in b["rows"]:
                    lines.append('<tr>')
                    for cell in row:
                        formatted_cell = str(cell).replace('\n', '<br/>')
                        lines.append(f'<td style="border: 1px solid #333; padding: 8px 10px;">{formatted_cell}</td>')
                    lines.append('</tr>')
                lines.append('</tbody></table>')
                
        # Closing
        lines.append('<div style="margin-top: 24px;">')
        if sign_block.get("thanks"):
            lines.append(f'<div style="margin-bottom: 12px;">{sign_block["thanks"]}</div>')
        lines.append(f'<div style="margin-bottom: 28px;">{sign_block["salutation"]}</div>')
        lines.append(f'<div style="font-weight: bold;">{sign_block["name"]}</div>')
        if sign_block.get("phone"):
            lines.append(f'<div>{sign_block["phoneLabel"]} {sign_block["phone"]}</div>')
        if sign_block.get("email"):
            lines.append(f'<div>{sign_block["emailLabel"]} {sign_block["email"]}</div>')
        lines.append('</div>')
        
        lines.append('</div>')
        return "\n".join(lines)
