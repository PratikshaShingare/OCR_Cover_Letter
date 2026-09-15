"""
Khanna Travels & Holidays — PDF Document Generator
Uses ReportLab Platypus to generate pristine, vector-precise PDFs
matching the reference layout and typography.
"""

import os
from typing import Dict, Any
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_RIGHT, TA_JUSTIFY, TA_LEFT


def build_pdf_cover_letter(doc_data: Dict[str, Any], output_path: str):
    """Generates a high-quality PDF cover letter directly matching the DOCX."""
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # 1-inch margins
    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        leftMargin=inch,
        rightMargin=inch,
        topMargin=inch,
        bottomMargin=inch
    )
    
    styles = getSampleStyleSheet()
    
    # Define custom styles
    date_style = ParagraphStyle(
        'CoverLetterDate',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=11,
        leading=14,
        alignment=TA_RIGHT,
        textColor=colors.black
    )
    
    recipient_style = ParagraphStyle(
        'CoverLetterRecipient',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=11,
        leading=14,
        textColor=colors.black
    )
    
    subject_style = ParagraphStyle(
        'CoverLetterSubject',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=15,
        textColor=colors.black
    )
    
    body_style = ParagraphStyle(
        'CoverLetterBody',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=11,
        leading=15,
        alignment=TA_JUSTIFY,
        textColor=colors.black
    )
    
    sign_style = ParagraphStyle(
        'CoverLetterSign',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=11,
        leading=14,
        textColor=colors.black
    )
    
    sign_bold_style = ParagraphStyle(
        'CoverLetterSignBold',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=11,
        leading=14,
        textColor=colors.black
    )
    
    story = []
    
    # 1. Date
    story.append(Paragraph(doc_data.get("date", ""), date_style))
    story.append(Spacer(1, 14))
    
    # 2. Recipient
    for line in doc_data.get("recipient", []):
        story.append(Paragraph(line, recipient_style))
    story.append(Spacer(1, 12))
    
    # 3. Subject
    subj_data = doc_data.get("subject", {})
    text = subj_data.get("text", "")
    if subj_data.get("underline", False):
        text = f"<u>{text}</u>"
    story.append(Paragraph(text, subject_style))
    story.append(Spacer(1, 12))
    
    # 4. Salutation
    story.append(Paragraph(doc_data.get("salutation", "Dear Sir/Ma’am,"), recipient_style))
    story.append(Spacer(1, 10))
    
    # 5. Blocks
    for block in doc_data.get("blocks", []):
        b_type = block.get("type")
        if b_type == "paragraph":
            story.append(Paragraph(block.get("content", ""), body_style))
            story.append(Spacer(1, 10))
            
        elif b_type == "table":
            headers = block.get("headers", [])
            rows = block.get("rows", [])
            
            # Convert cells to Paragraphs for text wrapping
            cell_hdr_style = ParagraphStyle(
                'TH', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=12
            )
            cell_data_style = ParagraphStyle(
                'TD', parent=styles['Normal'], fontName='Helvetica', fontSize=9.5, leading=12
            )
            
            table_data = []
            hdr_row = [Paragraph(f"<b>{h}</b>", cell_hdr_style) for h in headers]
            table_data.append(hdr_row)
            
            for row in rows:
                formatted_row = []
                for cell in row:
                    cell_html = str(cell).replace("\n", "<br/>")
                    formatted_row.append(Paragraph(cell_html, cell_data_style))
                table_data.append(formatted_row)
                
            # Column widths based on header count
            avail_width = 6.5 * inch
            col_width = avail_width / len(headers)
            
            # Specific custom column widths for Japan & Singapore
            col_widths = None
            if len(headers) == 3:  # Japan Hotel: NAME, DATE, CONTACT NO.
                col_widths = [3.2 * inch, 1.8 * inch, 1.5 * inch]
            elif len(headers) == 5: # Singapore Passengers: Sr, Name, Passport, Relation, Occupation
                col_widths = [0.5 * inch, 2.0 * inch, 1.3 * inch, 1.2 * inch, 1.5 * inch]
                
            t = Table(table_data, colWidths=col_widths or [col_width] * len(headers))
            t.setStyle(TableStyle([
                ('BOX', (0, 0), (-1, -1), 0.75, colors.black),
                ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#F8F9FA")),
                ('TOPPADDING', (0, 0), (-1, -1), 5),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                ('LEFTPADDING', (0, 0), (-1, -1), 6),
                ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ]))
            story.append(t)
            story.append(Spacer(1, 12))
            
    # 6. Closing & Sign-off
    sign = doc_data.get("signBlock", {})
    if sign.get("thanks"):
        story.append(Paragraph(sign["thanks"], recipient_style))
        story.append(Spacer(1, 4))
        
    story.append(Paragraph(sign.get("salutation", "Yours Faithfully,"), recipient_style))
    story.append(Spacer(1, 14))
    
    story.append(Paragraph(sign.get("name", ""), sign_bold_style))
    if sign.get("phone"):
        story.append(Paragraph(f"{sign.get('phoneLabel', 'Phone No.:')} {sign['phone']}", sign_style))
    if sign.get("email"):
        story.append(Paragraph(f"{sign.get('emailLabel', 'Email id:')} {sign['email']}", sign_style))
        
    doc.build(story)
    return output_path
