"""
Khanna Travels & Holidays — Template Registry API
Lists and resolves Base Europe Template and Country Overrides.
"""

from fastapi import APIRouter, HTTPException
from typing import List, Dict, Any
from ..document_generator.cover_letter_engine import CoverLetterEngine

router = APIRouter(prefix="/api/templates", tags=["Templates"])

AVAILABLE_FORMATS = [
    {
        "id": "standard",
        "name": "Standard",
        "description": "Standard European/Schengen base format. Multi-country itinerary breakdown, corporate employment & return-to-India commitment."
    },
    {
        "id": "japan",
        "name": "Japan",
        "description": "Consulate General of Japan format with family student/homemaker narrative and 3-Column Hotel Accommodation Table."
    },
    {
        "id": "singapore",
        "name": "Singapore",
        "description": "Consulate General of Singapore format with 5-column Passenger Table, underlined subject and multiple-entry request."
    }
]


@router.get("")
async def list_templates():
    """Returns the 3 available cover letter formats: Standard, Japan, Singapore."""
    return AVAILABLE_FORMATS


@router.get("/{format_id}")
async def get_format_template(format_id: str):
    engine = CoverLetterEngine()
    try:
        cfg = engine.load_template_config(format_id)
        return cfg
    except Exception as e:
        raise HTTPException(status_code=404, detail=f"Failed to load template for {format_id}: {str(e)}")
