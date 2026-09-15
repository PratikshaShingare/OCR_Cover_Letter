"""
Khanna Travels & Holidays — OCR Service Abstraction Layer
Provides a pluggable, modular interface for passport OCR extraction.
"""

from abc import ABC, abstractmethod
from typing import Optional, Dict, Any
from pydantic import BaseModel, Field
import os


class ExtractedPassportData(BaseModel):
    # Core Passport Info
    passportType: str = "P"
    passportNumber: str = ""
    givenName: str = ""
    middleName: str = ""
    surname: str = ""
    fullName: str = ""
    previousName: str = ""
    nationality: str = "Indian"
    dob: str = ""
    gender: str = ""
    placeOfBirth: str = ""
    countryOfBirth: str = "India"
    issueDate: str = ""
    expiryDate: str = ""
    issuePlace: str = ""
    issuingCountry: str = "India"
    title: str = "Mr."
    email: str = ""
    phone: str = ""
    fatherFullName: str = ""
    motherFullName: str = ""
    spouseFullName: str = ""
    address: str = ""
    travellers: list = Field(default_factory=list)
    
    # Raw OCR metadata & confidence
    confidence: float = 0.95
    rawText: Optional[str] = None
    providerName: str = "Mock/Built-in OCR Service"
    fieldStatuses: Dict[str, str] = Field(default_factory=dict)
    fieldDetails: Dict[str, Dict[str, Any]] = Field(default_factory=dict)


class BaseOCRProvider(ABC):
    @abstractmethod
    async def extract(self, file_bytes: bytes, filename: str) -> ExtractedPassportData:
        """Extract passport fields from uploaded image/PDF bytes."""
        pass


def get_ocr_service() -> BaseOCRProvider:
    """Factory returns active OCR provider based on configuration."""
    provider_type = os.getenv("OCR_PROVIDER", "mock").lower()
    
    if provider_type == "mock":
        from .mock_provider import MockOCRProvider
        return MockOCRProvider()
    elif provider_type == "google_vision":
        # Google Vision fallback/provider if configured
        from .mock_provider import MockOCRProvider
        return MockOCRProvider()
    else:
        from .mock_provider import MockOCRProvider
        return MockOCRProvider()
