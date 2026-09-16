"""
Khanna Travels & Holidays — Master Client / Application Data Schema
Single Centralized Source of Truth for:
- Passport OCR Data
- Human Verified Master Data
- Internal Automated Excel Client Data Sheet
- Country-Specific Visa Cover Letters (DOCX / PDF)
"""

from typing import List, Optional, Dict, Literal, Any
from pydantic import BaseModel, Field
from datetime import date, datetime, timezone


class ApplicantPersonal(BaseModel):
    title: str = Field(default="Mr.", description="Title: Mr., Mrs., Ms., Dr., Master")
    givenName: str = Field(default="", description="Given / First Name")
    middleName: str = Field(default="", description="Middle Name")
    surname: str = Field(default="", description="Surname / Last Name")
    fullName: str = Field(default="", description="Full Name as in Passport")
    previousName: str = Field(default="", description="Previous Name if applicable")
    gender: Literal["Male", "Female", "Other", ""] = Field(default="", description="Gender")
    dob: str = Field(default="", description="Date of Birth in YYYY-MM-DD")
    placeOfBirth: str = Field(default="", description="Place / City of Birth")
    countryOfBirth: str = Field(default="", description="Country of Birth")
    nationality: str = Field(default="", description="Nationality")


class ApplicantPassport(BaseModel):
    passportType: str = Field(default="P", description="Passport Type")
    passportNumber: str = Field(default="", description="Passport Number")
    issueDate: str = Field(default="", description="Passport Issue Date")
    expiryDate: str = Field(default="", description="Passport Expiry Date")
    issuePlace: str = Field(default="", description="Place of Issue")
    issuingCountry: str = Field(default="", description="Issuing Country")
    issuingAuthority: str = Field(default="", description="Issuing Authority")


class ApplicantAddress(BaseModel):
    addressLine1: str = Field(default="", description="Address Line 1")
    addressLine2: str = Field(default="", description="Address Line 2")
    city: str = Field(default="", description="City")
    state: str = Field(default="", description="State")
    country: str = Field(default="", description="Country")
    postalCode: str = Field(default="", description="Postal / PIN Code")
    currentResidentialAddress: str = Field(default="", description="Full Current Residential Address")
    permanentAddress: str = Field(default="", description="Permanent Address")
    countryOfResidence: str = Field(default="", description="Country of Residence")


class ApplicantFamily(BaseModel):
    fatherFullName: str = Field(default="", description="Father's Full Name")
    motherFullName: str = Field(default="", description="Mother's Full Name")
    spouseFullName: str = Field(default="", description="Spouse's Full Name")
    spousePassportNumber: str = Field(default="", description="Spouse Passport Number")
    spouseDob: str = Field(default="", description="Spouse DOB")
    spouseNationality: str = Field(default="", description="Spouse Nationality")
    spouseOccupation: str = Field(default="", description="Spouse Occupation")


class ApplicantContact(BaseModel):
    countryCode: str = Field(default="+91", description="Country Calling Code")
    mobileNumber: str = Field(default="", description="Mobile Number")
    emailAddress: str = Field(default="", description="Email Address")
    alternateContact: str = Field(default="", description="Alternate Contact Number")
    emergencyContact: str = Field(default="", description="Emergency Contact Person & Number")


class ApplicantEmployment(BaseModel):
    employmentStatus: str = Field(
        default="Employed", 
        description="Employed, Corporate Employee, Business / Self-Employed, Student, Homemaker, Retired, Other"
    )
    employerName: str = Field(default="", description="Employer Name / Organization")
    jobTitle: str = Field(default="", description="Job Title / Designation")
    department: str = Field(default="", description="Department")
    employmentStartDate: str = Field(default="", description="Employment Start Date")
    annualIncome: str = Field(default="", description="Annual / Monthly Income")
    businessName: str = Field(default="", description="Business Name if self-employed")
    businessType: str = Field(default="", description="Business Type")
    schoolCollegeName: str = Field(default="", description="School or College Name if student")
    courseName: str = Field(default="", description="Course if student")
    gradeClass: str = Field(default="", description="Grade / Class if student")
    otherDetails: str = Field(default="", description="Other employment notes")


class Traveller(BaseModel):
    id: str = Field(default="", description="Unique UUID for traveller")
    title: str = Field(default="", description="Title: Mr., Mrs., Ms., Master, Dr.")
    givenName: str = Field(default="", description="Given Name")
    middleName: str = Field(default="", description="Middle Name")
    surname: str = Field(default="", description="Surname")
    fullName: str = Field(default="", description="Full Name")
    passportNumber: str = Field(default="", description="Passport Number")
    issueDate: str = Field(default="", description="Passport Issue Date")
    expiryDate: str = Field(default="", description="Passport Expiry Date")
    issuePlace: str = Field(default="", description="Place of Issue")
    dob: str = Field(default="", description="Date of Birth")
    nationality: str = Field(default="", description="Nationality")
    gender: str = Field(default="", description="Gender")
    relationship: str = Field(default="", description="Self, Spouse, Child, Parent, Sibling, Friend, Colleague, Other")
    occupation: str = Field(default="", description="Occupation e.g. Corporate Employee, Homemaker, Student")
    employer: str = Field(default="", description="Employer / Organization")
    schoolCollege: str = Field(default="", description="School / College")
    gradeClass: str = Field(default="", description="Grade / Class")
    otherInfo: str = Field(default="", description="Other notes or special requirements")
    fieldStatuses: Dict[str, str] = Field(default_factory=dict)


class Hotel(BaseModel):
    id: str = Field(default="", description="Unique UUID for hotel")
    hotelName: str = Field(default="", description="Hotel Name")
    address: str = Field(default="", description="Address Line")
    city: str = Field(default="", description="City")
    country: str = Field(default="", description="Country")
    checkInDate: str = Field(default="", description="Check-in Date (YYYY-MM-DD)")
    checkOutDate: str = Field(default="", description="Check-out Date (YYYY-MM-DD)")
    numberOfNights: int = Field(default=0, description="Calculated Number of Nights")
    contactNumber: str = Field(default="", description="Hotel Contact Phone Number")
    bookingReference: str = Field(default="", description="Booking Confirmation / PNR")


class TravelDetails(BaseModel):
    destinationCountry: str = Field(default="France", description="Primary Destination Country")
    destinationCountries: List[str] = Field(default_factory=list, description="All Destination Countries")
    visaType: str = Field(default="Tourism", description="Visa Type: Tourism, Business, Transit, etc.")
    purposeOfTravel: str = Field(default="Tourism", description="Purpose: Tourism, Holiday, Leisure")
    travelStartDate: str = Field(default="", description="Departure / Travel Start Date")
    travelEndDate: str = Field(default="", description="Return / Travel End Date")
    numberOfDays: int = Field(default=0, description="Calculated duration in days")
    numberOfNights: int = Field(default=0, description="Calculated duration in nights")
    intendedArrivalDate: str = Field(default="", description="Intended Arrival Date")
    intendedDepartureDate: str = Field(default="", description="Intended Departure Date")
    entryType: str = Field(default="Single", description="Single, Double, Multiple")
    numberOfEntries: str = Field(default="Single Entry", description="Number of entries requested")
    countriesToBeVisited: str = Field(default="", description="Countries to be visited")
    citiesToBeVisited: str = Field(default="", description="Cities to visit")
    flightNumber: str = Field(default="", description="Outbound flight")
    departureAirport: str = Field(default="", description="Departure Airport")
    arrivalAirport: str = Field(default="", description="Arrival Airport")
    returnFlightNumber: str = Field(default="", description="Return flight")
    pnrBookingRef: str = Field(default="", description="Flight PNR / Booking reference")


class FinancialInfo(BaseModel):
    tripSponsor: str = Field(default="Self-funded", description="Self-funded, Jointly funded, Spouse funded, Parent funded, Company funded, Other sponsor")
    sponsorName: str = Field(default="", description="Name of sponsor if not self")
    sponsorRelation: str = Field(default="", description="Relation of sponsor")
    bankStatementAvailable: bool = Field(default=false if False else False, description="Executive confirmed bank statement exists")
    itrAvailable: bool = Field(default=False, description="Executive confirmed ITR exists")
    salarySlipsAvailable: bool = Field(default=False, description="Executive confirmed salary slips exist")
    employmentLetterAvailable: bool = Field(default=False, description="Executive confirmed employment/NOC letter exists")
    otherFinancialDocs: str = Field(default="", description="Other financial documents confirmed")
    notes: str = Field(default="", description="Financial sponsorship notes")


class AdditionalInfo(BaseModel):
    content: str = Field(default="", description="Custom notes or details not covered in predefined fields")
    includeInCoverLetter: bool = Field(default=False, description="Whether to include in generated cover letter")


class MasterApplicationData(BaseModel):
    applicationId: str = Field(default="", description="Unique Application ID")
    applicationStatus: str = Field(default="Draft", description="Draft, In-Review, Verified, Completed")
    createdAt: str = Field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    updatedAt: str = Field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    
    # Core master sections - completely empty by default
    personal: ApplicantPersonal = Field(default_factory=ApplicantPersonal)
    passport: ApplicantPassport = Field(default_factory=ApplicantPassport)
    address: ApplicantAddress = Field(default_factory=ApplicantAddress)
    family: ApplicantFamily = Field(default_factory=ApplicantFamily)
    contact: ApplicantContact = Field(default_factory=ApplicantContact)
    employment: ApplicantEmployment = Field(default_factory=ApplicantEmployment)
    
    # Dynamic Travellers & Accommodation
    travellers: List[Traveller] = Field(default_factory=list)
    travel: TravelDetails = Field(default_factory=TravelDetails)
    hotels: List[Hotel] = Field(default_factory=list)
    financial: FinancialInfo = Field(default_factory=FinancialInfo)
    additional: AdditionalInfo = Field(default_factory=AdditionalInfo)
    
    # Field Verification Status: fieldPath -> 'extracted' | 'verify' | 'not_found' | 'manual' | 'verified'
    fieldStatuses: Dict[str, str] = Field(default_factory=dict)
    fieldDetails: Dict[str, Dict[str, Any]] = Field(default_factory=dict)
    
    # Country Template Configuration Choice
    selectedCountry: str = Field(default="France", description="Target Country: France, Japan, Singapore, etc.")
    selectedTemplateId: str = Field(default="standard", description="Selected template ID: standard, japan, singapore")
    
    # Stored generated document (HTML for rich editor)
    generatedDocumentHtml: Optional[str] = Field(default=None, description="HTML content for rich editor")
    isDocumentManuallyEdited: bool = Field(default=False, description="Flag indicating executive made custom edits")
