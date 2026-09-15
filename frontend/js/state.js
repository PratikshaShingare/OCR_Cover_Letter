/**
 * Khanna Travels & Holidays — Application State Management
 * Maintains verified Master Client Data, active step, and event subscriptions.
 */

export function createEmptyApplication(appId = '') {
  return {
    applicationId: appId,
    applicationStatus: 'Draft',
    personal: {
      title: 'Mr.',
      givenName: '',
      middleName: '',
      surname: '',
      fullName: '',
      previousName: '',
      gender: '',
      dob: '',
      placeOfBirth: '',
      countryOfBirth: 'India',
      nationality: 'Indian'
    },
    passport: {
      passportType: 'P',
      passportNumber: '',
      issueDate: '',
      expiryDate: '',
      issuePlace: '',
      issuingCountry: 'India',
      issuingAuthority: ''
    },
    address: {
      addressLine1: '',
      addressLine2: '',
      city: '',
      state: '',
      country: 'India',
      postalCode: '',
      currentResidentialAddress: '',
      permanentAddress: '',
      countryOfResidence: 'India'
    },
    family: {
      fatherFullName: '',
      motherFullName: '',
      spouseFullName: '',
      spousePassportNumber: '',
      spouseDob: '',
      spouseNationality: 'Indian',
      spouseOccupation: ''
    },
    contact: {
      mobileNumber: '',
      emailAddress: '',
      alternateContact: '',
      emergencyContact: ''
    },
    employment: {
      employmentStatus: 'Employed',
      employerName: '',
      jobTitle: '',
      department: '',
      employmentStartDate: '',
      annualIncome: '',
      businessName: '',
      businessType: '',
      schoolCollegeName: '',
      courseName: '',
      gradeClass: '',
      otherDetails: ''
    },
    travellers: [],
    travel: {
      destinationCountry: 'France',
      destinationCountries: [],
      visaType: 'Tourism',
      purposeOfTravel: 'Tourism',
      travelStartDate: '',
      travelEndDate: '',
      numberOfDays: 0,
      numberOfNights: 0,
      entryType: 'Single',
      numberOfEntries: 'Single Entry',
      countriesToBeVisited: 'France',
      citiesToBeVisited: '',
      flightNumber: '',
      departureAirport: '',
      arrivalAirport: '',
      returnFlightNumber: '',
      pnrBookingRef: ''
    },
    hotels: [],
    financial: {
      tripSponsor: 'Self-funded',
      sponsorName: '',
      sponsorRelation: '',
      bankStatementAvailable: false,
      itrAvailable: false,
      salarySlipsAvailable: false,
      employmentLetterAvailable: false,
      otherFinancialDocs: ''
    },
    additional: {
      content: '',
      includeInCoverLetter: false
    },
    fieldStatuses: {},
    selectedCountry: 'France',
    selectedTemplateId: 'standard',
    generatedDocumentHtml: null,
    isDocumentManuallyEdited: false
  };
}

class AppState {
  constructor() {
    this.currentApplication = null;
    this.currentStep = 1;
    this.viewMode = 'dashboard'; // 'dashboard' | 'wizard'
    this.hasPassportFile = false;
    this.passportFileType = null;
    this.passportTotalPages = 1;
    this.passportCurrentPage = 1;
    this.listeners = [];
  }

  setApplication(app) {
    this.currentApplication = app;
    this.notify('application');
  }

  setStep(step) {
    if (step < 1) step = 1;
    if (step > 7) step = 7;
    this.currentStep = step;
    this.notify('step');
  }

  setViewMode(mode) {
    this.viewMode = mode;
    this.notify('viewMode');
  }

  setPassportFileStatus(hasFile, type = null, totalPages = 1) {
    this.hasPassportFile = hasFile;
    this.passportFileType = type;
    this.passportTotalPages = totalPages || 1;
    this.passportCurrentPage = 1;
    this.notify('passportFile');
  }

  setPassportCurrentPage(page) {
    if (page < 1) page = 1;
    if (page > this.passportTotalPages) page = this.passportTotalPages;
    this.passportCurrentPage = page;
    this.notify('passportPage');
  }

  subscribe(fn) {
    this.listeners.push(fn);
    return () => {
      this.listeners = this.listeners.filter(l => l !== fn);
    };
  }

  notify(eventType) {
    for (const fn of this.listeners) {
      fn(eventType, this);
    }
  }
}

export const State = new AppState();
