/**
 * Khanna Travels & Holidays — Client-Side Validation Service
 * Identifies missing fields according to country/template requirements.
 */

export const Validation = {
  validateForDocumentGeneration(app) {
    const errors = [];

    // Personal & Passport Requirements
    if (!app.personal?.fullName && !app.personal?.givenName) {
      errors.push('Applicant Full Name is required.');
    }
    if (!app.passport?.passportNumber) {
      errors.push('Passport Number is required.');
    }

    // Travel Requirements
    if (!app.travel?.travelStartDate) {
      errors.push('Travel Start Date is required.');
    }
    if (!app.travel?.travelEndDate) {
      errors.push('Travel End Date is required.');
    }

    // Template Specific Requirements
    const tpl = (app.selectedTemplateId || '').toLowerCase();
    if (tpl === 'japan') {
      if (!app.hotels || app.hotels.length === 0) {
        errors.push('Japan visa cover letter requires at least one hotel in Accommodation.');
      }
    }

    return {
      isValid: errors.length === 0,
      errors
    };
  }
};
