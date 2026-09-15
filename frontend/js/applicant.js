/**
 * Khanna Travels & Holidays — Applicant Information Module
 * Handles human verification of personal, passport, address, family, contact,
 * and employment data with progressive disclosure and field status badges.
 */

import { Api } from './api.js';
import { State } from './state.js';

export const ApplicantModule = {
  render(container) {
    const app = State.currentApplication;
    if (!app) return;

    const p = app.personal || {};
    const pass = app.passport || {};
    const addr = app.address || {};
    const fam = app.family || {};
    const cont = app.contact || {};
    const emp = app.employment || {};
    const statuses = app.fieldStatuses || {};
    const details = app.fieldDetails || {};

    const getBadge = (path, value) => {
      const detail = details[path];
      if (detail) {
        if (detail.status === 'extracted' || (detail.confidence && detail.confidence >= 0.85)) {
          return `<span class="status-badge badge-extracted">✓ Extracted from passport</span>`;
        }
        if (detail.status === 'verify') {
          return `<span class="status-badge badge-warning">⚠ Please verify</span>`;
        }
        if (detail.status === 'not_found' || !detail.value) {
          return `<span class="status-badge badge-neutral">⚠ Not found — enter manually</span>`;
        }
      }
      const st = statuses[path];
      if (st === 'ocr' && value) {
        return `<span class="status-badge badge-extracted">✓ Extracted from passport</span>`;
      }
      if (!value) {
        return `<span class="status-badge badge-neutral">⚠ Not found — enter manually</span>`;
      }
      return `<span class="status-badge badge-neutral">Manual Entry</span>`;
    };

    container.innerHTML = `
      <div style="max-width: 960px; margin: 0 auto;">
        <div style="margin-bottom: 24px;">
          <h2 style="font-size: 1.35rem; font-weight: 800; color: var(--dark); margin-bottom: 6px;">
            Step 2: Verify Applicant Master Data
          </h2>
          <p style="color: var(--muted); font-size: 0.88rem;">
            Review extracted passport fields and complete missing details. Data saved here becomes verified master client data.
          </p>
        </div>

        <form id="applicant-form">
          <!-- 1. PERSONAL INFORMATION -->
          <div class="accordion-section open">
            <div class="accordion-header">
              <div class="accordion-title">
                <span>1. Personal Information</span>
                <span class="accordion-badge">Core Details</span>
              </div>
              <svg class="accordion-toggle-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"></polyline></svg>
            </div>
            <div class="accordion-content">
              <div class="form-grid">
                <div class="form-group">
                  <div class="form-label-row">
                    <label class="form-label">Title</label>
                  </div>
                  <select class="form-control" name="personal.title">
                    <option value="Mr." ${p.title === 'Mr.' ? 'selected' : ''}>Mr.</option>
                    <option value="Mrs." ${p.title === 'Mrs.' ? 'selected' : ''}>Mrs.</option>
                    <option value="Ms." ${p.title === 'Ms.' ? 'selected' : ''}>Ms.</option>
                    <option value="Dr." ${p.title === 'Dr.' ? 'selected' : ''}>Dr.</option>
                    <option value="Master" ${p.title === 'Master' ? 'selected' : ''}>Master</option>
                  </select>
                </div>

                <div class="form-group">
                  <div class="form-label-row">
                    <label class="form-label">Full Name (as in Passport)</label>
                    ${getBadge('personal.fullName', p.fullName)}
                  </div>
                  <input type="text" class="form-control" name="personal.fullName" value="${p.fullName || ''}" placeholder="e.g. Rahul Sharma" />
                </div>

                <div class="form-group">
                  <div class="form-label-row">
                    <label class="form-label">Given Name</label>
                    ${getBadge('personal.givenName', p.givenName)}
                  </div>
                  <input type="text" class="form-control" name="personal.givenName" value="${p.givenName || ''}" placeholder="First Name" />
                </div>

                <div class="form-group">
                  <div class="form-label-row">
                    <label class="form-label">Surname</label>
                    ${getBadge('personal.surname', p.surname)}
                  </div>
                  <input type="text" class="form-control" name="personal.surname" value="${p.surname || ''}" placeholder="Last Name" />
                </div>

                <div class="form-group">
                  <div class="form-label-row">
                    <label class="form-label">Date of Birth</label>
                    ${getBadge('personal.dob', p.dob)}
                  </div>
                  <input type="text" class="form-control" name="personal.dob" value="${p.dob || ''}" placeholder="YYYY-MM-DD" />
                </div>

                <div class="form-group">
                  <div class="form-label-row">
                    <label class="form-label">Gender</label>
                    ${getBadge('personal.gender', p.gender)}
                  </div>
                  <select class="form-control" name="personal.gender">
                    <option value="" ${!p.gender ? 'selected' : ''}>-- Select Gender --</option>
                    <option value="Male" ${p.gender === 'Male' ? 'selected' : ''}>Male</option>
                    <option value="Female" ${p.gender === 'Female' ? 'selected' : ''}>Female</option>
                    <option value="Other" ${p.gender === 'Other' ? 'selected' : ''}>Other</option>
                  </select>
                </div>

                <div class="form-group">
                  <div class="form-label-row">
                    <label class="form-label">Place of Birth</label>
                    ${getBadge('personal.placeOfBirth', p.placeOfBirth)}
                  </div>
                  <input type="text" class="form-control" name="personal.placeOfBirth" value="${p.placeOfBirth || ''}" placeholder="City of Birth" />
                </div>

                <div class="form-group">
                  <div class="form-label-row">
                    <label class="form-label">Nationality</label>
                    ${getBadge('personal.nationality', p.nationality)}
                  </div>
                  <input type="text" class="form-control" name="personal.nationality" value="${p.nationality || 'Indian'}" placeholder="Nationality" />
                </div>
              </div>
            </div>
          </div>

          <!-- 2. PASSPORT INFORMATION -->
          <div class="accordion-section open">
            <div class="accordion-header">
              <div class="accordion-title">
                <span>2. Passport Information</span>
                <span class="accordion-badge">Document Data</span>
              </div>
              <svg class="accordion-toggle-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"></polyline></svg>
            </div>
            <div class="accordion-content">
              <div class="form-grid">
                <div class="form-group">
                  <div class="form-label-row">
                    <label class="form-label">Passport Number</label>
                    ${getBadge('passport.passportNumber', pass.passportNumber)}
                  </div>
                  <input type="text" class="form-control" name="passport.passportNumber" value="${pass.passportNumber || ''}" placeholder="e.g. Z1234567" />
                </div>

                <div class="form-group">
                  <div class="form-label-row">
                    <label class="form-label">Issue Date</label>
                    ${getBadge('passport.issueDate', pass.issueDate)}
                  </div>
                  <input type="text" class="form-control" name="passport.issueDate" value="${pass.issueDate || ''}" placeholder="YYYY-MM-DD" />
                </div>

                <div class="form-group">
                  <div class="form-label-row">
                    <label class="form-label">Expiry Date</label>
                    ${getBadge('passport.expiryDate', pass.expiryDate)}
                  </div>
                  <input type="text" class="form-control" name="passport.expiryDate" value="${pass.expiryDate || ''}" placeholder="YYYY-MM-DD" />
                </div>

                <div class="form-group">
                  <div class="form-label-row">
                    <label class="form-label">Place of Issue</label>
                    ${getBadge('passport.issuePlace', pass.issuePlace)}
                  </div>
                  <input type="text" class="form-control" name="passport.issuePlace" value="${pass.issuePlace || ''}" placeholder="e.g. Mumbai, Delhi" />
                </div>

                <div class="form-group">
                  <div class="form-label-row">
                    <label class="form-label">Issuing Country</label>
                  </div>
                  <input type="text" class="form-control" name="passport.issuingCountry" value="${pass.issuingCountry || 'India'}" />
                </div>
              </div>
            </div>
          </div>

          <!-- 3. ADDRESS INFORMATION -->
          <div class="accordion-section">
            <div class="accordion-header">
              <div class="accordion-title">
                <span>3. Address Information</span>
                <span class="accordion-badge">Residence</span>
              </div>
              <svg class="accordion-toggle-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"></polyline></svg>
            </div>
            <div class="accordion-content">
              <div class="form-grid">
                <div class="form-group col-span-2">
                  <div class="form-label-row">
                    <label class="form-label">Residential Address Line 1</label>
                    ${getBadge('address.addressLine1', addr.addressLine1 || addr.currentResidentialAddress)}
                  </div>
                  <input type="text" class="form-control" name="address.addressLine1" value="${addr.addressLine1 || ''}" placeholder="Flat / Building / Street" />
                </div>
                <div class="form-group col-span-2">
                  <label class="form-label">Address Line 2</label>
                  <input type="text" class="form-control" name="address.addressLine2" value="${addr.addressLine2 || ''}" placeholder="Area / Locality" />
                </div>
                <div class="form-group">
                  <label class="form-label">City</label>
                  <input type="text" class="form-control" name="address.city" value="${addr.city || ''}" placeholder="City" />
                </div>
                <div class="form-group">
                  <label class="form-label">State</label>
                  <input type="text" class="form-control" name="address.state" value="${addr.state || ''}" placeholder="State" />
                </div>
                <div class="form-group">
                  <label class="form-label">PIN / Postal Code</label>
                  <input type="text" class="form-control" name="address.postalCode" value="${addr.postalCode || ''}" placeholder="Postal Code" />
                </div>
                <div class="form-group">
                  <label class="form-label">Country</label>
                  <input type="text" class="form-control" name="address.country" value="${addr.country || 'India'}" />
                </div>
              </div>
            </div>
          </div>

          <!-- 4. FAMILY INFORMATION -->
          <div class="accordion-section">
            <div class="accordion-header">
              <div class="accordion-title">
                <span>4. Family Information</span>
                <span class="accordion-badge">Relations</span>
              </div>
              <svg class="accordion-toggle-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"></polyline></svg>
            </div>
            <div class="accordion-content">
              <div class="form-grid">
                <div class="form-group">
                  <div class="form-label-row">
                    <label class="form-label">Father's Full Name</label>
                    ${getBadge('family.fatherFullName', fam.fatherFullName)}
                  </div>
                  <input type="text" class="form-control" name="family.fatherFullName" value="${fam.fatherFullName || ''}" placeholder="Father's Name" />
                </div>
                <div class="form-group">
                  <div class="form-label-row">
                    <label class="form-label">Mother's Full Name</label>
                    ${getBadge('family.motherFullName', fam.motherFullName)}
                  </div>
                  <input type="text" class="form-control" name="family.motherFullName" value="${fam.motherFullName || ''}" placeholder="Mother's Name" />
                </div>
                <div class="form-group">
                  <div class="form-label-row">
                    <label class="form-label">Spouse Full Name</label>
                    ${getBadge('family.spouseFullName', fam.spouseFullName)}
                  </div>
                  <input type="text" class="form-control" name="family.spouseFullName" value="${fam.spouseFullName || ''}" placeholder="Spouse's Name (if married)" />
                </div>
                <div class="form-group">
                  <label class="form-label">Spouse Passport Number</label>
                  <input type="text" class="form-control" name="family.spousePassportNumber" value="${fam.spousePassportNumber || ''}" placeholder="Spouse Passport No." />
                </div>
              </div>
            </div>
          </div>

          <!-- 5. CONTACT INFORMATION -->
          <div class="accordion-section">
            <div class="accordion-header">
              <div class="accordion-title">
                <span>5. Contact Information</span>
                <span class="accordion-badge">Direct Contact</span>
              </div>
              <svg class="accordion-toggle-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"></polyline></svg>
            </div>
            <div class="accordion-content">
              <div class="form-grid">
                <div class="form-group">
                  <label class="form-label">Mobile Phone Number</label>
                  <input type="text" class="form-control" name="contact.mobileNumber" value="${cont.mobileNumber || ''}" placeholder="+91 98765 43210" />
                </div>
                <div class="form-group">
                  <label class="form-label">Email Address</label>
                  <input type="email" class="form-control" name="contact.emailAddress" value="${cont.emailAddress || ''}" placeholder="client@example.com" />
                </div>
                <div class="form-group">
                  <label class="form-label">Alternate Contact</label>
                  <input type="text" class="form-control" name="contact.alternateContact" value="${cont.alternateContact || ''}" placeholder="Phone or telephone" />
                </div>
              </div>
            </div>
          </div>

          <!-- 6. EMPLOYMENT INFORMATION (Progressive Disclosure) -->
          <div class="accordion-section open">
            <div class="accordion-header">
              <div class="accordion-title">
                <span>6. Employment Information</span>
                <span class="accordion-badge">Occupation Details</span>
              </div>
              <svg class="accordion-toggle-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"></polyline></svg>
            </div>
            <div class="accordion-content">
              <div class="form-group" style="margin-bottom: 18px;">
                <label class="form-label">Employment Status</label>
                <select class="form-control" id="emp-status-select" name="employment.employmentStatus">
                  <option value="Employed" ${emp.employmentStatus === 'Employed' ? 'selected' : ''}>Corporate Employee / Salaried</option>
                  <option value="Business / Self-Employed" ${emp.employmentStatus === 'Business / Self-Employed' ? 'selected' : ''}>Business / Self-Employed</option>
                  <option value="Student" ${emp.employmentStatus === 'Student' ? 'selected' : ''}>Student</option>
                  <option value="Homemaker" ${emp.employmentStatus === 'Homemaker' ? 'selected' : ''}>Homemaker</option>
                  <option value="Retired" ${emp.employmentStatus === 'Retired' ? 'selected' : ''}>Retired</option>
                  <option value="Other" ${emp.employmentStatus === 'Other' ? 'selected' : ''}>Other</option>
                </select>
              </div>

              <!-- Salaried Group -->
              <div id="group-salaried" class="form-grid" style="display: ${emp.employmentStatus === 'Employed' ? 'grid' : 'none'};">
                <div class="form-group">
                  <label class="form-label">Employer / Organization Name</label>
                  <input type="text" class="form-control" name="employment.employerName" value="${emp.employerName || ''}" placeholder="Company Name" />
                </div>
                <div class="form-group">
                  <label class="form-label">Job Title / Designation</label>
                  <input type="text" class="form-control" name="employment.jobTitle" value="${emp.jobTitle || ''}" placeholder="e.g. Senior Software Engineer" />
                </div>
                <div class="form-group">
                  <label class="form-label">Department</label>
                  <input type="text" class="form-control" name="employment.department" value="${emp.department || ''}" placeholder="e.g. Finance, Tech" />
                </div>
              </div>

              <!-- Self-Employed Group -->
              <div id="group-business" class="form-grid" style="display: ${emp.employmentStatus === 'Business / Self-Employed' ? 'grid' : 'none'};">
                <div class="form-group">
                  <label class="form-label">Business / Enterprise Name</label>
                  <input type="text" class="form-control" name="employment.businessName" value="${emp.businessName || ''}" placeholder="Registered Business Name" />
                </div>
                <div class="form-group">
                  <label class="form-label">Business Type / Nature</label>
                  <input type="text" class="form-control" name="employment.businessType" value="${emp.businessType || ''}" placeholder="e.g. Wholesale Trade, Consultancy" />
                </div>
              </div>

              <!-- Student Group -->
              <div id="group-student" class="form-grid" style="display: ${emp.employmentStatus === 'Student' ? 'grid' : 'none'};">
                <div class="form-group">
                  <label class="form-label">School / College Name</label>
                  <input type="text" class="form-control" name="employment.schoolCollegeName" value="${emp.schoolCollegeName || ''}" placeholder="Educational Institute" />
                </div>
                <div class="form-group">
                  <label class="form-label">Grade / Class / Degree</label>
                  <input type="text" class="form-control" name="employment.gradeClass" value="${emp.gradeClass || ''}" placeholder="e.g. 10th Grade, B.Com 2nd Year" />
                </div>
              </div>
            </div>
          </div>

          <!-- Navigation Buttons -->
          <div class="step-nav-bar">
            <button type="button" class="btn btn-secondary" id="btn-back-passport">
              ← Back to Passport
            </button>
            <button type="submit" class="btn btn-primary" id="btn-save-applicant">
              Confirm & Continue to Travellers →
            </button>
          </div>
        </form>
      </div>
    `;

    this.bindEvents(container);
  },

  bindEvents(container) {
    // Accordion toggle
    container.querySelectorAll('.accordion-header').forEach(header => {
      header.addEventListener('click', () => {
        header.parentElement.classList.toggle('open');
      });
    });

    // Progressive disclosure for employment
    const empSelect = container.querySelector('#emp-status-select');
    const groupSalaried = container.querySelector('#group-salaried');
    const groupBusiness = container.querySelector('#group-business');
    const groupStudent = container.querySelector('#group-student');

    if (empSelect) {
      empSelect.addEventListener('change', (e) => {
        const val = e.target.value;
        if (groupSalaried) groupSalaried.style.display = val === 'Employed' ? 'grid' : 'none';
        if (groupBusiness) groupBusiness.style.display = val === 'Business / Self-Employed' ? 'grid' : 'none';
        if (groupStudent) groupStudent.style.display = val === 'Student' ? 'grid' : 'none';
      });
    }

    // Back button
    const btnBack = container.querySelector('#btn-back-passport');
    if (btnBack) {
      btnBack.addEventListener('click', () => State.setStep(1));
    }

    // Form Submit / Save
    const form = container.querySelector('#applicant-form');
    if (form) {
      form.addEventListener('submit', async (e) => {
        e.preventDefault();
        const app = State.currentApplication;
        const formData = new FormData(form);

        for (const [key, val] of formData.entries()) {
          const parts = key.split('.');
          if (parts.length === 2) {
            const section = parts[0];
            const field = parts[1];
            if (!app[section]) app[section] = {};
            app[section][field] = val.trim();
          }
        }

        // Auto-compose full name if givenName and surname entered
        if (!app.personal.fullName && (app.personal.givenName || app.personal.surname)) {
          app.personal.fullName = `${app.personal.givenName} ${app.personal.surname}`.trim();
        }

        app.applicationStatus = 'Verified';
        app.updatedAt = new Date().toISOString();

        try {
          const saved = await Api.saveApplication(app);
          State.setApplication(saved);
          State.setStep(3); // Go to Travellers
        } catch (err) {
          alert(`Error saving applicant details: ${err.message}`);
        }
      });
    }
  }
};
