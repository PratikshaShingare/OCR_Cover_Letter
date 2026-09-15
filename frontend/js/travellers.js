/**
 * Khanna Travels & Holidays — Travellers & Companions Module
 * Dynamically manages accompanying family members and co-passengers.
 */

import { Api } from './api.js';
import { State } from './state.js';

export const TravellersModule = {
  render(container) {
    const app = State.currentApplication;
    if (!app) return;

    if (!app.travellers) app.travellers = [];

    container.innerHTML = `
      <div style="max-width: 960px; margin: 0 auto;">
        <div style="display: flex; justify-content: space-between; align-items: flex-end; margin-bottom: 24px;">
          <div>
            <h2 style="font-size: 1.35rem; font-weight: 800; color: var(--dark); margin-bottom: 6px;">
              Step 3: Accompanying Travellers
            </h2>
            <p style="color: var(--muted); font-size: 0.88rem;">
              Add family members or co-passengers travelling on the same trip. Leave empty if solo traveller.
            </p>
          </div>
          <button type="button" class="btn btn-primary btn-sm" id="btn-add-traveller">
            + Add Traveller
          </button>
        </div>

        <!-- Primary Applicant Card -->
        <div class="card" style="margin-bottom: 20px; background-color: #f8fafc; border-left: 4px solid var(--primary);">
          <div style="display: flex; justify-content: space-between; align-items: center;">
            <div>
              <span style="font-size: 0.72rem; font-weight: 700; color: var(--primary); text-transform: uppercase;">
                PRIMARY APPLICANT (SELF)
              </span>
              <h3 style="font-size: 1rem; font-weight: 700; color: var(--dark); margin-top: 2px;">
                ${app.personal?.fullName || 'Primary Applicant'}
              </h3>
              <div style="font-size: 0.82rem; color: var(--muted); margin-top: 4px;">
                Passport: <strong>${app.passport?.passportNumber || 'Not specified'}</strong> • 
                Occupation: <strong>${app.employment?.jobTitle || app.employment?.employmentStatus || 'Employed'}</strong>
              </div>
            </div>
            <span class="status-badge badge-neutral">Lead Traveller</span>
          </div>
        </div>

        <!-- Dynamic Travellers List -->
        <div id="travellers-list">
          ${app.travellers.length === 0 ? `
            <div class="card" style="text-align: center; padding: 36px 20px; color: var(--muted); margin-bottom: 20px;">
              <p style="font-size: 0.92rem; margin-bottom: 12px;">No accompanying travellers added.</p>
              <button type="button" class="btn btn-secondary btn-sm" id="btn-add-traveller-empty">
                + Add Accompanying Traveller
              </button>
            </div>
          ` : app.travellers.map((t, idx) => `
            <div class="card traveller-card" data-index="${idx}" style="margin-bottom: 16px; position: relative;">
              <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; border-bottom: 1px solid var(--border); padding-bottom: 10px;">
                <div style="font-weight: 700; font-size: 0.94rem; color: var(--dark);">
                  Traveller ${idx + 2}: ${t.fullName || t.relationship || 'Companion'}
                </div>
                <button type="button" class="btn btn-danger btn-sm btn-remove-traveller" data-index="${idx}">
                  Remove
                </button>
              </div>

              <div class="form-grid">
                <div class="form-group">
                  <label class="form-label">Relationship</label>
                  <select class="form-control t-field" data-field="relationship" data-index="${idx}">
                    <option value="Spouse" ${t.relationship === 'Spouse' ? 'selected' : ''}>Spouse / Wife / Husband</option>
                    <option value="Son" ${t.relationship === 'Son' ? 'selected' : ''}>Son</option>
                    <option value="Daughter" ${t.relationship === 'Daughter' ? 'selected' : ''}>Daughter</option>
                    <option value="Father" ${t.relationship === 'Father' ? 'selected' : ''}>Father</option>
                    <option value="Mother" ${t.relationship === 'Mother' ? 'selected' : ''}>Mother</option>
                    <option value="Brother" ${t.relationship === 'Brother' ? 'selected' : ''}>Brother</option>
                    <option value="Sister" ${t.relationship === 'Sister' ? 'selected' : ''}>Sister</option>
                    <option value="Other" ${t.relationship === 'Other' ? 'selected' : ''}>Other</option>
                  </select>
                </div>

                <div class="form-group">
                  <label class="form-label">Title</label>
                  <select class="form-control t-field" data-field="title" data-index="${idx}">
                    <option value="Mr." ${t.title === 'Mr.' ? 'selected' : ''}>Mr.</option>
                    <option value="Mrs." ${t.title === 'Mrs.' ? 'selected' : ''}>Mrs.</option>
                    <option value="Ms." ${t.title === 'Ms.' ? 'selected' : ''}>Ms.</option>
                    <option value="Master" ${t.title === 'Master' ? 'selected' : ''}>Master</option>
                  </select>
                </div>

                <div class="form-group">
                  <label class="form-label">Full Name</label>
                  <input type="text" class="form-control t-field" data-field="fullName" data-index="${idx}" value="${t.fullName || ''}" placeholder="Full Name as in Passport" />
                </div>

                <div class="form-group">
                  <label class="form-label">Passport Number</label>
                  <input type="text" class="form-control t-field" data-field="passportNumber" data-index="${idx}" value="${t.passportNumber || ''}" placeholder="Passport No." />
                </div>

                <div class="form-group">
                  <label class="form-label">Date of Birth</label>
                  <input type="text" class="form-control t-field" data-field="dob" data-index="${idx}" value="${t.dob || ''}" placeholder="YYYY-MM-DD" />
                </div>

                <div class="form-group">
                  <label class="form-label">Passport Issue Date</label>
                  <input type="text" class="form-control t-field" data-field="issueDate" data-index="${idx}" value="${t.issueDate || ''}" placeholder="YYYY-MM-DD" />
                </div>

                <div class="form-group">
                  <label class="form-label">Passport Issue Place</label>
                  <input type="text" class="form-control t-field" data-field="issuePlace" data-index="${idx}" value="${t.issuePlace || ''}" placeholder="e.g. Mumbai" />
                </div>

                <div class="form-group">
                  <label class="form-label">Occupation / Role</label>
                  <input type="text" class="form-control t-field" data-field="occupation" data-index="${idx}" value="${t.occupation || ''}" placeholder="e.g. Homemaker, Student, Architect" />
                </div>

                <div class="form-group">
                  <label class="form-label">Employer (if employed)</label>
                  <input type="text" class="form-control t-field" data-field="employer" data-index="${idx}" value="${t.employer || ''}" placeholder="Company Name" />
                </div>

                <div class="form-group">
                  <label class="form-label">School / College (if student)</label>
                  <input type="text" class="form-control t-field" data-field="schoolCollege" data-index="${idx}" value="${t.schoolCollege || ''}" placeholder="School / College Name" />
                </div>

                <div class="form-group">
                  <label class="form-label">Grade / Class (if student)</label>
                  <input type="text" class="form-control t-field" data-field="gradeClass" data-index="${idx}" value="${t.gradeClass || ''}" placeholder="e.g. 10th Grade" />
                </div>
              </div>
            </div>
          `).join('')}
        </div>

        <!-- Navigation Buttons -->
        <div class="step-nav-bar">
          <button type="button" class="btn btn-secondary" id="btn-back-applicant">
            ← Back to Applicant
          </button>
          <button type="button" class="btn btn-primary" id="btn-save-travellers">
            Save & Continue to Travel Details →
          </button>
        </div>
      </div>
    `;

    this.bindEvents(container);
  },

  bindEvents(container) {
    const app = State.currentApplication;

    const addTraveller = () => {
      if (!app.travellers) app.travellers = [];
      app.travellers.push({
        id: `T-${Date.now()}`,
        title: 'Mrs.',
        givenName: '',
        middleName: '',
        surname: '',
        fullName: '',
        passportNumber: '',
        issueDate: '',
        expiryDate: '',
        issuePlace: '',
        dob: '',
        nationality: 'Indian',
        gender: 'Female',
        relationship: 'Spouse',
        occupation: 'Homemaker',
        employer: '',
        schoolCollege: '',
        gradeClass: '',
        otherInfo: ''
      });
      this.render(container);
    };

    const btnAdd = container.querySelector('#btn-add-traveller');
    if (btnAdd) btnAdd.addEventListener('click', addTraveller);

    const btnAddEmpty = container.querySelector('#btn-add-traveller-empty');
    if (btnAddEmpty) btnAddEmpty.addEventListener('click', addTraveller);

    // Live binding for traveller inputs
    container.querySelectorAll('.t-field').forEach(input => {
      input.addEventListener('input', (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        const field = e.target.dataset.field;
        if (app.travellers && app.travellers[idx]) {
          app.travellers[idx][field] = e.target.value.trim();
        }
      });
      input.addEventListener('change', (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        const field = e.target.dataset.field;
        if (app.travellers && app.travellers[idx]) {
          app.travellers[idx][field] = e.target.value.trim();
        }
      });
    });

    // Remove traveller button
    container.querySelectorAll('.btn-remove-traveller').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        if (confirm('Remove this traveller from the application?')) {
          app.travellers.splice(idx, 1);
          this.render(container);
        }
      });
    });

    // Navigation
    const btnBack = container.querySelector('#btn-back-applicant');
    if (btnBack) btnBack.addEventListener('click', () => State.setStep(2));

    const btnSave = container.querySelector('#btn-save-travellers');
    if (btnSave) {
      btnSave.addEventListener('click', async () => {
        try {
          const saved = await Api.saveApplication(app);
          State.setApplication(saved);
          State.setStep(4); // Go to Travel Details
        } catch (err) {
          alert(`Error saving travellers: ${err.message}`);
        }
      });
    }
  }
};
