/**
 * Khanna Travels & Holidays — Travel, Accommodation & Financial Module
 * Handles trip duration calculation, dynamic hotel bookings, financial sponsorship,
 * and executive notes.
 */

import { Api } from './api.js';
import { State } from './state.js';

export const TravelModule = {
  render(container) {
    const app = State.currentApplication;
    if (!app) return;

    const tr = app.travel || {};
    const fin = app.financial || {};
    const add = app.additional || {};
    if (!app.hotels) app.hotels = [];

    container.innerHTML = `
      <div style="max-width: 960px; margin: 0 auto;">
        <div style="margin-bottom: 24px;">
          <h2 style="font-size: 1.35rem; font-weight: 800; color: var(--dark); margin-bottom: 6px;">
            Step 4: Travel, Accommodation & Financials
          </h2>
          <p style="color: var(--muted); font-size: 0.88rem;">
            Provide itinerary schedule, confirmed accommodation, and trip sponsorship details.
          </p>
        </div>

        <form id="travel-form">
          <!-- 1. TRIP DETAILS -->
          <div class="card" style="margin-bottom: 24px;">
            <h3 style="font-size: 1.05rem; font-weight: 700; color: var(--dark); margin-bottom: 16px; border-bottom: 1px solid var(--border); padding-bottom: 8px;">
              1. Trip & Itinerary Specifications
            </h3>

            <div class="form-grid">
              <div class="form-group">
                <label class="form-label">Destination Country</label>
                <input type="text" class="form-control" name="travel.destinationCountry" id="dest-country-input" value="${tr.destinationCountry || app.selectedCountry || 'France'}" placeholder="e.g. France, Japan, Singapore, Switzerland" />
              </div>

              <div class="form-group">
                <label class="form-label">Visa Category</label>
                <select class="form-control" name="travel.visaType">
                  <option value="Tourism" ${tr.visaType === 'Tourism' ? 'selected' : ''}>Tourist Visa</option>
                  <option value="Business" ${tr.visaType === 'Business' ? 'selected' : ''}>Business Visa</option>
                  <option value="Transit" ${tr.visaType === 'Transit' ? 'selected' : ''}>Transit Visa</option>
                </select>
              </div>

              <div class="form-group">
                <label class="form-label">Purpose of Travel</label>
                <input type="text" class="form-control" name="travel.purposeOfTravel" value="${tr.purposeOfTravel || 'Tourism'}" placeholder="e.g. Tourism, Holiday, Leisure" />
              </div>

              <div class="form-group">
                <label class="form-label">Travel Start Date (Departure)</label>
                <input type="date" class="form-control" name="travel.travelStartDate" id="date-start" value="${tr.travelStartDate || ''}" />
              </div>

              <div class="form-group">
                <label class="form-label">Travel End Date (Return)</label>
                <input type="date" class="form-control" name="travel.travelEndDate" id="date-end" value="${tr.travelEndDate || ''}" />
              </div>

              <div class="form-group">
                <label class="form-label">Calculated Duration</label>
                <div style="display: flex; gap: 8px;">
                  <input type="text" class="form-control" id="days-display" value="${tr.numberOfDays ? tr.numberOfDays + ' Days' : '—'}" readonly />
                  <input type="text" class="form-control" id="nights-display" value="${tr.numberOfNights ? tr.numberOfNights + ' Nights' : '—'}" readonly />
                </div>
              </div>

              <div class="form-group">
                <label class="form-label">Entry Type</label>
                <select class="form-control" name="travel.entryType">
                  <option value="Single" ${tr.entryType === 'Single' ? 'selected' : ''}>Single Entry</option>
                  <option value="Multiple" ${tr.entryType === 'Multiple' ? 'selected' : ''}>Multiple Entries</option>
                  <option value="Double" ${tr.entryType === 'Double' ? 'selected' : ''}>Double Entry</option>
                </select>
              </div>

              <div class="form-group">
                <label class="form-label">Countries to be Visited</label>
                <input type="text" class="form-control" name="travel.countriesToBeVisited" value="${tr.countriesToBeVisited || ''}" placeholder="e.g. France, Switzerland" />
              </div>
            </div>
          </div>

          <!-- 2. ACCOMMODATION / HOTELS -->
          <div class="card" style="margin-bottom: 24px;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; border-bottom: 1px solid var(--border); padding-bottom: 8px;">
              <h3 style="font-size: 1.05rem; font-weight: 700; color: var(--dark);">
                2. Accommodation & Hotel Bookings
              </h3>
              <button type="button" class="btn btn-secondary btn-sm" id="btn-add-hotel">
                + Add Hotel
              </button>
            </div>

            <div id="hotels-container">
              ${app.hotels.length === 0 ? `
                <div style="text-align: center; padding: 24px; color: var(--muted); background-color: #fafafa; border-radius: 6px;">
                  No hotels added yet. Click <strong>+ Add Hotel</strong> to add stay details.
                </div>
              ` : app.hotels.map((h, idx) => `
                <div class="hotel-item" data-index="${idx}" style="border: 1px solid var(--border); border-radius: 8px; padding: 16px; margin-bottom: 12px; background-color: #ffffff;">
                  <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                    <strong style="font-size: 0.9rem; color: var(--dark);">Hotel ${idx + 1}: ${h.hotelName || 'Pending name'}</strong>
                    <button type="button" class="btn btn-danger btn-sm btn-remove-hotel" data-index="${idx}">Remove</button>
                  </div>
                  <div class="form-grid">
                    <div class="form-group">
                      <label class="form-label">Hotel Name</label>
                      <input type="text" class="form-control h-field" data-field="hotelName" data-index="${idx}" value="${h.hotelName || ''}" placeholder="e.g. Tokyo Bay Hotel" />
                    </div>
                    <div class="form-group">
                      <label class="form-label">City</label>
                      <input type="text" class="form-control h-field" data-field="city" data-index="${idx}" value="${h.city || ''}" placeholder="City" />
                    </div>
                    <div class="form-group">
                      <label class="form-label">Check-in Date</label>
                      <input type="date" class="form-control h-field" data-field="checkInDate" data-index="${idx}" value="${h.checkInDate || ''}" />
                    </div>
                    <div class="form-group">
                      <label class="form-label">Check-out Date</label>
                      <input type="date" class="form-control h-field" data-field="checkOutDate" data-index="${idx}" value="${h.checkOutDate || ''}" />
                    </div>
                    <div class="form-group">
                      <label class="form-label">Contact Number</label>
                      <input type="text" class="form-control h-field" data-field="contactNumber" data-index="${idx}" value="${h.contactNumber || ''}" placeholder="Phone number" />
                    </div>
                    <div class="form-group">
                      <label class="form-label">Booking Reference / PNR</label>
                      <input type="text" class="form-control h-field" data-field="bookingReference" data-index="${idx}" value="${h.bookingReference || ''}" placeholder="Confirmation No." />
                    </div>
                  </div>
                </div>
              `).join('')}
            </div>
          </div>

          <!-- 3. FINANCIAL INFORMATION -->
          <div class="card" style="margin-bottom: 24px;">
            <h3 style="font-size: 1.05rem; font-weight: 700; color: var(--dark); margin-bottom: 16px; border-bottom: 1px solid var(--border); padding-bottom: 8px;">
              3. Trip Financial Sponsorship
            </h3>

            <div class="form-grid" style="margin-bottom: 16px;">
              <div class="form-group">
                <label class="form-label">Trip Sponsor</label>
                <select class="form-control" name="financial.tripSponsor">
                  <option value="Self-funded" ${fin.tripSponsor === 'Self-funded' ? 'selected' : ''}>Self-funded</option>
                  <option value="Jointly funded" ${fin.tripSponsor === 'Jointly funded' ? 'selected' : ''}>Jointly funded (Couple)</option>
                  <option value="Company funded" ${fin.tripSponsor === 'Company funded' ? 'selected' : ''}>Company / Employer funded</option>
                  <option value="Spouse funded" ${fin.tripSponsor === 'Spouse funded' ? 'selected' : ''}>Spouse funded</option>
                  <option value="Parent funded" ${fin.tripSponsor === 'Parent funded' ? 'selected' : ''}>Parent funded</option>
                  <option value="Other sponsor" ${fin.tripSponsor === 'Other sponsor' ? 'selected' : ''}>Other Sponsor</option>
                </select>
              </div>

              <div class="form-group">
                <label class="form-label">Sponsor Name & Relation (if not self)</label>
                <input type="text" class="form-control" name="financial.sponsorName" value="${fin.sponsorName || ''}" placeholder="e.g. Acme Corp / Father" />
              </div>
            </div>

            <div style="font-weight: 600; font-size: 0.84rem; color: #334155; margin-bottom: 10px;">
              Confirmed Supporting Financial Documents Attached:
            </div>
            <div style="display: flex; flex-wrap: wrap; gap: 20px;">
              <label class="checkbox-label">
                <input type="checkbox" name="financial.bankStatementAvailable" ${fin.bankStatementAvailable ? 'checked' : ''} />
                <span>Bank Statement</span>
              </label>
              <label class="checkbox-label">
                <input type="checkbox" name="financial.itrAvailable" ${fin.itrAvailable ? 'checked' : ''} />
                <span>ITR Acknowledgements (2-3 Yrs)</span>
              </label>
              <label class="checkbox-label">
                <input type="checkbox" name="financial.salarySlipsAvailable" ${fin.salarySlipsAvailable ? 'checked' : ''} />
                <span>Salary Slips (Last 3 Months)</span>
              </label>
              <label class="checkbox-label">
                <input type="checkbox" name="financial.employmentLetterAvailable" ${fin.employmentLetterAvailable ? 'checked' : ''} />
                <span>Employer NOC / Leave Approval Letter</span>
              </label>
            </div>
          </div>

          <!-- 4. ADDITIONAL EXECUTIVE NOTES -->
          <div class="card" style="margin-bottom: 24px;">
            <h3 style="font-size: 1.05rem; font-weight: 700; color: var(--dark); margin-bottom: 16px; border-bottom: 1px solid var(--border); padding-bottom: 8px;">
              4. Additional Information & Notes
            </h3>

            <div class="form-group" style="margin-bottom: 12px;">
              <label class="form-label">Executive Notes / Special Circumstances</label>
              <textarea class="form-control" name="additional.content" placeholder="Any special notes for this visa application...">${add.content || ''}</textarea>
            </div>

            <label class="checkbox-label">
              <input type="checkbox" name="additional.includeInCoverLetter" ${add.includeInCoverLetter ? 'checked' : ''} />
              <span>Include these notes as an additional paragraph in the cover letter</span>
            </label>
          </div>

          <!-- Navigation Buttons -->
          <div class="step-nav-bar">
            <button type="button" class="btn btn-secondary" id="btn-back-travellers">
              ← Back to Travellers
            </button>
            <button type="submit" class="btn btn-primary" id="btn-save-travel">
              Save & Continue to Template Selection →
            </button>
          </div>
        </form>
      </div>
    `;

    this.bindEvents(container);
  },

  bindEvents(container) {
    const app = State.currentApplication;

    // Date duration calculation
    const startInput = container.querySelector('#date-start');
    const endInput = container.querySelector('#date-end');
    const daysDisplay = container.querySelector('#days-display');
    const nightsDisplay = container.querySelector('#nights-display');

    const updateDuration = () => {
      const s = startInput.value;
      const e = endInput.value;
      if (s && e) {
        const d1 = new Date(s);
        const d2 = new Date(e);
        const diffMs = d2 - d1;
        if (diffMs >= 0) {
          const days = Math.round(diffMs / (1000 * 60 * 60 * 24)) + 1;
          const nights = Math.max(days - 1, 0);
          daysDisplay.value = `${days} Days`;
          nightsDisplay.value = `${nights} Nights`;
          app.travel.numberOfDays = days;
          app.travel.numberOfNights = nights;
          return;
        }
      }
      daysDisplay.value = '—';
      nightsDisplay.value = '—';
      app.travel.numberOfDays = 0;
      app.travel.numberOfNights = 0;
    };

    if (startInput) startInput.addEventListener('change', updateDuration);
    if (endInput) endInput.addEventListener('change', updateDuration);

    // Hotel management
    const btnAddHotel = container.querySelector('#btn-add-hotel');
    if (btnAddHotel) {
      btnAddHotel.addEventListener('click', () => {
        if (!app.hotels) app.hotels = [];
        app.hotels.push({
          id: `H-${Date.now()}`,
          hotelName: '',
          address: '',
          city: '',
          country: app.travel.destinationCountry || 'France',
          checkInDate: app.travel.travelStartDate || '',
          checkOutDate: app.travel.travelEndDate || '',
          contactNumber: '',
          bookingReference: ''
        });
        this.render(container);
      });
    }

    container.querySelectorAll('.h-field').forEach(input => {
      input.addEventListener('input', (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        const field = e.target.dataset.field;
        if (app.hotels && app.hotels[idx]) {
          app.hotels[idx][field] = e.target.value.trim();
        }
      });
    });

    container.querySelectorAll('.btn-remove-hotel').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        app.hotels.splice(idx, 1);
        this.render(container);
      });
    });

    // Navigation
    const btnBack = container.querySelector('#btn-back-travellers');
    if (btnBack) btnBack.addEventListener('click', () => State.setStep(3));

    // Form Submit / Save
    const form = container.querySelector('#travel-form');
    if (form) {
      form.addEventListener('submit', async (e) => {
        e.preventDefault();
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

        // Handle checkboxes explicitly
        app.financial.bankStatementAvailable = form.querySelector('input[name="financial.bankStatementAvailable"]').checked;
        app.financial.itrAvailable = form.querySelector('input[name="financial.itrAvailable"]').checked;
        app.financial.salarySlipsAvailable = form.querySelector('input[name="financial.salarySlipsAvailable"]').checked;
        app.financial.employmentLetterAvailable = form.querySelector('input[name="financial.employmentLetterAvailable"]').checked;
        app.additional.includeInCoverLetter = form.querySelector('input[name="additional.includeInCoverLetter"]').checked;

        // Keep selectedCountry aligned with destinationCountry
        if (app.travel.destinationCountry) {
          app.selectedCountry = app.travel.destinationCountry;
        }

        try {
          const saved = await Api.saveApplication(app);
          State.setApplication(saved);
          State.setStep(5); // Go to Template Selection
        } catch (err) {
          alert(`Error saving travel details: ${err.message}`);
        }
      });
    }
  }
};
