/**
 * Khanna Travels & Holidays — Cover Letter Format Selection Module
 * Provides clean 3-option dropdown: Standard, Japan, Singapore.
 */

import { Api } from './api.js';
import { State } from './state.js';

const FORMAT_DESCRIPTIONS = {
  standard: 'Approved European/Schengen base format. Narrative corporate employment, Schengen multi-country itinerary breakdown, and return-to-India commitment.',
  japan: 'Approved Consulate General of Japan format. Features family student/homemaker narrative and a dedicated 3-Column Hotel Accommodation Table (NAME | DATE | CONTACT NO.).',
  singapore: 'Approved Consulate General of Singapore format. Features a dedicated 5-Column Passenger Table, underlined subject line, and multiple-entry visa request.'
};

export const TemplatesModule = {
  render(container) {
    const app = State.currentApplication;
    if (!app) return;

    const currentFormat = (app.selectedTemplateId || 'standard').toLowerCase();

    container.innerHTML = `
      <div style="max-width: 720px; margin: 0 auto;">
        <div style="margin-bottom: 28px; text-align: center;">
          <h2 style="font-size: 1.35rem; font-weight: 800; color: var(--dark); margin-bottom: 6px;">
            Step 5: Choose Cover Letter Format
          </h2>
          <p style="color: var(--muted); font-size: 0.88rem;">
            Select the approved consular cover letter format for this application.
          </p>
        </div>

        <div class="card" style="padding: 32px 28px; margin-bottom: 24px;">
          <div class="form-group" style="margin-bottom: 20px;">
            <label class="form-label" style="font-size: 0.92rem; font-weight: 700; margin-bottom: 8px;">
              Cover Letter Format
            </label>
            <select class="form-control" id="template-format-select" style="font-size: 1rem; padding: 12px 16px;">
              <option value="standard" ${currentFormat === 'standard' ? 'selected' : ''}>Standard</option>
              <option value="japan" ${currentFormat === 'japan' ? 'selected' : ''}>Japan</option>
              <option value="singapore" ${currentFormat === 'singapore' ? 'selected' : ''}>Singapore</option>
            </select>
          </div>

          <!-- Format Information Box -->
          <div id="format-info-box" style="background-color: #f8fafc; border: 1px solid var(--border); border-radius: var(--radius-sm); padding: 16px; margin-bottom: 16px;">
            <div style="font-weight: 700; font-size: 0.86rem; color: var(--dark); margin-bottom: 4px;" id="format-info-title">
              ${currentFormat === 'japan' ? 'Japan Consular Format' : currentFormat === 'singapore' ? 'Singapore Consular Format' : 'Standard European / Schengen Base Format'}
            </div>
            <p style="font-size: 0.84rem; color: var(--slate); line-height: 1.5;" id="format-info-desc">
              ${FORMAT_DESCRIPTIONS[currentFormat] || FORMAT_DESCRIPTIONS.standard}
            </p>
          </div>

          <div style="font-size: 0.8rem; color: var(--muted); display: flex; align-items: center; gap: 6px;">
            <span>Target Destination Country:</span>
            <strong style="color: var(--dark);">${app.travel?.destinationCountry || app.selectedCountry || 'France'}</strong>
          </div>
        </div>

        <!-- Navigation Buttons -->
        <div class="step-nav-bar">
          <button type="button" class="btn btn-secondary" id="btn-back-travel">
            ← Back to Travel Details
          </button>
          <button type="button" class="btn btn-primary" id="btn-generate-cover-letter">
            Generate Cover Letter →
          </button>
        </div>
      </div>
    `;

    this.bindEvents(container);
  },

  bindEvents(container) {
    const app = State.currentApplication;

    const select = container.querySelector('#template-format-select');
    const infoTitle = container.querySelector('#format-info-title');
    const infoDesc = container.querySelector('#format-info-desc');

    if (select) {
      select.addEventListener('change', (e) => {
        const val = e.target.value;
        app.selectedTemplateId = val;
        infoTitle.textContent = val === 'japan' ? 'Japan Consular Format' : val === 'singapore' ? 'Singapore Consular Format' : 'Standard European / Schengen Base Format';
        infoDesc.textContent = FORMAT_DESCRIPTIONS[val] || FORMAT_DESCRIPTIONS.standard;
      });
    }

    const btnBack = container.querySelector('#btn-back-travel');
    if (btnBack) btnBack.addEventListener('click', () => State.setStep(4));

    const btnGenerate = container.querySelector('#btn-generate-cover-letter');
    if (btnGenerate) {
      btnGenerate.addEventListener('click', async () => {
        btnGenerate.disabled = true;
        btnGenerate.textContent = 'Generating Cover Letter...';

        try {
          // 1. Save chosen format
          await Api.saveApplication(app);

          // 2. Generate cover letter
          const res = await Api.generateCoverLetter(app);
          app.generatedDocumentHtml = res.document.html;
          app.isDocumentManuallyEdited = false;
          State.setApplication(app);

          // 3. Go to Step 6: Cover Letter Editor
          State.setStep(6);
        } catch (err) {
          alert(`Failed to generate cover letter: ${err.message}`);
          btnGenerate.disabled = false;
          btnGenerate.textContent = 'Generate Cover Letter →';
        }
      });
    }
  }
};
