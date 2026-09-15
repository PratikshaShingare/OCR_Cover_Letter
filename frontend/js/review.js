/**
 * Khanna Travels & Holidays — Review & Final Output Module
 * Provides binary DOCX and PDF downloads. Zero Excel buttons in UI.
 */

import { Api } from './api.js';
import { State } from './state.js';

export const ReviewModule = {
  render(container) {
    const app = State.currentApplication;
    if (!app) return;

    container.innerHTML = `
      <div style="max-width: 680px; margin: 20px auto 0;">
        <div class="card" style="text-align: center; padding: 48px 36px;">
          <!-- Success Badge Icon -->
          <div style="width: 64px; height: 64px; border-radius: 50%; background-color: var(--success-light); color: var(--success); display: flex; align-items: center; justify-content: center; margin: 0 auto 20px; box-shadow: 0 4px 12px rgba(5, 150, 105, 0.2);">
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"></polyline></svg>
          </div>

          <h2 style="font-size: 1.5rem; font-weight: 800; color: var(--dark); margin-bottom: 6px;">
            Cover Letter Ready
          </h2>
          <p style="color: var(--muted); font-size: 0.92rem; margin-bottom: 32px;">
            Your country-specific visa cover letter has been compiled from verified master client data.
          </p>

          <!-- Summary Meta Box -->
          <div style="background-color: #f8fafc; border: 1px solid var(--border); border-radius: var(--radius-md); padding: 18px 24px; text-align: left; margin-bottom: 32px;">
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 14px; font-size: 0.86rem;">
              <div>
                <span style="color: var(--muted); display: block; font-size: 0.76rem; text-transform: uppercase; font-weight: 600;">Application ID</span>
                <strong style="color: var(--dark);">${app.applicationId}</strong>
              </div>
              <div>
                <span style="color: var(--muted); display: block; font-size: 0.76rem; text-transform: uppercase; font-weight: 600;">Primary Applicant</span>
                <strong style="color: var(--dark);">${app.personal?.fullName || 'Applicant'}</strong>
              </div>
              <div>
                <span style="color: var(--muted); display: block; font-size: 0.76rem; text-transform: uppercase; font-weight: 600;">Destination Country</span>
                <strong style="color: var(--dark);">${app.selectedCountry || 'Visa'}</strong>
              </div>
              <div>
                <span style="color: var(--muted); display: block; font-size: 0.76rem; text-transform: uppercase; font-weight: 600;">Cover Letter Format</span>
                <strong style="color: var(--dark);">${(app.selectedTemplateId || 'Standard').toUpperCase()}</strong>
              </div>
            </div>
          </div>

          <!-- Primary Download Actions (DOCX & PDF ONLY) -->
          <div style="display: flex; flex-direction: column; gap: 14px; max-width: 400px; margin: 0 auto 28px;">
            <button type="button" class="btn btn-primary btn-lg" id="btn-download-docx" style="display: flex; align-items: center; justify-content: center; gap: 10px;">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
              Download DOCX (Word Document)
            </button>

            <button type="button" class="btn btn-secondary btn-lg" id="btn-download-pdf" style="display: flex; align-items: center; justify-content: center; gap: 10px; border-color: #cbd5e1;">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="12" y1="18" x2="12" y2="12"></line><line x1="9" y1="15" x2="15" y2="15"></line></svg>
              Download PDF (Vector Document)
            </button>
          </div>

          <!-- Secondary Actions -->
          <div style="display: flex; justify-content: center; gap: 16px;">
            <button type="button" class="btn btn-secondary btn-sm" id="btn-edit-doc">
              ✎ Edit Document
            </button>
            <button type="button" class="btn btn-secondary btn-sm" id="btn-finish-dashboard">
              Return to Applications
            </button>
          </div>
        </div>
      </div>
    `;

    this.bindEvents(container);
  },

  bindEvents(container) {
    const app = State.currentApplication;

    const btnDocx = container.querySelector('#btn-download-docx');
    if (btnDocx) {
      btnDocx.addEventListener('click', async () => {
        btnDocx.disabled = true;
        try {
          await Api.downloadDocx(app);
        } catch (err) {
          alert(`DOCX Download failed: ${err.message}`);
        } finally {
          btnDocx.disabled = false;
        }
      });
    }

    const btnPdf = container.querySelector('#btn-download-pdf');
    if (btnPdf) {
      btnPdf.addEventListener('click', async () => {
        btnPdf.disabled = true;
        try {
          await Api.downloadPdf(app);
        } catch (err) {
          alert(`PDF Download failed: ${err.message}`);
        } finally {
          btnPdf.disabled = false;
        }
      });
    }

    const btnEdit = container.querySelector('#btn-edit-doc');
    if (btnEdit) {
      btnEdit.addEventListener('click', () => State.setStep(6));
    }

    const btnFinish = container.querySelector('#btn-finish-dashboard');
    if (btnFinish) {
      btnFinish.addEventListener('click', () => State.setViewMode('dashboard'));
    }
  }
};
