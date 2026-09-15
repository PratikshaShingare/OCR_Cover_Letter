/**
 * Khanna Travels & Holidays — Cover Letter Editor Module
 * Provides live editable A4 paper preview, rich formatting toolbar,
 * and a safe regeneration prompt modal.
 */

import { Api } from './api.js';
import { State } from './state.js';

export const CoverLetterModule = {
  render(container) {
    const app = State.currentApplication;
    if (!app) return;

    const htmlContent = app.generatedDocumentHtml || '<p>Generating document content...</p>';

    container.innerHTML = `
      <div class="editor-container">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
          <div>
            <h2 style="font-size: 1.35rem; font-weight: 800; color: var(--dark); margin-bottom: 4px;">
              Step 6: Review & Edit Cover Letter
            </h2>
            <p style="color: var(--muted); font-size: 0.86rem;">
              Click anywhere on the document below to edit text, dates, or tables. Formatting changes are saved.
            </p>
          </div>
          <button type="button" class="btn btn-secondary btn-sm" id="btn-regenerate-doc" style="color: var(--primary); border-color: #93c5fd;">
            ↻ Regenerate from Master Data
          </button>
        </div>

        <!-- Formatting Toolbar -->
        <div class="editor-toolbar">
          <button type="button" class="tool-btn" data-cmd="bold" title="Bold (Ctrl+B)"><strong>B</strong></button>
          <button type="button" class="tool-btn" data-cmd="italic" title="Italic (Ctrl+I)"><em>I</em></button>
          <button type="button" class="tool-btn" data-cmd="underline" title="Underline (Ctrl+U)"><u>U</u></button>
          <div class="toolbar-sep"></div>
          <button type="button" class="tool-btn" data-cmd="justifyLeft" title="Align Left">Left</button>
          <button type="button" class="tool-btn" data-cmd="justifyCenter" title="Align Center">Center</button>
          <button type="button" class="tool-btn" data-cmd="justifyRight" title="Align Right">Right</button>
          <button type="button" class="tool-btn" data-cmd="justifyFull" title="Justify">Justify</button>
          <div class="toolbar-sep"></div>
          <button type="button" class="tool-btn" data-cmd="insertUnorderedList" title="Bullet List">• List</button>
          <button type="button" class="tool-btn" data-cmd="insertOrderedList" title="Numbered List">1. List</button>
          <div class="toolbar-sep"></div>
          <button type="button" class="tool-btn" data-cmd="undo" title="Undo">↶ Undo</button>
          <button type="button" class="tool-btn" data-cmd="redo" title="Redo">↷ Redo</button>
        </div>

        <!-- A4 Paper Viewport -->
        <div class="a4-paper-wrapper">
          <div class="a4-paper" id="cover-letter-paper" contenteditable="true" spellcheck="false">
            ${htmlContent}
          </div>
        </div>

        <!-- Navigation Buttons -->
        <div class="step-nav-bar">
          <button type="button" class="btn btn-secondary" id="btn-back-template">
            ← Back to Format Selection
          </button>
          <button type="button" class="btn btn-primary" id="btn-proceed-review">
            Proceed to Final Output →
          </button>
        </div>
      </div>

      <!-- Regenerate Confirmation Modal -->
      <div id="regen-modal" class="modal-overlay hidden">
        <div class="modal-dialog">
          <div class="modal-title">Regenerate Cover Letter?</div>
          <div class="modal-body">
            Regenerating the cover letter will re-compile the document from Master Client Data and <strong>will overwrite any manual text edits</strong> you have made in this editor.
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary" id="btn-cancel-regen">Cancel</button>
            <button type="button" class="btn btn-danger" id="btn-confirm-regen">Yes, Regenerate</button>
          </div>
        </div>
      </div>
    `;

    this.bindEvents(container);
  },

  bindEvents(container) {
    const app = State.currentApplication;
    const paper = container.querySelector('#cover-letter-paper');

    // ExecCommand for formatting
    container.querySelectorAll('.tool-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        const cmd = btn.dataset.cmd;
        if (cmd) {
          document.execCommand(cmd, false, null);
          paper.focus();
        }
      });
    });

    // Mark as manually edited on typing
    if (paper) {
      paper.addEventListener('input', () => {
        app.isDocumentManuallyEdited = true;
        app.generatedDocumentHtml = paper.innerHTML;
      });
    }

    // Regeneration Warning Modal
    const regenModal = container.querySelector('#regen-modal');
    const btnRegenerate = container.querySelector('#btn-regenerate-doc');
    const btnCancelRegen = container.querySelector('#btn-cancel-regen');
    const btnConfirmRegen = container.querySelector('#btn-confirm-regen');

    if (btnRegenerate && regenModal) {
      btnRegenerate.addEventListener('click', () => {
        if (app.isDocumentManuallyEdited) {
          regenModal.classList.remove('hidden');
        } else {
          this.doRegenerate(container);
        }
      });
    }

    if (btnCancelRegen && regenModal) {
      btnCancelRegen.addEventListener('click', () => regenModal.classList.add('hidden'));
    }

    if (btnConfirmRegen && regenModal) {
      btnConfirmRegen.addEventListener('click', async () => {
        regenModal.classList.add('hidden');
        await this.doRegenerate(container);
      });
    }

    // Navigation
    const btnBack = container.querySelector('#btn-back-template');
    if (btnBack) btnBack.addEventListener('click', () => State.setStep(5));

    const btnNext = container.querySelector('#btn-proceed-review');
    if (btnNext) {
      btnNext.addEventListener('click', async () => {
        if (paper) {
          app.generatedDocumentHtml = paper.innerHTML;
        }
        await Api.saveApplication(app);
        State.setStep(7); // Go to Final Review & Output
      });
    }
  },

  async doRegenerate(container) {
    const app = State.currentApplication;
    try {
      const res = await Api.generateCoverLetter(app);
      app.generatedDocumentHtml = res.document.html;
      app.isDocumentManuallyEdited = false;
      State.setApplication(app);
      this.render(container);
    } catch (err) {
      alert(`Regeneration failed: ${err.message}`);
    }
  }
};
