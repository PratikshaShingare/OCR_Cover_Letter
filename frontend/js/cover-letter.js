/**
 * Khanna Travels & Holidays — Cover Letter Editor Module
 * Comprehensive Word-like document editor with text styling, fonts, alignments,
 * dynamic table operations (rows/cols), page breaks, and safe master data regeneration.
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
            Regenerate from Master Data
          </button>
        </div>

        <!-- Word-like Rich Formatting Toolbar -->
        <div class="editor-toolbar" style="display: flex; flex-wrap: wrap; gap: 4px; padding: 8px; background: #f8fafc; border: 1px solid var(--border); border-radius: 6px; margin-bottom: 12px; align-items: center;">
          <!-- Font Family -->
          <select class="tool-select" id="tool-font-name" title="Font Family" style="padding: 4px 8px; font-size: 0.82rem; border: 1px solid var(--border); border-radius: 4px; background: #fff;">
            <option value="Calibri">Calibri</option>
            <option value="Arial">Arial</option>
            <option value="'Times New Roman', serif">Times New Roman</option>
            <option value="Georgia, serif">Georgia</option>
            <option value="'Segoe UI', sans-serif">Segoe UI</option>
            <option value="Verdana, sans-serif">Verdana</option>
          </select>

          <!-- Font Size -->
          <select class="tool-select" id="tool-font-size" title="Font Size" style="padding: 4px 8px; font-size: 0.82rem; border: 1px solid var(--border); border-radius: 4px; background: #fff;">
            <option value="2">10 pt</option>
            <option value="3" selected>11 pt</option>
            <option value="4">12 pt</option>
            <option value="5">14 pt</option>
            <option value="6">18 pt</option>
          </select>

          <div class="toolbar-sep" style="width: 1px; height: 20px; background: #cbd5e1; margin: 0 4px;"></div>

          <!-- Basic Formatting -->
          <button type="button" class="tool-btn" data-cmd="bold" title="Bold (Ctrl+B)"><strong>B</strong></button>
          <button type="button" class="tool-btn" data-cmd="italic" title="Italic (Ctrl+I)"><em>I</em></button>
          <button type="button" class="tool-btn" data-cmd="underline" title="Underline (Ctrl+U)"><u>U</u></button>
          <button type="button" class="tool-btn" data-cmd="strikeThrough" title="Strikethrough"><s>S</s></button>

          <div class="toolbar-sep" style="width: 1px; height: 20px; background: #cbd5e1; margin: 0 4px;"></div>

          <!-- Color Pickers -->
          <label title="Text Color" style="display: flex; align-items: center; gap: 2px; cursor: pointer; font-size: 0.78rem; font-weight: 600; color: var(--dark); padding: 2px 4px;">
            A
            <input type="color" id="tool-text-color" value="#000000" style="width: 18px; height: 18px; border: none; padding: 0; cursor: pointer;" />
          </label>

          <label title="Highlight Color" style="display: flex; align-items: center; gap: 2px; cursor: pointer; font-size: 0.78rem; font-weight: 600; color: var(--dark); padding: 2px 4px;">
            <span style="background: #fef08a; padding: 0 2px;">H</span>
            <input type="color" id="tool-bg-color" value="#ffffff" style="width: 18px; height: 18px; border: none; padding: 0; cursor: pointer;" />
          </label>

          <div class="toolbar-sep" style="width: 1px; height: 20px; background: #cbd5e1; margin: 0 4px;"></div>

          <!-- Alignment -->
          <button type="button" class="tool-btn" data-cmd="justifyLeft" title="Align Left">Left</button>
          <button type="button" class="tool-btn" data-cmd="justifyCenter" title="Align Center">Center</button>
          <button type="button" class="tool-btn" data-cmd="justifyRight" title="Align Right">Right</button>
          <button type="button" class="tool-btn" data-cmd="justifyFull" title="Justify">Justify</button>

          <div class="toolbar-sep" style="width: 1px; height: 20px; background: #cbd5e1; margin: 0 4px;"></div>

          <!-- Lists -->
          <button type="button" class="tool-btn" data-cmd="insertUnorderedList" title="Bullet List">Bullets</button>
          <button type="button" class="tool-btn" data-cmd="insertOrderedList" title="Numbered List">Numbers</button>

          <div class="toolbar-sep" style="width: 1px; height: 20px; background: #cbd5e1; margin: 0 4px;"></div>

          <!-- Table Controls Dropdown / Buttons -->
          <button type="button" class="tool-btn" id="btn-insert-table" title="Insert 3-Column Table">+ Table</button>
          <button type="button" class="tool-btn" id="btn-add-table-row" title="Add Row Below">+ Row</button>
          <button type="button" class="tool-btn" id="btn-del-table-row" title="Delete Current Row">- Row</button>
          <button type="button" class="tool-btn" id="btn-add-table-col" title="Add Column Right">+ Col</button>
          <button type="button" class="tool-btn" id="btn-del-table-col" title="Delete Current Column">- Col</button>

          <div class="toolbar-sep" style="width: 1px; height: 20px; background: #cbd5e1; margin: 0 4px;"></div>

          <!-- Page Break & History -->
          <button type="button" class="tool-btn" id="btn-insert-pagebreak" title="Insert Page Break">Page Break</button>
          <button type="button" class="tool-btn" data-cmd="undo" title="Undo (Ctrl+Z)">Undo</button>
          <button type="button" class="tool-btn" data-cmd="redo" title="Redo (Ctrl+Y)">Redo</button>
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
            Back to Format Selection
          </button>
          <button type="button" class="btn btn-primary" id="btn-proceed-review">
            Proceed to Final Output
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

    // Standard ExecCommand formatting buttons
    container.querySelectorAll('.tool-btn[data-cmd]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.preventDefault();
        const cmd = btn.dataset.cmd;
        if (cmd) {
          document.execCommand(cmd, false, null);
          if (paper) paper.focus();
        }
      });
    });

    // Font family change
    const fontSelect = container.querySelector('#tool-font-name');
    if (fontSelect) {
      fontSelect.addEventListener('change', (e) => {
        document.execCommand('fontName', false, e.target.value);
        if (paper) paper.focus();
      });
    }

    // Font size change
    const sizeSelect = container.querySelector('#tool-font-size');
    if (sizeSelect) {
      sizeSelect.addEventListener('change', (e) => {
        document.execCommand('fontSize', false, e.target.value);
        if (paper) paper.focus();
      });
    }

    // Text color change
    const textColor = container.querySelector('#tool-text-color');
    if (textColor) {
      textColor.addEventListener('input', (e) => {
        document.execCommand('foreColor', false, e.target.value);
        if (paper) paper.focus();
      });
    }

    // Highlight background color change
    const bgColor = container.querySelector('#tool-bg-color');
    if (bgColor) {
      bgColor.addEventListener('input', (e) => {
        document.execCommand('hiliteColor', false, e.target.value);
        if (paper) paper.focus();
      });
    }

    // Helper: Find current selected cell and table
    const getActiveCellAndTable = () => {
      const sel = window.getSelection();
      if (!sel || !sel.anchorNode) return { cell: null, row: null, table: null };
      let node = sel.anchorNode.nodeType === 3 ? sel.anchorNode.parentNode : sel.anchorNode;
      const cell = node.closest('td, th');
      const row = node.closest('tr');
      const table = node.closest('table');
      return { cell, row, table };
    };

    // Insert Table
    const btnInsertTable = container.querySelector('#btn-insert-table');
    if (btnInsertTable) {
      btnInsertTable.addEventListener('click', () => {
        const tableHtml = `
          <table style="width: 100%; border-collapse: collapse; margin: 12px 0;">
            <thead>
              <tr style="background-color: #f1f5f9;">
                <th style="border: 1px solid #cbd5e1; padding: 6px 10px; font-weight: 600; text-align: left;">Item</th>
                <th style="border: 1px solid #cbd5e1; padding: 6px 10px; font-weight: 600; text-align: left;">Description</th>
                <th style="border: 1px solid #cbd5e1; padding: 6px 10px; font-weight: 600; text-align: left;">Details</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td style="border: 1px solid #cbd5e1; padding: 6px 10px;">1</td>
                <td style="border: 1px solid #cbd5e1; padding: 6px 10px;">Sample Data</td>
                <td style="border: 1px solid #cbd5e1; padding: 6px 10px;">Confirmed</td>
              </tr>
              <tr>
                <td style="border: 1px solid #cbd5e1; padding: 6px 10px;">2</td>
                <td style="border: 1px solid #cbd5e1; padding: 6px 10px;">Sample Data</td>
                <td style="border: 1px solid #cbd5e1; padding: 6px 10px;">Confirmed</td>
              </tr>
            </tbody>
          </table>
          <p></p>
        `;
        document.execCommand('insertHTML', false, tableHtml);
        if (paper) paper.focus();
      });
    }

    // Add Row Below
    const btnAddRow = container.querySelector('#btn-add-table-row');
    if (btnAddRow) {
      btnAddRow.addEventListener('click', () => {
        const { row, table } = getActiveCellAndTable();
        if (!row || !table) {
          alert('Click inside a table cell to add a row.');
          return;
        }
        const colCount = row.cells.length;
        const newRow = document.createElement('tr');
        for (let i = 0; i < colCount; i++) {
          const newCell = document.createElement('td');
          newCell.style.border = '1px solid #cbd5e1';
          newCell.style.padding = '6px 10px';
          newCell.innerHTML = '&nbsp;';
          newRow.appendChild(newCell);
        }
        row.parentNode.insertBefore(newRow, row.nextSibling);
      });
    }

    // Delete Row
    const btnDelRow = container.querySelector('#btn-del-table-row');
    if (btnDelRow) {
      btnDelRow.addEventListener('click', () => {
        const { row, table } = getActiveCellAndTable();
        if (!row || !table) {
          alert('Click inside a table row to delete it.');
          return;
        }
        row.remove();
      });
    }

    // Add Column Right
    const btnAddCol = container.querySelector('#btn-add-table-col');
    if (btnAddCol) {
      btnAddCol.addEventListener('click', () => {
        const { cell, row, table } = getActiveCellAndTable();
        if (!cell || !row || !table) {
          alert('Click inside a table column to add a column.');
          return;
        }
        const cellIdx = cell.cellIndex;
        Array.from(table.rows).forEach(r => {
          const isHeader = r.parentNode.tagName === 'THEAD' || r.cells[0]?.tagName === 'TH';
          const newCell = document.createElement(isHeader ? 'th' : 'td');
          newCell.style.border = '1px solid #cbd5e1';
          newCell.style.padding = '6px 10px';
          newCell.innerHTML = isHeader ? 'New Header' : '&nbsp;';
          if (cellIdx + 1 < r.cells.length) {
            r.insertBefore(newCell, r.cells[cellIdx + 1]);
          } else {
            r.appendChild(newCell);
          }
        });
      });
    }

    // Delete Column
    const btnDelCol = container.querySelector('#btn-del-table-col');
    if (btnDelCol) {
      btnDelCol.addEventListener('click', () => {
        const { cell, table } = getActiveCellAndTable();
        if (!cell || !table) {
          alert('Click inside a table column to delete it.');
          return;
        }
        const cellIdx = cell.cellIndex;
        Array.from(table.rows).forEach(r => {
          if (r.cells[cellIdx]) {
            r.cells[cellIdx].remove();
          }
        });
      });
    }

    // Insert Page Break
    const btnPageBreak = container.querySelector('#btn-insert-pagebreak');
    if (btnPageBreak) {
      btnPageBreak.addEventListener('click', () => {
        const breakHtml = `
          <div class="doc-page-break" style="page-break-after: always; border-top: 2px dashed #94a3b8; margin: 24px 0; text-align: center; color: #94a3b8; font-size: 11px; font-weight: 600; user-select: none;">
            --- PAGE BREAK ---
          </div>
          <p></p>
        `;
        document.execCommand('insertHTML', false, breakHtml);
        if (paper) paper.focus();
      });
    }

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

