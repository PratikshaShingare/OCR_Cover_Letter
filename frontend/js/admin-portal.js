/**
 * Khanna Travels & Holidays — Admin Portal & Bulk Data Import
 * Provides executive tools for importing client records from .xlsx, .xls, and .csv files,
 * interactive column mapping, duplicate resolution strategies, and audit logging.
 */

import { Api } from './api.js';
import { State } from './state.js';

export const AdminPortalModule = {
  uploadData: null,

  async render(container) {
    container.innerHTML = `
      <div style="max-width: 980px; margin: 24px auto 40px;">
        <!-- Top Navigation -->
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
          <button type="button" class="btn btn-secondary btn-sm" id="btn-admin-back">
            &larr; Back to Applications
          </button>
          <div style="display: flex; gap: 10px;">
            <button type="button" class="btn btn-secondary btn-sm" id="btn-export-all-clients">
              Download All Clients Export (.xlsx)
            </button>
          </div>
        </div>

        <!-- Header Card -->
        <div class="card" style="padding: 28px 32px; margin-bottom: 24px;">
          <div style="display: flex; align-items: center; gap: 16px;">
            <div style="width: 48px; height: 48px; border-radius: 10px; background-color: var(--primary); color: #ffffff; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">
              <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="17 8 12 3 7 8"></polyline><line x1="12" y1="3" x2="12" y2="15"></line></svg>
            </div>
            <div>
              <h2 style="font-size: 1.35rem; font-weight: 800; color: var(--dark); margin: 0 0 4px;">
                ADMIN PORTAL &amp; BULK DATA INGESTION
              </h2>
              <p style="font-size: 0.88rem; color: var(--muted); margin: 0;">
                Import bulk client records from Excel (.xlsx, .xls) or CSV files into the centralized Khanna Travels database.
              </p>
            </div>
          </div>
        </div>

        <!-- Section 1: File Upload Dropzone -->
        <div class="card" style="padding: 28px; margin-bottom: 24px;">
          <h3 style="font-size: 1.05rem; font-weight: 700; color: var(--dark); margin-bottom: 14px;">
            1. Upload Spreadsheet or CSV File
          </h3>

          <div id="admin-dropzone" style="border: 2px dashed #cbd5e1; border-radius: 8px; padding: 36px 20px; text-align: center; background-color: #f8fafc; cursor: pointer; transition: border-color 0.2s;">
            <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2" style="margin: 0 auto 12px;"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="12" y1="18" x2="12" y2="12"></line><line x1="9" y1="15" x2="15" y2="15"></line></svg>
            <div style="font-weight: 600; color: var(--dark); font-size: 0.95rem; margin-bottom: 4px;">
              Click to select or drag and drop file here
            </div>
            <div style="font-size: 0.82rem; color: var(--muted);">
              Supported formats: .xlsx, .xls, .csv (Maximum file size: 20MB)
            </div>
            <input type="file" id="admin-file-input" accept=".xlsx,.xls,.csv" style="display: none;" />
          </div>

          <div id="admin-upload-status" style="margin-top: 14px; font-size: 0.88rem; display: none;"></div>
        </div>

        <!-- Section 2: Mapping & Configuration (Hidden until file uploaded) -->
        <div id="admin-mapping-section" class="card" style="padding: 28px; margin-bottom: 24px; display: none;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
            <div>
              <h3 style="font-size: 1.05rem; font-weight: 700; color: var(--dark); margin: 0 0 4px;">
                2. Verify Column Mapping
              </h3>
              <p style="font-size: 0.84rem; color: var(--muted); margin: 0;" id="mapping-file-summary"></p>
            </div>
          </div>

          <div style="overflow-x: auto; margin-bottom: 24px;">
            <table class="apps-table" id="mapping-table">
              <thead>
                <tr>
                  <th style="width: 25%;">Spreadsheet Column</th>
                  <th style="width: 35%;">Sample Value (Row 1)</th>
                  <th style="width: 40%;">Target System Field</th>
                </tr>
              </thead>
              <tbody id="mapping-table-body">
                <!-- Generated dynamically -->
              </tbody>
            </table>
          </div>

          <h3 style="font-size: 1.05rem; font-weight: 700; color: var(--dark); margin-bottom: 12px;">
            3. Duplicate Resolution Strategy
          </h3>

          <div style="display: flex; flex-direction: column; gap: 10px; margin-bottom: 24px; background: #f8fafc; padding: 16px; border-radius: 8px; border: 1px solid var(--border);">
            <label style="display: flex; align-items: center; gap: 10px; cursor: pointer; font-size: 0.9rem;">
              <input type="radio" name="duplicate-strategy" value="update" checked />
              <span><strong>Update existing records:</strong> Update application if passport number already exists in database</span>
            </label>
            <label style="display: flex; align-items: center; gap: 10px; cursor: pointer; font-size: 0.9rem;">
              <input type="radio" name="duplicate-strategy" value="skip" />
              <span><strong>Skip duplicate records:</strong> Keep existing database records and ignore duplicate rows</span>
            </label>
            <label style="display: flex; align-items: center; gap: 10px; cursor: pointer; font-size: 0.9rem;">
              <input type="radio" name="duplicate-strategy" value="create_new" />
              <span><strong>Create new records:</strong> Always generate a new application ID for every row</span>
            </label>
          </div>

          <div style="display: flex; gap: 12px; align-items: center;">
            <button type="button" class="btn btn-primary btn-lg" id="btn-execute-import">
              Execute Bulk Ingestion
            </button>
            <span id="execute-status" style="font-size: 0.88rem; color: var(--muted);"></span>
          </div>
        </div>

        <!-- Section 3: Execution Summary Card (Hidden until executed) -->
        <div id="admin-summary-card" class="card" style="padding: 24px; margin-bottom: 24px; display: none; background-color: #f0fdf4; border-color: #86efac;">
          <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 12px;">
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#16a34a" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>
            <h3 style="font-size: 1.05rem; font-weight: 700; color: #166534; margin: 0;">
              Bulk Import Successfully Completed
            </h3>
          </div>
          <div id="summary-details" style="font-size: 0.9rem; color: #15803d; line-height: 1.6; margin-bottom: 16px;"></div>
          <button type="button" class="btn btn-secondary btn-sm" id="btn-view-imported-apps">
            View Applications in Dashboard &rarr;
          </button>
        </div>

        <!-- Section 4: Historical Import Audit Log -->
        <div class="card" style="padding: 24px;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
            <h3 style="font-size: 1.05rem; font-weight: 700; color: var(--dark); margin: 0;">
              Import History &amp; Audit Log
            </h3>
            <button type="button" class="btn btn-secondary btn-sm" id="btn-refresh-history">
              Refresh History
            </button>
          </div>

          <div id="history-container">
            <div style="text-align: center; padding: 24px; color: var(--muted); font-size: 0.88rem;">
              Loading audit logs...
            </div>
          </div>
        </div>

      </div>
    `;

    this.bindEvents(container);
    await this.loadHistory(container);
  },

  bindEvents(container) {
    const btnBack = container.querySelector('#btn-admin-back');
    if (btnBack) {
      btnBack.onclick = () => State.setViewMode('dashboard');
    }

    const btnExportAll = container.querySelector('#btn-export-all-clients');
    if (btnExportAll) {
      btnExportAll.onclick = async () => {
        try {
          btnExportAll.disabled = true;
          btnExportAll.textContent = 'Generating...';
          await Api.downloadAllClientsExport();
        } catch (err) {
          alert('Export failed: ' + err.message);
        } finally {
          btnExportAll.disabled = false;
          btnExportAll.textContent = 'Download All Clients Export (.xlsx)';
        }
      };
    }

    const dropzone = container.querySelector('#admin-dropzone');
    const fileInput = container.querySelector('#admin-file-input');

    if (dropzone && fileInput) {
      dropzone.onclick = () => fileInput.click();

      dropzone.ondragover = (e) => {
        e.preventDefault();
        dropzone.style.borderColor = 'var(--primary)';
        dropzone.style.backgroundColor = '#eff6ff';
      };

      dropzone.ondragleave = (e) => {
        e.preventDefault();
        dropzone.style.borderColor = '#cbd5e1';
        dropzone.style.backgroundColor = '#f8fafc';
      };

      dropzone.ondrop = (e) => {
        e.preventDefault();
        dropzone.style.borderColor = '#cbd5e1';
        dropzone.style.backgroundColor = '#f8fafc';
        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
          this.handleFileUpload(container, e.dataTransfer.files[0]);
        }
      };

      fileInput.onchange = (e) => {
        if (e.target.files && e.target.files[0]) {
          this.handleFileUpload(container, e.target.files[0]);
        }
      };
    }

    const btnExecute = container.querySelector('#btn-execute-import');
    if (btnExecute) {
      btnExecute.onclick = () => this.executeImport(container);
    }

    const btnRefreshHistory = container.querySelector('#btn-refresh-history');
    if (btnRefreshHistory) {
      btnRefreshHistory.onclick = () => this.loadHistory(container);
    }

    const btnViewApps = container.querySelector('#btn-view-imported-apps');
    if (btnViewApps) {
      btnViewApps.onclick = () => State.setViewMode('dashboard');
    }
  },

  async handleFileUpload(container, file) {
    const statusDiv = container.querySelector('#admin-upload-status');
    const mappingSection = container.querySelector('#admin-mapping-section');
    const summaryCard = container.querySelector('#admin-summary-card');

    summaryCard.style.display = 'none';
    statusDiv.style.display = 'block';
    statusDiv.innerHTML = '<span style="color: var(--primary);">Analyzing file headers and rows...</span>';

    try {
      const data = await Api.uploadBulkImportFile(file);
      this.uploadData = data;

      statusDiv.innerHTML = `<span style="color: #16a34a;">Loaded <strong>${data.filename}</strong> (${data.totalRows} rows detected)</span>`;
      mappingSection.style.display = 'block';

      const fileSummary = container.querySelector('#mapping-file-summary');
      if (fileSummary) {
        fileSummary.textContent = `${data.filename} • ${data.totalRows} records found • Review column assignments below`;
      }

      this.renderMappingTable(container, data);
    } catch (err) {
      statusDiv.innerHTML = `<span style="color: var(--danger);">Upload failed: ${err.message}</span>`;
      mappingSection.style.display = 'none';
    }
  },

  renderMappingTable(container, data) {
    const tbody = container.querySelector('#mapping-table-body');
    if (!tbody) return;

    const sampleRow = (data.sampleRows && data.sampleRows[0]) || {};
    const targetFields = data.targetFields || [];
    const detected = data.detectedMappings || {};

    const optionsHtml = [
      '<option value="">-- (Ignore / Do Not Import) --</option>',
      ...targetFields.map(f => `<option value="${f.key}">${f.label}${f.required ? ' *' : ''}</option>`)
    ].join('');

    tbody.innerHTML = data.headers.map(header => {
      const suggestedField = detected[header] || '';
      const sampleVal = sampleRow[header] !== undefined ? String(sampleRow[header]) : '';

      return `
        <tr>
          <td>
            <strong>${header}</strong>
          </td>
          <td style="color: var(--muted); font-size: 0.84rem; max-width: 260px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">
            ${sampleVal || '<span style="color: #94a3b8;">(empty)</span>'}
          </td>
          <td>
            <select class="input-field mapping-select" data-header="${header}" style="width: 100%; font-size: 0.86rem; padding: 6px 10px;">
              ${optionsHtml}
            </select>
          </td>
        </tr>
      `;
    }).join('');

    // Pre-select detected options
    tbody.querySelectorAll('.mapping-select').forEach(sel => {
      const h = sel.dataset.header;
      if (detected[h]) {
        sel.value = detected[h];
      }
    });
  },

  async executeImport(container) {
    if (!this.uploadData || !this.uploadData.tempId) {
      alert('Please upload a file first.');
      return;
    }

    const selects = container.querySelectorAll('.mapping-select');
    const mappings = {};
    selects.forEach(sel => {
      const h = sel.dataset.header;
      const val = sel.value;
      if (val) {
        mappings[h] = val;
      }
    });

    const strategyEl = container.querySelector('input[name="duplicate-strategy"]:checked');
    const duplicateStrategy = strategyEl ? strategyEl.value : 'update';

    const btnExecute = container.querySelector('#btn-execute-import');
    const statusEl = container.querySelector('#execute-status');
    const summaryCard = container.querySelector('#admin-summary-card');
    const summaryDetails = container.querySelector('#summary-details');

    try {
      btnExecute.disabled = true;
      statusEl.textContent = 'Processing records... Please wait.';

      const result = await Api.executeBulkImport({
        tempId: this.uploadData.tempId,
        mappings: mappings,
        duplicateStrategy: duplicateStrategy
      });

      statusEl.textContent = '';
      summaryCard.style.display = 'block';
      summaryDetails.innerHTML = `
        <div><strong>Total Created:</strong> ${result.importedCount} new applications</div>
        <div><strong>Total Updated:</strong> ${result.updatedCount} existing records</div>
        <div><strong>Skipped:</strong> ${result.skippedCount}</div>
        <div><strong>Errors:</strong> ${result.errorsCount}</div>
      `;

      await this.loadHistory(container);
    } catch (err) {
      alert('Import execution failed: ' + err.message);
      statusEl.textContent = 'Failed: ' + err.message;
    } finally {
      btnExecute.disabled = false;
    }
  },

  async loadHistory(container) {
    const historyContainer = container.querySelector('#history-container');
    if (!historyContainer) return;

    try {
      const list = await Api.getImportHistory();
      if (!list || list.length === 0) {
        historyContainer.innerHTML = `
          <div style="text-align: center; padding: 20px; color: var(--muted); font-size: 0.86rem;">
            No previous imports recorded.
          </div>
        `;
        return;
      }

      historyContainer.innerHTML = `
        <table class="apps-table">
          <thead>
            <tr>
              <th>Filename</th>
              <th>Date &amp; Time</th>
              <th>Imported</th>
              <th>Updated</th>
              <th>Skipped</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            ${list.map(item => `
              <tr>
                <td><strong>${item.filename}</strong></td>
                <td style="font-size: 0.82rem; color: var(--muted);">${(item.created_at || item.imported_at) ? new Date(item.created_at || item.imported_at).toLocaleString() : '—'}</td>
                <td><span style="color: #16a34a; font-weight: 600;">+${item.imported_count}</span></td>
                <td>${item.updated_count}</td>
                <td>${item.skipped_count}</td>
                <td>
                  <span class="status-badge ${item.status === 'Completed' ? 'badge-extracted' : 'badge-neutral'}">
                    ${item.status}
                  </span>
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      `;
    } catch (err) {
      historyContainer.innerHTML = `
        <div style="color: var(--danger); font-size: 0.86rem; padding: 12px;">
          Failed to load import history: ${err.message}
        </div>
      `;
    }
  }
};
