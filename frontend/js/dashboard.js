/**
 * Khanna Travels & Holidays — Dashboard View Module
 * Renders the main portal with '+ New Application' and real database records.
 * Supports delete with confirmation. Zero fake applications.
 */

import { Api } from './api.js';
import { State } from './state.js';

export const DashboardModule = {
  async render(container) {
    container.innerHTML = `
      <div style="max-width: 820px; margin: 30px auto 0;">
        <!-- Welcome Card -->
        <div class="card" style="padding: 44px 36px; text-align: center; margin-bottom: 32px;">
          <div style="width: 54px; height: 54px; border-radius: 12px; background-color: var(--primary); color: #ffffff; display: flex; align-items: center; justify-content: center; margin: 0 auto 16px; box-shadow: 0 4px 10px rgba(30, 58, 138, 0.25);">
            <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"></path></svg>
          </div>

          <h1 style="font-size: 1.6rem; font-weight: 800; color: var(--dark); margin-bottom: 6px;">
            KHANNA TRAVELS & HOLIDAYS
          </h1>
          <p style="font-size: 0.92rem; color: var(--muted); margin-bottom: 28px;">
            Visa Document Automation System • Internal Executive Portal
          </p>

          <button type="button" class="btn btn-primary btn-lg" id="btn-create-app" style="margin: 0 auto;">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="16"></line><line x1="8" y1="12" x2="16" y2="12"></line></svg>
            + New Application
          </button>
        </div>

        <!-- Real Applications Section -->
        <div class="card" style="padding: 24px;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; flex-wrap: wrap; gap: 10px;">
            <h3 style="font-size: 1.05rem; font-weight: 700; color: var(--dark); margin: 0;">
              Applications
            </h3>
            <div style="display: flex; gap: 8px;">
              <button type="button" class="btn btn-secondary btn-sm" id="btn-dash-export-all" title="Download 4-sheet Excel for all clients">
                Export All Clients (.xlsx)
              </button>
              <button type="button" class="btn btn-secondary btn-sm" id="btn-dash-admin" title="Admin Bulk Ingestion & Audit">
                Admin &amp; Import
              </button>
              <button type="button" class="btn btn-secondary btn-sm" id="btn-refresh-apps" title="Refresh List">
                Refresh
              </button>
            </div>
          </div>

          <div id="apps-container">
            <div style="text-align: center; padding: 32px; color: var(--muted); font-size: 0.88rem;">
              Loading applications from database...
            </div>
          </div>
        </div>
      </div>

      <!-- Delete Confirmation Modal -->
      <div id="delete-modal" class="modal-overlay hidden">
        <div class="modal-dialog">
          <div class="modal-title">Delete Application?</div>
          <div class="modal-body" id="delete-modal-text">
            Delete this application and its associated client data and generated files?
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary" id="btn-cancel-delete">Cancel</button>
            <button type="button" class="btn btn-danger" id="btn-confirm-delete">Yes, Delete Application</button>
          </div>
        </div>
      </div>
    `;

    this.bindEvents(container);
    await this.loadApplications(container);
  },

  async loadApplications(container) {
    const appsContainer = container.querySelector('#apps-container');
    if (!appsContainer) return;

    try {
      const apps = await Api.getApplications();

      if (!apps || apps.length === 0) {
        appsContainer.innerHTML = `
          <div style="text-align: center; padding: 40px 20px; color: var(--muted); background-color: #fafafa; border-radius: 8px;">
            <div style="font-size: 1rem; font-weight: 600; color: #475569; margin-bottom: 4px;">
              No applications yet.
            </div>
            <div style="font-size: 0.86rem;">
              Click <strong>"+ New Application"</strong> above to create one.
            </div>
          </div>
        `;
        return;
      }

      appsContainer.innerHTML = `
        <table class="apps-table">
          <thead>
            <tr>
              <th>Application ID</th>
              <th>Applicant Name</th>
              <th>Destination</th>
              <th>Status</th>
              <th style="text-align: right;">Action</th>
            </tr>
          </thead>
          <tbody>
            ${apps.map(a => `
              <tr>
                <td><strong>${a.id}</strong></td>
                <td>${a.applicant_name || '—'}</td>
                <td>${a.destination_country || 'Visa'}</td>
                <td>
                  <span class="status-badge ${a.status === 'Verified' ? 'badge-extracted' : 'badge-neutral'}">
                    ${a.status || 'Draft'}
                  </span>
                </td>
                <td style="text-align: right;">
                  <button type="button" class="btn btn-secondary btn-sm btn-open-app" data-id="${a.id}" style="margin-right: 6px;">
                    Open
                  </button>
                  <button type="button" class="btn btn-secondary btn-sm btn-export-client" data-id="${a.id}" data-name="${a.applicant_name || 'Client'}" title="Download 4-sheet client Excel" style="margin-right: 6px;">
                    Export (.xlsx)
                  </button>
                  <button type="button" class="btn btn-secondary btn-sm btn-delete-app" data-id="${a.id}" style="color: var(--danger); border-color: #fca5a5;">
                    Delete
                  </button>
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      `;

      // Bind Open & Delete buttons
      appsContainer.querySelectorAll('.btn-open-app').forEach(btn => {
        btn.addEventListener('click', async (e) => {
          const id = e.currentTarget.dataset.id;
          btn.disabled = true;
          try {
            const appData = await Api.getApplication(id);
            State.setApplication(appData);
            // Check if passport file exists and total pages
            try {
              const pInfo = await Api.getPassportInfo(id);
              State.setPassportFileStatus(pInfo.hasPassport, pInfo.isPdf ? 'pdf' : 'image', pInfo.totalPages || 1);
            } catch {
              State.setPassportFileStatus(false, null, 1);
            }
            State.setViewMode('wizard');
            State.setStep(1);
          } catch (err) {
            alert(`Failed to open application: ${err.message}`);
            btn.disabled = false;
          }
        });
      });

      // Bind per-client multi-sheet Excel export
      appsContainer.querySelectorAll('.btn-export-client').forEach(btn => {
        btn.addEventListener('click', async (e) => {
          const id = e.currentTarget.dataset.id;
          const name = e.currentTarget.dataset.name;
          btn.disabled = true;
          btn.textContent = '...';
          try {
            await Api.downloadMultiSheetExport(id, name);
          } catch (err) {
            alert(`Export failed: ${err.message}`);
          } finally {
            btn.disabled = false;
            btn.textContent = 'Export (.xlsx)';
          }
        });
      });

      let targetDeleteId = null;
      const deleteModal = container.querySelector('#delete-modal');
      const deleteModalText = container.querySelector('#delete-modal-text');

      appsContainer.querySelectorAll('.btn-delete-app').forEach(btn => {
        btn.addEventListener('click', (e) => {
          targetDeleteId = e.currentTarget.dataset.id;
          deleteModalText.innerHTML = `Delete application <strong>${targetDeleteId}</strong> and its associated client data and generated files?`;
          deleteModal.classList.remove('hidden');
        });
      });

      const btnCancelDelete = container.querySelector('#btn-cancel-delete');
      if (btnCancelDelete) {
        btnCancelDelete.onclick = () => {
          deleteModal.classList.add('hidden');
          targetDeleteId = null;
        };
      }

      const btnConfirmDelete = container.querySelector('#btn-confirm-delete');
      if (btnConfirmDelete) {
        btnConfirmDelete.onclick = async () => {
          if (!targetDeleteId) return;
          try {
            await Api.deleteApplication(targetDeleteId);
            deleteModal.classList.add('hidden');
            targetDeleteId = null;
            await this.loadApplications(container);
          } catch (err) {
            alert(`Failed to delete: ${err.message}`);
          }
        };
      }

    } catch (err) {
      appsContainer.innerHTML = `
        <div style="text-align: center; padding: 20px; color: var(--danger);">
          Failed to load applications: ${err.message}
        </div>
      `;
    }
  },

  bindEvents(container) {
    const btnCreate = container.querySelector('#btn-create-app');
    if (btnCreate) {
      btnCreate.addEventListener('click', async () => {
        btnCreate.disabled = true;
        try {
          // Creates 100% blank application with next sequential ID (e.g. APP-2026-00001)
          const newApp = await Api.createNewApplication();
          State.setApplication(newApp);
          State.setPassportFileStatus(false, null);
          State.setViewMode('wizard');
          State.setStep(1);
        } catch (err) {
          alert(`Failed to create application: ${err.message}`);
          btnCreate.disabled = false;
        }
      });
    }

    const btnRefresh = container.querySelector('#btn-refresh-apps');
    if (btnRefresh) {
      btnRefresh.addEventListener('click', () => this.loadApplications(container));
    }

    const btnAdmin = container.querySelector('#btn-dash-admin');
    if (btnAdmin) {
      btnAdmin.addEventListener('click', () => State.setViewMode('admin'));
    }

    const btnExportAll = container.querySelector('#btn-dash-export-all');
    if (btnExportAll) {
      btnExportAll.addEventListener('click', async () => {
        try {
          btnExportAll.disabled = true;
          btnExportAll.textContent = 'Generating...';
          await Api.downloadAllClientsExport();
        } catch (err) {
          alert(`Export failed: ${err.message}`);
        } finally {
          btnExportAll.disabled = false;
          btnExportAll.textContent = 'Export All Clients (.xlsx)';
        }
      });
    }
  }
};
