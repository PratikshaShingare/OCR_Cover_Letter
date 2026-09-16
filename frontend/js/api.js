/**
 * Khanna Travels & Holidays — API Client
 * Centralized, reusable asynchronous service for backend endpoints.
 * Automatically resolves local vs production backend endpoint.
 */

export function resolveApiBaseUrl() {
  if (typeof window !== 'undefined') {
    // 1. URL parameter override: ?api=http://...
    const urlParams = new URLSearchParams(window.location.search);
    const queryApi = urlParams.get('api');
    if (queryApi) {
      return queryApi.replace(/\/+$/, '');
    }

    // 2. Global window override if injected
    if (window.__API_BASE_URL__) {
      return window.__API_BASE_URL__.replace(/\/+$/, '');
    }

    // 3. User configured localStorage override
    const stored = localStorage.getItem('API_BASE_URL');
    if (stored) {
      return stored.replace(/\/+$/, '');
    }
  }
  // Local development default
  return '/api';
}


const API_BASE = resolveApiBaseUrl();

export const Api = {
  // Application CRUD
  async getApplications() {
    const res = await fetch(`${API_BASE}/applications`);
    if (!res.ok) throw new Error('Failed to load applications');
    return res.json();
  },

  async getApplication(id) {
    const res = await fetch(`${API_BASE}/applications/${id}`);
    if (!res.ok) throw new Error(`Failed to load application ${id}`);
    return res.json();
  },

  async createNewApplication() {
    const res = await fetch(`${API_BASE}/applications/new`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' }
    });
    if (!res.ok) throw new Error('Failed to create new application');
    return res.json();
  },

  async saveApplication(appData) {
    const res = await fetch(`${API_BASE}/applications`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(appData)
    });
    if (!res.ok) throw new Error('Failed to save application');
    return res.json();
  },

  async deleteApplication(id) {
    const res = await fetch(`${API_BASE}/applications/${id}`, {
      method: 'DELETE'
    });
    if (!res.ok) throw new Error(`Failed to delete application ${id}`);
    return res.json();
  },

  // Passport Upload & Document Serving
  async uploadPassport(appId, file) {
    const formData = new FormData();
    formData.append('file', file);
    const res = await fetch(`${API_BASE}/applications/${appId}/passport`, {
      method: 'POST',
      body: formData
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({ detail: 'Upload failed' }));
      throw new Error(err.detail || 'Passport upload failed');
    }
    return res.json();
  },

  async uploadPassportFast(appId, file) {
    const formData = new FormData();
    formData.append('file', file);
    const res = await fetch(`${API_BASE}/applications/${appId}/passport?extract=false`, {
      method: 'POST',
      body: formData
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({ detail: 'Upload failed' }));
      throw new Error(err.detail || 'Passport upload failed');
    }
    return res.json();
  },

  async extractPassportOcr(appId) {
    const res = await fetch(`${API_BASE}/applications/${appId}/extract-ocr`, {
      method: 'POST'
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({ detail: 'OCR Extraction failed' }));
      throw new Error(err.detail || 'OCR Extraction failed');
    }
    return res.json();
  },


  getPassportFileUrl(appId) {
    return `${API_BASE}/applications/${appId}/passport-file?t=${Date.now()}`;
  },

  getPassportPreviewUrl(appId, page = 1) {
    return `${API_BASE}/applications/${appId}/passport-preview?page=${page}&t=${Date.now()}`;
  },

  async getPassportInfo(appId) {
    const res = await fetch(`${API_BASE}/applications/${appId}/passport-info`);
    if (!res.ok) return { hasPassport: false, totalPages: 0, isPdf: false };
    return res.json();
  },

  async removePassportFile(appId) {
    const res = await fetch(`${API_BASE}/applications/${appId}/passport-file`, {
      method: 'DELETE'
    });
    if (!res.ok) throw new Error('Failed to remove passport file');
    return res.json();
  },

  // Templates
  async getTemplates() {
    const res = await fetch(`${API_BASE}/templates`);
    if (!res.ok) throw new Error('Failed to load templates');
    return res.json();
  },

  // Documents
  async generateCoverLetter(appData) {
    const res = await fetch(`${API_BASE}/documents/generate-cover-letter`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(appData)
    });
    if (!res.ok) throw new Error('Failed to generate cover letter');
    return res.json();
  },

  async downloadDocx(appData) {
    const res = await fetch(`${API_BASE}/documents/download-docx`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(appData)
    });
    if (!res.ok) throw new Error('Failed to download DOCX');
    const blob = await res.blob();
    const filename = `Cover_Letter_${appData.selectedCountry || 'Visa'}_${(appData.personal?.fullName || 'Applicant').replace(/\s+/g, '_')}.docx`;
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
  },

  async downloadPdf(appData) {
    const res = await fetch(`${API_BASE}/documents/download-pdf`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(appData)
    });
    if (!res.ok) throw new Error('Failed to download PDF');
    const blob = await res.blob();
    const filename = `Cover_Letter_${appData.selectedCountry || 'Visa'}_${(appData.personal?.fullName || 'Applicant').replace(/\s+/g, '_')}.pdf`;
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
  },

  async downloadMasterExcel() {
    const res = await fetch(`${API_BASE}/excel/master`);
    if (!res.ok) throw new Error('Failed to download Master Excel workbook');
    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Khanna_Travels_Client_Master.xlsx';
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
  },

  async downloadMultiSheetExport(appId, applicantName = 'Client') {
    const res = await fetch(`${API_BASE}/excel/export-multi-sheet/${appId}`);
    if (!res.ok) throw new Error('Failed to download client multi-sheet export');
    const blob = await res.blob();
    const cleanName = (applicantName || 'Client').replace(/\s+/g, '_');
    const filename = `Khanna_Travels_${cleanName}_Export.xlsx`;
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
  },

  async downloadAllClientsExport() {
    const res = await fetch(`${API_BASE}/excel/export-all`);
    if (!res.ok) throw new Error('Failed to download executive clients export');
    const blob = await res.blob();
    const filename = `Khanna_Travels_All_Clients_Export.xlsx`;
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);
  },

  // Admin Bulk Import & Audit
  async uploadBulkImportFile(file) {
    const formData = new FormData();
    formData.append('file', file);
    const res = await fetch(`${API_BASE}/admin/import/upload`, {
      method: 'POST',
      body: formData
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({ detail: 'Upload failed' }));
      throw new Error(err.detail || 'Import file upload failed');
    }
    return res.json();
  },

  async executeBulkImport(payload) {
    const res = await fetch(`${API_BASE}/admin/import/execute`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({ detail: 'Execution failed' }));
      throw new Error(err.detail || 'Bulk import execution failed');
    }
    return res.json();
  },

  async getImportHistory() {
    const res = await fetch(`${API_BASE}/admin/import/history`);
    if (!res.ok) throw new Error('Failed to fetch import history');
    return res.json();
  }
};

