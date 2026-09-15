/**
 * Khanna Travels & Holidays — API Client
 * Centralized, reusable asynchronous service for backend endpoints.
 */

const API_BASE = '/api';

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
  }
};
