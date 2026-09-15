/**
 * Khanna Travels & Holidays — Application Controller & Router
 * Orchestrates navigation, reactive view transitions, and top-bar actions.
 */

import { Api } from './api.js';
import { State } from './state.js';
import { DashboardModule } from './dashboard.js';
import { PassportModule } from './passport.js';
import { ApplicantModule } from './applicant.js';
import { TravellersModule } from './travellers.js';
import { TravelModule } from './travel.js';
import { TemplatesModule } from './templates.js';
import { CoverLetterModule } from './cover-letter.js';
import { ReviewModule } from './review.js';

const STEPS = [
  { step: 1, label: 'Passport' },
  { step: 2, label: 'Applicant' },
  { step: 3, label: 'Travellers' },
  { step: 4, label: 'Travel' },
  { step: 5, label: 'Template' },
  { step: 6, label: 'Cover Letter' },
  { step: 7, label: 'Output' }
];

export const App = {
  init() {
    if (window.__KHANNA_APP_LOADED__) return;
    window.__KHANNA_APP_LOADED__ = true;
    this.bindHeader();
    this.renderStepper();

    // Subscribe to state changes
    State.subscribe((event) => {
      if (event === 'viewMode' || event === 'step' || event === 'application') {
        this.renderView();
      }
    });

    // Initial render
    this.renderView();
  },

  bindHeader() {
    const brand = document.getElementById('header-brand');
    const btnNewApp = document.getElementById('header-btn-new');
    const btnDashboard = document.getElementById('header-btn-dashboard');

    if (brand) {
      brand.addEventListener('click', () => State.setViewMode('dashboard'));
    }

    if (btnDashboard) {
      btnDashboard.addEventListener('click', () => State.setViewMode('dashboard'));
    }

    if (btnNewApp) {
      btnNewApp.addEventListener('click', async () => {
        try {
          const newApp = await Api.createNewApplication();
          State.setApplication(newApp);
          State.setPassportFileStatus(false, null);
          State.setViewMode('wizard');
          State.setStep(1);
        } catch (err) {
          alert(`Failed to create application: ${err.message}`);
        }
      });
    }
  },

  renderStepper() {
    const stepperContainer = document.getElementById('stepper-container');
    if (!stepperContainer) return;

    stepperContainer.innerHTML = STEPS.map(s => `
      <button type="button" class="stepper-step" data-step="${s.step}">
        <div class="step-circle">${s.step}</div>
        <div class="step-label">${s.label}</div>
      </button>
    `).join('');

    stepperContainer.querySelectorAll('.stepper-step').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const step = parseInt(e.currentTarget.dataset.step, 10);
        // Can only jump if an application is active
        if (State.currentApplication) {
          State.setStep(step);
        }
      });
    });
  },

  updateStepperUI() {
    const stepperBar = document.getElementById('stepper-bar');
    if (!stepperBar) return;

    if (State.viewMode === 'dashboard') {
      stepperBar.classList.add('hidden');
      return;
    }

    stepperBar.classList.remove('hidden');

    const current = State.currentStep;
    document.querySelectorAll('.stepper-step').forEach(btn => {
      const step = parseInt(btn.dataset.step, 10);
      btn.classList.remove('active', 'completed');
      if (step === current) {
        btn.classList.add('active');
      } else if (step < current) {
        btn.classList.add('completed');
      }
    });
  },

  renderView() {
    this.updateStepperUI();
    const appRoot = document.getElementById('app-root');
    if (!appRoot) return;

    if (State.viewMode === 'dashboard') {
      DashboardModule.render(appRoot);
      return;
    }

    // Wizard Step Routing
    switch (State.currentStep) {
      case 1:
        PassportModule.render(appRoot);
        break;
      case 2:
        ApplicantModule.render(appRoot);
        break;
      case 3:
        TravellersModule.render(appRoot);
        break;
      case 4:
        TravelModule.render(appRoot);
        break;
      case 5:
        TemplatesModule.render(appRoot);
        break;
      case 6:
        CoverLetterModule.render(appRoot);
        break;
      case 7:
        ReviewModule.render(appRoot);
        break;
      default:
        PassportModule.render(appRoot);
        break;
    }
  }
};

// Bootstrap application immediately if DOM ready, or on DOMContentLoaded
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => App.init());
} else {
  App.init();
}

