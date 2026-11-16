/**
 * main.js - Bootstrap Aplikasi Word Processor & AI Copilot Persuratan Sekolah
 */

import { store } from './state.js';
import {
  renderWordCanvas,
  triggerWordCanvasPrintWithChoice,
  triggerWordCanvasDownloadWithChoice
} from './ui/wordCanvas.js';
import {
  renderSmartWordCanvas,
  getActiveSmartDocModel,
  triggerSmartCanvasPrintWithChoice,
  triggerSmartCanvasDownloadWithChoice
} from './ui/smartWordCanvas.js';
import { renderOutgoingLettersView } from './ui/outgoingLettersView.js';
import { renderVaultView } from './ui/vaultView.js';
import { renderSchoolProfileView } from './ui/schoolProfileView.js';
import { renderSettingsView } from './ui/settingsView.js';
import { renderQrVerificationView } from './ui/verifyView.js';
import { renderDocx } from './docx/render.js';
import { downloadFile } from './utils/download.js';
import { showToast } from './ui/notification.js';
import { runSafeWindowPrint } from './ui/printChoiceModal.js';
import { initPwa } from './pwa.js';

document.addEventListener('DOMContentLoaded', () => {
  initApp();
});

function initApp() {
  const mainContainer = document.getElementById('main-content-view');
  const params = new URLSearchParams(window.location.search);
  const verifyParam = params.get('v');

  // Rute Publik QR Verifikasi Surat (?v=<verificationId>) — langsung tampilkan portal verifikasi tanpa editor TU
  if (verifyParam !== null) {
    renderQrVerificationView(mainContainer, verifyParam);
    return;
  }

  let currentTab = 'canvas';

  // 1. Tab Switching
  const navTabs = document.querySelectorAll('.nav-tab');
  navTabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const targetTab = tab.dataset.tab;
      currentTab = targetTab;

      navTabs.forEach(t => {
        t.classList.remove('active', 'btn-primary', 'btn-secondary');
      });
      tab.classList.add('active');
      tab.classList.remove('btn-secondary', 'btn-primary');

      renderTab(targetTab, mainContainer);
    });
  });

  function renderTab(tabName, container) {
    container.innerHTML = '';
    switch (tabName) {
      case 'canvas':
        renderWordCanvas(container);
        break;
      case 'smart-canvas':
        renderSmartWordCanvas(container);
        break;
      case 'outgoing':
        renderOutgoingLettersView(container);
        break;
      case 'vault':
        renderVaultView(container);
        break;
      case 'school':
        renderSchoolProfileView(container);
        break;
      case 'settings':
        renderSettingsView(container);
        break;
      default:
        renderWordCanvas(container);
    }
  }

  // 2. Header Quick Download (Mendukung Pilihan QR vs TTD Manual pada Kanvas & Word Pro)
  document.getElementById('header-btn-download').addEventListener('click', async () => {
    if (currentTab === 'smart-canvas') {
      triggerSmartCanvasDownloadWithChoice();
      return;
    }
    if (currentTab === 'canvas' || mainContainer.querySelector('#editable-sig-date')) {
      triggerWordCanvasDownloadWithChoice(mainContainer);
      return;
    }

    const docModel = store.getState().document || getActiveSmartDocModel();
    showToast('Menyusun berkas Word .docx resmi...', 'info');
    try {
      const blob = await renderDocx(docModel, 'blob');
      const safeTitle = (docModel.meta?.title || 'surat_dinas')
        .replace(/[^a-zA-Z0-9_\-]/g, '_')
        .slice(0, 45);
      downloadFile(blob, `${safeTitle}.docx`);
      showToast('Berkas .docx berhasil diunduh.', 'success');
    } catch (err) {
      showToast(`Gagal mengunduh: ${err.message}`, 'error');
    }
  });

  // 3. Header Quick Print (Pilihan Barcode QR Pengesahan vs Kosongan TTD Manual Pulpen)
  document.getElementById('header-btn-print').addEventListener('click', () => {
    if (currentTab === 'smart-canvas') {
      triggerSmartCanvasPrintWithChoice();
      return;
    }
    if (currentTab === 'canvas' || mainContainer.querySelector('#editable-sig-date')) {
      triggerWordCanvasPrintWithChoice(mainContainer);
      return;
    }
    runSafeWindowPrint();
  });

  // Inisialisasi PWA (Service Worker & Install App Prompt)
  initPwa();

  // Dukung shortcut PWA (?tab=outgoing|vault|school|settings|canvas)
  const requestedTab = new URLSearchParams(window.location.search).get('tab');
  const matchingTabBtn = requestedTab
    ? document.querySelector(`.nav-tab[data-tab="${CSS.escape(requestedTab)}"]`)
    : null;

  if (matchingTabBtn) {
    matchingTabBtn.click();
  } else {
    // Render awal Lembar Kerja Word
    renderTab('canvas', mainContainer);
  }
}
