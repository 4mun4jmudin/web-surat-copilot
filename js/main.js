/**
 * main.js - Bootstrap Aplikasi Word Processor & AI Copilot Persuratan Sekolah
 */

import { store } from './state.js';
import { renderWordCanvas } from './ui/wordCanvas.js';
import { renderSmartWordCanvas, getActiveSmartDocModel } from './ui/smartWordCanvas.js';
import { renderOutgoingLettersView } from './ui/outgoingLettersView.js';
import { renderVaultView } from './ui/vaultView.js';
import { renderSchoolProfileView } from './ui/schoolProfileView.js';
import { renderSettingsView } from './ui/settingsView.js';
import { renderDocx } from './docx/render.js';
import { downloadFile } from './utils/download.js';
import { showToast } from './ui/notification.js';
import { initPwa } from './pwa.js';

document.addEventListener('DOMContentLoaded', () => {
  initApp();
});

function initApp() {
  const mainContainer = document.getElementById('main-content-view');
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

  // 2. Header Quick Download (Mendukung Kanvas Word Pro, Dokumen Bawaan & Berkas Impor)
  document.getElementById('header-btn-download').addEventListener('click', async () => {
    const activeDocx = store.getState().activeUploadedDocx;

    // Jika sedang membuka berkas luar (.docx) pada Lembar Kerja standar, terapkan patcher non-destruktif
    if (currentTab !== 'smart-canvas' && activeDocx && activeDocx.rawBytes) {
      showToast('Menerapkan editan ke berkas asli (preservasi format 100% utuh)...', 'info');
      try {
        const { applyCanvasEditsToDocx } = await import('./docx/patch.js');
        const patchedBlob = await applyCanvasEditsToDocx(activeDocx.rawBytes, activeDocx.paragraphs || [], { outputType: 'blob' });
        const safeName = (activeDocx.fileName || 'dokumen_teredit.docx').replace(/\.docx$/i, '') + '_edited.docx';
        downloadFile(patchedBlob, safeName);
        showToast('Berkas .docx hasil edit (format asli 100% terjaga) berhasil diunduh.', 'success');
        return;
      } catch (err) {
        showToast(`Gagal mem-patch berkas: ${err.message}`, 'error');
        return;
      }
    }

    // Dokumen standar sistem atau Kanvas Word Pro
    const docModel = currentTab === 'smart-canvas'
      ? getActiveSmartDocModel()
      : store.getState().document;
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

  // 3. Header Quick Print + Tampilkan Konfirmasi Simpan ke Surat Keluar
  document.getElementById('header-btn-print').addEventListener('click', () => {
    const postPrintBar = document.getElementById('post-print-outgoing-bar');
    if (postPrintBar) postPrintBar.classList.remove('hidden');
    window.print();
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
