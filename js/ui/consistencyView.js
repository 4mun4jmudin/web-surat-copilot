/**
 * consistencyView.js - Pemeriksa Kualitas dan Konsistensi Dokumen
 */

import { store } from '../state.js';
import { checkDocumentConsistency, autoFixDocument } from '../checks/consistency.js';
import { showToast } from './notification.js';

export function renderConsistencyView(container) {
  if (!container) return;

  const currentDoc = store.getState().document;
  const { issues, summary } = checkDocumentConsistency(currentDoc);

  container.innerHTML = `
    <div class="consistency-view-container">
      <div class="card">
        <div class="card-header-flex">
          <div>
            <h3>Pemeriksa Kualitas & Konsistensi Dokumen</h3>
            <p class="text-muted">
              Mendeteksi nomor gambar/tabel loncat/dobel, rujukan bookmark hilang, heading melompat level, 
              paragraf kosong berlebih, dan standar tipografi akademik.
            </p>
          </div>
          <div>
            <button id="btn-run-autofix" class="btn btn-success" ${issues.length === 0 ? 'disabled' : ''}>
              🔧 Perbaiki Otomatis (Auto-Fix)
            </button>
          </div>
        </div>

        <div class="stat-summary-row">
          <div class="stat-box stat-error">
            <span class="stat-num">${summary.errors}</span>
            <span class="stat-label">Masalah Kritis</span>
          </div>
          <div class="stat-box stat-warning">
            <span class="stat-num">${summary.warnings}</span>
            <span class="stat-label">Peringatan</span>
          </div>
          <div class="stat-box stat-info">
            <span class="stat-num">${summary.infos}</span>
            <span class="stat-label">Saran Format</span>
          </div>
        </div>

        <div class="issues-list" id="issues-list-container">
          ${issues.length === 0 ? `
            <div class="alert alert-success">
              ✓ Dokumen sangat rapi! Tidak ditemukan inkonsistensi struktur, penomoran, maupun rujukan.
            </div>
          ` : issues.map((iss, idx) => `
            <div class="issue-card issue-${iss.severity}">
              <div class="issue-header">
                <span class="badge badge-${iss.severity}">${iss.severity.toUpperCase()}</span>
                <span class="issue-code">${iss.code}</span>
                ${iss.blockIndex !== undefined ? `<span class="issue-loc">Blok #${iss.blockIndex + 1}</span>` : ''}
              </div>
              <div class="issue-body">
                <p class="issue-msg">${escapeHtml(iss.message)}</p>
                ${iss.suggestion ? `<div class="issue-suggestion">💡 <em>${escapeHtml(iss.suggestion)}</em></div>` : ''}
              </div>
            </div>
          `).join('')}
        </div>

      </div>
    </div>
  `;

  const btnAutoFix = container.querySelector('#btn-run-autofix');
  if (btnAutoFix) {
    btnAutoFix.onclick = () => {
      const fixedDoc = autoFixDocument(store.getState().document);
      store.setState({ document: fixedDoc }, true);
      showToast('Perbaikan otomatis telah diterapkan ke dokumen.', 'success');
      renderConsistencyView(container);
    };
  }
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

