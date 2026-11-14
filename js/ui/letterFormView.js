/**
 * letterFormView.js - Formulir Pembuatan Surat Dinas Sekolah
 * Mengisi nilai field template secara interaktif dengan sinkronisasi pratinjau realtime.
 */

import { store } from '../state.js';
import { suratKeteranganTemplate } from '../letters/templates/suratKeterangan.js';
import { showToast } from './notification.js';

export function renderLetterFormView(container) {
  if (!container) return;

  const state = store.getState();
  const formValues = state.currentLetterForm;
  const options = state.letterOptions;

  container.innerHTML = `
    <div class="letter-form-container">
      <div class="card">
        <div class="card-header-flex">
          <div>
            <h3>Pembuatan Surat Administrasi Sekolah</h3>
            <p class="text-muted">Pilih jenis template surat, sesuaikan pengaturan kertas, dan lengkapi rincian data.</p>
          </div>
        </div>

        <!-- PENGATURAN TEMPLATE & KERTAS -->
        <div class="form-row" style="margin-bottom: 20px;">
          <div class="form-group col">
            <label>Jenis Surat:</label>
            <select id="select-letter-template" class="form-control">
              <option value="surat_keterangan" selected>Surat Keterangan Siswa Aktif</option>
              <option value="surat_undangan" disabled>Surat Undangan Dinas (Tahap C)</option>
              <option value="surat_tugas" disabled>Surat Tugas / Perjalanan Dinas (Tahap C)</option>
              <option value="surat_dispensasi" disabled>Surat Dispensasi Siswa (Tahap C)</option>
            </select>
          </div>
          <div class="form-group col">
            <label>Ukuran Kertas:</label>
            <select id="select-paper-size" class="form-control">
              <option value="a4" ${options.paperSize === 'a4' ? 'selected' : ''}>A4 (21,0 x 29,7 cm)</option>
              <option value="f4" ${options.paperSize === 'f4' ? 'selected' : ''}>F4 / Folio (21,5 x 33,0 cm)</option>
            </select>
          </div>
        </div>

        <hr/>
        <h4 style="margin: 16px 0 12px 0;">Data Isian: ${suratKeteranganTemplate.name}</h4>

        <div class="form-fields-grid" id="letter-fields-container"></div>

        <div class="btn-group" style="margin-top: 24px;">
          <button id="btn-reset-form" class="btn btn-secondary">↺ Muat Ulang Isian Contoh</button>
        </div>
      </div>
    </div>
  `;

  // Render fields
  const fieldsContainer = container.querySelector('#letter-fields-container');
  suratKeteranganTemplate.fields.forEach(field => {
    const val = formValues[field.name] !== undefined ? formValues[field.name] : (field.defaultValue || '');
    const fieldEl = document.createElement('div');
    fieldEl.className = 'form-group';

    if (field.type === 'textarea') {
      fieldEl.innerHTML = `
        <label>${field.label} ${field.required ? '<span class="text-danger">*</span>' : ''}:</label>
        <textarea class="form-control letter-field-input" data-name="${field.name}" rows="2" placeholder="${field.placeholder || ''}">${escapeHtml(val)}</textarea>
      `;
    } else if (field.type === 'select') {
      fieldEl.innerHTML = `
        <label>${field.label} ${field.required ? '<span class="text-danger">*</span>' : ''}:</label>
        <select class="form-control letter-field-input" data-name="${field.name}">
          ${(field.options || []).map(opt => `
            <option value="${opt}" ${val === opt ? 'selected' : ''}>${opt}</option>
          `).join('')}
        </select>
      `;
    } else {
      fieldEl.innerHTML = `
        <label>${field.label} ${field.required ? '<span class="text-danger">*</span>' : ''}:</label>
        <input type="text" class="form-control letter-field-input" data-name="${field.name}" value="${escapeHtml(val)}" placeholder="${field.placeholder || ''}">
      `;
    }

    fieldsContainer.appendChild(fieldEl);
  });

  // Event handler input realtime
  const inputs = container.querySelectorAll('.letter-field-input');
  inputs.forEach(input => {
    const handler = () => {
      const fieldName = input.dataset.name;
      const val = input.value;
      store.updateLetterForm({ [fieldName]: val });
    };
    input.addEventListener('input', handler);
    input.addEventListener('change', handler);
  });

  // Paper size change
  const paperSelect = container.querySelector('#select-paper-size');
  paperSelect.onchange = () => {
    store.updateLetterOptions({ paperSize: paperSelect.value });
    showToast(`Ukuran kertas diubah ke ${paperSelect.value.toUpperCase()}.`, 'info');
  };

  // Reset form
  container.querySelector('#btn-reset-form').onclick = () => {
    if (confirm('Kembalikan isian formulir ke contoh baku?')) {
      const initial = {};
      suratKeteranganTemplate.fields.forEach(f => {
        initial[f.name] = f.defaultValue || '';
      });
      store.updateLetterForm(initial);
      renderLetterFormView(container);
      showToast('Formulir telah direset ke nilai contoh.', 'success');
    }
  };
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

