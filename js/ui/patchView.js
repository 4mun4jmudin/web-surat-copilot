/**
 * patchView.js - Tampilan Antarmuka Mode B (Edit .docx yang Diunggah)
 * Memungkinkan modifikasi paragraf berbasis anchor tanpa merusak tata letak asli.
 */

import { patchUploadedDocx } from '../docx/patch.js';
import { downloadFile } from '../utils/download.js';
import { showToast } from './notification.js';

let uploadedFileBuffer = null;
let uploadedFileName = '';

/**
 * Merender antarmuka Mode B
 * @param {HTMLElement} container 
 */
export function renderPatchView(container) {
  if (!container) return;

  container.innerHTML = `
    <div class="patch-view-container">
      <div class="card">
        <h3>Mode B: Edit Dokumen Word (.docx) Tanpa Merusak Format</h3>
        <p class="text-muted">
          Unggah dokumen yang sudah rapi (skripsi/laporan/surat). Sistem akan mencari teks sasaran (anchor) 
          dan menyisipkan atau merevisi konten dengan menyalin gaya paragraf tetangga secara presisi.
        </p>

        <div class="upload-dropzone" id="patch-dropzone">
          <input type="file" id="patch-file-input" accept=".docx" class="hidden">
          <div class="dropzone-content">
            <span class="dropzone-icon">📁</span>
            <p><strong>Klik untuk memilih berkas .docx</strong> atau seret ke sini</p>
            <span class="file-name-label" id="file-name-label">Belum ada berkas dipilih</span>
          </div>
        </div>

        <div id="patch-operations-panel" class="hidden">
          <hr/>
          <h4>Konfigurasi Operasi Perubahan</h4>
          
          <div class="form-group">
            <label>Teks Sasaran (Anchor Teks):</label>
            <input type="text" id="patch-anchor-text" class="form-control" placeholder="Ketik kata/kalimat yang menjadi patokan pencarian di Word...">
            <small class="text-muted">Sistem akan menggabungkan seluruh teks run dalam satu paragraf untuk mencocokkan.</small>
          </div>

          <div class="form-row">
            <div class="form-group col">
              <label>Tindakan:</label>
              <select id="patch-action" class="form-control">
                <option value="insertAfter">Sisipkan Paragraf Sesudahnya (insertAfter)</option>
                <option value="insertBefore">Sisipkan Paragraf Sebelumnya (insertBefore)</option>
                <option value="replace">Ganti Konten Paragraf (replace)</option>
                <option value="delete">Hapus Paragraf Sasaran (delete)</option>
              </select>
            </div>
            <div class="form-group col-checkbox">
              <label class="checkbox-label">
                <input type="checkbox" id="patch-track-changes" checked>
                <span>Tandai sebagai <strong>Track Changes</strong> (Revisi Word &lt;w:ins&gt;/&lt;w:del&gt;)</span>
              </label>
            </div>
          </div>

          <div class="form-group" id="group-new-text">
            <label>Teks Baru yang Disisipkan / Pengganti:</label>
            <textarea id="patch-new-text" class="form-control" rows="4" placeholder="Tuliskan teks paragraf baru..."></textarea>
          </div>

          <div class="btn-group">
            <button id="btn-apply-patch" class="btn btn-primary">⚡ Terapkan Perubahan & Unduh Hasil</button>
          </div>
        </div>

        <div id="patch-report-panel" class="hidden">
          <hr/>
          <h4>Laporan Perubahan & Audit Keamanan</h4>
          <div id="patch-report-content"></div>
        </div>
      </div>
    </div>
  `;

  const dropzone = container.querySelector('#patch-dropzone');
  const fileInput = container.querySelector('#patch-file-input');
  const nameLabel = container.querySelector('#file-name-label');
  const opsPanel = container.querySelector('#patch-operations-panel');
  const reportPanel = container.querySelector('#patch-report-panel');
  const reportContent = container.querySelector('#patch-report-content');
  const btnApply = container.querySelector('#btn-apply-patch');

  dropzone.onclick = () => fileInput.click();
  fileInput.onchange = (e) => {
    const file = e.target.files[0];
    if (file) handleFileSelected(file);
  };

  dropzone.ondragover = (e) => {
    e.preventDefault();
    dropzone.classList.add('drag-over');
  };
  dropzone.ondragleave = () => dropzone.classList.remove('drag-over');
  dropzone.ondrop = (e) => {
    e.preventDefault();
    dropzone.classList.remove('drag-over');
    if (e.dataTransfer.files.length > 0) {
      handleFileSelected(e.dataTransfer.files[0]);
    }
  };

  function handleFileSelected(file) {
    if (!file.name.endsWith('.docx')) {
      showToast('Harap pilih berkas dengan format .docx.', 'error');
      return;
    }
    uploadedFileName = file.name;
    nameLabel.textContent = `${file.name} (${Math.round(file.size / 1024)} KB)`;
    
    const reader = new FileReader();
    reader.onload = (ev) => {
      uploadedFileBuffer = ev.target.result;
      opsPanel.classList.remove('hidden');
      showToast(`Berkas "${file.name}" siap diproses.`, 'success');
    };
    reader.readAsArrayBuffer(file);
  }

  btnApply.onclick = async () => {
    if (!uploadedFileBuffer) {
      showToast('Silakan pilih berkas .docx terlebih dahulu.', 'warning');
      return;
    }

    const anchorText = container.querySelector('#patch-anchor-text').value.trim();
    const action = container.querySelector('#patch-action').value;
    const newText = container.querySelector('#patch-new-text').value.trim();
    const trackChanges = container.querySelector('#patch-track-changes').checked;

    if (!anchorText) {
      showToast('Teks sasaran (anchor) wajib diisi.', 'warning');
      return;
    }
    if (action !== 'delete' && !newText) {
      showToast('Teks baru tidak boleh kosong.', 'warning');
      return;
    }

    btnApply.disabled = true;
    btnApply.textContent = 'Memproses modifikasi dokumen...';

    try {
      const operations = [
        {
          action,
          anchor: { type: 'text', value: anchorText },
          text: newText
        }
      ];

      const { patchedData, report } = await patchUploadedDocx(uploadedFileBuffer, operations, {
        trackChanges,
        author: 'AI Document Builder',
        outputType: 'blob'
      });

      // Tampilkan laporan perubahan
      reportPanel.classList.remove('hidden');
      reportContent.innerHTML = `
        <div class="alert ${report.safetyPassed ? 'alert-success' : 'alert-warning'}">
          <strong>${report.safetyPassed ? '✓ Audit Keamanan Lolos' : '⚠ Perhatian Keamanan'}</strong>:
          ${report.safetyPassed ? 'Struktur section break dan header/footer dokumen utuh tanpa pergeseran.' : report.warnings.join('<br>')}
        </div>
        <p><strong>Total Operasi:</strong> ${report.totalOperations} | <strong>Eksekusi Berhasil:</strong> ${report.executedChanges.length}</p>
        <div class="table-responsive">
          <table class="table">
            <thead>
              <tr><th>Aksi</th><th>Anchor</th><th>Detail Perubahan</th><th>Indeks</th></tr>
            </thead>
            <tbody>
              ${report.executedChanges.map(c => `
                <tr>
                  <td><span class="badge badge-info">${c.action}</span></td>
                  <td><em>"${escapeHtml(c.anchorText || '')}"</em></td>
                  <td>${escapeHtml(c.insertedText || c.newText || c.deletedText || '')}</td>
                  <td>#${c.index + 1}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      `;

      // Otomatis unduh
      const outName = uploadedFileName.replace(/\.docx$/i, '_patched.docx');
      downloadFile(patchedData, outName);
      showToast(`Perubahan berhasil diterapkan. Mengunduh ${outName}...`, 'success');

    } catch (err) {
      showToast(`Gagal mengedit berkas: ${err.message}`, 'error');
    } finally {
      btnApply.disabled = false;
      btnApply.textContent = '⚡ Terapkan Perubahan & Unduh Hasil';
    }
  };
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

