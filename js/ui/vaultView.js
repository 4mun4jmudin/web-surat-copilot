/**
 * vaultView.js - Antarmuka Gudang Penyimpanan Surat (Letter Vault) & Knowledge Base AI
 * Mendukung penyimpanan, pencarian instan, duplikasi, pembukaan kembali ke Lembar Kerja Word,
 * unggah berkas .docx luar dengan preservasi format, dan pengaturan referensi AI.
 */

import { store } from '../state.js';
import { parseDocxForEditing, applyCanvasEditsToDocx } from '../docx/patch.js';
import { renderDocx } from '../docx/render.js';
import { downloadFile } from '../utils/download.js';
import { showToast } from './notification.js';

let currentSearchQuery = '';
let currentCategoryFilter = 'all';
let filterOnlyAiRef = false;

export function renderVaultView(container) {
  if (!container) return;

  const allLetters = store.getLetterVault();
  const aiRefLetters = allLetters.filter(l => l.isAiReference);

  // Filter letters
  const filtered = allLetters.filter(item => {
    if (filterOnlyAiRef && !item.isAiReference) return false;
    if (currentCategoryFilter !== 'all' && item.category !== currentCategoryFilter) return false;
    if (currentSearchQuery) {
      const q = currentSearchQuery.toLowerCase();
      const matchTitle = (item.title || '').toLowerCase().includes(q);
      const matchNomor = (item.nomorSurat || '').toLowerCase().includes(q);
      const matchPerson = (item.targetPerson || '').toLowerCase().includes(q);
      const matchPerihal = (item.perihal || '').toLowerCase().includes(q);
      const matchSummary = (item.summary || '').toLowerCase().includes(q);
      return matchTitle || matchNomor || matchPerson || matchPerihal || matchSummary;
    }
    return true;
  });

  const categories = ['all', 'Surat Keterangan', 'Surat Tugas', 'Surat Undangan', 'Dokumen Impor'];

  container.innerHTML = `
    <div class="vault-view-container" style="max-width: 1100px; margin: 20px auto; padding: 0 16px; overflow-y: auto; height: 100%;">
      
      <!-- HEADER GUDANG SURAT -->
      <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 20px; flex-wrap: wrap; gap: 14px;">
        <div>
          <h2 style="font-size: 22px; font-weight: 700; color: var(--dark); display: flex; align-items: center; gap: 10px;">
            <span>📁</span> Gudang Penyimpanan & Arsip Surat Sekolah
          </h2>
          <p class="text-muted" style="margin-top: 4px; font-size: 13.5px;">
            Arsip seluruh dokumen yang pernah dibuat atau diunggah. Surat dapat dibuka kembali untuk diedit, 
            serta otomatis menjadi <strong>basis referensi gaya penulisan untuk AI Copilot</strong>.
          </p>
        </div>

        <div style="display: flex; gap: 8px;">
          <input type="file" id="input-upload-docx-vault" accept=".docx" class="hidden">
          <button id="btn-upload-docx-vault" class="btn btn-secondary" style="font-weight: 600;">
            📤 Unggah Berkas .docx Luar
          </button>
          <button id="btn-create-new-from-vault" class="btn btn-primary" style="font-weight: 600;">
            ➕ Buat Surat Baru di Word
          </button>
        </div>
      </div>

      <!-- KARTU STATISTIK & KNOWLEDGE BASE AI -->
      <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 12px; margin-bottom: 20px;">
        <div style="background: white; border: 1px solid var(--border); border-radius: 8px; padding: 14px; box-shadow: 0 2px 4px rgba(0,0,0,0.02);">
          <span style="font-size: 11.5px; font-weight: 700; color: #64748b; text-transform: uppercase;">Total Surat Tersimpan</span>
          <div style="font-size: 24px; font-weight: 800; color: var(--dark); margin-top: 4px;">${allLetters.length} <span style="font-size: 13px; font-weight: normal; color: #64748b;">dokumen</span></div>
        </div>

        <div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; padding: 14px;">
          <span style="font-size: 11.5px; font-weight: 700; color: #166534; text-transform: uppercase;">⭐ Acuan AI Copilot Aktif</span>
          <div style="font-size: 24px; font-weight: 800; color: #15803d; margin-top: 4px;">${aiRefLetters.length} <span style="font-size: 13px; font-weight: normal; color: #166534;">surat dijadikan referensi</span></div>
        </div>

        <div style="background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 8px; padding: 14px;">
          <span style="font-size: 11.5px; font-weight: 700; color: #1e40af; text-transform: uppercase;">Preservasi Format Word</span>
          <div style="font-size: 14px; font-weight: 600; color: #1e3a8a; margin-top: 6px;">100% Aman (Non-Destructive)</div>
          <small style="color: #3b82f6; font-size: 11px;">Header, styles, margin, dan tabel asli tidak rusak saat diedit.</small>
        </div>
      </div>

      <!-- DROPZONE AREA UNTUK UNGGAH .DOCX -->
      <div id="dropzone-docx-vault" style="border: 2px dashed #cbd5e1; border-radius: 8px; padding: 16px; text-align: center; background: #f8fafc; margin-bottom: 20px; cursor: pointer; transition: all 0.2s ease;">
        <span style="font-size: 24px; display: block; margin-bottom: 4px;">📥</span>
        <strong style="color: #334155; font-size: 13.5px;">Klik untuk mengunggah atau seret berkas Word (.docx) ke sini</strong>
        <p style="font-size: 11.5px; color: #64748b; margin: 4px 0 0 0;">
          Sistem akan menyimpan dokumen ke gudang arsip dan mengekstrak isinya agar bisa langsung diedit di lembar kerja tanpa merusak format bawaan.
        </p>
      </div>

      <!-- FILTER & PENCARIAN INSTAN -->
      <div style="background: white; border: 1px solid var(--border); border-radius: 8px; padding: 12px 16px; margin-bottom: 20px; display: flex; flex-wrap: wrap; gap: 12px; align-items: center; justify-content: space-between;">
        <div style="display: flex; gap: 8px; flex-wrap: wrap; align-items: center;">
          <span style="font-size: 12px; font-weight: 700; color: #475569;">Kategori:</span>
          ${categories.map(cat => {
            const isAct = currentCategoryFilter === cat;
            const label = cat === 'all' ? 'Semua Surat' : cat;
            return `
              <button class="btn btn-sm ${isAct ? 'btn-primary' : 'btn-secondary'} btn-filter-cat" data-cat="${cat}" style="font-size: 12px;">
                ${label}
              </button>
            `;
          }).join('')}

          <button class="btn btn-sm ${filterOnlyAiRef ? 'btn-warning' : 'btn-secondary'} btn-toggle-ai-filter" style="font-size: 12px; margin-left: 6px;">
            ${filterOnlyAiRef ? '★ Filter Acuan AI Aktif' : '☆ Hanya Acuan AI'}
          </button>
        </div>

        <div style="position: relative; min-width: 260px; flex: 1; max-width: 380px;">
          <input type="text" id="input-vault-search" class="form-control" style="font-size: 12.5px; padding-left: 28px;" 
            placeholder="Cari nomor surat, nama, perihal..." value="${escapeHtml(currentSearchQuery)}">
          <span style="position: absolute; left: 8px; top: 7px; color: #94a3b8; font-size: 14px;">🔍</span>
        </div>
      </div>

      <!-- DAFTAR KARTU ARSIP SURAT -->
      ${filtered.length === 0 ? `
        <div style="text-align: center; padding: 48px 20px; background: white; border-radius: 8px; border: 1px dashed #cbd5e1;">
          <span style="font-size: 40px; display: block; margin-bottom: 12px;">📭</span>
          <h4 style="color: #334155;">Tidak Ada Surat yang Cocok</h4>
          <p style="color: #64748b; font-size: 13px; max-width: 450px; margin: 4px auto 16px auto;">
            Belum ada surat yang sesuai dengan kata kunci atau filter pencarian. Buat surat baru atau unggah berkas .docx Anda.
          </p>
          <button id="btn-empty-reset" class="btn btn-secondary btn-sm">Reset Filter Pencarian</button>
        </div>
      ` : `
        <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); gap: 16px; margin-bottom: 40px;">
          ${filtered.map(letter => {
            const isRef = letter.isAiReference;
            const badgeBg = letter.category === 'Surat Tugas' ? '#dcfce7' : letter.category === 'Surat Undangan' ? '#f3e8ff' : '#dbeafe';
            const badgeColor = letter.category === 'Surat Tugas' ? '#166534' : letter.category === 'Surat Undangan' ? '#6b21a8' : '#1e40af';

            return `
              <div class="card vault-card" style="background: white; border: 1px solid var(--border); border-radius: 8px; padding: 16px; display: flex; flex-direction: column; justify-content: space-between; box-shadow: 0 2px 6px rgba(0,0,0,0.04); transition: transform 0.15s ease, box-shadow 0.15s ease;">
                <div>
                  <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                    <span style="background: ${badgeBg}; color: ${badgeColor}; font-size: 10.5px; font-weight: 700; padding: 2px 8px; border-radius: 4px;">
                      ${escapeHtml(letter.category || 'Surat Dinas')}
                    </span>
                    <button class="btn-toggle-ai-ref" data-id="${letter.id}" title="${isRef ? 'Surat ini dijadikan acuan AI. Klik untuk nonaktifkan.' : 'Jadikan surat ini sebagai acuan gaya penulisan AI Copilot.'}" 
                      style="background: ${isRef ? '#fef3c7' : '#f1f5f9'}; border: 1px solid ${isRef ? '#f59e0b' : '#cbd5e1'}; color: ${isRef ? '#b45309' : '#64748b'}; font-size: 11px; font-weight: 700; padding: 2px 8px; border-radius: 12px; cursor: pointer;">
                      ${isRef ? '★ Acuan AI' : '☆ Jadikan Acuan'}
                    </button>
                  </div>

                  <strong style="font-size: 14.5px; color: var(--dark); display: block; line-height: 1.35; margin-bottom: 4px;">
                    ${escapeHtml(letter.title)}
                  </strong>

                  <div style="font-family: 'Courier New', monospace; font-size: 11.5px; font-weight: 700; color: #0f766e; margin-bottom: 8px;">
                    ${escapeHtml(letter.nomorSurat || '-')}
                  </div>

                  <div style="font-size: 11.5px; color: #475569; margin-bottom: 8px;">
                    📅 <strong>${escapeHtml(letter.tanggalSurat || '-')}</strong>
                    ${letter.targetPerson && letter.targetPerson !== '-' ? `• 👤 <em>${escapeHtml(letter.targetPerson)}</em>` : ''}
                  </div>

                  <p style="font-size: 12px; color: #64748b; line-height: 1.45; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 4px; padding: 8px; margin: 0 0 12px 0;">
                    ${escapeHtml(letter.summary || 'Dokumen persuratan resmi sekolah.')}
                  </p>
                </div>

                <div style="border-top: 1px solid #f1f5f9; padding-top: 10px; display: flex; gap: 6px; flex-wrap: wrap;">
                  <button class="btn btn-sm btn-primary btn-load-letter" data-id="${letter.id}" style="flex: 1; min-width: 120px;">
                    ✏ Buka di Word
                  </button>
                  <button class="btn btn-sm btn-secondary btn-download-vault-letter" data-id="${letter.id}" title="Unduh berkas .docx">
                    ⬇ .docx
                  </button>
                  <button class="btn btn-sm btn-secondary btn-duplicate-letter" data-id="${letter.id}" title="Duplikasi Surat">
                    📑
                  </button>
                  <button class="btn btn-sm btn-secondary btn-delete-vault-letter" data-id="${letter.id}" title="Hapus dari Gudang" style="color: #dc2626;">
                    🗑
                  </button>
                </div>
              </div>
            `;
          }).join('')}
        </div>
      `}

    </div>
  `;

  bindVaultEvents(container);
}

function bindVaultEvents(container) {
  // 1. Search input
  const searchInput = container.querySelector('#input-vault-search');
  if (searchInput) {
    searchInput.oninput = (e) => {
      currentSearchQuery = e.target.value.trim();
      renderVaultView(container);
    };
  }

  // 2. Filter category
  container.querySelectorAll('.btn-filter-cat').forEach(btn => {
    btn.onclick = () => {
      currentCategoryFilter = btn.dataset.cat;
      renderVaultView(container);
    };
  });

  // 3. Toggle only AI filter
  const btnToggleAiFilter = container.querySelector('.btn-toggle-ai-filter');
  if (btnToggleAiFilter) {
    btnToggleAiFilter.onclick = () => {
      filterOnlyAiRef = !filterOnlyAiRef;
      renderVaultView(container);
    };
  }

  // 4. Reset empty filter
  const btnReset = container.querySelector('#btn-empty-reset');
  if (btnReset) {
    btnReset.onclick = () => {
      currentSearchQuery = '';
      currentCategoryFilter = 'all';
      filterOnlyAiRef = false;
      renderVaultView(container);
    };
  }

  // 5. Tombol Buat Surat Baru -> Switch ke tab canvas
  const btnCreate = container.querySelector('#btn-create-new-from-vault');
  if (btnCreate) {
    btnCreate.onclick = () => {
      const canvasTab = document.querySelector('.nav-tab[data-tab="canvas"]');
      if (canvasTab) canvasTab.click();
    };
  }

  // 6. Upload .docx handling (Drag & Drop & Button)
  const fileInput = container.querySelector('#input-upload-docx-vault');
  const dropzone = container.querySelector('#dropzone-docx-vault');
  const btnUpload = container.querySelector('#btn-upload-docx-vault');

  if (btnUpload && fileInput) {
    btnUpload.onclick = () => fileInput.click();
  }
  if (dropzone && fileInput) {
    dropzone.onclick = () => fileInput.click();
    dropzone.ondragover = (e) => { e.preventDefault(); dropzone.style.borderColor = 'var(--primary)'; dropzone.style.background = '#eff6ff'; };
    dropzone.ondragleave = () => { dropzone.style.borderColor = '#cbd5e1'; dropzone.style.background = '#f8fafc'; };
    dropzone.ondrop = (e) => {
      e.preventDefault();
      dropzone.style.borderColor = '#cbd5e1';
      dropzone.style.background = '#f8fafc';
      if (e.dataTransfer.files.length > 0) {
        processUploadedDocx(e.dataTransfer.files[0], container);
      }
    };
  }

  if (fileInput) {
    fileInput.onchange = (e) => {
      const file = e.target.files[0];
      if (file) processUploadedDocx(file, container);
    };
  }

  // 7. Toggle AI Reference per card
  container.querySelectorAll('.btn-toggle-ai-ref').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      const id = btn.dataset.id;
      store.toggleLetterAiReference(id);
      showToast('Status acuan referensi AI diperbarui.', 'info');
      renderVaultView(container);
    };
  });

  // 8. Buka di Word Canvas
  container.querySelectorAll('.btn-load-letter').forEach(btn => {
    btn.onclick = () => {
      const id = btn.dataset.id;
      const success = store.loadLetterToCanvas(id);
      if (success) {
        showToast('Dokumen berhasil dimuat ke Lembar Kerja Word.', 'success');
        const canvasTab = document.querySelector('.nav-tab[data-tab="canvas"]');
        if (canvasTab) canvasTab.click();
      } else {
        showToast('Gagal memuat dokumen.', 'error');
      }
    };
  });

  // 9. Unduh dokumen langsung dari kartu
  container.querySelectorAll('.btn-download-vault-letter').forEach(btn => {
    btn.onclick = async () => {
      const id = btn.dataset.id;
      const letter = store.getLetterVault().find(l => l.id === id);
      if (!letter) return;

      showToast(`Menyiapkan berkas .docx untuk ${letter.title}...`, 'info');
      try {
        let blob;
        if (letter.uploadedDocx?.rawBytes) {
          // Berkas yang diupload: Gunakan rawBytes atau patch
          blob = new Blob([letter.uploadedDocx.rawBytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        } else if (letter.document) {
          blob = await renderDocx(letter.document, 'blob');
        } else {
          blob = await renderDocx(store.getState().document, 'blob');
        }

        const safeTitle = (letter.title || 'surat_dinas').replace(/[^a-zA-Z0-9_\-]/g, '_').slice(0, 45);
        downloadFile(blob, `${safeTitle}.docx`);
        showToast('Berkas .docx berhasil diunduh.', 'success');
      } catch (err) {
        showToast(`Gagal mengunduh: ${err.message}`, 'error');
      }
    };
  });

  // 10. Duplikasi Surat
  container.querySelectorAll('.btn-duplicate-letter').forEach(btn => {
    btn.onclick = () => {
      const id = btn.dataset.id;
      const letter = store.getLetterVault().find(l => l.id === id);
      if (!letter) return;

      const duplicated = {
        ...JSON.parse(JSON.stringify(letter)),
        id: `letter-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        title: `(Salinan) ${letter.title}`,
        createdAt: new Date().toISOString()
      };

      store.addLetterVaultEntry(duplicated);
      showToast('Surat berhasil diduplikasi.', 'success');
      renderVaultView(container);
    };
  });

  // 11. Hapus Surat dari Gudang
  container.querySelectorAll('.btn-delete-vault-letter').forEach(btn => {
    btn.onclick = () => {
      const id = btn.dataset.id;
      const letter = store.getLetterVault().find(l => l.id === id);
      if (!letter) return;

      if (confirm(`Apakah Anda yakin ingin menghapus "${letter.title}" dari Gudang Surat?`)) {
        store.removeLetterVaultEntry(id);
        showToast('Surat dihapus dari gudang.', 'info');
        renderVaultView(container);
      }
    };
  });
}

/**
 * Memproses berkas .docx yang diunggah pengguna:
 * Mengekstrak paragraf & metadata, menyimpan ke Gudang Surat dengan mode preservasi format,
 * dan memuatnya ke Lembar Kerja Word.
 */
async function processUploadedDocx(file, container) {
  if (!file.name.endsWith('.docx')) {
    showToast('Harap pilih berkas dengan ekstensi .docx Microsoft Word.', 'error');
    return;
  }

  showToast(`Membaca dan memproses struktur "${file.name}"...`, 'info');

  try {
    const arrayBuffer = await file.arrayBuffer();
    const parsed = await parseDocxForEditing(arrayBuffer);

    // Buat entri baru di Gudang Surat
    const newEntry = store.addLetterVaultEntry({
      title: parsed.meta.title || file.name.replace(/\.docx$/i, ''),
      nomorSurat: parsed.meta.nomor || '-',
      tanggalSurat: parsed.meta.date || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
      category: 'Dokumen Impor',
      targetPerson: parsed.meta.signer || '-',
      perihal: parsed.meta.title,
      summary: `Dokumen .docx resmi diunggah: "${file.name}". Berisi ${parsed.editableParagraphs.length} paragraf dengan preservasi format asli 100%.`,
      isAiReference: true,
      source: 'upload',
      uploadedDocx: {
        rawBytes: new Uint8Array(arrayBuffer),
        fileName: file.name,
        parsedMeta: parsed.meta,
        structured: parsed.structured,
        paragraphs: parsed.editableParagraphs
      }
    });

    // Otomatis atur berkas aktif di Lembar Kerja Word
    store.setActiveUploadedDocx(newEntry.uploadedDocx);

    showToast(`Berkas "${file.name}" berhasil disimpan ke Gudang Surat & siap diedit!`, 'success');
    renderVaultView(container);

    // Tanya pengguna apakah ingin langsung membukanya di Word
    if (confirm(`Berkas "${file.name}" telah tersimpan di Gudang Surat.\n\nApakah Anda ingin langsung membukanya di Lembar Kerja Word untuk mengedit?`)) {
      const canvasTab = document.querySelector('.nav-tab[data-tab="canvas"]');
      if (canvasTab) canvasTab.click();
    }
  } catch (err) {
    showToast(`Gagal memproses berkas .docx: ${err.message}`, 'error');
  }
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str || '';
  return div.innerHTML;
}
