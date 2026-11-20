/**
 * vaultSaveHelper.js - Pembantu Bersama Status Badge & Proteksi Penyimpanan Master Terkunci (⭐)
 * Dipakai bersama oleh Lembar Kerja Word dan Word Pro agar perilaku Save / Save As konsisten.
 */

function escapeLocalHtml(str = '') {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Memperbarui teks & warna badge status file Gudang di bar status bawah kanvas.
 */
export function renderCanvasVaultBadge(badgeEl, activeEntry) {
  if (!badgeEl) return;
  if (activeEntry) {
    const isLockedMaster = Boolean(activeEntry.isLockedTemplate || activeEntry.isMasterTemplate);
    badgeEl.textContent = isLockedMaster
      ? `🔒⭐ Master Terkunci: ${activeEntry.title}`
      : `📂 File Gudang: ${activeEntry.title}`;
    badgeEl.classList.add('is-linked');
    badgeEl.classList.toggle('is-master-locked', isLockedMaster);
    badgeEl.title = isLockedMaster
      ? `"${activeEntry.title}" dikunci sebagai ⭐ Master Template. Tekan Ctrl+S untuk Save As salinan aman atau buka kunci; tekan Ctrl+Shift+S / F12 untuk Save As langsung.`
      : `Mengedit "${activeEntry.title}". Tekan Ctrl+S untuk Save (perbarui file ini) atau Ctrl+Shift+S / F12 untuk Save As (duplikat).`;
  } else {
    badgeEl.textContent = '📄 Draf Baru (Ctrl+S Simpan / Ctrl+Shift+S Save As)';
    badgeEl.classList.remove('is-linked', 'is-master-locked');
    badgeEl.title = 'Dokumen baru belum ditautkan ke Gudang. Tekan Ctrl+S untuk menyimpan pertama kali.';
  }
}

/**
 * Menampilkan dialog proteksi saat pengguna menekan Simpan (Ctrl+S) pada dokumen berstatus ⭐ Master Terkunci.
 */
export function openLockedMasterSaveChoiceModal(modalRoot, activeEntry, { clearAfterSave = false, onSaveAsCopy, onUnlockAndSave } = {}) {
  if (!modalRoot || !activeEntry) return;
  const safeTitle = escapeLocalHtml(activeEntry.title || 'Master Template');

  modalRoot.innerHTML = `
    <div class="modal-overlay">
      <div class="modal-box" style="max-width: 490px; width: 92%;">
        <div class="modal-header">
          <h3>🔒⭐ Dokumen Master Terkunci</h3>
          <button type="button" class="btn-icon" id="lock-save-close">&times;</button>
        </div>
        <div class="modal-body" style="display: grid; gap: 12px;">
          <div class="auth-alert" style="background: #fffbeb; border-color: #f59e0b; color: #92400e; font-size: 12.5px; line-height: 1.5;">
            <strong>"${safeTitle}"</strong> sedang ditandai sebagai <strong>⭐ Master Template Terkunci</strong> untuk menjaga keaslian acuan AI Copilot.
            Pilih <strong>Save As Salinan Baru</strong> agar Master asli tetap aman, atau buka kunci bila memang ingin menimpanya.
          </div>
        </div>
        <div class="modal-footer" style="display: flex; flex-wrap: wrap; gap: 8px; justify-content: flex-end;">
          <button type="button" class="btn btn-secondary" id="lock-save-cancel">Batal</button>
          <button type="button" class="btn btn-secondary" id="lock-save-unlock" style="border-color: #d97706; color: #92400e;" title="Membuka kunci Master ini dan menyimpan perubahan langsung ke dalamnya">
            🔓 Buka Kunci &amp; Timpa Master
          </button>
          <button type="button" class="btn btn-primary" id="lock-save-copy">
            📄 ${clearAfterSave ? 'Save As Salinan Baru & Kosongkan' : 'Save As Salinan Baru'}
          </button>
        </div>
      </div>
    </div>`;

  const close = () => { modalRoot.innerHTML = ''; };
  modalRoot.querySelector('#lock-save-close')?.addEventListener('click', close);
  modalRoot.querySelector('#lock-save-cancel')?.addEventListener('click', close);
  modalRoot.querySelector('#lock-save-copy')?.addEventListener('click', () => {
    close();
    if (typeof onSaveAsCopy === 'function') onSaveAsCopy();
  });
  modalRoot.querySelector('#lock-save-unlock')?.addEventListener('click', () => {
    close();
    if (typeof onUnlockAndSave === 'function') onUnlockAndSave();
  });
}
