/**
 * vaultSaveHelper.js - Pembantu Bersama Status Badge & Proteksi Penyimpanan Master Terkunci (⭐)
 * Dipakai bersama oleh Lembar Kerja Word dan Word Pro agar perilaku Save / Save As konsisten.
 */

import { extractDocumentMetadata, enhanceDocumentMetadataWithAi } from '../letters/smartNaming.js';
import { getActiveAiConfig } from '../ai/client.js';

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

/**
 * Toolbar ringkas untuk pindai ulang Format Sistem maupun Pindai AI pada modal simpan/catat surat.
 */
export function buildMetaAutoScanBarHtml({ statusText = 'Penamaan otomatis format sistem siap' } = {}) {
  return `
    <div class="meta-scan-bar" style="display:flex;align-items:center;justify-content:space-between;gap:8px;flex-wrap:wrap;padding:7px 10px;background:#f8fafc;border:1px solid #cbd5e1;border-radius:8px;margin-bottom:4px;">
      <span data-meta-scan-status="true" style="font-size:11px;font-weight:600;color:#334155;">${escapeLocalHtml(statusText)}</span>
      <div style="display:inline-flex;align-items:center;gap:6px;flex-wrap:wrap;">
        <button type="button" class="btn btn-sm btn-secondary" data-meta-action="system" style="font-size:11px;padding:3px 9px;height:auto;">Format Sistem</button>
        <button type="button" class="btn btn-sm btn-secondary" data-meta-action="ai" style="font-size:11px;padding:3px 9px;height:auto;border-color:#86efac;color:#156c3b;background:#ecfdf5;font-weight:600;">Pindai AI</button>
      </div>
    </div>
  `.trim();
}

/**
 * Menghubungkan tombol Format Sistem & Pindai AI serta auto-scan AI non-blocking bila API key aktif.
 */
export function wireMetaAutoScanBar({
  rootEl,
  getContext,
  fields = {},
  forTemplate = false,
  autoAiOnOpen = true,
  onMetaApplied = null
} = {}) {
  if (!rootEl) return () => {};
  const statusEl = rootEl.querySelector('[data-meta-scan-status="true"]');
  const btnSys = rootEl.querySelector('[data-meta-action="system"]');
  const btnAi = rootEl.querySelector('[data-meta-action="ai"]');

  let userEdited = false;
  let activeAbort = null;
  let disposed = false;

  const markEdited = () => { userEdited = true; };
  [fields.titleEl, fields.recipientEl, fields.perihalEl, fields.dateEl, fields.categoryEl]
    .filter(Boolean)
    .forEach((el) => el.addEventListener('input', markEdited));

  const applyMeta = (meta, sourceLabel = 'Format sistem diterapkan') => {
    if (!meta || disposed) return;
    if (fields.titleEl) {
      fields.titleEl.value = meta.formattedTitle || meta.rawTitle || fields.titleEl.value;
      fields.titleEl.dispatchEvent(new Event('input', { bubbles: true }));
      // Dispatch input tidak boleh dianggap manual edit saat baru di-set oleh sistem/AI
      userEdited = false;
    }
    if (fields.perihalEl) fields.perihalEl.value = meta.perihal || meta.keperluan || '';
    if (fields.recipientEl) fields.recipientEl.value = meta.recipientName || '';
    if (fields.dateEl && meta.tanggalSurat) fields.dateEl.value = meta.tanggalSurat;
    if (fields.categoryEl && meta.category) fields.categoryEl.value = meta.category;
    if (fields.typeSelectEl && meta.typeCode) fields.typeSelectEl.value = meta.typeCode;
    if (statusEl) statusEl.textContent = sourceLabel;
    if (typeof onMetaApplied === 'function') onMetaApplied(meta);
  };

  if (btnSys) {
    btnSys.addEventListener('click', () => {
      userEdited = false;
      const ctx = typeof getContext === 'function' ? getContext() : {};
      const meta = extractDocumentMetadata({ ...ctx, forTemplate });
      applyMeta(meta, 'Format sistem diterapkan ulang');
    });
  }

  const triggerAiScan = async (isAuto = false) => {
    const aiCfg = getActiveAiConfig();
    if (!aiCfg?.apiKey) {
      if (!isAuto && statusEl) {
        statusEl.textContent = 'API key AI belum aktif — memakai format sistem';
      }
      return;
    }
    if (activeAbort) {
      try { activeAbort.abort(); } catch {}
    }
    activeAbort = typeof AbortController !== 'undefined' ? new AbortController() : null;
    if (btnAi) btnAi.disabled = true;
    if (statusEl) statusEl.textContent = 'Memindai isi surat dengan AI...';

    try {
      const ctx = typeof getContext === 'function' ? getContext() : {};
      const meta = await enhanceDocumentMetadataWithAi({
        ...ctx,
        forTemplate,
        signal: activeAbort?.signal || null
      });
      if (disposed || !document.body.contains(rootEl)) return;
      if (isAuto && userEdited) {
        if (statusEl) statusEl.textContent = 'Editan manual dipertahankan';
        return;
      }
      applyMeta(
        meta,
        meta.scannedByAi ? 'Dipindai & dirapikan oleh AI' : 'Format sistem diterapkan (fallback AI)'
      );
    } catch {
      if (!disposed && statusEl) {
        statusEl.textContent = 'Format sistem diterapkan (AI tidak merespons)';
      }
    } finally {
      if (btnAi) btnAi.disabled = false;
    }
  };

  if (btnAi) {
    btnAi.addEventListener('click', () => triggerAiScan(false));
  }

  if (autoAiOnOpen && getActiveAiConfig()?.apiKey) {
    setTimeout(() => {
      if (!disposed && !userEdited) {
        triggerAiScan(true);
      }
    }, 60);
  }

  return () => {
    disposed = true;
    if (activeAbort) {
      try { activeAbort.abort(); } catch {}
    }
  };
}
