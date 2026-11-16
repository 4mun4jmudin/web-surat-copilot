/**
 * Print & Download Signature Mode Choice Modal + Shared QR Authorization Helpers
 *
 * Enforces:
 * - Barcode QR only appears if the letter is already declared SAH.
 * - On print/PDF or .docx download, user chooses between:
 *   1) Barcode QR Pengesahan (elektronik), OR
 *   2) Kosongan TTD Manual (untuk tanda tangan basah pakai pulpen).
 * - If not yet SAH, offers 1-click "Sahkan & Cetak dengan Barcode QR" (with selectable Kepsek / TU / Kepsek & TU)
 *   or immediate "Kosongan TTD Manual (Pulpen)".
 */

import { store } from '../state.js';
import { showToast } from './notification.js';
import {
  SIGNER_ROLE_OPTIONS,
  buildSignerSelectionInfo,
  buildQrCaptionLines,
  buildVerificationPayload,
  generateQrPngDataUrl,
  registerVerificationInCloud
} from '../letters/verification.js';

export function escPrintHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Shared HTML snippet for Signer Authorization Selector inside Catat Surat Keluar modals.
 */
export function buildSignerRoleSelectorHtml({
  radioName = 'out-signer-role',
  checkboxId = 'chk-out-authorize-sah',
  defaultRole = 'kepala_sekolah',
  school = {}
} = {}) {
  const roles = SIGNER_ROLE_OPTIONS.map(opt => {
    const info = buildSignerSelectionInfo(school, opt.id);
    const checked = opt.id === defaultRole ? 'checked' : '';
    return `
      <label class="signer-role-option">
        <input type="radio" name="${escPrintHtml(radioName)}" value="${escPrintHtml(opt.id)}" ${checked} />
        <div class="signer-role-card">
          <div class="signer-role-title">${escPrintHtml(opt.label)}</div>
          <div class="signer-role-name">${escPrintHtml(info.signerName)}</div>
        </div>
      </label>
    `;
  }).join('');

  return `
    <div class="authorize-box">
      <label class="authorize-toggle-row">
        <input type="checkbox" id="${escPrintHtml(checkboxId)}" checked />
        <div>
          <strong>🛡️ Nyatakan SAH &amp; Terbitkan Barcode QR Verifikasi</strong>
          <small>Barcode QR akan aktif saat cetak PDF / unduh .docx jika surat dinyatakan SAH.</small>
        </div>
      </label>
      <div class="signer-role-section">
        <div class="signer-role-heading">Pernyataan Pengesahan Atas Nama:</div>
        <div class="signer-role-grid">${roles}</div>
      </div>
    </div>
  `;
}

/**
 * Unified helper to declare SAH, generate Barcode QR PNG, sync to Netlify Blobs (/api/verify),
 * and record/lock the entry in Buku Surat Keluar.
 */
export async function authorizeAndRecordOutgoing({
  entryBase,
  school,
  signerRole = 'kepala_sekolah',
  authorizeSah = true,
  syncActiveDocMode = '', // 'standard' | 'smart' | ''
  clearAfterSave = false
} = {}) {
  const effectiveSchool = school || store.getState().schoolProfile || {};
  const signerInfo = buildSignerSelectionInfo(effectiveSchool, signerRole);

  if (!authorizeSah) {
    const recorded = store.addOutgoingLetter({
      ...entryBase,
      signerRole: signerInfo.signerRole,
      signerTitle: signerInfo.signerTitle,
      signerName: entryBase.signerName || signerInfo.signerName,
      signerNip: signerInfo.signerNip,
      signerRank: signerInfo.signerRank,
      isVerifiedSah: false,
      verificationStatus: 'unverified',
      cloudSynced: false,
      skipAutoAdvanceNumber: !clearAfterSave
    });
    if (!clearAfterSave && syncActiveDocMode) {
      store.setActiveDocumentVerification(null, syncActiveDocMode);
    }
    return { recorded, qrDataUrl: '', payload: null, cloudResult: { ok: false, skipped: true } };
  }

  const payload = buildVerificationPayload({
    entry: entryBase,
    school: effectiveSchool,
    signerRole
  });
  const qrDataUrl = generateQrPngDataUrl(payload.verificationUrl);
  const cloudResult = await registerVerificationInCloud(payload);

  const recorded = store.addOutgoingLetter({
    ...entryBase,
    signerRole: payload.signerRole,
    signerTitle: payload.signerTitle,
    signerName: payload.signerName,
    signerNip: payload.signerNip,
    signerRank: payload.signerRank,
    verificationId: payload.verificationId,
    verificationUrl: payload.verificationUrl,
    qrDataUrl,
    manageToken: payload.manageToken,
    verificationStatement: payload.verificationStatement,
    verificationDetails: payload.details,
    isVerifiedSah: true,
    verificationStatus: 'sah',
    cloudSynced: Boolean(cloudResult.ok),
    skipAutoAdvanceNumber: !clearAfterSave
  });

  if (!clearAfterSave && syncActiveDocMode) {
    store.setActiveDocumentVerification({
      outgoingId: recorded.id,
      isVerifiedSah: true,
      verificationStatus: 'sah',
      verificationId: payload.verificationId,
      verificationUrl: payload.verificationUrl,
      qrDataUrl,
      manageToken: payload.manageToken,
      signerRole: payload.signerRole,
      roleLabel: signerInfo.roleLabel,
      signerTitle: payload.signerTitle,
      signerName: payload.signerName,
      signerNip: payload.signerNip,
      signerRank: payload.signerRank,
      verificationStatement: payload.verificationStatement,
      nomorSurat: recorded.nomorSurat,
      cloudSynced: Boolean(cloudResult.ok)
    }, syncActiveDocMode);
  }

  if (!cloudResult.ok) {
    showToast(
      `Surat #${String(recorded.sequenceNumber).padStart(3, '0')} disahkan secara lokal, namun sinkron cloud gagal: ${cloudResult.error || 'offline'}. Gunakan tombol Sync di Surat Keluar saat online.`,
      'warning'
    );
  }

  return { recorded, payload, qrDataUrl, cloudResult };
}

const PAPER_MM_DIMENSIONS = {
  a4: { w: '210mm', h: '297mm' },
  f4: { w: '215mm', h: '330mm' },
  letter: { w: '215.9mm', h: '279.4mm' },
  legal: { w: '215.9mm', h: '355.6mm' }
};

/**
 * Synchronize @page physical size with active document dimensions (A4/F4/Letter/Legal + orientation),
 * reset desk scroll offset so Page 1 top is never clipped, and trigger window.print().
 */
export function runSafeWindowPrint({
  paperSize,
  orientation,
  onAfterPrint = null,
  fallbackTimeoutMs = 1500
} = {}) {
  const opts = typeof store?.getLetterOptions === 'function' ? store.getLetterOptions() : {};
  const sizeKey = String(paperSize || opts?.paperSize || 'a4').toLowerCase();
  const orient = String(orientation || opts?.orientation || 'portrait').toLowerCase() === 'landscape' ? 'landscape' : 'portrait';
  const dim = PAPER_MM_DIMENSIONS[sizeKey] || PAPER_MM_DIMENSIONS.a4;
  const pageSizeCss = orient === 'landscape' ? `${dim.h} ${dim.w}` : `${dim.w} ${dim.h}`;

  document.getElementById('dynamic-print-page-style')?.remove();
  const pageStyleEl = document.createElement('style');
  pageStyleEl.id = 'dynamic-print-page-style';
  pageStyleEl.textContent = `@page { size: ${pageSizeCss}; margin: 0mm; }`;
  document.head.appendChild(pageStyleEl);

  const desks = Array.from(document.querySelectorAll('.word-desk'));
  const savedScrolls = desks.map(el => ({ el, top: el.scrollTop, left: el.scrollLeft }));
  desks.forEach(el => {
    el.scrollTop = 0;
    el.scrollLeft = 0;
  });

  let cleaned = false;
  const cleanup = () => {
    if (cleaned) return;
    cleaned = true;
    window.removeEventListener('afterprint', cleanup);
    pageStyleEl.remove();
    savedScrolls.forEach(({ el, top, left }) => {
      if (el && document.body.contains(el)) {
        el.scrollTop = top;
        el.scrollLeft = left;
      }
    });
    if (typeof onAfterPrint === 'function') onAfterPrint();
  };

  window.addEventListener('afterprint', cleanup, { once: true });
  requestAnimationFrame(() => {
    window.print();
    setTimeout(cleanup, fallbackTimeoutMs);
  });
}

/**
 * Build crisp visual QR box HTML for print sheets (Lembar Kerja, Word Pro, Surat Keluar).
 * Displays single clean caption "Disahkan secara elektronik" (signer name is printed below).
 */
export function buildPrintQrBoxHtml({
  qrDataUrl,
  signerRole = 'kepala_sekolah',
  principalName = '',
  signerName = '',
  school = null
} = {}) {
  const activeSchool = school || (typeof store?.getState === 'function' ? store.getState().schoolProfile : {}) || {};
  const caption = buildQrCaptionLines({
    signerRole,
    school: activeSchool,
    principalName,
    signerName
  });

  return `
    <div class="qr-sig-print-box" data-temp-qr="true">
      <img src="${escPrintHtml(qrDataUrl)}" alt="Barcode QR Pengesahan Surat" class="qr-sig-img" />
      <div class="qr-sig-caption">
        <div>${escPrintHtml(caption.line1)}</div>
      </div>
    </div>
  `.trim();
}

/**
 * Temporarily inject QR code into the live canvas signature holder right before window.print(),
 * and restore the clean manual blank state immediately after printing completes.
 */
export function printWithTemporaryQrOverlay({
  mode = 'manual', // 'qr' | 'manual'
  standardHolderEl = null,
  smartSigCellEl = null,
  smartFallbackEditableEl = null,
  qrDataUrl = '',
  signerRole = 'kepala_sekolah',
  principalName = '',
  signerName = '',
  school = null,
  paperSize,
  orientation,
  onAfterPrint = null
} = {}) {
  let restoreFn = () => {};

  if (mode === 'qr' && qrDataUrl) {
    const qrHtml = buildPrintQrBoxHtml({
      qrDataUrl,
      signerRole,
      principalName,
      signerName,
      school
    });
    if (standardHolderEl) {
      const prevHtml = standardHolderEl.innerHTML;
      const hadActive = standardHolderEl.classList.contains('has-qr-print');
      standardHolderEl.innerHTML = qrHtml;
      standardHolderEl.classList.add('has-qr-print');
      restoreFn = () => {
        standardHolderEl.innerHTML = prevHtml;
        if (!hadActive) standardHolderEl.classList.remove('has-qr-print');
      };
    } else if (smartSigCellEl) {
      const prevHtml = smartSigCellEl.innerHTML;
      // Ganti baris kosong beruntun (<br><br><br>) di atas nama penanda tangan dengan QR
      const replaced = prevHtml.replace(/(<br\s*\/?>\s*){2,}/i, qrHtml);
      if (replaced !== prevHtml) {
        smartSigCellEl.innerHTML = replaced;
      } else {
        const strongNode = smartSigCellEl.querySelector('strong, u, b');
        if (strongNode) {
          strongNode.insertAdjacentHTML('beforebegin', qrHtml);
        } else {
          smartSigCellEl.insertAdjacentHTML('beforeend', qrHtml);
        }
      }
      restoreFn = () => {
        smartSigCellEl.innerHTML = prevHtml;
      };
    } else if (smartFallbackEditableEl) {
      const tempNode = document.createElement('div');
      tempNode.style.cssText = 'display:flex; justify-content:flex-end; margin-top:20px; page-break-inside:avoid; break-inside:avoid;';
      tempNode.innerHTML = `<div style="width:230px;">${qrHtml}</div>`;
      smartFallbackEditableEl.appendChild(tempNode);
      restoreFn = () => tempNode.remove();
    }
  }

  runSafeWindowPrint({
    paperSize,
    orientation,
    onAfterPrint: () => {
      restoreFn();
      if (typeof onAfterPrint === 'function') onAfterPrint();
    }
  });
}

export function openPrintSignatureChoiceModal({
  actionType = 'print', // 'print' | 'docx'
  isVerifiedSah = false,
  isRevoked = false,
  nomorSurat = '-',
  title = 'SURAT DINAS',
  qrDataUrl = '',
  defaultSignerRole = 'kepala_sekolah',
  signerLabel = 'Kepala Sekolah',
  signerName = '',
  onChooseQr,
  onChooseManual,
  onAuthorizeAndQr
} = {}) {
  document.getElementById('modal-print-signature-choice')?.remove();

  const verbLabel = actionType === 'docx' ? 'Unduh .docx' : 'Cetak / PDF';
  const canUseQrDirectly = Boolean(isVerifiedSah) && !isRevoked && Boolean(qrDataUrl);

  const overlay = document.createElement('div');
  overlay.id = 'modal-print-signature-choice';
  overlay.className = 'modal-overlay';

  overlay.innerHTML = `
    <div class="modal-card modal-card-choice" style="max-width: 610px; width: calc(100% - 24px);">
      <div class="modal-header" style="display: flex; align-items: center; justify-content: space-between; border-bottom: 1px solid #e2e8f0; padding-bottom: 12px; margin-bottom: 14px;">
        <div>
          <h3 style="margin: 0; font-size: 16px; color: #0f172a; display: flex; align-items: center; gap: 8px;">
            <span>${actionType === 'docx' ? '📥' : '🖨️'}</span>
            <span>Pilih Mode Tanda Tangan (${escPrintHtml(verbLabel)})</span>
          </h3>
          <p style="margin: 4px 0 0 0; font-size: 11.5px; color: #64748b;">
            <strong>${escPrintHtml(title)}</strong> • Nomor: <span style="font-family: monospace;">${escPrintHtml(nomorSurat)}</span>
          </p>
        </div>
        <button type="button" class="btn-icon-ghost" id="btn-close-print-choice" title="Tutup">✕</button>
      </div>

      ${isRevoked ? `
        <div class="choice-status-banner choice-banner-danger">
          <strong>⛔ Pengesahan Surat Ini Telah Dicabut</strong>
          <span>Barcode QR tidak aktif selama surat dalam status dicabut. Anda hanya dapat mencetak dengan ruang tanda tangan manual, atau sahkan ulang di menu Surat Keluar.</span>
        </div>
      ` : canUseQrDirectly ? `
        <div class="choice-status-banner choice-banner-sah">
          <strong>✓ Status Dokumen: SUDAH DINYATAKAN SAH</strong>
          <span>Disahkan atas nama <strong>${escPrintHtml(signerLabel)}</strong>${signerName ? ` (${escPrintHtml(signerName)})` : ''}. Pilih apakah ingin menampilkan Barcode QR atau dikosongkan untuk TTD pulpen.</span>
        </div>
      ` : `
        <div class="choice-status-banner choice-banner-warn">
          <strong>⚠️ Draft Belum Dinyatakan SAH</strong>
          <span>Barcode QR hanya terbit setelah surat dinyatakan <strong>SAH</strong> &amp; dicatat ke Surat Keluar. Anda bisa langsung mengesahkannya di bawah ini atau mencetak kosongan untuk TTD manual pulpen.</span>
        </div>
      `}

      <div class="choice-options-grid">
        ${canUseQrDirectly ? `
          <div class="choice-option-card choice-option-qr">
            <div class="choice-card-top">
              <div class="choice-qr-thumb">
                <img src="${escPrintHtml(qrDataUrl)}" alt="Preview QR Pengesahan" />
              </div>
              <div>
                <div class="choice-card-badge badge-sah">Pengesahan Digital</div>
                <h4>Pakai Barcode QR Pengesahan</h4>
                <p>QR Code resmi dicetak di pojok tanda tangan (atas nama <strong>${escPrintHtml(signerLabel)}</strong>) agar bisa dipindai kamera HP.</p>
              </div>
            </div>
            <button type="button" id="btn-choice-use-qr" class="btn btn-primary" style="width: 100%; margin-top: 10px;">
              ▣ ${escPrintHtml(verbLabel)} Pakai Barcode QR
            </button>
          </div>
        ` : !isRevoked && typeof onAuthorizeAndQr === 'function' ? `
          <div class="choice-option-card choice-option-qr">
            <div class="choice-card-top">
              <div class="choice-qr-placeholder">QR</div>
              <div>
                <div class="choice-card-badge badge-pending">Perlu Pengesahan SAH</div>
                <h4>Sahkan Surat &amp; Pakai Barcode QR</h4>
                <p>Otomatis mencatat ke Surat Keluar sebagai surat <strong>SAH</strong> dan menerbitkan Barcode QR verifikasi resmi.</p>
              </div>
            </div>
            <div style="margin-top: 10px;">
              <label style="display: block; font-size: 11px; font-weight: 700; color: #334155; margin-bottom: 5px;">
                Pernyataan Pengesahan Atas Nama:
              </label>
              <select id="choice-authorize-role" class="form-control" style="font-size: 12px; height: 34px;">
                ${SIGNER_ROLE_OPTIONS.map(opt => `
                  <option value="${escPrintHtml(opt.id)}" ${opt.id === defaultSignerRole ? 'selected' : ''}>
                    ${escPrintHtml(opt.label)} — ${escPrintHtml(opt.description)}
                  </option>
                `).join('')}
              </select>
            </div>
            <button type="button" id="btn-choice-authorize-qr" class="btn btn-primary" style="width: 100%; margin-top: 10px;">
              🛡️ Sahkan &amp; ${escPrintHtml(verbLabel)} Pakai Barcode QR
            </button>
          </div>
        ` : `
          <div class="choice-option-card choice-option-disabled">
            <div class="choice-card-top">
              <div class="choice-qr-placeholder">⛔</div>
              <div>
                <div class="choice-card-badge badge-revoked">QR Dinonaktifkan</div>
                <h4>Barcode QR Tidak Tersedia</h4>
                <p>Status surat saat ini tidak sah / dicabut. Gunakan opsi TTD manual atau sahkan ulang surat terlebih dahulu.</p>
              </div>
            </div>
          </div>
        `}

        <div class="choice-option-card choice-option-manual">
          <div class="choice-card-top">
            <div class="choice-manual-preview">
              <span>Area TTD</span>
              <span class="choice-pen-line">____________</span>
            </div>
            <div>
              <div class="choice-card-badge badge-manual">Tanda Tangan Basah</div>
              <h4>Kosongan TTD Manual (Pulpen)</h4>
              <p>Ruang tanda tangan tetap dikosongkan tanpa QR agar Kepala Sekolah / TU dapat tanda tangan &amp; stempel basah dengan pulpen.</p>
            </div>
          </div>
          <button type="button" id="btn-choice-manual-sign" class="btn btn-secondary" style="width: 100%; margin-top: 10px; border: 1px solid #cbd5e1; background: #fff; color: #0f172a;">
            ✍️ ${escPrintHtml(verbLabel)} Kosongan TTD Manual
          </button>
        </div>
      </div>

      <div style="display: flex; justify-content: flex-end; margin-top: 14px; border-top: 1px solid #e2e8f0; padding-top: 10px;">
        <button type="button" id="btn-cancel-print-choice" class="btn btn-secondary" style="font-size: 12px;">Batal</button>
      </div>
    </div>
  `;

  document.body.appendChild(overlay);

  const cleanup = () => {
    window.removeEventListener('keydown', onKeyDown);
    overlay.remove();
  };

  const onKeyDown = (e) => {
    if (e.key === 'Escape') cleanup();
  };
  window.addEventListener('keydown', onKeyDown);

  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) cleanup();
  });
  overlay.querySelector('#btn-close-print-choice').onclick = cleanup;
  overlay.querySelector('#btn-cancel-print-choice').onclick = cleanup;

  const btnQr = overlay.querySelector('#btn-choice-use-qr');
  if (btnQr) {
    btnQr.onclick = async () => {
      cleanup();
      if (typeof onChooseQr === 'function') await onChooseQr();
    };
  }

  const btnAuthQr = overlay.querySelector('#btn-choice-authorize-qr');
  if (btnAuthQr) {
    btnAuthQr.onclick = async () => {
      const selectedRole = overlay.querySelector('#choice-authorize-role')?.value || 'kepala_sekolah';
      btnAuthQr.disabled = true;
      btnAuthQr.textContent = 'Menyimpan pengesahan & membuat QR...';
      try {
        if (typeof onAuthorizeAndQr === 'function') {
          await onAuthorizeAndQr(selectedRole);
        }
      } finally {
        cleanup();
      }
    };
  }

  const btnManual = overlay.querySelector('#btn-choice-manual-sign');
  if (btnManual) {
    btnManual.onclick = async () => {
      cleanup();
      if (typeof onChooseManual === 'function') await onChooseManual();
    };
  }
}
