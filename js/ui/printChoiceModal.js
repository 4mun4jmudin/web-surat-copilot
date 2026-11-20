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
  applySignatureModeToDocModel,
  buildSignerSelectionInfo,
  buildQrCaptionLines,
  buildVerificationPayload,
  generateQrPngDataUrl,
  registerVerificationInCloud
} from '../letters/verification.js';
import {
  canSignWithRole,
  getSelectableSignerRoles,
  resolveDefaultAllowedSignerRole
} from '../auth/permissions.js';
import {
  blobToBase64,
  getStoredAuthToken,
  recordOutgoingCloud
} from '../sync/workspaceClient.js';
import { setLastKnownSyncRev } from '../sync/liveSync.js';
import { renderDocx } from '../docx/render.js';
import {
  computePrintableDimsMm,
  normalizeQrPlacement,
  findPreferredSignatureQrPlacement,
  calibratePlacementFromDom,
  buildFloatingQrElementHtml
} from './qrFloatingHelper.js';

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
  const st = store.getState();
  const selectable = getSelectableSignerRoles(st.authUser, st.authPermissions, SIGNER_ROLE_OPTIONS);
  const resolvedDefault = resolveDefaultAllowedSignerRole(st.authUser, st.authPermissions, defaultRole);
  const isDelegated = selectable.some(s => s.id === 'kepala_sekolah' && s.allowed && s.delegated);
  const canRequestAcc = st.authUser?.role === 'tu' && Boolean(getStoredAuthToken());
  const hasLockedKepsekOption = selectable.some(s => !s.allowed);

  const roles = selectable.map(opt => {
    const info = buildSignerSelectionInfo(school, opt.id);
    const checked = opt.id === resolvedDefault ? 'checked' : '';
    const allowSelection = opt.allowed || canRequestAcc;
    const disabledAttr = allowSelection ? '' : 'disabled';
    const requiresAccAttr = !opt.allowed && canRequestAcc ? '1' : '0';
    const lockTag = !opt.allowed
      ? (canRequestAcc
          ? `<div style="font-size:10px;color:#b45309;font-weight:700;margin-top:2px;">📨 Minta ACC Kepsek</div>`
          : `<div style="font-size:10px;color:#b45309;font-weight:700;margin-top:2px;">🔒 Perlu Izin Kepsek/Admin</div>`)
      : (opt.delegated
          ? `<div style="font-size:10px;color:#15803d;font-weight:700;margin-top:2px;">✓ Izin Delegasi Aktif</div>`
          : '');
    return `
      <label class="signer-role-option ${opt.allowed ? '' : (canRequestAcc ? 'is-request-acc' : 'is-locked')}" title="${escPrintHtml(opt.lockReason || opt.description)}">
        <input type="radio" name="${escPrintHtml(radioName)}" value="${escPrintHtml(opt.id)}" data-requires-acc="${requiresAccAttr}" ${checked} ${disabledAttr} />
        <div class="signer-role-card">
          <div class="signer-role-title">${escPrintHtml(opt.label)}</div>
          <div class="signer-role-name">${escPrintHtml(info.signerName)}</div>
          ${lockTag}
        </div>
      </label>
    `;
  }).join('');

  return `
    <div class="authorize-box">
      <label class="authorize-toggle-row">
        <input type="checkbox" id="${escPrintHtml(checkboxId)}" checked />
        <div>
          <strong>🛡️ Nyatakan SAH / Ajukan ACC &amp; Terbitkan Barcode QR</strong>
          <small>Barcode QR akan aktif setelah pengesahan SAH${isDelegated ? ' (Delegasi Pengesahan Kepala Sekolah ke TU Aktif)' : ''}. ${canRequestAcc && hasLockedKepsekOption ? 'Jika memilih Kepala Sekolah tanpa izin delegasi, surat otomatis masuk ke antrean ACC Kepala Sekolah.' : ''}</small>
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
 * Unified helper to declare SAH, generate Barcode QR PNG, sync to Netlify Cloud + Google Drive,
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
  const st = store.getState();
  const effectiveSchool = school || st.schoolProfile || {};
  const effectiveRole = signerRole || resolveDefaultAllowedSignerRole(st.authUser, st.authPermissions);

  if (authorizeSah) {
    const roleCheck = canSignWithRole(st.authUser, st.authPermissions, effectiveRole);
    if (!roleCheck.allowed) {
      showToast(roleCheck.reason || 'Anda tidak memiliki izin pengesahan untuk pejabat ini.', 'error');
      throw new Error(roleCheck.reason || 'Akses pengesahan ditolak.');
    }
  }

  const signerInfo = buildSignerSelectionInfo(effectiveSchool, effectiveRole);
  const authToken = getStoredAuthToken();

  let docxBase64 = '';
  if (authToken && entryBase?.document && typeof window !== 'undefined') {
    try {
      const previewDocForDrive = authorizeSah
        ? entryBase.document
        : applySignatureModeToDocModel(entryBase.document, { mode: 'manual' });
      const blob = await renderDocx(previewDocForDrive, 'blob');
      docxBase64 = await blobToBase64(blob);
    } catch {}
  }

  if (!authorizeSah) {
    const draftEntry = {
      ...entryBase,
      signerRole: signerInfo.signerRole,
      signerTitle: signerInfo.signerTitle,
      signerName: entryBase.signerName || signerInfo.signerName,
      signerNip: signerInfo.signerNip,
      signerRank: signerInfo.signerRank,
      isVerifiedSah: false,
      verificationStatus: 'unverified'
    };

    if (authToken) {
      const cloudSave = await recordOutgoingCloud({
        authorizeSah: false,
        signerRole: effectiveRole,
        docxBase64,
        entry: {
          ...draftEntry,
          documentSnapshot: entryBase.document || null
        }
      });
      if (cloudSave.ok && cloudSave.item) {
        const recorded = store.upsertOutgoingLetterFromCloud(cloudSave.item);
        if (!clearAfterSave && syncActiveDocMode) {
          store.setActiveDocumentVerification(null, syncActiveDocMode);
        } else if (clearAfterSave) {
          store.syncActiveDocumentNumberToNextOutgoing(true);
        }
        return { recorded, qrDataUrl: '', payload: null, cloudResult: cloudSave };
      }
    }

    const recorded = store.addOutgoingLetter({
      ...draftEntry,
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
    signerRole: effectiveRole
  });
  const qrDataUrl = generateQrPngDataUrl(payload.verificationUrl);

  if (authToken && entryBase?.document && typeof window !== 'undefined') {
    try {
      const qrDocForDrive = applySignatureModeToDocModel(entryBase.document, {
        mode: 'qr',
        qrDataUrl,
        verificationUrl: payload.verificationUrl,
        signerRole: effectiveRole,
        signerInfo,
        school: effectiveSchool,
        syncSignerFields: false
      });
      const qrBlob = await renderDocx(qrDocForDrive, 'blob');
      docxBase64 = await blobToBase64(qrBlob);
    } catch {}
  }

  if (authToken) {
    const cloudSave = await recordOutgoingCloud({
      authorizeSah: true,
      signerRole: effectiveRole,
      docxBase64,
      entry: {
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
        documentSnapshot: entryBase.document || null
      },
      verificationPayload: payload
    });

    if (cloudSave.forbidden) {
      showToast(cloudSave.error || 'Akses pengesahan ditolak server.', 'error');
      throw new Error(cloudSave.error || 'Forbidden');
    }

    if (cloudSave.ok && cloudSave.item) {
      const recorded = store.upsertOutgoingLetterFromCloud(cloudSave.item);
      if (cloudSave.syncRev) setLastKnownSyncRev(cloudSave.syncRev);
      if (clearAfterSave) {
        store.syncActiveDocumentNumberToNextOutgoing(true);
      } else if (syncActiveDocMode) {
        store.setActiveDocumentVerification({
          outgoingId: recorded.id,
          isVerifiedSah: true,
          verificationStatus: 'sah',
          verificationId: recorded.verificationId || payload.verificationId,
          verificationUrl: recorded.verificationUrl || payload.verificationUrl,
          qrDataUrl,
          manageToken: recorded.manageToken || payload.manageToken,
          signerRole: recorded.signerRole || payload.signerRole,
          roleLabel: signerInfo.roleLabel,
          signerTitle: recorded.signerTitle || payload.signerTitle,
          signerName: recorded.signerName || payload.signerName,
          signerNip: recorded.signerNip ?? payload.signerNip,
          signerRank: recorded.signerRank || payload.signerRank,
          verificationStatement: recorded.verificationStatement || payload.verificationStatement,
          nomorSurat: recorded.nomorSurat,
          cloudSynced: true
        }, syncActiveDocMode);
      }

      if (cloudSave.driveSync && !cloudSave.driveSync.ok && cloudSave.driveSync.error) {
        showToast(`Surat #${String(recorded.sequenceNumber).padStart(3, '0')} tersimpan di Server Cloud (Drive menyusul: ${cloudSave.driveSync.error}).`, 'info');
      }
      return { recorded, payload, qrDataUrl, cloudResult: cloudSave };
    }
  }

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

export function normalizeMarginToCm(val, fallbackCm = 2.0) {
  const num = Number(val);
  if (!Number.isFinite(num) || num <= 0) return fallbackCm;
  const cm = num > 50 ? num / 567 : num;
  return Number(Math.max(0.4, Math.min(8.0, cm)).toFixed(2));
}

export function buildPrintPageCssRule({ paperSize, orientation, margins } = {}) {
  const opts = typeof store?.getLetterOptions === 'function' ? store.getLetterOptions() : {};
  const sizeKey = String(paperSize || opts?.paperSize || 'a4').toLowerCase();
  const orient = String(orientation || opts?.orientation || 'portrait').toLowerCase() === 'landscape' ? 'landscape' : 'portrait';
  const dim = PAPER_MM_DIMENSIONS[sizeKey] || PAPER_MM_DIMENSIONS.a4;
  const pageSizeCss = orient === 'landscape' ? `${dim.h} ${dim.w}` : `${dim.w} ${dim.h}`;
  const rawMargins = margins || opts?.margins || { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 };
  const mTop = normalizeMarginToCm(rawMargins.top, 2.0);
  const mRight = normalizeMarginToCm(rawMargins.right, 2.0);
  const mBottom = normalizeMarginToCm(rawMargins.bottom, 2.0);
  const mLeft = normalizeMarginToCm(rawMargins.left, 2.5);
  return `@page { size: ${pageSizeCss}; margin: ${mTop}cm ${mRight}cm ${mBottom}cm ${mLeft}cm; }`;
}

export function syncGlobalPrintPageStyle(options = {}) {
  if (typeof document === 'undefined' || !document.head) return;
  let styleEl = document.getElementById('live-print-page-style');
  if (!styleEl) {
    styleEl = document.createElement('style');
    styleEl.id = 'live-print-page-style';
    document.head.appendChild(styleEl);
  }
  styleEl.textContent = buildPrintPageCssRule(options);
}

/**
 * Synchronize @page physical size and margins with active document dimensions (A4/F4/Letter/Legal + orientation),
 * reset desk scroll offset so Page 1 top is never clipped, and trigger window.print().
 */
export function runSafeWindowPrint({
  paperSize,
  orientation,
  margins,
  onAfterPrint = null,
  fallbackTimeoutMs = 1500
} = {}) {
  const cssRule = buildPrintPageCssRule({ paperSize, orientation, margins });
  syncGlobalPrintPageStyle({ paperSize, orientation, margins });

  document.getElementById('dynamic-print-page-style')?.remove();
  const pageStyleEl = document.createElement('style');
  pageStyleEl.id = 'dynamic-print-page-style';
  pageStyleEl.textContent = cssRule;
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
 * Temporarily inject floating QR code onto the live canvas paper sheet(s) right before window.print(),
 * without reflowing text/tables or causing blank overflow pages, and restore exact canvas state after printing.
 */
export function printWithTemporaryQrOverlay({
  mode = 'manual', // 'qr' | 'manual'
  pageSheetEls = null,
  qrPlacement = null,
  docModel = null,
  standardHolderEl = null,
  qrDataUrl = '',
  signerRole = 'kepala_sekolah',
  principalName = '',
  signerName = '',
  school = null,
  paperSize,
  orientation,
  margins,
  onAfterPrint = null
} = {}) {
  const restoreCallbacks = [];
  const sheets = Array.isArray(pageSheetEls) ? pageSheetEls.filter(Boolean) : [];
  const opts = typeof store?.getLetterOptions === 'function' ? store.getLetterOptions() : {};
  const dims = computePrintableDimsMm({
    paperSize: paperSize || opts?.paperSize || 'a4',
    orientation: orientation || opts?.orientation || 'portrait',
    margins: margins || opts?.margins
  });

  if (sheets.length > 0) {
    const stages = sheets.map((sheet, idx) => {
      const prevPrintable = sheet.style.getPropertyValue('--printable-h-mm');
      sheet.style.setProperty('--printable-h-mm', `${dims.printableHeightMm}mm`);
      let stage = sheet.querySelector(':scope > .floating-qr-stage');
      let createdTemp = false;
      if (!stage) {
        stage = document.createElement('div');
        stage.className = 'floating-qr-stage';
        stage.dataset.qrPage = String(idx);
        sheet.appendChild(stage);
        createdTemp = true;
      }
      const prevHtml = stage.innerHTML;
      restoreCallbacks.push(() => {
        if (prevPrintable) sheet.style.setProperty('--printable-h-mm', prevPrintable);
        else sheet.style.removeProperty('--printable-h-mm');
        if (createdTemp) stage.remove();
        else stage.innerHTML = prevHtml;
      });
      return stage;
    });

    if (mode === 'qr' && qrDataUrl) {
      let effectivePlacement = qrPlacement
        ? normalizeQrPlacement(qrPlacement, stages.length)
        : findPreferredSignatureQrPlacement(docModel, stages.length);
      if (!qrPlacement?.customPositioned) {
        effectivePlacement = calibratePlacementFromDom(stages, effectivePlacement, effectivePlacement.sizePx);
      }

      const floatingHtml = buildFloatingQrElementHtml({
        qrDataUrl,
        placement: effectivePlacement,
        signerRole,
        principalName,
        signerName,
        school: school || {},
        interactive: false,
        pageCount: stages.length
      });

      stages.forEach((stage, idx) => {
        stage.innerHTML = idx === effectivePlacement.pageIndex ? floatingHtml : '';
      });
    } else {
      // Mode manual TTD pulpen: pastikan barcode tersembunyi bersih saat dicetak
      stages.forEach((stage) => {
        stage.innerHTML = '';
      });
    }
  } else if (mode === 'qr' && qrDataUrl && standardHolderEl) {
    const qrHtml = buildPrintQrBoxHtml({
      qrDataUrl,
      signerRole,
      principalName,
      signerName,
      school
    });
    const prevHtml = standardHolderEl.innerHTML;
    const hadActive = standardHolderEl.classList.contains('has-qr-print');
    standardHolderEl.innerHTML = qrHtml;
    standardHolderEl.classList.add('has-qr-print');
    restoreCallbacks.push(() => {
      standardHolderEl.innerHTML = prevHtml;
      if (!hadActive) standardHolderEl.classList.remove('has-qr-print');
    });
  }

  runSafeWindowPrint({
    paperSize,
    orientation,
    margins,
    onAfterPrint: () => {
      for (const fn of restoreCallbacks) fn();
      if (typeof onAfterPrint === 'function') onAfterPrint();
    }
  });
}

/**
 * Helper terpadu untuk mencatat surat keluar sebagai antrean ACC Kepala Sekolah
 * (ketika TU belum diberi izin delegasi pengesahan oleh Kepala Sekolah).
 */
export async function requestAccForOutgoing({
  entryBase,
  school,
  requestedSignerRole = 'kepala_sekolah',
  requestNote = '',
  syncActiveDocMode = 'standard',
  clearAfterSave = false
} = {}) {
  const effectiveSchool = school || store.getState().schoolProfile || {};
  const signerInfo = buildSignerSelectionInfo(effectiveSchool, requestedSignerRole);
  const authToken = getStoredAuthToken();

  let docxBase64 = '';
  if (authToken && entryBase?.document && typeof window !== 'undefined') {
    try {
      const manualDoc = applySignatureModeToDocModel(entryBase.document, { mode: 'manual' });
      const blob = await renderDocx(manualDoc, 'blob');
      docxBase64 = await blobToBase64(blob);
    } catch {}
  }

  const pendingEntry = {
    ...entryBase,
    signerRole: signerInfo.signerRole,
    signerTitle: signerInfo.signerTitle,
    signerName: signerInfo.signerName,
    signerNip: signerInfo.signerNip,
    signerRank: signerInfo.signerRank,
    isVerifiedSah: false,
    verificationStatus: 'pending_acc',
    accStatus: 'pending'
  };

  if (authToken) {
    const cloudSave = await recordOutgoingCloud({
      authorizeSah: false,
      requestAcc: true,
      signerRole: signerInfo.signerRole,
      requestNote,
      docxBase64,
      entry: {
        ...pendingEntry,
        documentSnapshot: entryBase.document || null
      }
    });

    if (!cloudSave.ok || !cloudSave.item) {
      throw new Error(cloudSave.error || 'Gagal mengirim pengajuan ACC ke Kepala Sekolah.');
    }

    const recorded = store.upsertOutgoingLetterFromCloud(cloudSave.item);
    if (cloudSave.syncRev) setLastKnownSyncRev(cloudSave.syncRev);

    if (clearAfterSave) {
      store.syncActiveDocumentNumberToNextOutgoing(true);
    } else if (syncActiveDocMode) {
      store.setActiveDocumentVerification({
        outgoingId: recorded.id,
        isVerifiedSah: false,
        verificationStatus: 'pending_acc',
        accStatus: 'pending',
        signerRole: recorded.signerRole,
        roleLabel: signerInfo.roleLabel,
        signerTitle: recorded.signerTitle,
        signerName: recorded.signerName,
        nomorSurat: recorded.nomorSurat
      }, syncActiveDocMode);
    }

    return { recorded, cloudResult: cloudSave };
  }

  const recorded = store.addOutgoingLetter({
    ...pendingEntry,
    cloudSynced: false,
    skipAutoAdvanceNumber: !clearAfterSave
  });
  return { recorded, cloudResult: { ok: false, skipped: true } };
}

export function openPrintSignatureChoiceModal({
  actionType = 'print', // 'print' | 'docx'
  isVerifiedSah = false,
  isRevoked = false,
  isPendingAcc = false,
  isRejectedAcc = false,
  rejectedReason = '',
  nomorSurat = '-',
  title = 'SURAT DINAS',
  qrDataUrl = '',
  defaultSignerRole = 'kepala_sekolah',
  signerLabel = 'Kepala Sekolah',
  signerName = '',
  onChooseQr,
  onChooseManual,
  onAuthorizeAndQr,
  onRequestAcc
} = {}) {
  document.getElementById('modal-print-signature-choice')?.remove();

  const verbLabel = actionType === 'docx' ? 'Unduh .docx' : 'Cetak / PDF';
  const canUseQrDirectly = Boolean(isVerifiedSah) && !isRevoked && Boolean(qrDataUrl);
  const stNow = store.getState();
  const canRequestAccRole = typeof onRequestAcc === 'function' && stNow.authUser?.role === 'tu' && Boolean(getStoredAuthToken());

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
      ` : isPendingAcc ? `
        <div class="choice-status-banner choice-banner-pending">
          <strong style="display:flex;align-items:center;gap:8px;"><span class="acc-live-spinner acc-spinner-sm"></span> Menunggu ACC Kepala Sekolah</strong>
          <span>Surat ini berada dalam antrean ACC Kepala Sekolah. Barcode QR otomatis terbit tanpa refresh begitu di-ACC. Sementara itu Anda tetap bebas menggunakan TTD Manual atau membuat surat lainnya.</span>
        </div>
      ` : isRejectedAcc ? `
        <div class="choice-status-banner choice-banner-danger">
          <strong>❌ Permintaan ACC Ditolak Kepala Sekolah</strong>
          <span>Catatan Kepsek: ${escPrintHtml(rejectedReason || 'Perlu revisi sebelum pengesahan')}. Anda dapat mengajukan ulang permintaan ACC di bawah atau menggunakan TTD Manual.</span>
        </div>
      ` : `
        <div class="choice-status-banner choice-banner-warn">
          <strong>⚠️ Draft Belum Dinyatakan SAH</strong>
          <span>Barcode QR hanya terbit setelah surat dinyatakan <strong>SAH</strong>. ${canRequestAccRole ? 'Jika memilih Kepala Sekolah tanpa izin delegasi, Anda dapat langsung mengirim permintaan ACC ke Kepala Sekolah.' : 'Anda bisa langsung mengesahkannya di bawah ini atau mencetak kosongan TTD manual.'}</span>
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
        ` : isPendingAcc ? `
          <div class="choice-option-card choice-option-disabled">
            <div class="choice-card-top">
              <div class="choice-qr-placeholder"><span class="acc-live-spinner"></span></div>
              <div>
                <div class="choice-card-badge badge-pending">Antrean ACC Aktif</div>
                <h4>Menunggu ACC Kepala Sekolah</h4>
                <p>Barcode QR akan aktif otomatis secara real-time di aplikasi segera setelah Kepala Sekolah menekan tombol ACC.</p>
              </div>
            </div>
          </div>
        ` : !isRevoked && typeof onAuthorizeAndQr === 'function' ? `
          <div class="choice-option-card choice-option-qr">
            <div class="choice-card-top">
              <div class="choice-qr-placeholder">QR</div>
              <div>
                <div class="choice-card-badge badge-pending">Perlu Pengesahan SAH</div>
                <h4>Sahkan / Minta ACC Barcode QR</h4>
                <p>Mencatat ke Surat Keluar serta langsung menerbitkan Barcode QR atau mengajukan ke Antrean ACC Kepala Sekolah.</p>
              </div>
            </div>
            <div style="margin-top: 10px;">
              <label style="display: block; font-size: 11px; font-weight: 700; color: #334155; margin-bottom: 5px;">
                Pernyataan Pengesahan Atas Nama:
              </label>
              <select id="choice-authorize-role" class="form-control" style="font-size: 12px; height: 34px;">
                ${(() => {
                  const opts = getSelectableSignerRoles(stNow.authUser, stNow.authPermissions, SIGNER_ROLE_OPTIONS);
                  const defRole = resolveDefaultAllowedSignerRole(stNow.authUser, stNow.authPermissions, defaultSignerRole);
                  return opts.map(opt => {
                    const allowOpt = opt.allowed || canRequestAccRole;
                    const reqAcc = !opt.allowed && canRequestAccRole ? '1' : '0';
                    const optText = opt.allowed
                      ? `${opt.label} ${opt.delegated ? '(Izin Delegasi Aktif)' : `— ${opt.description}`}`
                      : (canRequestAccRole
                          ? `${opt.label} 📨 (Minta ACC Kepala Sekolah)`
                          : `${opt.label} 🔒 (Perlu Izin Kepsek/Admin)`);
                    return `<option value="${escPrintHtml(opt.id)}" data-requires-acc="${reqAcc}" ${opt.id === defRole ? 'selected' : ''} ${allowOpt ? '' : 'disabled'}>${escPrintHtml(optText)}</option>`;
                  }).join('');
                })()}
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

  const roleSelectEl = overlay.querySelector('#choice-authorize-role');
  const btnAuthQr = overlay.querySelector('#btn-choice-authorize-qr');
  if (btnAuthQr) {
    const syncAuthBtnState = () => {
      const selectedOpt = roleSelectEl?.options?.[roleSelectEl.selectedIndex];
      const requiresAcc = selectedOpt?.dataset?.requiresAcc === '1';
      btnAuthQr.textContent = requiresAcc
        ? '📨 Ajukan Minta ACC Kepala Sekolah'
        : `🛡️ Sahkan & ${verbLabel} Pakai Barcode QR`;
    };
    roleSelectEl?.addEventListener('change', syncAuthBtnState);
    syncAuthBtnState();

    btnAuthQr.onclick = async () => {
      const selectedOpt = roleSelectEl?.options?.[roleSelectEl.selectedIndex];
      const selectedRole = roleSelectEl?.value || 'kepala_sekolah';
      const requiresAcc = selectedOpt?.dataset?.requiresAcc === '1';
      btnAuthQr.disabled = true;
      btnAuthQr.textContent = requiresAcc
        ? 'Mengirim permintaan ACC ke Kepala Sekolah...'
        : 'Menyimpan pengesahan & membuat QR...';
      try {
        if (requiresAcc && typeof onRequestAcc === 'function') {
          await onRequestAcc(selectedRole);
        } else if (typeof onAuthorizeAndQr === 'function') {
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
