/**
 * qrFloatingHelper.js - Mesin Posisi Barcode QR Mengambang (Floating Layer) & Pratinjau Cetak Presisi
 *
 * Fitur utama:
 * 1. Mengubah Barcode QR menjadi objek mengambang bebas (seperti gambar di MS Word)
 *    yang tidak pernah memengaruhi alur paragraf/tabel dan tidak pernah menimbulkan halaman kosong saat cetak.
 * 2. Bebas digeser (drag-and-drop) antar posisi maupun antar halaman, serta diubah ukurannya (resize).
 * 3. Mendukung pemilihan lapisan: Di Depan Teks (In Front of Text) vs Di Belakang Teks (Behind Text).
 * 4. Mendukung toggle Tampil vs Sembunyikan di Kanvas (saat disembunyikan tetap muncul ketika Cetak/PDF & Surat Keluar).
 */

import { buildQrCaptionLines } from '../letters/verification.js';

const PAPER_DIMS_MM = {
  a4: { w: 210, h: 297 },
  f4: { w: 215, h: 330 },
  letter: { w: 215.9, h: 279.4 },
  legal: { w: 215.9, h: 355.6 }
};

function escapeAttrHtml(val) {
  return String(val ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function clampNum(val, min, max, fallback) {
  const n = Number(val);
  if (!Number.isFinite(n)) return fallback;
  return Math.max(min, Math.min(max, n));
}

function normCm(val, fallback = 2.0) {
  const n = Number(val);
  if (!Number.isFinite(n) || n <= 0) return fallback;
  const cm = n > 50 ? n / 567 : n;
  return clampNum(cm, 0.4, 8.0, fallback);
}

/**
 * Menghitung dimensi area cetak bersih (mm) setelah dikurangi margin untuk menjaga presisi print
 * agar elemen absolut tidak meluap dan memicu halaman kosong tambahan.
 */
export function computePrintableDimsMm({ paperSize = 'a4', orientation = 'portrait', margins = {} } = {}) {
  const key = String(paperSize || 'a4').toLowerCase();
  const base = PAPER_DIMS_MM[key] || PAPER_DIMS_MM.a4;
  const isLand = String(orientation || 'portrait').toLowerCase() === 'landscape';
  const sheetW = isLand ? base.h : base.w;
  const sheetH = isLand ? base.w : base.h;

  const mTopMm = normCm(margins?.top, 2.0) * 10;
  const mBottomMm = normCm(margins?.bottom, 2.0) * 10;
  const mLeftMm = normCm(margins?.left, 2.5) * 10;
  const mRightMm = normCm(margins?.right, 2.0) * 10;

  return {
    sheetWidthMm: sheetW,
    sheetHeightMm: sheetH,
    printableWidthMm: Number(Math.max(60, sheetW - mLeftMm - mRightMm).toFixed(2)),
    // Kurangi 2mm sebagai bantalan sub-pixel Chrome print agar tidak tumpah ke halaman baru
    printableHeightMm: Number(Math.max(80, sheetH - mTopMm - mBottomMm - 2).toFixed(2))
  };
}

/**
 * Menormalkan metadata posisi & visibilitas Barcode QR mengambang.
 */
export function normalizeQrPlacement(raw = null, pageCount = 1) {
  const maxPageIdx = Math.max(0, (Number(pageCount) || 1) - 1);
  const src = raw && typeof raw === 'object' ? raw : {};
  return {
    pageIndex: Math.min(maxPageIdx, Math.max(0, Math.round(Number(src.pageIndex) || 0))),
    xPct: Number(clampNum(src.xPct, 0, 92, 68).toFixed(2)),
    yPct: Number(clampNum(src.yPct, 0, 92, 64).toFixed(2)),
    sizePx: Math.round(clampNum(src.sizePx, 48, 190, 86)),
    layer: src.layer === 'behind' ? 'behind' : 'front',
    visibleOnCanvas: src.visibleOnCanvas !== false,
    showCaption: src.showCaption !== false,
    customPositioned: Boolean(src.customPositioned)
  };
}

/**
 * Menentukan halaman & koordinat awal terbaik berdasarkan isi DocumentModel
 * (misal: Surat Tugas halaman 1 dengan `qrSlot` atau blok `letterSignature`).
 */
export function findPreferredSignatureQrPlacement(docModel = null, pageCount = 1) {
  const blocks = Array.isArray(docModel?.blocks) ? docModel.blocks : [];
  let pageCursor = 0;
  let qrSlotPage = -1;
  let sigPage = -1;

  for (const b of blocks) {
    if (!b) continue;
    if (b.type === 'pageBreak') {
      pageCursor += 1;
      continue;
    }
    if (b.type === 'table' && Array.isArray(b.rows)) {
      const hasSlot = b.rows.some((r) => {
        const cells = Array.isArray(r) ? r : (r?.cells || []);
        return cells.some(c =>
          c?.qrSlot || (Array.isArray(c?.paragraphs) && c.paragraphs.some(p => p?.qrSlot))
        );
      });
      if (hasSlot && qrSlotPage === -1) {
        qrSlotPage = pageCursor;
      }
    }
    if (b.type === 'letterSignature') {
      sigPage = pageCursor;
    }
  }

  const targetPage = sigPage >= 0
    ? sigPage
    : (qrSlotPage >= 0 ? qrSlotPage : Math.max(0, pageCount - 1));

  return normalizeQrPlacement({
    pageIndex: targetPage,
    xPct: 68,
    yPct: sigPage >= 0 ? 64 : 58,
    sizePx: 86,
    layer: 'front',
    visibleOnCanvas: true,
    showCaption: true,
    customPositioned: false
  }, pageCount);
}

/**
 * Mengukur posisi slot TTD aktual di DOM untuk kalibrasi awal agar QR langsung mendarat di atas ruang TTD.
 */
export function calibratePlacementFromDom(stages = [], currentPlacement = {}, fallbackSizePx = 86) {
  if (!Array.isArray(stages) || stages.length === 0) return currentPlacement;
  if (currentPlacement?.customPositioned) return currentPlacement;

  for (let idx = 0; idx < stages.length; idx++) {
    const stage = stages[idx];
    const sheet = stage?.closest('.smart-paper-sheet, .word-paper-sheet');
    if (!sheet) continue;
    const targetEl = sheet.querySelector('[data-qr-slot="true"], .sig-space-holder, .smart-sig-cell');
    if (!targetEl) continue;

    const stageRect = stage.getBoundingClientRect();
    const tRect = targetEl.getBoundingClientRect();
    if (stageRect.width <= 10 || stageRect.height <= 10 || tRect.width <= 0) continue;

    const zoomScale = stageRect.width / Math.max(1, stage.offsetWidth || stageRect.width);
    const visualSize = fallbackSizePx * zoomScale;

    let xLeftPx = (tRect.left - stageRect.left) + Math.max(0, (tRect.width - visualSize) / 2);
    let yTopPx = (tRect.top - stageRect.top);
    if (targetEl.classList.contains('smart-sig-cell')) {
      yTopPx += Math.min(44 * zoomScale, tRect.height * 0.34);
    }

    const xPct = clampNum((xLeftPx / stageRect.width) * 100, 2, 85, 68);
    const yPct = clampNum((yTopPx / stageRect.height) * 100, 2, 85, 62);
    return normalizeQrPlacement({
      ...currentPlacement,
      pageIndex: idx,
      xPct,
      yPct,
      sizePx: fallbackSizePx
    }, stages.length);
  }
  return currentPlacement;
}

/**
 * Menghasilkan markup HTML elemen Barcode QR mengambang untuk kanvas interaktif maupun pratinjau/cetak.
 */
export function buildFloatingQrElementHtml({
  qrDataUrl = '',
  placement = {},
  signerRole = 'kepala_sekolah',
  principalName = '',
  signerName = '',
  school = {},
  interactive = false,
  pageCount = 1,
  repositioning = false
} = {}) {
  if (!qrDataUrl) return '';
  const p = normalizeQrPlacement(placement, pageCount);
  const caption = buildQrCaptionLines({ signerRole, school, principalName, signerName });
  const layerClass = p.layer === 'behind' ? 'layer-behind' : 'layer-front';
  const repositionClass = interactive && repositioning ? 'is-repositioning' : '';
  const canDragDirectly = interactive && (p.layer === 'front' || repositioning);

  const resizeHandleHtml = interactive
    ? `<div class="floating-qr-resize no-print" data-qr-resize="true" draggable="false" title="Tarik untuk ubah ukuran Barcode QR"></div>`
    : '';

  const interactiveTitle = interactive
    ? ' title="Seret untuk memindah • Klik kanan untuk opsi Depan/Belakang Teks, Halaman, Ukuran &amp; Sembunyikan"'
    : '';

  return `
    <div class="floating-qr-box ${layerClass} ${interactive ? 'is-interactive' : ''} ${repositionClass}"
      data-floating-qr="true"
      contenteditable="false"
      draggable="false"
      style="left: ${p.xPct}%; top: ${p.yPct}%; width: ${p.sizePx}px;">
      ${resizeHandleHtml}
      <div class="floating-qr-visual" ${canDragDirectly ? 'data-qr-drag="true"' : ''}${interactiveTitle}>
        <img src="${escapeAttrHtml(qrDataUrl)}" alt="Barcode QR Pengesahan" class="floating-qr-img" draggable="false" />
        ${p.showCaption ? `<div class="floating-qr-caption">${escapeAttrHtml(caption.line1)}</div>` : ''}
      </div>
    </div>
  `.trim();
}

/**
 * Menghasilkan markup HTML menu konteks klik kanan (ala MS Word) untuk mengatur Barcode QR.
 */
export function buildFloatingQrContextMenuHtml({ placement = {}, pageCount = 1 } = {}) {
  const p = normalizeQrPlacement(placement, pageCount);
  const totalPages = Math.max(1, Number(pageCount) || 1);
  const pageBtnsHtml = totalPages > 1
    ? `
      <div class="fqr-ctx-section">
        <div class="fqr-ctx-label">Pindah ke Halaman</div>
        <div class="fqr-ctx-pill-row">
          ${Array.from({ length: totalPages }, (_, i) => `
            <button type="button" class="fqr-ctx-pill ${p.pageIndex === i ? 'active' : ''}" data-qr-action="move-page" data-page="${i}">
              H${i + 1}
            </button>
          `).join('')}
        </div>
      </div>
    `
    : '';

  const behindDragOptionHtml = p.layer === 'behind'
    ? `
      <button type="button" class="fqr-ctx-item" data-qr-action="enable-reposition">
        <span class="fqr-ctx-check">✥</span>
        <span>Geser / Ubah Ukuran Saat Ini</span>
      </button>
    `
    : '';

  return `
    <div class="fqr-context-header">
      <span>Opsi Barcode QR</span>
      <span class="fqr-ctx-page-badge">Hal ${p.pageIndex + 1}/${totalPages}</span>
    </div>
    <div class="fqr-ctx-group">
      <button type="button" class="fqr-ctx-item ${p.layer === 'front' ? 'active' : ''}" data-qr-action="set-layer" data-layer="front">
        <span class="fqr-ctx-check">${p.layer === 'front' ? '✓' : ''}</span>
        <span>Di Depan Teks</span>
      </button>
      <button type="button" class="fqr-ctx-item ${p.layer === 'behind' ? 'active' : ''}" data-qr-action="set-layer" data-layer="behind">
        <span class="fqr-ctx-check">${p.layer === 'behind' ? '✓' : ''}</span>
        <span>Di Belakang Teks</span>
      </button>
      ${behindDragOptionHtml}
    </div>
    ${pageBtnsHtml}
    <div class="fqr-ctx-section">
      <div class="fqr-ctx-label-row">
        <span>Ukuran Barcode</span>
        <span class="fqr-ctx-size-val" data-qr-size-val>${p.sizePx} px</span>
      </div>
      <div class="fqr-ctx-size-controls">
        <button type="button" class="fqr-ctx-step-btn" data-qr-action="step-size" data-delta="-8" title="Perkecil">−</button>
        <button type="button" class="fqr-ctx-pill ${p.sizePx <= 74 ? 'active' : ''}" data-qr-action="set-size" data-size="70">Kecil</button>
        <button type="button" class="fqr-ctx-pill ${p.sizePx > 74 && p.sizePx < 102 ? 'active' : ''}" data-qr-action="set-size" data-size="86">Normal</button>
        <button type="button" class="fqr-ctx-pill ${p.sizePx >= 102 ? 'active' : ''}" data-qr-action="set-size" data-size="112">Besar</button>
        <button type="button" class="fqr-ctx-step-btn" data-qr-action="step-size" data-delta="8" title="Perbesar">+</button>
      </div>
    </div>
    <div class="fqr-ctx-group fqr-ctx-divider">
      <button type="button" class="fqr-ctx-item ${p.showCaption ? 'active' : ''}" data-qr-action="toggle-caption">
        <span class="fqr-ctx-check">${p.showCaption ? '✓' : ''}</span>
        <span>Teks "Disahkan secara elektronik"</span>
      </button>
      <button type="button" class="fqr-ctx-item fqr-ctx-item-danger" data-qr-action="hide-canvas">
        <span class="fqr-ctx-check">👁️</span>
        <span>Sembunyikan di Kanvas</span>
      </button>
    </div>
  `.trim();
}

let activeContextMenuCleanup = null;

/**
 * Menutup menu konteks Barcode QR yang sedang terbuka jika ada.
 */
export function closeActiveQrContextMenu() {
  if (typeof activeContextMenuCleanup === 'function') {
    const fn = activeContextMenuCleanup;
    activeContextMenuCleanup = null;
    fn();
  }
}

/**
 * Memperbarui tombol Tampil/Sembunyi Barcode QR di ribbon / status bar kanvas.
 */
export function updateCanvasQrToggleButton(btnEl, { verification = null, placement = null } = {}) {
  if (!btnEl) return;
  const isRevoked = verification?.verificationStatus === 'revoked';
  const hasSahQr = Boolean(verification?.isVerifiedSah && verification?.qrDataUrl && !isRevoked);

  if (!hasSahQr) {
    btnEl.classList.add('hidden');
    btnEl.classList.remove('active', 'is-qr-hidden-mode');
    return;
  }

  btnEl.classList.remove('hidden');
  const visible = placement?.visibleOnCanvas !== false;
  btnEl.classList.toggle('active', visible);
  btnEl.classList.toggle('is-qr-hidden-mode', !visible);

  const labelEl = btnEl.querySelector('[data-qr-btn-label]') || btnEl;
  labelEl.textContent = visible ? 'QR Tampil' : 'Tampil QR';
  btnEl.title = visible
    ? 'Barcode QR sedang tampil di kanvas (seret untuk geser, klik kanan pada QR untuk opsi Depan/Belakang Teks, Halaman, Ukuran, atau Sembunyikan).'
    : 'Barcode QR disembunyikan di kanvas agar fokus mengetik (tetap tercetak saat Print & muncul di Surat Keluar). Klik untuk munculkan di kanvas.';
}

/**
 * Memasang & mengikat interaksi drag-and-drop bebas, lintas halaman, resize sudut,
 * serta menu klik kanan (Depan/Belakang Teks, Pindah Halaman, Ukuran, Teks, Sembunyi).
 */
export function mountInteractiveFloatingQr({
  stages = [],
  verification = null,
  rawPlacement = null,
  docModel = null,
  school = {},
  getZoomScale = () => 1,
  onPlacementChange = null
} = {}) {
  closeActiveQrContextMenu();
  const stageEls = Array.from(stages || []).filter(Boolean);
  stageEls.forEach(s => {
    if (typeof s._fqrCleanup === 'function') {
      s._fqrCleanup();
      s._fqrCleanup = null;
    }
    s.innerHTML = '';
  });

  const isRevoked = verification?.verificationStatus === 'revoked';
  if (!verification?.isVerifiedSah || !verification?.qrDataUrl || isRevoked || stageEls.length === 0) {
    return null;
  }

  let placement = rawPlacement
    ? normalizeQrPlacement(rawPlacement, stageEls.length)
    : findPreferredSignatureQrPlacement(docModel, stageEls.length);

  if (!rawPlacement?.customPositioned) {
    placement = calibratePlacementFromDom(stageEls, placement, placement.sizePx);
  }

  if (placement.visibleOnCanvas === false) {
    return placement;
  }

  let isRepositioningBehind = false;

  const commit = (nextPatch, reRender = false) => {
    const merged = normalizeQrPlacement({ ...placement, ...nextPatch }, stageEls.length);
    placement = merged;
    if (typeof onPlacementChange === 'function') onPlacementChange(merged);
    if (reRender) {
      if (merged.visibleOnCanvas === false) {
        closeActiveQrContextMenu();
        stageEls.forEach(s => { s.innerHTML = ''; });
      } else {
        renderAtStage(merged);
      }
    }
    return merged;
  };

  const openContextMenuAt = (clientX, clientY) => {
    closeActiveQrContextMenu();
    if (typeof document === 'undefined' || !document.body) return;

    const menuEl = document.createElement('div');
    menuEl.className = 'fqr-context-menu no-print';
    menuEl.setAttribute('contenteditable', 'false');
    menuEl.innerHTML = buildFloatingQrContextMenuHtml({
      placement,
      pageCount: stageEls.length
    });

    const vw = (typeof window !== 'undefined' && window.innerWidth) ? window.innerWidth : 1280;
    const vh = (typeof window !== 'undefined' && window.innerHeight) ? window.innerHeight : 800;
    const estW = 240;
    const estH = 330;
    const initLeft = Math.max(8, Math.min(clientX, vw - estW - 8));
    const initTop = Math.max(8, Math.min(clientY, vh - estH - 8));
    menuEl.style.left = `${initLeft}px`;
    menuEl.style.top = `${initTop}px`;
    document.body.appendChild(menuEl);

    if (typeof menuEl.getBoundingClientRect === 'function') {
      const mRect = menuEl.getBoundingClientRect();
      if (mRect.width > 10 && mRect.height > 10) {
        const clampedL = Math.max(8, Math.min(clientX, vw - mRect.width - 8));
        const clampedT = Math.max(8, Math.min(clientY, vh - mRect.height - 8));
        menuEl.style.left = `${clampedL}px`;
        menuEl.style.top = `${clampedT}px`;
      }
    }

    const onOutsideDown = (ev) => {
      if (!ev.target?.closest?.('.fqr-context-menu')) {
        closeActiveQrContextMenu();
      }
    };
    const onEscKey = (ev) => {
      if (ev.key === 'Escape') closeActiveQrContextMenu();
    };
    const onScrollOrResize = () => closeActiveQrContextMenu();

    document.addEventListener('mousedown', onOutsideDown, true);
    document.addEventListener('pointerdown', onOutsideDown, true);
    document.addEventListener('keydown', onEscKey);
    if (typeof window !== 'undefined') {
      window.addEventListener('resize', onScrollOrResize);
    }

    activeContextMenuCleanup = () => {
      document.removeEventListener('mousedown', onOutsideDown, true);
      document.removeEventListener('pointerdown', onOutsideDown, true);
      document.removeEventListener('keydown', onEscKey);
      if (typeof window !== 'undefined') {
        window.removeEventListener('resize', onScrollOrResize);
      }
      menuEl.remove();
    };

    const syncMenuSizeUi = () => {
      const sizeValEl = menuEl.querySelector('[data-qr-size-val]');
      if (sizeValEl) sizeValEl.textContent = `${placement.sizePx} px`;
      menuEl.querySelectorAll('[data-qr-action="set-size"]').forEach((btn) => {
        const sz = Number(btn.dataset.size);
        const isSmall = sz <= 74 && placement.sizePx <= 74;
        const isNorm = sz > 74 && sz < 102 && placement.sizePx > 74 && placement.sizePx < 102;
        const isLarge = sz >= 102 && placement.sizePx >= 102;
        btn.classList.toggle('active', isSmall || isNorm || isLarge);
      });
    };

    menuEl.addEventListener('mousedown', (ev) => {
      ev.preventDefault();
      ev.stopPropagation();
    });
    menuEl.addEventListener('pointerdown', (ev) => {
      ev.stopPropagation();
    });
    menuEl.addEventListener('contextmenu', (ev) => {
      ev.preventDefault();
      ev.stopPropagation();
    });

    menuEl.querySelectorAll('[data-qr-action]').forEach((btn) => {
      btn.addEventListener('click', (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        const act = btn.dataset.qrAction;
        if (act === 'set-layer') {
          const nextLayer = btn.dataset.layer === 'behind' ? 'behind' : 'front';
          isRepositioningBehind = false;
          closeActiveQrContextMenu();
          commit({ layer: nextLayer, customPositioned: true }, true);
        } else if (act === 'enable-reposition') {
          isRepositioningBehind = true;
          closeActiveQrContextMenu();
          renderAtStage(placement);
          const curBox = stageEls[placement.pageIndex]?.querySelector('[data-floating-qr="true"]');
          curBox?.classList.add('is-selected');
        } else if (act === 'move-page') {
          const nextIdx = Number(btn.dataset.page);
          closeActiveQrContextMenu();
          if (!Number.isNaN(nextIdx)) {
            commit({ pageIndex: nextIdx, customPositioned: true }, true);
          }
        } else if (act === 'step-size') {
          const delta = Number(btn.dataset.delta || 0);
          const nextSize = Math.round(clampNum(placement.sizePx + delta, 48, 190, placement.sizePx));
          commit({ sizePx: nextSize, customPositioned: true }, false);
          const curBox = stageEls[placement.pageIndex]?.querySelector('[data-floating-qr="true"]');
          if (curBox) curBox.style.width = `${nextSize}px`;
          syncMenuSizeUi();
        } else if (act === 'set-size') {
          const nextSize = Math.round(clampNum(Number(btn.dataset.size), 48, 190, placement.sizePx));
          commit({ sizePx: nextSize, customPositioned: true }, false);
          const curBox = stageEls[placement.pageIndex]?.querySelector('[data-floating-qr="true"]');
          if (curBox) curBox.style.width = `${nextSize}px`;
          syncMenuSizeUi();
        } else if (act === 'toggle-caption') {
          closeActiveQrContextMenu();
          commit({ showCaption: !placement.showCaption, customPositioned: true }, true);
        } else if (act === 'hide-canvas') {
          isRepositioningBehind = false;
          closeActiveQrContextMenu();
          commit({ visibleOnCanvas: false }, true);
        }
      });
    });
  };

  function renderAtStage(targetPlacement) {
    placement = normalizeQrPlacement(targetPlacement, stageEls.length);
    stageEls.forEach((stage, idx) => {
      if (typeof stage._fqrCleanup === 'function') {
        stage._fqrCleanup();
        stage._fqrCleanup = null;
      }
      if (idx === placement.pageIndex) {
        stage.innerHTML = buildFloatingQrElementHtml({
          qrDataUrl: verification.qrDataUrl,
          placement,
          signerRole: verification.signerRole || 'kepala_sekolah',
          principalName: school.principalName || '',
          signerName: verification.signerName || '',
          school,
          interactive: true,
          pageCount: stageEls.length,
          repositioning: isRepositioningBehind
        });
      } else {
        stage.innerHTML = '';
      }
    });

    const activeStage = stageEls[placement.pageIndex];
    const box = activeStage?.querySelector('[data-floating-qr="true"]');
    if (!box) return;

    const ownerSheet = activeStage.closest?.('.smart-paper-sheet, .word-paper-sheet') || activeStage.parentNode;

    const clearSelectionState = (ev) => {
      if (ev.target?.closest?.('[data-floating-qr="true"], .fqr-context-menu')) return;
      box.classList.remove('is-selected');
      if (isRepositioningBehind) {
        isRepositioningBehind = false;
        renderAtStage(placement);
      }
    };

    const onSheetContextMenu = (ev) => {
      const directQr = ev.target?.closest?.('[data-floating-qr="true"]');
      if (directQr && directQr === box) {
        ev.preventDefault();
        ev.stopPropagation();
        box.classList.add('is-selected');
        openContextMenuAt(ev.clientX, ev.clientY);
        return;
      }
      // Saat mode Belakang Teks (pointer-events: none), klik kanan tetap mendeteksi area QR di balik teks
      if (placement.layer === 'behind') {
        const bRect = box.getBoundingClientRect();
        if (
          bRect.width > 0 && bRect.height > 0 &&
          ev.clientX >= bRect.left && ev.clientX <= bRect.right &&
          ev.clientY >= bRect.top && ev.clientY <= bRect.bottom
        ) {
          ev.preventDefault();
          ev.stopPropagation();
          openContextMenuAt(ev.clientX, ev.clientY);
        }
      }
    };

    if (ownerSheet) {
      ownerSheet.addEventListener('contextmenu', onSheetContextMenu, true);
    }
    if (typeof document !== 'undefined') {
      document.addEventListener('mousedown', clearSelectionState, true);
    }

    activeStage._fqrCleanup = () => {
      if (ownerSheet) {
        ownerSheet.removeEventListener('contextmenu', onSheetContextMenu, true);
      }
      if (typeof document !== 'undefined') {
        document.removeEventListener('mousedown', clearSelectionState, true);
      }
    };

    let activeGesture = null;

    box.addEventListener('dragstart', (e) => {
      e.preventDefault();
    });

    box.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      e.stopPropagation();
      box.classList.add('is-selected');
      openContextMenuAt(e.clientX, e.clientY);
    });

    const resolveEffectiveZoom = () => {
      const stageRect = activeStage.getBoundingClientRect();
      const offsetW = activeStage.offsetWidth || 0;
      if (offsetW > 10 && stageRect.width > 10) {
        return Math.max(0.25, stageRect.width / offsetW);
      }
      return Math.max(0.25, Number(getZoomScale()) || 1);
    };

    const attachGesture = (triggerEl, mode, onStart) => {
      if (!triggerEl) return;
      const begin = (ev) => {
        if (typeof ev.button === 'number' && ev.button !== 0) return;
        if (ev.target?.closest?.('[data-qr-resize]') && mode === 'drag') return;
        ev.preventDefault();
        ev.stopPropagation();
        if (activeGesture) return;

        closeActiveQrContextMenu();
        const session = onStart(ev);
        if (!session) return;
        activeGesture = mode;
        box.classList.add('is-selected');

        const finish = (endEv) => {
          window.removeEventListener('pointermove', onMove, true);
          window.removeEventListener('pointerup', finish, true);
          window.removeEventListener('pointercancel', finish, true);
          window.removeEventListener('mousemove', onMove, true);
          window.removeEventListener('mouseup', finish, true);
          window.removeEventListener('blur', finish);
          activeGesture = null;
          box.classList.remove('is-dragging');
          session.onEnd?.(endEv);
        };

        const onMove = (mEv) => {
          if (typeof mEv.buttons === 'number' && mEv.buttons === 0 && ev.type === 'mousedown') {
            finish(mEv);
            return;
          }
          box.classList.add('is-dragging');
          session.onMove(mEv);
        };

        window.addEventListener('pointermove', onMove, true);
        window.addEventListener('pointerup', finish, true);
        window.addEventListener('pointercancel', finish, true);
        window.addEventListener('mousemove', onMove, true);
        window.addEventListener('mouseup', finish, true);
        window.addEventListener('blur', finish);
      };

      triggerEl.addEventListener('pointerdown', begin);
      triggerEl.addEventListener('mousedown', begin);
    };

    // Resize proportional via bottom-right corner handle
    const resizeHandle = box.querySelector('[data-qr-resize="true"]');
    attachGesture(resizeHandle, 'resize', (ev) => {
      const startX = ev.clientX;
      const startSize = placement.sizePx;
      const zoom = resolveEffectiveZoom();
      return {
        onMove: (mEv) => {
          const dx = (mEv.clientX - startX) / zoom;
          const nextSize = Math.round(clampNum(startSize + dx, 48, 190, startSize));
          placement = { ...placement, sizePx: nextSize, customPositioned: true };
          box.style.width = `${nextSize}px`;
        },
        onEnd: () => {
          commit({ customPositioned: true }, false);
        }
      };
    });

    // Free drag directly on QR box (supports moving X/Y and crossing across pages)
    box.querySelectorAll('[data-qr-drag="true"]').forEach((dragZone) => {
      attachGesture(dragZone, 'drag', (ev) => {
        const startX = ev.clientX;
        const startY = ev.clientY;
        const startRect = box.getBoundingClientRect();
        const grabOffsetX = ev.clientX - startRect.left;
        const grabOffsetY = ev.clientY - startRect.top;
        let currentStageIdx = placement.pageIndex;
        let lastX = placement.xPct;
        let lastY = placement.yPct;
        let didDrag = false;

        return {
          onMove: (mEv) => {
            if (!didDrag && Math.hypot(mEv.clientX - startX, mEv.clientY - startY) < 3) return;
            didDrag = true;

            for (let i = 0; i < stageEls.length; i++) {
              const sRect = stageEls[i].getBoundingClientRect();
              if (sRect.width > 0 && mEv.clientY >= sRect.top - 18 && mEv.clientY <= sRect.bottom + 18) {
                if (i !== currentStageIdx) {
                  currentStageIdx = i;
                  stageEls[i].appendChild(box);
                }
                break;
              }
            }

            const curStage = stageEls[currentStageIdx] || activeStage;
            const curRect = curStage.getBoundingClientRect();
            if (curRect.width <= 10 || curRect.height <= 10) return;

            const boxRect = box.getBoundingClientRect();
            const maxLeft = Math.max(1, curRect.width - boxRect.width);
            const maxTop = Math.max(1, curRect.height - boxRect.height);

            const rawLeft = clampNum(mEv.clientX - grabOffsetX - curRect.left, 0, maxLeft, 0);
            const rawTop = clampNum(mEv.clientY - grabOffsetY - curRect.top, 0, maxTop, 0);

            lastX = Number(((rawLeft / curRect.width) * 100).toFixed(2));
            lastY = Number(((rawTop / curRect.height) * 100).toFixed(2));
            box.style.left = `${lastX}%`;
            box.style.top = `${lastY}%`;
          },
          onEnd: () => {
            if (!didDrag) return;
            const movedBetweenPages = currentStageIdx !== placement.pageIndex;
            commit({
              pageIndex: currentStageIdx,
              xPct: lastX,
              yPct: lastY,
              customPositioned: true
            }, movedBetweenPages);
          }
        };
      });
    });
  }

  renderAtStage(placement);
  return placement;
}
