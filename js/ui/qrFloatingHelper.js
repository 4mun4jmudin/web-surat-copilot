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

const SPPD_TRAVEL_EXCLUSION_RE = /(dengan kendaraan|tempat kedudukan pegawai|kantor yang dikunjungi|mobil\s*\/\s*umum|tiba kembali|berangkat lagi|dari pejabat pemberi perintah jalan|dari tempat yang dikunjungi)/i;
const SIGNATURE_BODY_PROSE_RE = /(yang bertanda tangan di bawah ini|diberi perintah jalan|menerangkan bahwa|dengan ini menugaskan|sebagaimana mestinya)/i;
const IDENTITY_LABEL_RE = /^\s*(?:nama(?:\s+lengkap)?|jabatan|unit\s+kerja|alamat|nip|nuptk|niy|pangkat|golongan|instansi|asal\s+sekolah|tempat\s*\/?\s*tgl)\s*:?$/i;
const KEY_VALUE_IDENTITY_LINE_RE = /^\s*(?::|\b(?:nama|jabatan|unit\s+kerja|alamat|pangkat|golongan)\b\s*:)/i;
const SIGNATURE_ROLE_RE = /\b(kepala\s+(?:sekolah|smk(?:\s+it)?|smks(?:\s+it)?|smp(?:\s+it)?|tata\s+usaha|madrasah|yayasan)|plt\.?\s*kepala|ketua\s+(?:yayasan|panitia)|wakil\s+kepala|koordinator|pembina|sekretaris|bendahara|yang\s+memberi\s+perintah|mengetahui)\b/i;
const NIP_OR_NUPTK_RE = /\b(?:nip|nuptk|niy|nrks)\b\.?\s*[:\-]/i;

function normalizeSigScanText(val = '') {
  return String(val || '').replace(/\s+/g, ' ').trim();
}

function buildKnownSignerTokens(school = {}) {
  const candidates = [
    school?.principalName,
    school?.tuName,
    'Yangyang Budiman',
    'Dian Nurhamdani'
  ];
  const tokens = [];
  for (const raw of candidates) {
    const str = normalizeSigScanText(raw);
    if (!str) continue;
    // Ambil nama inti sebelum gelar di belakang koma (mis. "Yangyang Budiman, S.Pd." -> "yangyang budiman")
    const base = str.split(',')[0].replace(/\b(dr|drs|dra|ir|h|hj|prof|kh|ust)\.?\b/gi, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
    if (base.length >= 5 && !tokens.includes(base)) {
      tokens.push(base);
    }
  }
  return tokens;
}

function matchesSignerName(text = '', signerTokens = []) {
  const lower = normalizeSigScanText(text).toLowerCase();
  if (!lower) return false;
  return signerTokens.some(tok => lower.includes(tok));
}

function isNonSignatureTravelCell(cellText = '', tableText = '', signerTokens = []) {
  if (SPPD_TRAVEL_EXCLUSION_RE.test(cellText)) return true;
  if (SPPD_TRAVEL_EXCLUSION_RE.test(tableText)) {
    // Dalam tabel SPPD Keterangan I/II, tolak sel kendaraan/tanggal/tujuan maupun sel kosong di baris tsb
    // kecuali sel memang berisi nama & jabatan kepala sekolah resmi.
    if (!matchesSignerName(cellText, signerTokens) && !SIGNATURE_ROLE_RE.test(cellText)) {
      return true;
    }
  }
  return false;
}

function extractBlockModelText(block) {
  if (!block || typeof block !== 'object') return '';
  if (Array.isArray(block.runs)) {
    return normalizeSigScanText(block.runs.map(r => r?.text || '').join(' '));
  }
  return normalizeSigScanText(block.text || '');
}

function extractCellModelText(cell) {
  if (!cell || typeof cell !== 'object') return '';
  if (Array.isArray(cell.paragraphs) && cell.paragraphs.length > 0) {
    return normalizeSigScanText(cell.paragraphs.map(p => extractBlockModelText(p)).join(' '));
  }
  return normalizeSigScanText(cell.text || '');
}

function estimateColumnCenterPct(colWidthsPct = null, colIndex = 0, totalCols = 1) {
  const cols = Math.max(1, totalCols);
  if (Array.isArray(colWidthsPct) && colWidthsPct.length === cols) {
    let acc = 0;
    for (let i = 0; i < colIndex; i++) acc += Number(colWidthsPct[i]) || (100 / cols);
    const w = Number(colWidthsPct[colIndex]) || (100 / cols);
    return clampNum(acc + w / 2 - 7, 10, 78, 64);
  }
  const slotW = 100 / cols;
  return clampNum(colIndex * slotW + slotW / 2 - 7, 10, 78, 64);
}

/**
 * Menentukan halaman & koordinat awal terbaik berdasarkan isi DocumentModel
 * (mendukung `qrSlot`, blok `letterSignature`, maupun tabel/paragraf tanda tangan bebas Word Pro
 * dan menghindari tabel kendaraan / keterangan SPPD halaman 3).
 */
export function findPreferredSignatureQrPlacement(docModel = null, pageCount = 1, school = {}) {
  const blocks = Array.isArray(docModel?.blocks) ? docModel.blocks : [];
  const signerTokens = buildKnownSignerTokens(school);
  const candidates = [];
  let pageCursor = 0;

  // Hitung jumlah blok per halaman untuk estimasi posisi vertikal
  const blocksPerPage = [0];
  for (const b of blocks) {
    if (!b) continue;
    if (b.type === 'pageBreak') {
      blocksPerPage.push(0);
    } else {
      blocksPerPage[blocksPerPage.length - 1] += 1;
    }
  }

  let blockIdxInPage = 0;
  for (const b of blocks) {
    if (!b) continue;
    if (b.type === 'pageBreak') {
      pageCursor += 1;
      blockIdxInPage = 0;
      continue;
    }

    const totalInPage = Math.max(1, blocksPerPage[pageCursor] || 1);
    const relY = clampNum(24 + (blockIdxInPage / totalInPage) * 52, 22, 74, 58);

    if (b.type === 'letterSignature') {
      candidates.push({
        pageIndex: pageCursor,
        xPct: 64,
        yPct: 58,
        score: 100,
        preferEarliest: true
      });
    } else if (b.type === 'identityTable') {
      // Jangan pernah memperlakukan tabel identitas ("Nama : Kepsek / Jabatan : Kepala Sekolah") sebagai area TTD
      blockIdxInPage += 1;
      continue;
    } else if (b.type === 'table' && Array.isArray(b.rows)) {
      const allCellTexts = [];
      b.rows.forEach((r) => {
        const cells = Array.isArray(r) ? r : (r?.cells || []);
        cells.forEach(c => allCellTexts.push(extractCellModelText(c)));
      });
      const tableText = normalizeSigScanText(allCellTexts.join(' '));
      const isTravelTable = SPPD_TRAVEL_EXCLUSION_RE.test(tableText);

      b.rows.forEach((r) => {
        const cells = Array.isArray(r) ? r : (r?.cells || []);
        const rowLeftLabel = normalizeSigScanText(cells[0] ? extractCellModelText(cells[0]) : '');
        const isIdentityRow = cells.length >= 2 && cells.length <= 3 && IDENTITY_LABEL_RE.test(rowLeftLabel);

        cells.forEach((c, cIdx) => {
          const cellText = extractCellModelText(c);
          const hasExplicitSlot = Boolean(c?.qrSlot || (Array.isArray(c?.paragraphs) && c.paragraphs.some(p => p?.qrSlot)));
          const xPct = estimateColumnCenterPct(b.colWidthsPct, cIdx, cells.length);

          if (hasExplicitSlot) {
            candidates.push({
              pageIndex: pageCursor,
              xPct,
              yPct: clampNum(relY, 34, 72, 56),
              score: 100,
              preferEarliest: true
            });
            return;
          }

          if (isIdentityRow || isTravelTable || isNonSignatureTravelCell(cellText, tableText, signerTokens)) {
            return;
          }
          if (!cellText || cellText.length > 260 || SIGNATURE_BODY_PROSE_RE.test(cellText) || KEY_VALUE_IDENTITY_LINE_RE.test(cellText)) return;

          const hasName = matchesSignerName(cellText, signerTokens);
          const hasRole = SIGNATURE_ROLE_RE.test(cellText);
          if (hasName || hasRole) {
            candidates.push({
              pageIndex: pageCursor,
              xPct,
              yPct: clampNum(relY, 36, 72, 58),
              score: hasName && hasRole ? 90 : (hasName ? 85 : 78),
              preferEarliest: b.rows.length <= 2
            });
          }
        });
      });
    } else if (b.type === 'paragraph' || b.type === 'heading') {
      const pText = extractBlockModelText(b);
      if (
        pText && pText.length <= 130
        && !SIGNATURE_BODY_PROSE_RE.test(pText)
        && !SPPD_TRAVEL_EXCLUSION_RE.test(pText)
        && !KEY_VALUE_IDENTITY_LINE_RE.test(pText)
      ) {
        const hasName = matchesSignerName(pText, signerTokens);
        const hasRole = SIGNATURE_ROLE_RE.test(pText) && !/:$/.test(pText);
        if (hasName || hasRole) {
          const isIndentedRight = Number(b.leftIndentCm || 0) >= 5.5
            || Number(b.firstLineIndentCm || 0) >= 5.5
            || b.alignment === 'right'
            || (typeof b.text === 'string' && /^\s{10,}/.test(b.text));
          candidates.push({
            pageIndex: pageCursor,
            xPct: isIndentedRight ? 64 : 62,
            yPct: clampNum(relY - 12, 24, 72, 56),
            score: hasName ? 84 : 72,
            preferEarliest: false
          });
        }
      }
    }

    blockIdxInPage += 1;
  }

  let best = null;
  for (const cand of candidates) {
    if (!best) {
      best = cand;
      continue;
    }
    if (cand.score !== best.score) {
      if (cand.score > best.score) best = cand;
      continue;
    }
    // Jika sama-sama blok tanda tangan resmi eksplisit (mis. ST hal 1 + SPPD hal 2), utamakan Surat Tugas utama
    if (cand.preferEarliest) {
      if (cand.pageIndex < best.pageIndex) best = cand;
    } else if (cand.pageIndex > best.pageIndex) {
      best = cand;
    }
  }

  const targetPage = best ? best.pageIndex : Math.max(0, pageCount - 1);

  return normalizeQrPlacement({
    pageIndex: targetPage,
    xPct: best ? best.xPct : 66,
    yPct: best ? best.yPct : 62,
    sizePx: 86,
    layer: 'front',
    visibleOnCanvas: true,
    showCaption: true,
    customPositioned: false
  }, pageCount);
}

function getTrimmedTextNodeRect(textNode, doc) {
  if (!textNode || !textNode.nodeValue || !doc || typeof doc.createRange !== 'function') return null;
  const raw = textNode.nodeValue;
  const start = raw.search(/\S/);
  if (start < 0) return null;
  const trimmedEndMatch = raw.match(/\S(?=\s*$)/);
  const end = trimmedEndMatch ? (trimmedEndMatch.index + 1) : raw.length;
  try {
    const range = doc.createRange();
    if (typeof range?.getBoundingClientRect !== 'function') return null;
    range.setStart(textNode, start);
    range.setEnd(textNode, Math.max(start + 1, end));
    const r = range.getBoundingClientRect();
    if (r && r.width > 0 && r.height > 0) return r;
  } catch {}
  return null;
}

function collectVisualTextSegments(rootEl) {
  const segments = [];
  if (!rootEl) return segments;
  const doc = rootEl.ownerDocument;
  const walk = (node) => {
    if (!node) return;
    if (node.nodeType === 3) {
      const txt = normalizeSigScanText(node.nodeValue || '');
      if (txt) {
        const hostEl = node.parentElement || node.parentNode;
        const rect = getTrimmedTextNodeRect(node, doc)
          || (hostEl && typeof hostEl.getBoundingClientRect === 'function'
            ? hostEl.getBoundingClientRect()
            : null);
        if (rect && rect.height > 0) {
          segments.push({ text: txt, rect, el: hostEl });
        }
      }
      return;
    }
    if (node.nodeType === 1) {
      if (node.classList?.contains('floating-qr-box') || node.classList?.contains('floating-qr-stage')) return;
      const children = node.childNodes || [];
      for (let i = 0; i < children.length; i++) walk(children[i]);
    }
  };
  walk(rootEl);
  return segments;
}

function measureContainerSignatureGap(containerEl, signerTokens, stageRect, zoomScale, fallbackSizePx, anchorSlotRect = null) {
  const segments = collectVisualTextSegments(containerEl);
  const cRect = typeof containerEl?.getBoundingClientRect === 'function' ? containerEl.getBoundingClientRect() : anchorSlotRect;
  if (!cRect) return null;

  let titleSeg = null;
  let nameSeg = null;

  for (let i = 0; i < segments.length; i++) {
    const seg = segments[i];
    const isName = matchesSignerName(seg.text, signerTokens);
    const isNip = NIP_OR_NUPTK_RE.test(seg.text);
    const isRole = SIGNATURE_ROLE_RE.test(seg.text) && !SIGNATURE_BODY_PROSE_RE.test(seg.text);
    const isBoldNameLine = !isRole && !isNip && seg.el && /^(STRONG|B|U)$/i.test(seg.el.tagName || '');

    if (isName && !nameSeg) {
      nameSeg = seg;
    } else if (isNip && !nameSeg && i > 0) {
      nameSeg = segments[i - 1];
    } else if (!nameSeg && isBoldNameLine && titleSeg && seg.rect.top > titleSeg.rect.bottom + 8) {
      nameSeg = seg;
    }

    if ((isRole || /mengetahui/i.test(seg.text) || (seg.text.length <= 45 && /,\s*\d{1,2}\s+[a-z]+\s+\d{4}/i.test(seg.text))) && !nameSeg) {
      titleSeg = seg;
    }
  }

  const centerX = nameSeg
    ? (nameSeg.rect.left + nameSeg.rect.right) / 2
    : (titleSeg ? (titleSeg.rect.left + titleSeg.rect.right) / 2 : (cRect.left + cRect.width / 2));

  let gapTop = titleSeg ? titleSeg.rect.bottom : (anchorSlotRect ? anchorSlotRect.top : cRect.top);
  let gapBottom = nameSeg ? nameSeg.rect.top : cRect.bottom;
  if (anchorSlotRect && anchorSlotRect.height > 12 && anchorSlotRect.top >= gapTop - 8 && anchorSlotRect.bottom <= gapBottom + 8) {
    gapTop = anchorSlotRect.top;
    gapBottom = anchorSlotRect.bottom;
  }

  return computePlacementFromGap({
    gapTop,
    gapBottom,
    centerX,
    stageRect,
    zoomScale,
    fallbackSizePx,
    hasNameBelow: Boolean(nameSeg)
  });
}

function computePlacementFromGap({ gapTop, gapBottom, centerX, stageRect, zoomScale, fallbackSizePx, hasNameBelow = true }) {
  const safeZoom = Math.max(0.25, zoomScale || 1);
  const gapPx = Math.max(0, gapBottom - gapTop);
  const gapUnscaled = gapPx / safeZoom;

  let sizePx = fallbackSizePx;
  const captionMargin = 14;
  if (hasNameBelow && gapUnscaled >= 22 && gapUnscaled < fallbackSizePx + captionMargin) {
    sizePx = Math.round(clampNum(gapUnscaled - 10, 50, fallbackSizePx, fallbackSizePx));
  } else if (hasNameBelow && gapUnscaled > 0 && gapUnscaled < 22) {
    sizePx = Math.min(fallbackSizePx, 58);
  }

  const visualBoxW = sizePx * safeZoom;
  const visualBoxH = (sizePx + captionMargin) * safeZoom;
  const xLeftPx = (centerX - stageRect.left) - (visualBoxW / 2);

  let yTopPx;
  if (gapPx >= visualBoxH) {
    yTopPx = (gapTop - stageRect.top) + (gapPx - visualBoxH) / 2;
  } else if (hasNameBelow) {
    // Pusatkan simetris tepat di tengah celah antara baris jabatan & nama penandatangan
    const gapMid = (gapTop + gapBottom) / 2;
    const centeredTop = (gapMid - stageRect.top) - (visualBoxH / 2);
    yTopPx = gapPx >= 24 * safeZoom
      ? centeredTop
      : (gapBottom - stageRect.top) - visualBoxH + (4 * safeZoom);
  } else {
    yTopPx = (gapTop - stageRect.top) + Math.min(20 * safeZoom, Math.max(0, gapPx * 0.25));
  }

  const xPct = clampNum((xLeftPx / stageRect.width) * 100, 2, 85, 66);
  const yPct = clampNum((yTopPx / stageRect.height) * 100, 2, 85, 62);
  return { xPct, yPct, sizePx };
}

/**
 * Mengukur posisi area tanda tangan aktual di DOM (baik slot eksplisit maupun deteksi
 * otomatis nama Kepsek/TU, jabatan, dan blok TTD bebas di Word Pro) agar QR presisi di ruang TTD.
 */
export function calibratePlacementFromDom(stages = [], currentPlacement = {}, fallbackSizePx = 86, options = {}) {
  if (!Array.isArray(stages) || stages.length === 0) return currentPlacement;
  if (currentPlacement?.customPositioned) return currentPlacement;

  const school = options?.school || {};
  const signerTokens = buildKnownSignerTokens(school);
  const domCandidates = [];

  for (let idx = 0; idx < stages.length; idx++) {
    const stage = stages[idx];
    const sheet = stage?.closest('.smart-paper-sheet, .word-paper-sheet');
    if (!sheet) continue;

    const stageRect = typeof stage.getBoundingClientRect === 'function' ? stage.getBoundingClientRect() : null;
    if (!stageRect || stageRect.width <= 10 || stageRect.height <= 10) continue;
    const zoomScale = stageRect.width / Math.max(1, stage.offsetWidth || stageRect.width);

    // Tier 1: Slot eksplisit ([data-qr-slot="true"], .sig-space-holder, .smart-sig-cell)
    const explicitSlots = sheet.querySelectorAll('[data-qr-slot="true"], .sig-space-holder, .smart-sig-cell');
    for (const targetEl of explicitSlots) {
      const tRect = typeof targetEl.getBoundingClientRect === 'function' ? targetEl.getBoundingClientRect() : null;
      if (!tRect || tRect.width <= 0) continue;

      if (targetEl.classList.contains('sig-space-holder') && tRect.height > 12) {
        const sigBox = targetEl.closest('.word-sig-box');
        const nameEl = sigBox?.querySelector('.sig-name');
        const nameRect = nameEl ? (getTrimmedTextNodeRect(nameEl.firstChild, nameEl.ownerDocument) || nameEl.getBoundingClientRect()) : null;
        const centerX = nameRect && nameRect.width > 0 ? (nameRect.left + nameRect.right) / 2 : (tRect.left + tRect.width / 2);
        const fit = computePlacementFromGap({
          gapTop: tRect.top,
          gapBottom: tRect.bottom,
          centerX,
          stageRect,
          zoomScale,
          fallbackSizePx,
          hasNameBelow: Boolean(nameEl)
        });
        domCandidates.push({ pageIndex: idx, ...fit, score: 100 });
        continue;
      }

      const tdHost = targetEl.closest('td, .word-sig-box');
      if (tdHost) {
        const fit = measureContainerSignatureGap(tdHost, signerTokens, stageRect, zoomScale, fallbackSizePx, tRect.height > 12 ? tRect : null);
        if (fit) {
          domCandidates.push({ pageIndex: idx, ...fit, score: 100 });
          continue;
        }
      }

      // Slot [data-qr-slot="true"] di paragraf bebas (bukan di dalam tabel)
      if (targetEl.matches('[data-qr-slot="true"]')) {
        const pageContainer = targetEl.closest('.smart-editable-page, .sheet-content-editable');
        const siblings = pageContainer ? Array.from(pageContainer.children) : [];
        const slotIndex = siblings.indexOf(targetEl);
        if (slotIndex >= 0) {
          const slice = siblings.slice(Math.max(0, slotIndex - 4), Math.min(siblings.length, slotIndex + 7));
          let gapTop = tRect.top;
          let gapBottom = tRect.bottom;
          let nameCenterX = tRect.left + tRect.width / 2;
          let foundNameBelow = false;
          for (const sib of slice) {
            const txt = normalizeSigScanText(sib.textContent || '');
            if (!txt) continue;
            const sibRect = getTrimmedTextNodeRect(sib.firstChild, sib.ownerDocument) || sib.getBoundingClientRect();
            const relPos = siblings.indexOf(sib);
            if (relPos < slotIndex && sibRect.height > 0) {
              gapTop = Math.max(gapTop, sibRect.bottom);
            } else if (relPos > slotIndex && !foundNameBelow && (matchesSignerName(txt, signerTokens) || NIP_OR_NUPTK_RE.test(txt) || sib.querySelector('strong, u'))) {
              gapBottom = sibRect.top;
              if (sibRect.width > 0) nameCenterX = (sibRect.left + sibRect.right) / 2;
              foundNameBelow = true;
            }
          }
          if (foundNameBelow) {
            const fit = computePlacementFromGap({
              gapTop,
              gapBottom,
              centerX: nameCenterX,
              stageRect,
              zoomScale,
              fallbackSizePx,
              hasNameBelow: true
            });
            domCandidates.push({ pageIndex: idx, ...fit, score: 99 });
            continue;
          }
        }
      }

      const visualSize = fallbackSizePx * zoomScale;
      const xLeftPx = (tRect.left - stageRect.left) + Math.max(0, (tRect.width - visualSize) / 2);
      let yTopPx = (tRect.top - stageRect.top);
      if (targetEl.classList.contains('smart-sig-cell')) {
        yTopPx += Math.min(40 * zoomScale, tRect.height * 0.30);
      }
      domCandidates.push({
        pageIndex: idx,
        xPct: clampNum((xLeftPx / stageRect.width) * 100, 2, 85, 66),
        yPct: clampNum((yTopPx / stageRect.height) * 100, 2, 85, 62),
        sizePx: fallbackSizePx,
        score: 96
      });
    }

    // Tier 2: Deteksi otomatis sel <td> tabel tanda tangan bebas (tanpa class .smart-sig-cell)
    const tdCells = sheet.querySelectorAll('td');
    for (const td of tdCells) {
      if (td.classList.contains('smart-sig-cell') || td.closest('.floating-qr-box, .sheet-kop-container, .smart-meta-table, .smart-identity-table, [data-smart-role="identity-table"]')) continue;
      const tdText = normalizeSigScanText(td.textContent || '');
      if (!tdText || tdText.length > 260 || SIGNATURE_BODY_PROSE_RE.test(tdText) || KEY_VALUE_IDENTITY_LINE_RE.test(tdText)) continue;
      const rowParent = td.parentElement || td.parentNode;
      const rowCells = rowParent ? Array.from(rowParent.children || []) : [];
      const rowLeftLabel = rowCells[0] ? normalizeSigScanText(rowCells[0].textContent || '') : '';
      if (rowCells.length >= 2 && rowCells.length <= 3 && IDENTITY_LABEL_RE.test(rowLeftLabel)) continue;
      const tableText = normalizeSigScanText(td.closest('table')?.textContent || '');
      if (isNonSignatureTravelCell(tdText, tableText, signerTokens)) continue;

      const hasName = matchesSignerName(tdText, signerTokens);
      const hasRole = SIGNATURE_ROLE_RE.test(tdText);
      if (!hasName && !hasRole) continue;

      const fit = measureContainerSignatureGap(td, signerTokens, stageRect, zoomScale, fallbackSizePx);
      if (fit) {
        const isRightMost = rowCells.length >= 2 && rowCells.indexOf(td) === rowCells.length - 1;
        domCandidates.push({
          pageIndex: idx,
          ...fit,
          score: (hasName && hasRole ? 90 : (hasName ? 86 : 78)) + (isRightMost ? 4 : 0)
        });
      }
    }

    // Tier 3: Deteksi otomatis blok paragraf <p>/<div> tanda tangan bebas di Word Pro (tanpa tabel)
    const pageContainer = sheet.querySelector('.smart-editable-page, .sheet-content-editable');
    if (pageContainer) {
      const blocks = Array.from(pageContainer.children || []).filter(el =>
        !el.closest?.('table, .floating-qr-box, .sheet-kop-container')
      );
      for (let bIdx = 0; bIdx < blocks.length; bIdx++) {
        const el = blocks[bIdx];
        const txt = normalizeSigScanText(el.textContent || '');
        if (!txt || txt.length > 100 || SIGNATURE_BODY_PROSE_RE.test(txt) || SPPD_TRAVEL_EXCLUSION_RE.test(txt) || KEY_VALUE_IDENTITY_LINE_RE.test(txt)) continue;

        const hasName = matchesSignerName(txt, signerTokens);
        const hasRole = SIGNATURE_ROLE_RE.test(txt) && !/:$/.test(txt);
        if (!hasName && !hasRole) continue;

        // Cari pasangan jabatan (di atas) dan nama penandatangan (di bawah) dalam rentang +-6 paragraf
        let titleIdx = hasRole ? bIdx : -1;
        let nameIdx = hasName ? bIdx : -1;

        if (hasName && titleIdx < 0) {
          for (let j = bIdx - 1; j >= Math.max(0, bIdx - 6); j--) {
            const prevTxt = normalizeSigScanText(blocks[j].textContent || '');
            if (!prevTxt) continue;
            if (SIGNATURE_BODY_PROSE_RE.test(prevTxt) || prevTxt.length > 90) break;
            if (SIGNATURE_ROLE_RE.test(prevTxt) || /mengetahui/i.test(prevTxt)) {
              titleIdx = j;
              break;
            }
          }
        } else if (hasRole && nameIdx < 0) {
          for (let j = bIdx + 1; j <= Math.min(blocks.length - 1, bIdx + 6); j++) {
            const nextTxt = normalizeSigScanText(blocks[j].textContent || '');
            if (!nextTxt) continue;
            if (matchesSignerName(nextTxt, signerTokens) || NIP_OR_NUPTK_RE.test(nextTxt)) {
              nameIdx = NIP_OR_NUPTK_RE.test(nextTxt) && j > bIdx + 1 && normalizeSigScanText(blocks[j - 1].textContent || '') ? j - 1 : j;
              break;
            }
            if (blocks[j].querySelector('strong, u') && nextTxt.length <= 65) {
              nameIdx = j;
              break;
            }
          }
        }

        if (titleIdx < 0 && nameIdx < 0) continue;
        const anchorTitleEl = titleIdx >= 0 ? blocks[titleIdx] : null;
        const anchorNameEl = nameIdx >= 0 ? blocks[nameIdx] : null;
        const tRect = anchorTitleEl ? (collectVisualTextSegments(anchorTitleEl)[0]?.rect || anchorTitleEl.getBoundingClientRect()) : null;
        const nRect = anchorNameEl ? (collectVisualTextSegments(anchorNameEl)[0]?.rect || anchorNameEl.getBoundingClientRect()) : null;
        if (!tRect && !nRect) continue;

        const centerX = nRect && nRect.width > 0
          ? (nRect.left + nRect.right) / 2
          : (tRect.left + tRect.width / 2);
        const gapTop = tRect ? tRect.bottom : (nRect.top - 52 * zoomScale);
        const gapBottom = nRect ? nRect.top : (tRect.bottom + 56 * zoomScale);

        const fit = computePlacementFromGap({
          gapTop,
          gapBottom,
          centerX,
          stageRect,
          zoomScale,
          fallbackSizePx,
          hasNameBelow: Boolean(nRect)
        });
        const onRightHalf = centerX >= stageRect.left + stageRect.width * 0.45;
        const paired = Boolean(anchorTitleEl && anchorNameEl && nameIdx > titleIdx);
        domCandidates.push({
          pageIndex: idx,
          ...fit,
          score: (paired ? 88 : 74) + (onRightHalf ? 4 : 0)
        });
      }
    }
  }

  if (domCandidates.length === 0) {
    return currentPlacement;
  }

  domCandidates.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    return a.pageIndex - b.pageIndex;
  });
  const winner = domCandidates[0];

  return normalizeQrPlacement({
    ...currentPlacement,
    pageIndex: winner.pageIndex,
    xPct: winner.xPct,
    yPct: winner.yPct,
    sizePx: winner.sizePx,
    customPositioned: false
  }, stages.length);
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
    : findPreferredSignatureQrPlacement(docModel, stageEls.length, school);

  if (!rawPlacement?.customPositioned) {
    placement = calibratePlacementFromDom(stageEls, placement, placement.sizePx, { docModel, school });
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
