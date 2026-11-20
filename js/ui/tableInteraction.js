/**
 * tableInteraction.js - Kontrol Interaktif Tabel ala Microsoft Word Asli
 *
 * Fitur:
 * 1. Table Move Handle (ikon panah 4-arah di pojok kiri atas tabel):
 *    - Klik untuk memilih seluruh isi tabel (whole-table selection)
 *    - Seret (drag & drop) untuk memindahkan posisi tabel ke antar paragraf / blok / halaman lain
 * 2. Table Border Resize Cursor:
 *    - Otomatis menampilkan kursor `col-resize` (↔) saat diarahkan ke garis batas kolom
 *    - Otomatis menampilkan kursor `row-resize` (↕) saat diarahkan ke garis batas baris
 * 3. Table Column / Row Boundary Drag-and-Drop:
 *    - Menggeser batas kolom & batas tinggi baris secara presisi (memperhitungkan skala zoom kertas)
 *    - Menampilkan garis panduan & indikator ukuran real-time (cm & %), serta menyinkronkan ke <colgroup> / DocumentModel
 */

const PX_PER_CM = 37.7952755906;
const BORDER_HIT_PX = 5;

export function pxToCm(px = 0) {
  return Math.max(0.4, Math.round((Number(px) / PX_PER_CM) * 100) / 100);
}

export function cmToPx(cm = 0) {
  return Number(cm) * PX_PER_CM;
}

/**
 * Menghitung lebar kolom baru saat garis batas kolom digeser (dalam satuan px & %).
 * Internal divider membagi ulang lebar antara kolom `colIdx` dan `colIdx + 1` tanpa mengubah lebar total tabel.
 * Garis tepi paling kanan (`isOuterRight`) mengubah lebar keseluruhan tabel (clamped min..max).
 */
export function computeColumnResizeWidths(
  initWidthsPx = [],
  colIdx = 0,
  deltaPx = 0,
  { minColPx = 28, isOuterRight = false, minTotalPx = 120, maxTotalPx = 1200 } = {}
) {
  const widths = initWidthsPx.map(w => Math.max(minColPx, Number(w) || minColPx));
  if (widths.length === 0) return { widthsPx: [], widthsPct: [], totalWidthPx: 0 };

  if (isOuterRight || colIdx >= widths.length - 1) {
    const initTotal = widths.reduce((a, b) => a + b, 0);
    const targetTotal = Math.max(minTotalPx, Math.min(maxTotalPx, initTotal + deltaPx));
    const ratio = initTotal > 0 ? targetTotal / initTotal : 1;
    const nextPx = widths.map(w => Math.max(minColPx, Math.round(w * ratio * 10) / 10));
    const nextSum = nextPx.reduce((a, b) => a + b, 0);
    const nextPct = nextPx.map(w => Math.round((w / nextSum) * 1000) / 10);
    return {
      widthsPx: nextPx,
      widthsPct: normalizePctSum(nextPct),
      totalWidthPx: Math.round(nextSum * 10) / 10
    };
  }

  const pairTotal = widths[colIdx] + widths[colIdx + 1];
  const minBound = Math.min(minColPx, pairTotal / 2);
  const nextLeft = Math.max(minBound, Math.min(pairTotal - minBound, widths[colIdx] + deltaPx));
  const nextRight = pairTotal - nextLeft;
  const nextPx = [...widths];
  nextPx[colIdx] = Math.round(nextLeft * 10) / 10;
  nextPx[colIdx + 1] = Math.round(nextRight * 10) / 10;

  const totalSum = nextPx.reduce((a, b) => a + b, 0);
  const rawPct = nextPx.map(w => (totalSum > 0 ? Math.round((w / totalSum) * 1000) / 10 : 0));
  return {
    widthsPx: nextPx,
    widthsPct: normalizePctSum(rawPct),
    totalWidthPx: Math.round(totalSum * 10) / 10
  };
}

function normalizePctSum(pctList = []) {
  if (!pctList.length) return [];
  const rounded = pctList.map(p => Math.round(p * 10) / 10);
  const diff = Math.round((100 - rounded.reduce((a, b) => a + b, 0)) * 10) / 10;
  if (Math.abs(diff) > 0 && Math.abs(diff) <= 1.0) {
    const lastIdx = rounded.length - 1;
    rounded[lastIdx] = Math.max(2, Math.round((rounded[lastIdx] + diff) * 10) / 10);
  }
  return rounded;
}

/**
 * Membangun peta sel logis 2D dari DOM <table> dengan memperhitungkan colspan & rowspan,
 * serta mengecualikan kolom tombol aksi mini (.id-row-action-cell).
 */
function buildDomTableGrid(tableEl) {
  const trList = [...tableEl.querySelectorAll('tr')];
  const occ = [];
  const cellMap = new Map();
  let totalCols = 0;

  trList.forEach((tr, r) => {
    occ[r] = occ[r] || [];
    let c = 0;
    const cells = [...tr.children].filter(
      el => (el.tagName === 'TD' || el.tagName === 'TH') && !el.classList.contains('id-row-action-cell')
    );
    for (const cell of cells) {
      while (occ[r][c]) c++;
      const cs = Math.max(1, parseInt(cell.getAttribute('colspan') || '1', 10) || 1);
      const rs = Math.max(1, parseInt(cell.getAttribute('rowspan') || '1', 10) || 1);
      const info = { cell, tr, startR: r, startC: c, colspan: cs, rowspan: rs };
      cellMap.set(cell, info);
      for (let dr = 0; dr < rs; dr++) {
        occ[r + dr] = occ[r + dr] || [];
        for (let dc = 0; dc < cs; dc++) {
          occ[r + dr][c + dc] = info;
        }
      }
      c += cs;
      totalCols = Math.max(totalCols, c);
    }
  });

  return { trList, occ, totalCols, cellMap };
}

function getElementScale(el) {
  if (!el) return 1;
  const w = el.offsetWidth || 0;
  const rW = el.getBoundingClientRect().width || 0;
  if (w > 10 && rW > 10) return rW / w;
  return 1;
}

function measureLogicalColumnWidthsPx(tableEl, grid, scale) {
  const { totalCols, occ } = grid;
  const cols = [...tableEl.querySelectorAll(':scope > colgroup > col')].filter(
    col => !col.classList.contains('id-action-col')
  );
  if (cols.length === totalCols && totalCols > 0) {
    const measured = cols.map(c => c.getBoundingClientRect().width / scale);
    if (measured.every(w => w > 1)) return measured;
  }

  const widths = new Array(totalCols).fill(0);
  // Prioritaskan sel dengan colspan === 1 untuk mengukur lebar kolom secara langsung
  for (let c = 0; c < totalCols; c++) {
    for (let r = 0; r < occ.length; r++) {
      const slot = occ[r]?.[c];
      if (slot && slot.colspan === 1) {
        widths[c] = slot.cell.getBoundingClientRect().width / scale;
        break;
      }
    }
  }
  const tblW = tableEl.getBoundingClientRect().width / scale;
  const actionCell = tableEl.querySelector('.id-row-action-cell');
  const actionW = actionCell ? actionCell.getBoundingClientRect().width / scale : 0;
  const printableTblW = Math.max(80, tblW - actionW);
  const knownSum = widths.reduce((a, b) => a + b, 0);
  const missingCount = widths.filter(w => w <= 1).length;
  if (missingCount > 0) {
    const fallbackW = Math.max(28, (printableTblW - knownSum) / missingCount);
    for (let c = 0; c < totalCols; c++) {
      if (widths[c] <= 1) widths[c] = fallbackW;
    }
  }
  return widths;
}

function applyColumnWidthsToTable(tableEl, widthsPct, totalTableWidthPct = null) {
  if (!tableEl || !Array.isArray(widthsPct) || widthsPct.length === 0) return;

  const isLegacyIdentity = tableEl.classList.contains('identity-table-word');
  const hasActionCol = Boolean(tableEl.querySelector('.id-row-action-cell'));

  if (totalTableWidthPct !== null) {
    tableEl.style.width = `${totalTableWidthPct}%`;
    tableEl.setAttribute('data-table-width-pct', String(totalTableWidthPct));
  }

  if (isLegacyIdentity && widthsPct.length >= 2) {
    tableEl.setAttribute('data-col-widths', `${widthsPct[0]},${widthsPct[1]}`);
    tableEl.querySelectorAll('tr').forEach((tr) => {
      const lbl = tr.querySelector('.id-label');
      const val = tr.querySelector('.id-val');
      if (lbl) lbl.style.width = `${widthsPct[0]}%`;
      if (val) val.style.width = `${widthsPct[1]}%`;
    });
    return;
  }

  tableEl.setAttribute('data-col-widths', widthsPct.join(','));
  let colgroup = [...tableEl.children].find(ch => ch.tagName === 'COLGROUP');
  if (!colgroup) {
    colgroup = document.createElement('colgroup');
    tableEl.insertBefore(colgroup, tableEl.firstChild);
  }
  const colsHtml = widthsPct.map(p => `<col style="width: ${p}%;">`).join('');
  colgroup.innerHTML = hasActionCol
    ? `${colsHtml}<col class="id-action-col" style="width: 38px;">`
    : colsHtml;
}

function detectBorderHit(tableEl, clientX, clientY) {
  if (!tableEl) return null;
  const grid = buildDomTableGrid(tableEl);
  const { totalCols, cellMap, trList } = grid;
  if (totalCols === 0 || trList.length === 0) return null;

  let bestCol = null;
  let bestRow = null;

  for (const [, info] of cellMap.entries()) {
    const rect = info.cell.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) continue;

    // 1. Batas kolom (kiri bila col > 0, atau kanan sel)
    if (clientY >= rect.top - BORDER_HIT_PX && clientY <= rect.bottom + BORDER_HIT_PX) {
      const distRight = Math.abs(clientX - rect.right);
      if (distRight <= BORDER_HIT_PX) {
        const colIdx = info.startC + info.colspan - 1;
        const isOuterRight = colIdx >= totalCols - 1;
        if (!bestCol || distRight < bestCol.dist) {
          bestCol = {
            axis: 'col',
            dist: distRight,
            table: tableEl,
            colIdx,
            isOuterRight,
            guideScreenX: rect.right,
            grid
          };
        }
      }
      if (info.startC > 0) {
        const distLeft = Math.abs(clientX - rect.left);
        if (distLeft <= BORDER_HIT_PX) {
          const colIdx = info.startC - 1;
          if (!bestCol || distLeft < bestCol.dist) {
            bestCol = {
              axis: 'col',
              dist: distLeft,
              table: tableEl,
              colIdx,
              isOuterRight: false,
              guideScreenX: rect.left,
              grid
            };
          }
        }
      }
    }

    // 2. Batas baris (atas bila row > 0, atau bawah sel)
    if (clientX >= rect.left - BORDER_HIT_PX && clientX <= rect.right + BORDER_HIT_PX) {
      const distBottom = Math.abs(clientY - rect.bottom);
      if (distBottom <= BORDER_HIT_PX) {
        const rowIdx = info.startR + info.rowspan - 1;
        const targetTr = trList[rowIdx];
        if (targetTr && (!bestRow || distBottom < bestRow.dist)) {
          bestRow = {
            axis: 'row',
            dist: distBottom,
            table: tableEl,
            rowIdx,
            tr: targetTr,
            guideScreenY: rect.bottom,
            grid
          };
        }
      }
      if (info.startR > 0) {
        const distTop = Math.abs(clientY - rect.top);
        if (distTop <= BORDER_HIT_PX) {
          const rowIdx = info.startR - 1;
          const targetTr = trList[rowIdx];
          if (targetTr && (!bestRow || distTop < bestRow.dist)) {
            bestRow = {
              axis: 'row',
              dist: distTop,
              table: tableEl,
              rowIdx,
              tr: targetTr,
              guideScreenY: rect.top,
              grid
            };
          }
        }
      }
    }
  }

  // Prioritaskan batas kolom jika jarak hampir sama agar memisahkan kolom dengan mulus
  if (bestCol && bestRow) {
    return bestCol.dist <= bestRow.dist + 1 ? bestCol : bestRow;
  }
  return bestCol || bestRow || null;
}

/**
 * Memasang kontrol interaktif tabel (Move Handle, Resize Cursor, Drag Boundary)
 * ke sebuah desk editor (`word-desk`). Bersih dari DOM contenteditable & bebas kebocoran listener.
 */
export function attachInteractiveTables({
  deskEl,
  tableSelector = 'table',
  excludeTableSelector = '.smart-sig-table, .meta-perihal-table',
  getMoveUnitEl = (tableEl) => tableEl,
  getDropTargets = () => [],
  performMove = null,
  onSelectWholeTable = null,
  onBeforeMutate = null,
  onMutate = null
} = {}) {
  if (!deskEl || typeof document === 'undefined') {
    return { destroy: () => {}, clearSelection: () => {} };
  }

  // Pastikan desk menjadi anchor posisi relatif untuk overlay handles & guide lines
  if (getComputedStyle(deskEl).position === 'static') {
    deskEl.style.position = 'relative';
  }

  const overlayRoot = document.createElement('div');
  overlayRoot.className = 'word-table-overlay-root no-print';
  overlayRoot.innerHTML = `
    <button type="button" class="word-table-move-handle hidden" title="Klik: pilih seluruh tabel • Seret: pindahkan posisi tabel" aria-label="Pilih atau pindahkan tabel">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
        <polyline points="5 9 2 12 5 15"></polyline>
        <polyline points="9 5 12 2 15 5"></polyline>
        <polyline points="15 19 12 22 9 19"></polyline>
        <polyline points="19 9 22 12 19 15"></polyline>
        <line x1="2" y1="12" x2="22" y2="12"></line>
        <line x1="12" y1="2" x2="12" y2="22"></line>
      </svg>
    </button>
    <div class="word-table-border-affordance hidden"></div>
    <div class="word-table-resize-guide-v hidden"></div>
    <div class="word-table-resize-guide-h hidden"></div>
    <div class="word-table-drop-line hidden"></div>
    <div class="word-table-size-tooltip hidden"></div>
  `;
  deskEl.appendChild(overlayRoot);

  const moveHandleEl = overlayRoot.querySelector('.word-table-move-handle');
  const affordanceEl = overlayRoot.querySelector('.word-table-border-affordance');
  const guideVEl = overlayRoot.querySelector('.word-table-resize-guide-v');
  const guideHEl = overlayRoot.querySelector('.word-table-resize-guide-h');
  const dropLineEl = overlayRoot.querySelector('.word-table-drop-line');
  const tooltipEl = overlayRoot.querySelector('.word-table-size-tooltip');

  let hoveredTable = null;
  let activeSelectedTable = null;
  let armedHit = null;
  let dragState = null;
  let moveHideTimer = null;

  function isEligibleTable(tableEl) {
    if (!tableEl || !deskEl.contains(tableEl)) return false;
    if (excludeTableSelector && tableEl.matches(excludeTableSelector)) return false;
    return tableEl.matches(tableSelector);
  }

  function toDeskCoords(clientX, clientY) {
    const r = deskEl.getBoundingClientRect();
    return {
      x: clientX - r.left + deskEl.scrollLeft,
      y: clientY - r.top + deskEl.scrollTop
    };
  }

  function positionMoveHandleForTable(tableEl) {
    if (!tableEl || !isEligibleTable(tableEl)) {
      moveHandleEl.classList.add('hidden');
      hoveredTable = null;
      return;
    }
    clearTimeout(moveHideTimer);
    hoveredTable = tableEl;
    const tRect = tableEl.getBoundingClientRect();
    const pos = toDeskCoords(tRect.left, tRect.top);
    moveHandleEl.style.left = `${Math.max(4, pos.x - 18)}px`;
    moveHandleEl.style.top = `${Math.max(4, pos.y - 18)}px`;
    moveHandleEl.classList.remove('hidden');
  }

  function scheduleHideMoveHandle() {
    clearTimeout(moveHideTimer);
    moveHideTimer = setTimeout(() => {
      if (!dragState && hoveredTable && hoveredTable !== activeSelectedTable) {
        moveHandleEl.classList.add('hidden');
        hoveredTable = null;
      }
    }, 320);
  }

  function clearWholeTableSelection() {
    if (activeSelectedTable) {
      activeSelectedTable.classList.remove('word-table-whole-selected');
      activeSelectedTable = null;
    }
  }

  function selectWholeTable(tableEl) {
    clearWholeTableSelection();
    if (!tableEl || !isEligibleTable(tableEl)) return;
    activeSelectedTable = tableEl;
    tableEl.classList.add('word-table-whole-selected');

    try {
      const sel = window.getSelection();
      if (sel) {
        const range = document.createRange();
        range.selectNode(tableEl);
        sel.removeAllRanges();
        sel.addRange(range);
      }
    } catch {}

    positionMoveHandleForTable(tableEl);
    onSelectWholeTable?.(tableEl);
  }

  function hideResizeFeedback() {
    affordanceEl.classList.add('hidden');
    guideVEl.classList.add('hidden');
    guideHEl.classList.add('hidden');
    dropLineEl.classList.add('hidden');
    tooltipEl.classList.add('hidden');
  }

  function showBorderAffordance(hit) {
    if (!hit) {
      affordanceEl.classList.add('hidden');
      return;
    }
    const tRect = hit.table.getBoundingClientRect();
    if (hit.axis === 'col') {
      const p = toDeskCoords(hit.guideScreenX, tRect.top);
      affordanceEl.className = 'word-table-border-affordance is-col';
      affordanceEl.style.left = `${p.x - 1}px`;
      affordanceEl.style.top = `${p.y}px`;
      affordanceEl.style.width = '3px';
      affordanceEl.style.height = `${Math.max(20, tRect.height)}px`;
    } else {
      const p = toDeskCoords(tRect.left, hit.guideScreenY);
      affordanceEl.className = 'word-table-border-affordance is-row';
      affordanceEl.style.left = `${p.x}px`;
      affordanceEl.style.top = `${p.y - 1}px`;
      affordanceEl.style.width = `${Math.max(40, tRect.width)}px`;
      affordanceEl.style.height = '3px';
    }
  }

  function setCursorVisual(axis) {
    if (axis === 'col') {
      deskEl.style.cursor = 'col-resize';
      document.body.classList.add('cursor-col-resize');
      document.body.classList.remove('cursor-row-resize');
    } else if (axis === 'row') {
      deskEl.style.cursor = 'row-resize';
      document.body.classList.add('cursor-row-resize');
      document.body.classList.remove('cursor-col-resize');
    } else {
      deskEl.style.cursor = '';
      document.body.classList.remove('cursor-col-resize', 'cursor-row-resize');
    }
  }

  function findTableNearPointer(targetEl, clientX, clientY) {
    const directTable = targetEl?.closest?.('table');
    if (isEligibleTable(directTable)) return directTable;

    // Izinkan deteksi garis batas luar tabel hingga ±6px di luar tepi tabel
    const allTables = [...deskEl.querySelectorAll(tableSelector)].filter(isEligibleTable);
    for (const tbl of allTables) {
      const r = tbl.getBoundingClientRect();
      if (
        clientX >= r.left - 6 && clientX <= r.right + 6 &&
        clientY >= r.top - 6 && clientY <= r.bottom + 6
      ) {
        return tbl;
      }
    }
    return null;
  }

  function findBestDropSlot(clientY, draggedUnit) {
    const targets = getDropTargets(draggedUnit) || [];
    let best = null;
    for (const el of targets) {
      if (!el || el === draggedUnit || draggedUnit.contains(el)) continue;
      const r = el.getBoundingClientRect();
      if (r.height <= 0 && r.width <= 0) continue;
      const midY = r.top + r.height / 2;
      const position = clientY < midY ? 'before' : 'after';
      const dist = Math.abs(clientY - (position === 'before' ? r.top : r.bottom));
      if (!best || dist < best.dist) {
        best = {
          targetEl: el,
          position,
          dist,
          lineScreenY: position === 'before' ? r.top - 2 : r.bottom + 2,
          lineLeft: r.left,
          lineWidth: Math.max(120, r.width)
        };
      }
    }
    return best;
  }

  const isFloatingQrTarget = (target) =>
    Boolean(target?.closest?.('.floating-qr-box, .floating-qr-toolbar, .floating-qr-resize, .floating-qr-stage, .fqr-context-menu'));

  const onDeskMouseMove = (e) => {
    if (dragState) return;
    if (isFloatingQrTarget(e.target)) {
      armedHit = null;
      setCursorVisual(null);
      return;
    }
    if (e.target.closest?.('.word-table-move-handle')) {
      clearTimeout(moveHideTimer);
      setCursorVisual(null);
      return;
    }

    const nearTbl = findTableNearPointer(e.target, e.clientX, e.clientY);
    if (nearTbl) {
      positionMoveHandleForTable(nearTbl);
      const hit = detectBorderHit(nearTbl, e.clientX, e.clientY);
      armedHit = hit;
      setCursorVisual(hit?.axis || null);
      showBorderAffordance(hit);
    } else {
      armedHit = null;
      setCursorVisual(null);
      affordanceEl.classList.add('hidden');
      scheduleHideMoveHandle();
    }
  };

  const onDeskMouseDown = (e) => {
    if (e.button !== 0) return;
    if (isFloatingQrTarget(e.target)) {
      armedHit = null;
      setCursorVisual(null);
      return;
    }

    // Bersihkan pilihan seluruh tabel bila klik biasa di dalam sel untuk mengetik
    if (!e.target.closest?.('.word-table-move-handle')) {
      clearWholeTableSelection();
    }

    if (!armedHit || !isEligibleTable(armedHit.table)) return;
    e.preventDefault();
    e.stopPropagation();

    const tbl = armedHit.table;
    const scale = getElementScale(tbl);
    onBeforeMutate?.();

    if (armedHit.axis === 'col') {
      const initWidthsPx = measureLogicalColumnWidthsPx(tbl, armedHit.grid, scale);
      const initSumPx = initWidthsPx.reduce((a, b) => a + b, 0);
      const parentBox = tbl.parentElement;
      const maxParentPx = parentBox
        ? (parentBox.getBoundingClientRect().width / scale)
        : initSumPx;

      dragState = {
        mode: 'col',
        table: tbl,
        colIdx: armedHit.colIdx,
        isOuterRight: armedHit.isOuterRight,
        scale,
        startX: e.clientX,
        startY: e.clientY,
        initWidthsPx,
        maxParentPx: Math.max(initSumPx, maxParentPx),
        latestWidthsPct: null,
        latestTableWidthPct: null
      };
    } else {
      const tr = armedHit.tr;
      const initRowH = tr.getBoundingClientRect().height / scale;
      dragState = {
        mode: 'row',
        table: tbl,
        tr,
        rowIdx: armedHit.rowIdx,
        scale,
        startX: e.clientX,
        startY: e.clientY,
        initRowH
      };
    }

    document.body.classList.add('word-table-dragging');
    attachDocDragListeners();
  };

  const onMoveHandleMouseDown = (e) => {
    if (e.button !== 0 || !hoveredTable) return;
    e.preventDefault();
    e.stopPropagation();
    const tbl = hoveredTable;
    const unitEl = getMoveUnitEl(tbl);
    dragState = {
      mode: 'move',
      table: tbl,
      unitEl,
      startX: e.clientX,
      startY: e.clientY,
      hasMoved: false,
      dropTarget: null
    };
    document.body.classList.add('word-table-dragging');
    attachDocDragListeners();
  };

  const onDocMouseMove = (e) => {
    if (!dragState) return;

    if (dragState.mode === 'col') {
      const dxScreen = e.clientX - dragState.startX;
      const dx = dxScreen / dragState.scale;
      const res = computeColumnResizeWidths(dragState.initWidthsPx, dragState.colIdx, dx, {
        minColPx: 28,
        isOuterRight: dragState.isOuterRight,
        minTotalPx: 140,
        maxTotalPx: dragState.maxParentPx
      });

      const tableWidthPct = dragState.isOuterRight && dragState.maxParentPx > 0
        ? Math.min(100, Math.max(35, Math.round((res.totalWidthPx / dragState.maxParentPx) * 1000) / 10))
        : null;

      dragState.latestWidthsPct = res.widthsPct;
      dragState.latestTableWidthPct = tableWidthPct;
      applyColumnWidthsToTable(dragState.table, res.widthsPct, tableWidthPct);

      // Update panduan garis vertikal & tooltip ukuran
      const tRect = dragState.table.getBoundingClientRect();
      const cellsInTable = [...dragState.table.querySelectorAll('td, th')].filter(
        el => !el.classList.contains('id-row-action-cell')
      );
      let lineScreenX = e.clientX;
      const gridNow = buildDomTableGrid(dragState.table);
      for (const cell of cellsInTable) {
        const info = gridNow.cellMap.get(cell);
        if (info && info.startC + info.colspan - 1 === dragState.colIdx) {
          lineScreenX = cell.getBoundingClientRect().right;
          break;
        }
      }

      const gp = toDeskCoords(lineScreenX, tRect.top);
      guideVEl.style.left = `${gp.x}px`;
      guideVEl.style.top = `${gp.y}px`;
      guideVEl.style.height = `${Math.max(24, tRect.height)}px`;
      guideVEl.classList.remove('hidden');

      const tp = toDeskCoords(lineScreenX + 10, e.clientY - 26);
      tooltipEl.style.left = `${tp.x}px`;
      tooltipEl.style.top = `${tp.y}px`;
      if (dragState.isOuterRight) {
        tooltipEl.textContent = `Lebar Tabel: ${pxToCm(res.totalWidthPx).toFixed(2)} cm (${tableWidthPct || 100}%)`;
      } else {
        const c1 = dragState.colIdx;
        const c2 = c1 + 1;
        tooltipEl.textContent = `Kol ${c1 + 1}: ${pxToCm(res.widthsPx[c1]).toFixed(1)} cm (${res.widthsPct[c1]}%)  |  Kol ${c2 + 1}: ${pxToCm(res.widthsPx[c2]).toFixed(1)} cm (${res.widthsPct[c2]}%)`;
      }
      tooltipEl.classList.remove('hidden');
      positionMoveHandleForTable(dragState.table);
      return;
    }

    if (dragState.mode === 'row') {
      const dyScreen = e.clientY - dragState.startY;
      const dy = dyScreen / dragState.scale;
      const nextPx = Math.max(22, Math.round(dragState.initRowH + dy));
      const cm = pxToCm(nextPx);

      dragState.tr.style.height = `${cm}cm`;
      dragState.tr.setAttribute('data-min-height', String(cm));

      const tRect = dragState.table.getBoundingClientRect();
      const rRect = dragState.tr.getBoundingClientRect();
      const gp = toDeskCoords(tRect.left, rRect.bottom);
      guideHEl.style.left = `${gp.x}px`;
      guideHEl.style.top = `${gp.y}px`;
      guideHEl.style.width = `${Math.max(40, tRect.width)}px`;
      guideHEl.classList.remove('hidden');

      const tp = toDeskCoords(e.clientX + 12, rRect.bottom + 6);
      tooltipEl.style.left = `${tp.x}px`;
      tooltipEl.style.top = `${tp.y}px`;
      tooltipEl.textContent = `Tinggi Baris ${dragState.rowIdx + 1}: ${cm.toFixed(2)} cm`;
      tooltipEl.classList.remove('hidden');
      return;
    }

    if (dragState.mode === 'move') {
      const dist = Math.hypot(e.clientX - dragState.startX, e.clientY - dragState.startY);
      if (dist > 5) {
        dragState.hasMoved = true;
      }
      if (!dragState.hasMoved) return;

      dragState.table.classList.add('word-table-dragging-source');
      const slot = findBestDropSlot(e.clientY, dragState.unitEl);
      dragState.dropTarget = slot;
      if (slot) {
        const dp = toDeskCoords(slot.lineLeft, slot.lineScreenY);
        dropLineEl.style.left = `${dp.x}px`;
        dropLineEl.style.top = `${dp.y}px`;
        dropLineEl.style.width = `${slot.lineWidth}px`;
        dropLineEl.classList.remove('hidden');

        const tp = toDeskCoords(e.clientX + 14, e.clientY - 12);
        tooltipEl.style.left = `${tp.x}px`;
        tooltipEl.style.top = `${tp.y}px`;
        tooltipEl.textContent = 'Lepaskan untuk memindahkan tabel ke sini';
        tooltipEl.classList.remove('hidden');
      } else {
        dropLineEl.classList.add('hidden');
      }
    }
  };

  const onDocMouseUp = () => {
    if (!dragState) return;
    const finished = dragState;
    dragState = null;

    document.body.classList.remove('word-table-dragging');
    detachDocDragListeners();
    hideResizeFeedback();
    setCursorVisual(null);
    finished.table?.classList.remove('word-table-dragging-source');

    if (finished.mode === 'move') {
      if (!finished.hasMoved) {
        // Klik tunggal pada handle => pilih seluruh isi tabel
        selectWholeTable(finished.table);
        return;
      }
      if (finished.dropTarget) {
        onBeforeMutate?.();
        const { targetEl, position } = finished.dropTarget;
        if (typeof performMove === 'function') {
          performMove(finished.unitEl, targetEl, position);
        } else {
          targetEl.insertAdjacentElement(position === 'before' ? 'beforebegin' : 'afterend', finished.unitEl);
        }
        clearWholeTableSelection();
        positionMoveHandleForTable(finished.table);
        onMutate?.('move', finished.table);
      }
      return;
    }

    if (finished.mode === 'col' || finished.mode === 'row') {
      positionMoveHandleForTable(finished.table);
      onMutate?.(finished.mode === 'col' ? 'resize-col' : 'resize-row', finished.table);
    }
  };

  function attachDocDragListeners() {
    document.addEventListener('mousemove', onDocMouseMove, true);
    document.addEventListener('mouseup', onDocMouseUp, true);
  }

  function detachDocDragListeners() {
    document.removeEventListener('mousemove', onDocMouseMove, true);
    document.removeEventListener('mouseup', onDocMouseUp, true);
  }

  const onDeskScroll = () => {
    if (activeSelectedTable && deskEl.contains(activeSelectedTable)) {
      positionMoveHandleForTable(activeSelectedTable);
    } else if (hoveredTable && deskEl.contains(hoveredTable)) {
      positionMoveHandleForTable(hoveredTable);
    }
  };

  const onDocKeyDown = (e) => {
    if (!deskEl.isConnected) {
      destroy();
      return;
    }
    if (!activeSelectedTable || !deskEl.contains(activeSelectedTable)) return;
    if (e.key === 'Escape') {
      clearWholeTableSelection();
      return;
    }
    if ((e.key === 'Delete' || e.key === 'Backspace') && activeSelectedTable) {
      const sel = window.getSelection();
      const isWholeNodeSelected = sel && !sel.isCollapsed && activeSelectedTable.contains(sel.anchorNode) && sel.anchorNode !== activeSelectedTable
        ? false
        : true;
      if (isWholeNodeSelected || sel?.toString()?.trim() === activeSelectedTable.innerText.trim()) {
        e.preventDefault();
        onBeforeMutate?.();
        const unitEl = getMoveUnitEl(activeSelectedTable);
        unitEl?.remove();
        clearWholeTableSelection();
        moveHandleEl.classList.add('hidden');
        onMutate?.('delete-table');
      }
    }
  };

  deskEl.addEventListener('mousemove', onDeskMouseMove);
  deskEl.addEventListener('mousedown', onDeskMouseDown, true);
  deskEl.addEventListener('scroll', onDeskScroll, { passive: true });
  moveHandleEl.addEventListener('mousedown', onMoveHandleMouseDown);
  document.addEventListener('keydown', onDocKeyDown);

  function destroy() {
    clearTimeout(moveHideTimer);
    detachDocDragListeners();
    setCursorVisual(null);
    document.body.classList.remove('word-table-dragging');
    deskEl.removeEventListener('mousemove', onDeskMouseMove);
    deskEl.removeEventListener('mousedown', onDeskMouseDown, true);
    deskEl.removeEventListener('scroll', onDeskScroll);
    moveHandleEl.removeEventListener('mousedown', onMoveHandleMouseDown);
    document.removeEventListener('keydown', onDocKeyDown);
    overlayRoot.remove();
  }

  return {
    destroy,
    selectWholeTable,
    clearSelection: clearWholeTableSelection,
    refreshHandle: () => {
      if (activeSelectedTable && deskEl.contains(activeSelectedTable)) {
        positionMoveHandleForTable(activeSelectedTable);
      } else {
        moveHandleEl.classList.add('hidden');
      }
    }
  };
}
