/**
 * surgicalEdit.js - Mesin Pengeditan Bedah (Surgical Patch-In-Place) untuk AI Copilot
 *
 * Tujuan utama:
 * Saat pengguna menyuruh AI mengubah DATA/TEKS (nama pegawai, tujuan, tanggal, nomor, PUEBI,
 * tambah baris pegawai, dll.) tanpa meminta ubah tata letak, AI HANYA menambal sel/paragraf
 * target di tempatnya. Seluruh tabel (SPPD, grid, 2-kolom), lebar kolom (<colgroup>), tinggi baris,
 * colspan/rowspan, borderStyle, pemisah halaman, dan slot QR tetap terkunci 100% utuh.
 */

import {
  parseHtmlNodes,
  serializeHtmlNodes,
  getNodePlainText,
  cloneAstNode
} from '../utils/htmlAst.js';
import { formatSmartNomorSurat } from '../letters/numbering.js';
import { formatSignerNipLine } from '../letters/verification.js';

const INLINE_FMT_TAGS = new Set(['strong', 'b', 'em', 'i', 'u', 'span', 'sup', 'sub', 'small']);
const WRAPPER_CONTAINER_TAGS = new Set(['div', 'blockquote', 'section']);
const BLOCK_CONTAINER_TAGS = new Set(['table', 'p', 'div', 'h1', 'h2', 'h3', 'h4', 'ul', 'ol', 'blockquote']);

/**
 * Menentukan apakah instruksi pengguna secara eksplisit meminta perombakan tata letak /
 * pembuatan ulang tabel / perubahan struktur halaman.
 * Secara default bernilai `false` (mode kunci tata letak & tabel).
 */
export function shouldAllowLayoutRestructure(instruction = '') {
  const txt = String(instruction || '').toLowerCase().trim();
  if (!txt) return false;

  const explicitLayoutPatterns = [
    /\b(ubah|ganti|atur|desain|susun|buat|bikin)\s+(ulang\s+)?(tata\s*letak|layout|struktur\s+dokumen|desain\s+surat|tampilan)\b/,
    /\b(rombak|desain\s+ulang|susun\s+ulang|tata\s+ulang|buat\s+dari\s+(awal|nol)|format\s+ulang\s+total)\b/,
    /\b(pindah|geser|tukar)\w*\s+(posisi\s+)?tabel\b/,
    /\b(buat|bikin|sisipkan|pecah|gabung|ubah)\w*\s+(tabel\s+baru|struktur\s+tabel|kolom\s+tabel|jadi\s+\d+\s+kolom)\b/,
    /\b(gabung|satukan)\s+(semua\s+)?halaman\b|\bhapus\s+halaman\b/
  ];

  return explicitLayoutPatterns.some((re) => re.test(txt));
}

function normalizeStr(v) {
  return String(v ?? '').replace(/\u00a0/g, ' ');
}

function hasNestedBlockChild(node) {
  return (node?.children || []).some(
    (ch) => ch?.type === 'element' && (BLOCK_CONTAINER_TAGS.has(ch.tag) || hasNestedBlockChild(ch))
  );
}

function textToInlineAst(text = '') {
  const clean = normalizeStr(text);
  const parts = clean.split(/\r?\n/);
  const nodes = [];
  parts.forEach((part, idx) => {
    if (part) nodes.push({ type: 'text', text: part });
    if (idx < parts.length - 1) {
      nodes.push({ type: 'element', tag: 'br', attrs: {}, children: [] });
    }
  });
  return nodes;
}

function isWhitespaceTextNode(node) {
  return node?.type === 'text' && !String(node.text || '').trim();
}

function findSingleFormattingChainLeaf(children = []) {
  let currentList = children;
  let innermost = null;

  for (let depth = 0; depth < 6; depth++) {
    const meaningful = (currentList || []).filter((c) => !isWhitespaceTextNode(c));
    if (meaningful.length !== 1) break;
    const only = meaningful[0];
    if (only.type === 'element' && INLINE_FMT_TAGS.has(only.tag)) {
      innermost = only;
      currentList = only.children || [];
    } else {
      break;
    }
  }
  return innermost;
}

/**
 * Memperbarui teks suatu elemen inline/paragraf sambil mempertahankan pembungkus gaya tunggal
 * (misal `<strong><u>Nama</u></strong>`) atau pola awalan statis + tag (`Tiba di <u>Lokasi</u>`).
 */
export function updateInlineElementTextPreservingStyle(elNode, newText = '') {
  if (!elNode || elNode.type !== 'element') return;
  const cleanText = normalizeStr(newText);

  // 1. Cek pola "Prefix Statis + <inlineTag>Isi</inlineTag>" (seperti "Tiba di <u>Dragon Motion</u>")
  const meaningful = (elNode.children || []).filter((c) => !isWhitespaceTextNode(c));
  if (
    meaningful.length === 2 &&
    meaningful[0].type === 'text' &&
    meaningful[1].type === 'element' &&
    INLINE_FMT_TAGS.has(meaningful[1].tag)
  ) {
    const origPrefix = meaningful[0].text;
    const trimmedPrefix = origPrefix.trimEnd();
    if (trimmedPrefix && cleanText.toLowerCase().startsWith(trimmedPrefix.toLowerCase())) {
      const rest = cleanText.slice(trimmedPrefix.length).replace(/^\s+/, '');
      const leaf = findSingleFormattingChainLeaf([meaningful[1]]) || meaningful[1];
      leaf.children = textToInlineAst(rest);
      elNode.children = [
        { type: 'text', text: origPrefix.endsWith(' ') ? origPrefix : `${trimmedPrefix} ` },
        meaningful[1]
      ];
      return;
    }
  }

  // 2. Cek apakah seluruh teks dibungkus satu rantai tag inline tunggal (<strong><u>...</u></strong>)
  const fmtLeaf = findSingleFormattingChainLeaf(elNode.children);
  if (fmtLeaf) {
    fmtLeaf.children = cleanText ? textToInlineAst(cleanText) : [{ type: 'text', text: '' }];
    return;
  }

  // 3. Fallback teks biasa (pertahankan atribut & tag induk <p> / <td>)
  if (!cleanText && elNode.tag === 'p') {
    elNode.children = [{ type: 'element', tag: 'br', attrs: {}, children: [] }];
  } else {
    elNode.children = textToInlineAst(cleanText);
  }
}

export function isQrSlotElement(node) {
  if (!node || node.type !== 'element') return false;
  if (node.attrs?.['data-qr-slot'] === 'true') return true;
  return String(node.attrs?.class || '').split(/\s+/).includes('smart-qr-slot');
}

function isSpacerParagraphElement(node) {
  if (!node || node.type !== 'element' || node.tag !== 'p' || isQrSlotElement(node)) return false;
  const txt = getNodePlainText(node).replace(/\u00a0/g, '').trim();
  return txt === '';
}

function getBlockParagraphChildren(cellNode) {
  return (cellNode?.children || []).filter(
    (c) => c.type === 'element' && (c.tag === 'p' || c.tag === 'div')
  );
}

/**
 * Memperbarui isi sel `<td>`/`<th>` tanpa menyentuh atribut tabel/sel (colspan, rowspan, align, width)
 * dan tanpa merusak `.smart-qr-slot` maupun tag `<p>` bertingkat di dalamnya.
 */
export function updateTableCellNodePreservingLayout(cellNode, newText = '', paragraphIndex = undefined) {
  if (!cellNode || cellNode.type !== 'element') return false;
  const cleanText = normalizeStr(newText);
  const blockKids = getBlockParagraphChildren(cellNode);

  if (blockKids.length > 0) {
    if (typeof paragraphIndex === 'number' && Number.isInteger(paragraphIndex) && paragraphIndex >= 0) {
      const targetP = blockKids[paragraphIndex];
      if (!targetP || isQrSlotElement(targetP)) return false;
      updateInlineElementTextPreservingStyle(targetP, cleanText);
      return true;
    }

    const lines = cleanText
      .split(/\r?\n/)
      .filter((line) => !/\[QR_SLOT_TERKUNCI\]/i.test(line));
    const hasQrSlot = blockKids.some(isQrSlotElement);

    if (hasQrSlot) {
      let editablePs = blockKids.filter((p) => !isQrSlotElement(p) && !isSpacerParagraphElement(p));
      if (editablePs.length === 0) {
        editablePs = blockKids.filter((p) => !isQrSlotElement(p));
      }
      if (editablePs.length === 0) return false;

      if (lines.length <= editablePs.length) {
        editablePs.forEach((pEl, idx) => {
          updateInlineElementTextPreservingStyle(pEl, lines[idx] !== undefined ? lines[idx] : getNodePlainText(pEl));
        });
      } else {
        editablePs.forEach((pEl, idx) => {
          if (idx < editablePs.length - 1) {
            updateInlineElementTextPreservingStyle(pEl, lines[idx]);
          } else {
            updateInlineElementTextPreservingStyle(pEl, lines.slice(idx).join('\n'));
          }
        });
      }
      return true;
    }

    // Sel multi-paragraf biasa tanpa QR (misal kolom CATATAN SPPD)
    const templateP = cloneAstNode(blockKids[blockKids.length - 1]);
    if (lines.length === 0) {
      blockKids.forEach((pEl, idx) => {
        if (idx === 0) updateInlineElementTextPreservingStyle(pEl, '');
      });
      cellNode.children = [blockKids[0]];
      return true;
    }

    const nonSpacerKids = blockKids.filter((p) => !isSpacerParagraphElement(p));
    const targets = nonSpacerKids.length > 0 ? nonSpacerKids : blockKids;
    const updatedBlocks = [];

    lines.forEach((line, idx) => {
      if (idx < targets.length) {
        updateInlineElementTextPreservingStyle(targets[idx], line);
        updatedBlocks.push(targets[idx]);
      } else {
        const cloned = cloneAstNode(templateP);
        // Paragraf tambahan jangan mewarisi tag strong/em penuh jika hanya lanjut isi biasa
        cloned.children = textToInlineAst(line);
        updatedBlocks.push(cloned);
      }
    });

    cellNode.children = updatedBlocks;
    return true;
  }

  updateInlineElementTextPreservingStyle(cellNode, cleanText);
  return true;
}

function collectTableAstRows(tableNode) {
  const rows = [];
  const walk = (n, parentChildren = null, inThead = false) => {
    if (!n || n.type !== 'element') return;
    if (n.tag === 'tr') {
      const cells = (n.children || []).filter(
        (c) =>
          c.type === 'element' &&
          (c.tag === 'td' || c.tag === 'th') &&
          !String(c.attrs?.class || '').split(/\s+/).includes('id-row-action-cell')
      );
      const isHeader =
        n.attrs?.['data-header'] === 'true' ||
        inThead ||
        (cells.length > 0 && cells.every((c) => c.tag === 'th'));
      rows.push({ tr: n, parentChildren, isHeader, cells });
      return;
    }
    const nextInThead = inThead || n.tag === 'thead';
    (n.children || []).forEach((child) => walk(child, n.children, nextInThead));
  };
  walk(tableNode, null, false);
  return rows;
}

function detectTableBorderStyle(tableNode) {
  const cls = String(tableNode?.attrs?.class || '').toLowerCase();
  if (tableNode?.attrs?.['data-border-style']) return tableNode.attrs['data-border-style'];
  if (cls.includes('smart-table-sppd')) return 'sppd';
  if (cls.includes('smart-table-none')) return 'none';
  return 'grid';
}

function parseTableColWidths(tableNode) {
  const attr = String(tableNode?.attrs?.['data-col-widths'] || '').trim();
  if (attr) {
    const parsed = attr.split(',').map((n) => parseFloat(n.trim())).filter((n) => !Number.isNaN(n) && n > 0);
    if (parsed.length > 0) return parsed;
  }
  const colgroup = (tableNode?.children || []).find((c) => c.type === 'element' && c.tag === 'colgroup');
  if (colgroup) {
    const cols = (colgroup.children || []).filter((c) => c.type === 'element' && c.tag === 'col');
    const widths = cols.map((c) => {
      const m = String(c.attrs?.style || '').match(/width\s*:\s*([\d.]+)\s*%/i);
      return m ? parseFloat(m[1]) : 0;
    });
    if (widths.length > 0 && widths.every((w) => w > 0)) return widths;
  }
  return [];
}

function parseCellIndexEntry(cellNode, colIndex) {
  const colspan = Math.max(1, parseInt(cellNode.attrs?.colspan || '1', 10) || 1);
  const rowspan = Math.max(1, parseInt(cellNode.attrs?.rowspan || '1', 10) || 1);
  const blockKids = getBlockParagraphChildren(cellNode);

  if (blockKids.length > 0) {
    const paragraphs = blockKids.map((pEl, pIndex) => {
      if (isQrSlotElement(pEl)) {
        return { pIndex, isQrSlot: true, text: '[QR_SLOT_TERKUNCI]' };
      }
      if (isSpacerParagraphElement(pEl)) {
        return { pIndex, isSpacer: true, text: '' };
      }
      return { pIndex, text: getNodePlainText(pEl).replace(/\u00a0/g, ' ').trim() };
    });
    const combinedText = paragraphs
      .filter((p) => !p.isSpacer)
      .map((p) => p.text)
      .join('\n');
    return {
      col: colIndex,
      ...(colspan > 1 ? { colspan } : {}),
      ...(rowspan > 1 ? { rowspan } : {}),
      ...(paragraphs.some((p) => p.isQrSlot) ? { hasQrSlot: true } : {}),
      text: combinedText,
      paragraphs: paragraphs.filter((p) => !p.isSpacer)
    };
  }

  return {
    col: colIndex,
    ...(colspan > 1 ? { colspan } : {}),
    ...(rowspan > 1 ? { rowspan } : {}),
    text: getNodePlainText(cellNode).replace(/\u00a0/g, ' ').trim()
  };
}

function parseSignatureLinesFromAst(sigTableNode) {
  const rows = collectTableAstRows(sigTableNode);
  const lines = [];
  rows.forEach((r) => {
    const sigCell = r.cells[r.cells.length - 1];
    if (!sigCell) return;
    const raw = getNodePlainText(sigCell).replace(/\u00a0/g, ' ');
    raw.split(/\r?\n/).map((l) => l.trim()).filter(Boolean).forEach((l) => lines.push(l));
  });
  return lines;
}

function parseMetaRecipientFromAst(metaTableNode) {
  const rows = collectTableAstRows(metaTableNode);
  const metaRows = [];
  const recipientLines = [];
  rows.forEach((r) => {
    const left = r.cells[0] ? getNodePlainText(r.cells[0]).replace(/\u00a0/g, ' ').trim() : '';
    const right = r.cells.length > 1
      ? getNodePlainText(r.cells[r.cells.length - 1]).replace(/\u00a0/g, ' ').trim()
      : '';
    if (left) {
      left.split(/\r?\n/).map((l) => l.trim()).filter(Boolean).forEach((l) => metaRows.push(l));
    }
    if (right) {
      right.split(/\r?\n/).map((l) => l.trim()).filter(Boolean).forEach((l) => recipientLines.push(l));
    }
  });
  return { metaRows, recipientLines };
}

/**
 * Membangun indeks struktural bedah (slot ID + peta sel AST) langsung dari array halaman Smart Word Pro
 * tanpa kehilangan atribut HTML, lebar kolom, tinggi baris, maupun slot QR.
 */
export function buildSurgicalIndexFromSmartPages(pages = [], schoolProfile = {}) {
  const astPages = [];
  const nodeMap = new Map();
  const orderedBlocks = [];
  let titleEntry = null;
  let nomorEntry = null;
  let titleText = '';
  let nomorSurat = '';

  pages.forEach((page, pageIndex) => {
    const ast = parseHtmlNodes(page?.html || '<p><br></p>');
    astPages.push({
      id: page?.id || `spage-${Date.now()}-${pageIndex + 1}`,
      isManualBreak: Boolean(page?.isManualBreak),
      ast
    });

    const pageSummaryBlocks = [];
    const counters = { b: 0, t: 0, id: 0, sig: 0, list: 0, meta: 0 };

    const visitNodes = (nodesList) => {
      if (!Array.isArray(nodesList)) return;
      for (const node of nodesList) {
        if (!node || node.type !== 'element') continue;
        const role = node.attrs?.['data-smart-role'] || '';

        if (role === 'title') {
          const text = getNodePlainText(node).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
          titleText = text || titleText;
          titleEntry = { pageIndex, node, parentArray: nodesList, kind: 'title' };
          continue;
        }

        if (role === 'nomor') {
          const text = getNodePlainText(node).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
          nomorSurat = text.replace(/^Nomor\s*:\s*/i, '').trim() || nomorSurat;
          nomorEntry = { pageIndex, node, parentArray: nodesList, kind: 'nomor' };
          continue;
        }

        if (role === 'meta-recipient') {
          const blockId = `p${pageIndex}_meta${counters.meta++}`;
          const parsedMeta = parseMetaRecipientFromAst(node);
          const entry = { blockId, pageIndex, kind: 'metaRecipient', node, parentArray: nodesList };
          nodeMap.set(blockId, entry);
          orderedBlocks.push(entry);
          pageSummaryBlocks.push({ blockId, type: 'metaRecipient', ...parsedMeta });
          continue;
        }

        if (node.tag === 'table') {
          if (role === 'signature') {
            const blockId = `p${pageIndex}_sig${counters.sig++}`;
            const entry = { blockId, pageIndex, kind: 'signature', node, parentArray: nodesList };
            nodeMap.set(blockId, entry);
            orderedBlocks.push(entry);
            pageSummaryBlocks.push({
              blockId,
              type: 'signature',
              lines: parseSignatureLinesFromAst(node)
            });
            continue;
          }

          const astRows = collectTableAstRows(node);
          if (role === 'identity-table') {
            const blockId = `p${pageIndex}_id${counters.id++}`;
            const entry = { blockId, pageIndex, kind: 'identityTable', node, parentArray: nodesList };
            nodeMap.set(blockId, entry);
            orderedBlocks.push(entry);
            pageSummaryBlocks.push({
              blockId,
              type: 'identityTable',
              rows: astRows.map((r, rIdx) => ({
                row: rIdx,
                label: r.cells[0] ? getNodePlainText(r.cells[0]).replace(/\u00a0/g, ' ').trim() : '',
                value: r.cells[1] ? getNodePlainText(r.cells[1]).replace(/\u00a0/g, ' ').trim() : ''
              }))
            });
            continue;
          }

          const blockId = `p${pageIndex}_t${counters.t++}`;
          const colWidthsPct = parseTableColWidths(node);
          const borderStyle = detectTableBorderStyle(node);
          const entry = { blockId, pageIndex, kind: 'table', node, parentArray: nodesList };
          nodeMap.set(blockId, entry);
          orderedBlocks.push(entry);
          pageSummaryBlocks.push({
            blockId,
            type: 'table',
            borderStyle,
            ...(colWidthsPct.length > 0 ? { colWidthsPct } : {}),
            rowCount: astRows.length,
            rows: astRows.map((r, rIdx) => ({
              row: rIdx,
              ...(r.isHeader ? { isHeader: true } : {}),
              cells: r.cells.map((c, cIdx) => parseCellIndexEntry(c, cIdx))
            }))
          });
          continue;
        }

        if (node.tag === 'ul' || node.tag === 'ol') {
          const blockId = `p${pageIndex}_list${counters.list++}`;
          const liNodes = (node.children || []).filter((c) => c.type === 'element' && c.tag === 'li');
          const entry = { blockId, pageIndex, kind: 'list', node, parentArray: nodesList };
          nodeMap.set(blockId, entry);
          orderedBlocks.push(entry);
          pageSummaryBlocks.push({
            blockId,
            type: 'list',
            ordered: node.tag === 'ol',
            items: liNodes.map((li, itemIndex) => ({
              itemIndex,
              text: getNodePlainText(li).replace(/\u00a0/g, ' ').trim()
            }))
          });
          continue;
        }

        if (WRAPPER_CONTAINER_TAGS.has(node.tag) && hasNestedBlockChild(node)) {
          visitNodes(node.children || []);
          continue;
        }

        if (['p', 'div', 'blockquote', 'h1', 'h2', 'h3', 'h4'].includes(node.tag)) {
          if (isQrSlotElement(node)) continue;
          const txt = getNodePlainText(node).replace(/\u00a0/g, ' ').trim();
          if (!txt) continue;
          const blockId = `p${pageIndex}_b${counters.b++}`;
          const entry = { blockId, pageIndex, kind: 'paragraph', node, parentArray: nodesList };
          nodeMap.set(blockId, entry);
          orderedBlocks.push(entry);
          pageSummaryBlocks.push({
            blockId,
            type: /^h[1-4]$/.test(node.tag) ? 'heading' : 'paragraph',
            text: txt
          });
        }
      }
    };

    visitNodes(ast);

    astPages[pageIndex].summary = {
      pageNumber: pageIndex + 1,
      blocks: pageSummaryBlocks
    };
  });

  const snapshot = {
    surgicalMode: 'smart_pages',
    totalPages: astPages.length,
    letterTitle: titleText || 'SURAT DINAS SEKOLAH',
    nomorSurat: nomorSurat || '',
    schoolName: schoolProfile?.schoolName || 'SMK IT Al-Hawari',
    pages: astPages.map((p) => p.summary)
  };

  return {
    astPages,
    nodeMap,
    orderedBlocks,
    roleNodes: { title: titleEntry, nomor: nomorEntry },
    snapshot
  };
}

function handleAddTableRowAst(tableNode, cells = [], insertAtRow = undefined) {
  const astRows = collectTableAstRows(tableNode);
  if (astRows.length === 0) return false;

  const nonHeaderRows = astRows.filter((r) => !r.isHeader);
  const matchingCellCount = [...nonHeaderRows]
    .reverse()
    .find((r) => r.cells.length === cells.length);
  const templateInfo = matchingCellCount || nonHeaderRows[nonHeaderRows.length - 1] || astRows[astRows.length - 1];
  if (!templateInfo || !templateInfo.parentChildren) return false;

  const clonedTr = cloneAstNode(templateInfo.tr);
  if (clonedTr.attrs && clonedTr.attrs['data-header']) {
    delete clonedTr.attrs['data-header'];
  }
  const clonedCells = (clonedTr.children || []).filter(
    (c) => c.type === 'element' && (c.tag === 'td' || c.tag === 'th')
  );
  clonedCells.forEach((cellEl, cIdx) => {
    if (cellEl.tag === 'th') cellEl.tag = 'td';
    updateTableCellNodePreservingLayout(cellEl, cells[cIdx] !== undefined ? String(cells[cIdx]) : '');
  });

  const targetParent = templateInfo.parentChildren;
  if (typeof insertAtRow === 'number' && Number.isInteger(insertAtRow) && insertAtRow >= 0 && insertAtRow < astRows.length) {
    const refInfo = astRows[insertAtRow];
    if (refInfo && refInfo.parentChildren === targetParent) {
      const idxInParent = targetParent.indexOf(refInfo.tr);
      if (idxInParent !== -1) {
        targetParent.splice(idxInParent, 0, clonedTr);
        return true;
      }
    }
  }
  targetParent.push(clonedTr);
  return true;
}

function handleDeleteTableRowAst(tableNode, rowIndex) {
  const astRows = collectTableAstRows(tableNode);
  if (rowIndex < 0 || rowIndex >= astRows.length) return false;
  const target = astRows[rowIndex];
  const hasOtherBodyRows = astRows.filter((r) => !r.isHeader).length > 1;
  if (target.isHeader && hasOtherBodyRows) return false;
  if (!target.parentChildren) return false;
  const idx = target.parentChildren.indexOf(target.tr);
  if (idx === -1) return false;
  target.parentChildren.splice(idx, 1);
  return true;
}

function handleSetSignatureAst(sigTableNode, patch = {}, schoolProfile = {}) {
  const rows = collectTableAstRows(sigTableNode);
  if (rows.length === 0) return false;
  const sigCell = rows[0].cells[rows[0].cells.length - 1];
  if (!sigCell) return false;

  const existingLines = parseSignatureLinesFromAst(sigTableNode);
  const firstLine = existingLines[0] || '';
  const commaIdx = firstLine.indexOf(',');
  const curCity = commaIdx > 0 ? firstLine.slice(0, commaIdx).trim() : (schoolProfile.city || 'Garut');
  const curDate = commaIdx > 0 ? firstLine.slice(commaIdx + 1).trim() : (firstLine || '');
  const curTitle = existingLines[1] || schoolProfile.principalTitle || 'Kepala Sekolah';
  const curName = existingLines[2] || schoolProfile.principalName || 'Dian Nurhadian, S.Pd.';
  const curNip = existingLines[3] || formatSignerNipLine(schoolProfile.principalNip || '-');

  const nextCity = patch.city ? String(patch.city).trim() : curCity;
  const nextDate = patch.date ? String(patch.date).trim() : curDate;
  const nextTitle = patch.title ? String(patch.title).trim() : curTitle;
  const nextName = patch.name ? String(patch.name).trim() : curName;
  const nextNip = patch.nip !== undefined
    ? (/^nip/i.test(String(patch.nip).trim()) ? String(patch.nip).trim() : formatSignerNipLine(patch.nip))
    : curNip;

  sigCell.children = [
    { type: 'text', text: `${nextCity}, ${nextDate}` },
    { type: 'element', tag: 'br', attrs: {}, children: [] },
    { type: 'text', text: nextTitle },
    { type: 'element', tag: 'br', attrs: {}, children: [] },
    { type: 'element', tag: 'br', attrs: {}, children: [] },
    { type: 'element', tag: 'br', attrs: {}, children: [] },
    { type: 'element', tag: 'br', attrs: {}, children: [] },
    {
      type: 'element',
      tag: 'strong',
      attrs: {},
      children: [{ type: 'element', tag: 'u', attrs: {}, children: [{ type: 'text', text: nextName }] }]
    },
    ...(nextNip
      ? [{ type: 'element', tag: 'br', attrs: {}, children: [] }, { type: 'text', text: nextNip }]
      : [])
  ];
  return true;
}

function handleSetMetaRecipientAst(metaTableNode, patch = {}) {
  const rows = collectTableAstRows(metaTableNode);
  if (rows.length === 0) return false;
  const leftCell = rows[0].cells[0];
  const rightCell = rows[0].cells[rows[0].cells.length - 1];
  let changed = false;

  if (leftCell && (patch.sifat !== undefined || patch.hal !== undefined || patch.perihal !== undefined || patch.lampiran !== undefined)) {
    const existing = getNodePlainText(leftCell).replace(/\u00a0/g, ' ').split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
    const updated = [];
    let hasSifat = false;
    let hasLampiran = false;
    let hasPerihal = false;

    existing.forEach((line) => {
      if (/^sifat\s*:/i.test(line)) {
        hasSifat = true;
        updated.push(patch.sifat !== undefined ? `Sifat : ${String(patch.sifat).trim()}` : line);
      } else if (/^lampiran\s*:/i.test(line)) {
        hasLampiran = true;
        updated.push(patch.lampiran !== undefined ? `Lampiran : ${String(patch.lampiran).trim()}` : line);
      } else if (/^(perihal|hal)\s*:/i.test(line)) {
        hasPerihal = true;
        const val = patch.hal ?? patch.perihal;
        updated.push(val !== undefined ? `Perihal : ${String(val).trim()}` : line);
      } else {
        updated.push(line);
      }
    });
    if (!hasSifat && patch.sifat !== undefined) updated.unshift(`Sifat : ${String(patch.sifat).trim()}`);
    if (!hasLampiran && patch.lampiran !== undefined && patch.lampiran !== '-') {
      updated.splice(1, 0, `Lampiran : ${String(patch.lampiran).trim()}`);
    }
    if (!hasPerihal && (patch.hal !== undefined || patch.perihal !== undefined)) {
      updated.push(`Perihal : ${String(patch.hal ?? patch.perihal).trim()}`);
    }
    leftCell.children = textToInlineAst(updated.join('\n'));
    changed = true;
  }

  if (rightCell && (Array.isArray(patch.tujuanLines) || patch.tujuan !== undefined)) {
    let lines = Array.isArray(patch.tujuanLines) ? patch.tujuanLines.map((l) => String(l).trim()).filter(Boolean) : null;
    if (!lines && patch.tujuan !== undefined) {
      const tRaw = String(patch.tujuan).trim();
      const ythLine = /^yth/i.test(tRaw) ? tRaw : `Yth. ${tRaw}`;
      lines = ['Kepada', ythLine, 'Di Tempat'];
    }
    if (lines && lines.length > 0) {
      rightCell.children = textToInlineAst(lines.join('\n'));
      changed = true;
    }
  }

  return changed;
}

function normalizeColonValue(newVal, existingVal = '') {
  const clean = String(newVal ?? '').trim();
  if (!clean) return ': ';
  if (clean.startsWith(':')) return clean;
  if (!existingVal || existingVal.trim().startsWith(':')) return `: ${clean}`;
  return `: ${clean}`;
}

function handleIdentityPatchAst(idTableNode, patch = {}) {
  const op = String(patch.op || '').trim();
  const astRows = collectTableAstRows(idTableNode);

  if (op === 'setIdentityRow') {
    const rIdx = Number(patch.row);
    if (!Number.isInteger(rIdx) || rIdx < 0 || rIdx >= astRows.length) return false;
    const cells = astRows[rIdx].cells;
    let changed = false;
    if (patch.label !== undefined && cells[0]) {
      updateInlineElementTextPreservingStyle(cells[0], String(patch.label).trim());
      changed = true;
    }
    if (patch.value !== undefined && cells[1]) {
      const prevVal = getNodePlainText(cells[1]);
      updateInlineElementTextPreservingStyle(cells[1], normalizeColonValue(patch.value, prevVal));
      changed = true;
    }
    return changed;
  }

  if (op === 'addIdentityRow') {
    const lastRow = astRows[astRows.length - 1];
    if (!lastRow || !lastRow.parentChildren) return false;
    const clonedTr = cloneAstNode(lastRow.tr);
    const clonedCells = (clonedTr.children || []).filter((c) => c.type === 'element' && (c.tag === 'td' || c.tag === 'th'));
    if (clonedCells[0]) updateInlineElementTextPreservingStyle(clonedCells[0], String(patch.label || '').trim());
    if (clonedCells[1]) updateInlineElementTextPreservingStyle(clonedCells[1], normalizeColonValue(patch.value || ''));
    const afterRow = Number(patch.afterRow);
    if (Number.isInteger(afterRow) && afterRow >= 0 && afterRow < astRows.length && astRows[afterRow].parentChildren === lastRow.parentChildren) {
      const pos = lastRow.parentChildren.indexOf(astRows[afterRow].tr);
      lastRow.parentChildren.splice(pos + 1, 0, clonedTr);
    } else {
      lastRow.parentChildren.push(clonedTr);
    }
    return true;
  }

  if (op === 'deleteIdentityRow') {
    return handleDeleteTableRowAst(idTableNode, Number(patch.row));
  }

  return false;
}

/**
 * Fallback cerdas bila LLM mengembalikan format JSON lama (`{ blocks: [...] }`)
 * saat pengguna TIDAK meminta perubahan tata letak.
 * Menyalin pembaruan teks secara sel-demi-sel ke AST yang ada tanpa merusak tabel asli.
 */
function pairCompatibleLegacyTables(targetEntries = [], legacyTables = [], getTargetColsFn, getLegacyColsFn) {
  if (targetEntries.length === legacyTables.length) {
    return targetEntries.map((entry, idx) => ({ entry, src: legacyTables[idx] }));
  }
  const pairs = [];
  let cursor = 0;
  for (const src of legacyTables) {
    const srcCols = getLegacyColsFn(src);
    let matchedIdx = -1;
    for (let i = cursor; i < targetEntries.length; i++) {
      if (getTargetColsFn(targetEntries[i]) === srcCols) {
        matchedIdx = i;
        break;
      }
    }
    if (matchedIdx !== -1) {
      pairs.push({ entry: targetEntries[matchedIdx], src });
      cursor = matchedIdx + 1;
    }
  }
  return pairs;
}

function getLegacyTableMaxCols(srcTable) {
  const rows = Array.isArray(srcTable?.rows) ? srcTable.rows : [];
  let maxC = 0;
  for (const r of rows) {
    const cells = Array.isArray(r) ? r : (r?.cells || []);
    if (cells.length > maxC) maxC = cells.length;
  }
  return maxC;
}

function reconcileLegacyBlocksOntoAstPages(indexObj, editedJson = {}, schoolProfile = {}, warnings = []) {
  let count = 0;
  const { roleNodes, orderedBlocks } = indexObj;

  if (editedJson.letterTitle && roleNodes.title?.node) {
    updateInlineElementTextPreservingStyle(roleNodes.title.node, String(editedJson.letterTitle).trim().toUpperCase());
    count++;
  }
  if (editedJson.nomorSurat && roleNodes.nomor?.node) {
    const rawNo = String(editedJson.nomorSurat).replace(/^Nomor\s*:\s*/i, '').trim();
    const formatted = editedJson.kodeJenisSurat
      ? formatSmartNomorSurat({ rawNomor: rawNo, forcedTypeCode: editedJson.kodeJenisSurat })
      : rawNo;
    updateInlineElementTextPreservingStyle(roleNodes.nomor.node, `Nomor: ${formatted}`);
    count++;
  }

  const metaEntry = orderedBlocks.find((b) => b.kind === 'metaRecipient');
  if (metaEntry && (editedJson.sifat || editedJson.hal || editedJson.tujuan)) {
    if (handleSetMetaRecipientAst(metaEntry.node, {
      sifat: editedJson.sifat,
      hal: editedJson.hal,
      tujuan: editedJson.tujuan
    })) {
      count++;
    }
  }

  const legacyBlocks = Array.isArray(editedJson.blocks) ? editedJson.blocks : [];
  const targetIdTables = orderedBlocks.filter((b) => b.kind === 'identityTable');
  const legacyIdTables = legacyBlocks.filter((b) => b?.type === 'identityTable' && Array.isArray(b.rows));
  targetIdTables.forEach((entry, idx) => {
    const src = legacyIdTables[idx];
    if (!src) return;
    const astRows = collectTableAstRows(entry.node);
    const minLen = Math.min(astRows.length, src.rows.length);
    for (let r = 0; r < minLen; r++) {
      if (handleIdentityPatchAst(entry.node, {
        op: 'setIdentityRow',
        row: r,
        label: src.rows[r].label,
        value: src.rows[r].value
      })) count++;
    }
    if (src.rows.length > astRows.length) {
      for (let r = astRows.length; r < src.rows.length; r++) {
        if (handleIdentityPatchAst(entry.node, {
          op: 'addIdentityRow',
          label: src.rows[r].label,
          value: src.rows[r].value
        })) count++;
      }
    }
  });

  const targetTables = orderedBlocks.filter((b) => b.kind === 'table');
  const legacyTables = legacyBlocks.filter((b) => b?.type === 'table' && Array.isArray(b.rows));
  const tablePairs = pairCompatibleLegacyTables(
    targetTables,
    legacyTables,
    (entry) => Math.max(0, ...collectTableAstRows(entry.node).map((r) => r.cells.length)),
    getLegacyTableMaxCols
  );
  tablePairs.forEach(({ entry, src }) => {
    const astRows = collectTableAstRows(entry.node);
    const srcRows = src.rows.map((r) => (Array.isArray(r) ? r : (r?.cells || [])));

    // Cek kompatibilitas kolom dasar agar tabel kompleks seperti SPPD Keterangan I/II tidak hancur
    const hasRowOrColSpan = astRows.some((r) =>
      r.cells.some((c) => Number(c.attrs?.colspan || 1) > 1 || Number(c.attrs?.rowspan || 1) > 1)
    );
    const bodyRows = astRows.filter((r) => !r.isHeader);
    const firstCols = astRows[0]?.cells?.length || 0;
    const srcFirstCols = srcRows[0]?.length || 0;

    if (hasRowOrColSpan && astRows.length !== srcRows.length) {
      warnings.push(`Tabel kompleks (${entry.blockId}) dikunci strukturnya karena jumlah baris berubah.`);
      // Tetap salin baris data yang cocok bila jumlah baris akhir sama setelah offset
    }

    const commonRows = Math.min(astRows.length, srcRows.length);
    for (let r = 0; r < commonRows; r++) {
      const targetCells = astRows[r].cells;
      const incomingCells = srcRows[r];
      if (targetCells.length !== incomingCells.length && hasRowOrColSpan) continue;
      const commonCols = Math.min(targetCells.length, incomingCells.length);
      for (let c = 0; c < commonCols; c++) {
        const cellItem = incomingCells[c];
        const cellText = typeof cellItem === 'string'
          ? cellItem
          : (Array.isArray(cellItem?.paragraphs)
              ? cellItem.paragraphs.map((p) => p?.text || '').join('\n')
              : String(cellItem?.text ?? ''));
        const curText = getNodePlainText(targetCells[c]).replace(/\u00a0/g, ' ').trim();
        if (cellText.trim() !== curText && updateTableCellNodePreservingLayout(targetCells[c], cellText)) {
          count++;
        }
      }
    }

    // Jika tabel datar tanpa merge cell (misal tabel daftar pegawai Surat Tugas), dukung tambah/kurangi baris data
    if (!hasRowOrColSpan && firstCols === srcFirstCols && bodyRows.length > 0) {
      if (srcRows.length > astRows.length) {
        for (let r = astRows.length; r < srcRows.length; r++) {
          const rowVals = srcRows[r].map((c) => (typeof c === 'string' ? c : String(c?.text ?? '')));
          if (handleAddTableRowAst(entry.node, rowVals)) count++;
        }
      } else if (srcRows.length < astRows.length && srcRows.length >= (astRows.length - bodyRows.length + 1)) {
        for (let r = astRows.length - 1; r >= srcRows.length; r--) {
          if (handleDeleteTableRowAst(entry.node, r)) count++;
        }
      }
    }
  });

  const targetParas = orderedBlocks.filter((b) => b.kind === 'paragraph');
  const legacyParas = legacyBlocks.filter((b) => b?.type === 'paragraph' && String(b.text || '').trim());
  if (targetParas.length === legacyParas.length) {
    targetParas.forEach((entry, idx) => {
      const nextTxt = String(legacyParas[idx].text || '').trim();
      const curTxt = getNodePlainText(entry.node).replace(/\u00a0/g, ' ').trim();
      if (nextTxt && nextTxt !== curTxt) {
        updateInlineElementTextPreservingStyle(entry.node, nextTxt);
        count++;
      }
    });
  }

  const targetSigs = orderedBlocks.filter((b) => b.kind === 'signature');
  if (editedJson.signer && targetSigs.length === 1) {
    if (handleSetSignatureAst(targetSigs[0].node, editedJson.signer, schoolProfile)) count++;
  }

  return count;
}

/**
 * Menerapkan patch bedah JSON dari AI ke halaman Smart Word Pro secara in-place.
 * Menjaga 100% tabel, lebar kolom, tinggi baris, colspan/rowspan, borderStyle, dan slot QR.
 */
export function applySurgicalPatchesToSmartPages(pages = [], editedJson = {}, options = {}) {
  const schoolProfile = options.schoolProfile || {};
  const allowLayoutRestructure = Boolean(options.allowLayoutRestructure);
  const warnings = [];
  const indexObj = buildSurgicalIndexFromSmartPages(pages, schoolProfile);
  const { astPages, nodeMap, roleNodes } = indexObj;

  let appliedCount = 0;
  let mode = 'surgical';
  const patches = Array.isArray(editedJson?.patches) ? editedJson.patches : null;

  if (patches && patches.length > 0) {
    for (const patch of patches) {
      if (!patch || typeof patch !== 'object') continue;
      const op = String(patch.op || '').trim();

      if (op === 'setField') {
        const target = String(patch.target || '').toLowerCase();
        if ((target === 'title' || target === 'lettertitle') && roleNodes.title?.node) {
          updateInlineElementTextPreservingStyle(roleNodes.title.node, String(patch.text || '').toUpperCase());
          appliedCount++;
        } else if ((target === 'nomor' || target === 'nomorsurat') && roleNodes.nomor?.node) {
          const clean = String(patch.text || '').trim();
          const finalNomor = /^nomor\s*:/i.test(clean) ? clean : `Nomor: ${clean}`;
          updateInlineElementTextPreservingStyle(roleNodes.nomor.node, finalNomor);
          appliedCount++;
        }
        continue;
      }

      if (op === 'appendPage') {
        const renderPageFn = options.renderBlocksToPageHtml;
        let newHtml = '<p><br></p>';
        if (typeof renderPageFn === 'function' && Array.isArray(patch.blocks)) {
          newHtml = renderPageFn(patch.blocks) || '<p><br></p>';
        } else if (Array.isArray(patch.blocks)) {
          newHtml = patch.blocks
            .map((b) => `<p>${String(b?.text || '').replace(/</g, '&lt;')}</p>`)
            .join('\n') || '<p><br></p>';
        }
        astPages.push({
          id: `spage-${Date.now()}-${astPages.length + 1}`,
          isManualBreak: true,
          ast: parseHtmlNodes(newHtml)
        });
        appliedCount++;
        continue;
      }

      const entry = nodeMap.get(patch.blockId);
      if (!entry) {
        if (op === 'setMetaRecipient') {
          const metaFallback = indexObj.orderedBlocks.find((b) => b.kind === 'metaRecipient');
          if (metaFallback && handleSetMetaRecipientAst(metaFallback.node, patch)) {
            appliedCount++;
            continue;
          }
        }
        warnings.push(`Blok target tidak ditemukan: ${patch.blockId || op}`);
        continue;
      }

      if ((op === 'setParagraph' || op === 'setText' || op === 'replaceText') && entry.kind === 'paragraph') {
        updateInlineElementTextPreservingStyle(entry.node, String(patch.text ?? ''));
        appliedCount++;
      } else if ((op === 'setIdentityRow' || op === 'addIdentityRow' || op === 'deleteIdentityRow') && entry.kind === 'identityTable') {
        if (handleIdentityPatchAst(entry.node, patch)) appliedCount++;
      } else if (op === 'setMetaRecipient' && entry.kind === 'metaRecipient') {
        if (handleSetMetaRecipientAst(entry.node, patch)) appliedCount++;
      } else if (op === 'setTableCell' && entry.kind === 'table') {
        const rows = collectTableAstRows(entry.node);
        const targetRow = rows[Number(patch.row)];
        const targetCell = targetRow?.cells?.[Number(patch.col)];
        if (targetCell && updateTableCellNodePreservingLayout(targetCell, String(patch.text ?? ''), patch.paragraphIndex)) {
          appliedCount++;
        } else {
          warnings.push(`Sel (${patch.row},${patch.col}) tidak ditemukan pada ${patch.blockId}`);
        }
      } else if (op === 'addTableRow' && entry.kind === 'table') {
        if (handleAddTableRowAst(entry.node, Array.isArray(patch.cells) ? patch.cells : [], patch.insertAtRow)) {
          appliedCount++;
        }
      } else if (op === 'deleteTableRow' && entry.kind === 'table') {
        if (handleDeleteTableRowAst(entry.node, Number(patch.row))) appliedCount++;
      } else if (op === 'setSignature' && entry.kind === 'signature') {
        if (handleSetSignatureAst(entry.node, patch, schoolProfile)) appliedCount++;
      } else if ((op === 'setList' || op === 'setListItem') && entry.kind === 'list') {
        const liNodes = (entry.node.children || []).filter((c) => c.type === 'element' && c.tag === 'li');
        if (op === 'setListItem') {
          const itemIdx = Number(patch.itemIndex);
          if (liNodes[itemIdx]) {
            updateInlineElementTextPreservingStyle(liNodes[itemIdx], String(patch.text ?? ''));
            appliedCount++;
          }
        } else if (Array.isArray(patch.items) && patch.items.length > 0) {
          const tplLi = cloneAstNode(liNodes[0] || { type: 'element', tag: 'li', attrs: {}, children: [] });
          entry.node.children = patch.items.map((txt, i) => {
            const li = liNodes[i] || cloneAstNode(tplLi);
            updateInlineElementTextPreservingStyle(li, String(txt ?? ''));
            return li;
          });
          appliedCount++;
        }
      } else if (op === 'insertParagraphAfter' || op === 'insertParagraphBefore') {
        const idxInParent = entry.parentArray.indexOf(entry.node);
        if (idxInParent !== -1 && patch.text !== undefined) {
          const newP = {
            type: 'element',
            tag: 'p',
            attrs: {},
            children: textToInlineAst(String(patch.text))
          };
          entry.parentArray.splice(op === 'insertParagraphAfter' ? idxInParent + 1 : idxInParent, 0, newP);
          appliedCount++;
        }
      } else if (op === 'deleteBlock') {
        // Lindungi tabel agar tidak terhapus tak sengaja kecuali mode restrukturisasi aktif
        if (entry.kind === 'table' && !allowLayoutRestructure) {
          warnings.push(`Tabel ${patch.blockId} dilindungi dari penghapusan otomatis.`);
          continue;
        }
        const idxInParent = entry.parentArray.indexOf(entry.node);
        if (idxInParent !== -1) {
          entry.parentArray.splice(idxInParent, 1);
          appliedCount++;
        }
      }
    }
  } else if (editedJson?.editScope === 'restructure' && allowLayoutRestructure) {
    return {
      appliedCount: 0,
      mode: 'restructure_needed',
      pages,
      meta: null,
      warnings
    };
  } else if (Array.isArray(editedJson?.blocks)) {
    mode = 'reconciled';
    appliedCount = reconcileLegacyBlocksOntoAstPages(indexObj, editedJson, schoolProfile, warnings);
  }

  const nextPages = astPages.map((p) => ({
    id: p.id,
    html: serializeHtmlNodes(p.ast) || '<p><br></p>',
    isManualBreak: p.isManualBreak
  }));

  const nextTitle = roleNodes.title?.node
    ? getNodePlainText(roleNodes.title.node).replace(/\s+/g, ' ').trim()
    : (editedJson.letterTitle || indexObj.snapshot.letterTitle);
  const nextNomor = roleNodes.nomor?.node
    ? getNodePlainText(roleNodes.nomor.node).replace(/^Nomor\s*:\s*/i, '').trim()
    : (editedJson.nomorSurat || indexObj.snapshot.nomorSurat);

  return {
    appliedCount,
    mode,
    pages: nextPages,
    meta: { title: nextTitle, nomorSurat: nextNomor },
    warnings
  };
}

// =========================================================================
// DUKUNGAN DOCUMENT MODEL LANGSUNG (LEMBAR KERJA STANDAR / wordCanvas.js)
// =========================================================================

function extractModelCellText(cell) {
  if (typeof cell === 'string') return cell;
  if (cell && Array.isArray(cell.paragraphs) && cell.paragraphs.length > 0) {
    return cell.paragraphs
      .filter((p) => !p.qrSlot)
      .map((p) => {
        if (Array.isArray(p.runs) && p.runs.length > 0) return p.runs.map((r) => r.text || '').join('');
        return String(p.text ?? '');
      })
      .join('\n');
  }
  if (cell && Array.isArray(cell.runs) && cell.runs.length > 0) {
    return cell.runs.map((r) => r.text || '').join('');
  }
  return String(cell?.text ?? '');
}

function updateModelCellPreservingStyle(rowCells, colIdx, newText, paragraphIndex = undefined) {
  const cur = rowCells[colIdx];
  const clean = normalizeStr(newText);
  if (cur === undefined) return false;

  if (typeof cur === 'string') {
    rowCells[colIdx] = clean;
    return true;
  }
  if (cur && typeof cur === 'object') {
    if (Array.isArray(cur.paragraphs) && cur.paragraphs.length > 0) {
      if (typeof paragraphIndex === 'number' && Number.isInteger(paragraphIndex) && paragraphIndex >= 0) {
        const targetP = cur.paragraphs[paragraphIndex];
        if (!targetP || targetP.qrSlot) return false;
        if (Array.isArray(targetP.runs) && targetP.runs.length > 0) {
          targetP.runs = [{ ...targetP.runs[0], text: clean }];
        } else {
          targetP.text = clean;
        }
        return true;
      }
      const editablePs = cur.paragraphs.filter((p) => !p.qrSlot && String(p.text ?? (p.runs?.[0]?.text || '')).trim() !== '');
      const lines = clean.split(/\r?\n/).filter((l) => !/\[QR_SLOT_TERKUNCI\]/i.test(l));
      const targets = editablePs.length > 0 ? editablePs : cur.paragraphs.filter((p) => !p.qrSlot);
      targets.forEach((p, idx) => {
        if (lines[idx] !== undefined) {
          const val = (idx === targets.length - 1 && lines.length > targets.length)
            ? lines.slice(idx).join('\n')
            : lines[idx];
          if (Array.isArray(p.runs) && p.runs.length > 0) p.runs = [{ ...p.runs[0], text: val }];
          else p.text = val;
        }
      });
      return true;
    }

    if (Array.isArray(cur.runs) && cur.runs.length > 0) {
      cur.runs = [{ ...cur.runs[0], text: clean }];
    }
    cur.text = clean;
    return true;
  }
  return false;
}

/**
 * Membangun indeks bedah dari `DocumentModel` (untuk Lembar Kerja / `wordCanvas.js`).
 */
export function buildSurgicalIndexFromDocModel(docModel = {}) {
  const blocks = Array.isArray(docModel?.blocks) ? docModel.blocks : [];
  const blockMap = new Map();
  const indexedBlocks = [];
  let titleBlock = null;
  let metaBlock = null;
  const counters = { b: 0, t: 0, id: 0, sig: 0, list: 0 };

  blocks.forEach((b, modelIndex) => {
    if (!b || b.type === 'letterKop' || b.type === 'pageBreak') return;

    if (b.type === 'letterTitle') {
      titleBlock = b;
      if (b.metaRecipient) metaBlock = b.metaRecipient;
      return;
    }
    if (b.type === 'letterMetaRecipient') {
      metaBlock = b;
      return;
    }
    if (b.type === 'identityTable') {
      const blockId = `p0_id${counters.id++}`;
      blockMap.set(blockId, { blockId, kind: 'identityTable', block: b, modelIndex });
      indexedBlocks.push({
        blockId,
        type: 'identityTable',
        rows: (b.rows || []).map((r, rIdx) => ({ row: rIdx, label: r.label || '', value: r.value || '' }))
      });
      return;
    }
    if (b.type === 'table') {
      const blockId = `p0_t${counters.t++}`;
      blockMap.set(blockId, { blockId, kind: 'table', block: b, modelIndex });
      const rowsSummary = (b.rows || []).map((rowEntry, rIdx) => {
        const isObjRow = rowEntry && !Array.isArray(rowEntry);
        const cells = isObjRow ? (rowEntry.cells || []) : (Array.isArray(rowEntry) ? rowEntry : []);
        return {
          row: rIdx,
          ...(isObjRow && rowEntry.isHeader ? { isHeader: true } : {}),
          cells: cells.map((cell, cIdx) => {
            const hasQr = Boolean(cell?.qrSlot || (Array.isArray(cell?.paragraphs) && cell.paragraphs.some((p) => p.qrSlot)));
            return {
              col: cIdx,
              ...(cell?.colspan > 1 ? { colspan: cell.colspan } : {}),
              ...(cell?.rowspan > 1 ? { rowspan: cell.rowspan } : {}),
              ...(hasQr ? { hasQrSlot: true } : {}),
              text: extractModelCellText(cell)
            };
          })
        };
      });
      indexedBlocks.push({
        blockId,
        type: 'table',
        borderStyle: b.borderStyle || 'grid',
        ...(Array.isArray(b.colWidthsPct) ? { colWidthsPct: b.colWidthsPct } : {}),
        rowCount: rowsSummary.length,
        rows: rowsSummary
      });
      return;
    }
    if (b.type === 'letterSignature') {
      const blockId = `p0_sig${counters.sig++}`;
      blockMap.set(blockId, { blockId, kind: 'signature', block: b, modelIndex });
      indexedBlocks.push({
        blockId,
        type: 'signature',
        lines: [`${b.city || 'Garut'}, ${b.date || ''}`, b.title || '', b.name || '', b.nip || '']
      });
      return;
    }
    if (b.type === 'list') {
      const blockId = `p0_list${counters.list++}`;
      blockMap.set(blockId, { blockId, kind: 'list', block: b, modelIndex });
      indexedBlocks.push({
        blockId,
        type: 'list',
        ordered: Boolean(b.ordered),
        items: (b.items || []).map((it, idx) => ({
          itemIndex: idx,
          text: typeof it === 'string' ? it : (Array.isArray(it) ? it.map((r) => r.text).join('') : String(it?.text ?? ''))
        }))
      });
      return;
    }
    if (b.type === 'paragraph' || b.type === 'heading') {
      const text = Array.isArray(b.runs) && b.runs.length > 0
        ? b.runs.map((r) => r.text || '').join('').trim()
        : String(b.text || '').trim();
      if (!text) return;
      const blockId = `p0_b${counters.b++}`;
      blockMap.set(blockId, { blockId, kind: 'paragraph', block: b, modelIndex });
      indexedBlocks.push({ blockId, type: b.type, text });
    }
  });

  const snapshot = {
    surgicalMode: 'doc_model',
    letterTitle: titleBlock?.title || docModel?.meta?.title || 'SURAT DINAS',
    nomorSurat: String(titleBlock?.nomor || '').replace(/^Nomor\s*:\s*/i, '').trim(),
    hideMetaRecipient: !metaBlock || metaBlock.isDeleted,
    ...(metaBlock && !metaBlock.isDeleted
      ? {
          sifat: (metaBlock.metaRows || []).find((r) => /sifat/i.test(r.label))?.value || 'Penting',
          hal: (metaBlock.metaRows || []).find((r) => /hal|perihal/i.test(r.label))?.value || '',
          tujuan: (metaBlock.recipientLines || []).join(' ')
        }
      : {}),
    pages: [{ pageNumber: 1, blocks: indexedBlocks }]
  };

  return { blockMap, titleBlock, metaBlock, snapshot };
}

/**
 * Menerapkan patch bedah AI ke salinan bersih `DocumentModel` dengan mengunci semua tabel & tata letak.
 */
export function applySurgicalPatchesToDocModel(docModel = {}, editedJson = {}, options = {}) {
  const allowLayoutRestructure = Boolean(options.allowLayoutRestructure);
  const cloned = JSON.parse(JSON.stringify(docModel));
  const { blockMap, titleBlock, metaBlock } = buildSurgicalIndexFromDocModel(cloned);
  const warnings = [];
  let appliedCount = 0;

  const patches = Array.isArray(editedJson?.patches) ? editedJson.patches : null;
  if (patches && patches.length > 0) {
    for (const patch of patches) {
      if (!patch || typeof patch !== 'object') continue;
      const op = String(patch.op || '').trim();

      if (op === 'setField' && titleBlock) {
        const target = String(patch.target || '').toLowerCase();
        if (target === 'title' || target === 'lettertitle') {
          titleBlock.title = String(patch.text || '').toUpperCase();
          if (cloned.meta) cloned.meta.title = titleBlock.title;
          appliedCount++;
        } else if (target === 'nomor' || target === 'nomorsurat') {
          const clean = String(patch.text || '').trim().replace(/^Nomor\s*:\s*/i, '');
          titleBlock.nomor = `Nomor: ${clean}`;
          if (cloned.letterData?.formData) cloned.letterData.formData.nomorSurat = clean;
          appliedCount++;
        }
        continue;
      }

      if (op === 'setMetaRecipient' && metaBlock) {
        if (Array.isArray(metaBlock.metaRows)) {
          if (patch.sifat !== undefined) {
            const row = metaBlock.metaRows.find((r) => /sifat/i.test(r.label));
            if (row) row.value = String(patch.sifat).trim();
          }
          if (patch.hal !== undefined || patch.perihal !== undefined) {
            const row = metaBlock.metaRows.find((r) => /hal|perihal/i.test(r.label));
            if (row) row.value = String(patch.hal ?? patch.perihal).trim();
          }
        }
        if (Array.isArray(patch.tujuanLines)) {
          metaBlock.recipientLines = patch.tujuanLines.map((l) => String(l).trim()).filter(Boolean);
          delete metaBlock.customHtml;
        } else if (patch.tujuan !== undefined) {
          const tRaw = String(patch.tujuan).trim();
          metaBlock.recipientLines = ['Kepada', /^yth/i.test(tRaw) ? tRaw : `Yth. ${tRaw}`, 'Di Tempat'];
          delete metaBlock.customHtml;
        }
        appliedCount++;
        continue;
      }

      const entry = blockMap.get(patch.blockId);
      if (!entry) {
        warnings.push(`Blok model tidak ditemukan: ${patch.blockId || op}`);
        continue;
      }

      if ((op === 'setParagraph' || op === 'setText' || op === 'replaceText') && entry.kind === 'paragraph') {
        const val = normalizeStr(patch.text ?? '');
        entry.block.text = val;
        const firstRun = Array.isArray(entry.block.runs) && entry.block.runs[0] ? entry.block.runs[0] : {};
        entry.block.runs = [{ ...firstRun, text: val }];
        appliedCount++;
      } else if (op === 'setIdentityRow' && entry.kind === 'identityTable') {
        const row = entry.block.rows?.[Number(patch.row)];
        if (row) {
          if (patch.label !== undefined) row.label = String(patch.label).trim();
          if (patch.value !== undefined) row.value = normalizeColonValue(patch.value, row.value);
          appliedCount++;
        }
      } else if (op === 'addIdentityRow' && entry.kind === 'identityTable') {
        entry.block.rows = entry.block.rows || [];
        entry.block.rows.push({
          label: String(patch.label || '').trim(),
          value: normalizeColonValue(patch.value || '')
        });
        appliedCount++;
      } else if (op === 'deleteIdentityRow' && entry.kind === 'identityTable') {
        const idx = Number(patch.row);
        if (Array.isArray(entry.block.rows) && idx >= 0 && idx < entry.block.rows.length) {
          entry.block.rows.splice(idx, 1);
          appliedCount++;
        }
      } else if (op === 'setTableCell' && entry.kind === 'table') {
        const rowEntry = entry.block.rows?.[Number(patch.row)];
        const cells = Array.isArray(rowEntry) ? rowEntry : rowEntry?.cells;
        if (cells && updateModelCellPreservingStyle(cells, Number(patch.col), patch.text ?? '', patch.paragraphIndex)) {
          appliedCount++;
        }
      } else if (op === 'addTableRow' && entry.kind === 'table') {
        const rows = entry.block.rows || (entry.block.rows = []);
        const newCells = Array.isArray(patch.cells) ? patch.cells.map((c) => String(c ?? '')) : [];
        const tpl = rows[rows.length - 1];
        if (tpl && !Array.isArray(tpl) && Array.isArray(tpl.cells)) {
          rows.push({ isHeader: false, cells: newCells.map((t, i) => (tpl.cells[i]?.align ? { text: t, align: tpl.cells[i].align } : { text: t })) });
        } else {
          rows.push(newCells);
        }
        appliedCount++;
      } else if (op === 'deleteTableRow' && entry.kind === 'table') {
        const idx = Number(patch.row);
        if (Array.isArray(entry.block.rows) && idx >= 0 && idx < entry.block.rows.length && entry.block.rows.length > 1) {
          entry.block.rows.splice(idx, 1);
          appliedCount++;
        }
      } else if (op === 'setSignature' && entry.kind === 'signature') {
        if (patch.city !== undefined) entry.block.city = String(patch.city).trim();
        if (patch.date !== undefined) entry.block.date = String(patch.date).trim();
        if (patch.title !== undefined) entry.block.title = String(patch.title).trim();
        if (patch.name !== undefined) entry.block.name = String(patch.name).trim();
        if (patch.nip !== undefined) entry.block.nip = String(patch.nip).trim();
        appliedCount++;
      } else if ((op === 'insertParagraphAfter' || op === 'insertParagraphBefore') && patch.text !== undefined) {
        const curIdx = cloned.blocks.indexOf(entry.block);
        if (curIdx !== -1) {
          const val = String(patch.text).trim();
          const newBlock = { type: 'paragraph', style: 'letterBody', text: val, runs: [{ text: val }] };
          cloned.blocks.splice(op === 'insertParagraphAfter' ? curIdx + 1 : curIdx, 0, newBlock);
          appliedCount++;
        }
      } else if (op === 'deleteBlock') {
        if (entry.kind === 'table' && !allowLayoutRestructure) {
          warnings.push(`Tabel ${patch.blockId} dilindungi.`);
          continue;
        }
        const curIdx = cloned.blocks.indexOf(entry.block);
        if (curIdx !== -1) {
          cloned.blocks.splice(curIdx, 1);
          appliedCount++;
        }
      }
    }

    return { appliedCount, mode: 'surgical', docModel: cloned, warnings };
  }

  if (editedJson?.editScope === 'restructure' && allowLayoutRestructure) {
    return { appliedCount: 0, mode: 'restructure_needed', docModel: cloned, warnings };
  }

  // Fallback reconciler untuk DocumentModel bila LLM mengirim blocks lama
  const legacyBlocks = Array.isArray(editedJson?.blocks) ? editedJson.blocks : [];
  if (titleBlock) {
    if (editedJson.letterTitle) {
      titleBlock.title = String(editedJson.letterTitle).trim().toUpperCase();
      appliedCount++;
    }
    if (editedJson.nomorSurat) {
      const clean = String(editedJson.nomorSurat).replace(/^Nomor\s*:\s*/i, '').trim();
      titleBlock.nomor = `Nomor: ${clean}`;
      appliedCount++;
    }
  }
  const existingTables = [...blockMap.values()].filter((e) => e.kind === 'table');
  const legacyTables = legacyBlocks.filter((b) => b?.type === 'table' && Array.isArray(b.rows));
  const modelPairs = pairCompatibleLegacyTables(
    existingTables,
    legacyTables,
    (entry) => {
      const rows = entry.block.rows || [];
      return Math.max(0, ...rows.map((r) => (Array.isArray(r) ? r.length : (r?.cells?.length || 0))));
    },
    getLegacyTableMaxCols
  );
  modelPairs.forEach(({ entry, src }) => {
    const rows = entry.block.rows || [];
    const commonR = Math.min(rows.length, src.rows.length);
    for (let r = 0; r < commonR; r++) {
      const dstCells = Array.isArray(rows[r]) ? rows[r] : rows[r]?.cells;
      const srcCells = Array.isArray(src.rows[r]) ? src.rows[r] : src.rows[r]?.cells;
      if (!dstCells || !srcCells) continue;
      const commonC = Math.min(dstCells.length, srcCells.length);
      for (let c = 0; c < commonC; c++) {
        const txt = extractModelCellText(srcCells[c]);
        if (txt !== extractModelCellText(dstCells[c]) && updateModelCellPreservingStyle(dstCells, c, txt)) {
          appliedCount++;
        }
      }
    }
  });

  const existingIdTables = [...blockMap.values()].filter((e) => e.kind === 'identityTable');
  const legacyIdTables = legacyBlocks.filter((b) => b?.type === 'identityTable' && Array.isArray(b.rows));
  existingIdTables.forEach((entry, iIdx) => {
    const src = legacyIdTables[iIdx];
    if (!src) return;
    const commonR = Math.min((entry.block.rows || []).length, src.rows.length);
    for (let r = 0; r < commonR; r++) {
      entry.block.rows[r].label = String(src.rows[r].label || entry.block.rows[r].label).trim();
      entry.block.rows[r].value = normalizeColonValue(src.rows[r].value, entry.block.rows[r].value);
      appliedCount++;
    }
  });

  const existingParas = [...blockMap.values()].filter((e) => e.kind === 'paragraph');
  const legacyParas = legacyBlocks.filter((b) => b?.type === 'paragraph' && String(b.text || '').trim());
  if (existingParas.length === legacyParas.length) {
    existingParas.forEach((entry, pIdx) => {
      const val = String(legacyParas[pIdx].text).trim();
      if (val) {
        entry.block.text = val;
        const firstRun = Array.isArray(entry.block.runs) && entry.block.runs[0] ? entry.block.runs[0] : {};
        entry.block.runs = [{ ...firstRun, text: val }];
        appliedCount++;
      }
    });
  }

  return { appliedCount, mode: 'reconciled', docModel: cloned, warnings };
}
