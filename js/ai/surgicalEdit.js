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
import {
  cmToTwips,
  ptToTwips,
  cssLenToCm,
  normalizeAlignment,
  normalizeFontFamily,
  normalizeListStyle,
  getListStyleSpec,
  parseListNodeMeta,
  extractPageSetupFromAiPayload,
  normalizePageSetupOptions
} from '../letters/wordFormat.js';

function upsertCssStyleProp(styleStr = '', propName = '', propValue = '') {
  const cleanProp = String(propName || '').trim().toLowerCase();
  const declarations = String(styleStr || '')
    .split(';')
    .map((s) => s.trim())
    .filter(Boolean)
    .filter((decl) => {
      const idx = decl.indexOf(':');
      if (idx === -1) return false;
      return decl.slice(0, idx).trim().toLowerCase() !== cleanProp;
    });
  if (propValue !== null && propValue !== undefined && String(propValue).trim() !== '') {
    declarations.push(`${cleanProp}: ${String(propValue).trim()}`);
  }
  return declarations.join('; ');
}

function readCssStyleProp(node, propName) {
  const style = String(node?.attrs?.style || '');
  const re = new RegExp(`(?:^|;)\\s*${propName}\\s*:\\s*([^;]+)`, 'i');
  const m = style.match(re);
  return m ? m[1].trim() : '';
}

function applyAstParagraphFormat(node, patch = {}) {
  if (!node || typeof node !== 'object') return false;
  node.attrs = node.attrs || {};
  let style = node.attrs.style || '';
  let changed = false;

  const align = normalizeAlignment(patch.alignment || patch.align || '');
  if (align) {
    const cssAlign = align === 'both' ? 'justify' : align;
    style = upsertCssStyleProp(style, 'text-align', cssAlign);
    changed = true;
  }
  if (patch.firstLineIndentCm !== undefined && Number.isFinite(Number(patch.firstLineIndentCm))) {
    const cm = Math.max(0, Math.min(8, Number(patch.firstLineIndentCm)));
    style = upsertCssStyleProp(style, 'text-indent', cm > 0 ? `${cm}cm` : '0cm');
    changed = true;
  }
  if (patch.leftIndentCm !== undefined && Number.isFinite(Number(patch.leftIndentCm))) {
    const cm = Math.max(0, Math.min(8, Number(patch.leftIndentCm)));
    style = upsertCssStyleProp(style, 'padding-left', cm > 0 ? `${cm}cm` : '0cm');
    changed = true;
  }
  if (patch.spacingBeforePt !== undefined && Number.isFinite(Number(patch.spacingBeforePt))) {
    const pt = Math.max(0, Math.min(72, Number(patch.spacingBeforePt)));
    style = upsertCssStyleProp(style, 'margin-top', `${pt}pt`);
    node.attrs['data-space-before'] = String(ptToTwips(pt));
    changed = true;
  }
  if (patch.spacingAfterPt !== undefined && Number.isFinite(Number(patch.spacingAfterPt))) {
    const pt = Math.max(0, Math.min(72, Number(patch.spacingAfterPt)));
    style = upsertCssStyleProp(style, 'margin-bottom', `${pt}pt`);
    node.attrs['data-space-after'] = String(ptToTwips(pt));
    changed = true;
  }
  if (patch.lineSpacing !== undefined && Number.isFinite(Number(patch.lineSpacing))) {
    const ls = Math.max(0.8, Math.min(3.5, Number(patch.lineSpacing)));
    style = upsertCssStyleProp(style, 'line-height', String(ls));
    changed = true;
  }
  if (patch.fontSizePt !== undefined && Number.isFinite(Number(patch.fontSizePt))) {
    const pt = Math.max(8, Math.min(48, Number(patch.fontSizePt)));
    style = upsertCssStyleProp(style, 'font-size', `${pt}pt`);
    changed = true;
  }
  if (patch.fontFamily !== undefined && String(patch.fontFamily).trim()) {
    const font = normalizeFontFamily(patch.fontFamily);
    style = upsertCssStyleProp(style, 'font-family', `'${font}', serif`);
    changed = true;
  }
  if (patch.bold !== undefined) {
    style = upsertCssStyleProp(style, 'font-weight', patch.bold ? 'bold' : 'normal');
    changed = true;
  }
  if (patch.italic !== undefined) {
    style = upsertCssStyleProp(style, 'font-style', patch.italic ? 'italic' : 'normal');
    changed = true;
  }
  if (patch.underline !== undefined) {
    style = upsertCssStyleProp(style, 'text-decoration', patch.underline ? 'underline' : 'none');
    changed = true;
  }

  if (changed) {
    if (style) node.attrs.style = style;
    else delete node.attrs.style;
  }
  return changed;
}

function applyListStyleToAstNode(listNode, listStyleRaw, startAtRaw) {
  if (!listNode || typeof listNode !== 'object') return false;
  listNode.attrs = listNode.attrs || {};
  const spec = getListStyleSpec(listStyleRaw, listNode.tag === 'ol');
  listNode.tag = spec.htmlTag;
  listNode.attrs['data-list-style'] = spec.key;
  listNode.attrs.style = upsertCssStyleProp(listNode.attrs.style || '', 'list-style-type', spec.cssType);
  if (spec.key === 'bullet') {
    delete listNode.attrs.type;
    delete listNode.attrs.start;
  } else {
    listNode.attrs.type = spec.htmlType;
    if (startAtRaw !== undefined && Number(startAtRaw) > 1) {
      listNode.attrs.start = String(Math.round(Number(startAtRaw)));
    } else if (startAtRaw !== undefined) {
      delete listNode.attrs.start;
    }
  }
  return true;
}

function applyTableFormatAst(tableNode, patch = {}) {
  if (!tableNode || tableNode.tag !== 'table') return false;
  tableNode.attrs = tableNode.attrs || {};
  let classes = String(tableNode.attrs.class || 'smart-word-table').split(/\s+/).filter(Boolean);
  if (!classes.includes('smart-word-table')) classes.unshift('smart-word-table');
  let changed = false;

  if (patch.borderStyle !== undefined) {
    const bs = String(patch.borderStyle).trim().toLowerCase();
    classes = classes.filter((c) => !['smart-table-grid', 'smart-table-sppd', 'smart-table-none'].includes(c));
    if (bs === 'sppd') classes.push('smart-table-sppd');
    else if (bs === 'none') classes.push('smart-table-none');
    else classes.push('smart-table-grid');
    changed = true;
  }
  if (patch.headerShading !== undefined) {
    classes = classes.filter((c) => c !== 'no-header-bg');
    if (!patch.headerShading) classes.push('no-header-bg');
    changed = true;
  }
  if (patch.headerBold !== undefined) {
    classes = classes.filter((c) => c !== 'no-header-bold');
    if (!patch.headerBold) classes.push('no-header-bold');
    changed = true;
  }
  if (patch.align !== undefined) {
    const align = normalizeAlignment(patch.align) || 'left';
    tableNode.attrs.style = upsertCssStyleProp(tableNode.attrs.style || '', 'text-align', align === 'both' ? 'justify' : align);
    changed = true;
  }
  if (Array.isArray(patch.colWidthsPct) && patch.colWidthsPct.length > 0) {
    const cleanWidths = patch.colWidthsPct.map((n) => Math.max(4, Math.min(90, Number(n) || 10)));
    tableNode.attrs['data-col-widths'] = cleanWidths.join(',');
    const colgroup = (tableNode.children || []).find((c) => c.type === 'element' && c.tag === 'colgroup');
    if (colgroup && Array.isArray(colgroup.children)) {
      const cols = colgroup.children.filter((c) => c.type === 'element' && c.tag === 'col');
      cols.forEach((col, i) => {
        if (cleanWidths[i] !== undefined) {
          col.attrs = col.attrs || {};
          col.attrs.style = upsertCssStyleProp(col.attrs.style || '', 'width', `${cleanWidths[i]}%`);
        }
      });
    }
    changed = true;
  }

  tableNode.attrs.class = [...new Set(classes)].join(' ');
  return changed;
}

function applyDocModelParagraphFormat(block, patch = {}) {
  if (!block || typeof block !== 'object') return false;
  let changed = false;
  const align = normalizeAlignment(patch.alignment || patch.align || '');
  if (align) {
    block.alignment = align;
    changed = true;
  }
  if (patch.firstLineIndentCm !== undefined && Number.isFinite(Number(patch.firstLineIndentCm))) {
    const cm = Math.max(0, Math.min(8, Number(patch.firstLineIndentCm)));
    block.firstLineIndentCm = cm;
    block.firstLineIndentTwips = cmToTwips(cm);
    changed = true;
  }
  if (patch.leftIndentCm !== undefined && Number.isFinite(Number(patch.leftIndentCm))) {
    const cm = Math.max(0, Math.min(8, Number(patch.leftIndentCm)));
    block.leftIndentCm = cm;
    block.leftIndentTwips = cmToTwips(cm);
    changed = true;
  }
  if (patch.spacingBeforePt !== undefined && Number.isFinite(Number(patch.spacingBeforePt))) {
    block.spacingBefore = ptToTwips(Math.max(0, Math.min(72, Number(patch.spacingBeforePt))));
    changed = true;
  }
  if (patch.spacingAfterPt !== undefined && Number.isFinite(Number(patch.spacingAfterPt))) {
    block.spacingAfter = ptToTwips(Math.max(0, Math.min(72, Number(patch.spacingAfterPt))));
    changed = true;
  }
  if (patch.lineSpacing !== undefined && Number.isFinite(Number(patch.lineSpacing))) {
    block.lineSpacing = Math.max(0.8, Math.min(3.5, Number(patch.lineSpacing)));
    changed = true;
  }
  if (patch.fontSizePt !== undefined && Number.isFinite(Number(patch.fontSizePt))) {
    const pt = Math.max(8, Math.min(48, Number(patch.fontSizePt)));
    block.fontSizePt = pt;
    if (Array.isArray(block.runs)) {
      block.runs = block.runs.map((r) => ({ ...r, fontSizeHalfPt: Math.round(pt * 2) }));
    }
    changed = true;
  }
  if (patch.fontFamily !== undefined && String(patch.fontFamily).trim()) {
    const font = normalizeFontFamily(patch.fontFamily);
    block.fontFamily = font;
    if (Array.isArray(block.runs)) {
      block.runs = block.runs.map((r) => ({ ...r, fontFamily: font }));
    }
    changed = true;
  }
  for (const flag of ['bold', 'italic', 'underline']) {
    if (patch[flag] !== undefined) {
      block[flag] = Boolean(patch[flag]);
      if (Array.isArray(block.runs)) {
        block.runs = block.runs.map((r) => ({ ...r, [flag]: Boolean(patch[flag]) }));
      }
      changed = true;
    }
  }
  return changed;
}

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
export function buildSurgicalIndexFromSmartPages(pages = [], schoolProfile = {}, options = {}) {
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
          const listMeta = parseListNodeMeta(node);
          const entry = { blockId, pageIndex, kind: 'list', node, parentArray: nodesList };
          nodeMap.set(blockId, entry);
          orderedBlocks.push(entry);
          pageSummaryBlocks.push({
            blockId,
            type: 'list',
            ordered: listMeta.ordered,
            listStyle: listMeta.listStyle,
            ...(listMeta.startAt > 1 ? { startAt: listMeta.startAt } : {}),
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
          const alignment = normalizeAlignment(readCssStyleProp(node, 'text-align'));
          const firstLineIndentCm = cssLenToCm(readCssStyleProp(node, 'text-indent'));
          const entry = { blockId, pageIndex, kind: 'paragraph', node, parentArray: nodesList };
          nodeMap.set(blockId, entry);
          orderedBlocks.push(entry);
          pageSummaryBlocks.push({
            blockId,
            type: /^h[1-4]$/.test(node.tag) ? 'heading' : 'paragraph',
            ...(alignment ? { alignment } : {}),
            ...(firstLineIndentCm ? { firstLineIndentCm } : {}),
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
    ...(options?.letterOptions ? { pageSetup: normalizePageSetupOptions(options.letterOptions) } : {}),
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

function recordSurgicalChange(changes, kind, targetLabel, beforeVal, afterVal) {
  const before = normalizeStr(beforeVal).trim();
  const after = normalizeStr(afterVal).trim();
  if (before === after && kind !== 'addRow' && kind !== 'deleteRow' && kind !== 'insertParagraph' && kind !== 'appendPage') {
    return;
  }
  changes.push({
    kind,
    targetLabel: String(targetLabel || kind),
    before: before.slice(0, 220),
    after: after.slice(0, 220)
  });
}

function reconcileLegacyBlocksOntoAstPages(indexObj, editedJson = {}, schoolProfile = {}, warnings = [], changes = []) {
  let count = 0;
  const { roleNodes, orderedBlocks } = indexObj;

  if (editedJson.letterTitle && roleNodes.title?.node) {
    const prev = getNodePlainText(roleNodes.title.node).trim();
    const next = String(editedJson.letterTitle).trim().toUpperCase();
    updateInlineElementTextPreservingStyle(roleNodes.title.node, next);
    recordSurgicalChange(changes, 'field', 'Judul Surat', prev, next);
    count++;
  }
  if (editedJson.nomorSurat && roleNodes.nomor?.node) {
    const prev = getNodePlainText(roleNodes.nomor.node).trim();
    const rawNo = String(editedJson.nomorSurat).replace(/^Nomor\s*:\s*/i, '').trim();
    const formatted = editedJson.kodeJenisSurat
      ? formatSmartNomorSurat({ rawNomor: rawNo, forcedTypeCode: editedJson.kodeJenisSurat })
      : rawNo;
    const next = `Nomor: ${formatted}`;
    updateInlineElementTextPreservingStyle(roleNodes.nomor.node, next);
    recordSurgicalChange(changes, 'field', 'Nomor Surat', prev, next);
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
          recordSurgicalChange(changes, 'tableCell', `Sel Tabel (${entry.blockId}, Brs ${r + 1} Kol ${c + 1})`, curText, cellText);
          count++;
        }
      }
    }

    // Jika tabel datar tanpa merge cell (misal tabel daftar pegawai Surat Tugas), dukung tambah/kurangi baris data
    if (!hasRowOrColSpan && firstCols === srcFirstCols && bodyRows.length > 0) {
      if (srcRows.length > astRows.length) {
        for (let r = astRows.length; r < srcRows.length; r++) {
          const rowVals = srcRows[r].map((c) => (typeof c === 'string' ? c : String(c?.text ?? '')));
          if (handleAddTableRowAst(entry.node, rowVals)) {
            recordSurgicalChange(changes, 'addRow', `Tambah Baris (${entry.blockId})`, '-', rowVals.join(' | '));
            count++;
          }
        }
      } else if (srcRows.length < astRows.length && srcRows.length >= (astRows.length - bodyRows.length + 1)) {
        for (let r = astRows.length - 1; r >= srcRows.length; r--) {
          if (handleDeleteTableRowAst(entry.node, r)) {
            recordSurgicalChange(changes, 'deleteRow', `Hapus Baris (${entry.blockId}, Brs ${r + 1})`, `Baris ${r + 1}`, '(dihapus)');
            count++;
          }
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
        recordSurgicalChange(changes, 'paragraph', `Paragraf (${entry.blockId})`, curTxt, nextTxt);
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
  const changes = [];
  const indexObj = buildSurgicalIndexFromSmartPages(pages, schoolProfile, options);
  const { astPages, nodeMap, roleNodes } = indexObj;

  let appliedCount = 0;
  let mode = 'surgical';
  const pageFormatPatches = [];
  const pendingListConversions = [];
  const patches = Array.isArray(editedJson?.patches) ? editedJson.patches : null;

  if (patches && patches.length > 0) {
    for (const patch of patches) {
      if (!patch || typeof patch !== 'object') continue;
      const op = String(patch.op || '').trim();

      if (op === 'setPageFormat') {
        pageFormatPatches.push(patch);
        const fmtKeys = Object.keys(patch).filter((k) => k !== 'op').join(', ') || 'halaman';
        recordSurgicalChange(changes, 'pageFormat', 'Format Halaman & Tipografi', '-', fmtKeys);
        appliedCount++;
        continue;
      }

      if (op === 'setParagraphFormat') {
        const targets = (patch.blockId === '*' || String(patch.target || '').toLowerCase() === 'all')
          ? indexObj.orderedBlocks.filter((e) => e.kind === 'paragraph')
          : Array.isArray(patch.blockIds)
            ? patch.blockIds.map((id) => nodeMap.get(id)).filter((e) => e && e.kind === 'paragraph')
            : (() => {
                const single = nodeMap.get(patch.blockId);
                return single && single.kind === 'paragraph' ? [single] : [];
              })();
        if (targets.length === 0) {
          warnings.push(`Paragraf target format tidak ditemukan: ${patch.blockId || '*'}`);
          continue;
        }
        let fmtApplied = 0;
        for (const t of targets) {
          if (applyAstParagraphFormat(t.node, patch)) fmtApplied++;
        }
        if (fmtApplied > 0) {
          const fmtDesc = Object.keys(patch).filter((k) => !['op', 'blockId', 'blockIds', 'target'].includes(k)).join(', ');
          recordSurgicalChange(changes, 'paragraphFormat', `Format ${targets.length} Paragraf`, '-', fmtDesc);
          appliedCount++;
        }
        continue;
      }

      if (op === 'convertParagraphsToList') {
        const ids = Array.isArray(patch.blockIds) ? patch.blockIds : [];
        const entries = ids.map((id) => nodeMap.get(id)).filter((e) => e && e.kind === 'paragraph');
        if (entries.length === 0) {
          warnings.push('Blok paragraf untuk dijadikan daftar (list) tidak ditemukan.');
          continue;
        }
        pendingListConversions.push({
          entries,
          listStyle: patch.listStyle ?? (patch.ordered === false ? 'bullet' : 'decimal'),
          startAt: patch.startAt
        });
        recordSurgicalChange(
          changes,
          'convertList',
          `Ubah ${entries.length} Paragraf Jadi Daftar (${patch.listStyle || 'decimal'})`,
          `${entries.length} paragraf`,
          `Daftar ${patch.listStyle || 'decimal'}`
        );
        appliedCount++;
        continue;
      }

      if (op === 'setField') {
        const target = String(patch.target || '').toLowerCase();
        if ((target === 'title' || target === 'lettertitle') && roleNodes.title?.node) {
          const prev = getNodePlainText(roleNodes.title.node).trim();
          const next = String(patch.text || '').toUpperCase();
          updateInlineElementTextPreservingStyle(roleNodes.title.node, next);
          recordSurgicalChange(changes, 'field', 'Judul Surat', prev, next);
          appliedCount++;
        } else if ((target === 'nomor' || target === 'nomorsurat') && roleNodes.nomor?.node) {
          const prev = getNodePlainText(roleNodes.nomor.node).trim();
          const clean = String(patch.text || '').trim();
          const finalNomor = /^nomor\s*:/i.test(clean) ? clean : `Nomor: ${clean}`;
          updateInlineElementTextPreservingStyle(roleNodes.nomor.node, finalNomor);
          recordSurgicalChange(changes, 'field', 'Nomor Surat', prev, finalNomor);
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
        recordSurgicalChange(changes, 'appendPage', `Tambah Halaman ${astPages.length}`, '(tidak ada)', `Halaman baru (${Array.isArray(patch.blocks) ? patch.blocks.length : 1} blok)`);
        appliedCount++;
        continue;
      }

      const entry = nodeMap.get(patch.blockId);
      if (!entry) {
        if (op === 'setMetaRecipient') {
          const metaFallback = indexObj.orderedBlocks.find((b) => b.kind === 'metaRecipient');
          if (metaFallback) {
            const prevMeta = getNodePlainText(metaFallback.node).replace(/\s+/g, ' ').trim();
            if (handleSetMetaRecipientAst(metaFallback.node, patch)) {
              recordSurgicalChange(changes, 'meta', 'Perihal & Tujuan Surat', prevMeta, getNodePlainText(metaFallback.node).replace(/\s+/g, ' ').trim());
              appliedCount++;
              continue;
            }
          }
        }
        warnings.push(`Blok target tidak ditemukan: ${patch.blockId || op}`);
        continue;
      }

      if ((op === 'setParagraph' || op === 'setText' || op === 'replaceText') && entry.kind === 'paragraph') {
        const prevTxt = getNodePlainText(entry.node).replace(/\u00a0/g, ' ').trim();
        const nextTxt = String(patch.text ?? '');
        updateInlineElementTextPreservingStyle(entry.node, nextTxt);
        recordSurgicalChange(changes, 'paragraph', `Paragraf Hal. ${entry.pageIndex + 1} (${entry.blockId})`, prevTxt, nextTxt);
        appliedCount++;
      } else if ((op === 'setIdentityRow' || op === 'addIdentityRow' || op === 'deleteIdentityRow') && entry.kind === 'identityTable') {
        const prevTxt = getNodePlainText(entry.node).replace(/\s+/g, ' ').trim();
        if (handleIdentityPatchAst(entry.node, patch)) {
          recordSurgicalChange(changes, 'identity', `Rincian Identitas Hal. ${entry.pageIndex + 1}`, prevTxt, getNodePlainText(entry.node).replace(/\s+/g, ' ').trim());
          appliedCount++;
        }
      } else if (op === 'setMetaRecipient' && entry.kind === 'metaRecipient') {
        const prevMeta = getNodePlainText(entry.node).replace(/\s+/g, ' ').trim();
        if (handleSetMetaRecipientAst(entry.node, patch)) {
          recordSurgicalChange(changes, 'meta', 'Perihal & Tujuan Surat', prevMeta, getNodePlainText(entry.node).replace(/\s+/g, ' ').trim());
          appliedCount++;
        }
      } else if (op === 'setTableFormat' && entry.kind === 'table') {
        if (applyTableFormatAst(entry.node, patch)) {
          const fmtKeys = Object.keys(patch).filter((k) => !['op', 'blockId'].includes(k)).join(', ');
          recordSurgicalChange(changes, 'tableFormat', `Format Tabel (${entry.blockId})`, '-', fmtKeys);
          appliedCount++;
        }
      } else if (op === 'setTableCell' && entry.kind === 'table') {
        const rows = collectTableAstRows(entry.node);
        const rIdx = Number(patch.row);
        const cIdx = Number(patch.col);
        const targetRow = rows[rIdx];
        const targetCell = targetRow?.cells?.[cIdx];
        const prevCell = targetCell ? getNodePlainText(targetCell).replace(/\u00a0/g, ' ').trim() : '';
        if (targetCell && updateTableCellNodePreservingLayout(targetCell, String(patch.text ?? ''), patch.paragraphIndex)) {
          recordSurgicalChange(
            changes,
            'tableCell',
            `Sel Tabel Hal. ${entry.pageIndex + 1} (Brs ${rIdx + 1}, Kol ${cIdx + 1})`,
            prevCell,
            String(patch.text ?? '')
          );
          appliedCount++;
        } else {
          warnings.push(`Sel (${patch.row},${patch.col}) tidak ditemukan pada ${patch.blockId}`);
        }
      } else if (op === 'addTableRow' && entry.kind === 'table') {
        const cells = Array.isArray(patch.cells) ? patch.cells : [];
        if (handleAddTableRowAst(entry.node, cells, patch.insertAtRow)) {
          recordSurgicalChange(changes, 'addRow', `Tambah Baris Tabel Hal. ${entry.pageIndex + 1}`, '(baris baru)', cells.join(' | '));
          appliedCount++;
        }
      } else if (op === 'deleteTableRow' && entry.kind === 'table') {
        const rIdx = Number(patch.row);
        const prevRow = collectTableAstRows(entry.node)[rIdx]?.cells?.map(c => getNodePlainText(c).trim()).join(' | ') || `Baris ${rIdx + 1}`;
        if (handleDeleteTableRowAst(entry.node, rIdx)) {
          recordSurgicalChange(changes, 'deleteRow', `Hapus Baris Tabel Hal. ${entry.pageIndex + 1}`, prevRow, '(dihapus)');
          appliedCount++;
        }
      } else if (op === 'setSignature' && entry.kind === 'signature') {
        const prevSig = parseSignatureLinesFromAst(entry.node).join(' · ');
        if (handleSetSignatureAst(entry.node, patch, schoolProfile)) {
          recordSurgicalChange(changes, 'signature', `Blok TTD Hal. ${entry.pageIndex + 1}`, prevSig, parseSignatureLinesFromAst(entry.node).join(' · '));
          appliedCount++;
        }
      } else if ((op === 'setList' || op === 'setListItem') && entry.kind === 'list') {
        const liNodes = (entry.node.children || []).filter((c) => c.type === 'element' && c.tag === 'li');
        if (op === 'setListItem') {
          const itemIdx = Number(patch.itemIndex);
          if (liNodes[itemIdx]) {
            const prevLi = getNodePlainText(liNodes[itemIdx]).trim();
            updateInlineElementTextPreservingStyle(liNodes[itemIdx], String(patch.text ?? ''));
            recordSurgicalChange(changes, 'list', `Daftar Poin #${itemIdx + 1}`, prevLi, String(patch.text ?? ''));
            appliedCount++;
          }
        } else {
          let listChanged = false;
          if (patch.listStyle !== undefined || patch.ordered !== undefined || patch.startAt !== undefined) {
            const prevStyle = parseListNodeMeta(entry.node).listStyle;
            const targetStyle = patch.listStyle !== undefined
              ? patch.listStyle
              : (patch.ordered === false ? 'bullet' : (prevStyle === 'bullet' ? 'decimal' : prevStyle));
            if (applyListStyleToAstNode(entry.node, targetStyle, patch.startAt)) {
              listChanged = true;
              if (!Array.isArray(patch.items) || patch.items.length === 0) {
                recordSurgicalChange(
                  changes,
                  'listStyle',
                  `Gaya Penomoran Daftar Hal. ${entry.pageIndex + 1}`,
                  prevStyle,
                  `${targetStyle}${Number(patch.startAt) > 1 ? ` (mulai ${patch.startAt})` : ''}`
                );
              }
            }
          }
          if (Array.isArray(patch.items) && patch.items.length > 0) {
            const prevList = liNodes.map(li => getNodePlainText(li).trim()).join('; ');
            const tplLi = cloneAstNode(liNodes[0] || { type: 'element', tag: 'li', attrs: {}, children: [] });
            entry.node.children = patch.items.map((txt, i) => {
              const li = liNodes[i] || cloneAstNode(tplLi);
              updateInlineElementTextPreservingStyle(li, String(txt ?? ''));
              return li;
            });
            recordSurgicalChange(changes, 'list', `Daftar Poin Hal. ${entry.pageIndex + 1}`, prevList, patch.items.join('; '));
            listChanged = true;
          }
          if (listChanged) appliedCount++;
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
          recordSurgicalChange(changes, 'insertParagraph', `Sisip Paragraf Hal. ${entry.pageIndex + 1}`, '(baru)', String(patch.text));
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
          const prevTxt = getNodePlainText(entry.node).trim();
          entry.parentArray.splice(idxInParent, 1);
          recordSurgicalChange(changes, 'deleteBlock', `Hapus Paragraf Hal. ${entry.pageIndex + 1}`, prevTxt, '(dihapus)');
          appliedCount++;
        }
      }
    }

    // Jalankan konversi paragraf menjadi daftar (ul/ol) secara terurut setelah seluruh patch blok selesai
    for (const conv of pendingListConversions) {
      const valid = conv.entries.filter((e) => Array.isArray(e.parentArray) && e.parentArray.includes(e.node));
      if (valid.length === 0) continue;
      const anchor = valid[0];
      const anchorIdx = anchor.parentArray.indexOf(anchor.node);
      if (anchorIdx === -1) continue;
      const spec = getListStyleSpec(conv.listStyle, true);
      const listNode = {
        type: 'element',
        tag: spec.htmlTag,
        attrs: {},
        children: valid.map((e) => ({
          type: 'element',
          tag: 'li',
          attrs: {},
          children: Array.isArray(e.node.children) && e.node.children.length > 0
            ? e.node.children.map(cloneAstNode)
            : textToInlineAst(getNodePlainText(e.node).trim())
        }))
      };
      applyListStyleToAstNode(listNode, conv.listStyle, conv.startAt);
      anchor.parentArray.splice(anchorIdx, 1, listNode);
      for (let k = 1; k < valid.length; k++) {
        const remIdx = valid[k].parentArray.indexOf(valid[k].node);
        if (remIdx !== -1) valid[k].parentArray.splice(remIdx, 1);
      }
    }
  } else if (editedJson?.editScope === 'restructure' && allowLayoutRestructure) {
    return {
      appliedCount: 0,
      mode: 'restructure_needed',
      pages,
      meta: null,
      warnings,
      changes
    };
  } else if (Array.isArray(editedJson?.blocks)) {
    mode = 'reconciled';
    appliedCount = reconcileLegacyBlocksOntoAstPages(indexObj, editedJson, schoolProfile, warnings, changes);
  }

  const extractedPageSetup = extractPageSetupFromAiPayload(
    {
      ...(editedJson && typeof editedJson === 'object' ? editedJson : {}),
      patches: pageFormatPatches
    },
    options.letterOptions || {}
  );
  if (extractedPageSetup && pageFormatPatches.length === 0) {
    recordSurgicalChange(changes, 'pageFormat', 'Format Halaman & Tipografi', '-', Object.keys(extractedPageSetup.partial).join(', '));
    appliedCount++;
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
    warnings,
    changes,
    ...(extractedPageSetup ? { pageSetupUpdates: extractedPageSetup.partial, pageSetup: extractedPageSetup.full } : {})
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
export function buildSurgicalIndexFromDocModel(docModel = {}, options = {}) {
  const blocks = Array.isArray(docModel?.blocks) ? docModel.blocks : [];
  const blockMap = new Map();
  const orderedBlocks = [];
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
      const entry = { blockId, kind: 'identityTable', block: b, modelIndex };
      blockMap.set(blockId, entry);
      orderedBlocks.push(entry);
      indexedBlocks.push({
        blockId,
        type: 'identityTable',
        rows: (b.rows || []).map((r, rIdx) => ({ row: rIdx, label: r.label || '', value: r.value || '' }))
      });
      return;
    }
    if (b.type === 'table') {
      const blockId = `p0_t${counters.t++}`;
      const entry = { blockId, kind: 'table', block: b, modelIndex };
      blockMap.set(blockId, entry);
      orderedBlocks.push(entry);
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
      const entry = { blockId, kind: 'signature', block: b, modelIndex };
      blockMap.set(blockId, entry);
      orderedBlocks.push(entry);
      indexedBlocks.push({
        blockId,
        type: 'signature',
        lines: [`${b.city || 'Garut'}, ${b.date || ''}`, b.title || '', b.name || '', b.nip || '']
      });
      return;
    }
    if (b.type === 'list') {
      const blockId = `p0_list${counters.list++}`;
      const entry = { blockId, kind: 'list', block: b, modelIndex };
      blockMap.set(blockId, entry);
      orderedBlocks.push(entry);
      const listStyle = normalizeListStyle(b.listStyle, b.ordered !== false);
      indexedBlocks.push({
        blockId,
        type: 'list',
        ordered: listStyle !== 'bullet',
        listStyle,
        ...(Number(b.startAt) > 1 ? { startAt: Number(b.startAt) } : {}),
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
      const entry = { blockId, kind: 'paragraph', block: b, modelIndex };
      blockMap.set(blockId, entry);
      orderedBlocks.push(entry);
      indexedBlocks.push({
        blockId,
        type: b.type,
        ...(b.alignment ? { alignment: b.alignment } : {}),
        ...(b.firstLineIndentCm ? { firstLineIndentCm: b.firstLineIndentCm } : {}),
        text
      });
    }
  });

  const baseOpts = options?.letterOptions || docModel?.pageSetup || docModel?.letterData?.options || null;
  const snapshot = {
    surgicalMode: 'doc_model',
    letterTitle: titleBlock?.title || docModel?.meta?.title || 'SURAT DINAS',
    nomorSurat: String(titleBlock?.nomor || '').replace(/^Nomor\s*:\s*/i, '').trim(),
    ...(baseOpts ? { pageSetup: normalizePageSetupOptions(baseOpts) } : {}),
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

  return { blockMap, orderedBlocks, titleBlock, metaBlock, snapshot };
}

/**
 * Menerapkan patch bedah AI ke salinan bersih `DocumentModel` dengan mengunci semua tabel & tata letak.
 */
export function applySurgicalPatchesToDocModel(docModel = {}, editedJson = {}, options = {}) {
  const allowLayoutRestructure = Boolean(options.allowLayoutRestructure);
  const cloned = JSON.parse(JSON.stringify(docModel));
  const { blockMap, orderedBlocks, titleBlock, metaBlock } = buildSurgicalIndexFromDocModel(cloned, options);
  const warnings = [];
  const changes = [];
  let appliedCount = 0;
  const pageFormatPatches = [];
  const pendingListConversions = [];

  const patches = Array.isArray(editedJson?.patches) ? editedJson.patches : null;
  if (patches && patches.length > 0) {
    for (const patch of patches) {
      if (!patch || typeof patch !== 'object') continue;
      const op = String(patch.op || '').trim();

      if (op === 'setPageFormat') {
        pageFormatPatches.push(patch);
        const fmtKeys = Object.keys(patch).filter((k) => k !== 'op').join(', ') || 'halaman';
        recordSurgicalChange(changes, 'pageFormat', 'Format Halaman & Tipografi', '-', fmtKeys);
        appliedCount++;
        continue;
      }

      if (op === 'setParagraphFormat') {
        const targets = (patch.blockId === '*' || String(patch.target || '').toLowerCase() === 'all')
          ? orderedBlocks.filter((e) => e.kind === 'paragraph')
          : Array.isArray(patch.blockIds)
            ? patch.blockIds.map((id) => blockMap.get(id)).filter((e) => e && e.kind === 'paragraph')
            : (() => {
                const single = blockMap.get(patch.blockId);
                return single && single.kind === 'paragraph' ? [single] : [];
              })();
        if (targets.length === 0) {
          warnings.push(`Paragraf model target format tidak ditemukan: ${patch.blockId || '*'}`);
          continue;
        }
        let fmtCount = 0;
        for (const t of targets) {
          if (applyDocModelParagraphFormat(t.block, patch)) fmtCount++;
        }
        if (fmtCount > 0) {
          const fmtDesc = Object.keys(patch).filter((k) => !['op', 'blockId', 'blockIds', 'target'].includes(k)).join(', ');
          recordSurgicalChange(changes, 'paragraphFormat', `Format ${targets.length} Paragraf`, '-', fmtDesc);
          appliedCount++;
        }
        continue;
      }

      if (op === 'convertParagraphsToList') {
        const ids = Array.isArray(patch.blockIds) ? patch.blockIds : [];
        const entries = ids.map((id) => blockMap.get(id)).filter((e) => e && e.kind === 'paragraph');
        if (entries.length === 0) {
          warnings.push('Blok paragraf model untuk dijadikan daftar tidak ditemukan.');
          continue;
        }
        pendingListConversions.push({
          entries,
          listStyle: patch.listStyle ?? (patch.ordered === false ? 'bullet' : 'decimal'),
          startAt: patch.startAt
        });
        recordSurgicalChange(
          changes,
          'convertList',
          `Ubah ${entries.length} Paragraf Jadi Daftar (${patch.listStyle || 'decimal'})`,
          `${entries.length} paragraf`,
          `Daftar ${patch.listStyle || 'decimal'}`
        );
        appliedCount++;
        continue;
      }

      if (op === 'setField' && titleBlock) {
        const target = String(patch.target || '').toLowerCase();
        if (target === 'title' || target === 'lettertitle') {
          const prev = titleBlock.title || '';
          titleBlock.title = String(patch.text || '').toUpperCase();
          if (cloned.meta) cloned.meta.title = titleBlock.title;
          recordSurgicalChange(changes, 'field', 'Judul Surat', prev, titleBlock.title);
          appliedCount++;
        } else if (target === 'nomor' || target === 'nomorsurat') {
          const prev = titleBlock.nomor || '';
          const clean = String(patch.text || '').trim().replace(/^Nomor\s*:\s*/i, '');
          titleBlock.nomor = `Nomor: ${clean}`;
          if (cloned.letterData?.formData) cloned.letterData.formData.nomorSurat = clean;
          recordSurgicalChange(changes, 'field', 'Nomor Surat', prev, titleBlock.nomor);
          appliedCount++;
        }
        continue;
      }

      if (op === 'setMetaRecipient' && metaBlock) {
        if (Array.isArray(metaBlock.metaRows)) {
          if (patch.sifat !== undefined) {
            const row = metaBlock.metaRows.find((r) => /sifat/i.test(r.label));
            if (row) {
              recordSurgicalChange(changes, 'meta', 'Sifat Surat', row.value, patch.sifat);
              row.value = String(patch.sifat).trim();
            }
          }
          if (patch.hal !== undefined || patch.perihal !== undefined) {
            const row = metaBlock.metaRows.find((r) => /hal|perihal/i.test(r.label));
            if (row) {
              recordSurgicalChange(changes, 'meta', 'Perihal', row.value, patch.hal ?? patch.perihal);
              row.value = String(patch.hal ?? patch.perihal).trim();
            }
          }
        }
        if (Array.isArray(patch.tujuanLines)) {
          const prevT = (metaBlock.recipientLines || []).join(' ');
          metaBlock.recipientLines = patch.tujuanLines.map((l) => String(l).trim()).filter(Boolean);
          recordSurgicalChange(changes, 'meta', 'Tujuan Surat', prevT, metaBlock.recipientLines.join(' '));
          delete metaBlock.customHtml;
        } else if (patch.tujuan !== undefined) {
          const prevT = (metaBlock.recipientLines || []).join(' ');
          const tRaw = String(patch.tujuan).trim();
          metaBlock.recipientLines = ['Kepada', /^yth/i.test(tRaw) ? tRaw : `Yth. ${tRaw}`, 'Di Tempat'];
          recordSurgicalChange(changes, 'meta', 'Tujuan Surat', prevT, metaBlock.recipientLines.join(' '));
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
        const prev = entry.block.text || (entry.block.runs || []).map(r => r.text).join('');
        const val = normalizeStr(patch.text ?? '');
        entry.block.text = val;
        const firstRun = Array.isArray(entry.block.runs) && entry.block.runs[0] ? entry.block.runs[0] : {};
        entry.block.runs = [{ ...firstRun, text: val }];
        recordSurgicalChange(changes, 'paragraph', `Paragraf (${entry.blockId})`, prev, val);
        appliedCount++;
      } else if (op === 'setIdentityRow' && entry.kind === 'identityTable') {
        const row = entry.block.rows?.[Number(patch.row)];
        if (row) {
          const prev = `${row.label} ${row.value}`;
          if (patch.label !== undefined) row.label = String(patch.label).trim();
          if (patch.value !== undefined) row.value = normalizeColonValue(patch.value, row.value);
          recordSurgicalChange(changes, 'identity', `Identitas Baris ${Number(patch.row) + 1}`, prev, `${row.label} ${row.value}`);
          appliedCount++;
        }
      } else if (op === 'addIdentityRow' && entry.kind === 'identityTable') {
        entry.block.rows = entry.block.rows || [];
        const newR = { label: String(patch.label || '').trim(), value: normalizeColonValue(patch.value || '') };
        entry.block.rows.push(newR);
        recordSurgicalChange(changes, 'identity', 'Tambah Baris Identitas', '-', `${newR.label} ${newR.value}`);
        appliedCount++;
      } else if (op === 'deleteIdentityRow' && entry.kind === 'identityTable') {
        const idx = Number(patch.row);
        if (Array.isArray(entry.block.rows) && idx >= 0 && idx < entry.block.rows.length) {
          const rem = entry.block.rows.splice(idx, 1)[0];
          recordSurgicalChange(changes, 'identity', 'Hapus Baris Identitas', `${rem.label} ${rem.value}`, '(dihapus)');
          appliedCount++;
        }
      } else if (op === 'setTableFormat' && entry.kind === 'table') {
        if (patch.borderStyle !== undefined) entry.block.borderStyle = String(patch.borderStyle).trim().toLowerCase();
        if (patch.headerShading !== undefined) entry.block.headerShading = Boolean(patch.headerShading);
        if (patch.headerBold !== undefined) entry.block.headerBold = Boolean(patch.headerBold);
        if (patch.align !== undefined) entry.block.align = normalizeAlignment(patch.align) || 'left';
        if (Array.isArray(patch.colWidthsPct) && patch.colWidthsPct.length > 0) {
          entry.block.colWidthsPct = patch.colWidthsPct.map((n) => Math.max(4, Math.min(90, Number(n) || 10)));
        }
        const fmtKeys = Object.keys(patch).filter((k) => !['op', 'blockId'].includes(k)).join(', ');
        recordSurgicalChange(changes, 'tableFormat', `Format Tabel (${entry.blockId})`, '-', fmtKeys);
        appliedCount++;
      } else if (op === 'setTableCell' && entry.kind === 'table') {
        const rIdx = Number(patch.row);
        const cIdx = Number(patch.col);
        const rowEntry = entry.block.rows?.[rIdx];
        const cells = Array.isArray(rowEntry) ? rowEntry : rowEntry?.cells;
        const prevTxt = cells?.[cIdx] ? extractModelCellText(cells[cIdx]) : '';
        if (cells && updateModelCellPreservingStyle(cells, cIdx, patch.text ?? '', patch.paragraphIndex)) {
          recordSurgicalChange(changes, 'tableCell', `Sel Tabel (${entry.blockId}, Brs ${rIdx + 1} Kol ${cIdx + 1})`, prevTxt, patch.text ?? '');
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
        recordSurgicalChange(changes, 'addRow', `Tambah Baris Tabel (${entry.blockId})`, '-', newCells.join(' | '));
        appliedCount++;
      } else if (op === 'deleteTableRow' && entry.kind === 'table') {
        const idx = Number(patch.row);
        if (Array.isArray(entry.block.rows) && idx >= 0 && idx < entry.block.rows.length && entry.block.rows.length > 1) {
          entry.block.rows.splice(idx, 1);
          recordSurgicalChange(changes, 'deleteRow', `Hapus Baris (${entry.blockId}, Brs ${idx + 1})`, `Baris ${idx + 1}`, '(dihapus)');
          appliedCount++;
        }
      } else if ((op === 'setList' || op === 'setListItem') && entry.kind === 'list') {
        entry.block.items = entry.block.items || [];
        if (op === 'setListItem') {
          const itemIdx = Number(patch.itemIndex);
          if (entry.block.items[itemIdx] !== undefined) {
            const prevLi = typeof entry.block.items[itemIdx] === 'string'
              ? entry.block.items[itemIdx]
              : String(entry.block.items[itemIdx]?.text ?? '');
            const nextTxt = String(patch.text ?? '');
            entry.block.items[itemIdx] = typeof entry.block.items[itemIdx] === 'string' ? nextTxt : { text: nextTxt };
            recordSurgicalChange(changes, 'list', `Daftar Poin #${itemIdx + 1}`, prevLi, nextTxt);
            appliedCount++;
          }
        } else {
          let listChanged = false;
          if (patch.listStyle !== undefined || patch.ordered !== undefined || patch.startAt !== undefined) {
            const prevStyle = normalizeListStyle(entry.block.listStyle, entry.block.ordered !== false);
            const targetStyle = patch.listStyle !== undefined
              ? normalizeListStyle(patch.listStyle, true)
              : (patch.ordered === false ? 'bullet' : (prevStyle === 'bullet' ? 'decimal' : prevStyle));
            entry.block.listStyle = targetStyle;
            entry.block.ordered = targetStyle !== 'bullet';
            if (patch.startAt !== undefined) {
              const s = Math.max(1, parseInt(patch.startAt, 10) || 1);
              if (s > 1) entry.block.startAt = s;
              else delete entry.block.startAt;
            }
            recordSurgicalChange(changes, 'listStyle', `Gaya Penomoran Daftar`, prevStyle, targetStyle);
            listChanged = true;
          }
          if (Array.isArray(patch.items) && patch.items.length > 0) {
            entry.block.items = patch.items.map((txt) => String(txt ?? ''));
            recordSurgicalChange(changes, 'list', `Daftar Poin (${entry.blockId})`, '-', entry.block.items.join('; '));
            listChanged = true;
          }
          if (listChanged) appliedCount++;
        }
      } else if (op === 'setSignature' && entry.kind === 'signature') {
        const prevSig = `${entry.block.city || ''}, ${entry.block.date || ''} · ${entry.block.name || ''}`;
        if (patch.city !== undefined) entry.block.city = String(patch.city).trim();
        if (patch.date !== undefined) entry.block.date = String(patch.date).trim();
        if (patch.title !== undefined) entry.block.title = String(patch.title).trim();
        if (patch.name !== undefined) entry.block.name = String(patch.name).trim();
        if (patch.nip !== undefined) entry.block.nip = String(patch.nip).trim();
        const nextSig = `${entry.block.city || ''}, ${entry.block.date || ''} · ${entry.block.name || ''}`;
        recordSurgicalChange(changes, 'signature', 'Blok TTD', prevSig, nextSig);
        appliedCount++;
      } else if ((op === 'insertParagraphAfter' || op === 'insertParagraphBefore') && patch.text !== undefined) {
        const curIdx = cloned.blocks.indexOf(entry.block);
        if (curIdx !== -1) {
          const val = String(patch.text).trim();
          const newBlock = { type: 'paragraph', style: 'letterBody', text: val, runs: [{ text: val }] };
          cloned.blocks.splice(op === 'insertParagraphAfter' ? curIdx + 1 : curIdx, 0, newBlock);
          recordSurgicalChange(changes, 'insertParagraph', 'Sisip Paragraf Baru', '(baru)', val);
          appliedCount++;
        }
      } else if (op === 'deleteBlock') {
        if (entry.kind === 'table' && !allowLayoutRestructure) {
          warnings.push(`Tabel ${patch.blockId} dilindungi.`);
          continue;
        }
        const curIdx = cloned.blocks.indexOf(entry.block);
        if (curIdx !== -1) {
          const remTxt = entry.block?.text || '(blok)';
          cloned.blocks.splice(curIdx, 1);
          recordSurgicalChange(changes, 'deleteBlock', `Hapus Blok (${entry.blockId})`, remTxt, '(dihapus)');
          appliedCount++;
        }
      }
    }

    for (const conv of pendingListConversions) {
      const valid = conv.entries.filter((e) => cloned.blocks.includes(e.block));
      if (valid.length === 0) continue;
      const spec = getListStyleSpec(conv.listStyle, true);
      const startAt = Math.max(1, parseInt(conv.startAt || 1, 10) || 1);
      const listBlock = {
        type: 'list',
        ordered: spec.htmlTag === 'ol',
        listStyle: spec.key,
        ...(startAt > 1 ? { startAt } : {}),
        items: valid.map((e) =>
          Array.isArray(e.block.runs) && e.block.runs.length > 0
            ? e.block.runs.map((r) => ({ ...r }))
            : [{ text: String(e.block.text || '') }]
        )
      };
      const firstIdx = cloned.blocks.indexOf(valid[0].block);
      if (firstIdx !== -1) {
        cloned.blocks.splice(firstIdx, 1, listBlock);
        for (let k = 1; k < valid.length; k++) {
          const remIdx = cloned.blocks.indexOf(valid[k].block);
          if (remIdx !== -1) cloned.blocks.splice(remIdx, 1);
        }
      }
    }

    const baseDocOpts = options.letterOptions || cloned.pageSetup || cloned.letterData?.options || {};
    const extractedPageSetup = extractPageSetupFromAiPayload(
      { ...(editedJson && typeof editedJson === 'object' ? editedJson : {}), patches: pageFormatPatches },
      baseDocOpts
    );
    if (extractedPageSetup) {
      const full = extractedPageSetup.full;
      cloned.paperSize = full.paperSize;
      cloned.profile = full.paperSize === 'f4' ? 'suratDinasF4' : (cloned.profile || 'suratDinas');
      cloned.orientation = full.orientation;
      cloned.customMargins = full.margins;
      cloned.fontFamily = full.fontFamily;
      cloned.fontSize = full.fontSize;
      cloned.lineSpacing = full.lineSpacing;
      cloned.paragraphSpacing = full.paragraphSpacing;
      cloned.pageNumbering = full.pageNumbering;
      cloned.pageSetup = full;
      if (cloned.letterData) {
        cloned.letterData.options = { ...(cloned.letterData.options || {}), ...full };
      }
      if (pageFormatPatches.length === 0) {
        recordSurgicalChange(changes, 'pageFormat', 'Format Halaman & Tipografi', '-', Object.keys(extractedPageSetup.partial).join(', '));
        appliedCount++;
      }
    }

    return {
      appliedCount,
      mode: 'surgical',
      docModel: cloned,
      warnings,
      changes,
      ...(extractedPageSetup ? { pageSetupUpdates: extractedPageSetup.partial, pageSetup: extractedPageSetup.full } : {})
    };
  }

  const topOnlyPageSetup = extractPageSetupFromAiPayload(
    editedJson,
    options.letterOptions || cloned.pageSetup || cloned.letterData?.options || {}
  );
  if (topOnlyPageSetup && (!Array.isArray(editedJson?.blocks) || editedJson.blocks.length === 0) && editedJson?.editScope !== 'restructure') {
    const full = topOnlyPageSetup.full;
    cloned.paperSize = full.paperSize;
    cloned.orientation = full.orientation;
    cloned.customMargins = full.margins;
    cloned.fontFamily = full.fontFamily;
    cloned.fontSize = full.fontSize;
    cloned.lineSpacing = full.lineSpacing;
    cloned.paragraphSpacing = full.paragraphSpacing;
    cloned.pageNumbering = full.pageNumbering;
    cloned.pageSetup = full;
    recordSurgicalChange(changes, 'pageFormat', 'Format Halaman & Tipografi', '-', Object.keys(topOnlyPageSetup.partial).join(', '));
    return {
      appliedCount: 1,
      mode: 'surgical',
      docModel: cloned,
      warnings,
      changes,
      pageSetupUpdates: topOnlyPageSetup.partial,
      pageSetup: topOnlyPageSetup.full
    };
  }

  if (editedJson?.editScope === 'restructure' && allowLayoutRestructure) {
    return { appliedCount: 0, mode: 'restructure_needed', docModel: cloned, warnings, changes };
  }

  // Fallback reconciler untuk DocumentModel bila LLM mengirim blocks lama
  const legacyBlocks = Array.isArray(editedJson?.blocks) ? editedJson.blocks : [];
  if (titleBlock) {
    if (editedJson.letterTitle) {
      const prev = titleBlock.title || '';
      titleBlock.title = String(editedJson.letterTitle).trim().toUpperCase();
      recordSurgicalChange(changes, 'field', 'Judul Surat', prev, titleBlock.title);
      appliedCount++;
    }
    if (editedJson.nomorSurat) {
      const prev = titleBlock.nomor || '';
      const clean = String(editedJson.nomorSurat).replace(/^Nomor\s*:\s*/i, '').trim();
      titleBlock.nomor = `Nomor: ${clean}`;
      recordSurgicalChange(changes, 'field', 'Nomor Surat', prev, titleBlock.nomor);
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
        const prev = extractModelCellText(dstCells[c]);
        if (txt !== prev && updateModelCellPreservingStyle(dstCells, c, txt)) {
          recordSurgicalChange(changes, 'tableCell', `Sel Tabel (${entry.blockId}, Brs ${r + 1} Kol ${c + 1})`, prev, txt);
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
      const prev = `${entry.block.rows[r].label} ${entry.block.rows[r].value}`;
      entry.block.rows[r].label = String(src.rows[r].label || entry.block.rows[r].label).trim();
      entry.block.rows[r].value = normalizeColonValue(src.rows[r].value, entry.block.rows[r].value);
      recordSurgicalChange(changes, 'identity', `Identitas Baris ${r + 1}`, prev, `${entry.block.rows[r].label} ${entry.block.rows[r].value}`);
      appliedCount++;
    }
  });

  const existingParas = [...blockMap.values()].filter((e) => e.kind === 'paragraph');
  const legacyParas = legacyBlocks.filter((b) => b?.type === 'paragraph' && String(b.text || '').trim());
  if (existingParas.length === legacyParas.length) {
    existingParas.forEach((entry, pIdx) => {
      const val = String(legacyParas[pIdx].text).trim();
      const prev = entry.block.text || '';
      if (val && val !== prev) {
        entry.block.text = val;
        const firstRun = Array.isArray(entry.block.runs) && entry.block.runs[0] ? entry.block.runs[0] : {};
        entry.block.runs = [{ ...firstRun, text: val }];
        recordSurgicalChange(changes, 'paragraph', `Paragraf (${entry.blockId})`, prev, val);
        appliedCount++;
      }
    });
  }

  return { appliedCount, mode: 'reconciled', docModel: cloned, warnings, changes };
}
