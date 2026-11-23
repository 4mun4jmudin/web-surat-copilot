/**
 * smartWordCanvas.js - Kanvas Editor Surat Word Pro (Mode Dokumen Berkelanjutan Multi-Halaman)
 *
 * Fitur utama:
 * - Editor konten bebas (free-flow MS Word style): TIDAK membagi tiap paragraf ke dalam kotak/blok terpisah.
 *   Pengguna dapat mengetik, menyeleksi teks lintas judul/paragraf/tabel/penutup, Enter, Backspace,
 *   Ctrl+A, dan Copy-Paste persis seperti Microsoft Word asli.
 * - Multi-Halaman Dinamis: otomatis mengalirkan konten ke Halaman 2, 3, dst. bila melebihi 1 kertas
 *   (A4 / F4 / Letter / Legal), mendukung penambahan halaman manual (+ Halaman / Ctrl+Enter), hapus halaman,
 *   serta ekspor .docx (<w:br w:type="page"/>) dan Cetak PDF multi-halaman utuh.
 * - Paritas Penuh Fitur Sistem + AI Copilot: Quick AI Edit, AI Copilot Surat Sekolah (Buat Baru, Edit Terbuka,
 *   Rapihkan Dokumen, PUEBI, Tanggal Hari Ini), Gudang API & Basis Pengetahuan, Galeri Kop Surat, Gudang Surat,
 *   Buku Agenda Surat Keluar dengan penomoran otomatis berurutan, Impor .docx, dan Pengaturan Kertas/Margin.
 *
 * Catatan: Modul ini terpisah 100% dari js/ui/wordCanvas.js agar kanvas lama tetap stabil tanpa perubahan.
 */

import { store, computeDefaultDesktopZoom } from '../state.js';
import { showToast } from './notification.js';
import {
  ICON_GALLERY, ICON_EYE, ICON_EYE_OFF, ICON_UPLOAD,
  ICON_SAVE, ICON_SEND_OUT, ICON_SPARKLE, ICON_WAND,
  ICON_CONVERT, ICON_SETTINGS, ICON_RULER, ICON_SHADING,
  ICON_ALIGN_LEFT, ICON_ALIGN_CENTER, ICON_ALIGN_RIGHT, ICON_ALIGN_JUSTIFY,
  ICON_LIST_BULLET, ICON_LIST_NUMBER, ICON_OUTDENT, ICON_INDENT,
  ICON_ERASER, ICON_TABLE, ICON_UNDO, ICON_REDO, ICON_PILCROW,
  ICON_GLOBE, ICON_MAXIMIZE, ICON_CASE, ICON_PORTRAIT,
  ICON_MAIL, ICON_FOLDER, ICON_BUILDING, ICON_ID_CARD,
  ICON_BRIEFCASE, ICON_USERS, ICON_CALENDAR, ICON_TARGET,
  ICON_PLUS, ICON_BOT_SPARK, ICON_OPEN_BOOK, ICON_PENCIL_NOTE,
  ICON_CLIPBOARD_CHECK, ICON_PAGE_ADD, ICON_COPY_DOC, ICON_DOC_UPLOAD,
  ICON_REFRESH
} from './icons.js';
import { parseDocxForEditing, convertDocxToDocumentModel } from '../docx/patch.js';
import { DEFAULT_KOP_DATA_URL } from '../letters/defaultKop.js';
import { generateLetterNumber, formatSmartNomorSurat, detectLetterTypeCode, LETTER_TYPE_CODES } from '../letters/numbering.js';
import { extractDocumentMetadata, buildDocumentDownloadFilename } from '../letters/smartNaming.js';
import {
  getSchoolCodeFromProfile,
  detectTargetUnitFromText,
  resolveDocumentSchoolProfile,
  resolveSchoolProfileForUnit,
  adaptDocumentModelForUnit
} from '../letters/schoolProfile.js';
import { callAiWithValidation, polishMessyDocumentWithAi, editDocumentWithAi, convertAiBlocksToDocumentBlocks } from '../ai/client.js';
import { COPILOT_SYSTEM_PROMPT, buildCopilotUserPrompt } from '../ai/prompts.js';
import { summarizeKnowledgeBaseStats, describeAiMatchFeedback } from '../ai/knowledgeBase.js';
import { renderCanvasVaultBadge, openLockedMasterSaveChoiceModal, buildMetaAutoScanBarHtml, wireMetaAutoScanBar } from './vaultSaveHelper.js';
import { mountCompactChatDrawer } from './aiChatWorkspace.js';
import { renderDocx } from '../docx/render.js';
import { downloadFile } from '../utils/download.js';
import { applySignatureModeToDocModel, buildSignerSelectionInfo, formatSignerNipLine } from '../letters/verification.js';
import { canSignWithRole, resolveDefaultAllowedSignerRole } from '../auth/permissions.js';
import {
  openPrintSignatureChoiceModal,
  buildSignerRoleSelectorHtml,
  authorizeAndRecordOutgoing,
  requestAccForOutgoing,
  printWithTemporaryQrOverlay,
  syncGlobalPrintPageStyle
} from './printChoiceModal.js';
import { buildSuratTugasSppdDocModel } from '../letters/templates/suratTugasSppd.js';
import { attachInteractiveTables } from './tableInteraction.js';
import {
  computePrintableDimsMm,
  mountInteractiveFloatingQr,
  updateCanvasQrToggleButton,
  normalizeQrPlacement,
  findPreferredSignatureQrPlacement
} from './qrFloatingHelper.js';
import {
  escapeHtmlText,
  decodeHtmlEntities,
  parseHtmlNodes,
  getNodePlainText
} from '../utils/htmlAst.js';
import {
  shouldAllowLayoutRestructure,
  buildSurgicalIndexFromSmartPages,
  applySurgicalPatchesToSmartPages
} from '../ai/surgicalEdit.js';
import {
  SUPPORTED_PAPER_SIZES,
  SUPPORTED_FONT_FAMILIES,
  LIST_STYLE_CONFIG,
  PAGE_NUMBERING_DEFAULT,
  cssLenToCm,
  normalizeAlignment,
  normalizeFontFamily,
  normalizeListStyle,
  getListStyleSpec,
  parseListNodeMeta,
  formatPageNumberLabel,
  normalizePageNumbering,
  normalizePageSetupOptions,
  extractPageSetupFromAiPayload
} from '../letters/wordFormat.js';
import { requestAccCloud } from '../sync/workspaceClient.js';
import { setLastKnownSyncRev } from '../sync/liveSync.js';

const SMART_ICON_SEARCH = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>`;
const SMART_ICON_HISTORY = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>`;

export { escapeHtmlText, decodeHtmlEntities, parseHtmlNodes };

// =========================================================================
// BAGIAN 1: PARSER HTML RINGAN & KONVERTER HTML <-> DOCUMENT MODEL (PURE JS)
// =========================================================================

function detectCssAlignment(node) {
  const style = (node?.attrs?.style || '').toLowerCase();
  const alignAttr = (node?.attrs?.align || '').toLowerCase();
  const m = style.match(/text-align\s*:\s*(left|center|right|justify)/);
  const val = m ? m[1] : alignAttr;
  if (val === 'center') return 'center';
  if (val === 'right') return 'right';
  if (val === 'justify') return 'both';
  if (val === 'left') return 'left';
  return '';
}

function detectCssVerticalAlignment(node) {
  const style = (node?.attrs?.style || '').toLowerCase();
  const valignAttr = (node?.attrs?.valign || '').toLowerCase();
  const m = style.match(/vertical-align\s*:\s*(top|middle|bottom|center)/);
  const val = m ? m[1] : valignAttr;
  if (val === 'middle' || val === 'center') return 'center';
  if (val === 'bottom') return 'bottom';
  if (val === 'top') return 'top';
  return '';
}

function parseSpacingTwips(node, side = 'after') {
  const dataKey = side === 'before' ? 'data-space-before' : 'data-space-after';
  const rawAttr = node?.attrs?.[dataKey];
  if (rawAttr !== undefined && rawAttr !== '') {
    const n = Number(rawAttr);
    if (!Number.isNaN(n) && n >= 0) return Math.round(n);
  }
  const style = (node?.attrs?.style || '').toLowerCase();
  const prop = side === 'before' ? 'margin-top' : 'margin-bottom';
  const m = style.match(new RegExp(`${prop}\\s*:\\s*([\\d.]+)\\s*(pt|px)`));
  if (!m) return undefined;
  const val = parseFloat(m[1]);
  if (Number.isNaN(val) || val < 0) return undefined;
  return m[2] === 'pt' ? Math.round(val * 20) : Math.round(val * 15);
}

function parseCssFontSizeHalfPt(node) {
  const style = (node?.attrs?.style || '').toLowerCase();
  const m = style.match(/font-size\s*:\s*([\d.]+)\s*pt/);
  if (!m) return undefined;
  const pt = parseFloat(m[1]);
  return Number.isNaN(pt) ? undefined : Math.round(pt * 2);
}

function normalizePreservedHtmlText(text = '', preserveNewlines = false) {
  const raw = String(text ?? '');
  if (preserveNewlines) {
    return raw.replace(/\r\n/g, '\n');
  }
  // Hanya ubah newline + indentasi sumber HTML antar baris menjadi spasi tunggal;
  // pertahankan spasi berulang ('   '), non-breaking space (\u00a0), dan tab (\t) milik pengguna.
  return raw.replace(/\r?\n[ \t]*/g, ' ');
}

function formatRunTextToHtml(text = '') {
  return escapeHtmlText(text)
    .replace(/ (?= )/g, '&nbsp;')
    .replace(/^ /, '&nbsp;')
    .replace(/\u00a0/g, '&nbsp;')
    .replace(/\r?\n/g, '<br>');
}

function extractInlineRuns(nodes, inherited = {}, opts = {}) {
  const runs = [];
  const preserveNewlines = Boolean(opts?.preserveNewlines);

  for (const node of nodes || []) {
    if (node.type === 'text') {
      const clean = normalizePreservedHtmlText(node.text, preserveNewlines);
      if (clean) {
        runs.push({
          text: clean,
          ...(inherited.bold ? { bold: true } : {}),
          ...(inherited.italic ? { italic: true } : {}),
          ...(inherited.underline ? { underline: true } : {}),
          ...(inherited.sup ? { sup: true } : {}),
          ...(inherited.sub ? { sub: true } : {}),
          ...(inherited.fontSizeHalfPt ? { fontSizeHalfPt: inherited.fontSizeHalfPt } : {}),
          ...(inherited.color ? { color: inherited.color } : {})
        });
      }
      continue;
    }

    if (node.type !== 'element') continue;
    if (node.tag === 'br') {
      runs.push({
        text: preserveNewlines ? '\n' : ' ',
        ...(inherited.bold ? { bold: true } : {}),
        ...(inherited.italic ? { italic: true } : {}),
        ...(inherited.underline ? { underline: true } : {}),
        ...(inherited.fontSizeHalfPt ? { fontSizeHalfPt: inherited.fontSizeHalfPt } : {}),
        ...(inherited.color ? { color: inherited.color } : {})
      });
      continue;
    }

    const style = (node.attrs?.style || '').toLowerCase();
    const fontHalf = parseCssFontSizeHalfPt(node);
    const colorMatch = style.match(/(?:^|;)\s*color\s*:\s*#([0-9a-f]{3,8})/i);
    const nextFmt = {
      bold: inherited.bold || node.tag === 'b' || node.tag === 'strong' || /font-weight\s*:\s*(bold|700|800|900)/.test(style),
      italic: inherited.italic || node.tag === 'i' || node.tag === 'em' || /font-style\s*:\s*italic/.test(style),
      underline: inherited.underline || node.tag === 'u' || /text-decoration[^;]*underline/.test(style),
      sup: inherited.sup || node.tag === 'sup',
      sub: inherited.sub || node.tag === 'sub',
      fontSizeHalfPt: fontHalf || inherited.fontSizeHalfPt,
      color: colorMatch ? colorMatch[1].toUpperCase() : inherited.color
    };

    const inner = extractInlineRuns(node.children, nextFmt, opts);
    runs.push(...inner);
  }

  // Merge adjacent runs with identical formatting cleanly
  const merged = [];
  for (const r of runs) {
    const prev = merged[merged.length - 1];
    if (
      prev &&
      Boolean(prev.bold) === Boolean(r.bold) &&
      Boolean(prev.italic) === Boolean(r.italic) &&
      Boolean(prev.underline) === Boolean(r.underline) &&
      Boolean(prev.sup) === Boolean(r.sup) &&
      Boolean(prev.sub) === Boolean(r.sub) &&
      (prev.fontSizeHalfPt || 0) === (r.fontSizeHalfPt || 0) &&
      (prev.color || '') === (r.color || '')
    ) {
      prev.text += r.text;
    } else {
      merged.push({ ...r });
    }
  }
  return merged;
}

function parseCellParagraphChild(pNode) {
  const align = detectCssAlignment(pNode);
  const spacingBefore = parseSpacingTwips(pNode, 'before');
  const spacingAfter = parseSpacingTwips(pNode, 'after');
  if (pNode?.attrs?.['data-qr-slot'] === 'true') {
    return {
      text: '',
      align: align || 'center',
      qrSlot: true,
      qrCaption: pNode.attrs?.['data-qr-caption'] || 'Disahkan secara elektronik',
      ...(spacingBefore !== undefined ? { spacingBefore } : {}),
      ...(spacingAfter !== undefined ? { spacingAfter } : {})
    };
  }
  const runs = extractInlineRuns(pNode.children, {}, { preserveNewlines: true });
  const rawText = getNodePlainText(pNode).replace(/\r\n/g, '\n').replace(/^[\r\n]+|[\r\n]+$/g, '');
  const nonEmptyRuns = runs.filter(r => (r.text || '').trim());
  const allBold = nonEmptyRuns.length > 0 && nonEmptyRuns.every(r => r.bold);
  const allItalic = nonEmptyRuns.length > 0 && nonEmptyRuns.every(r => r.italic);
  const allUnderline = nonEmptyRuns.length > 0 && nonEmptyRuns.every(r => r.underline);
  const firstFont = nonEmptyRuns[0]?.fontSizeHalfPt;
  const uniformFont = firstFont && nonEmptyRuns.every(r => r.fontSizeHalfPt === firstFont) ? firstFont : undefined;

  const res = {
    text: rawText,
    ...(align ? { align } : {}),
    ...(allBold ? { bold: true } : {}),
    ...(allItalic ? { italic: true } : {}),
    ...(allUnderline ? { underline: true } : {}),
    ...(uniformFont ? { fontSizeHalfPt: uniformFont } : {}),
    ...(spacingBefore !== undefined ? { spacingBefore } : {}),
    ...(spacingAfter !== undefined ? { spacingAfter } : {})
  };
  if (nonEmptyRuns.length > 1 && !(allBold && allItalic === Boolean(nonEmptyRuns[0]?.italic) && allUnderline === Boolean(nonEmptyRuns[0]?.underline))) {
    res.runs = runs;
  }
  return res;
}

function extractRichCellData(cellNode) {
  const rawNodeText = getNodePlainText(cellNode);
  const plainText = rawNodeText.trim()
    ? rawNodeText.replace(/\r?\n[ \t]*/g, ' ').replace(/^[\r\n]+|[\r\n]+$/g, '')
    : '';
  const multilineText = rawNodeText.trim()
    ? rawNodeText.replace(/\r\n/g, '\n').replace(/^[\r\n]+|[\r\n]+$/g, '')
    : '';
  const colspan = Math.max(1, parseInt(cellNode.attrs?.colspan || '1', 10) || 1);
  const rowspan = Math.max(1, parseInt(cellNode.attrs?.rowspan || '1', 10) || 1);
  const align = detectCssAlignment(cellNode);
  const vAlign = detectCssVerticalAlignment(cellNode);
  const qrSlot = cellNode.attrs?.['data-qr-slot'] === 'true';

  const blockChildren = (cellNode.children || []).filter(ch => ch.type === 'element' && (ch.tag === 'p' || ch.tag === 'div'));
  if (blockChildren.length > 0) {
    const paragraphs = blockChildren.map(parseCellParagraphChild);
    const modelObj = {
      paragraphs,
      ...(colspan > 1 ? { colspan } : {}),
      ...(rowspan > 1 ? { rowspan } : {}),
      ...(vAlign ? { vAlign } : {}),
      ...(qrSlot ? { qrSlot: true } : {})
    };
    return { plainText, isRich: true, modelObj };
  }

  const runs = extractInlineRuns(cellNode.children, {}, { preserveNewlines: true });
  const nonEmptyRuns = runs.filter(r => (r.text || '').trim());
  const hasInlineFmt = nonEmptyRuns.some(r => r.bold || r.italic || r.underline || r.fontSizeHalfPt || r.color);
  const hasMultiline = multilineText.includes('\n');
  const isRich = colspan > 1 || rowspan > 1 || Boolean(align) || Boolean(vAlign) || qrSlot || hasInlineFmt || hasMultiline;

  if (!isRich) {
    return { plainText, isRich: false, modelObj: plainText };
  }

  const allBold = nonEmptyRuns.length > 0 && nonEmptyRuns.every(r => r.bold);
  const allItalic = nonEmptyRuns.length > 0 && nonEmptyRuns.every(r => r.italic);
  const allUnderline = nonEmptyRuns.length > 0 && nonEmptyRuns.every(r => r.underline);
  const modelObj = {
    text: multilineText || plainText,
    ...(colspan > 1 ? { colspan } : {}),
    ...(rowspan > 1 ? { rowspan } : {}),
    ...(align ? { align } : {}),
    ...(vAlign ? { vAlign } : {}),
    ...(allBold ? { bold: true } : {}),
    ...(allItalic ? { italic: true } : {}),
    ...(allUnderline ? { underline: true } : {}),
    ...(qrSlot ? { qrSlot: true } : {})
  };
  if (hasInlineFmt && nonEmptyRuns.length > 1) {
    modelObj.paragraphs = multilineText.split(/\r?\n/).map(line => ({ text: line, ...(align ? { align } : {}), ...(allBold ? { bold: true } : {}) }));
  }
  return { plainText, isRich: true, modelObj };
}

function parseTableColWidthsPct(tableNode) {
  const attr = (tableNode?.attrs?.['data-col-widths'] || '').trim();
  if (attr) {
    const parsed = attr.split(',').map(v => parseFloat(v.trim())).filter(n => !Number.isNaN(n) && n > 0);
    if (parsed.length > 0) return parsed;
  }
  const colgroup = (tableNode?.children || []).find(c => c.type === 'element' && c.tag === 'colgroup');
  if (colgroup) {
    const cols = (colgroup.children || []).filter(c => c.type === 'element' && c.tag === 'col');
    const widths = cols.map(c => {
      const m = (c.attrs?.style || '').match(/width\s*:\s*([\d.]+)\s*%/i);
      return m ? parseFloat(m[1]) : 0;
    });
    if (widths.length > 0 && widths.every(w => w > 0)) return widths;
  }
  return [];
}

function collectTableRows(tableNode) {
  const rows = [];
  const walk = (n, inThead = false) => {
    if (!n || n.type !== 'element') return;
    if (n.tag === 'tr') {
      const cellNodes = (n.children || []).filter(c => c.type === 'element' && (c.tag === 'td' || c.tag === 'th'));
      const style = (n.attrs?.style || '').toLowerCase();
      let minHeightCm = parseFloat(n.attrs?.['data-min-height'] || '0') || 0;
      if (!minHeightCm) {
        const mh = style.match(/(?:min-)?height\s*:\s*([\d.]+)\s*(cm|px)/i);
        if (mh) {
          minHeightCm = mh[2] === 'cm' ? parseFloat(mh[1]) : Math.round((parseFloat(mh[1]) / 37.795) * 100) / 100;
        }
      }
      const parsedCells = cellNodes.map(extractRichCellData);
      rows.push({
        isHeader: n.attrs?.['data-header'] === 'true' || inThead || (cellNodes.length > 0 && cellNodes.every(c => c.tag === 'th')),
        minHeightCm,
        cells: parsedCells.map(pc => pc.plainText),
        richCells: parsedCells
      });
      return;
    }
    const nextInThead = inThead || n.tag === 'thead';
    (n.children || []).forEach(child => walk(child, nextInThead));
  };
  walk(tableNode, false);
  return rows;
}

function parseSignatureCellLines(sigTableNode) {
  const rows = [];
  const walk = (n) => {
    if (!n || n.type !== 'element') return;
    if (n.tag === 'tr') {
      const cells = (n.children || []).filter(c => c.type === 'element' && (c.tag === 'td' || c.tag === 'th'));
      const sigCell = cells[cells.length - 1];
      if (sigCell) {
        const htmlText = collectLinesFromCell(sigCell);
        rows.push(...htmlText);
      }
      return;
    }
    (n.children || []).forEach(walk);
  };
  walk(sigTableNode);
  return rows.map(l => l.trim()).filter(Boolean);
}

function collectLinesFromCell(cellNode) {
  const lines = [];
  let buf = '';
  const walk = (n) => {
    if (n.type === 'text') {
      buf += n.text;
      return;
    }
    if (n.type !== 'element') return;
    if (n.tag === 'br') {
      lines.push(buf);
      buf = '';
      return;
    }
    if (['p', 'div'].includes(n.tag)) {
      if (buf.trim()) {
        lines.push(buf);
        buf = '';
      }
      (n.children || []).forEach(walk);
      lines.push(buf);
      buf = '';
      return;
    }
    (n.children || []).forEach(walk);
  };
  (cellNode.children || []).forEach(walk);
  if (buf) lines.push(buf);
  return lines;
}

const CONTAINER_BLOCK_TAGS = new Set(['table', 'p', 'div', 'h1', 'h2', 'h3', 'h4', 'ul', 'ol', 'blockquote']);

function hasNestedBlockChild(node) {
  return (node?.children || []).some(ch => ch?.type === 'element' && (CONTAINER_BLOCK_TAGS.has(ch.tag) || hasNestedBlockChild(ch)));
}

/**
 * Mengonversi HTML satu halaman Smart Canvas menjadi daftar blok DocumentModel terstruktur.
 */
export function parseSmartPageHtmlToBlocks(html = '', schoolProfile = {}) {
  const blocks = [];

  const processBlockNodes = (rawNodes = []) => {
    const nodes = (rawNodes || []).filter(n => !(n.type === 'text' && !n.text.trim()));
    for (let idx = 0; idx < nodes.length; idx++) {
      const node = nodes[idx];

      if (node.type === 'text') {
        if (node.text.trim()) {
          const preserved = normalizePreservedHtmlText(node.text, false);
          blocks.push({ type: 'paragraph', style: 'letterBody', runs: [{ text: preserved }] });
        }
        continue;
      }

      if (node.type !== 'element') continue;
      const role = node.attrs?.['data-smart-role'] || '';

      // 1. Judul + Nomor (+ opsional Meta Recipient bila tepat mengikuti)
      if (role === 'title') {
        const titleText = getNodePlainText(node).replace(/\s+/g, ' ').trim();
        let nomorText = '';
        let metaRecipient = null;

        const nextNode = nodes[idx + 1];
        if (nextNode?.type === 'element' && nextNode.attrs?.['data-smart-role'] === 'nomor') {
          nomorText = getNodePlainText(nextNode).replace(/\s+/g, ' ').trim();
          idx++;
        }

        const afterNomorNode = nodes[idx + 1];
        if (afterNomorNode?.type === 'element' && afterNomorNode.attrs?.['data-smart-role'] === 'meta-recipient') {
          metaRecipient = parseMetaRecipientTable(afterNomorNode);
          idx++;
        }

        blocks.push({
          type: 'letterTitle',
          title: titleText,
          nomor: nomorText,
          ...(metaRecipient ? { metaRecipient } : {})
        });
        continue;
      }

      if (role === 'nomor') {
        const nomorText = getNodePlainText(node).replace(/\s+/g, ' ').trim();
        if (nomorText) {
          blocks.push({
            type: 'paragraph',
            style: 'letterBody',
            alignment: 'center',
            runs: [{ text: nomorText }]
          });
        }
        continue;
      }

      if (role === 'meta-recipient') {
        blocks.push({
          type: 'letterMetaRecipient',
          ...parseMetaRecipientTable(node)
        });
        continue;
      }

      // 2. Tabel (Identity Table / Signature Table / Bordered Word Grid Table)
      if (node.tag === 'table') {
        if (role === 'signature') {
          const lines = parseSignatureCellLines(node);
          const sigBlock = parseSignatureLinesToBlock(lines, schoolProfile);
          blocks.push(sigBlock);
          continue;
        }

        const rows = collectTableRows(node);
        if (rows.length === 0) continue;

        if (role === 'identity-table') {
          const idColWidths = parseTableColWidthsPct(node);
          blocks.push({
            type: 'identityTable',
            ...(idColWidths.length >= 2 ? { colWidthsPct: idColWidths } : {}),
            rows: rows.map(r => ({
              label: r.cells[0] || '',
              value: r.cells[1] !== undefined ? r.cells[1] : ''
            }))
          });
          continue;
        }

        const classAttr = (node.attrs?.class || '').toLowerCase();
        const borderStyle = node.attrs?.['data-border-style']
          || (classAttr.includes('smart-table-sppd') ? 'sppd' : classAttr.includes('smart-table-none') ? 'none' : 'grid');
        const headerShading = node.attrs?.['data-header-shading'] === 'false' || classAttr.includes('no-header-bg')
          ? false
          : (node.attrs?.['data-header-shading'] === 'true' ? true : undefined);
        const headerBold = node.attrs?.['data-header-bold'] === 'false' || classAttr.includes('no-header-bold')
          ? false
          : (node.attrs?.['data-header-bold'] === 'true' ? true : undefined);
        const colWidthsPct = parseTableColWidthsPct(node);
        const tableWidthPct = parseFloat(node.attrs?.['data-table-width-pct'] || '0') || 0;

        const hasRichRowMeta = rows.some((r, rIdx) => r.minHeightCm > 0 || (r.isHeader && rIdx > 0));
        const hasRichCells = rows.some(r => r.richCells.some(c => c.isRich));
        const isRichTable = borderStyle !== 'grid' || headerShading !== undefined || headerBold !== undefined || colWidthsPct.length > 0 || (tableWidthPct > 0 && tableWidthPct < 100) || hasRichRowMeta || hasRichCells;

        if (!isRichTable) {
          blocks.push({
            type: 'table',
            headerRow: rows[0]?.isHeader || false,
            rows: rows.map(r => r.cells)
          });
        } else {
          const formattedRows = rows.map(r => {
            const cellList = r.richCells.map(c => c.modelObj);
            if (hasRichRowMeta || r.minHeightCm > 0) {
              return {
                isHeader: Boolean(r.isHeader),
                ...(r.minHeightCm > 0 ? { minHeightCm: r.minHeightCm } : {}),
                cells: cellList
              };
            }
            return cellList;
          });
          blocks.push({
            type: 'table',
            borderStyle,
            headerRow: rows.some(r => r.isHeader),
            ...(headerShading !== undefined ? { headerShading } : {}),
            ...(headerBold !== undefined ? { headerBold } : {}),
            ...(colWidthsPct.length > 0 ? { colWidthsPct } : {}),
            ...(tableWidthPct > 0 && tableWidthPct < 100 ? { tableWidthPct } : {}),
            rows: formattedRows
          });
        }
        continue;
      }

      // 3. List (<ul> / <ol>)
      if (node.tag === 'ul' || node.tag === 'ol') {
        const listMeta = parseListNodeMeta(node);
        const listStyleAttr = String(node.attrs?.style || '');
        const listLeftMatch = listStyleAttr.match(/(?:^|;)\s*(?:margin-left|padding-left)\s*:\s*([^;]+)/i);
        const listLeftCm = listLeftMatch ? cssLenToCm(listLeftMatch[1]) : undefined;
        const listLineMatch = listStyleAttr.match(/(?:^|;)\s*line-height\s*:\s*([\d.]+)/i);
        const listLineSpacing = listLineMatch ? parseFloat(listLineMatch[1]) : undefined;
        const items = (node.children || [])
          .filter(c => c.type === 'element' && c.tag === 'li')
          .map(li => {
            const runs = extractInlineRuns(li.children);
            return runs.length > 0 ? runs : [{ text: getNodePlainText(li).trim() }];
          })
          .filter(item => item.some(r => r.text.trim()));

        if (items.length > 0) {
          blocks.push({
            type: 'list',
            ordered: listMeta.ordered,
            listStyle: listMeta.listStyle,
            ...(listMeta.startAt > 1 ? { startAt: listMeta.startAt } : {}),
            ...(listLeftCm && listLeftCm > 0 ? { leftIndentCm: listLeftCm } : {}),
            ...(Number.isFinite(listLineSpacing) ? { lineSpacing: listLineSpacing } : {}),
            items
          });
        }
        continue;
      }

      // 4. Headings (<h1>, <h2>, <h3>)
      if (['h1', 'h2', 'h3', 'h4'].includes(node.tag)) {
        const txtCheck = getNodePlainText(node).replace(/\u00a0/g, ' ').trim();
        if (!txtCheck) continue;
        const align = detectCssAlignment(node) || 'left';
        const hFontHalf = parseCssFontSizeHalfPt(node);
        const hRuns = extractInlineRuns(node.children, { bold: true, ...(hFontHalf ? { fontSizeHalfPt: hFontHalf } : {}) });
        blocks.push({
          type: 'paragraph',
          style: 'letterBody',
          alignment: align,
          runs: hRuns.length > 0 ? hRuns : [{ text: txtCheck, bold: true }]
        });
        continue;
      }

      // 4B. Container div pembungkus blok (<div ...><table>...</table></div>): buka isinya secara rekursif agar tabel tidak rata jadi teks
      if ((node.tag === 'div' || node.tag === 'blockquote' || node.tag === 'section') && hasNestedBlockChild(node)) {
        processBlockNodes(node.children || []);
        continue;
      }

      // 5. Paragraph / Leaf Div standar MS Word
      if (node.tag === 'p' || node.tag === 'div' || node.tag === 'blockquote') {
        const spacingBefore = parseSpacingTwips(node, 'before');
        const spacingAfter = parseSpacingTwips(node, 'after');
        const pFontHalf = parseCssFontSizeHalfPt(node);
        const pStyleStr = String(node.attrs?.style || '');
        const indentMatch = pStyleStr.match(/(?:^|;)\s*text-indent\s*:\s*([^;]+)/i);
        const leftMatch = pStyleStr.match(/(?:^|;)\s*(?:padding-left|margin-left)\s*:\s*([^;]+)/i);
        const lineMatch = pStyleStr.match(/(?:^|;)\s*line-height\s*:\s*([\d.]+)/i);
        const fontFamMatch = pStyleStr.match(/(?:^|;)\s*font-family\s*:\s*([^;]+)/i);
        const firstLineIndentCm = indentMatch ? cssLenToCm(indentMatch[1]) : undefined;
        const leftIndentCm = leftMatch ? cssLenToCm(leftMatch[1]) : undefined;
        const pLineSpacing = lineMatch ? parseFloat(lineMatch[1]) : undefined;
        const pFontFamily = fontFamMatch ? normalizeFontFamily(fontFamMatch[1], '') : '';
        const pBold = /font-weight\s*:\s*(?:bold|700)/i.test(pStyleStr);
        const pItalic = /font-style\s*:\s*italic/i.test(pStyleStr);
        const pUnderline = /text-decoration[^:]*:\s*[^;]*underline/i.test(pStyleStr);
        const rawTxt = getNodePlainText(node).replace(/\u00a0/g, ' ').trim();
        if (!rawTxt) {
          blocks.push({
            type: 'paragraph',
            style: 'letterBody',
            ...(spacingBefore !== undefined ? { spacingBefore } : {}),
            ...(spacingAfter !== undefined ? { spacingAfter } : {}),
            runs: [{ text: '' }]
          });
          continue;
        }
        const inherited = {
          ...(pFontHalf ? { fontSizeHalfPt: pFontHalf } : {}),
          ...(pFontFamily ? { fontFamily: pFontFamily } : {}),
          ...(pBold ? { bold: true } : {}),
          ...(pItalic ? { italic: true } : {}),
          ...(pUnderline ? { underline: true } : {})
        };
        const runs = extractInlineRuns(node.children, inherited);
        const alignment = detectCssAlignment(node);
        blocks.push({
          type: 'paragraph',
          style: 'letterBody',
          ...(alignment ? { alignment } : {}),
          ...(firstLineIndentCm !== undefined && firstLineIndentCm > 0 ? { firstLineIndentCm } : {}),
          ...(leftIndentCm !== undefined && leftIndentCm > 0 ? { leftIndentCm } : {}),
          ...(Number.isFinite(pLineSpacing) && pLineSpacing >= 0.8 ? { lineSpacing: pLineSpacing } : {}),
          ...(pFontFamily ? { fontFamily: pFontFamily } : {}),
          ...(pFontHalf ? { fontSizePt: Math.round((pFontHalf / 2) * 10) / 10 } : {}),
          ...(spacingBefore !== undefined ? { spacingBefore } : {}),
          ...(spacingAfter !== undefined ? { spacingAfter } : {}),
          runs: runs.length > 0 ? runs : [{ text: rawTxt, ...inherited }]
        });
      }
    }
  };

  processBlockNodes(parseHtmlNodes(html));
  return blocks;
}

function parseMetaRecipientTable(tableNode) {
  const rows = collectTableRows(tableNode);
  const metaRows = [];
  const recipientLines = [];

  for (const r of rows) {
    const leftText = (r.cells[0] || '').trim();
    const rightText = (r.cells[1] || '').trim();
    if (leftText) {
      // Bentuk "Sifat : Penting" atau 2/3 kolom
      const colonIdx = leftText.indexOf(':');
      if (colonIdx > 0) {
        metaRows.push({
          label: leftText.slice(0, colonIdx).trim(),
          value: leftText.slice(colonIdx + 1).trim()
        });
      } else if (r.cells.length >= 2 && (r.cells[1] || '').trim().startsWith(':')) {
        metaRows.push({
          label: leftText,
          value: (r.cells[1] || '').replace(/^:\s*/, '').trim()
        });
      } else {
        metaRows.push({ label: leftText, value: '' });
      }
    }
    if (rightText && r.cells.length <= 2) {
      recipientLines.push(rightText);
    } else if (r.cells[r.cells.length - 1] && r.cells.length > 2) {
      recipientLines.push(r.cells[r.cells.length - 1].trim());
    }
  }

  // Pisahkan isi sel bila memakai newline / multi baris dalam satu <td>
  const expandedRecipient = [];
  const rowList = [];
  const walkTr = (n) => {
    if (!n || n.type !== 'element') return;
    if (n.tag === 'tr') {
      rowList.push(n);
      return;
    }
    (n.children || []).forEach(walkTr);
  };
  walkTr(tableNode);

  if (rowList.length > 0) {
    const tr = rowList[0];
    const tds = (tr.children || []).filter(c => c.type === 'element' && (c.tag === 'td' || c.tag === 'th'));
    if (tds.length >= 1) {
      const leftLines = collectLinesFromCell(tds[0]).map(l => l.trim()).filter(Boolean);
      metaRows.length = 0;
      for (const line of leftLines) {
        const cIdx = line.indexOf(':');
        if (cIdx > 0) {
          metaRows.push({ label: line.slice(0, cIdx).trim(), value: line.slice(cIdx + 1).trim() });
        } else {
          metaRows.push({ label: line, value: '' });
        }
      }
    }
    if (tds.length >= 2) {
      const rightLines = collectLinesFromCell(tds[tds.length - 1]).map(l => l.trim()).filter(Boolean);
      expandedRecipient.push(...rightLines);
    }
  }

  return {
    metaRows,
    recipientLines: expandedRecipient.length > 0 ? expandedRecipient : recipientLines
  };
}

function parseSignatureLinesToBlock(lines = [], schoolProfile = {}) {
  const firstLine = lines[0] || '';
  let city = schoolProfile.city || 'Garut';
  let date = '';
  const cityDateMatch = firstLine.match(/^([^,]+),\s*(.+)$/);
  if (cityDateMatch) {
    city = cityDateMatch[1].trim();
    date = cityDateMatch[2].trim();
  } else if (firstLine) {
    date = firstLine;
  }

  const title = lines[1] || schoolProfile.principalTitle || 'Kepala Sekolah';
  let name = schoolProfile.principalName || 'Dian Nurhamdani, S.Pd.';
  let nip = schoolProfile.principalNip || '-';

  const nipLine = lines.find(l => /^nip\b/i.test(l));
  if (nipLine) {
    nip = nipLine.replace(/^nip\.?\s*:?\s*/i, '').trim() || '-';
  }

  const candidates = lines.slice(2).filter(l => !/^nip\b/i.test(l) && !/^\s*$/.test(l));
  if (candidates.length > 0) {
    name = candidates[candidates.length - 1];
  }

  return { city, date, title, name, nip, type: 'letterSignature' };
}

/**
 * Menggabungkan seluruh halaman Smart Word Canvas menjadi satu DocumentModel utuh untuk diekspor ke .docx.
 */
export function buildDocModelFromSmartPagesData({
  pages = [],
  schoolProfile = {},
  letterOptions = {},
  meta = {}
} = {}) {
  const profileName = letterOptions.paperSize === 'f4' ? 'suratDinasF4' : 'suratDinas';
  const showKop = letterOptions.showKop !== false;
  const blocks = [];

  if (showKop) {
    blocks.push({ type: 'letterKop', school: schoolProfile });
  }

  let extractedTitle = meta.title || 'Surat Dinas Sekolah';
  let extractedNomor = meta.nomorSurat || '';

  pages.forEach((page, pageIdx) => {
    if (pageIdx > 0) {
      blocks.push({ type: 'pageBreak' });
    }
    const pageBlocks = parseSmartPageHtmlToBlocks(page?.html || '', schoolProfile);
    for (const b of pageBlocks) {
      if (b.type === 'letterTitle') {
        if (b.title) extractedTitle = b.title;
        if (b.nomor) extractedNomor = b.nomor.replace(/^Nomor\s*:\s*/i, '').trim();
      }
      blocks.push(b);
    }
  });

  const normOpts = normalizePageSetupOptions(letterOptions);

  return {
    meta: {
      title: extractedTitle,
      author: schoolProfile.schoolName || 'SMK IT Al-Hawari',
      year: new Date().getFullYear().toString()
    },
    profile: profileName,
    paperSize: normOpts.paperSize,
    orientation: normOpts.orientation,
    customMargins: normOpts.margins,
    fontFamily: normOpts.fontFamily,
    fontSize: normOpts.fontSize,
    lineSpacing: normOpts.lineSpacing,
    paragraphSpacing: normOpts.paragraphSpacing,
    pageNumbering: normOpts.pageNumbering,
    pageSetup: normOpts,
    letterData: {
      templateId: 'smart_word_pro',
      schoolProfile,
      formData: {
        judulSurat: extractedTitle,
        nomorSurat: extractedNomor
      },
      options: normOpts
    },
    blocks
  };
}

function runsToHtml(runs = []) {
  if (!Array.isArray(runs) || runs.length === 0) return '';
  return runs.map((r) => {
    let html = formatRunTextToHtml(r.text || '');
    if (r.bold) html = `<strong>${html}</strong>`;
    if (r.italic) html = `<em>${html}</em>`;
    if (r.underline) html = `<u>${html}</u>`;
    if (r.sup) html = `<sup>${html}</sup>`;
    if (r.sub) html = `<sub>${html}</sub>`;
    const styleParts = [];
    if (r.fontFamily) {
      styleParts.push(`font-family: '${escapeHtmlText(r.fontFamily)}', serif`);
    }
    if (r.fontSizeHalfPt && r.fontSizeHalfPt !== 24) {
      styleParts.push(`font-size: ${Math.round((r.fontSizeHalfPt / 2) * 10) / 10}pt`);
    }
    if (r.color) {
      styleParts.push(`color: #${String(r.color).replace(/^#/, '')}`);
    }
    if (styleParts.length > 0) {
      html = `<span style="${styleParts.join('; ')}">${html}</span>`;
    }
    return html;
  }).join('');
}

function renderTableCellParagraphHtml(p = {}) {
  const spParts = [];
  if (p.align) spParts.push(`text-align: ${p.align === 'both' ? 'justify' : p.align};`);
  if (p.spacingBefore !== undefined && p.spacingBefore > 0) {
    spParts.push(`margin-top: ${(p.spacingBefore / 20).toFixed(1)}pt;`);
  }
  if (p.spacingAfter !== undefined && p.spacingAfter > 0) {
    spParts.push(`margin-bottom: ${(p.spacingAfter / 20).toFixed(1)}pt;`);
  }
  const dataAttrs = `${p.spacingBefore !== undefined ? ` data-space-before="${p.spacingBefore}"` : ''}${p.spacingAfter !== undefined ? ` data-space-after="${p.spacingAfter}"` : ''}`;

  if (p.qrSlot) {
    return `<p class="smart-qr-slot" data-qr-slot="true"${p.qrCaption ? ` data-qr-caption="${escapeHtmlText(p.qrCaption)}"` : ''}${dataAttrs} style="${spParts.join(' ') || 'text-align: center;'}"></p>`;
  }

  const inner = Array.isArray(p.runs) && p.runs.length > 0
    ? runsToHtml(p.runs)
    : runsToHtml([{
        text: p.text ?? '',
        bold: p.bold,
        italic: p.italic,
        underline: p.underline,
        fontSizeHalfPt: p.fontSizeHalfPt,
        color: p.color
      }]);
  const styleAttr = spParts.length > 0 ? ` style="${spParts.join(' ')}"` : '';
  return `<p${dataAttrs}${styleAttr}>${inner || '<br>'}</p>`;
}

function renderRichTableCellHtml(cell, tag = 'td') {
  if (cell && typeof cell === 'object' && !Array.isArray(cell)) {
    const attrs = [];
    if (cell.colspan > 1) attrs.push(`colspan="${cell.colspan}"`);
    if (cell.rowspan > 1) attrs.push(`rowspan="${cell.rowspan}"`);
    if (cell.qrSlot) attrs.push(`data-qr-slot="true"`);
    if (cell.qrCaption) attrs.push(`data-qr-caption="${escapeHtmlText(cell.qrCaption)}"`);

    const styles = [];
    if (cell.align) styles.push(`text-align: ${cell.align === 'both' ? 'justify' : cell.align};`);
    if (cell.vAlign) styles.push(`vertical-align: ${cell.vAlign === 'center' ? 'middle' : cell.vAlign};`);
    if (styles.length > 0) attrs.push(`style="${styles.join(' ')}"`);
    const attrStr = attrs.length > 0 ? ` ${attrs.join(' ')}` : '';

    if (Array.isArray(cell.paragraphs) && cell.paragraphs.length > 0) {
      return `<${tag}${attrStr}>${cell.paragraphs.map(renderTableCellParagraphHtml).join('')}</${tag}>`;
    }

    const inner = Array.isArray(cell.runs) && cell.runs.length > 0
      ? runsToHtml(cell.runs)
      : runsToHtml([{
          text: cell.text ?? '',
          bold: cell.bold,
          italic: cell.italic,
          underline: cell.underline,
          fontSizeHalfPt: cell.fontSizeHalfPt,
          color: cell.color
        }]);
    return `<${tag}${attrStr}>${inner}</${tag}>`;
  }
  return `<${tag}>${formatRunTextToHtml(cell ?? '')}</${tag}>`;
}

function renderTableBlockToHtml(b = {}) {
  const rows = b.rows || [];
  const borderStyle = b.borderStyle || 'grid';
  const classes = ['smart-word-table'];
  if (borderStyle === 'sppd') classes.push('smart-table-sppd');
  if (borderStyle === 'none') classes.push('smart-table-none');
  if (b.headerShading === false) classes.push('no-header-bg');
  if (b.headerBold === false) classes.push('no-header-bold');

  const attrs = [`class="${classes.join(' ')}"`];
  if (borderStyle !== 'grid') attrs.push(`data-border-style="${borderStyle}"`);
  if (b.headerShading === false) attrs.push('data-header-shading="false"');
  if (b.headerBold === false) attrs.push('data-header-bold="false"');
  if (Array.isArray(b.colWidthsPct) && b.colWidthsPct.length > 0) {
    attrs.push(`data-col-widths="${b.colWidthsPct.join(',')}"`);
  }
  if (b.tableWidthPct && Number(b.tableWidthPct) >= 30 && Number(b.tableWidthPct) < 100) {
    attrs.push(`data-table-width-pct="${b.tableWidthPct}"`);
    attrs.push(`style="width: ${b.tableWidthPct}%;"`);
  }

  const colgroupHtml = Array.isArray(b.colWidthsPct) && b.colWidthsPct.length > 0
    ? `<colgroup>${b.colWidthsPct.map(w => `<col style="width: ${w}%;">`).join('')}</colgroup>`
    : '';

  const headRows = [];
  const bodyRows = [];
  let reachedBody = false;

  rows.forEach((rowEntry, idx) => {
    const rowObj = (rowEntry && typeof rowEntry === 'object' && !Array.isArray(rowEntry))
      ? {
          cells: rowEntry.cells || [],
          isHeader: 'isHeader' in rowEntry ? Boolean(rowEntry.isHeader) : (b.headerRow !== false && idx === 0),
          minHeightCm: Number(rowEntry.minHeightCm) || 0
        }
      : {
          cells: Array.isArray(rowEntry) ? rowEntry : [],
          isHeader: b.headerRow !== false && idx === 0,
          minHeightCm: 0
        };

    if (!reachedBody && rowObj.isHeader) {
      headRows.push(rowObj);
    } else {
      reachedBody = true;
      bodyRows.push(rowObj);
    }
  });

  const renderRowHtml = (rowObj, tag) => {
    const rAttrs = [];
    if (rowObj.isHeader) rAttrs.push('data-header="true"');
    if (rowObj.minHeightCm > 0) {
      rAttrs.push(`data-min-height="${rowObj.minHeightCm}"`);
      rAttrs.push(`style="height: ${rowObj.minHeightCm}cm;"`);
    }
    const rAttrStr = rAttrs.length > 0 ? ` ${rAttrs.join(' ')}` : '';
    return `<tr${rAttrStr}>${(rowObj.cells || []).map(c => renderRichTableCellHtml(c, tag)).join('')}</tr>`;
  };

  const headHtml = headRows.length > 0
    ? `<thead>${headRows.map(r => renderRowHtml(r, 'th')).join('')}</thead>`
    : '';
  const bodyHtml = bodyRows.length > 0
    ? `<tbody>${bodyRows.map(r => renderRowHtml(r, 'td')).join('')}</tbody>`
    : '';

  return `<table ${attrs.join(' ')}>${colgroupHtml}${headHtml}${bodyHtml}</table>`;
}

/**
 * Mengonversi DocumentModel (baik bawaan, hasil AI, atau hasil impor .docx) ke daftar halaman Smart HTML.
 */
export function docModelToSmartPages(docModel, schoolProfile = {}) {
  const blocks = docModel?.blocks || [];
  const pages = [{ id: `spage-${Date.now()}-1`, html: '', isManualBreak: false }];
  let currentParts = [];

  const flushPage = (isManual = false) => {
    pages[pages.length - 1].html = currentParts.join('\n') || '<p><br></p>';
    currentParts = [];
    pages.push({ id: `spage-${Date.now()}-${pages.length + 1}`, html: '', isManualBreak: isManual });
  };

  for (const b of blocks) {
    if (!b || b.type === 'letterKop') continue;

    if (b.type === 'pageBreak') {
      flushPage(true);
      continue;
    }

    if (b.type === 'letterTitle') {
      const titleText = b.title || 'SURAT DINAS SEKOLAH';
      const nomorText = b.nomor || 'Nomor: -';
      currentParts.push(`<h2 class="smart-doc-title" data-smart-role="title" style="text-align: center;">${escapeHtmlText(titleText)}</h2>`);
      currentParts.push(`<p class="smart-doc-nomor" data-smart-role="nomor" style="text-align: center;">${escapeHtmlText(nomorText)}</p>`);

      if (b.metaRecipient) {
        currentParts.push(renderMetaRecipientTableHtml(b.metaRecipient));
      }
      continue;
    }

    if (b.type === 'letterMetaRecipient') {
      currentParts.push(renderMetaRecipientTableHtml(b));
      continue;
    }

    if (b.type === 'identityTable') {
      const hasCustomW = Array.isArray(b.colWidthsPct) && b.colWidthsPct.length >= 2;
      const w0 = hasCustomW ? b.colWidthsPct[0] : 30;
      const w1 = hasCustomW ? b.colWidthsPct[1] : 70;
      const colAttr = hasCustomW ? ` data-col-widths="${w0},${w1}"` : '';
      const colGroup = hasCustomW ? `<colgroup><col style="width: ${w0}%;"><col style="width: ${w1}%;"></colgroup>` : '';
      const rowsHtml = (b.rows || []).map(r => {
        const rawLabel = String(r.label ?? '').replace(/^[\r\n]+|[\r\n]+$/g, '');
        const rawVal = String(r.value ?? '').replace(/^[\r\n]+|[\r\n]+$/g, '');
        const trimmedVal = rawVal.trimStart();
        const normVal = trimmedVal.startsWith(':')
          ? rawVal
          : (trimmedVal ? `: ${rawVal}` : ': ');
        return `<tr><td style="width: ${w0}%;">${formatRunTextToHtml(rawLabel)}</td><td>${formatRunTextToHtml(normVal)}</td></tr>`;
      }).join('');
      currentParts.push(`<table class="smart-identity-table" data-smart-role="identity-table"${colAttr}>${colGroup}<tbody>${rowsHtml}</tbody></table>`);
      continue;
    }

    if (b.type === 'table') {
      currentParts.push(renderTableBlockToHtml(b));
      continue;
    }

    if (b.type === 'list') {
      const spec = getListStyleSpec(b.listStyle, b.ordered !== false);
      const startNum = Math.max(1, parseInt(b.startAt || 1, 10) || 1);
      const typeAttr = spec.htmlType ? ` type="${spec.htmlType}"` : '';
      const startAttr = spec.htmlTag === 'ol' && startNum > 1 ? ` start="${startNum}"` : '';
      const listStyles = [`list-style-type: ${spec.cssType}`];
      if (b.leftIndentCm && Number(b.leftIndentCm) > 0) listStyles.push(`margin-left: ${Number(b.leftIndentCm)}cm`);
      if (b.lineSpacing) listStyles.push(`line-height: ${b.lineSpacing}`);
      const liHtml = (b.items || []).map(item => {
        if (typeof item === 'string') return `<li>${formatRunTextToHtml(item)}</li>`;
        if (Array.isArray(item)) return `<li>${runsToHtml(item)}</li>`;
        return `<li>${formatRunTextToHtml(item?.text || '')}</li>`;
      }).join('');
      currentParts.push(`<${spec.htmlTag} data-list-style="${spec.key}"${typeAttr}${startAttr} style="${listStyles.join('; ')}">${liHtml}</${spec.htmlTag}>`);
      continue;
    }

    if (b.type === 'heading') {
      const lvl = Math.min(Math.max(Number(b.level) || 2, 1), 3);
      const inner = b.runs ? runsToHtml(b.runs) : formatRunTextToHtml(b.text || '');
      currentParts.push(`<h${lvl}>${inner}</h${lvl}>`);
      continue;
    }

    if (b.type === 'letterSignature') {
      currentParts.push(renderSignatureTableHtml(b, schoolProfile));
      continue;
    }

    if (b.type === 'paragraph') {
      const inner = b.runs ? runsToHtml(b.runs) : formatRunTextToHtml(b.text || '');
      const styleParts = [];
      if (b.alignment) {
        styleParts.push(`text-align: ${b.alignment === 'both' ? 'justify' : b.alignment};`);
      }
      if (b.firstLineIndentCm !== undefined && Number(b.firstLineIndentCm) > 0) {
        styleParts.push(`text-indent: ${Number(b.firstLineIndentCm)}cm;`);
      }
      if (b.leftIndentCm !== undefined && Number(b.leftIndentCm) > 0) {
        styleParts.push(`padding-left: ${Number(b.leftIndentCm)}cm;`);
      }
      if (b.lineSpacing !== undefined && Number(b.lineSpacing) > 0) {
        styleParts.push(`line-height: ${Number(b.lineSpacing)};`);
      }
      if (b.fontFamily) {
        styleParts.push(`font-family: '${escapeHtmlText(b.fontFamily)}', serif;`);
      }
      if (b.fontSizePt) {
        styleParts.push(`font-size: ${Number(b.fontSizePt)}pt;`);
      }
      if (b.bold !== undefined) {
        styleParts.push(`font-weight: ${b.bold ? 'bold' : 'normal'};`);
      }
      if (b.italic !== undefined) {
        styleParts.push(`font-style: ${b.italic ? 'italic' : 'normal'};`);
      }
      if (b.underline !== undefined) {
        styleParts.push(`text-decoration: ${b.underline ? 'underline' : 'none'};`);
      }
      if (b.spacingBefore !== undefined && b.spacingBefore > 0) {
        styleParts.push(`margin-top: ${(b.spacingBefore / 20).toFixed(1)}pt;`);
      }
      if (b.spacingAfter !== undefined && b.spacingAfter > 0) {
        styleParts.push(`margin-bottom: ${(b.spacingAfter / 20).toFixed(1)}pt;`);
      }
      const dataAttrs = `${b.spacingBefore !== undefined ? ` data-space-before="${b.spacingBefore}"` : ''}${b.spacingAfter !== undefined ? ` data-space-after="${b.spacingAfter}"` : ''}`;
      const styleAttr = styleParts.length > 0 ? ` style="${styleParts.join(' ')}"` : '';
      currentParts.push(`<p${dataAttrs}${styleAttr}>${inner || '<br>'}</p>`);
    }
  }

  pages[pages.length - 1].html = currentParts.join('\n') || '<p><br></p>';

  return {
    pages,
    meta: {
      title: docModel?.meta?.title || 'Surat Dinas Sekolah',
      nomorSurat: docModel?.letterData?.formData?.nomorSurat || ''
    }
  };
}

function renderMetaRecipientTableHtml(mr = {}) {
  const leftLines = (mr.metaRows || [])
    .map(m => `${escapeHtmlText(m.label || '')} : ${escapeHtmlText(m.value || '')}`)
    .join('<br>');
  const rightLines = (mr.recipientLines || [])
    .map(l => escapeHtmlText(l))
    .join('<br>');

  return `
<table class="smart-meta-table" data-smart-role="meta-recipient">
  <tbody>
    <tr>
      <td style="width: 55%;">${leftLines || 'Sifat : Penting<br>Perihal : Surat Dinas'}</td>
      <td style="width: 45%;">${rightLines || 'Kepada<br>Yth. Bapak/Ibu<br>Di Tempat'}</td>
    </tr>
  </tbody>
</table>`.trim();
}

export function renderSignatureTableHtml(sig = {}, school = {}) {
  const city = sig.city || school.city || 'Garut';
  const date = sig.date || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
  const title = sig.title || school.principalTitle || 'Kepala Sekolah';
  const name = sig.name || school.principalName || 'Dian Nurhamdani, S.Pd.';
  const nipRaw = sig.nip && sig.nip !== '-' ? sig.nip : (school.principalNip || '-');
  const nipLine = `<br>${escapeHtmlText(formatSignerNipLine(nipRaw))}`;

  return `
<table class="smart-sig-table" data-smart-role="signature">
  <tbody>
    <tr>
      <td class="smart-sig-spacer"></td>
      <td class="smart-sig-cell">
        ${escapeHtmlText(city)}, ${escapeHtmlText(date)}<br>
        ${escapeHtmlText(title)}<br><br><br><br>
        <strong><u>${escapeHtmlText(name)}</u></strong>${nipLine}
      </td>
    </tr>
  </tbody>
</table>`.trim();
}

/**
 * Mengekstrak snapshot JSON dari halaman Smart Word Canvas untuk dikirim ke `editDocumentWithAi`.
 */
export function extractSmartSnapshotFromPages(pages = [], schoolProfile = {}) {
  const docModel = buildDocModelFromSmartPagesData({ pages, schoolProfile });
  let letterTitle = 'SURAT DINAS SEKOLAH';
  let nomorSurat = '';
  let hideMetaRecipient = true;
  let sifat = 'Penting';
  let hal = '';
  let tujuan = '';
  const bodyBlocks = [];
  let signer = {
    city: schoolProfile.city || 'Garut',
    date: new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
    title: schoolProfile.principalTitle || 'Kepala Sekolah',
    name: schoolProfile.principalName || 'Dian Nurhamdani, S.Pd.',
    nip: schoolProfile.principalNip || '-'
  };

  for (const b of docModel.blocks) {
    if (b.type === 'letterTitle') {
      letterTitle = b.title || letterTitle;
      nomorSurat = (b.nomor || '').replace(/^Nomor\s*:\s*/i, '').trim();
      if (b.metaRecipient) {
        hideMetaRecipient = false;
        const sifatRow = (b.metaRecipient.metaRows || []).find(r => /sifat/i.test(r.label));
        const halRow = (b.metaRecipient.metaRows || []).find(r => /hal|perihal/i.test(r.label));
        if (sifatRow) sifat = sifatRow.value;
        if (halRow) hal = halRow.value;
        tujuan = (b.metaRecipient.recipientLines || []).join(' ');
      }
    } else if (b.type === 'letterMetaRecipient') {
      hideMetaRecipient = false;
      const sifatRow = (b.metaRows || []).find(r => /sifat/i.test(r.label));
      const halRow = (b.metaRows || []).find(r => /hal|perihal/i.test(r.label));
      if (sifatRow) sifat = sifatRow.value;
      if (halRow) hal = halRow.value;
      tujuan = (b.recipientLines || []).join(' ');
    } else if (b.type === 'identityTable') {
      bodyBlocks.push({ type: 'identityTable', rows: b.rows || [] });
    } else if (b.type === 'table') {
      bodyBlocks.push({
        type: 'table',
        borderStyle: b.borderStyle || 'grid',
        headerRow: b.headerRow !== false,
        ...(b.headerShading !== undefined ? { headerShading: b.headerShading } : {}),
        ...(b.headerBold !== undefined ? { headerBold: b.headerBold } : {}),
        ...(Array.isArray(b.colWidthsPct) ? { colWidthsPct: b.colWidthsPct } : {}),
        rows: b.rows || []
      });
    } else if (b.type === 'pageBreak') {
      bodyBlocks.push({ type: 'pageBreak' });
    } else if (b.type === 'paragraph') {
      const txt = (b.runs || []).map(r => r.text).join('').trim();
      if (txt) {
        bodyBlocks.push({
          type: 'paragraph',
          text: txt,
          ...(b.alignment ? { alignment: b.alignment } : {})
        });
      }
    } else if (b.type === 'list') {
      const itemsText = (b.items || []).map(it => (Array.isArray(it) ? it.map(r => r.text).join('') : String(it)));
      itemsText.forEach((it, idx) => bodyBlocks.push({ type: 'paragraph', text: `${idx + 1}. ${it}` }));
    } else if (b.type === 'letterSignature') {
      signer = {
        city: b.city || signer.city,
        date: b.date || signer.date,
        title: b.title || signer.title,
        name: b.name || signer.name,
        nip: b.nip || signer.nip
      };
    }
  }

  return {
    letterTitle,
    nomorSurat,
    hideMetaRecipient,
    sifat,
    hal,
    tujuan,
    blocks: bodyBlocks,
    signer
  };
}

/**
 * Mengonversi respons JSON dari AI Copilot (`COPILOT_SYSTEM_PROMPT`) atau AI Editor (`EDITOR_SYSTEM_PROMPT`)
 * menjadi DocumentModel lengkap siap ditampilkan di Smart Word Canvas.
 */
export function aiJsonToSmartDocModel(parsedJson = {}, schoolProfile = {}, letterOptions = {}) {
  const extracted = extractPageSetupFromAiPayload(parsedJson, letterOptions);
  const resolvedOpts = extracted ? extracted.full : normalizePageSetupOptions(letterOptions);

  if (parsedJson.templatePreset === 'surat_tugas_sppd') {
    const sppdModel = buildSuratTugasSppdDocModel(
      parsedJson.formData || parsedJson,
      schoolProfile,
      {
        ...resolvedOpts,
        paperSize: extracted?.partial?.paperSize || parsedJson.paperSize || resolvedOpts.paperSize || 'f4'
      }
    );
    sppdModel.pageSetup = normalizePageSetupOptions(
      { ...resolvedOpts, paperSize: sppdModel.paperSize || resolvedOpts.paperSize },
      resolvedOpts
    );
    sppdModel.fontFamily = sppdModel.pageSetup.fontFamily;
    sppdModel.fontSize = sppdModel.pageSetup.fontSize;
    sppdModel.lineSpacing = sppdModel.pageSetup.lineSpacing;
    sppdModel.paragraphSpacing = sppdModel.pageSetup.paragraphSpacing;
    sppdModel.pageNumbering = sppdModel.pageSetup.pageNumbering;
    return sppdModel;
  }

  const title = (parsedJson.letterTitle || parsedJson.title || 'SURAT DINAS SEKOLAH').trim();
  const contextText = `${title} ${parsedJson.hal || parsedJson.perihal || ''}`;
  const rawNomor = (parsedJson.nomorSurat || '').trim();
  const unitSchoolCode = getSchoolCodeFromProfile(schoolProfile);
  const nomorFinal = rawNomor
    ? formatSmartNomorSurat({
        rawNomor,
        contextText,
        forcedTypeCode: parsedJson.kodeJenisSurat || '',
        schoolCode: unitSchoolCode
      })
    : generateLetterNumber({
        typeCode: detectLetterTypeCode(contextText),
        sequence: store.getNextOutgoingSequence()
      });

  const blocks = [{ type: 'letterKop', school: schoolProfile }];
  const hideMeta = parsedJson.hideMetaRecipient === true;

  const metaRecipient = hideMeta
    ? null
    : {
        metaRows: [
          { label: 'Sifat', value: parsedJson.sifat || 'Penting' },
          ...(parsedJson.lampiran && parsedJson.lampiran !== '-' ? [{ label: 'Lampiran', value: parsedJson.lampiran }] : []),
          { label: 'Perihal', value: parsedJson.hal || parsedJson.perihal || title }
        ],
        recipientLines: [
          'Kepada',
          parsedJson.tujuan
            ? (parsedJson.tujuan.toLowerCase().startsWith('yth') ? parsedJson.tujuan : `Yth. ${parsedJson.tujuan}`)
            : 'Yth. Bapak/Ibu',
          'Di Tempat'
        ]
      };

  blocks.push(
    {
      type: 'letterTitle',
      title: title.toUpperCase(),
      nomor: `Nomor: ${nomorFinal}`,
      ...(metaRecipient ? { metaRecipient } : {})
    },
    ...convertAiBlocksToDocumentBlocks(parsedJson.blocks)
  );

  if (parsedJson.hideSignature !== true) {
    const sig = parsedJson.signer || {};
    blocks.push({
      type: 'letterSignature',
      city: sig.city || schoolProfile.city || 'Garut',
      date: sig.date || parsedJson.tanggalSurat || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
      title: sig.title || schoolProfile.principalTitle || 'Kepala Sekolah',
      name: sig.name || schoolProfile.principalName || 'Dian Nurhamdani, S.Pd.',
      nip: sig.nip && sig.nip !== '-' ? sig.nip : (schoolProfile.principalNip || '-')
    });
  }

  if (Array.isArray(parsedJson.tembusan) && parsedJson.tembusan.length > 0) {
    blocks.push({
      type: 'paragraph',
      style: 'letterBody',
      alignment: 'left',
      runs: [{ text: 'Tembusan:', bold: true }]
    });
    parsedJson.tembusan.forEach((t) => {
      blocks.push({
        type: 'paragraph',
        style: 'letterBody',
        alignment: 'left',
        runs: [{ text: String(t) }]
      });
    });
  }

  const resolvedPaper = resolvedOpts.paperSize || 'a4';

  return {
    meta: {
      title,
      author: schoolProfile.schoolName || 'SMK IT Al-Hawari',
      year: new Date().getFullYear().toString()
    },
    profile: resolvedPaper === 'f4' ? 'suratDinasF4' : 'suratDinas',
    paperSize: resolvedPaper,
    orientation: resolvedOpts.orientation,
    customMargins: resolvedOpts.margins,
    fontFamily: resolvedOpts.fontFamily,
    fontSize: resolvedOpts.fontSize,
    lineSpacing: resolvedOpts.lineSpacing,
    paragraphSpacing: resolvedOpts.paragraphSpacing,
    pageNumbering: resolvedOpts.pageNumbering,
    pageSetup: resolvedOpts,
    ...(extracted ? { extractedPageSetup: extracted.partial } : {}),
    letterData: {
      templateId: 'smart_word_ai',
      schoolProfile,
      formData: {
        judulSurat: title,
        nomorSurat: nomorFinal
      },
      options: resolvedOpts
    },
    blocks
  };
}

// =========================================================================
// BAGIAN 2: BRIDGE GETTER UNTUK HEADER DOWNLOAD GLOBAL & STATE HELPERS
// =========================================================================

let activeLivePagesExtractor = null;
let activeSmartAuthorizeHandler = null;
let activeSmartRequestAccHandler = null;
let activeSmartPagesRoot = null;

/**
 * Mengambil DocumentModel terkini dari Kanvas Word Pro (dipanggil header Unduh .docx atau Simpan).
 */
export function getActiveSmartDocModel() {
  const state = store.getState();
  const school = state.smartDocument?.schoolProfile
    || resolveDocumentSchoolProfile(state.smartDocument?.pendingDocModel || state.smartDocument, store.getSchoolProfile(), state.kopGallery || []);
  const opts = state.letterOptions || {};
  const pages = activeLivePagesExtractor
    ? activeLivePagesExtractor()
    : (state.smartDocument?.pages?.length
        ? state.smartDocument.pages
        : state.smartDocument?.pendingDocModel
        ? docModelToSmartPages(state.smartDocument.pendingDocModel, school).pages
        : docModelToSmartPages(state.document, school).pages);

  const model = buildDocModelFromSmartPagesData({
    pages,
    schoolProfile: school,
    letterOptions: opts,
    meta: state.smartDocument?.meta || {}
  });
  model.targetUnitId = state.smartDocument?.targetUnitId || school?.unitId || store.getActiveUnitId();
  if (state.smartDocument?.verification) {
    model.verification = state.smartDocument.verification;
  }
  if (state.smartDocument?.qrPlacement) {
    model.qrPlacement = state.smartDocument.qrPlacement;
  }
  return model;
}

/**
 * Helper global Cetak / PDF untuk Kanvas Word Pro dengan Pilihan Barcode QR vs Kosongan TTD Manual.
 */
export function triggerSmartCanvasPrintWithChoice() {
  const smartDoc = store.getState().smartDocument;
  const verification = smartDoc?.verification || null;
  const isRevoked = verification?.verificationStatus === 'revoked';
  const isPendingAcc = verification?.verificationStatus === 'pending_acc' || verification?.accStatus === 'pending';
  const isRejectedAcc = verification?.verificationStatus === 'rejected';
  const isVerifiedSah = Boolean(verification?.isVerifiedSah && !isRevoked && verification?.qrDataUrl);
  const model = getActiveSmartDocModel();
  const titleText = model.letterData?.formData?.judulSurat || model.meta?.title || 'SURAT DINAS';
  const nomorText = model.letterData?.formData?.nomorSurat || verification?.nomorSurat || '-';

  const runSmartPrint = (mode, qrDataUrl = '', overrideRole = null, overrideVerif = null) => {
    const stackEl = activeSmartPagesRoot || document.getElementById('smart-pages-stack');
    const sheetEls = stackEl ? Array.from(stackEl.querySelectorAll('.smart-paper-sheet')) : [];
    const school = store.getEffectiveSchoolForDocument(model);
    const v = overrideVerif || verification;
    const effectiveRole = overrideRole || v?.signerRole || 'kepala_sekolah';
    const latestSmartDoc = store.getState().smartDocument;

    const printOpts = store.getLetterOptions();
    printWithTemporaryQrOverlay({
      mode,
      pageSheetEls: sheetEls,
      qrPlacement: latestSmartDoc?.qrPlacement || null,
      docModel: model,
      qrDataUrl,
      signerRole: effectiveRole,
      principalName: school.principalName || '',
      signerName: v?.signerName || '',
      school,
      paperSize: printOpts.paperSize,
      orientation: printOpts.orientation,
      margins: printOpts.margins
    });
  };

  openPrintSignatureChoiceModal({
    actionType: 'print',
    isVerifiedSah,
    isRevoked,
    isPendingAcc,
    isRejectedAcc,
    rejectedReason: verification?.rejectedReason || '',
    nomorSurat: nomorText,
    title: titleText,
    qrDataUrl: verification?.qrDataUrl || '',
    defaultSignerRole: verification?.signerRole || 'kepala_sekolah',
    signerLabel: verification?.roleLabel || verification?.signerTitle || 'Kepala Sekolah',
    signerName: verification?.signerName || '',
    school: store.getEffectiveSchoolForDocument(model),
    onChooseQr: () => runSmartPrint('qr', verification.qrDataUrl, verification?.signerRole || 'kepala_sekolah'),
    onChooseManual: () => runSmartPrint('manual'),
    onAuthorizeAndQr: async (selectedRole) => {
      if (typeof activeSmartAuthorizeHandler !== 'function') return;
      const result = await activeSmartAuthorizeHandler(selectedRole, false);
      runSmartPrint('qr', result.qrDataUrl, selectedRole, result.payload);
    },
    onRequestAcc: async (selectedRole) => {
      if (typeof activeSmartRequestAccHandler !== 'function') return;
      const res = await activeSmartRequestAccHandler(selectedRole, false);
      showToast(`⏳ Surat #${String(res.recorded.sequenceNumber).padStart(3, '0')} masuk Antrean ACC Kepala Sekolah!`, 'info');
    }
  });
}

/**
 * Helper global Unduh .docx untuk Kanvas Word Pro dengan Pilihan Barcode QR vs Kosongan TTD Manual.
 */
export function triggerSmartCanvasDownloadWithChoice() {
  const smartDoc = store.getState().smartDocument;
  const verification = smartDoc?.verification || null;
  const isRevoked = verification?.verificationStatus === 'revoked';
  const isPendingAcc = verification?.verificationStatus === 'pending_acc' || verification?.accStatus === 'pending';
  const isRejectedAcc = verification?.verificationStatus === 'rejected';
  const isVerifiedSah = Boolean(verification?.isVerifiedSah && !isRevoked && verification?.qrDataUrl);
  const model = getActiveSmartDocModel();
  const titleText = model.letterData?.formData?.judulSurat || model.meta?.title || 'SURAT DINAS';
  const nomorText = model.letterData?.formData?.nomorSurat || verification?.nomorSurat || '-';

  const runSmartDownload = async (mode, overrideVerif = null) => {
    const v = overrideVerif || verification;
    const liveModel = getActiveSmartDocModel();
    const school = store.getEffectiveSchoolForDocument(liveModel);
    const signerInfo = v ? buildSignerSelectionInfo(school, v.signerRole) : null;
    const hasQrSlotInTable = liveModel.blocks.some(b =>
      b.type === 'table' && (b.rows || []).some(r =>
        (Array.isArray(r) ? r : r?.cells || []).some(c =>
          c?.qrSlot || (Array.isArray(c?.paragraphs) && c.paragraphs.some(p => p?.qrSlot))
        )
      )
    );

    if (mode === 'qr' && v?.qrDataUrl && !liveModel.blocks.some(b => b.type === 'letterSignature') && !hasQrSlotInTable) {
      liveModel.blocks.push({
        type: 'letterSignature',
        city: school.city || 'Garut',
        date: new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
        title: signerInfo?.signerTitle || school.principalTitle || 'Kepala Sekolah',
        name: signerInfo?.signerName || school.principalName || 'Dian Nurhamdani, S.Pd.',
        nip: signerInfo?.signerNip ?? (school.principalNip || '-'),
        rank: signerInfo?.signerRank || ''
      });
    }

    const finalModel = applySignatureModeToDocModel(liveModel, {
      mode,
      qrDataUrl: mode === 'qr' ? (v?.qrDataUrl || '') : '',
      verificationUrl: mode === 'qr' ? (v?.verificationUrl || '') : '',
      signerRole: v?.signerRole || 'kepala_sekolah',
      signerInfo,
      school,
      syncSignerFields: false
    });

    showToast(`Menyusun berkas Word Pro .docx (${mode === 'qr' ? 'dengan Barcode QR' : 'kosongan TTD manual'})...`, 'info');
    try {
      const blob = await renderDocx(finalModel, 'blob');
      const dlMeta = extractDocumentMetadata({
        docModel: finalModel,
        smartPages: store.getState().smartDocument?.pages || [],
        schoolProfile: school
      });
      const fileName = buildDocumentDownloadFilename({
        kind: v?.outgoingId ? 'outgoing' : 'vault',
        sequenceNumber: v?.sequenceNumber || 0,
        typeCode: dlMeta.typeCode,
        title: dlMeta.formattedTitle,
        perihal: dlMeta.perihal,
        recipientName: dlMeta.recipientName,
        tanggalSurat: dlMeta.tanggalSurat,
        unitId: school?.unitId,
        schoolProfile: school
      });
      downloadFile(blob, fileName);
      showToast(`Berkas ${fileName} berhasil diunduh.`, 'success');
    } catch (err) {
      showToast(`Gagal mengunduh: ${err.message}`, 'error');
    }
  };

  openPrintSignatureChoiceModal({
    actionType: 'docx',
    isVerifiedSah,
    isRevoked,
    isPendingAcc,
    isRejectedAcc,
    rejectedReason: verification?.rejectedReason || '',
    nomorSurat: nomorText,
    title: titleText,
    qrDataUrl: verification?.qrDataUrl || '',
    defaultSignerRole: verification?.signerRole || 'kepala_sekolah',
    signerLabel: verification?.roleLabel || verification?.signerTitle || 'Kepala Sekolah',
    signerName: verification?.signerName || '',
    school: store.getEffectiveSchoolForDocument(model),
    onChooseQr: () => runSmartDownload('qr'),
    onChooseManual: () => runSmartDownload('manual'),
    onAuthorizeAndQr: async (selectedRole) => {
      if (typeof activeSmartAuthorizeHandler !== 'function') return;
      const result = await activeSmartAuthorizeHandler(selectedRole, false);
      const roleInfo = buildSignerSelectionInfo(store.getEffectiveSchoolForDocument(model), selectedRole);
      await runSmartDownload('qr', {
        qrDataUrl: result.qrDataUrl,
        verificationUrl: result.payload?.verificationUrl,
        signerRole: selectedRole,
        roleLabel: roleInfo.roleLabel
      });
    },
    onRequestAcc: async (selectedRole) => {
      if (typeof activeSmartRequestAccHandler !== 'function') return;
      const res = await activeSmartRequestAccHandler(selectedRole, false);
      showToast(`⏳ Surat #${String(res.recorded.sequenceNumber).padStart(3, '0')} masuk Antrean ACC Kepala Sekolah!`, 'info');
    }
  });
}

// =========================================================================
// BAGIAN 3: RENDER UI KANVAS WORD PRO & LOGIKA MULTI-HALAMAN MS WORD
// =========================================================================

const PAPER_LABELS = {
  a4: 'A4 (21 x 29,7 cm)',
  f4: 'F4 / Folio (21,5 x 33 cm)',
  letter: 'Letter (21,6 x 27,9 cm)',
  legal: 'Legal (21,6 x 35,6 cm)'
};

export function createEmptySmartDocument() {
  return {
    pages: [{ id: `spage-${Date.now()}-1`, html: '<p><br></p>', isManualBreak: false }],
    meta: { title: 'Dokumen Baru', nomorSurat: '' },
    isBlankDraft: true
  };
}

export function renderSmartWordCanvas(container) {
  const state = store.getState();
  const school = state.schoolProfile;

  // Konsumsi pendingDocModel (misalnya dari Gudang Surat / Template SPPD) sebelum membaca paperSize
  let smartDoc = state.smartDocument;
  if (smartDoc?.pendingDocModel) {
    const pendingModel = smartDoc.pendingDocModel;
    const pendingSchool = smartDoc.schoolProfile
      || resolveDocumentSchoolProfile(pendingModel, school, state.kopGallery || []);
    const converted = docModelToSmartPages(pendingModel, pendingSchool);
    const pendingSetupPatch = {};
    if (pendingModel.pageSetup && typeof pendingModel.pageSetup === 'object') {
      Object.assign(pendingSetupPatch, pendingModel.pageSetup);
    }
    if (pendingModel.paperSize) pendingSetupPatch.paperSize = pendingModel.paperSize;
    if (pendingModel.orientation) pendingSetupPatch.orientation = pendingModel.orientation;
    if (pendingModel.customMargins) pendingSetupPatch.margins = pendingModel.customMargins;
    if (pendingModel.fontFamily) pendingSetupPatch.fontFamily = pendingModel.fontFamily;
    if (pendingModel.fontSize) pendingSetupPatch.fontSize = pendingModel.fontSize;
    if (pendingModel.lineSpacing) pendingSetupPatch.lineSpacing = pendingModel.lineSpacing;
    if (pendingModel.paragraphSpacing !== undefined) pendingSetupPatch.paragraphSpacing = pendingModel.paragraphSpacing;
    if (pendingModel.pageNumbering !== undefined) pendingSetupPatch.pageNumbering = pendingModel.pageNumbering;
    if (Object.keys(pendingSetupPatch).length > 0) {
      store.updateLetterOptions(pendingSetupPatch);
    }
    smartDoc = {
      pages: converted.pages,
      meta: converted.meta || smartDoc.meta || {},
      targetUnitId: pendingModel.targetUnitId || smartDoc.targetUnitId || pendingSchool.unitId,
      schoolProfile: pendingSchool,
      ...(smartDoc.verification ? { verification: smartDoc.verification } : {}),
      ...(pendingModel.qrPlacement || smartDoc.qrPlacement ? { qrPlacement: pendingModel.qrPlacement || smartDoc.qrPlacement } : {})
    };
    store.setState({ smartDocument: smartDoc }, false);
  } else if (!smartDoc || !Array.isArray(smartDoc.pages) || smartDoc.pages.length === 0) {
    smartDoc = createEmptySmartDocument();
    store.setState({ smartDocument: smartDoc }, false);
  }

  const opts = normalizePageSetupOptions(store.getLetterOptions());
  const paperSize = opts.paperSize || 'a4';
  const orientation = opts.orientation || 'portrait';
  const customMargins = opts.margins || { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 };
  const fontFamily = opts.fontFamily || 'Times New Roman';
  const fontSize = opts.fontSize || '12';
  const lineSpacing = opts.lineSpacing || '1.15';
  const paragraphSpacing = opts.paragraphSpacing || '6';
  const showKop = opts.showKop !== false;
  syncGlobalPrintPageStyle(opts);

  const initialZoomPct = state.zoomUserAdjusted ? (state.zoomLevel || 100) : computeDefaultDesktopZoom();
  let currentZoom = Number((initialZoomPct / 100).toFixed(2));
  let copilotVisible = typeof window !== 'undefined' ? window.innerWidth >= 1120 : true;
  let showFormattingMarks = false;
  let activePageIndex = 0;
  let lastSavedRange = null;
  const smartUndoStack = [];
  const smartRedoStack = [];

  const apiVault = state.apiVault || [];
  let selectedApiId = apiVault.find(a => a.isActive)?.id || apiVault[0]?.id || '';
  const kbStats = summarizeKnowledgeBaseStats(state.letterVault || []);

  const kopUrl = school.kopImage || DEFAULT_KOP_DATA_URL;
  const initDims = computePrintableDimsMm({ paperSize, orientation, margins: customMargins });
  const marginStyles = `
    --margin-top: ${customMargins.top}cm;
    --margin-bottom: ${customMargins.bottom}cm;
    --margin-left: ${customMargins.left}cm;
    --margin-right: ${customMargins.right}cm;
    --printable-h-mm: ${initDims.printableHeightMm}mm;
    --doc-font: '${fontFamily}', serif;
    --doc-font-size: ${fontSize}pt;
    --doc-line-height: ${lineSpacing};
    --doc-para-after: ${paragraphSpacing}pt;
  `;

  container.innerHTML = `
    <div class="word-workspace smart-word-workspace">
      <!-- RIBBON CARD-GROUPED (GAYA MS WORD ASLI + SMART FITUR) -->
      <div class="word-ribbon no-print">
        <!-- Group 1: Halaman -->
        <div class="ribbon-group">
          <div class="ribbon-rows-wrapper">
            <div class="ribbon-row">
              <select class="ribbon-select" id="smart-paper-select" title="Ukuran Kertas">
                <option value="a4" ${paperSize === 'a4' ? 'selected' : ''}>A4</option>
                <option value="f4" ${paperSize === 'f4' ? 'selected' : ''}>F4</option>
                <option value="letter" ${paperSize === 'letter' ? 'selected' : ''}>Letter</option>
                <option value="legal" ${paperSize === 'legal' ? 'selected' : ''}>Legal</option>
              </select>
              <div class="ribbon-select-icon-wrap" title="Orientasi Kertas">
                ${ICON_PORTRAIT}
                <select class="ribbon-select" id="smart-orientation-select">
                  <option value="portrait" ${orientation === 'portrait' ? 'selected' : ''}>Potret</option>
                  <option value="landscape" ${orientation === 'landscape' ? 'selected' : ''}>Lanskap</option>
                </select>
              </div>
            </div>
            <div class="ribbon-row">
              <div class="smart-popover-wrap">
                <button class="ribbon-btn ribbon-btn-compact" id="smart-btn-margin-menu" title="Pilih Margin Cepat">
                  ${ICON_RULER} <span>Margin</span>
                </button>
                <div class="smart-popover-menu hidden" id="smart-margin-popover">
                  <button class="smart-popover-item" data-margin="2,2,2,2.5">Standar Dinas <span>2 / 2,5 cm</span></button>
                  <button class="smart-popover-item" data-margin="2.54,2.54,2.54,2.54">Normal Word <span>2,54 cm</span></button>
                  <button class="smart-popover-item" data-margin="1.27,1.27,1.27,1.27">Sempit (Narrow) <span>1,27 cm</span></button>
                  <button class="smart-popover-item" data-margin="3,3,3,3">Lebar (Wide) <span>3,0 cm</span></button>
                </div>
              </div>
              <button class="ribbon-btn ribbon-btn-compact" id="smart-btn-format-modal" title="Pengaturan Kertas, Margin, Spasi & Nomor Halaman">
                ${ICON_SETTINGS} <span>Format</span>
              </button>
            </div>
          </div>
          <div class="ribbon-label">Halaman</div>
        </div>

        <!-- Group 2: Font & Teks -->
        <div class="ribbon-group">
          <div class="ribbon-rows-wrapper">
            <div class="ribbon-row">
              <select class="ribbon-select" id="smart-block-style" title="Gaya Paragraf / Judul" style="max-width: 96px;">
                <option value="p">Normal</option>
                <option value="h1">Judul Besar (H1)</option>
                <option value="h2">Subjudul (H2)</option>
                <option value="h3">Sub-bagian (H3)</option>
              </select>
              <select class="ribbon-select" id="smart-font-family" title="Jenis Huruf">
                ${SUPPORTED_FONT_FAMILIES.map(f =>
                  `<option value="${f}" ${fontFamily === f ? 'selected' : ''}>${f}</option>`
                ).join('')}
              </select>
              <select class="ribbon-select" id="smart-font-size" title="Ukuran Huruf" style="width: 52px;">
                ${['8', '9', '10', '11', '12', '13', '14', '16', '18', '20', '22', '24', '28'].map(sz =>
                  `<option value="${sz}" ${String(fontSize) === sz ? 'selected' : ''}>${sz}</option>`
                ).join('')}
              </select>
              <button class="ribbon-btn-icon" id="smart-btn-font-up" title="Perbesar Huruf">A<sup>+</sup></button>
              <button class="ribbon-btn-icon" id="smart-btn-font-down" title="Perkecil Huruf">A<sup>-</sup></button>
              <button class="ribbon-btn-icon" id="smart-btn-change-case" title="Ubah Kapitalisasi (UPPER / lower / Title)">${ICON_CASE}</button>
              <button class="ribbon-btn-icon" id="smart-btn-clear-fmt" title="Bersihkan Pemformatan">${ICON_ERASER}</button>
            </div>
            <div class="ribbon-row">
              <button class="ribbon-btn-icon" data-cmd="bold" title="Tebal (Ctrl+B)"><b>B</b></button>
              <button class="ribbon-btn-icon" data-cmd="italic" title="Miring (Ctrl+I)"><i>I</i></button>
              <button class="ribbon-btn-icon" data-cmd="underline" title="Garis Bawah (Ctrl+U)"><u>U</u></button>
              <button class="ribbon-btn-icon" data-cmd="strikeThrough" title="Coret"><s>S</s></button>
              <button class="ribbon-btn-icon" data-cmd="subscript" title="Subskrip">x<sub>2</sub></button>
              <button class="ribbon-btn-icon" data-cmd="superscript" title="Superskrip">x<sup>2</sup></button>
              <div class="ribbon-divider"></div>
              <label class="ribbon-btn-icon ribbon-btn-color-wrap" title="Warna Sorotan Stabilo">
                <span style="background:#fef08a; padding: 0 4px; border-radius:3px; font-weight:700; font-size:11px;">ab</span>
                <input type="color" id="smart-highlight-color" class="ribbon-color-picker-input" value="#fef08a">
              </label>
              <label class="ribbon-btn-icon ribbon-btn-color-wrap" title="Warna Teks">
                <span style="border-bottom: 2.5px solid #dc2626; font-weight:700; font-size:12px;">A</span>
                <input type="color" id="smart-text-color" class="ribbon-color-picker-input" value="#dc2626">
              </label>
            </div>
          </div>
          <div class="ribbon-label">Font &amp; Teks</div>
        </div>

        <!-- Group 3: Paragraf -->
        <div class="ribbon-group">
          <div class="ribbon-rows-wrapper">
            <div class="ribbon-row">
              <button class="ribbon-btn-icon" data-cmd="justifyLeft" title="Rata Kiri">${ICON_ALIGN_LEFT}</button>
              <button class="ribbon-btn-icon" data-cmd="justifyCenter" title="Rata Tengah">${ICON_ALIGN_CENTER}</button>
              <button class="ribbon-btn-icon" data-cmd="justifyRight" title="Rata Kanan">${ICON_ALIGN_RIGHT}</button>
              <button class="ribbon-btn-icon" data-cmd="justifyFull" title="Rata Kiri-Kanan">${ICON_ALIGN_JUSTIFY}</button>
              <div class="ribbon-divider"></div>
              <button class="ribbon-btn-icon" id="smart-btn-pilcrow" title="Tampilkan / Sembunyikan Tanda Paragraf">${ICON_PILCROW}</button>
              <button class="ribbon-btn-icon" id="smart-btn-shading" title="Latar Paragraf Lembut">${ICON_SHADING}</button>
            </div>
            <div class="ribbon-row">
              <button class="ribbon-btn-icon" data-cmd="insertUnorderedList" title="Daftar Poin (Bullet)">${ICON_LIST_BULLET}</button>
              <button class="ribbon-btn-icon" data-cmd="insertOrderedList" title="Daftar Bernomor (1, 2, 3)">${ICON_LIST_NUMBER}</button>
              <select class="ribbon-select" id="smart-list-style-select" title="Jenis Penomoran Daftar (1,2,3 / a,b,c / A,B,C / i,ii / I,II / Bullet)" style="max-width: 72px;">
                <option value="">Nomor...</option>
                <option value="decimal">1, 2, 3</option>
                <option value="lower-alpha">a, b, c</option>
                <option value="upper-alpha">A, B, C</option>
                <option value="lower-roman">i, ii, iii</option>
                <option value="upper-roman">I, II, III</option>
                <option value="bullet">• Bullet</option>
              </select>
              <button class="ribbon-btn-icon" data-cmd="outdent" title="Kurangi Indentasi">${ICON_OUTDENT}</button>
              <button class="ribbon-btn-icon" data-cmd="indent" title="Tambah Indentasi">${ICON_INDENT}</button>
              <select class="ribbon-select" id="smart-line-spacing" title="Spasi Baris">
                <option value="1.0" ${String(lineSpacing) === '1.0' ? 'selected' : ''}>1,0</option>
                <option value="1.15" ${String(lineSpacing) === '1.15' ? 'selected' : ''}>1,15</option>
                <option value="1.5" ${String(lineSpacing) === '1.5' ? 'selected' : ''}>1,5</option>
                <option value="2.0" ${String(lineSpacing) === '2.0' ? 'selected' : ''}>2,0</option>
              </select>
              <div class="smart-popover-wrap">
                <button class="ribbon-btn ribbon-btn-compact" id="smart-btn-para-format" title="Inden Paragraf (First-Line, Left Indent) &amp; Jarak Paragraf">
                  <span>Inden</span>
                </button>
                <div class="smart-popover-menu hidden" id="smart-para-format-popover" style="min-width: 245px;">
                  <div style="font-size:11px;font-weight:700;color:#1e293b;padding:2px 6px;">Inden Baris Pertama (Paragraf Aktif)</div>
                  <button class="smart-popover-item" data-first-indent="0">Tanpa Inden (Rata) <span>0 cm</span></button>
                  <button class="smart-popover-item" data-first-indent="1">Inden Baris Pertama <span>1,0 cm</span></button>
                  <button class="smart-popover-item" data-first-indent="1.25">Inden Standar Word <span>1,25 cm</span></button>
                  <hr style="border:none;border-top:1px solid #e2e8f0;margin:4px 0;">
                  <div style="font-size:11px;font-weight:700;color:#1e293b;padding:2px 6px;">Inden Kiri Paragraf Aktif</div>
                  <button class="smart-popover-item" data-left-indent="0">Reset Inden Kiri <span>0 cm</span></button>
                  <button class="smart-popover-item" data-left-indent="0.75">Geser Masuk Kiri <span>0,75 cm</span></button>
                  <button class="smart-popover-item" data-left-indent="1.5">Geser Masuk Kiri <span>1,5 cm</span></button>
                  <hr style="border:none;border-top:1px solid #e2e8f0;margin:4px 0;">
                  <div style="font-size:11px;font-weight:700;color:#1e293b;padding:2px 6px;">Jarak Antar Paragraf (Global)</div>
                  <button class="smart-popover-item" data-para-spacing="0">Rapat (Tanpa Jarak) <span>0 pt</span></button>
                  <button class="smart-popover-item" data-para-spacing="6">Standar Word <span>6 pt</span></button>
                  <button class="smart-popover-item" data-para-spacing="10">Lega (After 10 pt) <span>10 pt</span></button>
                </div>
              </div>
            </div>
          </div>
          <div class="ribbon-label">Paragraf</div>
        </div>

        <!-- Group 4: Sisipkan (Irit Ruang 2 Baris + Hero ST+SPPD) -->
        <div class="ribbon-group">
          <div class="ribbon-hybrid-row">
            <div class="ribbon-rows-wrapper ribbon-compact-rows">
              <div class="ribbon-row">
                <button class="ribbon-tile is-compact-row" id="smart-btn-add-page" title="Tambah Halaman Baru (Ctrl+Enter)">
                  ${ICON_PAGE_ADD}
                  <span>+ Hal</span>
                </button>
                <button class="ribbon-tile is-compact-row" id="smart-btn-insert-table" title="Sisipkan Tabel Grid / SPPD / Data">
                  ${ICON_TABLE}
                  <span>Tabel</span>
                </button>
                <button class="ribbon-tile is-compact-row" id="smart-btn-insert-2col" title="Sisipkan Tata Letak 2 Kolom Berdampingan Tanpa Garis (Catatan + TTD)">
                  ${ICON_TABLE}
                  <span>2 Kolom</span>
                </button>
              </div>
              <div class="ribbon-row">
                <button class="ribbon-tile is-compact-row" id="smart-btn-insert-sig" title="Sisipkan Blok Tanda Tangan">
                  ${ICON_PENCIL_NOTE}
                  <span>TTD</span>
                </button>
                <button class="ribbon-tile is-compact-row hidden" id="smart-btn-toggle-qr-canvas" title="Tampilkan atau Sembunyikan Barcode QR di Kanvas (saat disembunyikan tetap tercetak &amp; muncul di Surat Keluar)">
                  ${ICON_CLIPBOARD_CHECK}
                  <span data-qr-btn-label="1">QR</span>
                </button>
                <button class="ribbon-tile is-compact-row" id="smart-btn-kop-gallery" title="Pilih / Kelola Kop Surat">
                  ${ICON_GALLERY}
                  <span>Kop</span>
                </button>
                <button class="ribbon-tile is-compact-row ${showKop ? 'active' : ''}" id="smart-btn-toggle-kop" title="Tampilkan / Sembunyikan Kop Surat">
                  ${showKop ? ICON_EYE : ICON_EYE_OFF}
                  <span>Tampil Kop</span>
                </button>
              </div>
            </div>
            <button class="ribbon-tile is-hero smart-ribbon-highlight" id="smart-btn-template-sppd" title="Terapkan Template Resmi 3 Halaman: Surat Tugas + SPPD + Keterangan (1:1 MS Word)">
              ${ICON_TARGET}
              <span>ST + SPPD</span>
            </button>
          </div>
          <div class="ribbon-label">Sisipkan</div>
        </div>

        <!-- Group 5: Dokumen (Irit Ruang 2 Baris + Hero Copilot AI agar 1 Baris Penuh Layar) -->
        <div class="ribbon-group">
          <div class="ribbon-hybrid-row">
            <div class="ribbon-rows-wrapper ribbon-compact-rows">
              <div class="ribbon-row">
                <button class="ribbon-tile is-compact-row" id="smart-btn-save-vault" title="Simpan perubahan ke dokumen aktif di Gudang Surat (Ctrl+S)">
                  ${ICON_SAVE}
                  <span>Simpan</span>
                </button>
                <button class="ribbon-tile is-compact-row" id="smart-btn-save-as-vault" title="Save As / Duplikat sebagai dokumen baru di Gudang Surat (Ctrl+Shift+S / F12)">
                  ${ICON_COPY_DOC}
                  <span>Save As</span>
                </button>
                <button class="ribbon-tile is-compact-row" id="smart-btn-version-history" title="Riwayat Versi Dokumen Gudang &amp; Pulihkan Versi Sebelumnya">
                  ${SMART_ICON_HISTORY}
                  <span>Riwayat</span>
                </button>
                <button class="ribbon-tile is-compact-row" id="smart-btn-find-replace" title="Cari &amp; Ganti Teks Lintas Halaman (Ctrl+F / Ctrl+H)">
                  ${SMART_ICON_SEARCH}
                  <span>Cari</span>
                </button>
              </div>
              <div class="ribbon-row">
                <button class="ribbon-btn-icon-sq" id="smart-btn-undo" title="Urungkan Perubahan (Ctrl+Z)">
                  ${ICON_UNDO}
                </button>
                <button class="ribbon-btn-icon-sq" id="smart-btn-redo" title="Ulangi Perubahan (Ctrl+Y)">
                  ${ICON_REDO}
                </button>
                <button class="ribbon-btn-icon-sq" id="smart-btn-refresh" title="Reset / Kosongkan Kertas ke Lembar Baru">
                  ${ICON_REFRESH}
                </button>
                <button class="ribbon-tile is-compact-row" id="smart-btn-pull-canvas1" title="Salin Isi Surat dari Lembar Kerja Lama ke Mode Word Pro">
                  ${ICON_COPY_DOC}
                  <span>Salin 1</span>
                </button>
                <button class="ribbon-tile is-compact-row" id="smart-btn-import-docx" title="Impor Berkas .docx ke Kanvas Word Pro">
                  ${ICON_DOC_UPLOAD}
                  <span>Impor</span>
                </button>
              </div>
            </div>
            <button class="ribbon-tile is-hero ${copilotVisible ? 'active' : ''}" id="smart-btn-toggle-copilot" title="Tampilkan / Sembunyikan AI Copilot">
              ${ICON_BOT_SPARK}
              <span>Copilot AI</span>
            </button>
          </div>
          <div class="ribbon-label">Dokumen</div>
        </div>
        <input type="file" id="smart-docx-upload-input" accept=".docx" style="display:none" />
        <input type="file" id="smart-kop-upload-input" accept="image/png,image/jpeg,image/webp,image/svg+xml" style="display:none" />
      </div>

      <!-- CONTEXT BAR ALAT TABEL WORD (MUNCUL OTOMATIS SAAT KURSOR DI DALAM TABEL) -->
      <div class="smart-table-context-bar hidden no-print" id="smart-table-context-bar">
        <span class="smart-tbl-badge">Alat Tabel</span>
        <div class="smart-tbl-group">
          <button type="button" class="smart-tbl-btn" data-tbl-act="row-above" title="Tambah baris di atas">+ Baris ↑</button>
          <button type="button" class="smart-tbl-btn" data-tbl-act="row-below" title="Tambah baris di bawah">+ Baris ↓</button>
          <button type="button" class="smart-tbl-btn smart-tbl-btn-danger" data-tbl-act="row-del" title="Hapus baris aktif">− Baris</button>
        </div>
        <div class="smart-tbl-group">
          <button type="button" class="smart-tbl-btn" data-tbl-act="col-right" title="Tambah kolom di kanan">+ Kolom →</button>
          <button type="button" class="smart-tbl-btn smart-tbl-btn-danger" data-tbl-act="col-del" title="Hapus kolom aktif">− Kolom</button>
        </div>
        <div class="smart-tbl-group">
          <button type="button" class="smart-tbl-btn" data-tbl-act="merge-sel" id="sm-btn-merge-sel" title="Gabung sel yang dipilih dengan Shift+Klik">⧉ Gabung Seleksi</button>
          <button type="button" class="smart-tbl-btn" data-tbl-act="merge-right" title="Gabung sel ini dengan sel sebelah kanan (Colspan)">⧉ →</button>
          <button type="button" class="smart-tbl-btn" data-tbl-act="merge-down" title="Gabung sel ini dengan sel di bawahnya (Rowspan)">⧉ ↓</button>
          <button type="button" class="smart-tbl-btn" data-tbl-act="unmerge" title="Pisah kembali sel yang digabung (Unmerge)">⊟ Pisah Sel</button>
        </div>
        <div class="smart-tbl-group">
          <select id="sm-ctx-border" class="ribbon-select ribbon-select-xs" title="Model Garis Tabel">
            <option value="grid">Garis Penuh (Grid)</option>
            <option value="sppd">Garis SPPD (Tanpa Samping)</option>
            <option value="none">Tanpa Garis (2-Kolom)</option>
          </select>
          <button type="button" class="smart-tbl-btn" data-tbl-act="toggle-header-bg" title="Ubah warna latar Header Abu-abu vs Putih Bersih">Header Abu/Putih</button>
          <select id="sm-ctx-row-height" class="ribbon-select ribbon-select-xs" title="Tinggi Minimum Baris (untuk Area Tanda Tangan / Stempel)">
            <option value="0">Tinggi Otomatis</option>
            <option value="2.5">Tinggi Stamp 2.5 cm</option>
            <option value="4.5">Tinggi Stamp 4.5 cm</option>
            <option value="5.5">Tinggi Stamp 5.5 cm</option>
          </select>
          <button type="button" class="smart-tbl-btn" data-tbl-act="align-left" title="Rata Kiri Sel">Kiri</button>
          <button type="button" class="smart-tbl-btn" data-tbl-act="align-center" title="Rata Tengah Sel">Tengah</button>
          <button type="button" class="smart-tbl-btn" data-tbl-act="align-right" title="Rata Kanan Sel">Kanan</button>
          <button type="button" class="smart-tbl-btn smart-tbl-btn-danger" data-tbl-act="del-table" title="Hapus seluruh tabel">Hapus Tabel</button>
        </div>
        <span class="smart-tbl-hint">Tip: Shift+Klik beberapa sel lalu Gabung Seleksi, atau Tab di sel akhir untuk tambah baris.</span>
      </div>

      <!-- STAGE UTAMA: KANVAS KERTAS MULTI-HALAMAN (KIRI) + SIDEBAR COPILOT (KANAN) -->
      <div class="word-main-stage">
        <div class="word-main-left">
          <div class="word-desk" id="smart-word-desk">
            <div class="word-sheet-wrapper" style="width:100%; display:flex; flex-direction:column; align-items:center;">

              <!-- BAR EDIT AI CEPAT -->
              <div class="quick-ai-edit-bar no-print" style="max-width: 920px; width: 100%; margin-bottom: 14px;">
                <div class="quick-ai-edit-main">
                  <div class="quick-ai-badge">
                    ${ICON_SPARKLE}
                    <span>AI Edit Word Pro</span>
                  </div>
                  <input
                    type="text"
                    id="smart-quick-ai-input"
                    class="quick-ai-input"
                    placeholder="Ketik instruksi edit cepat (mis: tambah halaman lampiran jadwal kegiatan 2 hari, rapikan ejaan PUEBI)..."
                  />
                  <button class="quick-ai-btn" id="smart-btn-quick-ai-apply">
                    ${ICON_SPARKLE} <span>Terapkan Edit</span>
                  </button>
                  <button class="quick-ai-outgoing-btn" id="smart-btn-record-outgoing" title="Catat ke Buku Surat Keluar Resmi">
                    ${ICON_SEND_OUT} <span>Catat Surat Keluar</span>
                  </button>
                </div>
                <div class="quick-ai-chips-row">
                  <span style="font-size: 11px; color: #475569; font-weight: 600;">Contoh Instruksi:</span>
                  <button class="quick-ai-chip" data-prompt="Tambahkan rincian Hari/Tanggal, Pukul, Tempat, dan Acara pada isi surat">+ Tambah Jadwal &amp; Tempat</button>
                  <button class="quick-ai-chip" data-prompt="Perbaiki ejaan PUEBI dan tingkatkan kalimat agar lebih resmi kedinasan">+ Perbaiki Kata &amp; PUEBI</button>
                  <button class="quick-ai-chip" data-prompt="Ubah tujuan surat Kepada Yth menjadi Kepala Dinas Pendidikan Kabupaten Garut">+ Ubah Kepada Yth</button>
                  <button class="quick-ai-chip" data-prompt="Tambahkan bagian Lampiran daftar peserta rapat pada halaman baru">+ Tambah Halaman Lampiran</button>
                </div>
                <div id="smart-quick-ai-status" class="quick-ai-status hidden"></div>
              </div>

              <!-- BANNER FEEDBACK REVISI KEPALA SEKOLAH (AKTIF BILA DIBUKA DARI ANTREAN ACC DITOLAK) -->
              <div id="smart-kepsek-revision-banner" class="smart-kepsek-rev-banner hidden no-print" style="max-width: 920px; width: 100%; margin-bottom: 12px;"></div>

              <!-- BAR CARI & GANTI CEPAT (CTRL+F / CTRL+H) -->
              <div id="smart-find-replace-bar" class="smart-find-replace-bar hidden no-print" style="max-width: 920px; width: 100%; margin-bottom: 12px;">
                <div class="smart-fr-row">
                  <div class="smart-fr-inputs">
                    <input type="text" id="smart-find-input" class="smart-fr-input" placeholder="Cari kata / tanggal / nama..." autocomplete="off" />
                    <input type="text" id="smart-replace-input" class="smart-fr-input" placeholder="Ganti dengan..." autocomplete="off" />
                    <label class="smart-fr-case-label" title="Cocokkan huruf besar/kecil (Case Sensitive)">
                      <input type="checkbox" id="smart-find-case" />
                      <span>Aa</span>
                    </label>
                  </div>
                  <div class="smart-fr-controls">
                    <span id="smart-find-count" class="smart-fr-count">0/0</span>
                    <button type="button" class="smart-tbl-btn" id="smart-find-prev" title="Hasil Sebelumnya (Shift+Enter)">▲</button>
                    <button type="button" class="smart-tbl-btn" id="smart-find-next" title="Hasil Berikutnya (Enter)">▼</button>
                    <button type="button" class="smart-tbl-btn" id="smart-replace-one" title="Ganti yang sedang disorot">Ganti</button>
                    <button type="button" class="smart-tbl-btn smart-tbl-btn-primary" id="smart-replace-all" title="Ganti semua kecocokan di semua halaman">Ganti Semua</button>
                    <button type="button" class="smart-tbl-btn" id="smart-find-close" title="Tutup Cari & Ganti (Esc)">✕</button>
                  </div>
                </div>
              </div>

              <!-- PANEL COMPARE / URUNGKAN REVISI AI TERAKHIR -->
              <div id="smart-ai-compare-bar" class="smart-ai-compare-bar hidden no-print" style="max-width: 920px; width: 100%; margin-bottom: 12px;">
                <div class="smart-cmp-info">
                  <span id="smart-cmp-summary">✨ Revisi AI diterapkan (tata letak &amp; tabel aman).</span>
                </div>
                <div class="smart-cmp-actions">
                  <button type="button" class="smart-tbl-btn smart-tbl-btn-primary" id="smart-btn-open-compare">👁️ Bandingkan Perubahan</button>
                  <button type="button" class="smart-tbl-btn smart-tbl-btn-danger" id="smart-btn-rollback-ai">↶ Urungkan Revisi AI</button>
                  <button type="button" class="smart-tbl-btn" id="smart-btn-dismiss-compare" title="Tutup bar perbandingan">✓</button>
                </div>
              </div>

              <!-- TUMPUKAN HALAMAN KERTAS MS WORD BEBAS & BERKELANJUTAN -->
              <div class="smart-pages-stack" id="smart-pages-stack" style="${marginStyles} transform: scale(${currentZoom});"></div>
            </div>
          </div>
        </div>

        <!-- SIDEBAR AI COPILOT KANAN (PERCAKAPAN GAYA CHATGPT TERPADU DENGAN AI ASISTEN) -->
        <aside class="copilot-drawer ${copilotVisible ? '' : 'hidden'} no-print" id="smart-copilot-drawer"></aside>
      </div>

      <!-- STATUS BAR BAWAH -->
      <div class="word-status-bar no-print">
        <div class="status-left">
          <span class="status-item" id="smart-status-pages">${ICON_PAGE_ADD} Halaman 1 dari 1</span>
          <span class="status-sep">|</span>
          <span class="status-item status-word-count" id="smart-status-words">0 kata</span>
          <span class="status-sep">|</span>
          <span class="status-item">${ICON_GLOBE} Bahasa Indonesia (PUEBI) &bull; Mode Bebas MS Word</span>
          <span class="status-sep">|</span>
          <span class="status-item status-vault-pill" id="smart-status-vault-badge">📄 Draf Baru</span>
          <span class="status-sep">|</span>
          <span class="status-item" id="smart-status-save-state" title="Status penyimpanan otomatis lokal & sinkronisasi Gudang Surat">✓ Draf lokal siap</span>
          <span class="status-sep">|</span>
          <button type="button" id="smart-btn-paste-mode-pill" class="status-item status-paste-pill" title="Klik untuk ganti mode tempel (Bersih Otomatis vs Tempel Teks Polos). Tekan Ctrl+Shift+V kapan saja untuk tempel polos.">🧹 Paste Bersih: Aktif</button>
        </div>
        <div class="status-right">
          <div class="zoom-slider-wrap">
            <button class="zoom-btn" id="smart-zoom-out" title="Perkecil Tampilan">-</button>
            <input type="range" id="smart-zoom-slider" min="50" max="150" value="${initialZoomPct}" step="5" />
            <button class="zoom-btn" id="smart-zoom-in" title="Perbesar Tampilan">+</button>
          </div>
          <span class="status-zoom-pill" id="smart-zoom-label">${initialZoomPct}%</span>
          <button class="status-fit-btn" id="smart-zoom-fit" title="Sesuaikan Lebar Layar">${ICON_MAXIMIZE}</button>
        </div>
      </div>

      <!-- MODAL AREA (SMART-SCOPED) -->
      <div id="smart-modal-root"></div>
    </div>
  `;

  // =========================================================================
  // LOGIKA RENDER HALAMAN, PAGINASI OTOMATIS & INTERAKSI WORD ASLI
  // =========================================================================

  const pagesStack = container.querySelector('#smart-pages-stack');
  const statusPagesEl = container.querySelector('#smart-status-pages');
  const statusWordsEl = container.querySelector('#smart-status-words');
  const statusSaveStateEl = container.querySelector('#smart-status-save-state');
  const pasteModePillEl = container.querySelector('#smart-btn-paste-mode-pill');
  const modalRoot = container.querySelector('#smart-modal-root');

  let reflowTimer = null;
  let saveTimer = null;
  let pastePlainTextMode = false;
  let hasUnsavedVaultChanges = false;
  let lastLocalSavedTime = '';
  let lastAiCompareState = null;

  function updateSmartSaveIndicator({ savedToVault = false } = {}) {
    if (!statusSaveStateEl) return;
    const activeEntry = store.getActiveVaultEntry('smart');
    if (savedToVault) {
      hasUnsavedVaultChanges = false;
    }
    const nowTime = new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
    if (activeEntry) {
      const verCount = Array.isArray(activeEntry.versionHistory) ? activeEntry.versionHistory.length : 1;
      if (hasUnsavedVaultChanges) {
        statusSaveStateEl.innerHTML = `<span class="save-dot is-pending"></span> Belum simpan Gudang (Ctrl+S)${lastLocalSavedTime ? ` · Lokal ${lastLocalSavedTime}` : ''}`;
      } else {
        statusSaveStateEl.innerHTML = `<span class="save-dot is-synced"></span> Tersinkron Gudang · ${verCount} versi (${lastLocalSavedTime || nowTime})`;
      }
    } else {
      statusSaveStateEl.innerHTML = `<span class="save-dot is-local"></span> Draf lokal tersimpan${lastLocalSavedTime ? ` · ${lastLocalSavedTime}` : ''}`;
    }
  }

  pasteModePillEl?.addEventListener('click', () => {
    pastePlainTextMode = !pastePlainTextMode;
    pasteModePillEl.textContent = pastePlainTextMode ? '📋 Paste Polos: Aktif' : '🧹 Paste Bersih: Aktif';
    pasteModePillEl.classList.toggle('is-plain', pastePlainTextMode);
    showToast(
      pastePlainTextMode
        ? 'Mode Paste Polos aktif: semua teks yang ditempel otomatis bebas format/gaya sumber.'
        : 'Mode Paste Bersih aktif: tebal/miring/tabel tetap terjaga, gaya liar sumber dibersihkan.',
      'info'
    );
  });

  function collectLivePagesData() {
    if (!pagesStack || !pagesStack.isConnected) {
      return store.getState().smartDocument?.pages || smartDoc.pages || [];
    }
    const pageEls = [...pagesStack.querySelectorAll('.smart-editable-page')];
    if (pageEls.length === 0) return smartDoc.pages || [];
    return pageEls.map((el, idx) => ({
      id: el.dataset.pageId || `spage-${Date.now()}-${idx + 1}`,
      html: el.innerHTML,
      isManualBreak: el.dataset.manualBreak === 'true'
    }));
  }

  activeLivePagesExtractor = collectLivePagesData;

  function pushUndoSnapshot() {
    const snap = JSON.stringify(collectLivePagesData());
    if (smartUndoStack[smartUndoStack.length - 1] !== snap) {
      smartUndoStack.push(snap);
      if (smartUndoStack.length > 40) smartUndoStack.shift();
      smartRedoStack.length = 0;
    }
  }

  function schedulePersistSmartDoc() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      const livePages = collectLivePagesData();
      const latestState = store.getState().smartDocument || {};
      const latestVerif = latestState.verification;
      const latestQrPlacement = latestState.qrPlacement;
      smartDoc = {
        ...smartDoc,
        pages: livePages,
        ...(latestVerif ? { verification: latestVerif } : {}),
        ...(latestQrPlacement ? { qrPlacement: latestQrPlacement } : {})
      };
      store.setState({ smartDocument: smartDoc }, false);
      lastLocalSavedTime = new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
      hasUnsavedVaultChanges = Boolean(store.getActiveVaultEntry('smart'));
      updateSmartSaveIndicator();
      updateStatusCounters();
    }, 250);
  }

  function refreshSmartFloatingQrOverlay() {
    const latestState = store.getState().smartDocument || {};
    const verification = latestState.verification || null;
    const stages = [...pagesStack.querySelectorAll('.floating-qr-stage')];
    const toggleBtn = container.querySelector('#smart-btn-toggle-qr-canvas');
    const curOpts = store.getLetterOptions();
    const dims = computePrintableDimsMm(curOpts);
    pagesStack.style.setProperty('--printable-h-mm', `${dims.printableHeightMm}mm`);
    pagesStack.querySelectorAll('.smart-paper-sheet').forEach((sh) => {
      sh.style.setProperty('--printable-h-mm', `${dims.printableHeightMm}mm`);
    });

    const activeModel = getActiveSmartDocModel();
    const effectiveSchool = store.getEffectiveSchoolForDocument(activeModel) || store.getSchoolProfile();
    const mountedPlacement = mountInteractiveFloatingQr({
      stages,
      verification,
      rawPlacement: latestState.qrPlacement || null,
      docModel: activeModel,
      school: effectiveSchool,
      getZoomScale: () => currentZoom,
      onPlacementChange: (nextPlacement) => {
        store.setActiveQrPlacement(nextPlacement, 'smart');
        updateCanvasQrToggleButton(toggleBtn, {
          verification: store.getState().smartDocument?.verification,
          placement: nextPlacement
        });
      }
    });

    const prevQp = latestState.qrPlacement;
    const shouldSyncCalibrated = mountedPlacement && verification?.isVerifiedSah && (
      !prevQp || (
        !prevQp.customPositioned && (
          mountedPlacement.pageIndex !== prevQp.pageIndex ||
          Math.abs(mountedPlacement.xPct - prevQp.xPct) > 0.5 ||
          Math.abs(mountedPlacement.yPct - prevQp.yPct) > 0.5 ||
          mountedPlacement.sizePx !== prevQp.sizePx
        )
      )
    );
    if (shouldSyncCalibrated) {
      store.setActiveQrPlacement(mountedPlacement, 'smart');
    }
    updateCanvasQrToggleButton(toggleBtn, {
      verification,
      placement: store.getState().smartDocument?.qrPlacement || mountedPlacement
    });
  }

  function updateStatusCounters() {
    const pageEls = [...pagesStack.querySelectorAll('.smart-editable-page')];
    const totalPages = Math.max(1, pageEls.length);
    const curIdx = Math.min(Math.max(activePageIndex, 0), totalPages - 1);
    const totalWords = pageEls
      .map(el => (el.innerText || el.textContent || '').trim())
      .join(' ')
      .split(/\s+/)
      .filter(Boolean).length;

    if (statusPagesEl) {
      statusPagesEl.innerHTML = `${ICON_PAGE_ADD} Halaman ${curIdx + 1} dari ${totalPages}`;
    }
    if (statusWordsEl) {
      statusWordsEl.textContent = `${totalWords} kata`;
    }

    // Sinkronkan pill di atas setiap kertas
    const pills = pagesStack.querySelectorAll('.smart-page-pill');
    pills.forEach((pill, idx) => {
      pill.textContent = `Halaman ${idx + 1} dari ${totalPages}`;
    });
  }

  function applyZoom({ fromUser = false } = {}) {
    pagesStack.style.transform = `scale(${currentZoom})`;
    const pct = Math.round(currentZoom * 100);
    container.querySelector('#smart-zoom-label').textContent = `${pct}%`;
    container.querySelector('#smart-zoom-slider').value = String(pct);
    store.setZoomLevel(pct, { fromUser });
  }

  function createPageDomElement(pageData, pageIdx) {
    const curOpts = store.getLetterOptions();
    const curSchool = store.getSchoolProfile();
    const pSize = curOpts.paperSize || 'a4';
    const orient = curOpts.orientation || 'portrait';
    const curShowKop = curOpts.showKop !== false;
    const curKopUrl = curSchool.kopImage || DEFAULT_KOP_DATA_URL;

    const wrapper = document.createElement('div');
    wrapper.className = 'smart-page-wrapper';

    const headerBar = document.createElement('div');
    headerBar.className = 'smart-page-header-bar no-print';
    headerBar.innerHTML = `
      <span class="smart-page-pill">Halaman ${pageIdx + 1}</span>
      <div class="smart-page-actions">
        ${pageIdx > 0 ? `<button type="button" class="smart-page-del-btn" data-action="delete-page">✕ Hapus Halaman</button>` : ''}
      </div>
    `;

    const curDims = computePrintableDimsMm(curOpts);
    const curPn = normalizePageNumbering(curOpts.pageNumbering);
    const sheet = document.createElement('div');
    sheet.className = `smart-paper-sheet paper-${pSize} ${orient === 'landscape' ? 'landscape' : ''}`;
    sheet.style.cssText = `
      --margin-top: ${curOpts.margins?.top ?? 2.0}cm;
      --margin-bottom: ${curOpts.margins?.bottom ?? 2.0}cm;
      --margin-left: ${curOpts.margins?.left ?? 2.5}cm;
      --margin-right: ${curOpts.margins?.right ?? 2.0}cm;
      --printable-h-mm: ${curDims.printableHeightMm}mm;
      --doc-font: '${curOpts.fontFamily || 'Times New Roman'}', serif;
      --doc-font-size: ${curOpts.fontSize || '12'}pt;
      --doc-line-height: ${curOpts.lineSpacing || '1.15'};
      --doc-para-after: ${curOpts.paragraphSpacing ?? '6'}pt;
    `;

    if (pageIdx === 0 && curShowKop) {
      const kopDiv = document.createElement('div');
      kopDiv.className = 'sheet-kop-container';
      kopDiv.id = 'smart-kop-container';
      kopDiv.innerHTML = `<img src="${curKopUrl}" alt="Kop Surat" class="sheet-kop-image" />`;
      kopDiv.title = 'Klik ganda untuk mengganti Kop Surat';
      kopDiv.addEventListener('dblclick', openKopGalleryModal);
      sheet.appendChild(kopDiv);
    }

    const editable = document.createElement('div');
    editable.className = 'smart-editable-page';
    editable.contentEditable = 'true';
    editable.spellcheck = false;
    editable.dataset.pageId = pageData.id || `spage-${Date.now()}-${pageIdx + 1}`;
    editable.dataset.pageIndex = String(pageIdx);
    editable.dataset.manualBreak = pageData.isManualBreak ? 'true' : 'false';
    editable.innerHTML = pageData.html || '<p><br></p>';

    bindEditablePageEvents(editable);

    headerBar.querySelector('[data-action="delete-page"]')?.addEventListener('click', () => {
      const wrappers = [...pagesStack.querySelectorAll('.smart-page-wrapper')];
      if (wrappers.length <= 1) return;
      pushUndoSnapshot();
      const targetIdx = wrappers.indexOf(wrapper);
      const prevEditable = wrappers[targetIdx - 1]?.querySelector('.smart-editable-page');
      const curText = (editable.innerText || '').trim();
      if (curText && prevEditable) {
        while (editable.firstChild) {
          prevEditable.appendChild(editable.firstChild);
        }
      }
      wrapper.remove();
      reindexPageHeaders();
      scheduleReflowPagination();
      schedulePersistSmartDoc();
    });

    const pageNumSlot = document.createElement('div');
    pageNumSlot.className = `smart-page-number-slot no-edit ${curPn.enabled ? `is-${curPn.position}` : 'hidden'}`;
    pageNumSlot.setAttribute('aria-hidden', 'true');
    if (curPn.enabled) {
      pageNumSlot.textContent = formatPageNumberLabel((curPn.startAt || 1) + pageIdx, curPn.format);
    }

    const qrStage = document.createElement('div');
    qrStage.className = 'floating-qr-stage';
    qrStage.dataset.qrPage = String(pageIdx);

    sheet.appendChild(editable);
    sheet.appendChild(pageNumSlot);
    sheet.appendChild(qrStage);
    wrapper.appendChild(headerBar);
    wrapper.appendChild(sheet);
    return wrapper;
  }

  function renderAllSmartPages(pagesList) {
    const latestOpts = normalizePageSetupOptions(store.getLetterOptions());
    const latestDims = computePrintableDimsMm(latestOpts);
    pagesStack.style.setProperty('--margin-top', `${latestOpts.margins.top}cm`);
    pagesStack.style.setProperty('--margin-bottom', `${latestOpts.margins.bottom}cm`);
    pagesStack.style.setProperty('--margin-left', `${latestOpts.margins.left}cm`);
    pagesStack.style.setProperty('--margin-right', `${latestOpts.margins.right}cm`);
    pagesStack.style.setProperty('--printable-h-mm', `${latestDims.printableHeightMm}mm`);
    pagesStack.style.setProperty('--doc-font', `'${latestOpts.fontFamily}', serif`);
    pagesStack.style.setProperty('--doc-font-size', `${latestOpts.fontSize}pt`);
    pagesStack.style.setProperty('--doc-line-height', `${latestOpts.lineSpacing}`);
    pagesStack.style.setProperty('--doc-para-after', `${latestOpts.paragraphSpacing}pt`);

    pagesStack.innerHTML = '';
    const list = (Array.isArray(pagesList) && pagesList.length > 0)
      ? pagesList
      : [{ id: `spage-${Date.now()}-1`, html: '<p><br></p>', isManualBreak: false }];

    list.forEach((p, idx) => {
      pagesStack.appendChild(createPageDomElement(p, idx));
    });

    reindexPageHeaders();
    updateStatusCounters();
    // Jalankan pasca-render check untuk auto-overflow bila teks panjang lalu sinkronkan QR mengambang
    requestAnimationFrame(() => {
      reflowPagination();
      refreshSmartFloatingQrOverlay();
    });
  }

  function reindexPageHeaders() {
    const wrappers = [...pagesStack.querySelectorAll('.smart-page-wrapper')];
    const total = wrappers.length;
    const pn = normalizePageNumbering(store.getLetterOptions().pageNumbering);
    wrappers.forEach((w, idx) => {
      const editable = w.querySelector('.smart-editable-page');
      if (editable) editable.dataset.pageIndex = String(idx);
      const qrStage = w.querySelector('.floating-qr-stage');
      if (qrStage) qrStage.dataset.qrPage = String(idx);
      const pill = w.querySelector('.smart-page-pill');
      if (pill) pill.textContent = `Halaman ${idx + 1} dari ${total}`;
      const numSlot = w.querySelector('.smart-page-number-slot');
      if (numSlot) {
        if (pn.enabled) {
          numSlot.className = `smart-page-number-slot no-edit is-${pn.position}`;
          numSlot.textContent = formatPageNumberLabel((pn.startAt || 1) + idx, pn.format);
        } else {
          numSlot.className = 'smart-page-number-slot no-edit hidden';
          numSlot.textContent = '';
        }
      }
    });
  }

  /**
   * Menyimpan posisi kursor relatif terhadap node tertentu sebelum node dipindahkan antar halaman.
   */
  function captureCaretInNode(targetNode) {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return null;
    const range = sel.getRangeAt(0);
    if (!targetNode.contains(range.startContainer)) return null;

    // Hitung offset karakter dari awal targetNode
    const preRange = document.createRange();
    preRange.selectNodeContents(targetNode);
    preRange.setEnd(range.startContainer, range.startOffset);
    return { charOffset: preRange.toString().length };
  }

  function restoreCaretInNode(targetNode, saved) {
    if (!saved || !targetNode) return;
    let remaining = saved.charOffset;
    const walker = document.createTreeWalker(targetNode, NodeFilter.SHOW_TEXT, null);
    let node;
    while ((node = walker.nextNode())) {
      const len = node.nodeValue.length;
      if (remaining <= len) {
        const sel = window.getSelection();
        const r = document.createRange();
        r.setStart(node, Math.max(0, remaining));
        r.collapse(true);
        sel.removeAllRanges();
        sel.addRange(r);
        targetNode.closest('.smart-editable-page')?.focus();
        targetNode.scrollIntoView?.({ block: 'nearest' });
        return;
      }
      remaining -= len;
    }
    const sel = window.getSelection();
    const r = document.createRange();
    if (!(targetNode.textContent || '').trim()) {
      r.setStart(targetNode, 0);
      r.collapse(true);
    } else {
      r.selectNodeContents(targetNode);
      r.collapse(false);
    }
    sel.removeAllRanges();
    sel.addRange(r);
    targetNode.closest('.smart-editable-page')?.focus();
    targetNode.scrollIntoView?.({ block: 'nearest' });
  }

  function placeCaretAtStart(el) {
    if (!el) return;
    const sel = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(el);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    el.focus();
  }

  function placeCaretAtEnd(el) {
    if (!el) return;
    const sel = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(el);
    range.collapse(false);
    sel.removeAllRanges();
    sel.addRange(range);
    el.focus();
  }

  /**
   * Mesin Paginasi Otomatis (Auto-Overflow & Underflow Pull ala Microsoft Word):
   * - Jika konten melebihi tinggi cetak kertas, elemen paling bawah dipindah otomatis ke halaman berikutnya.
   * - Jika ada halaman berikutnya dan ruang cukup (dan bukan manualBreak), tarik elemen teratas kembali.
   */
  function reflowPagination() {
    let changed = false;
    let guard = 0;

    while (guard < 40) {
      guard++;
      let passMoved = false;
      const pageEls = [...pagesStack.querySelectorAll('.smart-editable-page')];

      for (let i = 0; i < pageEls.length; i++) {
        const curEl = pageEls[i];
        const maxH = curEl.clientHeight;
        if (!maxH || maxH < 40) continue;

        // A. OVERFLOW CHECK: apakah scrollHeight melebihi ruang kertas?
        if (curEl.scrollHeight > maxH + 3) {
          // Pastikan ada halaman i+1
          let nextEl = pageEls[i + 1];
          if (!nextEl) {
            const newWrapper = createPageDomElement(
              { id: `spage-${Date.now()}-${i + 2}`, html: '', isManualBreak: false },
              i + 1
            );
            pagesStack.appendChild(newWrapper);
            nextEl = newWrapper.querySelector('.smart-editable-page');
            // Bersihkan placeholder <p><br></p> awal di halaman baru
            if (nextEl.children.length === 1 && nextEl.firstElementChild?.tagName === 'P' && nextEl.firstElementChild.innerHTML === '<br>') {
              nextEl.innerHTML = '';
            }
          }

          if (curEl.children.length > 1) {
            const movingNode = curEl.lastElementChild;
            if ((movingNode?.tagName === 'TABLE' || movingNode?.tagName === 'UL' || movingNode?.tagName === 'OL') &&
                splitStructuredBlockTail(movingNode, curEl, nextEl)) {
              changed = true;
              passMoved = true;
              break;
            }
            const savedCaret = captureCaretInNode(movingNode);
            nextEl.insertBefore(movingNode, nextEl.firstChild);
            if (savedCaret) {
              restoreCaretInNode(movingNode, savedCaret);
            }
            changed = true;
            passMoved = true;
            break;
          } else if (curEl.firstElementChild) {
            // Kasus 1 elemen besar tunggal (misal 1 paragraf sangat panjang / 1 tabel banyak baris / list panjang)
            const splitOk = splitOversizedSingleBlock(curEl, nextEl, maxH);
            if (splitOk) {
              changed = true;
              passMoved = true;
              break;
            }
          }
        }

        // B. UNDERFLOW CHECK: tarik elemen pertama dari halaman i+1 bila bukan manualBreak
        const nextEl = pageEls[i + 1];
        if (nextEl && nextEl.dataset.manualBreak !== 'true' && nextEl.firstElementChild) {
          const candidate = nextEl.firstElementChild;
          const curSaved = captureCaretInNode(candidate);
          curEl.appendChild(candidate);
          if (curEl.scrollHeight <= maxH + 2) {
            if (curSaved) restoreCaretInNode(candidate, curSaved);
            changed = true;
            passMoved = true;
            // Jika halaman i+1 jadi kosong, hapus lembar halaman tersebut
            if (!(nextEl.innerText || '').trim() && !nextEl.querySelector('table, img')) {
              nextEl.closest('.smart-page-wrapper')?.remove();
            }
            break;
          } else {
            // Tidak muat -> kembalikan
            nextEl.insertBefore(candidate, nextEl.firstChild);
          }
        }
      }

      if (!passMoved) break;
    }

    // Hapus halaman kosong berlebih di paling belakang (kecuali sengaja ditambah via + Halaman)
    const finalPages = [...pagesStack.querySelectorAll('.smart-editable-page')];
    for (let i = finalPages.length - 1; i >= 1; i--) {
      const pEl = finalPages[i];
      const hasContent = Boolean((pEl.innerText || '').trim() || pEl.querySelector('table, img'));
      if (!hasContent && pEl.dataset.manualBreak !== 'true') {
        pEl.closest('.smart-page-wrapper')?.remove();
        changed = true;
      } else {
        break;
      }
    }

    if (changed) {
      reindexPageHeaders();
      updateStatusCounters();
      schedulePersistSmartDoc();
      refreshSmartFloatingQrOverlay();
    }
  }

  function splitStructuredBlockTail(blockEl, curEl, nextEl) {
    if (!blockEl) return false;
    if (blockEl.tagName === 'UL' || blockEl.tagName === 'OL') {
      if (blockEl.children.length <= 1) return false;
      const tag = blockEl.tagName.toLowerCase();
      let destList = nextEl.firstElementChild;
      if (!destList || destList.tagName !== blockEl.tagName || destList.dataset.splitGroup !== (blockEl.dataset.splitGroup || '')) {
        const groupId = blockEl.dataset.splitGroup || `slist-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
        blockEl.dataset.splitGroup = groupId;
        destList = document.createElement(tag);
        destList.dataset.splitGroup = groupId;
        nextEl.insertBefore(destList, nextEl.firstChild);
      }
      const lastLi = blockEl.lastElementChild;
      const caret = captureCaretInNode(lastLi);
      destList.insertBefore(lastLi, destList.firstChild);
      if (caret) restoreCaretInNode(lastLi, caret);
      return true;
    }

    if (blockEl.tagName === 'TABLE' && !blockEl.classList.contains('smart-sig-table')) {
      if (blockEl.querySelector('[rowspan]') || blockEl.classList.contains('smart-table-none')) {
        return false;
      }
      const tbody = blockEl.querySelector('tbody') || blockEl;
      const rows = [...tbody.querySelectorAll('tr')];
      if (rows.length <= 1) return false;
      const firstRowBottom = rows[0].getBoundingClientRect().top - curEl.getBoundingClientRect().top + rows[0].offsetHeight;
      if (curEl.children.length > 1 && firstRowBottom > curEl.clientHeight) {
        return false;
      }
      const groupId = blockEl.dataset.splitGroup || `stbl-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
      blockEl.dataset.splitGroup = groupId;
      let destTable = nextEl.firstElementChild;
      if (!destTable || destTable.tagName !== 'TABLE' || destTable.dataset.splitGroup !== groupId) {
        destTable = document.createElement('table');
        destTable.className = blockEl.className;
        destTable.dataset.splitGroup = groupId;
        const destBody = document.createElement('tbody');
        destTable.appendChild(destBody);
        nextEl.insertBefore(destTable, nextEl.firstChild);
      }
      const destBody = destTable.querySelector('tbody') || destTable;
      const movingRow = rows[rows.length - 1];
      const caret = captureCaretInNode(movingRow);
      destBody.insertBefore(movingRow, destBody.firstChild);
      if (caret) restoreCaretInNode(movingRow, caret);
      return true;
    }
    return false;
  }

  function splitOversizedSingleBlock(curEl, nextEl, maxH) {
    const child = curEl.firstElementChild;
    if (!child) return false;

    if (splitStructuredBlockTail(child, curEl, nextEl)) {
      return true;
    }

    if (child.tagName === 'P' || child.tagName === 'DIV') {
      const text = child.innerText || child.textContent || '';
      const words = text.split(/\s+/).filter(Boolean);
      if (words.length < 16) return false;
      const splitAt = Math.ceil(words.length * 0.65);
      const keepText = words.slice(0, splitAt).join(' ');
      const moveText = words.slice(splitAt).join(' ');
      child.textContent = keepText;
      const overflowP = document.createElement('p');
      overflowP.textContent = moveText;
      nextEl.insertBefore(overflowP, nextEl.firstChild);
      if (curEl.scrollHeight <= maxH + 2) return true;
      return true;
    }

    return false;
  }

  function scheduleReflowPagination() {
    clearTimeout(reflowTimer);
    reflowTimer = setTimeout(() => {
      reflowPagination();
    }, 90);
  }

  function bindEditablePageEvents(editable) {
    editable.addEventListener('focus', () => {
      activePageIndex = Number(editable.dataset.pageIndex || 0);
      updateStatusCounters();
    });

    editable.addEventListener('keyup', () => {
      const sel = window.getSelection();
      if (sel && sel.rangeCount > 0 && editable.contains(sel.anchorNode)) {
        lastSavedRange = sel.getRangeAt(0).cloneRange();
      }
    });

    editable.addEventListener('mouseup', () => {
      const sel = window.getSelection();
      if (sel && sel.rangeCount > 0 && editable.contains(sel.anchorNode)) {
        lastSavedRange = sel.getRangeAt(0).cloneRange();
      }
    });

    editable.addEventListener('input', () => {
      updateStatusCounters();
      scheduleReflowPagination();
      schedulePersistSmartDoc();
    });

    editable.addEventListener('blur', () => {
      clearTimeout(saveTimer);
      const livePages = collectLivePagesData();
      const latestState = store.getState().smartDocument || {};
      smartDoc = {
        ...smartDoc,
        pages: livePages,
        ...(latestState.verification ? { verification: latestState.verification } : {}),
        ...(latestState.qrPlacement ? { qrPlacement: latestState.qrPlacement } : {})
      };
      store.setState({ smartDocument: smartDoc }, false);
    });

    editable.addEventListener('paste', (e) => {
      const html = e.clipboardData?.getData('text/html');
      const plain = e.clipboardData?.getData('text/plain');
      if ((pastePlainTextMode || e.shiftKey) && plain) {
        e.preventDefault();
        pushUndoSnapshot();
        const paragraphs = plain
          .split(/\r?\n\s*\r?\n/)
          .map(p => `<p>${escapeHtmlText(p.trim()).replace(/\r?\n/g, '<br>')}</p>`)
          .join('');
        document.execCommand('insertHTML', false, paragraphs);
        scheduleReflowPagination();
        schedulePersistSmartDoc();
        return;
      }
      if (html) {
        e.preventDefault();
        pushUndoSnapshot();
        const sanitized = sanitizePastedWordHtml(html);
        document.execCommand('insertHTML', false, sanitized);
        scheduleReflowPagination();
        schedulePersistSmartDoc();
        showToast('✨ Format tempelan dirapikan otomatis ke gaya standar surat.', 'info');
      } else if (plain) {
        e.preventDefault();
        pushUndoSnapshot();
        const paragraphs = plain
          .split(/\r?\n\s*\r?\n/)
          .map(p => `<p>${escapeHtmlText(p.trim()).replace(/\r?\n/g, '<br>')}</p>`)
          .join('');
        document.execCommand('insertHTML', false, paragraphs);
        scheduleReflowPagination();
        schedulePersistSmartDoc();
      }
    });

    editable.addEventListener('keydown', (e) => {
      // Ctrl+Enter -> Tambah Halaman Baru (Page Break seperti MS Word)
      if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault();
        insertManualPageBreakAtSelection(editable);
        return;
      }

      // Tab di dalam tabel / paragraf
      if (e.key === 'Tab') {
        const sel = window.getSelection();
        const td = sel?.anchorNode?.parentElement?.closest('td, th');
        if (td) {
          const table = td.closest('table');
          const allCells = table ? [...table.querySelectorAll('td, th')] : [];
          const curIdx = allCells.indexOf(td);
          if (curIdx >= 0) {
            e.preventDefault();
            const targetCell = e.shiftKey ? allCells[curIdx - 1] : allCells[curIdx + 1];
            if (targetCell) {
              placeCaretAtEnd(targetCell);
            } else if (!e.shiftKey && table && !table.classList.contains('smart-sig-table')) {
              const lastTr = td.closest('tr');
              const addedRow = insertRowRelative?.(table, lastTr, 'below');
              const firstNewCell = addedRow?.querySelector('td, th');
              if (firstNewCell) placeCaretAtStart(firstNewCell);
            }
          }
          return;
        }
        e.preventDefault();
        document.execCommand('insertHTML', false, '&nbsp;&nbsp;&nbsp;&nbsp;');
        return;
      }

      // Backspace di awal halaman > 0 -> gabungkan kembali dengan halaman sebelumnya
      if (e.key === 'Backspace') {
        const curIdx = Number(editable.dataset.pageIndex || 0);
        if (curIdx > 0 && isCaretAtStartOfElement(editable)) {
          e.preventDefault();
          const wrappers = [...pagesStack.querySelectorAll('.smart-page-wrapper')];
          const prevEditable = wrappers[curIdx - 1]?.querySelector('.smart-editable-page');
          if (prevEditable) {
            pushUndoSnapshot();
            editable.dataset.manualBreak = 'false';
            const isEmpty = !(editable.innerText || '').trim() && !editable.querySelector('table, img');
            if (isEmpty) {
              editable.closest('.smart-page-wrapper')?.remove();
              placeCaretAtEnd(prevEditable);
            } else {
              const prevEnd = prevEditable.lastElementChild || prevEditable;
              while (editable.firstChild) {
                prevEditable.appendChild(editable.firstChild);
              }
              editable.closest('.smart-page-wrapper')?.remove();
              placeCaretAtEnd(prevEnd);
            }
            reindexPageHeaders();
            scheduleReflowPagination();
            schedulePersistSmartDoc();
          }
        }
      }
    });
  }

  function isCaretAtStartOfElement(containerEl) {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0 || !sel.isCollapsed) return false;
    const range = sel.getRangeAt(0);
    if (!containerEl.contains(range.startContainer)) return false;
    const pre = document.createRange();
    pre.selectNodeContents(containerEl);
    pre.setEnd(range.startContainer, range.startOffset);
    return pre.toString().length === 0;
  }

  /**
   * Memecah halaman aktif tepat pada posisi kursor (atau menambah halaman baru kosong di belakang).
   */
  function insertManualPageBreakAtSelection(targetEditable = null) {
    pushUndoSnapshot();
    const wrappers = [...pagesStack.querySelectorAll('.smart-page-wrapper')];
    const activeEditable =
      targetEditable ||
      wrappers[activePageIndex]?.querySelector('.smart-editable-page') ||
      wrappers[wrappers.length - 1]?.querySelector('.smart-editable-page');

    if (!activeEditable) return;
    const curWrapper = activeEditable.closest('.smart-page-wrapper');
    const curIdx = wrappers.indexOf(curWrapper);

    let trailingHtml = '<p><br></p>';
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0 && activeEditable.contains(sel.anchorNode)) {
      const range = sel.getRangeAt(0);
      const afterRange = document.createRange();
      afterRange.setStart(range.endContainer, range.endOffset);
      const endTarget = activeEditable.lastChild || activeEditable;
      if (endTarget.nodeType === 3) {
        afterRange.setEnd(endTarget, endTarget.nodeValue.length);
      } else {
        afterRange.setEndAfter(endTarget);
      }
      const frag = afterRange.extractContents();
      const tmp = document.createElement('div');
      tmp.appendChild(frag);
      const extractedHtml = tmp.innerHTML.trim();
      if (extractedHtml) {
        trailingHtml = extractedHtml;
      }
    }

    const newPageData = {
      id: `spage-${Date.now()}-${curIdx + 2}`,
      html: trailingHtml,
      isManualBreak: true
    };
    const newWrapper = createPageDomElement(newPageData, curIdx + 1);
    curWrapper.after(newWrapper);
    reindexPageHeaders();
    updateStatusCounters();
    schedulePersistSmartDoc();

    const newEditable = newWrapper.querySelector('.smart-editable-page');
    if (newEditable) {
      placeCaretAtStart(newEditable);
      newWrapper.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
    showToast(`Halaman ${curIdx + 2} ditambahkan.`, 'info');
  }

  function sanitizePastedWordHtml(rawHtml = '') {
    if (typeof DOMParser === 'undefined') return rawHtml;
    const doc = new DOMParser().parseFromString(rawHtml, 'text/html');
    const allowedTags = new Set([
      'p', 'div', 'br', 'h1', 'h2', 'h3', 'h4',
      'b', 'strong', 'i', 'em', 'u', 's', 'strike', 'sub', 'sup',
      'ul', 'ol', 'li', 'table', 'colgroup', 'col', 'thead', 'tbody', 'tr', 'th', 'td', 'span'
    ]);

    const cleanNode = (node) => {
      if (node.nodeType === 3) return document.createTextNode(node.nodeValue);
      if (node.nodeType !== 1) return document.createTextNode('');
      const tag = node.tagName.toLowerCase();
      if (['script', 'style', 'meta', 'link', 'xml', 'o:p'].includes(tag)) {
        return document.createTextNode('');
      }
      if (!allowedTags.has(tag)) {
        const frag = document.createDocumentFragment();
        [...node.childNodes].forEach(c => frag.appendChild(cleanNode(c)));
        return frag;
      }
      if (tag === 'span') {
        const frag = document.createDocumentFragment();
        [...node.childNodes].forEach(c => frag.appendChild(cleanNode(c)));
        return frag;
      }
      if (tag === 'div') {
        const hasBlockKids = [...node.childNodes].some(
          c => c.nodeType === 1 && /^(p|div|table|ul|ol|h[1-6])$/i.test(c.tagName)
        );
        if (hasBlockKids) {
          const frag = document.createDocumentFragment();
          [...node.childNodes].forEach(c => frag.appendChild(cleanNode(c)));
          return frag;
        }
      }
      const el = document.createElement(tag === 'div' ? 'p' : tag);
      const align = node.style?.textAlign;
      if (align && ['left', 'center', 'right', 'justify'].includes(align)) {
        el.style.textAlign = align;
      }
      if (tag === 'td' || tag === 'th') {
        const cs = parseInt(node.getAttribute('colspan') || '1', 10);
        const rs = parseInt(node.getAttribute('rowspan') || '1', 10);
        if (cs > 1) el.setAttribute('colspan', String(cs));
        if (rs > 1) el.setAttribute('rowspan', String(rs));
        const va = node.style?.verticalAlign;
        if (va && ['top', 'middle', 'bottom'].includes(va)) {
          el.style.verticalAlign = va;
        }
      }
      if (tag === 'col' && node.style?.width) {
        el.style.width = node.style.width;
      }
      if (tag === 'table') el.className = 'smart-word-table';
      [...node.childNodes].forEach(c => el.appendChild(cleanNode(c)));
      return el;
    };

    const out = document.createElement('div');
    [...doc.body.childNodes].forEach(c => out.appendChild(cleanNode(c)));
    return out.innerHTML;
  }

  // Render halaman awal
  renderAllSmartPages(smartDoc.pages);

  // =========================================================================
  // TOOLBAR FORMATTING, RIBBON & MODALS HANDLERS
  // =========================================================================

  const tableCtxBar = container.querySelector('#smart-table-context-bar');
  const smCtxBorderSel = container.querySelector('#sm-ctx-border');
  const smCtxRowHeightSel = container.querySelector('#sm-ctx-row-height');
  const smBtnMergeSel = container.querySelector('#sm-btn-merge-sel');
  let activeTableCell = null;
  let tableSelAnchorCell = null;
  let selectedTableCells = [];

  function clearTableCellSelection() {
    container.querySelectorAll('.smart-cell-selected').forEach(c => c.classList.remove('smart-cell-selected'));
    selectedTableCells = [];
    if (smBtnMergeSel) smBtnMergeSel.disabled = true;
  }

  function buildLogicalGrid(tableEl) {
    const trList = [...tableEl.querySelectorAll('tr')];
    const occ = [];
    const cellMap = new Map();
    let totalCols = 0;

    trList.forEach((tr, r) => {
      occ[r] = occ[r] || [];
      let c = 0;
      const cells = [...tr.children].filter(el => el.tagName === 'TD' || el.tagName === 'TH');
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

  function syncColGroupEvenly(tableEl, colCount) {
    if (colCount <= 0) return;
    const widths = Array.from({ length: colCount }, () => Math.round((100 / colCount) * 10) / 10);
    let colgroup = [...tableEl.children].find(ch => ch.tagName === 'COLGROUP');
    if (colgroup) {
      colgroup.innerHTML = widths.map(w => `<col style="width: ${w}%;">`).join('');
      tableEl.setAttribute('data-col-widths', widths.join(','));
    } else if (tableEl.getAttribute('data-col-widths')) {
      tableEl.setAttribute('data-col-widths', widths.join(','));
    }
  }

  function insertRowRelative(tableEl, trEl, position = 'below') {
    if (!tableEl || !trEl) return null;
    pushUndoSnapshot();
    const { trList, occ, totalCols } = buildLogicalGrid(tableEl);
    const rIdx = trList.indexOf(trEl);
    if (rIdx < 0) return null;

    const boundaryTop = position === 'above' ? rIdx - 1 : rIdx;
    const spannedCols = new Set();
    if (boundaryTop >= 0 && occ[boundaryTop] && occ[boundaryTop + 1]) {
      for (let c = 0; c < totalCols; c++) {
        const upper = occ[boundaryTop][c];
        if (upper && upper.startR + upper.rowspan - 1 > boundaryTop) {
          if (!spannedCols.has(c)) {
            const newRs = upper.rowspan + 1;
            upper.cell.setAttribute('rowspan', String(newRs));
            for (let dc = 0; dc < upper.colspan; dc++) {
              spannedCols.add(upper.startC + dc);
            }
          }
        }
      }
    }

    const useHeaderTag = Boolean(trEl.closest('thead')) && !(position === 'below' && !trEl.nextElementSibling && tableEl.querySelector('tbody'));
    const cellTag = useHeaderTag ? 'th' : 'td';
    const neededCols = Math.max(1, totalCols - spannedCols.size);
    const newTr = document.createElement('tr');
    if (useHeaderTag) newTr.setAttribute('data-header', 'true');
    for (let i = 0; i < neededCols; i++) {
      const cell = document.createElement(cellTag);
      cell.innerHTML = '<br>';
      newTr.appendChild(cell);
    }

    if (position === 'above') {
      trEl.parentElement.insertBefore(newTr, trEl);
    } else if (trEl.closest('thead') && !trEl.nextElementSibling && tableEl.querySelector('tbody')) {
      const tbody = tableEl.querySelector('tbody');
      tbody.insertBefore(newTr, tbody.firstChild);
    } else {
      trEl.parentElement.insertBefore(newTr, trEl.nextSibling);
    }
    updateStatusCounters();
    scheduleReflowPagination();
    schedulePersistSmartDoc();
    return newTr;
  }

  function deleteRowAt(tableEl, trEl) {
    if (!tableEl || !trEl) return;
    pushUndoSnapshot();
    const { trList, occ, totalCols } = buildLogicalGrid(tableEl);
    const rIdx = trList.indexOf(trEl);
    if (rIdx < 0) return;

    if (trList.length <= 1) {
      tableEl.remove();
      clearTableCellSelection();
      updateTableContextBar(null);
      scheduleReflowPagination();
      schedulePersistSmartDoc();
      return;
    }

    const handled = new Set();
    for (let c = 0; c < totalCols; c++) {
      const slot = occ[rIdx]?.[c];
      if (!slot || handled.has(slot.cell)) continue;
      handled.add(slot.cell);
      if (slot.startR < rIdx) {
        const nextRs = slot.rowspan - 1;
        if (nextRs <= 1) slot.cell.removeAttribute('rowspan');
        else slot.cell.setAttribute('rowspan', String(nextRs));
      } else if (slot.startR === rIdx && slot.rowspan > 1 && trList[rIdx + 1]) {
        const nextTr = trList[rIdx + 1];
        const nextRs = slot.rowspan - 1;
        if (nextRs <= 1) slot.cell.removeAttribute('rowspan');
        else slot.cell.setAttribute('rowspan', String(nextRs));
        let anchorAfter = null;
        for (let k = slot.startC + slot.colspan; k < totalCols; k++) {
          const cand = occ[rIdx + 1]?.[k];
          if (cand && cand.startR === rIdx + 1) {
            anchorAfter = cand.cell;
            break;
          }
        }
        nextTr.insertBefore(slot.cell, anchorAfter);
      }
    }

    trEl.remove();
    clearTableCellSelection();
    updateTableContextBar(null);
    scheduleReflowPagination();
    schedulePersistSmartDoc();
  }

  function insertColRight(tableEl, cellEl) {
    if (!tableEl || !cellEl) return;
    pushUndoSnapshot();
    const { trList, occ, totalCols, cellMap } = buildLogicalGrid(tableEl);
    const info = cellMap.get(cellEl);
    if (!info) return;
    const insertAfterCol = info.startC + info.colspan - 1;
    const handledSpan = new Set();

    trList.forEach((tr, r) => {
      const slotAtCol = occ[r]?.[insertAfterCol];
      if (slotAtCol && slotAtCol.startC + slotAtCol.colspan - 1 > insertAfterCol) {
        if (!handledSpan.has(slotAtCol.cell)) {
          handledSpan.add(slotAtCol.cell);
          slotAtCol.cell.setAttribute('colspan', String(slotAtCol.colspan + 1));
        }
        return;
      }
      const tag = (tr.closest('thead') || tr.dataset.header === 'true') ? 'th' : 'td';
      const newCell = document.createElement(tag);
      newCell.innerHTML = '<br>';
      let nextAnchor = null;
      for (let k = insertAfterCol + 1; k < totalCols; k++) {
        const cand = occ[r]?.[k];
        if (cand && cand.startR === r) {
          nextAnchor = cand.cell;
          break;
        }
      }
      tr.insertBefore(newCell, nextAnchor);
    });

    syncColGroupEvenly(tableEl, totalCols + 1);
    scheduleReflowPagination();
    schedulePersistSmartDoc();
  }

  function deleteColAt(tableEl, cellEl) {
    if (!tableEl || !cellEl) return;
    pushUndoSnapshot();
    const { trList, occ, totalCols, cellMap } = buildLogicalGrid(tableEl);
    if (totalCols <= 1) {
      tableEl.remove();
      clearTableCellSelection();
      updateTableContextBar(null);
      scheduleReflowPagination();
      schedulePersistSmartDoc();
      return;
    }
    const info = cellMap.get(cellEl);
    if (!info) return;
    const targetCol = info.startC;
    const handled = new Set();

    trList.forEach((_, r) => {
      const slot = occ[r]?.[targetCol];
      if (!slot || handled.has(slot.cell)) return;
      handled.add(slot.cell);
      if (slot.colspan > 1) {
        const nextCs = slot.colspan - 1;
        if (nextCs <= 1) slot.cell.removeAttribute('colspan');
        else slot.cell.setAttribute('colspan', String(nextCs));
      } else {
        slot.cell.remove();
      }
    });

    syncColGroupEvenly(tableEl, totalCols - 1);
    clearTableCellSelection();
    updateTableContextBar(null);
    scheduleReflowPagination();
    schedulePersistSmartDoc();
  }

  function mergeCellRange(tableEl, targetCells) {
    if (!tableEl || !Array.isArray(targetCells) || targetCells.length < 2) return false;
    const { occ, cellMap } = buildLogicalGrid(tableEl);
    const infos = targetCells.map(c => cellMap.get(c)).filter(Boolean);
    if (infos.length < 2) return false;

    let minR = Math.min(...infos.map(i => i.startR));
    let maxR = Math.max(...infos.map(i => i.startR + i.rowspan - 1));
    let minC = Math.min(...infos.map(i => i.startC));
    let maxC = Math.max(...infos.map(i => i.startC + i.colspan - 1));

    let expanded = true;
    let guard = 0;
    while (expanded && guard < 20) {
      expanded = false;
      guard++;
      for (let r = minR; r <= maxR; r++) {
        for (let c = minC; c <= maxC; c++) {
          const s = occ[r]?.[c];
          if (!s) continue;
          if (s.startR < minR) { minR = s.startR; expanded = true; }
          if (s.startR + s.rowspan - 1 > maxR) { maxR = s.startR + s.rowspan - 1; expanded = true; }
          if (s.startC < minC) { minC = s.startC; expanded = true; }
          if (s.startC + s.colspan - 1 > maxC) { maxC = s.startC + s.colspan - 1; expanded = true; }
        }
      }
    }

    const uniqueCells = [];
    const seen = new Set();
    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        const s = occ[r]?.[c];
        if (s && !seen.has(s.cell)) {
          seen.add(s.cell);
          uniqueCells.push(s.cell);
        }
      }
    }
    if (uniqueCells.length < 2) return false;

    pushUndoSnapshot();
    const anchor = uniqueCells[0];
    const parts = [];
    uniqueCells.forEach((cell) => {
      const txt = (cell.innerText || cell.textContent || '').trim();
      if (txt) {
        parts.push(cell.innerHTML.replace(/<br>\s*$/i, ''));
      }
    });
    if (parts.length > 0) {
      anchor.innerHTML = parts.join('<br>');
    }
    uniqueCells.slice(1).forEach(c => c.remove());

    const cs = maxC - minC + 1;
    const rs = maxR - minR + 1;
    if (cs > 1) anchor.setAttribute('colspan', String(cs));
    else anchor.removeAttribute('colspan');
    if (rs > 1) anchor.setAttribute('rowspan', String(rs));
    else anchor.removeAttribute('rowspan');

    clearTableCellSelection();
    activeTableCell = anchor;
    placeCaretAtEnd(anchor);
    updateTableContextBar(anchor);
    scheduleReflowPagination();
    schedulePersistSmartDoc();
    return true;
  }

  function unmergeCell(tableEl, cellEl) {
    if (!tableEl || !cellEl) return false;
    const { trList, occ, totalCols, cellMap } = buildLogicalGrid(tableEl);
    const info = cellMap.get(cellEl);
    if (!info || (info.colspan <= 1 && info.rowspan <= 1)) {
      showToast('Sel aktif tidak dalam kondisi tergabung.', 'info');
      return false;
    }

    pushUndoSnapshot();
    const { startR, startC, colspan, rowspan } = info;
    cellEl.removeAttribute('colspan');
    cellEl.removeAttribute('rowspan');

    for (let r = startR; r < startR + rowspan; r++) {
      const tr = trList[r];
      if (!tr) continue;
      const tag = (tr.closest('thead') || tr.dataset.header === 'true') ? 'th' : 'td';
      const countNeeded = r === startR ? colspan - 1 : colspan;
      let anchorAfter = null;
      for (let k = startC + colspan; k < totalCols; k++) {
        const cand = occ[r]?.[k];
        if (cand && cand.startR === r) {
          anchorAfter = cand.cell;
          break;
        }
      }
      for (let n = 0; n < countNeeded; n++) {
        const newCell = document.createElement(tag);
        newCell.innerHTML = '<br>';
        tr.insertBefore(newCell, anchorAfter);
      }
    }

    clearTableCellSelection();
    placeCaretAtEnd(cellEl);
    updateTableContextBar(cellEl);
    scheduleReflowPagination();
    schedulePersistSmartDoc();
    return true;
  }

  function updateTableContextBar(cellEl) {
    if (!tableCtxBar) return;
    const tableEl = cellEl?.closest('table');
    if (!tableEl || tableEl.classList.contains('smart-sig-table')) {
      activeTableCell = null;
      tableCtxBar.classList.add('hidden');
      return;
    }
    activeTableCell = cellEl;
    tableCtxBar.classList.remove('hidden');

    if (smCtxBorderSel) {
      const curBorder = tableEl.getAttribute('data-border-style')
        || (tableEl.classList.contains('smart-table-sppd') ? 'sppd' : tableEl.classList.contains('smart-table-none') ? 'none' : 'grid');
      smCtxBorderSel.value = curBorder;
    }
    if (smCtxRowHeightSel) {
      const tr = cellEl.closest('tr');
      const curH = tr?.getAttribute('data-min-height') || '0';
      const matchingOpt = [...smCtxRowHeightSel.options].find(o => o.value === curH);
      smCtxRowHeightSel.value = matchingOpt ? curH : '0';
    }
    if (smBtnMergeSel) {
      smBtnMergeSel.disabled = selectedTableCells.length < 2;
      smBtnMergeSel.textContent = selectedTableCells.length >= 2
        ? `⧉ Gabung (${selectedTableCells.length} Sel)`
        : '⧉ Gabung Seleksi';
    }
  }

  const onSmartSelectionChange = () => {
    if (!container.isConnected) {
      document.removeEventListener('selectionchange', onSmartSelectionChange);
      return;
    }
    const sel = window.getSelection();
    const anchor = sel?.anchorNode;
    if (!anchor || !pagesStack.contains(anchor)) return;
    const cellEl = (anchor.nodeType === 1 ? anchor : anchor.parentElement)?.closest('td, th');
    updateTableContextBar(cellEl || null);
  };
  document.addEventListener('selectionchange', onSmartSelectionChange);

  pagesStack.addEventListener('mousedown', (e) => {
    if (e.target?.closest?.('.floating-qr-box, .floating-qr-toolbar, .floating-qr-resize, .floating-qr-stage, .fqr-context-menu')) {
      return;
    }
    const cellEl = e.target?.closest?.('td, th');
    const tableEl = cellEl?.closest?.('table');
    if (!cellEl || !tableEl || tableEl.classList.contains('smart-sig-table')) {
      if (!e.shiftKey) clearTableCellSelection();
      return;
    }
    if (e.shiftKey && tableSelAnchorCell && tableSelAnchorCell.closest('table') === tableEl) {
      e.preventDefault();
      const { occ, cellMap } = buildLogicalGrid(tableEl);
      const a = cellMap.get(tableSelAnchorCell);
      const b = cellMap.get(cellEl);
      if (a && b) {
        clearTableCellSelection();
        const minR = Math.min(a.startR, b.startR);
        const maxR = Math.max(a.startR + a.rowspan - 1, b.startR + b.rowspan - 1);
        const minC = Math.min(a.startC, b.startC);
        const maxC = Math.max(a.startC + a.colspan - 1, b.startC + b.colspan - 1);
        const seen = new Set();
        for (let r = minR; r <= maxR; r++) {
          for (let c = minC; c <= maxC; c++) {
            const s = occ[r]?.[c];
            if (s && !seen.has(s.cell)) {
              seen.add(s.cell);
              s.cell.classList.add('smart-cell-selected');
              selectedTableCells.push(s.cell);
            }
          }
        }
      }
      updateTableContextBar(cellEl);
      return;
    }
    clearTableCellSelection();
    tableSelAnchorCell = cellEl;
    updateTableContextBar(cellEl);
  });

  tableCtxBar?.querySelectorAll('[data-tbl-act]').forEach((btn) => {
    btn.addEventListener('mousedown', e => e.preventDefault());
    btn.addEventListener('click', () => {
      const cell = activeTableCell;
      const table = cell?.closest('table');
      if (!cell || !table) return;
      const act = btn.dataset.tblAct;
      const tr = cell.closest('tr');

      if (act === 'row-above') {
        const added = insertRowRelative(table, tr, 'above');
        if (added?.firstElementChild) placeCaretAtStart(added.firstElementChild);
      } else if (act === 'row-below') {
        const added = insertRowRelative(table, tr, 'below');
        if (added?.firstElementChild) placeCaretAtStart(added.firstElementChild);
      } else if (act === 'row-del') {
        deleteRowAt(table, tr);
      } else if (act === 'col-right') {
        insertColRight(table, cell);
      } else if (act === 'col-del') {
        deleteColAt(table, cell);
      } else if (act === 'merge-sel') {
        if (!mergeCellRange(table, selectedTableCells)) {
          showToast('Shift+Klik minimal 2 sel bersebelahan terlebih dahulu.', 'info');
        }
      } else if (act === 'merge-right') {
        const { occ, cellMap } = buildLogicalGrid(table);
        const info = cellMap.get(cell);
        const neighbor = info ? occ[info.startR]?.[info.startC + info.colspan] : null;
        if (neighbor && neighbor.cell !== cell) mergeCellRange(table, [cell, neighbor.cell]);
        else showToast('Tidak ada sel sebelah kanan yang bisa digabung.', 'info');
      } else if (act === 'merge-down') {
        const { occ, cellMap } = buildLogicalGrid(table);
        const info = cellMap.get(cell);
        const neighbor = info ? occ[info.startR + info.rowspan]?.[info.startC] : null;
        if (neighbor && neighbor.cell !== cell) mergeCellRange(table, [cell, neighbor.cell]);
        else showToast('Tidak ada sel di bawahnya yang bisa digabung.', 'info');
      } else if (act === 'unmerge') {
        unmergeCell(table, cell);
      } else if (act === 'toggle-header-bg') {
        pushUndoSnapshot();
        const currentlyNoBg = table.classList.contains('no-header-bg') || table.getAttribute('data-header-shading') === 'false';
        if (currentlyNoBg) {
          table.classList.remove('no-header-bg');
          table.removeAttribute('data-header-shading');
        } else {
          table.classList.add('no-header-bg');
          table.setAttribute('data-header-shading', 'false');
        }
        schedulePersistSmartDoc();
      } else if (act === 'align-left' || act === 'align-center' || act === 'align-right') {
        pushUndoSnapshot();
        const alignVal = act.replace('align-', '');
        const targets = selectedTableCells.length > 0 ? selectedTableCells : [cell];
        targets.forEach(c => {
          c.style.textAlign = alignVal;
          c.querySelectorAll('p').forEach(p => { p.style.textAlign = alignVal; });
        });
        schedulePersistSmartDoc();
      } else if (act === 'del-table') {
        pushUndoSnapshot();
        table.remove();
        clearTableCellSelection();
        updateTableContextBar(null);
        scheduleReflowPagination();
        schedulePersistSmartDoc();
      }
    });
  });

  smCtxBorderSel?.addEventListener('change', (e) => {
    const table = activeTableCell?.closest('table');
    if (!table) return;
    pushUndoSnapshot();
    const mode = e.target.value;
    table.classList.remove('smart-table-sppd', 'smart-table-none');
    if (mode === 'sppd') {
      table.classList.add('smart-word-table', 'smart-table-sppd');
      table.setAttribute('data-border-style', 'sppd');
    } else if (mode === 'none') {
      table.classList.add('smart-word-table', 'smart-table-none');
      table.setAttribute('data-border-style', 'none');
    } else {
      table.classList.add('smart-word-table');
      table.removeAttribute('data-border-style');
    }
    schedulePersistSmartDoc();
  });

  smCtxRowHeightSel?.addEventListener('change', (e) => {
    const tr = activeTableCell?.closest('tr');
    if (!tr) return;
    pushUndoSnapshot();
    const cm = parseFloat(e.target.value || '0') || 0;
    if (cm > 0) {
      tr.setAttribute('data-min-height', String(cm));
      tr.style.height = `${cm}cm`;
    } else {
      tr.removeAttribute('data-min-height');
      tr.style.height = '';
    }
    scheduleReflowPagination();
    schedulePersistSmartDoc();
  });

  // Hubungkan Move Handle (panah 4-arah pojok kiri atas), Resize Cursor (col/row), & Drag Boundary MS Word
  const smartDeskEl = container.querySelector('#smart-word-desk');
  const smartTableInteraction = attachInteractiveTables({
    deskEl: smartDeskEl,
    tableSelector: '.smart-editable-page table',
    excludeTableSelector: '.smart-sig-table',
    getMoveUnitEl: (tbl) => tbl,
    getDropTargets: (draggedTbl) =>
      [...pagesStack.querySelectorAll('.smart-editable-page > *')].filter(
        el => el !== draggedTbl && !draggedTbl.contains(el)
      ),
    performMove: (unitEl, targetEl, position) => {
      unitEl.removeAttribute('data-split-group');
      targetEl.insertAdjacentElement(position === 'before' ? 'beforebegin' : 'afterend', unitEl);
    },
    onSelectWholeTable: (tbl) => {
      const firstCell = tbl.querySelector('th, td');
      if (firstCell) updateTableContextBar(firstCell);
      showToast('Seluruh isi tabel dipilih — klik & seret ikon panah 4-arah untuk memindahkan, Ctrl+C salin, atau Delete hapus.', 'info');
    },
    onBeforeMutate: () => pushUndoSnapshot(),
    onMutate: (kind) => {
      if (kind === 'move' || kind === 'delete-table') clearTableCellSelection();
      if (kind === 'delete-table') updateTableContextBar(null);
      updateStatusCounters();
      scheduleReflowPagination();
      schedulePersistSmartDoc();
    }
  });

  function execSmartCommand(cmd, value = null) {
    const wrappers = [...pagesStack.querySelectorAll('.smart-page-wrapper')];
    const targetEditable = wrappers[activePageIndex]?.querySelector('.smart-editable-page')
      || wrappers[0]?.querySelector('.smart-editable-page');
    if (targetEditable && lastSavedRange && targetEditable.contains(lastSavedRange.startContainer)) {
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(lastSavedRange);
    } else if (targetEditable) {
      targetEditable.focus();
    }

    pushUndoSnapshot();
    document.execCommand(cmd, false, value);
    updateStatusCounters();
    scheduleReflowPagination();
    schedulePersistSmartDoc();
  }

  container.querySelectorAll('[data-cmd]').forEach((btn) => {
    btn.addEventListener('mousedown', (e) => e.preventDefault());
    btn.addEventListener('click', () => {
      const cmd = btn.dataset.cmd;
      execSmartCommand(cmd);
    });
  });

  function syncSmartRibbonFromOptions(nextOpts = {}) {
    const o = normalizePageSetupOptions(nextOpts, store.getLetterOptions());
    const pSel = container.querySelector('#smart-paper-select');
    if (pSel) pSel.value = o.paperSize;
    const oSel = container.querySelector('#smart-orientation-select');
    if (oSel) oSel.value = o.orientation;
    const fSel = container.querySelector('#smart-font-family');
    if (fSel) {
      if (![...fSel.options].some(opt => opt.value === o.fontFamily)) {
        const optEl = document.createElement('option');
        optEl.value = o.fontFamily;
        optEl.textContent = o.fontFamily;
        fSel.appendChild(optEl);
      }
      fSel.value = o.fontFamily;
    }
    const szSel = container.querySelector('#smart-font-size');
    if (szSel) {
      if (![...szSel.options].some(opt => opt.value === String(o.fontSize))) {
        const optEl = document.createElement('option');
        optEl.value = String(o.fontSize);
        optEl.textContent = String(o.fontSize);
        szSel.appendChild(optEl);
      }
      szSel.value = String(o.fontSize);
    }
    const lsSel = container.querySelector('#smart-line-spacing');
    if (lsSel) lsSel.value = String(o.lineSpacing);
    renderWordHeaderToggleButton(o.showKop !== false);
  }

  // Block Style (Normal / H1 / H2 / H3)
  container.querySelector('#smart-block-style')?.addEventListener('change', (e) => {
    execSmartCommand('formatBlock', `<${e.target.value}>`);
  });

  // Jenis Penomoran Daftar Lengkap Word (1,2,3 / a,b,c / A,B,C / i,ii,iii / I,II,III / Bullet)
  container.querySelector('#smart-list-style-select')?.addEventListener('change', (e) => {
    const val = e.target.value;
    if (!val) return;
    const spec = getListStyleSpec(val, true);
    const wrappers = [...pagesStack.querySelectorAll('.smart-page-wrapper')];
    const targetEditable = wrappers[activePageIndex]?.querySelector('.smart-editable-page')
      || wrappers[0]?.querySelector('.smart-editable-page');
    if (targetEditable && lastSavedRange && targetEditable.contains(lastSavedRange.startContainer)) {
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(lastSavedRange);
    } else if (targetEditable) {
      targetEditable.focus();
    }

    pushUndoSnapshot();
    const sel = window.getSelection();
    let listEl = (sel?.anchorNode?.nodeType === 1 ? sel.anchorNode : sel?.anchorNode?.parentElement)?.closest('ol, ul');
    if (!listEl || !pagesStack.contains(listEl)) {
      document.execCommand(spec.htmlTag === 'ul' ? 'insertUnorderedList' : 'insertOrderedList', false, null);
      const afterSel = window.getSelection();
      listEl = (afterSel?.anchorNode?.nodeType === 1 ? afterSel.anchorNode : afterSel?.anchorNode?.parentElement)?.closest('ol, ul');
    }
    if (listEl && pagesStack.contains(listEl)) {
      if (listEl.tagName.toLowerCase() !== spec.htmlTag) {
        const replacement = document.createElement(spec.htmlTag);
        while (listEl.firstChild) replacement.appendChild(listEl.firstChild);
        listEl.replaceWith(replacement);
        listEl = replacement;
      }
      listEl.setAttribute('data-list-style', spec.key);
      listEl.style.listStyleType = spec.cssType;
      if (spec.htmlType) listEl.setAttribute('type', spec.htmlType);
      else listEl.removeAttribute('type');
    }
    e.target.value = '';
    updateStatusCounters();
    scheduleReflowPagination();
    schedulePersistSmartDoc();
  });

  // Font & Typography Options
  container.querySelector('#smart-paper-select')?.addEventListener('change', (e) => {
    store.updateLetterOptions({ paperSize: e.target.value });
    syncGlobalPrintPageStyle(store.getLetterOptions());
    renderAllSmartPages(collectLivePagesData());
    showToast(`Ukuran kertas diatur ke ${PAPER_LABELS[e.target.value] || e.target.value.toUpperCase()}`, 'info');
  });

  container.querySelector('#smart-orientation-select')?.addEventListener('change', (e) => {
    store.updateLetterOptions({ orientation: e.target.value });
    syncGlobalPrintPageStyle(store.getLetterOptions());
    renderAllSmartPages(collectLivePagesData());
  });

  container.querySelector('#smart-font-family')?.addEventListener('change', (e) => {
    store.updateLetterOptions({ fontFamily: e.target.value });
    renderAllSmartPages(collectLivePagesData());
  });

  container.querySelector('#smart-font-size')?.addEventListener('change', (e) => {
    store.updateLetterOptions({ fontSize: e.target.value });
    renderAllSmartPages(collectLivePagesData());
  });

  container.querySelector('#smart-line-spacing')?.addEventListener('change', (e) => {
    store.updateLetterOptions({ lineSpacing: e.target.value });
    renderAllSmartPages(collectLivePagesData());
  });

  container.querySelector('#smart-btn-font-up')?.addEventListener('click', () => {
    const sel = container.querySelector('#smart-font-size');
    const nextIdx = Math.min(sel.selectedIndex + 1, sel.options.length - 1);
    sel.selectedIndex = nextIdx;
    store.updateLetterOptions({ fontSize: sel.value });
    renderAllSmartPages(collectLivePagesData());
  });

  container.querySelector('#smart-btn-font-down')?.addEventListener('click', () => {
    const sel = container.querySelector('#smart-font-size');
    const prevIdx = Math.max(sel.selectedIndex - 1, 0);
    sel.selectedIndex = prevIdx;
    store.updateLetterOptions({ fontSize: sel.value });
    renderAllSmartPages(collectLivePagesData());
  });

  // Change Case siklus: UPPER -> lower -> Title Case
  let caseStep = 0;
  container.querySelector('#smart-btn-change-case')?.addEventListener('mousedown', e => e.preventDefault());
  container.querySelector('#smart-btn-change-case')?.addEventListener('click', () => {
    const sel = window.getSelection();
    if (!sel || sel.isCollapsed) {
      showToast('Sorot terlebih dahulu kata atau kalimat yang ingin diubah huruf kapitalnya.', 'info');
      return;
    }
    const text = sel.toString();
    caseStep = (caseStep + 1) % 3;
    let replaced = text;
    if (caseStep === 1) replaced = text.toUpperCase();
    else if (caseStep === 2) replaced = text.toLowerCase();
    else replaced = text.replace(/\w\S*/g, w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
    execSmartCommand('insertText', replaced);
  });

  container.querySelector('#smart-btn-clear-fmt')?.addEventListener('mousedown', e => e.preventDefault());
  container.querySelector('#smart-btn-clear-fmt')?.addEventListener('click', () => execSmartCommand('removeFormat'));

  container.querySelector('#smart-highlight-color')?.addEventListener('input', (e) => {
    execSmartCommand('hiliteColor', e.target.value);
  });

  container.querySelector('#smart-text-color')?.addEventListener('input', (e) => {
    execSmartCommand('foreColor', e.target.value);
  });

  container.querySelector('#smart-btn-pilcrow')?.addEventListener('click', (e) => {
    showFormattingMarks = !showFormattingMarks;
    pagesStack.classList.toggle('smart-show-marks', showFormattingMarks);
    e.currentTarget.classList.toggle('active', showFormattingMarks);
  });

  container.querySelector('#smart-btn-shading')?.addEventListener('mousedown', e => e.preventDefault());
  container.querySelector('#smart-btn-shading')?.addEventListener('click', () => {
    execSmartCommand('hiliteColor', '#f1f5f9');
  });

  // Inden Paragraf & Spasi Antar Paragraf Popover
  const paraFmtBtn = container.querySelector('#smart-btn-para-format');
  const paraFmtPop = container.querySelector('#smart-para-format-popover');
  paraFmtBtn?.addEventListener('mousedown', e => e.preventDefault());
  paraFmtBtn?.addEventListener('click', (e) => {
    e.stopPropagation();
    paraFmtPop?.classList.toggle('hidden');
  });
  const getActiveSmartParagraphEl = () => {
    const sel = window.getSelection();
    const anchor = (lastSavedRange && pagesStack.contains(lastSavedRange.startContainer))
      ? lastSavedRange.startContainer
      : sel?.anchorNode;
    const el = anchor?.nodeType === 1 ? anchor : anchor?.parentElement;
    return el?.closest?.('.smart-editable-page p, .smart-editable-page div, .smart-editable-page li') || null;
  };
  paraFmtPop?.querySelectorAll('[data-first-indent]').forEach((item) => {
    item.addEventListener('mousedown', e => e.preventDefault());
    item.addEventListener('click', () => {
      const pEl = getActiveSmartParagraphEl();
      const cm = Number(item.dataset.firstIndent || 0);
      if (!pEl) {
        showToast('Klik dulu paragraf yang ingin diberi inden baris pertama.', 'info');
      } else {
        pushUndoSnapshot();
        pEl.style.textIndent = cm > 0 ? `${cm}cm` : '0cm';
        scheduleReflowPagination();
        schedulePersistSmartDoc();
      }
      paraFmtPop.classList.add('hidden');
    });
  });
  paraFmtPop?.querySelectorAll('[data-left-indent]').forEach((item) => {
    item.addEventListener('mousedown', e => e.preventDefault());
    item.addEventListener('click', () => {
      const pEl = getActiveSmartParagraphEl();
      const cm = Number(item.dataset.leftIndent || 0);
      if (!pEl) {
        showToast('Klik dulu paragraf yang ingin digeser inden kirinya.', 'info');
      } else {
        pushUndoSnapshot();
        pEl.style.paddingLeft = cm > 0 ? `${cm}cm` : '0cm';
        scheduleReflowPagination();
        schedulePersistSmartDoc();
      }
      paraFmtPop.classList.add('hidden');
    });
  });
  paraFmtPop?.querySelectorAll('[data-para-spacing]').forEach((item) => {
    item.addEventListener('click', () => {
      const pt = item.dataset.paraSpacing || '6';
      store.updateLetterOptions({ paragraphSpacing: pt });
      renderAllSmartPages(collectLivePagesData());
      paraFmtPop.classList.add('hidden');
      showToast(`Jarak antar paragraf diatur ke ${pt} pt.`, 'success');
    });
  });

  // Margin Popover
  const marginBtn = container.querySelector('#smart-btn-margin-menu');
  const marginPop = container.querySelector('#smart-margin-popover');
  marginBtn?.addEventListener('click', (e) => {
    e.stopPropagation();
    marginPop.classList.toggle('hidden');
  });
  marginPop?.querySelectorAll('[data-margin]').forEach((item) => {
    item.addEventListener('click', () => {
      const [top, right, bottom, left] = item.dataset.margin.split(',').map(Number);
      store.updateLetterOptions({ margins: { top, right, bottom, left } });
      syncGlobalPrintPageStyle(store.getLetterOptions());
      marginPop.classList.add('hidden');
      renderAllSmartPages(collectLivePagesData());
      showToast('Margin halaman berhasil diterapkan.', 'success');
    });
  });
  document.addEventListener('click', () => {
    marginPop?.classList.add('hidden');
    paraFmtPop?.classList.add('hidden');
  }, { once: false });

  // + Halaman (Page Break)
  container.querySelector('#smart-btn-add-page')?.addEventListener('mousedown', e => e.preventDefault());
  container.querySelector('#smart-btn-add-page')?.addEventListener('click', () => insertManualPageBreakAtSelection());

  // Toggle Tampil / Sembunyi Barcode QR Mengambang di Kanvas
  container.querySelector('#smart-btn-toggle-qr-canvas')?.addEventListener('click', () => {
    const latestSmart = store.getState().smartDocument || {};
    const stages = [...pagesStack.querySelectorAll('.floating-qr-stage')];
    const current = latestSmart.qrPlacement
      ? normalizeQrPlacement(latestSmart.qrPlacement, Math.max(1, stages.length))
      : findPreferredSignatureQrPlacement(getActiveSmartDocModel(), Math.max(1, stages.length));
    const next = { ...current, visibleOnCanvas: !(current.visibleOnCanvas !== false) };
    store.setActiveQrPlacement(next, 'smart');
    refreshSmartFloatingQrOverlay();
    showToast(
      next.visibleOnCanvas
        ? '▣ Barcode QR ditampilkan di kanvas — Anda bebas menggeser posisi, mengubah ukuran, atau memilih Depan/Belakang Teks.'
        : '👁️ Barcode QR disembunyikan di kanvas (tetap otomatis muncul saat Cetak/PDF & Surat Keluar).',
      'info'
    );
  });

  const onSmartVerificationSync = () => {
    if (!pagesStack.isConnected) {
      window.removeEventListener('docbuilder:verification-updated', onSmartVerificationSync);
      return;
    }
    refreshSmartFloatingQrOverlay();
  };
  window.addEventListener('docbuilder:verification-updated', onSmartVerificationSync);

  // Toggle Kop Surat
  container.querySelector('#smart-btn-toggle-kop')?.addEventListener('click', () => {
    const cur = store.getLetterOptions().showKop !== false;
    store.updateLetterOptions({ showKop: !cur });
    renderWordHeaderToggleButton(!cur);
    renderAllSmartPages(collectLivePagesData());
  });

  function renderWordHeaderToggleButton(isShown) {
    const btn = container.querySelector('#smart-btn-toggle-kop');
    if (!btn) return;
    btn.classList.toggle('active', isShown);
    btn.innerHTML = `${isShown ? ICON_EYE : ICON_EYE_OFF}<span>Tampil Kop</span>`;
  }

  // Insert Signature
  container.querySelector('#smart-btn-insert-sig')?.addEventListener('mousedown', e => e.preventDefault());
  container.querySelector('#smart-btn-insert-sig')?.addEventListener('click', () => {
    const sigHtml = renderSignatureTableHtml({}, store.getSchoolProfile()) + '<p><br></p>';
    execSmartCommand('insertHTML', sigHtml);
    showToast('Blok tanda tangan resmi berhasil disisipkan.', 'success');
  });

  // Undo & Redo
  container.querySelector('#smart-btn-undo')?.addEventListener('click', () => {
    if (smartUndoStack.length > 0) {
      smartRedoStack.push(JSON.stringify(collectLivePagesData()));
      const prev = JSON.parse(smartUndoStack.pop());
      renderAllSmartPages(prev);
      schedulePersistSmartDoc();
    } else {
      document.execCommand('undo');
    }
  });

  container.querySelector('#smart-btn-redo')?.addEventListener('click', () => {
    if (smartRedoStack.length > 0) {
      smartUndoStack.push(JSON.stringify(collectLivePagesData()));
      const next = JSON.parse(smartRedoStack.pop());
      renderAllSmartPages(next);
      schedulePersistSmartDoc();
    } else {
      document.execCommand('redo');
    }
  });

  // Salin isi surat dari Lembar Kerja pertama
  container.querySelector('#smart-btn-pull-canvas1')?.addEventListener('click', () => {
    pushUndoSnapshot();
    const curState = store.getState();
    const converted = docModelToSmartPages(curState.document, curState.schoolProfile);
    smartDoc = { pages: converted.pages, meta: converted.meta };
    store.setState({ smartDocument: smartDoc }, false);
    store.setActiveVaultId(null, 'smart');
    renderAllSmartPages(smartDoc.pages);
    updateSmartVaultBadge?.();
    showToast('Surat dari Lembar Kerja berhasil disalin ke Mode Word Pro.', 'success');
  });

  // Impor .docx ke Kanvas Word Pro
  const docxInput = container.querySelector('#smart-docx-upload-input');
  container.querySelector('#smart-btn-import-docx')?.addEventListener('click', () => docxInput.click());
  docxInput?.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      showToast(`Mengonversi berkas "${file.name}" ke Word Pro...`, 'info');
      const buffer = await file.arrayBuffer();
      const parsedDocx = await parseDocxForEditing(buffer);
      const model = convertDocxToDocumentModel(parsedDocx, store.getSchoolProfile(), store.getLetterOptions());
      pushUndoSnapshot();
      const converted = docModelToSmartPages(model, store.getSchoolProfile());
      smartDoc = { pages: converted.pages, meta: { ...converted.meta, title: file.name.replace(/\.docx$/i, '') } };
      store.setState({ smartDocument: smartDoc }, false);
      store.setActiveVaultId(null, 'smart');
      renderAllSmartPages(smartDoc.pages);
      updateSmartVaultBadge?.();
      showToast(`Berkas "${file.name}" siap diedit di Kanvas Word Pro!`, 'success');
    } catch (err) {
      showToast('Gagal mengimpor .docx: ' + err.message, 'error');
    } finally {
      docxInput.value = '';
    }
  });

  // Toggle Copilot Drawer
  const copilotDrawer = container.querySelector('#smart-copilot-drawer');
  const toggleCopilotBtn = container.querySelector('#smart-btn-toggle-copilot');
  function setCopilotVisible(show) {
    copilotVisible = show;
    copilotDrawer.classList.toggle('hidden', !show);
    toggleCopilotBtn.classList.toggle('active', show);
  }
  toggleCopilotBtn?.addEventListener('click', () => setCopilotVisible(!copilotVisible));
  container.querySelector('#smart-copilot-close')?.addEventListener('click', () => setCopilotVisible(false));
  container.querySelector('#smart-btn-open-vault-nav')?.addEventListener('click', () => {
    document.querySelector('.nav-tab[data-tab="vault"]')?.click();
  });

  // Zoom controls
  container.querySelector('#smart-zoom-out')?.addEventListener('click', () => {
    currentZoom = Math.max(0.5, Number((currentZoom - 0.1).toFixed(2)));
    applyZoom({ fromUser: true });
  });
  container.querySelector('#smart-zoom-in')?.addEventListener('click', () => {
    currentZoom = Math.min(1.5, Number((currentZoom + 0.1).toFixed(2)));
    applyZoom({ fromUser: true });
  });
  container.querySelector('#smart-zoom-slider')?.addEventListener('input', (e) => {
    currentZoom = Number(e.target.value) / 100;
    applyZoom({ fromUser: true });
  });
  container.querySelector('#smart-zoom-fit')?.addEventListener('click', () => {
    const desk = container.querySelector('#smart-word-desk');
    const sheet = pagesStack.querySelector('.smart-paper-sheet');
    if (desk && sheet) {
      const ratio = (desk.clientWidth - 48) / sheet.offsetWidth;
      currentZoom = Math.max(0.5, Math.min(1.2, Number(ratio.toFixed(2))));
      applyZoom({ fromUser: true });
    }
  });

  let smartResizeRaf = null;
  const onSmartViewportResize = () => {
    if (!pagesStack.isConnected) {
      window.removeEventListener('resize', onSmartViewportResize);
      return;
    }
    if (store.getState().zoomUserAdjusted) return;
    if (smartResizeRaf) cancelAnimationFrame(smartResizeRaf);
    smartResizeRaf = requestAnimationFrame(() => {
      const nextZoom = Number((computeDefaultDesktopZoom() / 100).toFixed(2));
      if (Math.abs(currentZoom - nextZoom) >= 0.01) {
        currentZoom = nextZoom;
        applyZoom({ fromUser: false });
      }
    });
  };
  window.addEventListener('resize', onSmartViewportResize);

  // =========================================================================
  // LOGIKA AI EDIT & AI COPILOT WORD PRO
  // =========================================================================

  container.querySelectorAll('.quick-ai-chip').forEach((chip) => {
    chip.addEventListener('click', () => {
      const inp = container.querySelector('#smart-quick-ai-input');
      inp.value = chip.dataset.prompt;
      inp.focus();
    });
  });

  const compareBarEl = container.querySelector('#smart-ai-compare-bar');
  const compareSummaryEl = container.querySelector('#smart-cmp-summary');

  function showAiCompareBar(stateObj) {
    lastAiCompareState = stateObj;
    if (!compareBarEl || !stateObj) return;
    if (compareSummaryEl) {
      compareSummaryEl.textContent = stateObj.summaryText || 'Revisi AI berhasil diterapkan (tata letak & tabel tetap aman).';
    }
    compareBarEl.classList.remove('hidden');
  }

  function hideAiCompareBar() {
    lastAiCompareState = null;
    compareBarEl?.classList.add('hidden');
  }

  function rollbackLastAiRevision() {
    if (!lastAiCompareState || !Array.isArray(lastAiCompareState.prevPages)) return;
    pushUndoSnapshot();
    const latestState = store.getState().smartDocument || {};
    smartDoc = {
      ...smartDoc,
      pages: JSON.parse(JSON.stringify(lastAiCompareState.prevPages)),
      meta: lastAiCompareState.prevMeta || smartDoc.meta,
      ...(latestState.verification ? { verification: latestState.verification } : {}),
      ...(latestState.qrPlacement ? { qrPlacement: latestState.qrPlacement } : {})
    };
    store.setState({ smartDocument: smartDoc }, false);
    renderAllSmartPages(smartDoc.pages);
    hideAiCompareBar();
    showToast('↶ Revisi AI dibatalkan. Dokumen dikembalikan ke kondisi persis sebelum diedit AI.', 'info');
  }

  function openAiCompareModal() {
    if (!lastAiCompareState) return;
    const changes = Array.isArray(lastAiCompareState.changes) ? lastAiCompareState.changes : [];
    const rowsHtml = changes.length === 0
      ? `<tr><td colspan="3" style="text-align:center; padding:14px; color:#64748b;">Perubahan diterapkan pada dokumen (${escapeHtmlText(lastAiCompareState.summaryText || 'Revisi AI')}).</td></tr>`
      : changes.map((c, idx) => `
        <tr>
          <td style="font-weight:700; color:#334155; font-size:12px;">${idx + 1}. ${escapeHtmlText(c.targetLabel || c.kind)}</td>
          <td style="background:#fef2f2; color:#991b1b; font-size:12px; white-space:pre-wrap;">${escapeHtmlText(c.before || '-')}</td>
          <td style="background:#ecfdf5; color:#14532d; font-size:12px; white-space:pre-wrap;">${escapeHtmlText(c.after || '-')}</td>
        </tr>
      `).join('');

    modalRoot.innerHTML = `
      <div class="modal-overlay">
        <div class="modal-card" style="max-width: 820px; width: 95%; max-height: 88vh; display: flex; flex-direction: column;">
          <div class="modal-header">
            <h3>👁️ Perbandingan Revisi AI (Sebelum vs Sesudah)</h3>
            <button type="button" class="modal-close" id="smart-modal-close">&times;</button>
          </div>
          <div class="modal-body" style="overflow-y: auto;">
            <p style="font-size:12.5px; color:#475569; margin:0 0 12px 0;">
              <strong>Ringkasan AI:</strong> ${escapeHtmlText(lastAiCompareState.summaryText || '-')}
            </p>
            <table class="smart-compare-table">
              <thead>
                <tr>
                  <th style="width: 24%;">Bagian Target</th>
                  <th style="width: 38%;">Sebelum Revisi</th>
                  <th style="width: 38%;">Sesudah Revisi AI</th>
                </tr>
              </thead>
              <tbody>${rowsHtml}</tbody>
            </table>
          </div>
          <div class="modal-footer" style="justify-content: space-between;">
            <button type="button" class="btn btn-danger" id="smart-cmp-modal-rollback">↶ Urungkan Semua Revisi AI Ini</button>
            <button type="button" class="btn btn-primary" id="smart-cmp-modal-keep">✓ Pertahankan Hasil</button>
          </div>
        </div>
      </div>
    `;
    modalRoot.querySelector('#smart-modal-close').onclick = closeSmartModal;
    modalRoot.querySelector('#smart-cmp-modal-keep').onclick = closeSmartModal;
    modalRoot.querySelector('#smart-cmp-modal-rollback').onclick = () => {
      closeSmartModal();
      rollbackLastAiRevision();
    };
  }

  container.querySelector('#smart-btn-open-compare')?.addEventListener('click', openAiCompareModal);
  container.querySelector('#smart-btn-rollback-ai')?.addEventListener('click', rollbackLastAiRevision);
  container.querySelector('#smart-btn-dismiss-compare')?.addEventListener('click', hideAiCompareBar);

  async function runSmartAiEdit(instruction) {
    const qStatus = container.querySelector('#smart-quick-ai-status');
    const cStatus = container.querySelector('#smart-copilot-status');
    const msg = 'AI sedang membaca peta dokumen & menambal bagian target tanpa mengubah tata letak...';
    [qStatus, cStatus].forEach((el) => {
      if (el) {
        el.textContent = msg;
        el.classList.remove('hidden', 'error');
      }
    });

    try {
      const livePages = collectLivePagesData();
      const prevPagesBackup = JSON.parse(JSON.stringify(livePages));
      const prevMetaBackup = smartDoc.meta ? JSON.parse(JSON.stringify(smartDoc.meta)) : null;
      const prevOptsBackup = normalizePageSetupOptions(store.getLetterOptions());
      const stateForSmart = store.getState();
      const baseSmartSchool = smartDoc?.schoolProfile || store.getEffectiveSchoolForDocument(smartDoc);
      const explicitSwitchUnit = detectTargetUnitFromText(instruction, null);
      const schoolProfile = explicitSwitchUnit
        ? resolveSchoolProfileForUnit(explicitSwitchUnit, stateForSmart.schoolProfile, stateForSmart.kopGallery || [])
        : baseSmartSchool;
      const allowLayoutRestructure = shouldAllowLayoutRestructure(instruction);
      const hasExistingContent = livePages.some((p) => {
        const txt = String(p?.html || '').replace(/<[^>]+>/g, '').trim();
        return txt.length > 10 || /<table/i.test(p?.html || '');
      });
      const indexObj = buildSurgicalIndexFromSmartPages(livePages, schoolProfile, { letterOptions: prevOptsBackup });
      const preferredApiId = container.querySelector('#smart-copilot-model-select')?.value || selectedApiId;
      const editedJson = await editDocumentWithAi(indexObj.snapshot, instruction, {
        schoolProfile,
        ...(explicitSwitchUnit ? { targetUnitId: explicitSwitchUnit } : {}),
        preferredApiId,
        allowLayoutRestructure
      });

      pushUndoSnapshot();
      const latestState = store.getState().smartDocument || {};
      const shouldFullRebuild =
        !hasExistingContent ||
        (allowLayoutRestructure &&
          editedJson?.editScope === 'restructure' &&
          Array.isArray(editedJson.blocks) &&
          editedJson.blocks.length > 0 &&
          (!Array.isArray(editedJson.patches) || editedJson.patches.length === 0));

      let summaryText = editedJson.summaryOfChanges || 'Revisi AI berhasil diterapkan ke Kanvas Word Pro.';
      let recordedChanges = [];

      if (shouldFullRebuild) {
        const docModel = aiJsonToSmartDocModel(editedJson, schoolProfile, prevOptsBackup);
        docModel.targetUnitId = schoolProfile.unitId;
        if (docModel.extractedPageSetup || docModel.pageSetup) {
          store.updateLetterOptions(docModel.extractedPageSetup || docModel.pageSetup);
          syncGlobalPrintPageStyle(store.getLetterOptions());
          syncSmartRibbonFromOptions(store.getLetterOptions());
        }
        const converted = docModelToSmartPages(docModel, schoolProfile);
        smartDoc = {
          pages: converted.pages,
          meta: converted.meta,
          targetUnitId: schoolProfile.unitId,
          schoolProfile,
          ...(latestState.verification ? { verification: latestState.verification } : {}),
          ...(latestState.qrPlacement ? { qrPlacement: latestState.qrPlacement } : {})
        };
        recordedChanges = [{ kind: 'rebuild', targetLabel: 'Seluruh Dokumen', before: 'Struktur sebelumnya', after: summaryText }];
      } else {
        const patchResult = applySurgicalPatchesToSmartPages(livePages, editedJson, {
          schoolProfile,
          allowLayoutRestructure,
          letterOptions: prevOptsBackup,
          renderBlocksToPageHtml: (aiBlocks = []) => {
            const miniDoc = { blocks: convertAiBlocksToDocumentBlocks(aiBlocks) };
            return docModelToSmartPages(miniDoc, schoolProfile).pages.map((mp) => mp.html).filter(Boolean).join('\n');
          }
        });
        if (patchResult.pageSetupUpdates) {
          store.updateLetterOptions(patchResult.pageSetupUpdates);
          syncGlobalPrintPageStyle(store.getLetterOptions());
          syncSmartRibbonFromOptions(store.getLetterOptions());
        }
        smartDoc = {
          ...smartDoc,
          pages: patchResult.pages,
          meta: patchResult.meta || smartDoc.meta,
          targetUnitId: schoolProfile.unitId,
          schoolProfile,
          ...(latestState.verification ? { verification: latestState.verification } : {}),
          ...(latestState.qrPlacement ? { qrPlacement: latestState.qrPlacement } : {})
        };
        recordedChanges = patchResult.changes || [];
        if (patchResult.appliedCount > 0) {
          summaryText = `${summaryText} (${patchResult.appliedCount} bagian diperbarui, tata letak & tabel tetap utuh)`;
        }
      }

      store.setState({ smartDocument: smartDoc }, false);
      hasUnsavedVaultChanges = Boolean(store.getActiveVaultEntry('smart'));
      lastLocalSavedTime = new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
      updateSmartSaveIndicator();
      renderAllSmartPages(smartDoc.pages);
      showAiCompareBar({
        prevPages: prevPagesBackup,
        prevMeta: prevMetaBackup,
        summaryText,
        changes: recordedChanges
      });

      [qStatus, cStatus].forEach((el) => {
        if (el) el.textContent = summaryText;
      });
      showToast(summaryText, 'success');
    } catch (err) {
      [qStatus, cStatus].forEach((el) => {
        if (el) {
          el.textContent = 'Gagal memproses AI Edit: ' + err.message;
          el.classList.add('error');
        }
      });
      showToast('Gagal memproses AI Edit: ' + err.message, 'error');
    }
  }

  container.querySelector('#smart-btn-quick-ai-apply')?.addEventListener('click', () => {
    const val = container.querySelector('#smart-quick-ai-input').value.trim();
    if (!val) {
      showToast('Ketikkan instruksi perubahan atau klik salah satu contoh instruksi.', 'warning');
      return;
    }
    runSmartAiEdit(val);
  });

  container.querySelector('#smart-quick-ai-input')?.addEventListener('keydown', (e) => {
    if (e.isComposing) return;
    if (e.key === 'Enter') {
      e.preventDefault();
      e.stopPropagation();
      container.querySelector('#smart-btn-quick-ai-apply')?.click();
    }
  });

  const smartCopilotDrawerEl = container.querySelector('#smart-copilot-drawer');
  if (smartCopilotDrawerEl) {
    mountCompactChatDrawer(smartCopilotDrawerEl, {
      editorKind: 'smart-canvas',
      title: 'AI Copilot Word Pro',
      getSnapshot: () => {
        const livePages = collectLivePagesData();
        const hasExisting = livePages.some((p) => {
          const txt = String(p?.html || '').replace(/<[^>]+>/g, '').trim();
          return txt.length > 10 || /<table/i.test(p?.html || '');
        });
        if (!hasExisting) return null;
        return buildSurgicalIndexFromSmartPages(livePages, store.getEffectiveSchoolForDocument(smartDoc), {
          letterOptions: normalizePageSetupOptions(store.getLetterOptions())
        }).snapshot;
      },
      onApplyCreateDraft: (draft, unitCtx = {}) => {
        pushUndoSnapshot();
        const stateNow = store.getState();
        const prevSmart = JSON.parse(JSON.stringify(smartDoc));
        const prevOpts = normalizePageSetupOptions(store.getLetterOptions());
        const unitId = draft?._targetUnitId
          || unitCtx.targetUnitId
          || draft?.targetUnit
          || detectTargetUnitFromText(`${draft?.letterTitle || ''} ${draft?.nomorSurat || ''} ${draft?.signer?.name || ''}`, null)
          || store.getActiveUnitId();
        const targetSchool = draft?._targetSchoolProfile
          || unitCtx.schoolProfile
          || resolveSchoolProfileForUnit(unitId, stateNow.schoolProfile, stateNow.kopGallery || []);
        const docModel = aiJsonToSmartDocModel(draft, targetSchool, prevOpts);
        docModel.targetUnitId = targetSchool.unitId;
        if (docModel.extractedPageSetup || docModel.pageSetup) {
          store.updateLetterOptions(docModel.extractedPageSetup || docModel.pageSetup);
          syncGlobalPrintPageStyle(store.getLetterOptions());
          syncSmartRibbonFromOptions(store.getLetterOptions());
        }
        const converted = docModelToSmartPages(docModel, targetSchool);
        smartDoc = {
          pages: converted.pages,
          meta: converted.meta,
          targetUnitId: targetSchool.unitId,
          schoolProfile: targetSchool
        };
        store.setState({ smartDocument: smartDoc }, false);
        store.setActiveVaultId(null, 'smart');
        renderAllSmartPages(smartDoc.pages);
        updateSmartVaultBadge?.();
        return {
          changes: [{ targetLabel: 'Surat Baru', before: '-', after: draft?.letterTitle || 'Draf diterapkan di Word Pro' }],
          undoBackup: { workspace: 'smart-canvas', prevSmartDocument: prevSmart, prevOpts }
        };
      },
      onApplySurgicalEdit: (editPlan, userInstruction, unitCtx = {}) => {
        pushUndoSnapshot();
        const stateNow = store.getState();
        const livePages = collectLivePagesData();
        const prevSmart = JSON.parse(JSON.stringify(smartDoc));
        const prevOpts = normalizePageSetupOptions(store.getLetterOptions());
        const explicitUnit = detectTargetUnitFromText(userInstruction || '', null);
        const resolvedUnit = explicitUnit || unitCtx.targetUnitId || smartDoc?.targetUnitId || '';
        const schoolProfile = resolvedUnit
          ? resolveSchoolProfileForUnit(resolvedUnit, stateNow.schoolProfile, stateNow.kopGallery || [])
          : store.getEffectiveSchoolForDocument(smartDoc);
        const latestState = stateNow.smartDocument || {};
        const allowLayoutRestructure = shouldAllowLayoutRestructure(userInstruction || '');
        let recordedChanges = [];

        if (allowLayoutRestructure && editPlan?.editScope === 'restructure' && Array.isArray(editPlan.blocks) && editPlan.blocks.length > 0 && (!Array.isArray(editPlan.patches) || editPlan.patches.length === 0)) {
          const docModel = aiJsonToSmartDocModel(editPlan, schoolProfile, prevOpts);
          docModel.targetUnitId = schoolProfile.unitId;
          if (docModel.extractedPageSetup || docModel.pageSetup) {
            store.updateLetterOptions(docModel.extractedPageSetup || docModel.pageSetup);
            syncGlobalPrintPageStyle(store.getLetterOptions());
            syncSmartRibbonFromOptions(store.getLetterOptions());
          }
          const converted = docModelToSmartPages(docModel, schoolProfile);
          smartDoc = {
            pages: converted.pages,
            meta: converted.meta,
            targetUnitId: schoolProfile.unitId,
            schoolProfile,
            ...(latestState.verification ? { verification: latestState.verification } : {}),
            ...(latestState.qrPlacement ? { qrPlacement: latestState.qrPlacement } : {})
          };
          recordedChanges = [{ targetLabel: 'Dokumen Lengkap', before: 'Struktur lama', after: editPlan.summaryOfChanges || 'Diperbarui' }];
        } else {
          const patchResult = applySurgicalPatchesToSmartPages(livePages, editPlan, {
            schoolProfile,
            allowLayoutRestructure,
            letterOptions: prevOpts,
            renderBlocksToPageHtml: (aiBlocks = []) => {
              const miniDoc = { blocks: convertAiBlocksToDocumentBlocks(aiBlocks) };
              return docModelToSmartPages(miniDoc, schoolProfile).pages.map((mp) => mp.html).filter(Boolean).join('\n');
            }
          });
          if (patchResult.pageSetupUpdates) {
            store.updateLetterOptions(patchResult.pageSetupUpdates);
            syncGlobalPrintPageStyle(store.getLetterOptions());
            syncSmartRibbonFromOptions(store.getLetterOptions());
          }
          smartDoc = {
            ...smartDoc,
            pages: patchResult.pages,
            meta: patchResult.meta || smartDoc.meta,
            targetUnitId: schoolProfile.unitId,
            schoolProfile,
            ...(latestState.verification ? { verification: latestState.verification } : {}),
            ...(latestState.qrPlacement ? { qrPlacement: latestState.qrPlacement } : {})
          };
          recordedChanges = patchResult.changes || [];
        }

        store.setState({ smartDocument: smartDoc }, false);
        hasUnsavedVaultChanges = Boolean(store.getActiveVaultEntry('smart'));
        lastLocalSavedTime = new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
        updateSmartSaveIndicator();
        renderAllSmartPages(smartDoc.pages);
        return {
          changes: recordedChanges,
          undoBackup: { workspace: 'smart-canvas', prevSmartDocument: prevSmart, prevOpts }
        };
      },
      onUndoAction: (backup) => {
        if (backup?.prevOpts) {
          store.updateLetterOptions(backup.prevOpts);
          syncGlobalPrintPageStyle(store.getLetterOptions());
          syncSmartRibbonFromOptions(store.getLetterOptions());
        }
        if (backup?.prevSmartDocument) {
          smartDoc = backup.prevSmartDocument;
          store.setState({ smartDocument: smartDoc }, false);
          renderAllSmartPages(smartDoc.pages);
        }
      },
      onQuickAction: async (act) => {
        if (act === 'polish') {
          const rawText = collectLivePagesData()
            .map(p => (p.html || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim())
            .join('\n\n')
            .trim();
          if (!rawText) return;
          showToast('AI sedang merapihkan naskah menjadi format surat dinas resmi...', 'info');
          try {
            const polishedModel = await polishMessyDocumentWithAi(rawText, {
              ...store.getLetterOptions(),
              schoolProfile: store.getSchoolProfile()
            });
            pushUndoSnapshot();
            if (polishedModel.pageSetup) {
              store.updateLetterOptions(polishedModel.pageSetup);
              syncGlobalPrintPageStyle(store.getLetterOptions());
              syncSmartRibbonFromOptions(store.getLetterOptions());
            }
            const converted = docModelToSmartPages(polishedModel, store.getSchoolProfile());
            smartDoc = { pages: converted.pages, meta: converted.meta };
            store.setState({ smartDocument: smartDoc }, false);
            renderAllSmartPages(smartDoc.pages);
            showToast('Naskah berhasil dirapihkan menjadi format surat resmi.', 'success');
          } catch (err) {
            showToast('Gagal merapihkan: ' + err.message, 'error');
          }
        } else if (act === 'puebi') {
          runSmartAiEdit('Perbaiki seluruh kata, imbuhan, tanda baca, dan kalimat surat ini ke standar resmi kedinasan PUEBI tanpa mengubah substansi maupun rincian tabel.');
        } else if (act === 'today') {
          const todayStr = new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
          runSmartAiEdit(`Perbarui tanggal tanda tangan surat ini ke tanggal hari ini: ${todayStr}. Pertahankan isi lainnya.`);
        }
      },
      onClose: () => setCopilotVisible(false),
      onOpenVault: () => document.querySelector('.nav-tab[data-tab="vault"]')?.click()
    });
  }

  // =========================================================================
  // MODALS: FORMAT KERTAS, GALERI KOP, SISIPKAN TABEL, SIMPAN GUDANG, SURAT KELUAR
  // =========================================================================

  function closeSmartModal() {
    modalRoot.innerHTML = '';
  }

  container.querySelector('#smart-btn-format-modal')?.addEventListener('click', () => {
    const o = normalizePageSetupOptions(store.getLetterOptions());
    const m = o.margins || { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 };
    const pn = o.pageNumbering || PAGE_NUMBERING_DEFAULT;
    modalRoot.innerHTML = `
      <div class="modal-overlay">
        <div class="modal-box" style="max-width: 560px;">
          <div class="modal-header">
            <h3>Pengaturan Halaman, Tipografi &amp; Nomor Halaman</h3>
            <button class="btn-icon" id="smart-modal-close">✕</button>
          </div>
          <div class="modal-body" style="display:grid; gap:12px;">
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
              <div>
                <label class="form-label">Ukuran Kertas</label>
                <select id="sm-paper" class="form-input">
                  <option value="a4" ${o.paperSize === 'a4' ? 'selected' : ''}>A4</option>
                  <option value="f4" ${o.paperSize === 'f4' ? 'selected' : ''}>F4 / Folio</option>
                  <option value="letter" ${o.paperSize === 'letter' ? 'selected' : ''}>Letter</option>
                  <option value="legal" ${o.paperSize === 'legal' ? 'selected' : ''}>Legal</option>
                </select>
              </div>
              <div>
                <label class="form-label">Orientasi</label>
                <select id="sm-orient" class="form-input">
                  <option value="portrait" ${o.orientation !== 'landscape' ? 'selected' : ''}>Potret</option>
                  <option value="landscape" ${o.orientation === 'landscape' ? 'selected' : ''}>Lanskap</option>
                </select>
              </div>
            </div>
            <div style="display:grid; grid-template-columns:repeat(4,1fr); gap:8px;">
              <div><label class="form-label">Atas (cm)</label><input type="number" step="0.1" id="sm-top" class="form-input" value="${m.top}"></div>
              <div><label class="form-label">Bawah (cm)</label><input type="number" step="0.1" id="sm-bottom" class="form-input" value="${m.bottom}"></div>
              <div><label class="form-label">Kiri (cm)</label><input type="number" step="0.1" id="sm-left" class="form-input" value="${m.left}"></div>
              <div><label class="form-label">Kanan (cm)</label><input type="number" step="0.1" id="sm-right" class="form-input" value="${m.right}"></div>
            </div>
            <div style="display:grid; grid-template-columns:2fr 1fr 1fr 1fr; gap:8px;">
              <div>
                <label class="form-label">Font Utama</label>
                <select id="sm-font-family" class="form-input">
                  ${SUPPORTED_FONT_FAMILIES.map(f =>
                    `<option value="${f}" ${o.fontFamily === f ? 'selected' : ''}>${f}</option>`
                  ).join('')}
                </select>
              </div>
              <div>
                <label class="form-label">Ukuran (pt)</label>
                <select id="sm-font-size" class="form-input">
                  ${['9', '10', '11', '12', '13', '14', '16', '18'].map(sz =>
                    `<option value="${sz}" ${String(o.fontSize) === sz ? 'selected' : ''}>${sz} pt</option>`
                  ).join('')}
                </select>
              </div>
              <div>
                <label class="form-label">Spasi Baris</label>
                <select id="sm-line-spacing" class="form-input">
                  <option value="1.0" ${String(o.lineSpacing) === '1.0' ? 'selected' : ''}>1,0</option>
                  <option value="1.15" ${String(o.lineSpacing) === '1.15' ? 'selected' : ''}>1,15</option>
                  <option value="1.5" ${String(o.lineSpacing) === '1.5' ? 'selected' : ''}>1,5</option>
                  <option value="2.0" ${String(o.lineSpacing) === '2.0' ? 'selected' : ''}>2,0</option>
                </select>
              </div>
              <div>
                <label class="form-label">Jarak Paragraf</label>
                <select id="sm-para-spacing" class="form-input">
                  ${['0', '3', '6', '8', '12'].map(pt =>
                    `<option value="${pt}" ${String(o.paragraphSpacing) === pt ? 'selected' : ''}>${pt} pt</option>`
                  ).join('')}
                </select>
              </div>
            </div>
            <div style="padding:10px;background:#f8fafc;border:1px solid #cbd5e1;border-radius:8px;display:grid;gap:8px;">
              <label style="display:flex;align-items:center;gap:8px;font-weight:600;font-size:12.5px;cursor:pointer;">
                <input type="checkbox" id="sm-pn-enabled" ${pn.enabled ? 'checked' : ''} style="width:16px;height:16px;">
                <span>Tampilkan Nomor Halaman Otomatis (Kanvas, Cetak &amp; .docx)</span>
              </label>
              <div style="display:grid;grid-template-columns:1.2fr 1.2fr 0.8fr;gap:8px;">
                <div>
                  <label class="form-label">Posisi</label>
                  <select id="sm-pn-pos" class="form-input">
                    <option value="bottom-center" ${pn.position === 'bottom-center' ? 'selected' : ''}>Bawah Tengah</option>
                    <option value="bottom-right" ${pn.position === 'bottom-right' ? 'selected' : ''}>Bawah Kanan</option>
                    <option value="top-right" ${pn.position === 'top-right' ? 'selected' : ''}>Atas Kanan</option>
                  </select>
                </div>
                <div>
                  <label class="form-label">Format Nomor</label>
                  <select id="sm-pn-fmt" class="form-input">
                    <option value="decimal" ${pn.format === 'decimal' ? 'selected' : ''}>1, 2, 3 (Angka)</option>
                    <option value="lowerRoman" ${pn.format === 'lowerRoman' ? 'selected' : ''}>i, ii, iii (Romawi Kecil)</option>
                    <option value="upperRoman" ${pn.format === 'upperRoman' ? 'selected' : ''}>I, II, III (Romawi Besar)</option>
                  </select>
                </div>
                <div>
                  <label class="form-label">Mulai Dari</label>
                  <input type="number" min="1" max="999" id="sm-pn-start" class="form-input" value="${pn.startAt || 1}">
                </div>
              </div>
            </div>
          </div>
          <div class="modal-footer">
            <button class="btn btn-secondary" id="smart-modal-cancel">Batal</button>
            <button class="btn btn-primary" id="smart-modal-apply">Terapkan Format</button>
          </div>
        </div>
      </div>`;
    modalRoot.querySelector('#smart-modal-close').onclick = closeSmartModal;
    modalRoot.querySelector('#smart-modal-cancel').onclick = closeSmartModal;
    modalRoot.querySelector('#smart-modal-apply').onclick = () => {
      store.updateLetterOptions({
        paperSize: modalRoot.querySelector('#sm-paper').value,
        orientation: modalRoot.querySelector('#sm-orient').value,
        margins: {
          top: Number(modalRoot.querySelector('#sm-top').value) || 2.0,
          bottom: Number(modalRoot.querySelector('#sm-bottom').value) || 2.0,
          left: Number(modalRoot.querySelector('#sm-left').value) || 2.5,
          right: Number(modalRoot.querySelector('#sm-right').value) || 2.0
        },
        fontFamily: modalRoot.querySelector('#sm-font-family').value,
        fontSize: modalRoot.querySelector('#sm-font-size').value,
        lineSpacing: modalRoot.querySelector('#sm-line-spacing').value,
        paragraphSpacing: modalRoot.querySelector('#sm-para-spacing').value,
        pageNumbering: {
          enabled: Boolean(modalRoot.querySelector('#sm-pn-enabled').checked),
          position: modalRoot.querySelector('#sm-pn-pos').value,
          format: modalRoot.querySelector('#sm-pn-fmt').value,
          startAt: Math.max(1, parseInt(modalRoot.querySelector('#sm-pn-start').value || '1', 10) || 1)
        }
      });
      syncGlobalPrintPageStyle(store.getLetterOptions());
      syncSmartRibbonFromOptions(store.getLetterOptions());
      closeSmartModal();
      renderAllSmartPages(collectLivePagesData());
      showToast('Pengaturan halaman, tipografi & nomor halaman diterapkan.', 'success');
    };
  });

  function openKopGalleryModal() {
    const gallery = store.getKopGallery();
    const curKop = store.getSchoolProfile().kopImage || DEFAULT_KOP_DATA_URL;
    modalRoot.innerHTML = `
      <div class="modal-overlay">
        <div class="modal-box" style="max-width: 680px;">
          <div class="modal-header">
            <h3>Galeri Kop Surat Sekolah</h3>
            <button class="btn-icon" id="smart-modal-close">✕</button>
          </div>
          <div class="modal-body" style="max-height:60vh; overflow-y:auto; display:grid; gap:10px;">
            <div style="display:flex; justify-content:space-between; align-items:center;">
              <span style="font-size:12px; color:#475569;">Pilih kop surat resmi atau unggah gambar kop PNG/JPG kustom:</span>
              <button class="btn btn-sm btn-secondary" id="sm-kop-upload-btn">${ICON_UPLOAD} Unggah Kop</button>
            </div>
            ${gallery.map(item => `
              <div style="border:2px solid ${item.dataUrl === curKop ? '#15803d' : '#e2e8f0'}; border-radius:10px; padding:10px; background:#fff;">
                <img src="${item.dataUrl}" style="width:100%; max-height:110px; object-fit:contain; border-bottom:1px solid #f1f5f9; margin-bottom:6px;">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                  <strong style="font-size:13px;">${escapeHtmlText(item.name)}</strong>
                  <button class="btn btn-sm ${item.dataUrl === curKop ? 'btn-primary' : 'btn-secondary'}" data-pick-kop="${item.id}">
                    ${item.dataUrl === curKop ? 'Sedang Aktif' : 'Gunakan Kop Ini'}
                  </button>
                </div>
              </div>`).join('')}
          </div>
        </div>
      </div>`;
    modalRoot.querySelector('#smart-modal-close').onclick = closeSmartModal;
    modalRoot.querySelector('#sm-kop-upload-btn').onclick = () => {
      container.querySelector('#smart-kop-upload-input').click();
    };
    modalRoot.querySelectorAll('[data-pick-kop]').forEach((btn) => {
      btn.onclick = () => {
        store.applyKopFromGallery(btn.dataset.pickKop);
        syncGlobalPrintPageStyle(store.getLetterOptions());
      closeSmartModal();
        renderAllSmartPages(collectLivePagesData());
        showToast('Kop surat resmi berhasil diperbarui.', 'success');
      };
    });
  }

  container.querySelector('#smart-btn-kop-gallery')?.addEventListener('click', openKopGalleryModal);
  container.querySelector('#smart-kop-upload-input')?.addEventListener('change', (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const added = store.addKopGalleryEntry({
        name: file.name.replace(/\.[^.]+$/, ''),
        dataUrl: reader.result,
        category: 'Unggahan'
      });
      store.applyKopFromGallery(added.id);
      syncGlobalPrintPageStyle(store.getLetterOptions());
      closeSmartModal();
      renderAllSmartPages(collectLivePagesData());
      showToast('Kop surat berhasil diunggah dan diterapkan.', 'success');
    };
    reader.readAsDataURL(file);
    e.target.value = '';
  });

  // Shortcut 1-klik Tata Letak 2 Kolom Berdampingan (Tanpa Garis)
  container.querySelector('#smart-btn-insert-2col')?.addEventListener('click', () => {
    const html = `
<table class="smart-word-table smart-table-none" data-border-style="none" data-header-shading="false" data-header-bold="false" data-col-widths="50,50">
  <colgroup><col style="width: 50%;"><col style="width: 50%;"></colgroup>
  <tbody>
    <tr>
      <td><p><strong>Keterangan / Catatan Kiri:</strong></p><p>Isi teks kolom sebelah kiri...</p></td>
      <td style="text-align: center;"><p>Garut, ${escapeHtmlText(new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }))}</p><p>${escapeHtmlText(store.getSchoolProfile().principalTitle || 'Kepala SMK IT Al-Hawari')}</p><p><br></p><p><br></p><p><strong><u>${escapeHtmlText(store.getSchoolProfile().principalName || 'Dian Nurhadian, S.Pd.')}</u></strong></p></td>
    </tr>
  </tbody>
</table><p><br></p>`.trim();
    syncGlobalPrintPageStyle(store.getLetterOptions());
    execSmartCommand('insertHTML', html);
    showToast('Tata letak 2 kolom berdampingan tanpa garis berhasil disisipkan.', 'success');
  });

  // Wizard Template Resmi Surat Tugas & SPPD (3 Halaman 1:1 Microsoft Word)
  function openSuratTugasSppdModal() {
    const school = store.getSchoolProfile();
    const opts = store.getLetterOptions();
    const todayStr = new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
    const nextNum = generateLetterNumber({
      typeCode: 'ST',
      sequence: store.getNextOutgoingSequence()
    });

    modalRoot.innerHTML = `
      <div class="modal-overlay">
        <div class="modal-box" style="max-width: 640px; max-height: 90vh; overflow-y: auto;">
          <div class="modal-header" style="background: #f0fdf4; border-bottom-color: #bbf7d0;">
            <h3 style="color: #14532d;">${ICON_TARGET} Template Resmi Surat Tugas &amp; SPPD (1:1 Word)</h3>
            <button class="btn-icon" id="smart-modal-close">✕</button>
          </div>
          <div class="modal-body" style="display:grid; gap:10px;">
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
              <div>
                <label class="form-label">Paket Halaman Dokumen</label>
                <select id="sppd-wiz-pkg" class="form-input">
                  <option value="both" selected>Surat Tugas + SPPD Lengkap (3 Halaman)</option>
                  <option value="st">Surat Tugas Saja (1 Halaman)</option>
                  <option value="sppd">SPPD + Halaman Keterangan Saja (2 Halaman)</option>
                </select>
              </div>
              <div>
                <label class="form-label">Ukuran Kertas Cetak</label>
                <select id="sppd-wiz-paper" class="form-input">
                  <option value="f4" ${(opts.paperSize || 'f4') === 'f4' ? 'selected' : ''}>F4 / Folio (21,5 x 33 cm)</option>
                  <option value="a4" ${opts.paperSize === 'a4' ? 'selected' : ''}>A4 (21 x 29,7 cm)</option>
                </select>
              </div>
            </div>
            <div>
              <label class="form-label">Nomor Surat Tugas</label>
              <input id="sppd-wiz-nomor" class="form-input" value="${escapeHtmlText(nextNum)}" />
            </div>
            <div>
              <label class="form-label">Daftar Pegawai / Guru yang Ditugaskan (1 baris per orang: Nama | Jabatan)</label>
              <textarea id="sppd-wiz-assignees" class="form-input" rows="3">Taufik Dwi Setyo Ramadhan, S.Pd. | Guru
Siti Laras Maemunah, S.Pd. | Guru</textarea>
            </div>
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
              <div>
                <label class="form-label">Maksud Tugas (di Surat Tugas)</label>
                <input id="sppd-wiz-maksud-st" class="form-input" value="Penjemputan Peserta Didik PKL" />
              </div>
              <div>
                <label class="form-label">Tujuan Lokasi / Instansi</label>
                <input id="sppd-wiz-lokasi" class="form-input" value="DRAGON MOTION PICTURE ASIA" />
              </div>
            </div>
            <div>
              <label class="form-label">Maksud Perjalanan Dinas (Poin 6 SPPD)</label>
              <input id="sppd-wiz-maksud-sppd" class="form-input" value="Penjemputan siswa kelas 12 Multimedia ke tempat PKL" />
            </div>
            <div style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:10px;">
              <div>
                <label class="form-label">Kendaraan</label>
                <input id="sppd-wiz-kendaraan" class="form-input" value="Mobil" />
              </div>
              <div>
                <label class="form-label">Jenis Kendaraan (Hal 3)</label>
                <input id="sppd-wiz-jenis-kend" class="form-input" value="Roda 4" />
              </div>
              <div>
                <label class="form-label">Lama Perjalanan (Hari)</label>
                <input id="sppd-wiz-lama" class="form-input" value="1" />
              </div>
            </div>
            <div style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:10px;">
              <div>
                <label class="form-label">Tanggal Surat</label>
                <input id="sppd-wiz-tgl-surat" class="form-input" value="${escapeHtmlText(todayStr)}" />
              </div>
              <div>
                <label class="form-label">Tanggal Berangkat</label>
                <input id="sppd-wiz-tgl-berangkat" class="form-input" value="${escapeHtmlText(todayStr)}" />
              </div>
              <div>
                <label class="form-label">Tanggal Kembali</label>
                <input id="sppd-wiz-tgl-kembali" class="form-input" value="${escapeHtmlText(todayStr)}" />
              </div>
            </div>
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
              <div>
                <label class="form-label">Tempat Tiba (Pojok Bawah Hal 1)</label>
                <input id="sppd-wiz-tiba" class="form-input" value="Dragon Motion Picture" />
              </div>
              <div>
                <label class="form-label">Pemberi Perintah / Penanda Tangan</label>
                <input id="sppd-wiz-pemberi" class="form-input" value="${escapeHtmlText(school.principalName || 'Dian Nurhadian, S.Pd.')}" />
              </div>
            </div>
          </div>
          <div class="modal-footer">
            <button class="btn btn-secondary" id="smart-modal-cancel">Batal</button>
            <button class="btn btn-primary" id="sppd-wiz-apply">✨ Terapkan ke Kanvas Word Pro (1:1)</button>
          </div>
        </div>
      </div>`;
    modalRoot.querySelector('#smart-modal-close').onclick = closeSmartModal;
    modalRoot.querySelector('#smart-modal-cancel').onclick = closeSmartModal;
    modalRoot.querySelector('#sppd-wiz-apply').onclick = () => {
      const pkg = modalRoot.querySelector('#sppd-wiz-pkg').value;
      const paperSize = modalRoot.querySelector('#sppd-wiz-paper').value || 'f4';
      const curOpts = store.getLetterOptions();
      store.updateLetterOptions({ paperSize });
      const paperSelect = container.querySelector('#smart-paper-select');
      if (paperSelect) paperSelect.value = paperSize;

      const formData = {
        nomorSurat: modalRoot.querySelector('#sppd-wiz-nomor').value.trim(),
        assignees: modalRoot.querySelector('#sppd-wiz-assignees').value,
        maksudTugas: modalRoot.querySelector('#sppd-wiz-maksud-st').value.trim(),
        tujuanLokasi: modalRoot.querySelector('#sppd-wiz-lokasi').value.trim(),
        maksudSppd: modalRoot.querySelector('#sppd-wiz-maksud-sppd').value.trim(),
        kendaraan: modalRoot.querySelector('#sppd-wiz-kendaraan').value.trim() || 'Mobil',
        kendaraanSingkat: modalRoot.querySelector('#sppd-wiz-jenis-kend').value.trim() || 'Roda 4',
        lamaHari: modalRoot.querySelector('#sppd-wiz-lama').value.trim() || '1',
        tanggalSurat: modalRoot.querySelector('#sppd-wiz-tgl-surat').value.trim() || todayStr,
        tglBerangkat: modalRoot.querySelector('#sppd-wiz-tgl-berangkat').value.trim() || todayStr,
        tglKembali: modalRoot.querySelector('#sppd-wiz-tgl-kembali').value.trim() || todayStr,
        tempatTiba: modalRoot.querySelector('#sppd-wiz-tiba').value.trim() || 'Dragon Motion Picture',
        pemberiNama: modalRoot.querySelector('#sppd-wiz-pemberi').value.trim() || school.principalName
      };

      pushUndoSnapshot();
      const docModel = buildSuratTugasSppdDocModel(formData, school, {
        includeSuratTugas: pkg !== 'sppd',
        includeSppd: pkg !== 'st',
        showKop: curOpts.showKop !== false,
        paperSize,
        orientation: 'portrait',
        margins: curOpts.margins
      });
      const converted = docModelToSmartPages(docModel, school);
      smartDoc = { pages: converted.pages, meta: converted.meta };
      store.setState({ smartDocument: smartDoc }, false);
      store.setActiveVaultId(null, 'smart');
      syncGlobalPrintPageStyle(store.getLetterOptions());
      closeSmartModal();
      renderAllSmartPages(smartDoc.pages);
      updateSmartVaultBadge?.();
      showToast('Template Surat Tugas & SPPD berhasil diterapkan 1:1 ke Word Pro!', 'success');
    };
  }

  container.querySelector('#smart-btn-template-sppd')?.addEventListener('click', openSuratTugasSppdModal);

  // Modal Sisipkan Tabel Word / Rincian
  container.querySelector('#smart-btn-insert-table')?.addEventListener('click', () => {
    modalRoot.innerHTML = `
      <div class="modal-overlay">
        <div class="modal-box" style="max-width: 470px;">
          <div class="modal-header">
            <h3>Sisipkan Tabel ke Dokumen</h3>
            <button class="btn-icon" id="smart-modal-close">✕</button>
          </div>
          <div class="modal-body" style="display:grid; gap:12px;">
            <div>
              <label class="form-label">Jenis Tabel &amp; Garis Word</label>
              <select id="sm-tbl-type" class="form-input">
                <option value="grid">Tabel Berbingkai Penuh — Header Abu (Jadwal / Panitia)</option>
                <option value="grid-clean">Tabel Berbingkai Penuh — Header Putih (No | Nama | Jabatan)</option>
                <option value="sppd">Tabel SPPD — Garis Horizontal &amp; Tengah Saja</option>
                <option value="none-2col">Layout 2 Kolom Berdampingan Tanpa Garis (Catatan + TTD)</option>
                <option value="identity">Tabel Rincian Tanpa Garis (Label : Nilai)</option>
              </select>
            </div>
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
              <div><label class="form-label">Jumlah Baris</label><input type="number" id="sm-tbl-rows" class="form-input" min="1" max="25" value="4"></div>
              <div><label class="form-label">Jumlah Kolom</label><input type="number" id="sm-tbl-cols" class="form-input" min="2" max="10" value="3"></div>
            </div>
          </div>
          <div class="modal-footer">
            <button class="btn btn-secondary" id="smart-modal-cancel">Batal</button>
            <button class="btn btn-primary" id="smart-modal-insert-tbl">Sisipkan Tabel</button>
          </div>
        </div>
      </div>`;
    modalRoot.querySelector('#smart-modal-close').onclick = closeSmartModal;
    modalRoot.querySelector('#smart-modal-cancel').onclick = closeSmartModal;
    modalRoot.querySelector('#smart-modal-insert-tbl').onclick = () => {
      const tType = modalRoot.querySelector('#sm-tbl-type').value;
      const nRows = Math.max(1, Number(modalRoot.querySelector('#sm-tbl-rows').value) || 4);
      const nCols = Math.max(2, Number(modalRoot.querySelector('#sm-tbl-cols').value) || 3);

      let html = '';
      if (tType === 'identity') {
        const rowsHtml = Array.from({ length: nRows }, (_, i) =>
          `<tr><td style="width:30%;">Rincian ${i + 1}</td><td>: ........................................</td></tr>`
        ).join('');
        html = `<table class="smart-identity-table" data-smart-role="identity-table"><tbody>${rowsHtml}</tbody></table><p><br></p>`;
      } else if (tType === 'none-2col') {
        const bodyRows = Array.from({ length: nRows }, () =>
          `<tr><td><p>Isi kolom kiri...</p></td><td><p>Isi kolom kanan...</p></td></tr>`
        ).join('');
        html = `<table class="smart-word-table smart-table-none" data-border-style="none" data-header-shading="false" data-header-bold="false" data-col-widths="50,50"><colgroup><col style="width:50%;"><col style="width:50%;"></colgroup><tbody>${bodyRows}</tbody></table><p><br></p>`;
      } else if (tType === 'sppd') {
        const bodyRows = Array.from({ length: nRows }, (_, i) =>
          `<tr><td>${i + 1}.   Uraian poin SPPD</td><td>Isi rincian...</td></tr>`
        ).join('');
        html = `<table class="smart-word-table smart-table-sppd" data-border-style="sppd" data-header-shading="false" data-header-bold="false" data-col-widths="46,54"><colgroup><col style="width:46%;"><col style="width:54%;"></colgroup><tbody>${bodyRows}</tbody></table><p><br></p>`;
      } else if (tType === 'grid-clean') {
        const thCells = Array.from({ length: nCols }, (_, c) => `<th style="text-align: center;">Kolom ${c + 1}</th>`).join('');
        const bodyRows = Array.from({ length: Math.max(1, nRows - 1) }, () =>
          `<tr>${Array.from({ length: nCols }, () => '<td>Isi data...</td>').join('')}</tr>`
        ).join('');
        html = `<table class="smart-word-table no-header-bg no-header-bold" data-header-shading="false" data-header-bold="false"><thead><tr data-header="true">${thCells}</tr></thead><tbody>${bodyRows}</tbody></table><p><br></p>`;
      } else {
        const thCells = Array.from({ length: nCols }, (_, c) => `<th>Kolom ${c + 1}</th>`).join('');
        const bodyRows = Array.from({ length: Math.max(1, nRows - 1) }, () =>
          `<tr>${Array.from({ length: nCols }, () => '<td>Isi data...</td>').join('')}</tr>`
        ).join('');
        html = `<table class="smart-word-table"><thead><tr>${thCells}</tr></thead><tbody>${bodyRows}</tbody></table><p><br></p>`;
      }
      syncGlobalPrintPageStyle(store.getLetterOptions());
      closeSmartModal();
      execSmartCommand('insertHTML', html);
      showToast('Tabel berhasil disisipkan.', 'success');
    };
  });

  function hasSmartCanvasContent() {
    const pageEls = [...pagesStack.querySelectorAll('.smart-editable-page')];
    if (pageEls.length > 1) return true;
    return pageEls.some((el) => {
      const txt = (el.innerText || el.textContent || '').replace(/\u00a0/g, ' ').trim();
      if (txt.length > 0) return true;
      return Boolean(el.querySelector('table, img, ul, ol'));
    });
  }

  const statusVaultBadgeEl = container.querySelector('#smart-status-vault-badge');

  function updateSmartVaultBadge() {
    renderCanvasVaultBadge(statusVaultBadgeEl, store.getActiveVaultEntry('smart'));
    updateSmartSaveIndicator({ savedToVault: true });
  }
  updateSmartVaultBadge();

  function resetSmartCanvasToEmpty(notify = true) {
    pushUndoSnapshot();
    smartDoc = createEmptySmartDocument();
    store.setState({ smartDocument: smartDoc }, false);
    store.setActiveVaultId(null, 'smart');
    hideAiCompareBar();
    renderAllSmartPages(smartDoc.pages);
    updateSmartVaultBadge();
    if (notify) {
      showToast('Lembar Word Pro berhasil dikosongkan ke kertas baru.', 'info');
    }
  }

  function writeSmartUpdateToEntry(targetEntry, { unlock = false, clearAfterSave = false } = {}) {
    const snap = extractSmartSnapshotFromPages(collectLivePagesData(), store.getSchoolProfile());
    const liveDocModel = getActiveSmartDocModel();
    const summaryText = (snap.blocks || [])
      .filter(b => b.type === 'paragraph' && b.text)
      .map(b => b.text)
      .join(' ')
      .slice(0, 160);

    const versionNote = unlock
      ? 'Buka kunci & simpan perubahan'
      : (lastAiCompareState ? 'Simpan pasca revisi AI' : 'Simpan perubahan (Ctrl+S)');

    store.updateLetterVaultEntry(targetEntry.id, {
      nomorSurat: snap.nomorSurat || targetEntry.nomorSurat || '-',
      tanggalSurat: snap.signer?.date || targetEntry.tanggalSurat || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
      targetPerson: snap.tujuan || targetEntry.targetPerson || '-',
      perihal: snap.hal || targetEntry.perihal || targetEntry.title,
      summary: summaryText || targetEntry.summary || 'Dokumen persuratan dari Kanvas Word Pro.',
      document: liveDocModel,
      ...(unlock ? { isLockedTemplate: false, isMasterTemplate: false } : {})
    }, { versionNote });
    lastLocalSavedTime = new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
    syncGlobalPrintPageStyle(store.getLetterOptions());
    updateSmartVaultBadge();
    if (clearAfterSave) {
      resetSmartCanvasToEmpty(false);
      showToast(`✓ Perubahan disimpan ke "${targetEntry.title}" & kertas dikosongkan!`, 'success');
    } else if (unlock) {
      showToast(`🔓 Kunci dibuka & perubahan disimpan ke "${targetEntry.title}". Anda dapat mengunci Master ⭐ kembali di Gudang Surat.`, 'success');
    } else {
      showToast(`✓ Tersimpan ke "${targetEntry.title}" (Ctrl+S) — file Gudang diperbarui tanpa duplikat.`, 'success');
    }
  }

  function performSmartSave({ clearAfterSave = false } = {}) {
    const activeEntry = store.getActiveVaultEntry('smart');
    if (!activeEntry) {
      openSmartSaveVaultModal({ mode: 'save', clearAfterSave });
      return;
    }
    if (activeEntry.isLockedTemplate || activeEntry.isMasterTemplate) {
      openLockedMasterSaveChoiceModal(modalRoot, activeEntry, {
        clearAfterSave,
        onSaveAsCopy: () => openSmartSaveVaultModal({ mode: 'save-as', clearAfterSave }),
        onUnlockAndSave: () => writeSmartUpdateToEntry(activeEntry, { unlock: true, clearAfterSave })
      });
      return;
    }
    writeSmartUpdateToEntry(activeEntry, { clearAfterSave });
  }

  function openSmartSaveVaultModal({ mode = 'save', clearAfterSave = false } = {}) {
    const isSaveAs = mode === 'save-as';
    const activeEntry = store.getActiveVaultEntry('smart');
    const livePages = collectLivePagesData();
    const liveModel = getActiveSmartDocModel();
    const docSchool = store.getEffectiveSchoolForDocument(liveModel) || store.getSchoolProfile();
    const snap = extractSmartSnapshotFromPages(livePages, docSchool);
    const autoMeta = extractDocumentMetadata({
      docModel: liveModel,
      smartPages: livePages,
      schoolProfile: docSchool,
      fallbackTitle: snap.letterTitle,
      fallbackNomor: snap.nomorSurat,
      fallbackDate: snap.signer?.date || '',
      fallbackPerihal: snap.hal || '',
      fallbackRecipient: snap.tujuan || ''
    });
    const defaultTitle = isSaveAs
      ? `Salinan - ${activeEntry?.title || autoMeta.formattedTitle}`
      : autoMeta.formattedTitle;

    modalRoot.innerHTML = `
      <div class="modal-overlay">
        <div class="modal-box" style="max-width: 510px;">
          <div class="modal-header">
            <h3>${isSaveAs ? ICON_COPY_DOC : ICON_SAVE} ${clearAfterSave ? 'Simpan ke Gudang Sebelum Mengosongkan' : (isSaveAs ? 'Save As / Duplikat ke Gudang Surat' : 'Simpan Dokumen Baru ke Gudang Surat')}</h3>
            <button class="btn-icon" id="smart-modal-close">✕</button>
          </div>
          <div class="modal-body" style="display:grid; gap:10px;">
            ${isSaveAs && activeEntry ? `<div class="auth-alert" style="background:#eff6ff; border-color:#bfdbfe; color:#1e3a8a; font-size:12px;">Membuat salinan independen baru dari <strong>${escapeHtmlText(activeEntry.title)}</strong> agar dokumen asli tetap aman tanpa tertimpa.</div>` : ''}
            <div id="sm-v-scan-slot">${buildMetaAutoScanBarHtml()}</div>
            <div>
              <label class="form-label">Judul Arsip (Terformat Otomatis)</label>
              <input id="sm-v-title" class="form-input" value="${escapeHtmlText(defaultTitle)}">
              <div id="sm-v-dup-warn" class="hidden" style="margin-top:6px; padding:7px 10px; background:#fffbeb; border:1px solid #fde68a; border-radius:6px; font-size:11.5px; color:#92400e;"></div>
            </div>
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:8px;">
              <div><label class="form-label">Kategori</label><input id="sm-v-cat" class="form-input" value="${escapeHtmlText(activeEntry?.category || autoMeta.category)}"></div>
              <div><label class="form-label">Tanggal Surat</label><input id="sm-v-date" class="form-input" value="${escapeHtmlText(autoMeta.tanggalSurat)}"></div>
            </div>
            <div><label class="form-label">Atas Nama / Pihak Terkait</label><input id="sm-v-person" class="form-input" value="${escapeHtmlText(autoMeta.recipientName)}"></div>
            <div><label class="form-label">Nomor Surat</label><input id="sm-v-nomor" class="form-input" value="${escapeHtmlText(autoMeta.nomorSurat || snap.nomorSurat)}"></div>
            <div><label class="form-label">Keperluan / Perihal</label><input id="sm-v-hal" class="form-input" value="${escapeHtmlText(autoMeta.perihal)}"></div>
          </div>
          <div class="modal-footer">
            <button class="btn btn-secondary" id="smart-modal-cancel">Batal</button>
            <button class="btn btn-secondary hidden" id="smart-modal-overwrite-btn" title="Timpa file gudang dengan judul yang sama">Timpa File Sama</button>
            <button class="btn btn-primary" id="smart-modal-save-ok">${clearAfterSave ? 'Simpan &amp; Kosongkan Kertas' : (isSaveAs ? 'Save As (Simpan Duplikat)' : 'Simpan ke Gudang')}</button>
          </div>
        </div>
      </div>`;

    const titleInput = modalRoot.querySelector('#sm-v-title');
    const dupWarnEl = modalRoot.querySelector('#sm-v-dup-warn');
    const overwriteBtn = modalRoot.querySelector('#smart-modal-overwrite-btn');
    let matchedDup = null;

    const cleanupSmartVaultScan = wireMetaAutoScanBar({
      rootEl: modalRoot.querySelector('.modal-body'),
      getContext: () => ({
        docModel: getActiveSmartDocModel(),
        smartPages: collectLivePagesData(),
        schoolProfile: docSchool
      }),
      fields: {
        titleEl: titleInput,
        categoryEl: modalRoot.querySelector('#sm-v-cat'),
        dateEl: modalRoot.querySelector('#sm-v-date'),
        recipientEl: modalRoot.querySelector('#sm-v-person'),
        perihalEl: modalRoot.querySelector('#sm-v-hal')
      },
      forTemplate: false,
      autoAiOnOpen: !isSaveAs
    });

    const closeSmartSaveModal = () => {
      cleanupSmartVaultScan();
      closeSmartModal();
    };

    const checkDupTitle = () => {
      const val = titleInput?.value.trim() || '';
      matchedDup = store.findLetterVaultByTitle(val);
      if (matchedDup) {
        if (matchedDup.isLockedTemplate || matchedDup.isMasterTemplate) {
          dupWarnEl.textContent = `🔒 "${matchedDup.title}" berstatus ⭐ Master Terkunci dan dilindungi agar tidak tertimpa. Ubah judul untuk menyimpan Salinan Baru.`;
          dupWarnEl.classList.remove('hidden');
          overwriteBtn?.classList.add('hidden');
        } else {
          dupWarnEl.textContent = `⚠️ Dokumen berjudul "${matchedDup.title}" sudah ada di Gudang. Ganti judul untuk membuat file duplikat terpisah, atau klik "Timpa File Sama".`;
          dupWarnEl.classList.remove('hidden');
          overwriteBtn?.classList.remove('hidden');
        }
      } else {
        dupWarnEl.classList.add('hidden');
        overwriteBtn?.classList.add('hidden');
      }
    };
    titleInput?.addEventListener('input', checkDupTitle);
    checkDupTitle();

    modalRoot.querySelector('#smart-modal-close').onclick = closeSmartSaveModal;
    modalRoot.querySelector('#smart-modal-cancel').onclick = closeSmartSaveModal;

    const buildEntryPayload = (titleVal) => ({
      title: titleVal,
      category: modalRoot.querySelector('#sm-v-cat').value.trim() || autoMeta.category,
      nomorSurat: modalRoot.querySelector('#sm-v-nomor').value.trim() || '-',
      tanggalSurat: modalRoot.querySelector('#sm-v-date')?.value.trim() || autoMeta.tanggalSurat,
      targetPerson: modalRoot.querySelector('#sm-v-person')?.value.trim() || autoMeta.recipientName || '-',
      perihal: modalRoot.querySelector('#sm-v-hal').value.trim() || titleVal,
      summary: autoMeta.summary || 'Dokumen persuratan dari Kanvas Word Pro.',
      document: getActiveSmartDocModel(),
      isAiReference: true,
      source: 'created'
    });

    if (overwriteBtn) {
      overwriteBtn.onclick = () => {
        if (!matchedDup) return;
        const titleVal = titleInput.value.trim() || matchedDup.title;
        store.updateLetterVaultEntry(matchedDup.id, buildEntryPayload(titleVal));
        store.setActiveVaultId(matchedDup.id, 'smart');
        syncGlobalPrintPageStyle(store.getLetterOptions());
        closeSmartModal();
        if (clearAfterSave) {
          resetSmartCanvasToEmpty(false);
          showToast(`✓ File "${titleVal}" di Gudang diperbarui & kertas dikosongkan!`, 'success');
        } else {
          updateSmartVaultBadge();
          showToast(`✓ File "${titleVal}" di Gudang Surat berhasil diperbarui!`, 'success');
        }
      };
    }

    modalRoot.querySelector('#smart-modal-save-ok').onclick = () => {
      const titleVal = titleInput.value.trim() || defaultTitle;
      const created = store.addLetterVaultEntry(buildEntryPayload(titleVal));
      store.setActiveVaultId(created.id, 'smart');
      syncGlobalPrintPageStyle(store.getLetterOptions());
      closeSmartModal();
      if (clearAfterSave) {
        resetSmartCanvasToEmpty(false);
        showToast(`"${titleVal}" tersimpan ke Gudang Surat & kertas siap dipakai ulang!`, 'success');
      } else {
        updateSmartVaultBadge();
        showToast(
          isSaveAs
            ? `✓ Save As berhasil! Salinan "${titleVal}" dibuat & kini aktif diedit (Ctrl+S).`
            : `✓ "${titleVal}" disimpan ke Gudang Surat! Edit berikutnya cukup tekan Ctrl+S.`,
          'success'
        );
      }
    };
  }

  function buildSmartEntryBase(recordedNomorOverride = '') {
    const liveModel = getActiveSmartDocModel();
    const docSchool = store.getEffectiveSchoolForDocument(liveModel);
    const livePages = collectLivePagesData();
    const snap = extractSmartSnapshotFromPages(livePages, docSchool);
    const autoMeta = extractDocumentMetadata({
      docModel: liveModel,
      smartPages: livePages,
      schoolProfile: docSchool,
      fallbackTitle: snap.letterTitle,
      fallbackNomor: snap.nomorSurat,
      fallbackDate: snap.signer?.date || '',
      fallbackPerihal: snap.hal || '',
      fallbackRecipient: snap.tujuan || ''
    });
    const nextInfo = store.getNextFormattedOutgoingNomor({
      title: autoMeta.formattedTitle || snap.letterTitle,
      perihal: autoMeta.perihal || snap.hal,
      currentNomor: autoMeta.nomorSurat || snap.nomorSurat,
      forcedTypeCode: autoMeta.typeCode,
      schoolProfile: docSchool
    });
    return {
      snap,
      autoMeta,
      nextInfo,
      school: docSchool,
      entryBase: {
        nomorSurat: recordedNomorOverride || nextInfo.nomorSurat,
        sequenceNumber: nextInfo.sequenceNumber,
        typeCode: nextInfo.typeCode,
        year: nextInfo.year,
        title: autoMeta.formattedTitle || snap.letterTitle,
        perihal: autoMeta.perihal || snap.hal || snap.letterTitle,
        recipientName: autoMeta.recipientName || snap.tujuan || 'Umum',
        tanggalSurat: autoMeta.tanggalSurat || snap.signer?.date || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
        signerName: snap.signer?.name || docSchool?.principalName || 'Kepala Sekolah',
        summary: autoMeta.summary,
        sourceMode: 'smart',
        document: liveModel,
        smartPagesSnapshot: {
          pages: livePages,
          letterOptions: store.getLetterOptions()
        }
      }
    };
  }

  activeSmartPagesRoot = pagesStack;
  activeSmartAuthorizeHandler = async (signerRole = 'kepala_sekolah', clearAfterSave = false) => {
    const { entryBase, school: docSchool } = buildSmartEntryBase('');
    const result = await authorizeAndRecordOutgoing({
      entryBase,
      school: docSchool,
      signerRole,
      authorizeSah: true,
      syncActiveDocMode: 'smart',
      clearAfterSave
    });
    if (!clearAfterSave) refreshSmartFloatingQrOverlay();
    return result;
  };
  activeSmartRequestAccHandler = async (signerRole = 'kepala_sekolah', clearAfterSave = false) => {
    const { entryBase, school: docSchool } = buildSmartEntryBase('');
    return requestAccForOutgoing({
      entryBase,
      school: docSchool,
      requestedSignerRole: signerRole,
      syncActiveDocMode: 'smart',
      clearAfterSave
    });
  };

  function openSmartRecordOutgoingModal({ clearAfterSave = false } = {}) {
    const { snap, autoMeta, nextInfo, school } = buildSmartEntryBase('');

    modalRoot.innerHTML = `
      <div class="modal-overlay">
        <div class="modal-box" style="max-width: 560px;">
          <div class="modal-header">
            <h3>${ICON_CLIPBOARD_CHECK} ${clearAfterSave ? 'Sahkan &amp; Catat ke Surat Keluar Sebelum Mengosongkan' : 'Catat &amp; Sahkan Surat Keluar'}</h3>
            <button class="btn-icon" id="smart-modal-close">✕</button>
          </div>
          <div class="modal-body" style="display:grid; gap:10px;">
            <div id="sm-out-scan-slot">${buildMetaAutoScanBarHtml()}</div>
            <div style="display:grid; grid-template-columns: 1fr 165px; gap:8px;">
              <div>
                <label class="form-label">Nomor Surat Resmi (Otomatis Berurutan)</label>
                <input id="sm-out-nomor" class="form-input" value="${escapeHtmlText(nextInfo.nomorSurat)}">
              </div>
              <div>
                <label class="form-label">Kode Jenis Surat</label>
                <select id="sm-out-type-code" class="form-input">
                  ${Object.values(LETTER_TYPE_CODES).map(c => `<option value="${c.code}" ${c.code === nextInfo.typeCode ? 'selected' : ''}>${c.code} — ${c.label.replace(/\s*\([^)]+\)$/, '')}</option>`).join('')}
                </select>
              </div>
            </div>
            <div><label class="form-label">Judul Surat (Terformat Otomatis)</label><input id="sm-out-title" class="form-input" value="${escapeHtmlText(autoMeta.formattedTitle)}"></div>
            <div style="display:grid; grid-template-columns: 1fr 165px; gap:8px;">
              <div><label class="form-label">Atas Nama / Pihak Ditugaskan / Penerima</label><input id="sm-out-to" class="form-input" value="${escapeHtmlText(autoMeta.recipientName)}"></div>
              <div><label class="form-label">Tanggal Surat</label><input id="sm-out-date" class="form-input" value="${escapeHtmlText(autoMeta.tanggalSurat)}"></div>
            </div>
            <div><label class="form-label">Keperluan / Perihal</label><input id="sm-out-hal" class="form-input" value="${escapeHtmlText(autoMeta.perihal)}"></div>
            ${buildSignerRoleSelectorHtml({
              radioName: 'sm-out-signer-role',
              checkboxId: 'sm-chk-authorize-sah',
              defaultRole: store.getState().smartDocument?.verification?.signerRole || 'kepala_sekolah',
              school
            })}
          </div>
          <div class="modal-footer">
            <button class="btn btn-secondary" id="smart-modal-cancel">Batal</button>
            <button class="btn btn-primary" id="smart-modal-save-out">${clearAfterSave ? '🛡️ Sahkan, Catat &amp; Kosongkan' : '🛡️ Sahkan &amp; Simpan Surat Keluar'}</button>
          </div>
        </div>
      </div>`;

    const titleEl = modalRoot.querySelector('#sm-out-title');
    const halEl = modalRoot.querySelector('#sm-out-hal');
    const toEl = modalRoot.querySelector('#sm-out-to');
    const dateEl = modalRoot.querySelector('#sm-out-date');
    const nomorEl = modalRoot.querySelector('#sm-out-nomor');
    const typeSelectEl = modalRoot.querySelector('#sm-out-type-code');

    const syncNomorForType = (forcedCode) => {
      const updatedSeq = store.getNextFormattedOutgoingNomor({
        title: titleEl?.value || autoMeta.formattedTitle,
        perihal: halEl?.value || autoMeta.perihal,
        currentNomor: snap.nomorSurat,
        forcedTypeCode: forcedCode || typeSelectEl?.value || nextInfo.typeCode,
        schoolProfile: school
      });
      if (nomorEl && updatedSeq?.nomorSurat) nomorEl.value = updatedSeq.nomorSurat;
    };
    typeSelectEl?.addEventListener('change', () => syncNomorForType(typeSelectEl.value));

    const cleanupOutScan = wireMetaAutoScanBar({
      rootEl: modalRoot.querySelector('.modal-body'),
      getContext: () => ({
        docModel: getActiveSmartDocModel(),
        smartPages: collectLivePagesData(),
        schoolProfile: school
      }),
      fields: {
        titleEl,
        perihalEl: halEl,
        recipientEl: toEl,
        dateEl,
        typeSelectEl
      },
      forTemplate: false,
      autoAiOnOpen: true,
      onMetaApplied: (m) => {
        if (m?.typeCode) syncNomorForType(m.typeCode);
      }
    });

    const closeOutModal = () => {
      cleanupOutScan();
      closeSmartModal();
    };
    modalRoot.querySelector('#smart-modal-close').onclick = closeOutModal;
    modalRoot.querySelector('#smart-modal-cancel').onclick = closeOutModal;
    const btnSaveOut = modalRoot.querySelector('#smart-modal-save-out');
    const syncSmartSaveLabel = () => {
      const checked = modalRoot.querySelector('input[name="sm-out-signer-role"]:checked');
      const authChecked = modalRoot.querySelector('#sm-chk-authorize-sah')?.checked !== false;
      const reqAcc = authChecked && checked?.dataset?.requiresAcc === '1';
      if (btnSaveOut) {
        btnSaveOut.textContent = reqAcc
          ? (clearAfterSave ? '📨 Ajukan ACC, Catat & Kosongkan' : '📨 Ajukan ACC & Simpan Surat Keluar')
          : authChecked
          ? (clearAfterSave ? '🛡️ Sahkan, Catat & Kosongkan' : '🛡️ Sahkan & Simpan Surat Keluar')
          : (clearAfterSave ? '💾 Catat & Kosongkan' : '💾 Simpan Surat Keluar');
      }
    };
    modalRoot.querySelectorAll('input[name="sm-out-signer-role"], #sm-chk-authorize-sah').forEach(el => {
      el.addEventListener('change', syncSmartSaveLabel);
    });
    syncSmartSaveLabel();

    btnSaveOut.onclick = async () => {
      const stNow = store.getState();
      const selectedRadio = modalRoot.querySelector('input[name="sm-out-signer-role"]:checked');
      const selectedRole = selectedRadio?.value
        || resolveDefaultAllowedSignerRole(stNow.authUser, stNow.authPermissions, 'kepala_sekolah');
      const authorizeSah = modalRoot.querySelector('#sm-chk-authorize-sah')?.checked !== false;
      const roleAllowed = canSignWithRole(stNow.authUser, stNow.authPermissions, selectedRole).allowed;
      const requiresAcc = authorizeSah && (selectedRadio?.dataset?.requiresAcc === '1' || !roleAllowed);

      const recordedNomor = (nomorEl?.value || '').trim();
      const chosenTypeCode = (typeSelectEl?.value || nextInfo.typeCode || '').toUpperCase();
      const titleVal = (titleEl?.value || '').trim() || autoMeta.formattedTitle;
      const halVal = (halEl?.value || '').trim() || titleVal;
      const toVal = (toEl?.value || '').trim() || autoMeta.recipientName || 'Umum';
      const dateVal = (dateEl?.value || '').trim() || autoMeta.tanggalSurat;

      const liveModel = getActiveSmartDocModel();

      btnSaveOut.disabled = true;
      btnSaveOut.textContent = requiresAcc
        ? 'Mengirim Antrean ACC Kepala Sekolah...'
        : authorizeSah
        ? 'Menyimpan & Mengesahkan QR...'
        : 'Menyimpan ke Surat Keluar...';

      try {
        const entryBase = {
          nomorSurat: recordedNomor,
          sequenceNumber: nextInfo.sequenceNumber,
          typeCode: chosenTypeCode,
          year: nextInfo.year,
          title: titleVal,
          perihal: halVal,
          recipientName: toVal,
          tanggalSurat: dateVal,
          signerName: snap.signer?.name || school?.principalName || 'Kepala Sekolah',
          summary: autoMeta.summary,
          sourceMode: 'smart',
          document: liveModel,
          smartPagesSnapshot: {
            pages: collectLivePagesData(),
            letterOptions: store.getLetterOptions()
          }
        };

        if (requiresAcc) {
          const { recorded } = await requestAccForOutgoing({
            entryBase,
            school,
            requestedSignerRole: selectedRole,
            syncActiveDocMode: 'smart',
            clearAfterSave
          });
          syncGlobalPrintPageStyle(store.getLetterOptions());
      closeSmartModal();
          if (clearAfterSave) {
            resetSmartCanvasToEmpty(false);
          }
          showToast(
            `⏳ Surat #${String(recorded.sequenceNumber).padStart(3, '0')} (${recorded.nomorSurat}) masuk Antrean ACC Kepala Sekolah! Anda bebas lanjut membuat surat lain.`,
            'info'
          );
          return;
        }

        const { recorded, cloudResult } = await authorizeAndRecordOutgoing({
          entryBase,
          school,
          signerRole: selectedRole,
          authorizeSah,
          syncActiveDocMode: 'smart',
          clearAfterSave
        });

        syncGlobalPrintPageStyle(store.getLetterOptions());
      closeSmartModal();
        if (clearAfterSave) {
          resetSmartCanvasToEmpty(false);
          showToast(`Tercatat (${recorded.nomorSurat})${authorizeSah ? ' & DISAHKAN' : ''} ke Surat Keluar, kertas dikosongkan!`, 'success');
        } else {
          refreshSmartFloatingQrOverlay();
          showToast(
            `Tercatat! Nomor ${recorded.nomorSurat}${authorizeSah ? (cloudResult.ok ? ' DISAHKAN (Barcode QR Aktif di Kanvas — bebas digeser atau disembunyikan)' : ' DISAHKAN (Lokal)') : ' tersimpan'} ke Surat Keluar.`,
            'success'
          );
        }
      } catch (err) {
        showToast(err?.message || 'Gagal mencatat Surat Keluar.', 'error');
      } finally {
        btnSaveOut.disabled = false;
        syncSmartSaveLabel();
      }
    };
  }

  function openSmartRefreshConfirmModal() {
    if (!hasSmartCanvasContent()) {
      resetSmartCanvasToEmpty(false);
      showToast('Lembar Word Pro sudah kosong dan siap digunakan.', 'info');
      return;
    }

    modalRoot.innerHTML = `
      <div class="modal-overlay">
        <div class="modal-box" style="max-width: 480px;">
          <div class="modal-header" style="background: #fffbeb; border-bottom-color: #fde68a;">
            <h3 style="color: #92400e;">${ICON_REFRESH} Kosongkan / Refresh Dokumen?</h3>
            <button class="btn-icon" id="smart-modal-close">✕</button>
          </div>
          <div class="modal-body">
            <p style="margin: 0 0 8px 0; line-height: 1.5; color: #334155;">
              Dokumen di <strong>Word Pro</strong> saat ini berisi teks atau elemen aktif. Sebelum dikosongkan ke lembar baru, pilih tindakan yang Anda inginkan:
            </p>
            <div class="refresh-choice-list">
              <button type="button" class="refresh-choice-btn refresh-choice-vault" id="sm-ref-vault">
                <div>
                  <span>Simpan ke Gudang Surat lalu Kosongkan</span>
                  <small>Arsipkan naskah ke Gudang Surat &amp; jadikan acuan AI sebelum kertas dibersihkan</small>
                </div>
                <span>${ICON_SAVE}</span>
              </button>
              <button type="button" class="refresh-choice-btn refresh-choice-outgoing" id="sm-ref-outgoing">
                <div>
                  <span>Simpan ke Surat Keluar lalu Kosongkan</span>
                  <small>Catat resmi ke Buku Agenda Surat Keluar dengan nomor berurutan lalu bersihkan</small>
                </div>
                <span>${ICON_SEND_OUT}</span>
              </button>
              <button type="button" class="refresh-choice-btn refresh-choice-discard" id="sm-ref-discard">
                <div>
                  <span>Kosongkan Langsung Tanpa Simpan</span>
                  <small>Hapus seluruh isi dokumen saat ini dan mulai dari kertas kosong</small>
                </div>
                <span>${ICON_REFRESH}</span>
              </button>
            </div>
          </div>
          <div class="modal-footer">
            <button class="btn btn-secondary" id="smart-modal-cancel">Batal (Pertahankan Isi)</button>
          </div>
        </div>
      </div>`;

    modalRoot.querySelector('#smart-modal-close').onclick = closeSmartModal;
    modalRoot.querySelector('#smart-modal-cancel').onclick = closeSmartModal;
    modalRoot.querySelector('#sm-ref-vault').onclick = () => performSmartSave({ clearAfterSave: true });
    modalRoot.querySelector('#sm-ref-outgoing').onclick = () => openSmartRecordOutgoingModal({ clearAfterSave: true });
    modalRoot.querySelector('#sm-ref-discard').onclick = () => {
      syncGlobalPrintPageStyle(store.getLetterOptions());
      closeSmartModal();
      resetSmartCanvasToEmpty(true);
    };
  }

  // Simpan (Save: Ctrl+S) & Save As (Duplikat: Ctrl+Shift+S / F12)
  container.querySelector('#smart-btn-save-vault')?.addEventListener('click', () => performSmartSave({ clearAfterSave: false }));
  container.querySelector('#smart-btn-save-as-vault')?.addEventListener('click', () => openSmartSaveVaultModal({ mode: 'save-as', clearAfterSave: false }));

  // =========================================================================
  // CARI & GANTI CEPAT (CTRL+F / CTRL+H) LINTAS HALAMAN WORD PRO
  // =========================================================================
  const findBarEl = container.querySelector('#smart-find-replace-bar');
  const findInputEl = container.querySelector('#smart-find-input');
  const replaceInputEl = container.querySelector('#smart-replace-input');
  const findCaseEl = container.querySelector('#smart-find-case');
  const findCountEl = container.querySelector('#smart-find-count');
  let findMatches = [];
  let activeFindIdx = -1;

  function clearFindHighlights() {
    if (typeof CSS !== 'undefined' && CSS.highlights) {
      CSS.highlights.delete('smart-find-all');
      CSS.highlights.delete('smart-find-current');
    }
  }

  function collectSmartFindMatches() {
    findMatches = [];
    const query = findInputEl?.value || '';
    if (!query) {
      activeFindIdx = -1;
      clearFindHighlights();
      if (findCountEl) findCountEl.textContent = '0/0';
      return;
    }

    const caseSensitive = Boolean(findCaseEl?.checked);
    const needle = caseSensitive ? query : query.toLowerCase();
    const pagesEls = [...pagesStack.querySelectorAll('.smart-editable-page')];

    pagesEls.forEach((pageEl) => {
      const walker = document.createTreeWalker(pageEl, NodeFilter.SHOW_TEXT, null);
      let textNode;
      while ((textNode = walker.nextNode())) {
        const raw = textNode.nodeValue || '';
        if (!raw) continue;
        const hay = caseSensitive ? raw : raw.toLowerCase();
        let pos = 0;
        while ((pos = hay.indexOf(needle, pos)) !== -1) {
          findMatches.push({ node: textNode, start: pos, end: pos + query.length });
          pos += Math.max(1, query.length);
        }
      }
    });

    if (findMatches.length === 0) {
      activeFindIdx = -1;
      clearFindHighlights();
    } else {
      if (activeFindIdx < 0 || activeFindIdx >= findMatches.length) activeFindIdx = 0;
      paintSmartFindHighlights();
    }
    if (findCountEl) {
      findCountEl.textContent = findMatches.length > 0 ? `${activeFindIdx + 1}/${findMatches.length}` : '0/0';
    }
  }

  function paintSmartFindHighlights() {
    clearFindHighlights();
    if (findMatches.length === 0 || typeof Range === 'undefined') return;
    if (typeof CSS !== 'undefined' && CSS.highlights && typeof Highlight !== 'undefined') {
      const allRanges = [];
      let curRange = null;
      findMatches.forEach((m, i) => {
        try {
          const r = document.createRange();
          r.setStart(m.node, m.start);
          r.setEnd(m.node, m.end);
          if (i === activeFindIdx) curRange = r;
          else allRanges.push(r);
        } catch {}
      });
      if (allRanges.length > 0) CSS.highlights.set('smart-find-all', new Highlight(...allRanges));
      if (curRange) CSS.highlights.set('smart-find-current', new Highlight(curRange));
    }
  }

  function jumpToFindMatch(step = 1) {
    if (findMatches.length === 0) collectSmartFindMatches();
    if (findMatches.length === 0) return;
    activeFindIdx = (activeFindIdx + step + findMatches.length) % findMatches.length;
    const match = findMatches[activeFindIdx];
    paintSmartFindHighlights();
    if (findCountEl) findCountEl.textContent = `${activeFindIdx + 1}/${findMatches.length}`;
    try {
      const r = document.createRange();
      r.setStart(match.node, match.start);
      r.setEnd(match.node, match.end);
      const sel = window.getSelection();
      if (sel) {
        sel.removeAllRanges();
        sel.addRange(r);
      }
      match.node.parentElement?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    } catch {}
  }

  function openSmartFindReplaceBar({ focusReplace = false } = {}) {
    if (!findBarEl) return;
    findBarEl.classList.remove('hidden');
    const selectedTxt = window.getSelection()?.toString()?.trim();
    if (selectedTxt && selectedTxt.length <= 60 && findInputEl && !findInputEl.value) {
      findInputEl.value = selectedTxt;
    }
    collectSmartFindMatches();
    if (focusReplace && findInputEl?.value && replaceInputEl) {
      replaceInputEl.focus();
      replaceInputEl.select();
    } else if (findInputEl) {
      findInputEl.focus();
      findInputEl.select();
    }
  }

  function closeSmartFindReplaceBar() {
    findBarEl?.classList.add('hidden');
    clearFindHighlights();
    findMatches = [];
    activeFindIdx = -1;
  }

  container.querySelector('#smart-btn-find-replace')?.addEventListener('click', () => {
    if (findBarEl && !findBarEl.classList.contains('hidden')) closeSmartFindReplaceBar();
    else openSmartFindReplaceBar();
  });
  container.querySelector('#smart-find-close')?.addEventListener('click', closeSmartFindReplaceBar);
  findInputEl?.addEventListener('input', () => { activeFindIdx = 0; collectSmartFindMatches(); });
  findCaseEl?.addEventListener('change', () => { activeFindIdx = 0; collectSmartFindMatches(); });
  container.querySelector('#smart-find-next')?.addEventListener('click', () => jumpToFindMatch(1));
  container.querySelector('#smart-find-prev')?.addEventListener('click', () => jumpToFindMatch(-1));

  findInputEl?.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      jumpToFindMatch(e.shiftKey ? -1 : 1);
    } else if (e.key === 'Escape') {
      closeSmartFindReplaceBar();
    }
  });
  replaceInputEl?.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeSmartFindReplaceBar();
  });

  container.querySelector('#smart-replace-one')?.addEventListener('click', () => {
    if (findMatches.length === 0) collectSmartFindMatches();
    if (activeFindIdx < 0 || !findMatches[activeFindIdx]) return;
    const replacement = replaceInputEl?.value ?? '';
    const m = findMatches[activeFindIdx];
    pushUndoSnapshot();
    const curVal = m.node.nodeValue || '';
    m.node.nodeValue = curVal.slice(0, m.start) + replacement + curVal.slice(m.end);
    schedulePersistSmartDoc();
    scheduleReflowPagination();
    collectSmartFindMatches();
    if (findMatches.length > 0) jumpToFindMatch(0);
    showToast('1 teks berhasil diganti.', 'success');
  });

  container.querySelector('#smart-replace-all')?.addEventListener('click', () => {
    collectSmartFindMatches();
    if (findMatches.length === 0) {
      showToast('Tidak ada teks yang cocok untuk diganti.', 'warning');
      return;
    }
    const replacement = replaceInputEl?.value ?? '';
    const totalReplaced = findMatches.length;
    pushUndoSnapshot();
    // Ganti mundur per node agar indeks offset teks sebelumnya tidak bergeser
    const byNode = new Map();
    findMatches.forEach((m) => {
      if (!byNode.has(m.node)) byNode.set(m.node, []);
      byNode.get(m.node).push(m);
    });
    byNode.forEach((list, textNode) => {
      let val = textNode.nodeValue || '';
      for (let i = list.length - 1; i >= 0; i--) {
        const m = list[i];
        val = val.slice(0, m.start) + replacement + val.slice(m.end);
      }
      textNode.nodeValue = val;
    });
    schedulePersistSmartDoc();
    scheduleReflowPagination();
    collectSmartFindMatches();
    showToast(`✓ ${totalReplaced} kecocokan berhasil diganti di seluruh halaman.`, 'success');
  });

  // =========================================================================
  // RIWAYAT VERSI DOKUMEN GUDANG & SNAPSHOT SESI (VERSION HISTORY)
  // =========================================================================
  function openSmartVersionHistoryModal() {
    const activeEntry = store.getActiveVaultEntry('smart');
    const vaultHistory = Array.isArray(activeEntry?.versionHistory) ? activeEntry.versionHistory : [];
    const sessionUndoCount = smartUndoStack.length;

    const vaultRowsHtml = !activeEntry
      ? `<div style="padding:14px; background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; font-size:12.5px; color:#475569;">
           Dokumen ini masih berupa draf baru. Simpan ke <strong>Gudang Surat (Ctrl+S)</strong> terlebih dahulu agar setiap revisinya tercatat permanen di riwayat versi.
         </div>`
      : vaultHistory.length === 0
      ? `<div style="padding:14px; background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; font-size:12.5px; color:#475569;">
           Belum ada riwayat versi tersimpan untuk <strong>${escapeHtmlText(activeEntry.title)}</strong>. Tekan <strong>Ctrl+S</strong> untuk menyimpan versi baru.
         </div>`
      : `<div style="display:grid; gap:8px;">
          ${vaultHistory.map((ver, idx) => {
            const when = ver.savedAt ? new Date(ver.savedAt).toLocaleString('id-ID', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '-';
            return `
              <div style="display:flex; justify-content:space-between; align-items:center; gap:10px; padding:10px 12px; border:1px solid ${idx === 0 ? '#86efac' : '#e2e8f0'}; background:${idx === 0 ? '#f0fdf4' : '#ffffff'}; border-radius:10px;">
                <div style="min-width:0;">
                  <div style="font-size:12.5px; font-weight:700; color:#0f172a;">
                    ${idx === 0 ? '✓ Versi Aktif Terbaru · ' : `Versi #${vaultHistory.length - idx} · `}${escapeHtmlText(ver.note || 'Simpan Perubahan')}
                  </div>
                  <div style="font-size:11.5px; color:#64748b; margin-top:2px;">
                    🕒 ${escapeHtmlText(when)} · 👤 ${escapeHtmlText(ver.savedBy || 'Pengguna')} · No: ${escapeHtmlText(ver.nomorSurat || '-')}
                  </div>
                </div>
                ${idx > 0 && ver.document ? `<button type="button" class="btn btn-sm btn-secondary btn-smart-restore-ver" data-ver-id="${escapeHtmlText(ver.versionId)}">Pulihkan Versi</button>` : '<span class="status-chip status-chip-green" style="font-size:10px;">Terkini</span>'}
              </div>
            `;
          }).join('')}
        </div>`;

    modalRoot.innerHTML = `
      <div class="modal-overlay">
        <div class="modal-card" style="max-width: 680px; width: 95%; max-height: 88vh; display: flex; flex-direction: column;">
          <div class="modal-header">
            <h3>🕒 Riwayat Versi Dokumen${activeEntry ? `: ${escapeHtmlText(activeEntry.title)}` : ''}</h3>
            <button type="button" class="modal-close" id="smart-modal-close">&times;</button>
          </div>
          <div class="modal-body" style="overflow-y: auto; display: grid; gap: 14px;">
            <div>
              <h4 style="margin:0 0 8px 0; font-size:13px; color:#0f172a;">Riwayat Tersimpan di Gudang Surat (Maks 10 Versi)</h4>
              ${vaultRowsHtml}
            </div>
            <div style="padding-top:10px; border-top:1px solid #e2e8f0; display:flex; justify-content:space-between; align-items:center; gap:10px;">
              <div style="font-size:12px; color:#475569;">
                <strong>Undo Sesi Aktif:</strong> ${sessionUndoCount} langkah perubahan lokal tersedia di memori kanvas.
              </div>
              <button type="button" class="btn btn-sm btn-secondary" id="smart-modal-step-undo" ${sessionUndoCount === 0 ? 'disabled' : ''}>↶ Undo 1 Langkah</button>
            </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-primary" id="smart-modal-close-btn">Tutup</button>
          </div>
        </div>
      </div>
    `;
    modalRoot.querySelector('#smart-modal-close').onclick = closeSmartModal;
    modalRoot.querySelector('#smart-modal-close-btn').onclick = closeSmartModal;
    modalRoot.querySelector('#smart-modal-step-undo')?.addEventListener('click', () => {
      closeSmartModal();
      container.querySelector('#smart-btn-undo')?.click();
    });
    modalRoot.querySelectorAll('.btn-smart-restore-ver').forEach((btn) => {
      btn.addEventListener('click', () => {
        if (!activeEntry) return;
        const restored = store.restoreLetterVaultVersion(activeEntry.id, btn.dataset.verId);
        if (!restored || !restored.document) {
          showToast('Gagal memulihkan versi dokumen.', 'error');
          return;
        }
        pushUndoSnapshot();
        const converted = docModelToSmartPages(restored.document, store.getSchoolProfile());
        smartDoc = { pages: converted.pages, meta: converted.meta };
        store.setState({ smartDocument: smartDoc }, false);
        renderAllSmartPages(smartDoc.pages);
        updateSmartVaultBadge();
        closeSmartModal();
        showToast('✓ Versi dokumen sebelumnya berhasil dipulihkan!', 'success');
      });
    });
  }

  container.querySelector('#smart-btn-version-history')?.addEventListener('click', openSmartVersionHistoryModal);

  // =========================================================================
  // FEEDBACK REVISI KEPALA SEKOLAH TERHUBUNG LANGSUNG KE AI SURGICAL
  // =========================================================================
  const kepsekRevBannerEl = container.querySelector('#smart-kepsek-revision-banner');

  function renderKepsekRevisionBanner() {
    const rev = store.getPendingKepsekRevision();
    if (!kepsekRevBannerEl) return;
    if (!rev) {
      kepsekRevBannerEl.classList.add('hidden');
      kepsekRevBannerEl.innerHTML = '';
      return;
    }

    kepsekRevBannerEl.classList.remove('hidden');
    kepsekRevBannerEl.innerHTML = `
      <div class="smart-rev-head">
        <div>
          <span class="status-chip status-chip-red" style="font-size:10.5px;">📝 Catatan Revisi Kepala Sekolah</span>
          <strong style="margin-left:6px; font-size:13px; color:#0f172a;">${escapeHtmlText(rev.title || 'Surat')} (${escapeHtmlText(rev.nomorSurat || '-')})</strong>
        </div>
        <button type="button" class="smart-tbl-btn" id="smart-rev-dismiss" title="Sembunyikan banner revisi">✕ Tutup</button>
      </div>
      <div class="smart-rev-reason">"${escapeHtmlText(rev.reason || 'Perlu penyesuaian data sebelum ACC QR.')}"</div>
      <div class="smart-rev-actions">
        <button type="button" class="btn btn-sm btn-primary" id="smart-rev-auto-ai">🤖 Perbaiki Otomatis Sesuai Catatan (AI Surgical)</button>
        <button type="button" class="btn btn-sm btn-success" id="smart-rev-resubmit-acc">📨 Kirim Ajukan Ulang ACC Ini</button>
      </div>
    `;

    kepsekRevBannerEl.querySelector('#smart-rev-dismiss').onclick = () => {
      store.clearPendingKepsekRevision();
      renderKepsekRevisionBanner();
    };

    const runKepsekAiFix = () => {
      const promptTxt = `Lakukan perbaikan bedah (surgical) pada surat terbuka sesuai catatan revisi Kepala Sekolah berikut TANPA merombak tata letak atau struktur tabel: "${rev.reason}".`;
      return runSmartAiEdit(promptTxt);
    };

    kepsekRevBannerEl.querySelector('#smart-rev-auto-ai').onclick = runKepsekAiFix;

    kepsekRevBannerEl.querySelector('#smart-rev-resubmit-acc').onclick = async (e) => {
      const btn = e.currentTarget;
      btn.disabled = true;
      btn.textContent = 'Mengirim ACC Ulang...';
      try {
        const livePages = collectLivePagesData();
        const activeModel = getActiveSmartDocModel();
        const snap = extractSmartSnapshotFromPages(livePages, store.getSchoolProfile());
        const summaryText = (snap.blocks || []).filter(b => b.type === 'paragraph' && b.text).map(b => b.text).join(' ').slice(0, 160);
        const res = await requestAccCloud({
          outgoingId: rev.outgoingId,
          signerRole: rev.signerRole || 'kepala_sekolah',
          requestNote: `Sudah direvisi sesuai catatan Kepsek: ${rev.reason}`.slice(0, 200),
          title: snap.letterTitle || rev.title,
          recipientName: snap.tujuan || undefined,
          perihal: snap.hal || undefined,
          tanggalSurat: snap.signer?.date || undefined,
          summary: summaryText || undefined,
          documentSnapshot: { ...activeModel, pages: livePages }
        });
        if (!res.ok || !res.item) {
          throw new Error(res.error || 'Gagal mengajukan ulang ACC ke server.');
        }
        store.upsertOutgoingLetterFromCloud(res.item);
        if (res.syncRev) setLastKnownSyncRev(res.syncRev);
        store.clearPendingKepsekRevision();
        renderKepsekRevisionBanner();
        showToast('📨 Surat hasil revisi berhasil diajukan ulang ke Antrean ACC Kepala Sekolah!', 'success');
      } catch (err) {
        btn.disabled = false;
        btn.textContent = '📨 Kirim Ajukan Ulang ACC Ini';
        showToast(err.message, 'error');
      }
    };

    if (rev.autoRunAi) {
      store.setPendingKepsekRevision({ ...rev, autoRunAi: false });
      setTimeout(() => runKepsekAiFix(), 180);
    }
  }
  renderKepsekRevisionBanner();

  // Tombol Refresh / Bersihkan Lembar Kerja Word Pro
  container.querySelector('#smart-btn-refresh')?.addEventListener('click', openSmartRefreshConfirmModal);

  // Modal Catat Surat Keluar Resmi (berbagi nomor urut buku agenda!)
  container.querySelector('#smart-btn-record-outgoing')?.addEventListener('click', () => openSmartRecordOutgoingModal({ clearAfterSave: false }));

  // Shortcut Global Word Pro (Ctrl+S = Save, Ctrl+Shift+S / F12 = Save As, Ctrl+F / Ctrl+H = Cari & Ganti)
  const smartWorkspaceRoot = container.querySelector('.smart-word-workspace');
  const onSmartGlobalKeyDown = (e) => {
    if (!smartWorkspaceRoot || !smartWorkspaceRoot.isConnected) {
      document.removeEventListener('keydown', onSmartGlobalKeyDown);
      clearFindHighlights();
      return;
    }
    const key = e.key.toLowerCase();
    const isSaveAs = ((e.ctrlKey || e.metaKey) && e.shiftKey && key === 's') || e.key === 'F12';
    const isSave = (e.ctrlKey || e.metaKey) && !e.shiftKey && key === 's';
    const isFind = (e.ctrlKey || e.metaKey) && (key === 'f' || key === 'h');
    if (isSaveAs) {
      e.preventDefault();
      openSmartSaveVaultModal({ mode: 'save-as', clearAfterSave: false });
    } else if (isSave) {
      e.preventDefault();
      performSmartSave({ clearAfterSave: false });
    } else if (isFind) {
      e.preventDefault();
      openSmartFindReplaceBar({ focusReplace: key === 'h' });
    }
  };
  document.addEventListener('keydown', onSmartGlobalKeyDown);
}
