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

import { store } from '../state.js';
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
import { generateLetterNumber, formatSmartNomorSurat, detectLetterTypeCode } from '../letters/numbering.js';
import { callAiWithValidation, polishMessyDocumentWithAi, editDocumentWithAi } from '../ai/client.js';
import { COPILOT_SYSTEM_PROMPT, buildCopilotUserPrompt } from '../ai/prompts.js';

// =========================================================================
// BAGIAN 1: PARSER HTML RINGAN & KONVERTER HTML <-> DOCUMENT MODEL (PURE JS)
// =========================================================================

const VOID_TAGS = new Set(['br', 'hr', 'img', 'input', 'meta', 'link']);

export function escapeHtmlText(str = '') {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export function decodeHtmlEntities(str = '') {
  return String(str)
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&#(\d+);/g, (_, dec) => String.fromCharCode(Number(dec)))
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
}

function parseAttributes(attrStr = '') {
  const attrs = {};
  const re = /([a-zA-Z0-9_:-]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+)))?/g;
  let m;
  while ((m = re.exec(attrStr)) !== null) {
    const key = m[1].toLowerCase();
    const val = m[2] ?? m[3] ?? m[4] ?? '';
    attrs[key] = decodeHtmlEntities(val);
  }
  return attrs;
}

/**
 * Parser HTML zero-dependency (berjalan identik di Browser maupun Node.js unit test).
 * Menghasilkan daftar node AST: { type: 'element'|'text', tag, attrs, children, text }.
 */
export function parseHtmlNodes(html = '') {
  const root = { type: 'element', tag: 'root', attrs: {}, children: [] };
  const stack = [root];
  const tokenRe = /<!--[\s\S]*?-->|<\/?([a-zA-Z][a-zA-Z0-9]*)([^>]*)>|([^<]+)/g;
  let match;

  while ((match = tokenRe.exec(String(html))) !== null) {
    if (match[3] !== undefined) {
      const txt = decodeHtmlEntities(match[3]);
      if (txt) {
        stack[stack.length - 1].children.push({ type: 'text', text: txt });
      }
      continue;
    }
    if (!match[1]) continue;

    const fullTag = match[0];
    const tag = match[1].toLowerCase();
    const rawAttrs = match[2] || '';
    const isClosing = fullTag.startsWith('</');
    const isSelfClosing = fullTag.endsWith('/>') || VOID_TAGS.has(tag);

    if (isClosing) {
      for (let i = stack.length - 1; i >= 1; i--) {
        if (stack[i].tag === tag) {
          stack.length = i;
          break;
        }
      }
    } else {
      const node = {
        type: 'element',
        tag,
        attrs: parseAttributes(rawAttrs),
        children: []
      };
      stack[stack.length - 1].children.push(node);
      if (!isSelfClosing) {
        stack.push(node);
      }
    }
  }

  return root.children;
}

function getNodePlainText(node) {
  if (!node) return '';
  if (node.type === 'text') return node.text;
  if (node.tag === 'br') return '\n';
  return (node.children || []).map(getNodePlainText).join('');
}

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

function extractInlineRuns(nodes, inherited = {}) {
  const runs = [];

  for (const node of nodes || []) {
    if (node.type === 'text') {
      const clean = node.text.replace(/\s+/g, ' ');
      if (clean) {
        runs.push({
          text: clean,
          ...(inherited.bold ? { bold: true } : {}),
          ...(inherited.italic ? { italic: true } : {}),
          ...(inherited.underline ? { underline: true } : {}),
          ...(inherited.sup ? { sup: true } : {}),
          ...(inherited.sub ? { sub: true } : {})
        });
      }
      continue;
    }

    if (node.type !== 'element') continue;
    if (node.tag === 'br') {
      runs.push({ text: ' ', ...inherited });
      continue;
    }

    const style = (node.attrs?.style || '').toLowerCase();
    const nextFmt = {
      bold: inherited.bold || node.tag === 'b' || node.tag === 'strong' || /font-weight\s*:\s*(bold|700|800|900)/.test(style),
      italic: inherited.italic || node.tag === 'i' || node.tag === 'em' || /font-style\s*:\s*italic/.test(style),
      underline: inherited.underline || node.tag === 'u' || /text-decoration[^;]*underline/.test(style),
      sup: inherited.sup || node.tag === 'sup',
      sub: inherited.sub || node.tag === 'sub'
    };

    const inner = extractInlineRuns(node.children, nextFmt);
    runs.push(...inner);
  }

  // Merge adjacent runs with identical formatting and trim leading/trailing overall whitespace cleanly
  const merged = [];
  for (const r of runs) {
    const prev = merged[merged.length - 1];
    if (
      prev &&
      Boolean(prev.bold) === Boolean(r.bold) &&
      Boolean(prev.italic) === Boolean(r.italic) &&
      Boolean(prev.underline) === Boolean(r.underline) &&
      Boolean(prev.sup) === Boolean(r.sup) &&
      Boolean(prev.sub) === Boolean(r.sub)
    ) {
      prev.text += r.text;
    } else {
      merged.push({ ...r });
    }
  }
  return merged;
}

function collectTableRows(tableNode) {
  const rows = [];
  const walk = (n) => {
    if (!n || n.type !== 'element') return;
    if (n.tag === 'tr') {
      const cells = (n.children || []).filter(c => c.type === 'element' && (c.tag === 'td' || c.tag === 'th'));
      rows.push({
        isHeader: cells.length > 0 && cells.every(c => c.tag === 'th'),
        cells: cells.map(c => getNodePlainText(c).replace(/\s+/g, ' ').trim())
      });
      return;
    }
    (n.children || []).forEach(walk);
  };
  walk(tableNode);
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

/**
 * Mengonversi HTML satu halaman Smart Canvas menjadi daftar blok DocumentModel terstruktur.
 */
export function parseSmartPageHtmlToBlocks(html = '', schoolProfile = {}) {
  const nodes = parseHtmlNodes(html).filter(n => !(n.type === 'text' && !n.text.trim()));
  const blocks = [];

  for (let idx = 0; idx < nodes.length; idx++) {
    const node = nodes[idx];

    if (node.type === 'text') {
      const txt = node.text.trim();
      if (txt) {
        blocks.push({ type: 'paragraph', style: 'letterBody', runs: [{ text: txt }] });
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
        blocks.push({
          type: 'identityTable',
          rows: rows.map(r => ({
            label: r.cells[0] || '',
            value: r.cells[1] !== undefined ? r.cells[1] : ''
          }))
        });
        continue;
      }

      // Word custom grid table
      blocks.push({
        type: 'table',
        headerRow: rows[0]?.isHeader || false,
        rows: rows.map(r => r.cells)
      });
      continue;
    }

    // 3. List (<ul> / <ol>)
    if (node.tag === 'ul' || node.tag === 'ol') {
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
          ordered: node.tag === 'ol',
          items
        });
      }
      continue;
    }

    // 4. Headings (<h1>, <h2>, <h3>)
    if (['h1', 'h2', 'h3', 'h4'].includes(node.tag)) {
      const txt = getNodePlainText(node).replace(/\s+/g, ' ').trim();
      if (!txt) continue;
      // Render heading sebagai paragraf tebal rapi untuk surat dinas agar tidak memicu penomoran BAB otomatis di Word
      const align = detectCssAlignment(node) || 'left';
      blocks.push({
        type: 'paragraph',
        style: 'letterBody',
        alignment: align,
        runs: [{ text: txt, bold: true }]
      });
      continue;
    }

    // 5. Paragraph / Div standar MS Word
    if (node.tag === 'p' || node.tag === 'div') {
      const rawTxt = getNodePlainText(node).replace(/\u00a0/g, ' ').trim();
      if (!rawTxt) {
        // Baris kosong tetap dipertahankan sebagai spasi paragraf
        blocks.push({ type: 'paragraph', style: 'letterBody', runs: [{ text: '' }] });
        continue;
      }
      const runs = extractInlineRuns(node.children);
      const alignment = detectCssAlignment(node);
      blocks.push({
        type: 'paragraph',
        style: 'letterBody',
        ...(alignment ? { alignment } : {}),
        runs: runs.length > 0 ? runs : [{ text: rawTxt }]
      });
    }
  }

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
  let name = schoolProfile.principalName || 'Dian Nurhadian, S.Pd.';
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

  return {
    meta: {
      title: extractedTitle,
      author: schoolProfile.schoolName || 'SMKS IT Al-Hawari',
      year: new Date().getFullYear().toString()
    },
    profile: profileName,
    paperSize: letterOptions.paperSize || 'a4',
    orientation: letterOptions.orientation || 'portrait',
    customMargins: letterOptions.margins || { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 },
    letterData: {
      templateId: 'smart_word_pro',
      schoolProfile,
      formData: {
        judulSurat: extractedTitle,
        nomorSurat: extractedNomor
      },
      options: letterOptions
    },
    blocks
  };
}

function runsToHtml(runs = []) {
  if (!Array.isArray(runs) || runs.length === 0) return '';
  return runs.map((r) => {
    let html = escapeHtmlText(r.text || '');
    if (r.bold) html = `<strong>${html}</strong>`;
    if (r.italic) html = `<em>${html}</em>`;
    if (r.underline) html = `<u>${html}</u>`;
    if (r.sup) html = `<sup>${html}</sup>`;
    if (r.sub) html = `<sub>${html}</sub>`;
    return html;
  }).join('');
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
      const rowsHtml = (b.rows || []).map(r => {
        const val = (r.value || '').trim();
        const normVal = val.startsWith(':') ? val : (val ? `: ${val}` : ': ');
        return `<tr><td style="width: 30%;">${escapeHtmlText(r.label || '')}</td><td>${escapeHtmlText(normVal)}</td></tr>`;
      }).join('');
      currentParts.push(`<table class="smart-identity-table" data-smart-role="identity-table"><tbody>${rowsHtml}</tbody></table>`);
      continue;
    }

    if (b.type === 'table') {
      const rows = b.rows || [];
      const hasHeader = b.headerRow !== false && rows.length > 0;
      const headHtml = hasHeader
        ? `<thead><tr>${(rows[0] || []).map(c => `<th>${escapeHtmlText(typeof c === 'object' ? c.text : c)}</th>`).join('')}</tr></thead>`
        : '';
      const bodyRows = hasHeader ? rows.slice(1) : rows;
      const bodyHtml = `<tbody>${bodyRows.map(row => `<tr>${(row || []).map(c => `<td>${escapeHtmlText(typeof c === 'object' ? c.text : c)}</td>`).join('')}</tr>`).join('')}</tbody>`;
      currentParts.push(`<table class="smart-word-table">${headHtml}${bodyHtml}</table>`);
      continue;
    }

    if (b.type === 'list') {
      const tag = b.ordered ? 'ol' : 'ul';
      const liHtml = (b.items || []).map(item => {
        if (typeof item === 'string') return `<li>${escapeHtmlText(item)}</li>`;
        if (Array.isArray(item)) return `<li>${runsToHtml(item)}</li>`;
        return `<li>${escapeHtmlText(item?.text || '')}</li>`;
      }).join('');
      currentParts.push(`<${tag}>${liHtml}</${tag}>`);
      continue;
    }

    if (b.type === 'heading') {
      const lvl = Math.min(Math.max(Number(b.level) || 2, 1), 3);
      const inner = b.runs ? runsToHtml(b.runs) : escapeHtmlText(b.text || '');
      currentParts.push(`<h${lvl}>${inner}</h${lvl}>`);
      continue;
    }

    if (b.type === 'letterSignature') {
      currentParts.push(renderSignatureTableHtml(b, schoolProfile));
      continue;
    }

    if (b.type === 'paragraph') {
      const inner = b.runs ? runsToHtml(b.runs) : escapeHtmlText(b.text || '');
      const alignCss = b.alignment
        ? ` style="text-align: ${b.alignment === 'both' ? 'justify' : b.alignment};"`
        : '';
      currentParts.push(`<p${alignCss}>${inner || '<br>'}</p>`);
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
  const name = sig.name || school.principalName || 'Dian Nurhadian, S.Pd.';
  const nipRaw = sig.nip && sig.nip !== '-' ? sig.nip : (school.principalNip || '-');
  const nipLine = nipRaw && nipRaw !== '-' ? `<br>NIP. ${escapeHtmlText(nipRaw)}` : '';

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
    name: schoolProfile.principalName || 'Dian Nurhadian, S.Pd.',
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
    } else if (b.type === 'paragraph') {
      const txt = (b.runs || []).map(r => r.text).join('').trim();
      if (txt) bodyBlocks.push({ type: 'paragraph', text: txt });
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
  const title = (parsedJson.letterTitle || parsedJson.title || 'SURAT DINAS SEKOLAH').trim();
  const contextText = `${title} ${parsedJson.hal || parsedJson.perihal || ''}`;
  const rawNomor = (parsedJson.nomorSurat || '').trim();
  const nomorFinal = rawNomor
    ? formatSmartNomorSurat({
        rawNomor,
        contextText,
        forcedTypeCode: parsedJson.kodeJenisSurat || ''
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

  blocks.push({
    type: 'letterTitle',
    title: title.toUpperCase(),
    nomor: `Nomor: ${nomorFinal}`,
    ...(metaRecipient ? { metaRecipient } : {})
  });

  if (Array.isArray(parsedJson.blocks)) {
    for (const b of parsedJson.blocks) {
      if (b.type === 'paragraph' && b.text) {
        blocks.push({
          type: 'paragraph',
          style: 'letterBody',
          runs: [{ text: b.text.trim() }]
        });
      } else if (b.type === 'identityTable' && Array.isArray(b.rows)) {
        blocks.push({
          type: 'identityTable',
          rows: b.rows.map(r => ({
            label: r.label || '',
            value: (r.value || '').startsWith(':') ? r.value : `: ${r.value || ''}`
          }))
        });
      }
    }
  }

  const sig = parsedJson.signer || {};
  blocks.push({
    type: 'letterSignature',
    city: sig.city || schoolProfile.city || 'Garut',
    date: sig.date || parsedJson.tanggalSurat || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
    title: sig.title || schoolProfile.principalTitle || 'Kepala Sekolah',
    name: sig.name || schoolProfile.principalName || 'Dian Nurhadian, S.Pd.',
    nip: sig.nip && sig.nip !== '-' ? sig.nip : (schoolProfile.principalNip || '-')
  });

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

  return {
    meta: {
      title,
      author: schoolProfile.schoolName || 'SMKS IT Al-Hawari',
      year: new Date().getFullYear().toString()
    },
    profile: letterOptions.paperSize === 'f4' ? 'suratDinasF4' : 'suratDinas',
    paperSize: letterOptions.paperSize || 'a4',
    orientation: letterOptions.orientation || 'portrait',
    customMargins: letterOptions.margins || { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 },
    letterData: {
      templateId: 'smart_word_ai',
      schoolProfile,
      formData: {
        judulSurat: title,
        nomorSurat: nomorFinal
      },
      options: letterOptions
    },
    blocks
  };
}

// =========================================================================
// BAGIAN 2: BRIDGE GETTER UNTUK HEADER DOWNLOAD GLOBAL & STATE HELPERS
// =========================================================================

let activeLivePagesExtractor = null;

/**
 * Mengambil DocumentModel terkini dari Kanvas Word Pro (dipanggil header Unduh .docx atau Simpan).
 */
export function getActiveSmartDocModel() {
  const state = store.getState();
  const school = state.schoolProfile;
  const opts = state.letterOptions || {};
  const pages = activeLivePagesExtractor
    ? activeLivePagesExtractor()
    : (state.smartDocument?.pages || docModelToSmartPages(state.document, school).pages);

  return buildDocModelFromSmartPagesData({
    pages,
    schoolProfile: school,
    letterOptions: opts,
    meta: state.smartDocument?.meta || {}
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
  const opts = state.letterOptions || {};
  const paperSize = opts.paperSize || 'a4';
  const orientation = opts.orientation || 'portrait';
  const customMargins = opts.margins || { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 };
  const fontFamily = opts.fontFamily || 'Times New Roman';
  const fontSize = opts.fontSize || '12';
  const lineSpacing = opts.lineSpacing || '1.15';
  const showKop = opts.showKop !== false;

  // Default halaman Word Pro dimulai dalam keadaan KOSONG bersih (tanpa teks contoh bawaan)
  let smartDoc = state.smartDocument;
  if (!smartDoc || !Array.isArray(smartDoc.pages) || smartDoc.pages.length === 0) {
    smartDoc = createEmptySmartDocument();
    store.setState({ smartDocument: smartDoc }, false);
  }

  let currentZoom = 1.0;
  let copilotVisible = true;
  let showFormattingMarks = false;
  let activePageIndex = 0;
  let lastSavedRange = null;
  const smartUndoStack = [];
  const smartRedoStack = [];

  const apiVault = state.apiVault || [];
  let selectedApiId = apiVault.find(a => a.isActive)?.id || apiVault[0]?.id || '';
  const kbCount = (state.letterVault || []).length;

  const kopUrl = school.kopImage || DEFAULT_KOP_DATA_URL;
  const marginStyles = `
    --margin-top: ${customMargins.top}cm;
    --margin-bottom: ${customMargins.bottom}cm;
    --margin-left: ${customMargins.left}cm;
    --margin-right: ${customMargins.right}cm;
    --doc-font: '${fontFamily}', serif;
    --doc-font-size: ${fontSize}pt;
    --doc-line-height: ${lineSpacing};
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
                  <button class="smart-popover-item" data-margin="2,2,2.5,2">Standar Dinas <span>2 / 2,5 cm</span></button>
                  <button class="smart-popover-item" data-margin="2.54,2.54,2.54,2.54">Normal Word <span>2,54 cm</span></button>
                  <button class="smart-popover-item" data-margin="1.27,1.27,1.27,1.27">Sempit (Narrow) <span>1,27 cm</span></button>
                  <button class="smart-popover-item" data-margin="3,3,3,3">Lebar (Wide) <span>3,0 cm</span></button>
                </div>
              </div>
              <button class="ribbon-btn ribbon-btn-compact" id="smart-btn-format-modal" title="Pengaturan Kertas & Margin Kustom">
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
                <option value="Times New Roman" ${fontFamily === 'Times New Roman' ? 'selected' : ''}>Times New Roman</option>
                <option value="Arial" ${fontFamily === 'Arial' ? 'selected' : ''}>Arial</option>
                <option value="Calibri" ${fontFamily === 'Calibri' ? 'selected' : ''}>Calibri</option>
                <option value="Georgia" ${fontFamily === 'Georgia' ? 'selected' : ''}>Georgia</option>
                <option value="Garamond" ${fontFamily === 'Garamond' ? 'selected' : ''}>Garamond</option>
              </select>
              <select class="ribbon-select" id="smart-font-size" title="Ukuran Huruf" style="width: 50px;">
                ${['9', '10', '11', '12', '13', '14', '16', '18'].map(sz =>
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
              <button class="ribbon-btn-icon" data-cmd="insertOrderedList" title="Daftar Bernomor">${ICON_LIST_NUMBER}</button>
              <button class="ribbon-btn-icon" data-cmd="outdent" title="Kurangi Indentasi">${ICON_OUTDENT}</button>
              <button class="ribbon-btn-icon" data-cmd="indent" title="Tambah Indentasi">${ICON_INDENT}</button>
              <select class="ribbon-select" id="smart-line-spacing" title="Spasi Baris">
                <option value="1.0" ${String(lineSpacing) === '1.0' ? 'selected' : ''}>Spasi 1,0</option>
                <option value="1.15" ${String(lineSpacing) === '1.15' ? 'selected' : ''}>Spasi 1,15</option>
                <option value="1.5" ${String(lineSpacing) === '1.5' ? 'selected' : ''}>Spasi 1,5</option>
                <option value="2.0" ${String(lineSpacing) === '2.0' ? 'selected' : ''}>Spasi 2,0</option>
              </select>
            </div>
          </div>
          <div class="ribbon-label">Paragraf</div>
        </div>

        <!-- Group 4: Sisipkan -->
        <div class="ribbon-group">
          <div class="ribbon-tiles-row">
            <button class="ribbon-tile" id="smart-btn-add-page" title="Tambah Halaman Baru (Ctrl+Enter)">
              ${ICON_PAGE_ADD}
              <span>+ Halaman</span>
            </button>
            <button class="ribbon-tile" id="smart-btn-insert-table" title="Sisipkan Tabel Grid / Tabel Data">
              ${ICON_TABLE}
              <span>Tabel</span>
            </button>
            <button class="ribbon-tile" id="smart-btn-insert-sig" title="Sisipkan Blok Tanda Tangan">
              ${ICON_PENCIL_NOTE}
              <span>TTD</span>
            </button>
            <button class="ribbon-tile" id="smart-btn-kop-gallery" title="Pilih / Kelola Kop Surat">
              ${ICON_GALLERY}
              <span>Kop Surat</span>
            </button>
            <button class="ribbon-tile ${showKop ? 'active' : ''}" id="smart-btn-toggle-kop" title="Tampilkan / Sembunyikan Kop Surat">
              ${showKop ? ICON_EYE : ICON_EYE_OFF}
              <span>Tampil Kop</span>
            </button>
          </div>
          <div class="ribbon-label">Sisipkan</div>
        </div>

        <!-- Group 5: Dokumen -->
        <div class="ribbon-group">
          <div class="ribbon-tiles-row">
            <button class="ribbon-tile" id="smart-btn-save-vault" title="Simpan Dokumen ke Gudang Surat">
              ${ICON_SAVE}
              <span>Simpan</span>
            </button>
            <button class="ribbon-tile" id="smart-btn-refresh" title="Refresh / Kosongkan Kertas ke Lembar Baru">
              ${ICON_REFRESH}
              <span>Refresh</span>
            </button>
            <button class="ribbon-tile" id="smart-btn-undo" title="Urungkan Perubahan">${ICON_UNDO}<span>Undo</span></button>
            <button class="ribbon-tile" id="smart-btn-redo" title="Ulangi Perubahan">${ICON_REDO}<span>Redo</span></button>
            <button class="ribbon-tile" id="smart-btn-pull-canvas1" title="Salin Isi Surat dari Lembar Kerja Lama ke Mode Word Pro">
              ${ICON_COPY_DOC}
              <span>Salin Surat 1</span>
            </button>
            <button class="ribbon-tile" id="smart-btn-import-docx" title="Impor Berkas .docx ke Kanvas Word Pro">
              ${ICON_DOC_UPLOAD}
              <span>Impor .docx</span>
            </button>
            <button class="ribbon-tile active" id="smart-btn-toggle-copilot" title="Tampilkan / Sembunyikan AI Copilot">
              ${ICON_BOT_SPARK}
              <span>Copilot AI</span>
            </button>
          </div>
          <div class="ribbon-label">Dokumen</div>
        </div>
        <input type="file" id="smart-docx-upload-input" accept=".docx" style="display:none" />
        <input type="file" id="smart-kop-upload-input" accept="image/png,image/jpeg,image/webp,image/svg+xml" style="display:none" />
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

              <!-- TUMPUKAN HALAMAN KERTAS MS WORD BEBAS & BERKELANJUTAN -->
              <div class="smart-pages-stack" id="smart-pages-stack" style="${marginStyles}"></div>
            </div>
          </div>
        </div>

        <!-- SIDEBAR AI COPILOT KANAN (DOCKED) -->
        <aside class="copilot-drawer no-print" id="smart-copilot-drawer">
          <div class="copilot-header">
            <div class="copilot-title">
              ${ICON_BOT_SPARK}
              <h4>AI Copilot Word Pro</h4>
            </div>
            <button class="copilot-close-btn" id="smart-copilot-close" title="Tutup Panel Copilot">✕</button>
          </div>

          <div class="copilot-body">
            <div class="copilot-model-card">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px; gap:6px;">
                <span style="font-size:11.5px; font-weight:700; color:#1e293b;">Model AI dari Gudang API</span>
                <span class="smart-badge-pro">Word Multi-Page</span>
              </div>
              <select class="form-control copilot-model-select" id="smart-copilot-model-select" style="width:100%;">
                ${apiVault.map(api =>
                  `<option value="${api.id}" ${api.id === selectedApiId ? 'selected' : ''}>[${(api.provider || 'AI').toUpperCase()}] ${escapeHtmlText(api.label || api.name || api.model)}</option>`
                ).join('')}
              </select>
            </div>

            <div class="copilot-kb-card">
              <div style="display:flex; align-items:flex-start; gap:7px; min-width:0;">
                <span style="color:#15803d; display:inline-flex; margin-top:1px; flex-shrink:0;">${ICON_OPEN_BOOK}</span>
                <span><strong>Basis Pengetahuan:</strong> ${kbCount} arsip surat aktif sebagai acuan gaya</span>
              </div>
              <button type="button" class="copilot-kb-btn" id="smart-btn-open-vault-nav">Kelola Gudang</button>
            </div>

            <div class="copilot-desc">
              Ketik surat dinas bebas atau dokumen panjang (skrip rapat, tata tertib, lampiran multi-halaman). Konten otomatis mengalir seperti Microsoft Word asli tanpa kotak per paragraf.
            </div>

            <div class="copilot-chips">
              <button type="button" class="copilot-chip" data-prompt="Buatkan Surat Undangan Rapat Dewan Guru SMKS IT Al-Hawari beserta lampiran susunan acara dan daftar hadir pada halaman 2">${ICON_MAIL} Undangan + Lampiran</button>
              <button type="button" class="copilot-chip" data-prompt="Buatkan Surat Edaran Tata Usaha pemberitahuan pembagian rapor semester ganjil untuk orang tua/wali">${ICON_BUILDING} Surat Tata Usaha (TU)</button>
              <button type="button" class="copilot-chip" data-prompt="Buatkan Surat Keputusan (SK) Kepala Sekolah tentang Pembagian Tugas Guru Mengajar beserta rincian lengkap pada halaman ke-2">${ICON_BRIEFCASE} SK Multi-Halaman (KEP)</button>
              <button type="button" class="copilot-chip" data-prompt="Buatkan Surat Keterangan Siswa Aktif atas nama [Nama Siswa] NIS [NIS] kelas [Kelas]">${ICON_ID_CARD} Ket. Siswa Aktif (KET)</button>
              <button type="button" class="copilot-chip" data-prompt="Buatkan Surat Tugas Guru mengikuti Workshop Kurikulum Merdeka beserta rincian jadwal kegiatan">${ICON_TARGET} Surat Tugas (ST)</button>
              <button type="button" class="copilot-chip" data-prompt="Buatkan Surat Panggilan Orang Tua/Wali Siswa terkait pembinaan kedisiplinan siswa">${ICON_USERS} Panggilan Ortu (PGL)</button>
            </div>

            <div class="form-group" style="margin-top: 4px; margin-bottom: 10px;">
              <label class="copilot-prompt-label" for="smart-copilot-prompt">${ICON_PENCIL_NOTE} <span>Perintah Pembuatan Baru atau Instruksi Edit Surat Terbuka</span></label>
              <textarea
                id="smart-copilot-prompt"
                class="form-control copilot-textarea"
                rows="4"
                spellcheck="false"
                placeholder="Ketik instruksi edit (mis: tambah jadwal 09.00 WIB, perbaiki kalimat...) ATAU perintah buat surat baru multi-halaman..."
              ></textarea>
            </div>

            <div style="display: flex; flex-direction: column; gap: 7px;">
              <button type="button" class="btn copilot-primary-cta" id="smart-copilot-btn-edit">
                ${ICON_SPARKLE} <span>Edit / Tambah pada Surat Terbuka</span>
              </button>
              <button type="button" class="btn copilot-secondary-cta" id="smart-copilot-btn-create">
                ${ICON_DOC_UPLOAD} <span>Buat Surat Baru dari Awal</span>
              </button>
            </div>

            <hr style="margin: 14px 0; border: 0; border-top: 1px solid #e2e8f0;" />
            <h5 class="copilot-quick-title">${ICON_SPARKLE} <span>Aksi Cepat AI Copilot:</span></h5>
            <div class="btn-group-vertical" style="display: flex; flex-direction: column; gap: 6px;">
              <button type="button" class="copilot-quick-btn copilot-quick-btn-polish" id="smart-copilot-quick-polish">
                <span style="color:#15803d; display:inline-flex;">${ICON_WAND}</span> <span>Rapihkan Teks Berantakan (AI)</span>
              </button>
              <button type="button" class="copilot-quick-btn" id="smart-copilot-quick-puebi">
                <span style="color:#475569; display:inline-flex;">${ICON_OPEN_BOOK}</span> <span>Rapikan Bahasa Standar PUEBI</span>
              </button>
              <button type="button" class="copilot-quick-btn" id="smart-copilot-quick-today">
                <span style="color:#475569; display:inline-flex;">${ICON_CALENDAR}</span> <span>Perbarui Tanggal Surat ke Hari Ini</span>
              </button>
            </div>

            <div id="smart-copilot-status" class="copilot-status-area hidden"></div>
          </div>
        </aside>
      </div>

      <!-- STATUS BAR BAWAH -->
      <div class="word-status-bar no-print">
        <div class="status-left">
          <span class="status-item" id="smart-status-pages">${ICON_PAGE_ADD} Halaman 1 dari 1</span>
          <span class="status-sep">|</span>
          <span class="status-item status-word-count" id="smart-status-words">0 kata</span>
          <span class="status-sep">|</span>
          <span class="status-item">${ICON_GLOBE} Bahasa Indonesia (PUEBI) &bull; Mode Bebas MS Word</span>
        </div>
        <div class="status-right">
          <div class="zoom-slider-wrap">
            <button class="zoom-btn" id="smart-zoom-out" title="Perkecil Tampilan">-</button>
            <input type="range" id="smart-zoom-slider" min="50" max="150" value="100" step="5" />
            <button class="zoom-btn" id="smart-zoom-in" title="Perbesar Tampilan">+</button>
          </div>
          <span class="status-zoom-pill" id="smart-zoom-label">100%</span>
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
  const modalRoot = container.querySelector('#smart-modal-root');

  let reflowTimer = null;
  let saveTimer = null;

  function collectLivePagesData() {
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
      smartDoc = { ...smartDoc, pages: livePages };
      store.setState({ smartDocument: smartDoc }, false);
      updateStatusCounters();
    }, 250);
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

  function applyZoom() {
    pagesStack.style.transform = `scale(${currentZoom})`;
    container.querySelector('#smart-zoom-label').textContent = `${Math.round(currentZoom * 100)}%`;
    container.querySelector('#smart-zoom-slider').value = String(Math.round(currentZoom * 100));
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

    const sheet = document.createElement('div');
    sheet.className = `smart-paper-sheet paper-${pSize} ${orient === 'landscape' ? 'landscape' : ''}`;
    sheet.style.cssText = `
      --margin-top: ${curOpts.margins?.top ?? 2.0}cm;
      --margin-bottom: ${curOpts.margins?.bottom ?? 2.0}cm;
      --margin-left: ${curOpts.margins?.left ?? 2.5}cm;
      --margin-right: ${curOpts.margins?.right ?? 2.0}cm;
      --doc-font: '${curOpts.fontFamily || 'Times New Roman'}', serif;
      --doc-font-size: ${curOpts.fontSize || '12'}pt;
      --doc-line-height: ${curOpts.lineSpacing || '1.15'};
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

    sheet.appendChild(editable);
    wrapper.appendChild(headerBar);
    wrapper.appendChild(sheet);
    return wrapper;
  }

  function renderAllSmartPages(pagesList) {
    pagesStack.innerHTML = '';
    const list = (Array.isArray(pagesList) && pagesList.length > 0)
      ? pagesList
      : [{ id: `spage-${Date.now()}-1`, html: '<p><br></p>', isManualBreak: false }];

    list.forEach((p, idx) => {
      pagesStack.appendChild(createPageDomElement(p, idx));
    });

    reindexPageHeaders();
    updateStatusCounters();
    // Jalankan pasca-render check untuk auto-overflow bila teks panjang
    requestAnimationFrame(() => reflowPagination());
  }

  function reindexPageHeaders() {
    const wrappers = [...pagesStack.querySelectorAll('.smart-page-wrapper')];
    const total = wrappers.length;
    wrappers.forEach((w, idx) => {
      const editable = w.querySelector('.smart-editable-page');
      if (editable) editable.dataset.pageIndex = String(idx);
      const pill = w.querySelector('.smart-page-pill');
      if (pill) pill.textContent = `Halaman ${idx + 1} dari ${total}`;
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
        return;
      }
      remaining -= len;
    }
    const sel = window.getSelection();
    const r = document.createRange();
    r.selectNodeContents(targetNode);
    r.collapse(false);
    sel.removeAllRanges();
    sel.addRange(r);
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
    }
  }

  function splitOversizedSingleBlock(curEl, nextEl, maxH) {
    const child = curEl.firstElementChild;
    if (!child) return false;

    if (child.tagName === 'UL' || child.tagName === 'OL') {
      if (child.children.length <= 1) return false;
      const tag = child.tagName.toLowerCase();
      let destList = nextEl.firstElementChild;
      if (!destList || destList.tagName !== child.tagName) {
        destList = document.createElement(tag);
        nextEl.insertBefore(destList, nextEl.firstChild);
      }
      const lastLi = child.lastElementChild;
      destList.insertBefore(lastLi, destList.firstChild);
      return true;
    }

    if (child.tagName === 'TABLE') {
      const tbody = child.querySelector('tbody') || child;
      const rows = [...tbody.querySelectorAll('tr')];
      if (rows.length <= 1) return false;
      let destTable = nextEl.firstElementChild;
      if (!destTable || destTable.tagName !== 'TABLE') {
        destTable = document.createElement('table');
        destTable.className = child.className;
        const destBody = document.createElement('tbody');
        destTable.appendChild(destBody);
        nextEl.insertBefore(destTable, nextEl.firstChild);
      }
      const destBody = destTable.querySelector('tbody') || destTable;
      destBody.insertBefore(rows[rows.length - 1], destBody.firstChild);
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
      smartDoc = { ...smartDoc, pages: livePages };
      store.setState({ smartDocument: smartDoc }, false);
    });

    editable.addEventListener('paste', (e) => {
      const html = e.clipboardData?.getData('text/html');
      const plain = e.clipboardData?.getData('text/plain');
      if (html) {
        e.preventDefault();
        pushUndoSnapshot();
        const sanitized = sanitizePastedWordHtml(html);
        document.execCommand('insertHTML', false, sanitized);
        scheduleReflowPagination();
        schedulePersistSmartDoc();
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
      'ul', 'ol', 'li', 'table', 'thead', 'tbody', 'tr', 'th', 'td', 'span'
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
      const el = document.createElement(tag);
      const align = node.style?.textAlign;
      if (align && ['left', 'center', 'right', 'justify'].includes(align)) {
        el.style.textAlign = align;
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

  // Block Style (Normal / H1 / H2 / H3)
  container.querySelector('#smart-block-style')?.addEventListener('change', (e) => {
    execSmartCommand('formatBlock', `<${e.target.value}>`);
  });

  // Font & Typography Options
  container.querySelector('#smart-paper-select')?.addEventListener('change', (e) => {
    store.updateLetterOptions({ paperSize: e.target.value });
    renderAllSmartPages(collectLivePagesData());
    showToast(`Ukuran kertas diatur ke ${PAPER_LABELS[e.target.value] || e.target.value.toUpperCase()}`, 'info');
  });

  container.querySelector('#smart-orientation-select')?.addEventListener('change', (e) => {
    store.updateLetterOptions({ orientation: e.target.value });
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
      marginPop.classList.add('hidden');
      renderAllSmartPages(collectLivePagesData());
      showToast('Margin halaman berhasil diterapkan.', 'success');
    });
  });
  document.addEventListener('click', () => marginPop?.classList.add('hidden'), { once: false });

  // + Halaman (Page Break)
  container.querySelector('#smart-btn-add-page')?.addEventListener('mousedown', e => e.preventDefault());
  container.querySelector('#smart-btn-add-page')?.addEventListener('click', () => insertManualPageBreakAtSelection());

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
    renderAllSmartPages(smartDoc.pages);
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
      renderAllSmartPages(smartDoc.pages);
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
    applyZoom();
  });
  container.querySelector('#smart-zoom-in')?.addEventListener('click', () => {
    currentZoom = Math.min(1.5, Number((currentZoom + 0.1).toFixed(2)));
    applyZoom();
  });
  container.querySelector('#smart-zoom-slider')?.addEventListener('input', (e) => {
    currentZoom = Number(e.target.value) / 100;
    applyZoom();
  });
  container.querySelector('#smart-zoom-fit')?.addEventListener('click', () => {
    const desk = container.querySelector('#smart-word-desk');
    const sheet = pagesStack.querySelector('.smart-paper-sheet');
    if (desk && sheet) {
      const ratio = (desk.clientWidth - 48) / sheet.offsetWidth;
      currentZoom = Math.max(0.5, Math.min(1.2, Number(ratio.toFixed(2))));
      applyZoom();
    }
  });

  // =========================================================================
  // LOGIKA AI EDIT & AI COPILOT WORD PRO
  // =========================================================================

  container.querySelector('#smart-copilot-model-select')?.addEventListener('change', (e) => {
    selectedApiId = e.target.value;
    store.setActiveApi(selectedApiId);
  });

  container.querySelectorAll('.quick-ai-chip').forEach((chip) => {
    chip.addEventListener('click', () => {
      const inp = container.querySelector('#smart-quick-ai-input');
      inp.value = chip.dataset.prompt;
      inp.focus();
    });
  });

  container.querySelectorAll('.copilot-chip').forEach((chip) => {
    chip.addEventListener('click', () => {
      const ta = container.querySelector('#smart-copilot-prompt');
      ta.value = chip.dataset.prompt;
      ta.focus();
    });
  });

  async function runSmartAiEdit(instruction) {
    const qStatus = container.querySelector('#smart-quick-ai-status');
    const cStatus = container.querySelector('#smart-copilot-status');
    const msg = 'AI sedang membaca seluruh halaman dokumen dan menerapkan revisi...';
    [qStatus, cStatus].forEach((el) => {
      if (el) {
        el.textContent = msg;
        el.classList.remove('hidden', 'error');
      }
    });

    try {
      const snapshot = extractSmartSnapshotFromPages(collectLivePagesData(), store.getSchoolProfile());
      const preferredApiId = container.querySelector('#smart-copilot-model-select')?.value || selectedApiId;
      const editedJson = await editDocumentWithAi(snapshot, instruction, {
        schoolProfile: store.getSchoolProfile(),
        preferredApiId
      });

      pushUndoSnapshot();
      const docModel = aiJsonToSmartDocModel(editedJson, store.getSchoolProfile(), store.getLetterOptions());
      const converted = docModelToSmartPages(docModel, store.getSchoolProfile());
      smartDoc = { pages: converted.pages, meta: converted.meta };
      store.setState({ smartDocument: smartDoc }, false);
      renderAllSmartPages(smartDoc.pages);

      const summary = editedJson.summaryOfChanges || 'Revisi AI berhasil diterapkan ke Kanvas Word Pro.';
      [qStatus, cStatus].forEach((el) => {
        if (el) el.textContent = summary;
      });
      showToast(summary, 'success');
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
    if (e.key === 'Enter') {
      e.preventDefault();
      container.querySelector('#smart-btn-quick-ai-apply').click();
    }
  });

  container.querySelector('#smart-copilot-btn-edit')?.addEventListener('click', () => {
    const val = container.querySelector('#smart-copilot-prompt').value.trim();
    if (!val) {
      showToast('Isi perintah pengeditan terlebih dahulu.', 'warning');
      return;
    }
    runSmartAiEdit(val);
  });

  container.querySelector('#smart-copilot-btn-create')?.addEventListener('click', async () => {
    const userPrompt = container.querySelector('#smart-copilot-prompt').value.trim();
    if (!userPrompt) {
      showToast('Ketikkan jenis surat atau deskripsi dokumen yang ingin dibuat.', 'warning');
      return;
    }
    const cStatus = container.querySelector('#smart-copilot-status');
    cStatus.textContent = 'AI Copilot sedang menyusun draf surat resmi baru...';
    cStatus.classList.remove('hidden', 'error');

    try {
      const school = store.getSchoolProfile();
      const dateStr = new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
      const preferredApiId = container.querySelector('#smart-copilot-model-select')?.value || selectedApiId;
      const references = store.getAiReferences(4);

      const fullPrompt = buildCopilotUserPrompt({
        userPrompt,
        schoolName: school.schoolName,
        principalName: school.principalName,
        dateStr,
        references
      });

      const parsedJson = await callAiWithValidation({
        prompt: fullPrompt,
        systemPrompt: COPILOT_SYSTEM_PROMPT,
        preferredApiId,
        validator: (parsed) => {
          const errors = [];
          if (!parsed || typeof parsed !== 'object') errors.push('Output harus berupa objek JSON.');
          if (!parsed?.letterTitle && !parsed?.title) errors.push('Field letterTitle wajib ada.');
          if (!Array.isArray(parsed?.blocks) || parsed.blocks.length === 0) {
            errors.push('Field blocks harus array berisi minimal 1 blok.');
          }
          return { valid: errors.length === 0, errors };
        }
      });

      pushUndoSnapshot();
      const docModel = aiJsonToSmartDocModel(parsedJson, school, store.getLetterOptions());
      const converted = docModelToSmartPages(docModel, school);
      smartDoc = { pages: converted.pages, meta: converted.meta };
      store.setState({ smartDocument: smartDoc }, false);
      renderAllSmartPages(smartDoc.pages);

      cStatus.textContent = 'Surat baru berhasil dibuat ke dalam Kanvas Word Pro!';
      showToast('Surat baru selesai disusun AI ke Word Pro.', 'success');
    } catch (err) {
      cStatus.textContent = 'Gagal membuat surat AI: ' + err.message;
      cStatus.classList.add('error');
      showToast('Gagal membuat surat AI: ' + err.message, 'error');
    }
  });

  // 3 Aksi Cepat Copilot
  container.querySelector('#smart-copilot-quick-polish')?.addEventListener('click', async () => {
    const rawText = collectLivePagesData()
      .map(p => (p.html || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim())
      .join('\n\n')
      .trim();
    if (!rawText) return;

    const cStatus = container.querySelector('#smart-copilot-status');
    cStatus.textContent = 'AI sedang merapihkan naskah mentah menjadi struktur surat dinas resmi...';
    cStatus.classList.remove('hidden', 'error');

    try {
      const preferredApiId = container.querySelector('#smart-copilot-model-select')?.value || selectedApiId;
      const polishedModel = await polishMessyDocumentWithAi(rawText, {
        ...store.getLetterOptions(),
        schoolProfile: store.getSchoolProfile(),
        preferredApiId
      });
      pushUndoSnapshot();
      const converted = docModelToSmartPages(polishedModel, store.getSchoolProfile());
      smartDoc = { pages: converted.pages, meta: converted.meta };
      store.setState({ smartDocument: smartDoc }, false);
      renderAllSmartPages(smartDoc.pages);
      cStatus.textContent = 'Dokumen berhasil dirapihkan AI!';
      showToast('Naskah berhasil dirapihkan menjadi format surat resmi.', 'success');
    } catch (err) {
      cStatus.textContent = 'Gagal merapihkan: ' + err.message;
      cStatus.classList.add('error');
    }
  });

  container.querySelector('#smart-copilot-quick-puebi')?.addEventListener('click', () => {
    runSmartAiEdit('Perbaiki seluruh kata, imbuhan, tanda baca, dan kalimat surat ini ke standar resmi kedinasan PUEBI tanpa mengubah substansi maupun rincian tabel.');
  });

  container.querySelector('#smart-copilot-quick-today')?.addEventListener('click', () => {
    const todayStr = new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
    runSmartAiEdit(`Perbarui tanggal tanda tangan surat ini ke tanggal hari ini: ${todayStr}. Pertahankan isi lainnya.`);
  });

  // =========================================================================
  // MODALS: FORMAT KERTAS, GALERI KOP, SISIPKAN TABEL, SIMPAN GUDANG, SURAT KELUAR
  // =========================================================================

  function closeSmartModal() {
    modalRoot.innerHTML = '';
  }

  container.querySelector('#smart-btn-format-modal')?.addEventListener('click', () => {
    const o = store.getLetterOptions();
    const m = o.margins || { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 };
    modalRoot.innerHTML = `
      <div class="modal-overlay">
        <div class="modal-box" style="max-width: 480px;">
          <div class="modal-header">
            <h3>Pengaturan Halaman &amp; Tipografi</h3>
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
        }
      });
      closeSmartModal();
      renderAllSmartPages(collectLivePagesData());
      showToast('Pengaturan kertas & margin disimpan.', 'success');
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
      closeSmartModal();
      renderAllSmartPages(collectLivePagesData());
      showToast('Kop surat berhasil diunggah dan diterapkan.', 'success');
    };
    reader.readAsDataURL(file);
    e.target.value = '';
  });

  // Modal Sisipkan Tabel Word / Rincian
  container.querySelector('#smart-btn-insert-table')?.addEventListener('click', () => {
    modalRoot.innerHTML = `
      <div class="modal-overlay">
        <div class="modal-box" style="max-width: 440px;">
          <div class="modal-header">
            <h3>Sisipkan Tabel ke Dokumen</h3>
            <button class="btn-icon" id="smart-modal-close">✕</button>
          </div>
          <div class="modal-body" style="display:grid; gap:12px;">
            <div>
              <label class="form-label">Jenis Tabel Word</label>
              <select id="sm-tbl-type" class="form-input">
                <option value="grid">Tabel Berbingkai Penuh (Jadwal / Panitia / Rekap)</option>
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
      } else {
        const thCells = Array.from({ length: nCols }, (_, c) => `<th>Kolom ${c + 1}</th>`).join('');
        const bodyRows = Array.from({ length: Math.max(1, nRows - 1) }, () =>
          `<tr>${Array.from({ length: nCols }, () => '<td>Isi data...</td>').join('')}</tr>`
        ).join('');
        html = `<table class="smart-word-table"><thead><tr>${thCells}</tr></thead><tbody>${bodyRows}</tbody></table><p><br></p>`;
      }
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

  function resetSmartCanvasToEmpty(notify = true) {
    pushUndoSnapshot();
    smartDoc = createEmptySmartDocument();
    store.setState({ smartDocument: smartDoc }, false);
    renderAllSmartPages(smartDoc.pages);
    if (notify) {
      showToast('Lembar Word Pro berhasil dikosongkan ke kertas baru.', 'info');
    }
  }

  function openSmartSaveVaultModal({ clearAfterSave = false } = {}) {
    const snap = extractSmartSnapshotFromPages(collectLivePagesData(), store.getSchoolProfile());
    const defaultTitle = snap.letterTitle && snap.letterTitle !== 'SURAT DINAS SEKOLAH'
      ? snap.letterTitle
      : (smartDoc?.meta?.title && smartDoc.meta.title !== 'Dokumen Baru' ? smartDoc.meta.title : 'Dokumen Surat Sekolah');
    const summaryText = (snap.blocks || [])
      .filter(b => b.type === 'paragraph' && b.text)
      .map(b => b.text)
      .join(' ')
      .slice(0, 160);

    modalRoot.innerHTML = `
      <div class="modal-overlay">
        <div class="modal-box" style="max-width: 460px;">
          <div class="modal-header">
            <h3>${ICON_SAVE} ${clearAfterSave ? 'Simpan ke Gudang Sebelum Mengosongkan' : 'Simpan ke Gudang Surat'}</h3>
            <button class="btn-icon" id="smart-modal-close">✕</button>
          </div>
          <div class="modal-body" style="display:grid; gap:10px;">
            <div><label class="form-label">Judul Arsip</label><input id="sm-v-title" class="form-input" value="${escapeHtmlText(defaultTitle)}"></div>
            <div><label class="form-label">Kategori</label><input id="sm-v-cat" class="form-input" value="Surat Dinas"></div>
            <div><label class="form-label">Nomor Surat</label><input id="sm-v-nomor" class="form-input" value="${escapeHtmlText(snap.nomorSurat)}"></div>
            <div><label class="form-label">Perihal</label><input id="sm-v-hal" class="form-input" value="${escapeHtmlText(snap.hal)}"></div>
          </div>
          <div class="modal-footer">
            <button class="btn btn-secondary" id="smart-modal-cancel">Batal</button>
            <button class="btn btn-primary" id="smart-modal-save-ok">${clearAfterSave ? 'Simpan &amp; Kosongkan Kertas' : 'Simpan ke Gudang'}</button>
          </div>
        </div>
      </div>`;
    modalRoot.querySelector('#smart-modal-close').onclick = closeSmartModal;
    modalRoot.querySelector('#smart-modal-cancel').onclick = closeSmartModal;
    modalRoot.querySelector('#smart-modal-save-ok').onclick = () => {
      const liveDocModel = getActiveSmartDocModel();
      const titleVal = modalRoot.querySelector('#sm-v-title').value.trim() || defaultTitle;
      store.addLetterVaultEntry({
        title: titleVal,
        category: modalRoot.querySelector('#sm-v-cat').value.trim() || 'Surat Dinas',
        nomorSurat: modalRoot.querySelector('#sm-v-nomor').value.trim() || '-',
        tanggalSurat: snap.signer?.date || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
        targetPerson: snap.tujuan || '-',
        perihal: modalRoot.querySelector('#sm-v-hal').value.trim() || titleVal,
        summary: summaryText || 'Dokumen persuratan dari Kanvas Word Pro.',
        document: liveDocModel,
        isAiReference: true,
        source: 'created'
      });
      closeSmartModal();
      if (clearAfterSave) {
        resetSmartCanvasToEmpty(false);
        showToast(`"${titleVal}" tersimpan ke Gudang Surat & kertas siap dipakai ulang!`, 'success');
      } else {
        showToast('Dokumen Word Pro berhasil disimpan ke Gudang Surat!', 'success');
      }
    };
  }

  function openSmartRecordOutgoingModal({ clearAfterSave = false } = {}) {
    const snap = extractSmartSnapshotFromPages(collectLivePagesData(), store.getSchoolProfile());
    const nextInfo = store.getNextFormattedOutgoingNomor({
      title: snap.letterTitle,
      perihal: snap.hal,
      currentNomor: snap.nomorSurat
    });
    const firstPara = (snap.blocks || []).find(b => b.type === 'paragraph' && b.text)?.text || '';

    modalRoot.innerHTML = `
      <div class="modal-overlay">
        <div class="modal-box" style="max-width: 500px;">
          <div class="modal-header">
            <h3>${ICON_CLIPBOARD_CHECK} ${clearAfterSave ? 'Catat ke Surat Keluar Sebelum Mengosongkan' : 'Catat ke Buku Surat Keluar'}</h3>
            <button class="btn-icon" id="smart-modal-close">✕</button>
          </div>
          <div class="modal-body" style="display:grid; gap:10px;">
            <div>
              <label class="form-label">Nomor Surat Resmi (Otomatis Berurutan)</label>
              <input id="sm-out-nomor" class="form-input" value="${escapeHtmlText(nextInfo.nomorSurat)}">
            </div>
            <div><label class="form-label">Judul / Jenis Surat</label><input id="sm-out-title" class="form-input" value="${escapeHtmlText(snap.letterTitle)}"></div>
            <div><label class="form-label">Perihal</label><input id="sm-out-hal" class="form-input" value="${escapeHtmlText(snap.hal || snap.letterTitle)}"></div>
            <div><label class="form-label">Tujuan / Kepada Yth.</label><input id="sm-out-to" class="form-input" value="${escapeHtmlText(snap.tujuan)}"></div>
          </div>
          <div class="modal-footer">
            <button class="btn btn-secondary" id="smart-modal-cancel">Batal</button>
            <button class="btn btn-primary" id="smart-modal-save-out">${clearAfterSave ? 'Catat Keluar &amp; Kosongkan' : 'Simpan Surat Keluar'}</button>
          </div>
        </div>
      </div>`;
    modalRoot.querySelector('#smart-modal-close').onclick = closeSmartModal;
    modalRoot.querySelector('#smart-modal-cancel').onclick = closeSmartModal;
    modalRoot.querySelector('#smart-modal-save-out').onclick = () => {
      const liveModel = getActiveSmartDocModel();
      const recordedNomor = modalRoot.querySelector('#sm-out-nomor').value.trim();
      const titleVal = modalRoot.querySelector('#sm-out-title').value.trim() || snap.letterTitle;
      const recorded = store.addOutgoingLetter({
        nomorSurat: recordedNomor,
        sequenceNumber: nextInfo.sequenceNumber,
        typeCode: nextInfo.typeCode,
        year: nextInfo.year,
        title: titleVal,
        perihal: modalRoot.querySelector('#sm-out-hal').value.trim() || titleVal,
        recipientName: modalRoot.querySelector('#sm-out-to').value.trim() || 'Umum / Tanpa Nama Penerima',
        tanggalSurat: snap.signer?.date || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
        signerName: snap.signer?.name || store.getSchoolProfile()?.principalName || 'Kepala Sekolah',
        summary: firstPara.slice(0, 160),
        document: liveModel
      });

      closeSmartModal();
      if (clearAfterSave) {
        resetSmartCanvasToEmpty(false);
        showToast(`Tercatat (${recorded.nomorSurat}) ke Surat Keluar & kertas dikosongkan!`, 'success');
      } else {
        const afterInfo = store.getNextFormattedOutgoingNomor({
          title: snap.letterTitle,
          perihal: snap.hal
        });
        const nomorEl = pagesStack.querySelector('[data-smart-role="nomor"]');
        if (nomorEl) {
          nomorEl.textContent = `Nomor: ${afterInfo.nomorSurat}`;
          schedulePersistSmartDoc();
        }
        showToast(`Tercatat! Nomor ${recorded.nomorSurat} disimpan ke Surat Keluar.`, 'success');
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
    modalRoot.querySelector('#sm-ref-vault').onclick = () => openSmartSaveVaultModal({ clearAfterSave: true });
    modalRoot.querySelector('#sm-ref-outgoing').onclick = () => openSmartRecordOutgoingModal({ clearAfterSave: true });
    modalRoot.querySelector('#sm-ref-discard').onclick = () => {
      closeSmartModal();
      resetSmartCanvasToEmpty(true);
    };
  }

  // Modal Simpan ke Gudang Surat
  container.querySelector('#smart-btn-save-vault')?.addEventListener('click', () => openSmartSaveVaultModal({ clearAfterSave: false }));

  // Tombol Refresh / Bersihkan Lembar Kerja Word Pro
  container.querySelector('#smart-btn-refresh')?.addEventListener('click', openSmartRefreshConfirmModal);

  // Modal Catat Surat Keluar Resmi (berbagi nomor urut buku agenda!)
  container.querySelector('#smart-btn-record-outgoing')?.addEventListener('click', () => openSmartRecordOutgoingModal({ clearAfterSave: false }));

  // Expose helper download lokal bila dibutuhkan oleh shortcut
  container.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
      container.querySelector('#smart-btn-save-vault')?.click();
    }
  });
}
