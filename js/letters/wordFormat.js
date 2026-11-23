/**
 * wordFormat.js - Normalisasi & Peta Format Dokumen Lengkap Microsoft Word
 * Digunakan bersama oleh State, Word Pro Canvas, Lembar Kerja, Renderer .docx,
 * Prompt AI, dan AI Smart Surgical Patcher (zero-dependency, Browser & Node.js).
 */

export const SUPPORTED_PAPER_SIZES = Object.freeze(['a4', 'f4', 'letter', 'legal']);

export const SUPPORTED_FONT_FAMILIES = Object.freeze([
  'Times New Roman',
  'Arial',
  'Calibri',
  'Georgia',
  'Garamond',
  'Cambria',
  'Tahoma',
  'Verdana',
  'Segoe UI',
  'Courier New'
]);

export const LIST_STYLE_CONFIG = Object.freeze({
  decimal: {
    key: 'decimal',
    label: '1, 2, 3 (Angka)',
    htmlTag: 'ol',
    htmlType: '1',
    cssType: 'decimal',
    ooxmlFmt: 'decimal',
    ooxmlNumId: '2',
    abstractNumId: '2'
  },
  'lower-alpha': {
    key: 'lower-alpha',
    label: 'a, b, c (Huruf Kecil)',
    htmlTag: 'ol',
    htmlType: 'a',
    cssType: 'lower-alpha',
    ooxmlFmt: 'lowerLetter',
    ooxmlNumId: '4',
    abstractNumId: '4'
  },
  'upper-alpha': {
    key: 'upper-alpha',
    label: 'A, B, C (Huruf Besar)',
    htmlTag: 'ol',
    htmlType: 'A',
    cssType: 'upper-alpha',
    ooxmlFmt: 'upperLetter',
    ooxmlNumId: '5',
    abstractNumId: '5'
  },
  'lower-roman': {
    key: 'lower-roman',
    label: 'i, ii, iii (Romawi Kecil)',
    htmlTag: 'ol',
    htmlType: 'i',
    cssType: 'lower-roman',
    ooxmlFmt: 'lowerRoman',
    ooxmlNumId: '6',
    abstractNumId: '6'
  },
  'upper-roman': {
    key: 'upper-roman',
    label: 'I, II, III (Romawi Besar)',
    htmlTag: 'ol',
    htmlType: 'I',
    cssType: 'upper-roman',
    ooxmlFmt: 'upperRoman',
    ooxmlNumId: '7',
    abstractNumId: '7'
  },
  bullet: {
    key: 'bullet',
    label: '• Bullet (Poin)',
    htmlTag: 'ul',
    htmlType: '',
    cssType: 'disc',
    ooxmlFmt: 'bullet',
    ooxmlNumId: '3',
    abstractNumId: '3'
  }
});

export const PAGE_NUMBERING_DEFAULT = Object.freeze({
  enabled: false,
  position: 'bottom-center', // 'bottom-center' | 'bottom-right' | 'top-right'
  format: 'decimal',         // 'decimal' | 'lowerRoman' | 'upperRoman'
  startAt: 1
});

export function cmToTwips(cm = 0) {
  const n = Number(cm);
  if (!Number.isFinite(n) || n < 0) return 0;
  return Math.round(n * 567);
}

export function ptToTwips(pt = 0) {
  const n = Number(pt);
  if (!Number.isFinite(n) || n < 0) return 0;
  return Math.round(n * 20);
}

export function cssLenToCm(rawCss = '') {
  const s = String(rawCss || '').trim().toLowerCase();
  if (!s) return undefined;
  const m = s.match(/^(-?[\d.]+)\s*(cm|mm|in|pt|px)?$/);
  if (!m) return undefined;
  const val = parseFloat(m[1]);
  if (!Number.isFinite(val)) return undefined;
  const unit = m[2] || 'cm';
  let cm = val;
  if (unit === 'mm') cm = val / 10;
  else if (unit === 'in') cm = val * 2.54;
  else if (unit === 'pt') cm = val / 28.3465;
  else if (unit === 'px') cm = val / 37.7953;
  return Math.round(cm * 100) / 100;
}

export function cssLenToTwips(rawCss = '') {
  const cm = cssLenToCm(rawCss);
  if (cm === undefined) return undefined;
  return cmToTwips(cm);
}

export function clampMarginCm(val, fallback = 2.0) {
  const n = Number(val);
  if (!Number.isFinite(n) || n <= 0) return fallback;
  const cm = n > 50 ? n / 567 : n;
  return Number(Math.max(0.4, Math.min(8.0, cm)).toFixed(2));
}

export function normalizeAlignment(raw = '') {
  const v = String(raw || '').trim().toLowerCase();
  if (v === 'left' || v === 'kiri') return 'left';
  if (v === 'center' || v === 'tengah') return 'center';
  if (v === 'right' || v === 'kanan') return 'right';
  if (v === 'justify' || v === 'both' || v === 'rata' || v === 'kiri-kanan') return 'both';
  return '';
}

export function normalizeFontFamily(raw = '', fallback = 'Times New Roman') {
  const s = String(raw || '').replace(/['"]/g, '').split(',')[0].trim();
  if (!s) return fallback;
  const matched = SUPPORTED_FONT_FAMILIES.find(f => f.toLowerCase() === s.toLowerCase());
  return matched || s.slice(0, 60);
}

export function normalizeListStyle(rawStyle = '', ordered = true) {
  const trimmed = String(rawStyle || '').trim();
  if (trimmed === 'A') return 'upper-alpha';
  if (trimmed === 'I') return 'upper-roman';
  const s = trimmed.toLowerCase();
  if (s === 'a' || s === 'lower-alpha' || s === 'loweralpha' || s === 'lower-latin' || s === 'lowerletter') {
    return 'lower-alpha';
  }
  if (s === 'upper-alpha' || s === 'upperalpha' || s === 'upper-latin' || s === 'upperletter') {
    return 'upper-alpha';
  }
  if (s === 'i' || s === 'lower-roman' || s === 'lowerroman') {
    return 'lower-roman';
  }
  if (s === 'upper-roman' || s === 'upperroman') {
    return 'upper-roman';
  }
  if (s === '1' || s === 'decimal' || s === 'angka' || s === 'number') {
    return 'decimal';
  }
  if (s === 'bullet' || s === 'disc' || s === 'circle' || s === 'square' || s === 'unordered' || s === 'ul') {
    return 'bullet';
  }
  return ordered ? 'decimal' : 'bullet';
}

export function getListStyleSpec(rawStyle = '', ordered = true) {
  const key = normalizeListStyle(rawStyle, ordered);
  return LIST_STYLE_CONFIG[key] || LIST_STYLE_CONFIG.decimal;
}

export function parseListNodeMeta(listNode) {
  const tag = String(listNode?.tag || '').toLowerCase();
  const isOl = tag === 'ol';
  const typeAttr = listNode?.attrs?.type || '';
  const dataStyle = listNode?.attrs?.['data-list-style'] || '';
  const styleCss = String(listNode?.attrs?.style || '').toLowerCase();
  const cssMatch = styleCss.match(/list-style-type\s*:\s*([a-z-]+)/);
  const rawStyle = dataStyle || typeAttr || (cssMatch ? cssMatch[1] : '');
  const listStyle = normalizeListStyle(rawStyle, isOl);
  const startAttr = parseInt(listNode?.attrs?.start || '1', 10);
  const startAt = Number.isInteger(startAttr) && startAttr > 0 ? startAttr : 1;
  return {
    ordered: listStyle !== 'bullet',
    listStyle,
    startAt
  };
}

export function toRomanNumeral(num = 1, upper = true) {
  const n = Math.max(1, Math.min(3999, Math.floor(Number(num) || 1)));
  const table = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'],
    [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'],
    [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']
  ];
  let rem = n;
  let out = '';
  for (const [val, sym] of table) {
    while (rem >= val) {
      out += sym;
      rem -= val;
    }
  }
  return upper ? out : out.toLowerCase();
}

export function formatPageNumberLabel(pageNum = 1, format = 'decimal') {
  const n = Math.max(1, Math.floor(Number(pageNum) || 1));
  const f = String(format || 'decimal').toLowerCase();
  if (f === 'lowerroman' || f === 'lower-roman') return toRomanNumeral(n, false);
  if (f === 'upperroman' || f === 'upper-roman') return toRomanNumeral(n, true);
  if (f === 'lowerletter' || f === 'lower-alpha') {
    return String.fromCharCode(97 + ((n - 1) % 26));
  }
  return String(n);
}

export function normalizePageNumbering(raw = {}, base = PAGE_NUMBERING_DEFAULT) {
  const fallback = { ...PAGE_NUMBERING_DEFAULT, ...(base && typeof base === 'object' ? base : {}) };
  if (raw === false || raw === 'none' || raw === 'off') {
    return { ...fallback, enabled: false };
  }
  if (raw === true || raw === 'on') {
    return { ...fallback, enabled: true };
  }
  if (typeof raw === 'string') {
    const s = raw.trim().toLowerCase();
    const pos = ['bottom-center', 'bottom-right', 'top-right'].includes(s) ? s : fallback.position;
    return { ...fallback, enabled: true, position: pos };
  }
  if (!raw || typeof raw !== 'object') return fallback;

  const posRaw = String(raw.position || fallback.position || 'bottom-center').trim().toLowerCase();
  const position = ['bottom-center', 'bottom-right', 'top-right'].includes(posRaw) ? posRaw : 'bottom-center';
  const fmtRaw = String(raw.format || fallback.format || 'decimal').trim().toLowerCase();
  const format = fmtRaw === 'lowerroman' || fmtRaw === 'lower-roman'
    ? 'lowerRoman'
    : fmtRaw === 'upperroman' || fmtRaw === 'upper-roman'
      ? 'upperRoman'
      : 'decimal';
  const startAt = Math.max(1, Math.min(999, parseInt(raw.startAt ?? fallback.startAt ?? 1, 10) || 1));
  return {
    enabled: raw.enabled !== undefined ? Boolean(raw.enabled) : Boolean(fallback.enabled),
    position,
    format,
    startAt
  };
}

export function normalizePageSetupOptions(raw = {}, base = {}) {
  const curMargins = base?.margins && typeof base.margins === 'object'
    ? base.margins
    : { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 };

  const inMargins = raw?.margins && typeof raw.margins === 'object' ? raw.margins : null;
  const nextMargins = {
    top: clampMarginCm(inMargins?.top ?? raw?.marginTop ?? curMargins.top, curMargins.top ?? 2.0),
    bottom: clampMarginCm(inMargins?.bottom ?? raw?.marginBottom ?? curMargins.bottom, curMargins.bottom ?? 2.0),
    left: clampMarginCm(inMargins?.left ?? raw?.marginLeft ?? curMargins.left, curMargins.left ?? 2.5),
    right: clampMarginCm(inMargins?.right ?? raw?.marginRight ?? curMargins.right, curMargins.right ?? 2.0)
  };

  const paperRaw = String(raw?.paperSize || base?.paperSize || 'a4').trim().toLowerCase();
  const paperSize = SUPPORTED_PAPER_SIZES.includes(paperRaw) ? paperRaw : 'a4';
  const orientRaw = String(raw?.orientation || base?.orientation || 'portrait').trim().toLowerCase();
  const orientation = orientRaw === 'landscape' ? 'landscape' : 'portrait';

  const fontFamily = normalizeFontFamily(raw?.fontFamily || base?.fontFamily || 'Times New Roman');
  const fsNum = parseFloat(raw?.fontSize ?? base?.fontSize ?? 12);
  const fontSize = String(Number.isFinite(fsNum) ? Math.max(8, Math.min(36, Math.round(fsNum))) : 12);

  const lsNum = parseFloat(raw?.lineSpacing ?? base?.lineSpacing ?? 1.15);
  const lineSpacing = Number.isFinite(lsNum) && lsNum >= 0.8 && lsNum <= 3.0
    ? (String(lsNum) === '1' ? '1.0' : String(Number(lsNum.toFixed(2))))
    : String(base?.lineSpacing || '1.15');

  const psNum = parseFloat(raw?.paragraphSpacing ?? base?.paragraphSpacing ?? 6);
  const paragraphSpacing = String(Number.isFinite(psNum) ? Math.max(0, Math.min(36, Math.round(psNum))) : 6);

  const showKop = raw?.showKop !== undefined
    ? Boolean(raw.showKop)
    : (base?.showKop !== undefined ? Boolean(base.showKop) : true);

  const pageNumbering = normalizePageNumbering(
    raw?.pageNumbering !== undefined ? raw.pageNumbering : base?.pageNumbering,
    base?.pageNumbering
  );

  return {
    paperSize,
    orientation,
    margins: nextMargins,
    fontFamily,
    fontSize,
    lineSpacing,
    paragraphSpacing,
    showKop,
    pageNumbering
  };
}

/**
 * Mengekstrak pembaruan Page Setup / Format Dokumen dari JSON respons AI (baik `pageSetup`
 * di level atas maupun operasi patch `setPageFormat` di array `patches`).
 */
export function extractPageSetupFromAiPayload(aiJson = {}, baseOptions = {}) {
  if (!aiJson || typeof aiJson !== 'object') return null;

  const patchFormatEntries = Array.isArray(aiJson.patches)
    ? aiJson.patches.filter(p => p && typeof p === 'object' && p.op === 'setPageFormat')
    : [];

  const srcList = [];
  if (aiJson.pageSetup && typeof aiJson.pageSetup === 'object') {
    srcList.push(aiJson.pageSetup);
  }
  const topLevelKeys = [
    'paperSize', 'orientation', 'margins',
    'marginTop', 'marginBottom', 'marginLeft', 'marginRight',
    'fontFamily', 'fontSize', 'lineSpacing', 'paragraphSpacing',
    'showKop', 'pageNumbering'
  ];
  const topPatch = {};
  for (const k of topLevelKeys) {
    if (aiJson[k] !== undefined) topPatch[k] = aiJson[k];
  }
  if (Object.keys(topPatch).length > 0) {
    srcList.push(topPatch);
  }
  for (const pf of patchFormatEntries) {
    srcList.push(pf);
  }

  if (srcList.length === 0) return null;

  let acc = normalizePageSetupOptions({}, baseOptions);
  const changedFields = new Set();

  for (const src of srcList) {
    const mergedRaw = {};
    if (src.paperSize !== undefined) { mergedRaw.paperSize = src.paperSize; changedFields.add('paperSize'); }
    if (src.orientation !== undefined) { mergedRaw.orientation = src.orientation; changedFields.add('orientation'); }
    if (src.margins !== undefined || src.marginTop !== undefined || src.marginBottom !== undefined || src.marginLeft !== undefined || src.marginRight !== undefined) {
      mergedRaw.margins = {
        ...(acc.margins || {}),
        ...(src.margins && typeof src.margins === 'object' ? src.margins : {}),
        ...(src.marginTop !== undefined ? { top: src.marginTop } : {}),
        ...(src.marginBottom !== undefined ? { bottom: src.marginBottom } : {}),
        ...(src.marginLeft !== undefined ? { left: src.marginLeft } : {}),
        ...(src.marginRight !== undefined ? { right: src.marginRight } : {})
      };
      changedFields.add('margins');
    }
    if (src.fontFamily !== undefined) { mergedRaw.fontFamily = src.fontFamily; changedFields.add('fontFamily'); }
    if (src.fontSize !== undefined) { mergedRaw.fontSize = src.fontSize; changedFields.add('fontSize'); }
    if (src.lineSpacing !== undefined) { mergedRaw.lineSpacing = src.lineSpacing; changedFields.add('lineSpacing'); }
    if (src.paragraphSpacing !== undefined) { mergedRaw.paragraphSpacing = src.paragraphSpacing; changedFields.add('paragraphSpacing'); }
    if (src.showKop !== undefined) { mergedRaw.showKop = src.showKop; changedFields.add('showKop'); }
    if (src.pageNumbering !== undefined) { mergedRaw.pageNumbering = src.pageNumbering; changedFields.add('pageNumbering'); }
    acc = normalizePageSetupOptions(mergedRaw, acc);
  }

  if (changedFields.size === 0) return null;

  const partial = {};
  for (const key of changedFields) {
    partial[key] = acc[key];
  }
  return { partial, full: acc };
}
