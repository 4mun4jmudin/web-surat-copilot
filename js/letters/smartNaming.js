/**
 * smartNaming.js - Ekstraktor Metadata Surat Otomatis (Format Sistem + Pemindai AI)
 * & Pemformat Penamaan Arsip Surat Keluar, Gudang Surat, serta Unduhan .docx.
 *
 * Menganalisis DocumentModel, halaman Word Pro, atau berkas .docx impor untuk menghasilkan:
 * - Kode Jenis Surat (ST, KET, UND, KEP, PGL, PBR, PRM, DISP, BA, REK, PKL, KUR, KES, TU)
 * - Judul arsip terstruktur & rapih: [Jenis Surat] — [Nama/Pihak] — [Keperluan] ([Tanggal])
 * - Atas Nama / Penerima / Pihak Ditugaskan (menghindari nama Kepsek penandatangan bila ia pemberi tugas)
 * - Keperluan / Perihal & Tanggal Surat
 * - Kategori Gudang Surat & Nama Berkas Unduhan .docx yang rapi dan bebas karakter liar
 */

import { detectLetterTypeCode, LETTER_TYPE_CODES } from './numbering.js';
import { SCHOOL_UNITS, normalizeUnitId, resolveDocumentSchoolProfile } from './schoolProfile.js';
import { callAiApi, cleanJsonOutput, getActiveAiConfig } from '../ai/client.js';

const INDONESIAN_MONTHS = [
  'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
];

const SHORT_MONTH_MAP = {
  januari: 'Jan', februari: 'Feb', pebruari: 'Feb', maret: 'Mar', april: 'Apr',
  mei: 'Mei', juni: 'Jun', juli: 'Jul', agustus: 'Agt', september: 'Sep',
  oktober: 'Okt', november: 'Nov', nopember: 'Nov', desember: 'Des'
};

const TYPE_TO_HUMAN_LABEL = {
  ST: 'Surat Tugas',
  SPPD: 'SPPD',
  KET: 'Surat Keterangan',
  UND: 'Surat Undangan',
  KEP: 'Surat Keputusan',
  PGL: 'Surat Panggilan',
  PBR: 'Surat Pemberitahuan',
  PRM: 'Surat Permohonan',
  DISP: 'Surat Dispensasi',
  BA: 'Berita Acara',
  REK: 'Surat Rekomendasi',
  PKL: 'Surat PKL / Prakerin',
  KUR: 'Surat Kurikulum',
  KES: 'Surat Kesiswaan',
  TU: 'Surat Tata Usaha'
};

const TYPE_TO_VAULT_CATEGORY = {
  ST: 'Surat Tugas',
  KET: 'Surat Keterangan',
  UND: 'Surat Undangan',
  PGL: 'Surat Panggilan',
  KEP: 'Surat Keputusan',
  PBR: 'Surat Pemberitahuan',
  PRM: 'Surat Permohonan',
  DISP: 'Surat Dispensasi',
  BA: 'Berita Acara',
  REK: 'Surat Rekomendasi',
  PKL: 'Surat PKL',
  KUR: 'Surat Kurikulum',
  KES: 'Surat Kesiswaan',
  TU: 'Surat Dinas'
};

const GENERIC_TITLES_RE = /^(surat dinas(\s+sekolah)?|dokumen(\s+baru|\s+persuratan|\s+surat\s+sekolah|\s+tanpa\s+judul)?|untitled|new\s+document|template\s+baru|salinan\s*-?\s*.*)$/i;
const PLACEHOLDER_VALUE_RE = /^(?:\[.*\]|\.{2,}|_{2,}|-+|pihak\s+terkait|di\s+tempat|tempat|di|kepada|yth\.?|umum\s*\/?\s*tanpa\s+nama\s+penerima|nama\s+siswa|nama\s+guru)$/i;
const SIGNATORY_INTRO_RE = /yang\s+bertanda\s+tangan\s+di\s+bawah\s+ini\s*:?\s*$/i;
const TARGET_INTRO_RE = /(menerangkan\s+bahwa|menugaskan\s+kepada|memberikan\s+tugas\s+kepada|memanggil\s+orang\s+tua|identitas\s+(?:siswa|pegawai|peserta)|daftar\s+(?:nama|guru|siswa|peserta)|pegawai\s+yang\s+diberi\s+perintah)/i;

function cleanSpaces(text = '') {
  return String(text || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function stripHtmlTags(html = '') {
  return cleanSpaces(
    String(html || '')
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/(?:p|div|tr|li|h[1-6])>/gi, '\n')
      .replace(/<\/(?:td|th)>/gi, ' | ')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/gi, ' ')
      .replace(/&amp;/gi, '&')
      .replace(/&lt;/gi, '<')
      .replace(/&gt;/gi, '>')
      .replace(/&quot;/gi, '"')
  );
}

function normalizeTextLinesFromHtml(html = '') {
  return String(html || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(?:p|div|tr|li|h[1-6])>/gi, '\n')
    .replace(/<\/(?:td|th)>/gi, '\t')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .split(/\r?\n/)
    .map(l => cleanSpaces(l))
    .filter(Boolean);
}

function toTitleCaseSmart(str = '') {
  const s = cleanSpaces(str);
  if (!s) return '';
  const minor = new Set(['dan', 'atau', 'di', 'ke', 'dari', 'pada', 'untuk', 'yang', 'dengan', 'dalam', 'atas', 'oleh', 'sebagai', 'serta', 'ini', 'tersebut']);
  const acronyms = new Set([
    'SPPD', 'PKL', 'OSIS', 'PUEBI', 'SMK', 'SMKS', 'SMP', 'SMPS', 'IT', 'PPDB', 'MPLS',
    'LDKS', 'PAS', 'PAT', 'PTS', 'STS', 'SAS', 'US', 'UKK', 'ANBK', 'P5', 'SK', 'SP', 'TU',
    'DUDI', 'KIP', 'PIP', 'BOS', 'BPMU', 'NISN', 'NIP', 'NUPTK', 'QR', 'PDF', 'LKS', 'MGMP',
    'MKKS', 'WIB', 'WITA', 'WIT', 'RT', 'RW', 'II', 'III', 'IV', 'VI', 'VII', 'VIII', 'IX', 'XI', 'XII',
    'S.Pd.', 'M.Pd.', 'S.Kom.', 'S.T.', 'M.T.', 'S.Ag.', 'S.E.'
  ]);
  return s.split(/\s+/).map((word, idx) => {
    const cleanW = word.replace(/^[("'[]+|[)"'\],.;:!?]+$/g, '');
    if (acronyms.has(cleanW) || acronyms.has(cleanW.toUpperCase())) {
      const prefix = word.slice(0, word.indexOf(cleanW));
      const suffix = word.slice(word.indexOf(cleanW) + cleanW.length);
      const exactAcc = acronyms.has(cleanW) ? cleanW : cleanW.toUpperCase();
      return `${prefix}${exactAcc}${suffix}`;
    }
    const lower = word.toLowerCase();
    if (idx > 0 && minor.has(cleanW.toLowerCase())) {
      return lower;
    }
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  }).join(' ');
}

function normalizeHeadlineCase(raw = '') {
  const s = cleanSpaces(raw).replace(/^[:\-–—\s]+|[:\-–—\s]+$/g, '');
  if (!s) return '';
  // Pertahankan frasa singkatan dan tampilkan huruf kapital awal tiap kata yang rapi
  if (s === s.toUpperCase() || s === s.toLowerCase()) {
    return toTitleCaseSmart(s);
  }
  return s;
}

function buildSignatoryRootTokens(school = {}, signerName = '') {
  const rawNames = [
    school?.principalName,
    school?.tuName,
    signerName,
    'Yangyang Budiman',
    'Dian Nurhamdani'
  ];
  const out = [];
  for (const r of rawNames) {
    const clean = cleanSpaces(r).split(',')[0].toLowerCase().trim();
    if (clean.length >= 4 && !out.includes(clean)) out.push(clean);
  }
  return out;
}

function isSignatoryPerson(name = '', signatoryTokens = []) {
  const lower = cleanSpaces(name).split(',')[0].toLowerCase();
  if (!lower) return false;
  return signatoryTokens.some(t => lower.includes(t) || t.includes(lower));
}

function isMeaningfulPersonOrRecipient(val = '') {
  const s = cleanSpaces(val).replace(/^[:\s]+/, '').trim();
  if (!s || s.length < 2 || s.length > 110) return false;
  if (PLACEHOLDER_VALUE_RE.test(s)) return false;
  if (/^\d+$/.test(s)) return false;
  if (/^kec\.?\s+/i.test(s) || /^kab\.?\s+/i.test(s) || /^jalan\s+|^jl\./i.test(s)) return false;
  return true;
}

/**
 * Mengubah format tanggal apa pun ("2026-11-23", "23, November 2026", "Garut, 23 November 2026")
 * ke format panjang Indonesia ("23 November 2026") dan format pendek ("23 Nov 2026").
 */
export function normalizeIndonesianDateText(rawDate = '') {
  const raw = cleanSpaces(rawDate);
  if (!raw) return { full: '', short: '' };

  // ISO YYYY-MM-DD
  const iso = raw.match(/\b(20\d{2})-(\d{1,2})-(\d{1,2})\b/);
  if (iso) {
    const y = iso[1];
    const mIdx = Math.max(0, Math.min(11, parseInt(iso[2], 10) - 1));
    const d = parseInt(iso[3], 10);
    const fullM = INDONESIAN_MONTHS[mIdx];
    const shortM = SHORT_MONTH_MAP[fullM.toLowerCase()] || fullM.slice(0, 3);
    return { full: `${d} ${fullM} ${y}`, short: `${d} ${shortM} ${y}` };
  }

  // DD [,] Bulan YYYY (opsional rentang "23 - 26 November 2026")
  const idMatch = raw.match(/(\d{1,2}(?:\s*[-–]\s*\d{1,2})?)\s*,?\s+(januari|februari|pebruari|maret|april|mei|juni|juli|agustus|september|oktober|november|nopember|desember)\s+(20\d{2})/i);
  if (idMatch) {
    const dayPart = cleanSpaces(idMatch[1]).replace(/\s*[-–]\s*/, '–');
    const monthLower = idMatch[2].toLowerCase();
    const monthCap = monthLower === 'pebruari' ? 'Februari' : (monthLower === 'nopember' ? 'November' : (monthLower.charAt(0).toUpperCase() + monthLower.slice(1)));
    const shortM = SHORT_MONTH_MAP[monthLower] || monthCap.slice(0, 3);
    const year = idMatch[3];
    return {
      full: `${dayPart} ${monthCap} ${year}`.replace(/–/, ' - '),
      short: `${dayPart} ${shortM} ${year}`
    };
  }

  // DD/MM/YYYY atau DD-MM-YYYY
  const slash = raw.match(/\b(\d{1,2})[\/\-](\d{1,2})[\/\-](20\d{2})\b/);
  if (slash) {
    const d = parseInt(slash[1], 10);
    const mIdx = Math.max(0, Math.min(11, parseInt(slash[2], 10) - 1));
    const y = slash[3];
    const fullM = INDONESIAN_MONTHS[mIdx];
    const shortM = SHORT_MONTH_MAP[fullM.toLowerCase()] || fullM.slice(0, 3);
    return { full: `${d} ${fullM} ${y}`, short: `${d} ${shortM} ${y}` };
  }

  return { full: raw, short: raw };
}

function formatPeopleList(names = []) {
  const unique = [];
  const seen = new Set();
  for (const raw of names) {
    const clean = cleanSpaces(raw);
    const key = clean.toLowerCase();
    if (!isMeaningfulPersonOrRecipient(clean) || seen.has(key)) continue;
    seen.add(key);
    unique.push(clean);
  }
  if (unique.length === 0) return '';
  if (unique.length === 1) return unique[0];
  if (unique.length === 2) return `${unique[0]} & ${unique[1]}`;
  return `${unique[0]} dkk (${unique.length} orang)`;
}

function extractTextFromBlock(block) {
  if (!block || typeof block !== 'object') return '';
  if (block.type === 'letterTitle') return cleanSpaces(`${block.title || ''} ${block.nomor || ''}`);
  if (block.type === 'paragraph' || block.type === 'heading') {
    if (Array.isArray(block.runs)) return cleanSpaces(block.runs.map(r => r?.text || '').join(''));
    return cleanSpaces(block.text || '');
  }
  return '';
}

function extractCellText(cell) {
  if (!cell) return '';
  if (typeof cell === 'string') return cleanSpaces(cell);
  if (Array.isArray(cell.paragraphs) && cell.paragraphs.length > 0) {
    return cleanSpaces(cell.paragraphs.map(p => extractTextFromBlock(p)).join(' '));
  }
  return cleanSpaces(cell.text || '');
}

function findAssigneeNameColumnIndex(headerCells = [], secondRowCells = []) {
  const idx = headerCells.findIndex(c => /^(nama(\s+(lengkap|guru|siswa|peserta|pegawai))?)$/i.test(cleanSpaces(c)));
  if (idx < 0) return -1;
  const secondVal = cleanSpaces(headerCells[1] || '');
  const nextFirstCol = cleanSpaces(secondRowCells[0] || '');
  const looksLikeKvIdentity = idx === 0 && (
    /^:\s*/.test(secondVal)
    || /^(?:kelas|nisn|nip|nuptk|jabatan|alamat|tempat|unit\s*kerja|ttl|tanggal\s+lahir|jenis\s+kelamin|tempat\s*\/?\s*tanggal)/i.test(nextFirstCol)
  );
  return looksLikeKvIdentity ? -1 : idx;
}

function collectFromDocModel(docModel = null, schoolProfile = {}) {
  const blocks = Array.isArray(docModel?.blocks) ? docModel.blocks : [];
  let titleText = '';
  let nomorText = '';
  let perihalText = '';
  let signatureDate = '';
  let signerName = '';
  const allParagraphs = [];
  const signatoryTokens = buildSignatoryRootTokens(schoolProfile);

  const assigneeNames = [];
  const identityCandidates = [];
  const ythRecipients = [];
  const sppdMeta = { pegawai: '', maksud: '', tempat: '', tglBerangkat: '' };

  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i];
    if (!b || b.type === 'letterKop') continue;

    if (b.type === 'letterTitle') {
      if (!titleText && b.title) titleText = cleanSpaces(b.title);
      if (!nomorText && b.nomor) nomorText = cleanSpaces(String(b.nomor).replace(/^Nomor\s*:\s*/i, ''));
      if (b.metaRecipient) {
        const halRow = (b.metaRecipient.metaRows || []).find(r => /perihal|hal/i.test(r.label || ''));
        if (halRow?.value && !perihalText) perihalText = cleanSpaces(halRow.value);
        for (const l of (b.metaRecipient.recipientLines || [])) {
          const line = cleanSpaces(l);
          if (/^yth\.?/i.test(line)) {
            const c = cleanSpaces(line.replace(/^yth\.?\s*/i, ''));
            if (isMeaningfulPersonOrRecipient(c)) ythRecipients.push(c);
          }
        }
      }
      continue;
    }

    if (b.type === 'letterMetaRecipient') {
      const halRow = (b.metaRows || []).find(r => /perihal|hal/i.test(r.label || ''));
      if (halRow?.value && !perihalText) perihalText = cleanSpaces(halRow.value);
      for (const l of (b.recipientLines || [])) {
        const line = cleanSpaces(l);
        if (/^yth\.?/i.test(line)) {
          const c = cleanSpaces(line.replace(/^yth\.?\s*/i, ''));
          if (isMeaningfulPersonOrRecipient(c)) ythRecipients.push(c);
        }
      }
      continue;
    }

    if (b.type === 'letterSignature') {
      if (b.date) signatureDate = cleanSpaces(b.date);
      if (b.name) signerName = cleanSpaces(b.name);
      continue;
    }

    if (b.type === 'paragraph' || b.type === 'heading') {
      const txt = extractTextFromBlock(b);
      if (txt) {
        allParagraphs.push(txt);
        if (!titleText && txt.length <= 75 && /^surat\s+/i.test(txt)) {
          titleText = txt;
        }
        if (!nomorText && /^nomor\s*:/i.test(txt)) {
          nomorText = cleanSpaces(txt.replace(/^nomor\s*:\s*/i, ''));
        }
      }
      continue;
    }

    if (b.type === 'list' && Array.isArray(b.items)) {
      b.items.forEach((it) => {
        const txt = Array.isArray(it) ? cleanSpaces(it.map(r => r?.text || '').join('')) : cleanSpaces(it?.text || it || '');
        if (txt) allParagraphs.push(txt);
      });
      continue;
    }

    if (b.type === 'identityTable' && Array.isArray(b.rows)) {
      const prevTxt = allParagraphs[allParagraphs.length - 1] || '';
      const isSignatoryBlock = SIGNATORY_INTRO_RE.test(prevTxt) || b.rows.some(r =>
        /jabatan/i.test(r.label || '') && /kepala\s+sekolah|kepala\s+smk|kepala\s+smp/i.test(r.value || '')
      );
      if (!isSignatoryBlock || TARGET_INTRO_RE.test(prevTxt)) {
        const nameRow = b.rows.find(r => /\bnama\b/i.test(r.label || ''));
        if (nameRow?.value) {
          const val = cleanSpaces(String(nameRow.value).replace(/^:\s*/, ''));
          if (isMeaningfulPersonOrRecipient(val) && !isSignatoryPerson(val, signatoryTokens)) {
            identityCandidates.push(val);
          }
        }
      }
      continue;
    }

    if (b.type === 'table' && Array.isArray(b.rows)) {
      // 1. Cek tabel daftar nama peserta / guru yang ditugaskan (kolom "Nama")
      const firstRow = b.rows[0];
      const headerCells = (Array.isArray(firstRow) ? firstRow : (firstRow?.cells || [])).map(extractCellText);
      const secondRow = b.rows[1];
      const secondCells = (Array.isArray(secondRow) ? secondRow : (secondRow?.cells || [])).map(extractCellText);
      const nameColIdx = findAssigneeNameColumnIndex(headerCells, secondCells);
      if (nameColIdx >= 0 && b.rows.length > 1) {
        for (let rIdx = 1; rIdx < b.rows.length; rIdx++) {
          const rowCells = Array.isArray(b.rows[rIdx]) ? b.rows[rIdx] : (b.rows[rIdx]?.cells || []);
          const cellVal = extractCellText(rowCells[nameColIdx]);
          if (isMeaningfulPersonOrRecipient(cellVal) && !isSignatoryPerson(cellVal, signatoryTokens)) {
            assigneeNames.push(cellVal);
          }
        }
        continue;
      }

      // 2. Cek tabel key-value identitas atau SPPD (mis. "Pegawai yang diberi perintah", "Maksud", "Tempat tujuan")
      const prevTxt = allParagraphs[allParagraphs.length - 1] || '';
      const tableKeyValSignatory = SIGNATORY_INTRO_RE.test(prevTxt) || b.rows.some((r) => {
        const cs = Array.isArray(r) ? r : (r?.cells || []);
        const l = extractCellText(cs[0]).toLowerCase();
        const v = extractCellText(cs[1] || cs[cs.length - 1]).toLowerCase();
        return l.includes('jabatan') && /kepala\s+(sekolah|smk|smp)/i.test(v);
      });

      for (const r of b.rows) {
        const cs = Array.isArray(r) ? r : (r?.cells || []);
        if (cs.length < 2) continue;
        const label = extractCellText(cs[0]);
        const value = cleanSpaces(cs.slice(1).map(c => extractCellText(c)).join(' ')).replace(/^:\s*/, '');
        if (!value) continue;

        if (/pegawai yang diberi perintah|nama\s+(?:pegawai|guru|siswa)/i.test(label) && isMeaningfulPersonOrRecipient(value)) {
          if (!isSignatoryPerson(value, signatoryTokens)) sppdMeta.pegawai = sppdMeta.pegawai || value;
        } else if (/^\s*(?:\d+\.?\s*)?nama\b/i.test(label) && !tableKeyValSignatory) {
          if (isMeaningfulPersonOrRecipient(value) && !isSignatoryPerson(value, signatoryTokens)) {
            identityCandidates.push(value);
          }
        }
        if (/maksud\s+perjalanan|keperluan/i.test(label) && value.length > 3) {
          sppdMeta.maksud = sppdMeta.maksud || value;
        } else if (/tempat\s+tujuan|lokasi\s+tujuan/i.test(label) && value.length > 2) {
          sppdMeta.tempat = sppdMeta.tempat || value;
        } else if (/tanggal\s+berangkat/i.test(label) && value.length > 3) {
          sppdMeta.tglBerangkat = sppdMeta.tglBerangkat || value;
        }
      }
    }
  }

  return {
    titleText,
    nomorText,
    perihalText,
    signatureDate,
    signerName,
    allParagraphs,
    assigneeNames,
    identityCandidates,
    ythRecipients,
    sppdMeta
  };
}

function collectFromSmartPagesHtml(pages = [], schoolProfile = {}) {
  const htmlCombined = (Array.isArray(pages) ? pages : []).map(p => p?.html || '').join('\n');
  if (!htmlCombined.trim()) return null;

  const lines = normalizeTextLinesFromHtml(htmlCombined);
  const allParagraphs = [];
  let titleText = '';
  let nomorText = '';
  let perihalText = '';
  let signatureDate = '';
  const assigneeNames = [];
  const identityCandidates = [];
  const ythRecipients = [];
  const signatoryTokens = buildSignatoryRootTokens(schoolProfile);

  // Ekstrak sel <tr>...</tr> untuk tabel nama / identitas
  const trMatches = [...htmlCombined.matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/gi)];
  let activeNameColIdx = -1;
  let tableSkipSignatory = false;

  // Reset per tabel
  const tableMatches = [...htmlCombined.matchAll(/<table\b[^>]*>([\s\S]*?)<\/table>/gi)];
  for (const tm of tableMatches) {
    const tHtml = tm[1];
    const tLower = stripHtmlTags(tHtml).toLowerCase();
    const isMetaTbl = /perihal|hal\s*:/i.test(tLower) && /kepada|yth/i.test(tLower);
    const isSignatoryTbl = /jabatan\s*:?\s*kepala\s+(?:sekolah|smk|smp)/i.test(tLower);
    const rows = [...tHtml.matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/gi)];
    activeNameColIdx = -1;
    tableSkipSignatory = isSignatoryTbl;

    const parsedRows = rows.map(rm =>
      [...rm[1].matchAll(/<(?:td|th)\b[^>]*>([\s\S]*?)<\/(?:td|th)>/gi)].map(cm => stripHtmlTags(cm[1]))
    );
    parsedRows.forEach((cells, rIdx) => {
      if (cells.length === 0) return;

      if (rIdx === 0) {
        const nextCells = parsedRows[1] || [];
        const candidateIdx = findAssigneeNameColumnIndex(cells, nextCells);
        if (candidateIdx >= 0) {
          activeNameColIdx = candidateIdx;
          return;
        }
      }

      if (activeNameColIdx >= 0 && cells[activeNameColIdx]) {
        const candidate = cells[activeNameColIdx];
        if (isMeaningfulPersonOrRecipient(candidate) && !isSignatoryPerson(candidate, signatoryTokens)) {
          assigneeNames.push(candidate);
        }
        return;
      }

      if (isMetaTbl) {
        const rawJoined = cells.join('\n');
        const halM = rawJoined.match(/(?:perihal|hal)\s*:\s*([^\n|]+)/i);
        if (halM && !perihalText) perihalText = cleanSpaces(halM[1]);
        const ythM = rawJoined.match(/yth\.?\s*([^\n|]+)/i);
        if (ythM) {
          const val = cleanSpaces(ythM[1]);
          if (isMeaningfulPersonOrRecipient(val)) ythRecipients.push(val);
        }
        return;
      }

      if (!tableSkipSignatory && cells.length >= 2) {
        const lbl = cells[0];
        const val = cleanSpaces(cells.slice(1).join(' ')).replace(/^:\s*/, '');
        if (/^\s*(?:\d+\.?\s*)?nama\b/i.test(lbl) && isMeaningfulPersonOrRecipient(val) && !isSignatoryPerson(val, signatoryTokens)) {
          identityCandidates.push(val);
        } else if (/maksud\s+perjalanan/i.test(lbl) && !perihalText && val.length > 3) {
          perihalText = val;
        }
      }
    });
  }

  const lineMeta = extractMetadataFromLines(lines);
  allParagraphs.push(...lines);
  if (!titleText) titleText = lineMeta.titleText;
  if (!nomorText) nomorText = lineMeta.nomorText;
  if (!perihalText) perihalText = lineMeta.perihalText;
  if (!signatureDate) signatureDate = lineMeta.signatureDate;
  ythRecipients.push(...lineMeta.ythRecipients);

  // Tangkap juga header h2 title eksplisit bila ada
  const hTitleMatch = htmlCombined.match(/<h[1-3][^>]*class=["'][^"']*smart-doc-title[^"']*["'][^>]*>([\s\S]*?)<\/h[1-3]>/i);
  if (hTitleMatch) titleText = stripHtmlTags(hTitleMatch[1]) || titleText;
  const pNomorMatch = htmlCombined.match(/<p[^>]*class=["'][^"']*smart-doc-nomor[^"']*["'][^>]*>([\s\S]*?)<\/p>/i);
  if (pNomorMatch) {
    const nVal = stripHtmlTags(pNomorMatch[1]).replace(/^Nomor\s*:\s*/i, '').trim();
    if (nVal && nVal !== '-') nomorText = nVal;
  }

  return {
    titleText,
    nomorText,
    perihalText,
    signatureDate,
    allParagraphs,
    assigneeNames,
    identityCandidates,
    ythRecipients,
    trCount: trMatches.length
  };
}

function extractMetadataFromLines(lines = []) {
  let titleText = '';
  let nomorText = '';
  let perihalText = '';
  let signatureDate = '';
  const ythRecipients = [];

  for (const line of lines) {
    if (!line) continue;
    if (!titleText && line.length <= 75 && /^(?:surat\s+|berita\s+acara\b|pengumuman\b|rekomendasi\b|dispensasi\b)/i.test(line)) {
      titleText = line;
    }
    if (!nomorText && /^nomor\s*:/i.test(line)) {
      nomorText = cleanSpaces(line.replace(/^nomor\s*:\s*/i, ''));
    }
    if (!perihalText && /^(?:perihal|hal)\s*:/i.test(line)) {
      perihalText = cleanSpaces(line.replace(/^(?:perihal|hal)\s*:\s*/i, ''));
    }
    if (/(?:^|\b)(?:kepada\s+)?yth\.?\s+/i.test(line)) {
      const val = cleanSpaces(line.replace(/^.*?(?:kepada\s+)?yth\.?\s*/i, ''));
      if (isMeaningfulPersonOrRecipient(val)) ythRecipients.push(val);
    }
    if (!signatureDate && /(?:garut|bandung|jakarta|cibiuk|tasikmalaya|[a-z]+)\s*,\s*\d{1,2}\s+[a-z]+\s+20\d{2}/i.test(line)) {
      const dm = line.match(/(\d{1,2}\s+[a-z]+\s+20\d{2})/i);
      if (dm) signatureDate = dm[1];
    }
  }

  return { titleText, nomorText, perihalText, signatureDate, ythRecipients };
}

function resolveLetterTypeCodeSmart(rawTitle = '', nomorSurat = '', detectionBlob = '') {
  const heading = cleanSpaces(rawTitle).toLowerCase();
  if (/\bsurat\s+keterangan\b/.test(heading)) return 'KET';
  if (/\bsurat\s+tugas\b|\bsppd\b/.test(heading)) return 'ST';
  if (/\bsurat\s+undangan\b|\bundangan\b/.test(heading)) return 'UND';
  if (/\bsurat\s+keputusan\b/.test(heading)) return 'KEP';
  if (/\bsurat\s+panggilan\b/.test(heading)) return 'PGL';
  if (/\bsurat\s+pemberitahuan\b|\bedaran\b/.test(heading)) return 'PBR';
  if (/\bsurat\s+permohonan\b/.test(heading)) return 'PRM';
  if (/\bsurat\s+dispensasi\b/.test(heading)) return 'DISP';
  if (/\bberita\s+acara\b/.test(heading)) return 'BA';
  if (/\bsurat\s+rekomendasi\b/.test(heading)) return 'REK';
  if (/\bsurat\s+(?:pkl|prakerin)\b/.test(heading)) return 'PKL';

  const nomorMatch = String(nomorSurat || '').match(/(?:^|\/)(UND|TU|KEP|KET|ST|PGL|PBR|PRM|DISP|BA|REK|PKL|KUR|KES)(?:\.|\/)/i);
  if (nomorMatch) return nomorMatch[1].toUpperCase();

  return detectLetterTypeCode(detectionBlob || 'TU');
}

function extractPurposeFromBodyParagraphs(paragraphs = [], judulDasar = '') {
  const joined = paragraphs.join('\n');

  // 1. "Untuk Penjemputan Peserta Didik PKL Bertempat di ..."
  const untukMatch = joined.match(/\buntuk\s+((?:melaksanakan|mengikuti|menghadiri|penjemputan|pengantaran|pendampingan|monitoring|evaluasi|koordinasi|kegiatan)[^.\n]{4,110})/i)
    || joined.match(/\buntuk\s+([A-Z][^.\n]{4,100})/);
  if (untukMatch) {
    let val = cleanSpaces(untukMatch[1]);
    val = val.replace(/\s+pada\s+hari\b.*$/i, '').replace(/\s+demikian\b.*$/i, '').trim();
    if (val.length >= 6) return normalizeHeadlineCase(val);
  }

  // 2. "dalam rangka ..." / "mengikuti kegiatan ..."
  const rangkaMatch = joined.match(/\b(?:dalam\s+rangka|acara|kegiatan)\s*:?\s*([^.\n]{5,95})/i);
  if (rangkaMatch) {
    const val = cleanSpaces(rangkaMatch[1]).replace(/\s+yang\s+insya\s+allah.*$/i, '').trim();
    if (val.length >= 5) return normalizeHeadlineCase(val);
  }

  // 3. "sebagai syarat ..." (umumnya di Surat Keterangan)
  const syaratMatch = joined.match(/\b(?:sebagai\s+syarat|untuk\s+keperluan|dipergunakan\s+(?:untuk|sebagai))\s+(?:administrasi\s+)?([^.\n]{5,90})/i);
  if (syaratMatch) {
    return normalizeHeadlineCase(cleanSpaces(syaratMatch[1]));
  }

  // 4. Dari judul spesifik mis. "SURAT KETERANGAN PENERIMA BANTUAN" -> "Penerima Bantuan"
  if (judulDasar && !GENERIC_TITLES_RE.test(judulDasar)) {
    const stripped = cleanSpaces(
      judulDasar
        .replace(/^surat\s+(?:keterangan|tugas|undangan|panggilan|keputusan|pemberitahuan|permohonan|dispensasi|rekomendasi)\s*/i, '')
    );
    if (stripped.length >= 4 && !/^siswa$/i.test(stripped)) {
      return normalizeHeadlineCase(stripped);
    }
  }

  // 5. Kalimat substantif pertama non-boilerplate
  const subst = paragraphs.find((p) =>
    p.length >= 18
    && !SIGNATORY_INTRO_RE.test(p)
    && !/^demikian\s+surat\s+ini/i.test(p)
    && !/^yang\s+bertanda\s+tangan/i.test(p)
    && !/^nomor\s*:/i.test(p)
    && !/^garut\s*,/i.test(p)
  );
  if (subst) {
    return cleanSpaces(subst.split(/[.!?]/)[0]).slice(0, 95);
  }

  return '';
}

function extractInlineRecipientFromParagraphs(paragraphs = [], signatoryTokens = []) {
  const joined = paragraphs.join('\n');

  const assignedMatch = joined.match(/(?:memberikan\s+tugas\s+kepada|menugaskan\s+(?:kepada\s+)?|memanggil\s+(?:kepada\s+)?)\s*([A-Z][^\n:.,]{2,60})/i);
  if (assignedMatch) {
    const candidate = cleanSpaces(assignedMatch[1]).replace(/\s+(?:untuk|agar|supaya|sebagai|pada)\b.*$/i, '').trim();
    if (isMeaningfulPersonOrRecipient(candidate) && !isSignatoryPerson(candidate, signatoryTokens)) {
      return candidate;
    }
  }

  const namaInline = joined.match(/\bNama\s*(?:Siswa|Guru|Pegawai)?\s*:\s*([^\n|]{2,70})/i);
  if (namaInline) {
    const val = cleanSpaces(namaInline[1]);
    if (isMeaningfulPersonOrRecipient(val) && !isSignatoryPerson(val, signatoryTokens)) {
      return val;
    }
  }

  return '';
}

function extractDateFromText(texts = []) {
  for (const t of texts) {
    if (!t) continue;
    const norm = normalizeIndonesianDateText(t);
    if (norm.full && norm.full !== cleanSpaces(t)) return norm;
  }
  return { full: '', short: '' };
}

function cleanFilenameSlug(str = '', maxLen = 28) {
  const slug = cleanSpaces(str)
    .replace(/&/g, 'dan')
    .replace(/[^a-zA-Z0-9\s\-]/g, ' ')
    .trim()
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-');
  if (slug.length <= maxLen) return slug;
  const cut = slug.slice(0, maxLen).replace(/-[^-]*$/, '');
  return cut || slug.slice(0, maxLen);
}

/**
 * Mengekstrak seluruh metadata rapih dari dokumen secara deterministik (Format Sistem Instan).
 * Mendukung DocumentModel, halaman Smart Word Pro, dan berkas .docx impor.
 */
export function extractDocumentMetadata({
  docModel = null,
  smartPages = null,
  uploadedDocx = null,
  schoolProfile = null,
  fallbackTitle = '',
  fallbackNomor = '',
  fallbackDate = '',
  fallbackPerihal = '',
  fallbackRecipient = '',
  forTemplate = false
} = {}) {
  const school = schoolProfile || resolveDocumentSchoolProfile(docModel || {}, {}, []);
  const signatoryTokens = buildSignatoryRootTokens(school);

  const modelData = docModel ? collectFromDocModel(docModel, school) : null;
  const pageData = Array.isArray(smartPages) && smartPages.length > 0
    ? collectFromSmartPagesHtml(smartPages, school)
    : null;

  const docxParas = Array.isArray(uploadedDocx?.paragraphs)
    ? uploadedDocx.paragraphs.map(p => cleanSpaces(p?.currentText ?? p?.text ?? p ?? '')).filter(Boolean)
    : [];
  const docxLineMeta = docxParas.length > 0 ? extractMetadataFromLines(docxParas) : null;
  const docxMeta = uploadedDocx?.parsedMeta || {};
  const docxStructured = uploadedDocx?.structured || {};

  const combinedParagraphs = [
    ...(modelData?.allParagraphs || []),
    ...(pageData?.allParagraphs || []),
    ...docxParas
  ];

  // 1. Raw heading & nomor surat
  const safeFallbackTitle = /_[a-z0-9]+|\(\d+\)/i.test(fallbackTitle || '') ? '' : fallbackTitle;
  let rawTitle = cleanSpaces(
    modelData?.titleText
    || pageData?.titleText
    || docxStructured?.title?.text
    || docxLineMeta?.titleText
    || docxMeta?.title
    || safeFallbackTitle
    || docModel?.meta?.title
    || ''
  );
  if (GENERIC_TITLES_RE.test(rawTitle)) rawTitle = '';

  const nomorSurat = cleanSpaces(
    modelData?.nomorText
    || pageData?.nomorText
    || (docxStructured?.nomor?.text || '').replace(/^Nomor\s*:\s*/i, '')
    || docxLineMeta?.nomorText
    || docxMeta?.nomor
    || fallbackNomor
    || ''
  );

  // 2. Deteksi Kode Jenis Surat
  const detectionBlob = [
    rawTitle,
    nomorSurat,
    modelData?.perihalText,
    pageData?.perihalText,
    docxLineMeta?.perihalText,
    combinedParagraphs.slice(0, 6).join(' ')
  ].filter(Boolean).join(' ');

  const isSppdDoc = /\bSPPD\b|surat\s+perintah\s+perjalanan\s+dinas/i.test(
    `${detectionBlob} ${combinedParagraphs.join(' ')}`
  );
  const typeCode = resolveLetterTypeCodeSmart(rawTitle, nomorSurat, detectionBlob);
  const baseJenisLabel = TYPE_TO_HUMAN_LABEL[typeCode] || LETTER_TYPE_CODES[typeCode]?.label?.replace(/\s*\([^)]+\)$/, '') || 'Surat Dinas';

  const jenisSurat = rawTitle && !GENERIC_TITLES_RE.test(rawTitle)
    ? normalizeHeadlineCase(rawTitle)
    : (isSppdDoc && typeCode === 'ST' ? 'Surat Tugas & SPPD' : baseJenisLabel);

  // 3. Atas Nama / Penerima / Pihak Ditugaskan
  const directAssignees = [
    ...(modelData?.assigneeNames || []),
    ...(pageData?.assigneeNames || [])
  ];
  const sppdPegawaiSplit = directAssignees.length === 0 && modelData?.sppdMeta?.pegawai
    ? modelData.sppdMeta.pegawai.split(/\s*[\/;]\s*/).map(s => cleanSpaces(s)).filter(Boolean)
    : [];
  const mergedAssignees = [
    ...directAssignees,
    ...sppdPegawaiSplit,
    ...(modelData?.identityCandidates || []),
    ...(pageData?.identityCandidates || [])
  ];

  let recipientName = formatPeopleList(mergedAssignees);
  if (!recipientName) {
    const ythList = [
      ...(modelData?.ythRecipients || []),
      ...(pageData?.ythRecipients || []),
      ...(docxLineMeta?.ythRecipients || [])
    ];
    recipientName = formatPeopleList(ythList);
  }
  if (!recipientName) {
    recipientName = extractInlineRecipientFromParagraphs(combinedParagraphs, signatoryTokens);
  }
  if (!recipientName && isMeaningfulPersonOrRecipient(fallbackRecipient) && !isSignatoryPerson(fallbackRecipient, signatoryTokens)) {
    recipientName = cleanSpaces(fallbackRecipient);
  }
  if (!recipientName) {
    recipientName = forTemplate ? '' : 'Umum';
  }

  // 4. Keperluan / Perihal
  let keperluan = cleanSpaces(
    modelData?.perihalText
    || pageData?.perihalText
    || docxLineMeta?.perihalText
    || modelData?.sppdMeta?.maksud
    || ''
  );
  if (keperluan && modelData?.sppdMeta?.tempat && !keperluan.toLowerCase().includes(modelData.sppdMeta.tempat.toLowerCase())) {
    keperluan = `${keperluan} (${modelData.sppdMeta.tempat})`;
  }
  if (!keperluan || GENERIC_TITLES_RE.test(keperluan) || keperluan.toLowerCase() === jenisSurat.toLowerCase()) {
    keperluan = extractPurposeFromBodyParagraphs(combinedParagraphs, rawTitle)
      || (modelData?.sppdMeta?.tempat ? `Perjalanan Dinas ke ${modelData.sppdMeta.tempat}` : '')
      || cleanSpaces(fallbackPerihal)
      || cleanSpaces(fallbackTitle);
  }
  keperluan = normalizeHeadlineCase(keperluan);
  if (keperluan.length > 95) {
    keperluan = `${keperluan.slice(0, 92).trim()}...`;
  }

  // 5. Tanggal Surat
  const dateObj = extractDateFromText([
    modelData?.signatureDate,
    pageData?.signatureDate,
    docxLineMeta?.signatureDate,
    modelData?.sppdMeta?.tglBerangkat,
    docxMeta?.date,
    fallbackDate,
    ...combinedParagraphs.slice(-10)
  ]);
  const todayFallback = new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
  const tanggalSurat = dateObj.full || normalizeIndonesianDateText(todayFallback).full || todayFallback;
  const shortDate = dateObj.short || normalizeIndonesianDateText(tanggalSurat).short;

  // 6. Rangkai Judul Rapih (tanpa pengulangan antar segmen)
  const titleParts = [jenisSurat || baseJenisLabel];
  if (recipientName && recipientName.toLowerCase() !== 'umum' && !PLACEHOLDER_VALUE_RE.test(recipientName)) {
    const recShort = recipientName.length > 42 ? `${recipientName.slice(0, 39).trim()}...` : recipientName;
    if (!titleParts[0].toLowerCase().includes(recShort.toLowerCase())) {
      titleParts.push(recShort);
    }
  }
  if (
    keperluan
    && !titleParts.some(p => p.toLowerCase() === keperluan.toLowerCase() || p.toLowerCase().includes(keperluan.toLowerCase()))
  ) {
    const kepShort = keperluan.length > 52 ? `${keperluan.slice(0, 49).trim()}...` : keperluan;
    titleParts.push(kepShort);
  }

  let formattedTitle = titleParts.join(' — ');
  if (shortDate && !forTemplate) {
    formattedTitle = `${formattedTitle} (${shortDate})`;
  }

  const unitIdResolved = normalizeUnitId(school?.unitId || docModel?.targetUnitId || 'smk');
  const unitBadge = SCHOOL_UNITS[unitIdResolved]?.unitShortBadge || 'SMK';
  const category = TYPE_TO_VAULT_CATEGORY[typeCode] || 'Surat Dinas';
  const summaryParts = [];
  if (keperluan) summaryParts.push(`Keperluan: ${keperluan}.`);
  if (recipientName && recipientName !== 'Umum') summaryParts.push(`Pihak: ${recipientName}.`);
  if (tanggalSurat) summaryParts.push(`Tanggal: ${tanggalSurat}.`);
  const summary = summaryParts.join(' ').slice(0, 220) || 'Dokumen persuratan resmi sekolah.';

  return {
    typeCode,
    jenisSurat: jenisSurat || baseJenisLabel,
    category,
    rawTitle: rawTitle || jenisSurat.toUpperCase(),
    formattedTitle,
    perihal: keperluan || jenisSurat || baseJenisLabel,
    keperluan: keperluan || baseJenisLabel,
    recipientName: recipientName || 'Umum',
    tanggalSurat,
    shortDate,
    nomorSurat: nomorSurat || '-',
    unitBadge,
    summary
  };
}

/**
 * Menyusun judul terstruktur dari input manual atau hasil AI agar tetap konsisten dan rapih.
 */
export function buildFormattedDocumentTitle({
  typeCode = '',
  jenisSurat = '',
  recipientName = '',
  keperluan = '',
  shortDate = '',
  tanggalSurat = '',
  forTemplate = false
} = {}) {
  const baseLabel = cleanSpaces(jenisSurat)
    || TYPE_TO_HUMAN_LABEL[String(typeCode || '').toUpperCase()]
    || 'Surat Dinas';
  const parts = [normalizeHeadlineCase(baseLabel)];

  const cleanRec = cleanSpaces(recipientName);
  if (cleanRec && cleanRec.toLowerCase() !== 'umum' && !PLACEHOLDER_VALUE_RE.test(cleanRec)) {
    const recShort = cleanRec.length > 42 ? `${cleanRec.slice(0, 39).trim()}...` : cleanRec;
    if (!parts[0].toLowerCase().includes(recShort.toLowerCase())) {
      parts.push(recShort);
    }
  }

  const cleanKep = cleanSpaces(keperluan);
  if (cleanKep && !parts.some(p => p.toLowerCase().includes(cleanKep.toLowerCase()))) {
    const kepShort = cleanKep.length > 52 ? `${cleanKep.slice(0, 49).trim()}...` : cleanKep;
    parts.push(normalizeHeadlineCase(kepShort));
  }

  const dateTag = cleanSpaces(shortDate) || normalizeIndonesianDateText(tanggalSurat).short;
  return !forTemplate && dateTag
    ? `${parts.join(' — ')} (${dateTag})`
    : parts.join(' — ');
}

/**
 * Menghasilkan nama berkas .docx yang bersih, deskriptif, dan aman untuk Windows/Drive/Linux.
 */
export function buildDocumentDownloadFilename({
  kind = 'vault', // 'outgoing' | 'vault'
  sequenceNumber = 0,
  typeCode = '',
  title = '',
  perihal = '',
  recipientName = '',
  tanggalSurat = '',
  unitId = '',
  schoolProfile = null
} = {}) {
  const seqPrefix = kind === 'outgoing' && Number(sequenceNumber) > 0
    ? String(Number(sequenceNumber)).padStart(3, '0')
    : '';
  const code = String(typeCode || detectLetterTypeCode(`${title} ${perihal}`)).toUpperCase().slice(0, 6) || 'DOC';
  const unitTag = SCHOOL_UNITS[normalizeUnitId(unitId || schoolProfile?.unitId || 'smk')]?.unitShortBadge || 'SMK';

  // Ambil inti jenis/subjek dari title atau perihal
  const titleSegments = cleanSpaces(title).split(/\s*[—–]\s*/).filter(Boolean);
  const jenisSeg = titleSegments[0] || TYPE_TO_HUMAN_LABEL[code] || 'Surat';
  const jenisSlug = cleanFilenameSlug(jenisSeg, 24) || code;
  const orangRaw = (recipientName && recipientName !== '-' && !/^umum/i.test(recipientName))
    ? recipientName
    : (titleSegments[1] || '');
  const orangSlug = cleanFilenameSlug(orangRaw, 24);
  const kepRaw = cleanSpaces(perihal).toLowerCase() !== cleanSpaces(jenisSeg).toLowerCase()
    ? perihal
    : (titleSegments[2] || '');
  const kepSlug = cleanFilenameSlug(kepRaw, 26);
  const dateObj = normalizeIndonesianDateText(tanggalSurat);
  const dateSlug = cleanFilenameSlug(dateObj.short || '', 14);

  const chunks = [];
  if (seqPrefix) chunks.push(seqPrefix);
  chunks.push(kind === 'outgoing' ? code : jenisSlug);
  if (orangSlug && orangSlug.toLowerCase() !== jenisSlug.toLowerCase()) chunks.push(orangSlug);
  if (kepSlug && kepSlug.toLowerCase() !== jenisSlug.toLowerCase() && kepSlug.toLowerCase() !== orangSlug.toLowerCase()) {
    chunks.push(kepSlug);
  }
  if (dateSlug) chunks.push(dateSlug);
  if (unitTag) chunks.push(unitTag);

  const joined = chunks.filter(Boolean).join('_').slice(0, 88).replace(/_+$/g, '');
  return `${joined || 'Dokumen_Surat_Sekolah'}.docx`;
}

/**
 * Memindai dan merapikan metadata surat memakai AI aktif (jika API key tersedia),
 * digabung dengan validasi format sistem agar hasilnya terstruktur & dapat diandalkan.
 */
export async function enhanceDocumentMetadataWithAi({
  docModel = null,
  smartPages = null,
  uploadedDocx = null,
  schoolProfile = null,
  forTemplate = false,
  signal = null
} = {}) {
  const sysMeta = extractDocumentMetadata({
    docModel,
    smartPages,
    uploadedDocx,
    schoolProfile,
    forTemplate
  });

  const aiCfg = getActiveAiConfig();
  if (!aiCfg?.apiKey) {
    return { ...sysMeta, scannedByAi: false };
  }

  const lines = [];
  if (docModel) {
    const m = collectFromDocModel(docModel, schoolProfile || {});
    lines.push(...m.allParagraphs);
    if (m.assigneeNames.length > 0) lines.push(`Daftar Ditugaskan: ${m.assigneeNames.join(', ')}`);
    if (m.identityCandidates.length > 0) lines.push(`Identitas Subjek: ${m.identityCandidates.join(', ')}`);
    if (m.sppdMeta.maksud || m.sppdMeta.tempat) {
      lines.push(`SPPD Maksud: ${m.sppdMeta.maksud}; Tujuan: ${m.sppdMeta.tempat}; Tanggal: ${m.sppdMeta.tglBerangkat}`);
    }
  }
  if (Array.isArray(smartPages)) {
    smartPages.forEach((p, idx) => {
      const txt = stripHtmlTags(p?.html || '');
      if (txt) lines.push(`[Hal ${idx + 1}] ${txt}`);
    });
  }
  if (Array.isArray(uploadedDocx?.paragraphs)) {
    uploadedDocx.paragraphs.slice(0, 35).forEach((p) => {
      const t = cleanSpaces(p?.currentText ?? p?.text ?? p ?? '');
      if (t) lines.push(t);
    });
  }

  const compactText = lines.join('\n').slice(0, 3200);
  if (!compactText.trim()) {
    return { ...sysMeta, scannedByAi: false };
  }

  const systemPrompt = `
Kamu asisten administrasi tata usaha sekolah Indonesia. Tugasmu mengekstrak metadata penamaan surat agar rapi dan konsisten.
Kembalikan HANYA objek JSON valid tanpa markdown:
{
  "typeCode": "ST | KET | UND | KEP | PGL | PBR | PRM | DISP | BA | REK | PKL | KUR | KES | TU",
  "jenisSurat": "Misal: Surat Tugas / Surat Keterangan Penerima Bantuan / Surat Undangan",
  "recipientName": "Nama orang/guru/siswa yang ditugaskan, diterangkan, atau dituju (JANGAN tulis nama Kepala Sekolah pemberi perintah jika surat adalah Surat Tugas/Penugasan untuk guru/siswa; jika lebih dari 2 orang tulis format: Nama Pertama dkk (N orang))",
  "keperluan": "Keperluan / perihal / lokasi kegiatan singkat & jelas (maks 8 kata)",
  "tanggalSurat": "Tanggal surat format DD Bulan YYYY (misal: 23 November 2026)"
}
  `.trim();

  const prompt = `Analisis teks surat berikut dan ekstrak metadata penamaannya:\n\n${compactText}\n\nReferensi awal sistem:\n${JSON.stringify({
    typeCode: sysMeta.typeCode,
    jenisSurat: sysMeta.jenisSurat,
    recipientName: sysMeta.recipientName,
    keperluan: sysMeta.keperluan,
    tanggalSurat: sysMeta.tanggalSurat
  })}`;

  try {
    const raw = await callAiApi({
      prompt,
      systemPrompt,
      signal
    });
    const parsed = JSON.parse(cleanJsonOutput(raw) || '{}');

    const nextCode = LETTER_TYPE_CODES[String(parsed.typeCode || '').toUpperCase()]
      ? String(parsed.typeCode).toUpperCase()
      : sysMeta.typeCode;
    const nextJenis = normalizeHeadlineCase(parsed.jenisSurat || sysMeta.jenisSurat);
    const rawRec = cleanSpaces(parsed.recipientName || '');
    const signatoryTokens = buildSignatoryRootTokens(schoolProfile || {});
    const nextRec = (isMeaningfulPersonOrRecipient(rawRec) && !(nextCode === 'ST' && isSignatoryPerson(rawRec, signatoryTokens)))
      ? rawRec
      : sysMeta.recipientName;
    const nextKep = normalizeHeadlineCase(parsed.keperluan || sysMeta.keperluan);
    const dateNorm = normalizeIndonesianDateText(parsed.tanggalSurat || sysMeta.tanggalSurat);
    const nextDate = dateNorm.full || sysMeta.tanggalSurat;
    const nextShortDate = dateNorm.short || sysMeta.shortDate;

    const formattedTitle = buildFormattedDocumentTitle({
      typeCode: nextCode,
      jenisSurat: nextJenis,
      recipientName: nextRec,
      keperluan: nextKep,
      shortDate: nextShortDate,
      tanggalSurat: nextDate,
      forTemplate
    });

    return {
      ...sysMeta,
      typeCode: nextCode,
      jenisSurat: nextJenis,
      category: TYPE_TO_VAULT_CATEGORY[nextCode] || sysMeta.category,
      recipientName: nextRec,
      keperluan: nextKep,
      perihal: nextKep || sysMeta.perihal,
      tanggalSurat: nextDate,
      shortDate: nextShortDate,
      formattedTitle,
      scannedByAi: true
    };
  } catch {
    return { ...sysMeta, scannedByAi: false };
  }
}
