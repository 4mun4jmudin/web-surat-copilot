/**
 * knowledgeBase.js - Mesin Pembelajar Template & Relevansi Gudang Surat untuk AI Copilot
 *
 * Menganalisis seluruh dokumen di Gudang Surat:
 * 1. Mengekstrak cetak biru struktur lengkap (judul, Sifat/Kepada, susunan paragraf, tabel rincian,
 *    tabel grid/SPPD beserta borderStyle/colWidthsPct/colspan/rowspan, pageBreak, & tanda tangan)
 *    baik dari DocumentModel kanvas maupun dokumen .docx unggahan.
 * 2. Mencocokkan permintaan pengguna terhadap dokumen di Gudang Surat dengan bobot relevansi,
 *    memprioritaskan dokumen bertanda Bintang Emas Terkunci (⭐ Master Template) sebagai acuan utama.
 * 3. Mengembalikan mode `'matched'` (duplikasi & adaptasi tata letak acuan terbaik) atau
 *    `'fresh'` (buat baru dari awal bila belum ada acuan sejenis di Gudang).
 */

import { detectLetterTypeCode } from '../letters/numbering.js';

const STOPWORDS = new Set([
  'buat', 'buatkan', 'tolong', 'mohon', 'minta', 'bikinkan', 'coba', 'surat',
  'dokumen', 'yang', 'untuk', 'kepada', 'dari', 'dan', 'atau', 'dengan',
  'pada', 'di', 'ke', 'dalam', 'sebagai', 'tentang', 'terkait', 'ini', 'itu',
  'agar', 'supaya', 'oleh', 'akan', 'resmi', 'sekolah', 'smk', 'smks', 'al', 'hawari',
  'kode', 'format', 'lampiran', 'halaman'
]);

function normalizeTokens(text = '') {
  return String(text || '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(/\s+/)
    .map(w => w.trim())
    .filter(w => w.length >= 2 && !STOPWORDS.has(w));
}

function truncateText(text = '', maxLen = 400) {
  const s = String(text || '').trim();
  if (s.length <= maxLen) return s;
  return `${s.slice(0, maxLen - 3)}...`;
}

function compactTableRowsForBlueprint(rows = [], maxRows = 10) {
  return (Array.isArray(rows) ? rows.slice(0, maxRows) : []).map((rowEntry) => {
    const isRowObj = rowEntry && typeof rowEntry === 'object' && !Array.isArray(rowEntry);
    const cells = isRowObj ? (rowEntry.cells || []) : (Array.isArray(rowEntry) ? rowEntry : []);
    const compactCells = cells.map((c) => {
      if (c && typeof c === 'object' && !Array.isArray(c)) {
        const txt = Array.isArray(c.paragraphs)
          ? c.paragraphs.map(p => p.text || '').filter(Boolean).join('\n')
          : String(c.text ?? '');
        const simpleObj = {
          text: truncateText(txt, 160),
          ...(c.align ? { align: c.align } : {}),
          ...(c.colspan > 1 ? { colspan: c.colspan } : {}),
          ...(c.rowspan > 1 ? { rowspan: c.rowspan } : {}),
          ...(c.bold ? { bold: true } : {})
        };
        return Object.keys(simpleObj).length === 1 ? simpleObj.text : simpleObj;
      }
      return truncateText(String(c ?? ''), 140);
    });
    if (isRowObj && ('isHeader' in rowEntry || rowEntry.minHeightCm)) {
      return {
        ...(rowEntry.isHeader !== undefined ? { isHeader: Boolean(rowEntry.isHeader) } : {}),
        ...(rowEntry.minHeightCm ? { minHeightCm: rowEntry.minHeightCm } : {}),
        cells: compactCells
      };
    }
    return compactCells;
  });
}

/**
 * Mengekstrak cetak biru struktur & tata letak utuh dari satu entri Gudang Surat untuk dipelajari AI.
 */
export function extractDocumentBlueprintForAi(entry = {}) {
  if (!entry || typeof entry !== 'object') return null;

  // 1. Jika entri berasal dari DocumentModel (Kanvas Word Pro / Lembar Kerja / Builder)
  const doc = entry.document;
  if (doc && Array.isArray(doc.blocks)) {
    const titleBlock = doc.blocks.find(b => b.type === 'letterTitle');
    const standaloneMeta = doc.blocks.find(b => b.type === 'letterMetaRecipient');
    const metaRecip = titleBlock?.metaRecipient || standaloneMeta || null;
    const sigBlock = doc.blocks.find(b => b.type === 'letterSignature');

    const blocksStructure = [];
    let hasTableSig = false;

    for (const b of doc.blocks) {
      if (!b || b.type === 'letterKop' || b.type === 'letterTitle' || b.type === 'letterMetaRecipient' || b.type === 'letterSignature') {
        continue;
      }
      if (b.type === 'pageBreak') {
        blocksStructure.push({ type: 'pageBreak' });
        continue;
      }
      if (b.type === 'heading') {
        blocksStructure.push({
          type: 'heading',
          text: truncateText(b.text || (b.runs || []).map(r => r.text).join(''), 200),
          ...(b.alignment ? { alignment: b.alignment } : {})
        });
        continue;
      }
      if (b.type === 'paragraph') {
        const txt = (b.runs ? b.runs.map(r => r.text).join('') : (b.text || '')).trim();
        if (!txt) continue;
        blocksStructure.push({
          type: 'paragraph',
          ...(b.alignment ? { alignment: b.alignment } : {}),
          ...(b.runs?.every(r => r.bold) ? { bold: true } : {}),
          text: truncateText(txt, 420)
        });
        continue;
      }
      if (b.type === 'identityTable' && Array.isArray(b.rows)) {
        blocksStructure.push({
          type: 'identityTable',
          ...(Array.isArray(b.colWidthsPct) ? { colWidthsPct: b.colWidthsPct } : {}),
          rows: b.rows.slice(0, 12).map(r => ({
            label: String(r.label || ''),
            value: truncateText(String(r.value || ''), 140)
          }))
        });
        continue;
      }
      if (b.type === 'table' && Array.isArray(b.rows)) {
        const flatStr = JSON.stringify(b.rows);
        if (flatStr.includes('qrSlot') || flatStr.includes('Kepala ')) {
          hasTableSig = true;
        }
        blocksStructure.push({
          type: 'table',
          borderStyle: b.borderStyle || 'grid',
          headerRow: b.headerRow !== false,
          ...(b.headerShading !== undefined ? { headerShading: b.headerShading } : {}),
          ...(Array.isArray(b.colWidthsPct) ? { colWidthsPct: b.colWidthsPct } : {}),
          rows: compactTableRowsForBlueprint(b.rows, 10)
        });
        continue;
      }
      if (b.type === 'list' && Array.isArray(b.items)) {
        blocksStructure.push({
          type: 'list',
          ordered: b.ordered !== false,
          items: b.items.slice(0, 10).map(it =>
            truncateText(Array.isArray(it) ? it.map(r => r.text).join('') : String(it), 180)
          )
        });
      }
    }

    const metaRows = metaRecip?.metaRows || [];
    const sifatRow = metaRows.find(r => /sifat/i.test(r.label || ''));
    const halRow = metaRows.find(r => /hal|perihal/i.test(r.label || ''));
    const tplId = doc.letterData?.templateId || '';

    return {
      sourceType: 'documentModel',
      ...(tplId === 'surat_tugas_sppd' ? { templatePreset: 'surat_tugas_sppd' } : {}),
      paperSize: doc.paperSize || 'a4',
      letterTitle: titleBlock?.title || entry.title || 'SURAT DINAS',
      kodeJenisSurat: detectLetterTypeCode(`${titleBlock?.title || entry.title} ${entry.category || ''} ${entry.perihal || ''}`),
      nomorPola: titleBlock?.nomor || entry.nomorSurat || '-',
      hideMetaRecipient: !metaRecip,
      ...(metaRecip ? {
        sifat: sifatRow?.value || 'Penting',
        hal: halRow?.value || entry.perihal || '',
        tujuan: (metaRecip.recipientLines || []).join(' / ')
      } : {}),
      blocksStructure: blocksStructure.slice(0, 18),
      hideSignature: Boolean(!sigBlock && hasTableSig),
      ...(sigBlock ? {
        signer: {
          city: sigBlock.city || 'Garut',
          title: sigBlock.title || 'Kepala Sekolah',
          name: sigBlock.name || 'Dian Nurhadian, S.Pd.'
        }
      } : {})
    };
  }

  // 2. Jika entri berasal dari berkas .docx impor yang diunggah pengguna
  const uploaded = entry.uploadedDocx;
  if (uploaded && uploaded.structured) {
    const st = uploaded.structured;
    const editedMap = new Map(
      (uploaded.paragraphs || []).map(p => [p.index, p.currentText !== undefined ? p.currentText : p.text])
    );
    const getLiveText = (node) => {
      if (!node) return '';
      if (node.index !== undefined && editedMap.has(node.index)) {
        return String(editedMap.get(node.index) || '').trim();
      }
      return String(node.text || '').trim();
    };

    const blocksStructure = [];
    for (const p of st.bodyParagraphs || []) {
      const txt = getLiveText(p);
      if (txt) blocksStructure.push({ type: 'paragraph', text: truncateText(txt, 420) });
    }
    if (Array.isArray(st.identityRows) && st.identityRows.length > 0) {
      blocksStructure.push({
        type: 'identityTable',
        rows: st.identityRows.slice(0, 12).map(r => ({
          label: r.label || '',
          value: `: ${String(r.value || '').replace(/^:\s*/, '')}`
        }))
      });
    }
    const closingTxt = getLiveText(st.closing);
    if (closingTxt) {
      blocksStructure.push({ type: 'paragraph', text: truncateText(closingTxt, 320) });
    }

    const metaRows = st.metaRows || [];
    const sifatRow = metaRows.find(r => /sifat/i.test(r.label || ''));
    const halRow = metaRows.find(r => /hal|perihal/i.test(r.label || ''));

    return {
      sourceType: 'uploadedDocx',
      paperSize: 'a4',
      letterTitle: getLiveText(st.title) || entry.title || 'SURAT DINAS',
      kodeJenisSurat: detectLetterTypeCode(`${entry.title} ${entry.category || ''} ${entry.perihal || ''}`),
      nomorPola: getLiveText(st.nomor) || entry.nomorSurat || '-',
      hideMetaRecipient: metaRows.length === 0 && !(st.recipientLines?.length),
      ...(metaRows.length > 0 ? {
        sifat: sifatRow?.value || 'Penting',
        hal: halRow?.value || entry.perihal || '',
        tujuan: (st.recipientLines || []).join(' / ')
      } : {}),
      blocksStructure: blocksStructure.slice(0, 16),
      signer: {
        city: st.signature?.city || 'Garut',
        title: st.signature?.title || 'Kepala Sekolah',
        name: st.signature?.signer || 'Dian Nurhadian, S.Pd.'
      }
    };
  }

  return null;
}

/**
 * Mengukur apakah jenis surat terdeteksi spesifik di dalam teks (bukan fallback default 'TU').
 */
function detectExplicitLetterCode(text = '') {
  const lower = String(text || '').toLowerCase();
  if (!lower.trim()) return '';
  if (/tata\s+usaha|\btu\b/.test(lower)) return 'TU';
  const code = detectLetterTypeCode(lower);
  return code === 'TU' ? '' : code;
}

/**
 * Memilih dan mengurutkan dokumen acuan dari seluruh Gudang Surat sesuai relevansinya terhadap prompt AI.
 * - Semua file aktif (`isAiReference !== false`) dapat dipelajari AI.
 * - Dokumen bertanda Bintang Emas Terkunci (`isMasterTemplate`) mendapat prioritas skor tertinggi
 *   jika relevan dengan jenis/topik surat yang diminta.
 * - Jika belum ada dokumen serupa di Gudang, mengembalikan `matchMode: 'fresh'` agar AI membuat format baru dari awal.
 */
export function selectRelevantAiReferences(userPrompt = '', vaultList = [], { limit = 3 } = {}) {
  const pool = (Array.isArray(vaultList) ? vaultList : []).filter(v => v && (v.isAiReference !== false || v.isMasterTemplate));
  const masterCount = pool.filter(v => Boolean(v.isMasterTemplate)).length;

  const query = String(userPrompt || '').trim();
  const promptTokens = normalizeTokens(query);
  const explicitPromptCode = detectExplicitLetterCode(query);
  const asksSppd = /\bsppd\b|perjalanan\s+dinas/i.test(query);

  if (!query) {
    const sorted = [...pool].sort((a, b) => {
      if (Boolean(b.isMasterTemplate) !== Boolean(a.isMasterTemplate)) {
        return b.isMasterTemplate ? 1 : -1;
      }
      return new Date(b.updatedAt || b.createdAt || 0) - new Date(a.updatedAt || a.createdAt || 0);
    });
    const picked = sorted.slice(0, limit).map(entry => formatMatchedRef(entry, entry.isMasterTemplate ? 100 : 10));
    return {
      matchMode: picked.length > 0 ? 'matched' : 'fresh',
      totalKnowledgeCount: pool.length,
      masterCount,
      topReference: picked[0] || null,
      references: picked
    };
  }

  const scored = [];
  for (const doc of pool) {
    const headerText = `${doc.title || ''} ${doc.category || ''} ${doc.perihal || ''}`;
    const bodyText = `${doc.summary || ''} ${doc.targetPerson || ''} ${extractSearchableBlockText(doc)}`;
    const headerTokens = new Set(normalizeTokens(headerText));
    const fullTokens = new Set([...headerTokens, ...normalizeTokens(bodyText)]);
    const fullLower = `${headerText} ${bodyText}`.toLowerCase();

    const explicitDocCode = detectExplicitLetterCode(headerText);
    const docCode = explicitDocCode || detectLetterTypeCode(headerText);

    // Jika user meminta jenis surat spesifik (mis. UND / ST / KET / PGL / KEP / BA) dan kandidat
    // secara eksplisit adalah jenis surat yang berbeda sama sekali, jangan paksakan jadi acuan
    if (explicitPromptCode && explicitDocCode && explicitPromptCode !== explicitDocCode) {
      if (!(asksSppd && /sppd|perjalanan\s+dinas/i.test(fullLower))) {
        continue;
      }
    }

    let lexicalScore = 0;

    if (explicitPromptCode && explicitPromptCode === docCode) {
      lexicalScore += 45;
    }
    if (asksSppd && /sppd|perjalanan\s+dinas/i.test(fullLower)) {
      lexicalScore += 45;
    }

    let tokenHits = 0;
    for (const tok of promptTokens) {
      if (headerTokens.has(tok)) {
        lexicalScore += 18;
        tokenHits++;
      } else if (fullTokens.has(tok)) {
        lexicalScore += 8;
        tokenHits++;
      }
    }

    if (lexicalScore < 18 && tokenHits === 0) {
      continue;
    }

    // Jika kecocokan cukup relevan, berikan bobot ekstra tinggi bagi Master Terkunci ⭐
    if (lexicalScore >= 18) {
      const masterBonus = doc.isMasterTemplate ? 100 : 0;
      const hasRichDocBonus = (doc.document?.blocks?.length || doc.uploadedDocx?.structured) ? 12 : 0;
      scored.push({
        entry: doc,
        score: lexicalScore + masterBonus + hasRichDocBonus
      });
    }
  }

  scored.sort((a, b) => b.score - a.score);
  const picked = scored.slice(0, limit).map(s => formatMatchedRef(s.entry, s.score));

  return {
    matchMode: picked.length > 0 ? 'matched' : 'fresh',
    totalKnowledgeCount: pool.length,
    masterCount,
    topReference: picked[0] || null,
    references: picked
  };
}

function extractSearchableBlockText(entry) {
  const blocks = entry?.document?.blocks || [];
  if (!blocks.length) return '';
  return blocks
    .slice(0, 12)
    .map((b) => {
      if (b.type === 'paragraph') return b.runs ? b.runs.map(r => r.text).join('') : (b.text || '');
      if (b.type === 'letterTitle') return b.title || '';
      if (b.type === 'identityTable') return (b.rows || []).map(r => `${r.label} ${r.value}`).join(' ');
      if (b.type === 'table') return (b.rows || []).slice(0, 4).map(r => JSON.stringify(r)).join(' ');
      return '';
    })
    .join(' ');
}

function formatMatchedRef(entry, score = 0) {
  return {
    id: entry.id,
    title: entry.title,
    nomorSurat: entry.nomorSurat,
    category: entry.category,
    targetPerson: entry.targetPerson,
    perihal: entry.perihal,
    summary: entry.summary,
    isAiReference: entry.isAiReference !== false,
    isMasterTemplate: Boolean(entry.isMasterTemplate),
    isLockedTemplate: Boolean(entry.isLockedTemplate || entry.isMasterTemplate),
    relevanceScore: score,
    blueprint: extractDocumentBlueprintForAi(entry)
  };
}

/**
 * Menghitung statistik Basis Pengetahuan Gudang Surat (arsip aktif vs Master Terkunci ⭐).
 */
export function summarizeKnowledgeBaseStats(vaultList = []) {
  const list = Array.isArray(vaultList) ? vaultList : [];
  const totalVault = list.length;
  const totalActive = list.filter(d => d && (d.isAiReference !== false || d.isMasterTemplate)).length;
  const masterCount = list.filter(d => d && d.isMasterTemplate).length;
  return { totalVault, totalActive, masterCount };
}

/**
 * Menghasilkan pesan umpan balik ringkas apakah AI mempelajari Master Terkunci ⭐, arsip Gudang, atau membuat baru dari nol.
 */
export function describeAiMatchFeedback(kbSelection = {}) {
  const top = kbSelection?.topReference;
  if (kbSelection?.matchMode === 'matched' && top) {
    if (top.isMasterTemplate) {
      return `AI mempelajari & meniru tata letak ⭐ Master Terkunci: "${top.title}".`;
    }
    return `AI mempelajari pola & gaya dari arsip Gudang: "${top.title}".`;
  }
  return 'Belum ada contoh sejenis di Gudang — AI merancang surat baru dari nol.';
}
