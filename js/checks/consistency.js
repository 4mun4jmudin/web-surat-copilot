/**
 * consistency.js - Pemeriksa Konsistensi Dokumen Akademik
 * Memeriksa masalah struktural, penomoran, rujukan, tipografi, dan kepatuhan standar.
 */

import { getProfile } from '../docx/render.js';

/**
 * @typedef {Object} ConsistencyIssue
 * @property {'error'|'warning'|'info'} severity
 * @property {string} code
 * @property {string} message
 * @property {number} [blockIndex]
 * @property {string} [suggestion]
 */

/**
 * Menjalankan pemeriksaan konsistensi menyeluruh terhadap Document Model
 * @param {object} doc Document Model
 * @returns {{ issues: ConsistencyIssue[], summary: { errors: number, warnings: number, infos: number } }}
 */
export function checkDocumentConsistency(doc) {
  const issues = [];

  if (!doc || !Array.isArray(doc.blocks)) {
    return {
      issues: [{ severity: 'error', code: 'INVALID_DOC', message: 'Dokumen tidak memiliki daftar blok yang valid.' }],
      summary: { errors: 1, warnings: 0, infos: 0 }
    };
  }

  const profile = getProfile(doc.profile);

  // 1. Cek Ukuran Kertas & Margin Profil
  if (profile.page.width !== 11906 || profile.page.height !== 16838) {
    issues.push({
      severity: 'warning',
      code: 'NON_A4_PAPER',
      message: `Ukuran kertas pada profil "${profile.name}" bukan standar A4 (21 x 29.7 cm).`,
      suggestion: 'Gunakan ukuran kertas A4 untuk dokumen formal/akademik.'
    });
  }

  // Pengumpulan data untuk pelacakan
  const registeredBookmarks = new Set();
  const registeredReferences = new Set();
  const crossReferences = [];
  const citations = [];
  
  let lastHeadingLevel = 0;
  let figureCounter = 0;
  let tableCounter = 0;
  let consecutiveEmptyCount = 0;

  doc.blocks.forEach((block, index) => {
    // 2. Cek Paragraf Kosong Berlebih
    if (block.type === 'paragraph') {
      const text = (block.runs ? block.runs.map(r => r.text).join('') : (block.text || '')).trim();
      if (text === '') {
        consecutiveEmptyCount++;
        if (consecutiveEmptyCount >= 2) {
          issues.push({
            severity: 'warning',
            code: 'EXCESSIVE_EMPTY_PARAGRAPHS',
            blockIndex: index,
            message: `Ditemukan paragraf kosong berturut-turut pada blok #${index + 1}.`,
            suggestion: 'Hapus spasi paragraf manual; gunakan pengaturan spasi before/after pada profil.'
          });
        }
      } else {
        consecutiveEmptyCount = 0;
      }
    } else {
      consecutiveEmptyCount = 0;
    }

    // 3. Cek Hierarki Heading (Level melompat)
    if (block.type === 'heading') {
      const currentLevel = block.level;
      if (lastHeadingLevel > 0 && currentLevel > lastHeadingLevel + 1) {
        issues.push({
          severity: 'warning',
          code: 'HEADING_LEVEL_SKIPPED',
          blockIndex: index,
          message: `Heading pada blok #${index + 1} melompat dari level ${lastHeadingLevel} langsung ke level ${currentLevel}.`,
          suggestion: `Gunakan heading level ${lastHeadingLevel + 1} sebelum turun ke level yang lebih dalam.`
        });
      }
      lastHeadingLevel = currentLevel;

      // Cek kapitalisasi Heading 1 (Standar akademik: ALL CAPS)
      if (currentLevel === 1 && block.text) {
        const isAllCaps = block.text === block.text.toUpperCase();
        if (!isAllCaps && profile.headings.heading1.allCaps) {
          issues.push({
            severity: 'info',
            code: 'HEADING1_CAPITALIZATION',
            blockIndex: index,
            message: `Judul bab "${block.text}" belum berupa HURUF KAPITAL SEMUA.`,
            suggestion: 'Ubah teks judul bab menjadi huruf kapital penuh sesuai profil.'
          });
        }
      }

      if (block.id) registeredBookmarks.add(block.id);
    }

    // 4. Catat Gambar & Bookmark
    if (block.type === 'figure') {
      figureCounter++;
      if (block.id) registeredBookmarks.add(block.id);
      if (!block.caption || block.caption.trim() === '') {
        issues.push({
          severity: 'error',
          code: 'FIGURE_MISSING_CAPTION',
          blockIndex: index,
          message: `Gambar #${figureCounter} (blok #${index + 1}) tidak memiliki caption.`,
          suggestion: 'Tambahkan judul/keterangan gambar yang jelas.'
        });
      }
    }

    // 5. Catat Tabel & Bookmark
    if (block.type === 'table') {
      tableCounter++;
      if (block.id) registeredBookmarks.add(block.id);
      if (!block.caption || block.caption.trim() === '') {
        issues.push({
          severity: 'warning',
          code: 'TABLE_MISSING_CAPTION',
          blockIndex: index,
          message: `Tabel #${tableCounter} (blok #${index + 1}) tidak memiliki caption.`,
          suggestion: 'Sertakan judul tabel di atas tabel sesuai format baku.'
        });
      }
    }

    // 6. Catat Cross-Reference
    if (block.type === 'crossRef') {
      crossReferences.push({ blockIndex: index, targetId: block.targetId });
    }

    // 7. Catat Reference & Citation
    if (block.type === 'reference') {
      registeredReferences.add(block.id);
      if (block.id) registeredBookmarks.add(block.id);
    }

    if (block.type === 'citation') {
      citations.push({ blockIndex: index, refId: block.refId });
    }
  });

  // 8. Cek Rujukan Silang yang Rusak / Target Tidak Ada
  crossReferences.forEach(cr => {
    if (!registeredBookmarks.has(cr.targetId)) {
      issues.push({
        severity: 'error',
        code: 'BROKEN_CROSS_REFERENCE',
        blockIndex: cr.blockIndex,
        message: `Rujukan silang ke "${cr.targetId}" pada blok #${cr.blockIndex + 1} tidak memiliki target bookmark yang cocok.`,
        suggestion: `Pastikan ada elemen tabel, gambar, atau bab dengan ID "${cr.targetId}".`
      });
    }
  });

  // 9. Cek Sitasi yang Tidak Terdaftar di Daftar Pustaka
  citations.forEach(cit => {
    if (!registeredReferences.has(cit.refId)) {
      issues.push({
        severity: 'warning',
        code: 'UNRESOLVED_CITATION',
        blockIndex: cit.blockIndex,
        message: `Sitasi [${cit.refId}] pada blok #${cit.blockIndex + 1} tidak ditemukan di daftar referensi/pustaka.`,
        suggestion: `Tambahkan rujukan untuk "${cit.refId}" pada bagian Daftar Pustaka.`
      });
    }
  });

  const summary = {
    errors: issues.filter(i => i.severity === 'error').length,
    warnings: issues.filter(i => i.severity === 'warning').length,
    infos: issues.filter(i => i.severity === 'info').length
  };

  return {
    issues,
    summary
  };
}

/**
 * Melakukan perbaikan otomatis (auto-fix) untuk isu-isu umum (paragraf kosong, heading all-caps)
 * @param {object} doc 
 * @returns {object} Dokumen yang telah diperbaiki
 */
export function autoFixDocument(doc) {
  if (!doc || !Array.isArray(doc.blocks)) return doc;

  const profile = getProfile(doc.profile);
  const fixedBlocks = [];
  let prevWasEmpty = false;

  for (const block of doc.blocks) {
    if (block.type === 'paragraph') {
      const text = (block.runs ? block.runs.map(r => r.text).join('') : (block.text || '')).trim();
      if (text === '') {
        if (prevWasEmpty) {
          // Lewati paragraf kosong berlebih
          continue;
        }
        prevWasEmpty = true;
      } else {
        prevWasEmpty = false;
      }
      fixedBlocks.push(block);
    } else if (block.type === 'heading') {
      prevWasEmpty = false;
      if (block.level === 1 && profile.headings.heading1.allCaps && block.text) {
        fixedBlocks.push({
          ...block,
          text: block.text.toUpperCase()
        });
      } else {
        fixedBlocks.push(block);
      }
    } else {
      prevWasEmpty = false;
      fixedBlocks.push(block);
    }
  }

  return {
    ...doc,
    blocks: fixedBlocks
  };
}

