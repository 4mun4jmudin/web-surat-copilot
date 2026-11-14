/**
 * schema.js - Validator Mandiri untuk Document Model
 * Memvalidasi struktur JSON kontrak antara AI dan Renderer Word (.docx)
 * Mendukung dokumen laporan akademik serta Sistem Persuratan Dinas Sekolah.
 */

/**
 * Memvalidasi sebuah Document Model JSON.
 * @param {any} doc 
 * @returns {{ valid: boolean, errors: string[] }}
 */
export function validateDocument(doc) {
  const errors = [];

  if (!doc || typeof doc !== 'object' || Array.isArray(doc)) {
    return { valid: false, errors: ['Dokumen harus berupa objek JSON valid.'] };
  }

  // 1. Validasi meta
  if (!doc.meta || typeof doc.meta !== 'object') {
    errors.push('Bagian "meta" wajib ada dan harus berupa objek.');
  } else {
    if (!doc.meta.title || typeof doc.meta.title !== 'string') {
      errors.push('meta.title wajib ada dan harus berupa teks.');
    }
  }

  // 2. Validasi profile
  const validProfiles = ['suratdinas', 'suratdinasf4', 'pkl', 'skripsi', 'jurnal'];
  if (!doc.profile || typeof doc.profile !== 'string') {
    errors.push('Bagian "profile" wajib ada.');
  } else if (!validProfiles.includes(doc.profile.toLowerCase())) {
    errors.push(`Profil "${doc.profile}" tidak dikenali. Pilihan valid: ${validProfiles.join(', ')}.`);
  }

  // 3. Validasi blocks
  if (!Array.isArray(doc.blocks)) {
    errors.push('Bagian "blocks" wajib ada dan harus berupa array.');
    return { valid: false, errors };
  }

  if (doc.blocks.length === 0) {
    errors.push('Dokumen tidak boleh kosong ("blocks" memiliki 0 elemen).');
  }

  doc.blocks.forEach((block, index) => {
    const prefix = `Blok #${index + 1}`;
    if (!block || typeof block !== 'object') {
      errors.push(`${prefix}: harus berupa objek.`);
      return;
    }

    if (!block.type || typeof block.type !== 'string') {
      errors.push(`${prefix}: properti "type" wajib diisi.`);
      return;
    }

    switch (block.type) {
      // --- KOMPONEN SURAT DINAS ---
      case 'letterKop':
        if (!block.school || typeof block.school !== 'object') {
          errors.push(`${prefix} (letterKop): properti "school" wajib diisi dengan objek profil sekolah.`);
        }
        break;

      case 'letterTitle':
        if (!block.title || typeof block.title !== 'string') {
          errors.push(`${prefix} (letterTitle): "title" wajib ada.`);
        }
        break;

      case 'letterMetaRecipient':
        break;

      case 'identityTable':
        if (!Array.isArray(block.rows)) {
          errors.push(`${prefix} (identityTable): "rows" harus berupa array label-value.`);
        }
        break;

      case 'letterSignature':
        if (!block.name || typeof block.name !== 'string') {
          errors.push(`${prefix} (letterSignature): "name" pejabat penandatangan wajib diisi.`);
        }
        break;

      // --- KOMPONEN UMUM ---
      case 'heading':
        if (![1, 2, 3].includes(block.level)) {
          errors.push(`${prefix} (heading): "level" harus berupa angka 1, 2, atau 3.`);
        }
        if (typeof block.text !== 'string' || block.text.trim() === '') {
          errors.push(`${prefix} (heading): "text" wajib ada dan tidak boleh kosong.`);
        }
        break;

      case 'paragraph':
        if (block.text !== undefined && typeof block.text !== 'string') {
          errors.push(`${prefix} (paragraph): "text" harus berupa string.`);
        }
        if (block.runs !== undefined) {
          if (!Array.isArray(block.runs)) {
            errors.push(`${prefix} (paragraph): "runs" harus berupa array.`);
          } else {
            block.runs.forEach((run, rIdx) => {
              if (!run || typeof run !== 'object' || typeof run.text !== 'string') {
                errors.push(`${prefix} run #${rIdx + 1}: wajib memiliki properti "text" bertipe string.`);
              }
            });
          }
        }
        if (block.text === undefined && (!block.runs || block.runs.length === 0)) {
          errors.push(`${prefix} (paragraph): harus menyertakan "text" atau "runs".`);
        }
        break;

      case 'list':
        if (typeof block.ordered !== 'boolean') {
          errors.push(`${prefix} (list): "ordered" wajib bernilai boolean (true/false).`);
        }
        if (!Array.isArray(block.items) || block.items.length === 0) {
          errors.push(`${prefix} (list): "items" harus berupa array yang tidak kosong.`);
        }
        break;

      case 'table':
        if (!Array.isArray(block.rows) || block.rows.length === 0) {
          errors.push(`${prefix} (table): "rows" wajib ada dan berupa array baris.`);
        }
        break;

      case 'figure':
        if (!block.caption || typeof block.caption !== 'string') {
          errors.push(`${prefix} (figure): "caption" wajib diisi.`);
        }
        break;

      case 'pageBreak':
        break;

      case 'sectionBreak':
        if (block.numberFormat && !['none', 'roman', 'decimal'].includes(block.numberFormat)) {
          errors.push(`${prefix} (sectionBreak): "numberFormat" harus salah satu dari: none, roman, decimal.`);
        }
        break;

      case 'toc':
      case 'listOfFigures':
      case 'listOfTables':
        break;

      case 'reference':
        if (!block.id || !block.text) {
          errors.push(`${prefix} (reference): "id" dan "text" wajib ada.`);
        }
        break;

      case 'citation':
        if (!block.refId) {
          errors.push(`${prefix} (citation): "refId" rujukan wajib diisi.`);
        }
        break;

      case 'crossRef':
        if (!block.targetId) {
          errors.push(`${prefix} (crossRef): "targetId" bookmark wajib diisi.`);
        }
        break;

      default:
        errors.push(`${prefix}: tipe blok "${block.type}" tidak didukung.`);
    }
  });

  return {
    valid: errors.length === 0,
    errors
  };
}

export function normalizeDocument(doc) {
  if (!doc || !Array.isArray(doc.blocks)) return doc;

  const normalizedBlocks = doc.blocks.map(block => {
    if (block.type === 'paragraph') {
      if (!block.runs) {
        return {
          ...block,
          runs: [{ text: block.text || '' }]
        };
      }
    }
    return block;
  });

  return {
    ...doc,
    blocks: normalizedBlocks
  };
}
