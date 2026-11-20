/**
 * patch.js - Mode B: Mengedit Berkas .docx yang Diunggah Secara Aman
 * 
 * Prinsip:
 * 1. Tidak meregenerasi dokumen dari nol. Hanya memanipulasi word/document.xml.
 * 2. Menggabungkan teks dari seluruh <w:r> dalam <w:p> untuk pencarian anchor yang akurat.
 * 3. Menyalin pPr dan rPr dari paragraf tetangga agar format persis sama.
 * 4. Melindungi sectPr, header/footer, styles, numbering, bookmark, dan gambar.
 * 5. Mendukung rekonsiliasi Track Changes (<w:ins>, <w:del>).
 * 6. Menghasilkan Safety Check dan Laporan Perubahan komprehensif.
 */

import { readZipPackage } from '../utils/zip.js';
import { formatSmartNomorSurat } from '../letters/numbering.js';
import { parseXmlAsync, serializeXmlAsync, escapeXml } from '../utils/xml.js';

/**
 * Ekstraksi seluruh paragraf dari DOM document.xml
 * @param {Document} docXml 
 * @returns {Array<object>}
 */
export function extractParagraphs(docXml) {
  const pNodes = Array.from(docXml.getElementsByTagName('w:p'));
  
  return pNodes.map((pNode, index) => {
    // Kumpulkan semua w:t di dalam paragraph ini
    const tNodes = Array.from(pNode.getElementsByTagName('w:t'));
    const fullText = tNodes.map(t => t.textContent || '').join('');

    // Cari pPr, pStyle, & alignment
    const pPr = pNode.getElementsByTagName('w:pPr')[0] || null;
    let styleName = null;
    let hasSectPr = false;
    let alignment = 'left';

    if (pPr) {
      const pStyle = pPr.getElementsByTagName('w:pStyle')[0];
      if (pStyle) {
        styleName = pStyle.getAttribute('w:val');
      }
      if (pPr.getElementsByTagName('w:sectPr').length > 0) {
        hasSectPr = true;
      }
      const jc = pPr.getElementsByTagName('w:jc')[0];
      if (jc) {
        const val = jc.getAttribute('w:val');
        if (val === 'center') alignment = 'center';
        else if (val === 'right') alignment = 'right';
        else if (val === 'both') alignment = 'justify';
      }
    }

    // Ambil contoh rPr pertama
    const firstR = pNode.getElementsByTagName('w:r')[0];
    const rPr = firstR ? firstR.getElementsByTagName('w:rPr')[0] : null;

    // Cek tebal (bold)
    let isBold = false;
    const bTags = pNode.getElementsByTagName('w:b');
    if (bTags.length > 0) {
      const bVal = bTags[0].getAttribute('w:val');
      if (bVal !== 'false' && bVal !== '0') isBold = true;
    }

    // Cek garis bawah (underline)
    let isUnderline = false;
    const uTags = pNode.getElementsByTagName('w:u');
    if (uTags.length > 0) {
      const uVal = uTags[0].getAttribute('w:val');
      if (uVal && uVal !== 'none') isUnderline = true;
    }

    // Cek apakah di dalam tabel
    let isTable = false;
    let curr = pNode.parentNode;
    while (curr && curr.nodeType === 1) {
      if (curr.tagName === 'w:tc' || curr.tagName === 'w:tbl') {
        isTable = true;
        break;
      }
      curr = curr.parentNode;
    }

    return {
      index,
      node: pNode,
      fullText,
      pPr,
      rPr,
      styleName,
      hasSectPr,
      alignment,
      isBold,
      isUnderline,
      isTable
    };
  });
}

/**
 * Mencari paragraf anchor berdasarkan kriteria
 * @param {Array<object>} paragraphs 
 * @param {object} query { type: 'text'|'exact'|'heading'|'index', value: string|number }
 * @returns {object|null}
 */
export function findAnchorParagraph(paragraphs, query) {
  if (!query || query.value === undefined || query.value === null || query.value === '') return null;

  switch (query.type) {
    case 'index': {
      const targetIdx = Number(query.value);
      return paragraphs.find(p => p.index === targetIdx) || null;
    }

    case 'exact': {
      const targetText = String(query.value).trim();
      return paragraphs.find(p => p.fullText.trim() === targetText) || null;
    }

    case 'heading': {
      const queryText = String(query.value).toLowerCase().trim();
      return paragraphs.find(p => {
        const isHead = p.styleName && /heading/i.test(p.styleName);
        return isHead && p.fullText.toLowerCase().includes(queryText);
      }) || null;
    }

    case 'text':
    default: {
      const queryText = String(query.value).toLowerCase().trim();
      return paragraphs.find(p => p.fullText.toLowerCase().includes(queryText)) || null;
    }
  }
}

/**
 * Membuat elemen paragraf baru dengan mewarisi pPr dan rPr dari tetangga
 * @param {Document} docXml 
 * @param {string} textContent 
 * @param {Element|null} templatePPr 
 * @param {Element|null} templateRPr 
 * @param {object} [options]
 * @returns {Element}
 */
export function createClonedParagraph(docXml, textContent, templatePPr, templateRPr, options = {}) {
  const { trackChanges = false, author = 'AI Document Builder', trackId = 101 } = options;
  const p = docXml.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:p');

  // Kloning pPr jika ada, jangan bawa w:sectPr
  if (templatePPr) {
    const clonedPPr = templatePPr.cloneNode(true);
    const sectPrInside = clonedPPr.getElementsByTagName('w:sectPr')[0];
    if (sectPrInside) {
      clonedPPr.removeChild(sectPrInside);
    }
    p.appendChild(clonedPPr);
  }

  // Buat w:r
  const r = docXml.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:r');
  if (templateRPr) {
    r.appendChild(templateRPr.cloneNode(true));
  }

  const t = docXml.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:t');
  if (/^\s|\s$/.test(textContent)) {
    t.setAttribute('xml:space', 'preserve');
  }
  t.textContent = textContent;
  r.appendChild(t);

  if (trackChanges) {
    const ins = docXml.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:ins');
    ins.setAttribute('w:id', String(trackId));
    ins.setAttribute('w:author', author);
    ins.setAttribute('w:date', new Date().toISOString());
    ins.appendChild(r);
    p.appendChild(ins);
  } else {
    p.appendChild(r);
  }

  return p;
}

/**
 * Memodifikasi teks di paragraf eksisting dengan dukungan Track Changes
 * @param {Document} docXml 
 * @param {Element} pNode 
 * @param {string} newText 
 * @param {object} [options]
 */
export function replaceParagraphContent(docXml, pNode, newText, options = {}) {
  const { trackChanges = false, author = 'AI Document Builder', trackId = 102 } = options;

  // Dapatkan rPr dari run pertama bila ada
  const firstR = pNode.getElementsByTagName('w:r')[0];
  const templateRPr = firstR ? firstR.getElementsByTagName('w:rPr')[0] : null;

  if (trackChanges) {
    // Tandai run yang lama sebagai <w:del>
    const existingRuns = Array.from(pNode.getElementsByTagName('w:r'));
    const oldText = existingRuns.map(r => r.textContent || '').join('');

    // Hapus runs lama dari pNode (kecuali pPr)
    existingRuns.forEach(r => pNode.removeChild(r));

    const del = docXml.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:del');
    del.setAttribute('w:id', String(trackId));
    del.setAttribute('w:author', author);
    del.setAttribute('w:date', new Date().toISOString());

    const delR = docXml.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:r');
    if (templateRPr) delR.appendChild(templateRPr.cloneNode(true));
    const delT = docXml.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:delText');
    delT.textContent = oldText;
    delR.appendChild(delT);
    del.appendChild(delR);
    pNode.appendChild(del);

    // Tambahkan teks baru sebagai <w:ins>
    const ins = docXml.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:ins');
    ins.setAttribute('w:id', String(trackId + 1));
    ins.setAttribute('w:author', author);
    ins.setAttribute('w:date', new Date().toISOString());

    const insR = docXml.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:r');
    if (templateRPr) insR.appendChild(templateRPr.cloneNode(true));
    const insT = docXml.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:t');
    if (/^\s|\s$/.test(newText)) insT.setAttribute('xml:space', 'preserve');
    insT.textContent = newText;
    insR.appendChild(insT);
    ins.appendChild(insR);
    pNode.appendChild(ins);
  } else {
    // Hapus semua w:r
    const runs = Array.from(pNode.getElementsByTagName('w:r'));
    runs.forEach(r => pNode.removeChild(r));

    // Tambahkan w:r baru
    const r = docXml.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:r');
    if (templateRPr) r.appendChild(templateRPr.cloneNode(true));
    const t = docXml.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:t');
    if (/^\s|\s$/.test(newText)) t.setAttribute('xml:space', 'preserve');
    t.textContent = newText;
    r.appendChild(t);
    pNode.appendChild(r);
  }
}

/**
 * Menjalankan operasi patch pada file .docx (Mode B)
 * @param {Blob|Uint8Array|ArrayBuffer} docxData 
 * @param {Array<object>} operations 
 * @param {object} [options]
 * @returns {Promise<{ patchedData: any, report: object }>}
 */
export async function patchUploadedDocx(docxData, operations = [], options = {}) {
  const { trackChanges = false, author = 'AI Document Builder', outputType = 'uint8array' } = options;

  const zip = await readZipPackage(docxData);
  const docXmlFile = zip.file('word/document.xml');

  if (!docXmlFile) {
    throw new Error('Berkas yang diunggah bukan dokumen Word .docx yang valid (word/document.xml tidak ditemukan).');
  }

  // Simpan XML asli untuk validasi keamanan dan perbandingan diff
  const originalXmlText = await docXmlFile.async('text');
  const docXml = await parseXmlAsync(originalXmlText);

  // Ambil sectPr awal untuk safety check
  const initialSectPrCount = docXml.getElementsByTagName('w:sectPr').length;

  const report = {
    timestamp: new Date().toISOString(),
    totalOperations: operations.length,
    executedChanges: [],
    safetyPassed: true,
    warnings: []
  };

  let trackIdCounter = 500;

  for (let i = 0; i < operations.length; i++) {
    const op = operations[i];
    trackIdCounter += 2;

    // Scan ulang paragraf terkini setiap iterasi agar indeks dan DOM akurat
    const currentParagraphs = extractParagraphs(docXml);
    const anchor = findAnchorParagraph(currentParagraphs, op.anchor);

    if (!anchor) {
      report.warnings.push(`Operasi #${i + 1} (${op.action}): Anchor tidak ditemukan untuk kueri "${JSON.stringify(op.anchor)}"`);
      continue;
    }

    const targetNode = anchor.node;
    const parent = targetNode.parentNode;

    switch (op.action) {
      case 'insertAfter': {
        const newP = createClonedParagraph(docXml, op.text, anchor.pPr, anchor.rPr, {
          trackChanges,
          author,
          trackId: trackIdCounter
        });
        if (targetNode.nextSibling) {
          parent.insertBefore(newP, targetNode.nextSibling);
        } else {
          parent.appendChild(newP);
        }
        report.executedChanges.push({
          action: 'insertAfter',
          anchorText: anchor.fullText.slice(0, 80),
          insertedText: op.text.slice(0, 80),
          index: anchor.index
        });
        break;
      }

      case 'insertBefore': {
        const newP = createClonedParagraph(docXml, op.text, anchor.pPr, anchor.rPr, {
          trackChanges,
          author,
          trackId: trackIdCounter
        });
        parent.insertBefore(newP, targetNode);
        report.executedChanges.push({
          action: 'insertBefore',
          anchorText: anchor.fullText.slice(0, 80),
          insertedText: op.text.slice(0, 80),
          index: anchor.index
        });
        break;
      }

      case 'replace': {
        const oldText = anchor.fullText;
        replaceParagraphContent(docXml, targetNode, op.text, {
          trackChanges,
          author,
          trackId: trackIdCounter
        });
        report.executedChanges.push({
          action: 'replace',
          oldText: oldText.slice(0, 80),
          newText: op.text.slice(0, 80),
          index: anchor.index
        });
        break;
      }

      case 'delete': {
        const oldText = anchor.fullText;
        if (anchor.hasSectPr) {
          // Jangan hapus node jika memiliki sectPr! Hapus teksnya saja
          replaceParagraphContent(docXml, targetNode, '', {
            trackChanges,
            author,
            trackId: trackIdCounter
          });
          report.warnings.push(`Paragraf #${anchor.index} mengandung batas halaman (sectPr). Paragraf tidak dihapus, hanya konten teks yang dikosongkan.`);
        } else {
          if (trackChanges) {
            replaceParagraphContent(docXml, targetNode, '', { trackChanges: true, author, trackId: trackIdCounter });
          } else {
            parent.removeChild(targetNode);
          }
        }
        report.executedChanges.push({
          action: 'delete',
          deletedText: oldText.slice(0, 80),
          index: anchor.index
        });
        break;
      }

      default:
        report.warnings.push(`Aksi tidak dikenal: ${op.action}`);
    }
  }

  // --- SAFETY AUDIT CHECK ---
  // 1. Pastikan jumlah sectPr tidak berkurang
  const finalSectPrCount = docXml.getElementsByTagName('w:sectPr').length;
  if (finalSectPrCount !== initialSectPrCount) {
    report.safetyPassed = false;
    report.warnings.push(`PERINGATAN KEAMANAN: Jumlah section break berubah (sebelum: ${initialSectPrCount}, sesudah: ${finalSectPrCount}). Struktur halaman mungkin bergeser.`);
  }

  // 2. Serialisasi XML kembali
  const patchedXmlString = await serializeXmlAsync(docXml);

  // Perbarui file word/document.xml di dalam zip
  zip.file('word/document.xml', patchedXmlString);

  // Generate ulang paket docx
  const patchedData = await zip.generateAsync({
    type: outputType,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
  });

  return {
    patchedData,
    report,
    originalXml: originalXmlText,
    patchedXml: patchedXmlString
  };
}

/**
 * Mengekstrak dokumen .docx yang diunggah menjadi struktur yang dapat diedit langsung di Canvas
 * tanpa merusak format bawaan (menjaga bytes asli untuk patcher).
 * @param {Blob|Uint8Array|ArrayBuffer} docxData 
 * @returns {Promise<object>}
 */
export async function parseDocxForEditing(docxData) {
  const zip = await readZipPackage(docxData);
  const docXmlFile = zip.file('word/document.xml');

  if (!docXmlFile) {
    throw new Error('Berkas bukan format .docx Word yang valid (word/document.xml tidak ditemukan).');
  }

  const xmlText = await docXmlFile.async('text');
  const docXml = await parseXmlAsync(xmlText);
  const allParagraphs = extractParagraphs(docXml);

  const nonBlank = allParagraphs.filter(p => p.fullText.trim().length > 0);

  // 1. Cek apakah dokumen ini dari SMK IT Al-Hawari
  const isAlHawari = nonBlank.some(p => /al-hawari|cibiuk|nunuk\s*kartini|smk\s*islam\s*terpadu/i.test(p.fullText));

  // 2. Deteksi Judul Dokumen yang Akurat (Title)
  // Syarat judul: singkat (<= 75 kar), bukan kalimat penjelas (tidak diawali Demikian/Yang bertanda/Sehubungan/dsb),
  // dan cocok dengan pola judul surat dinas resmi (SURAT PANGGILAN, SURAT KETERANGAN, SURAT TUGAS, dsb).
  let detectedTitle = '';
  let titleIndex = -1;

  for (const p of nonBlank) {
    const text = p.fullText.trim();
    if (text.length > 75) continue;
    if (/^(demikian|yang\s+bertanda|sehubungan|berdasarkan|menindaklanjuti|mengharap|kepada|alamat|dinas|pemerintah)/i.test(text)) continue;

    // Pola judul surat resmi
    if (/^(surat\s+[a-z0-9\s\/\-]+|berita\s+acara\s+[a-z0-9\s\/\-]+|keputusan\s+[a-z0-9\s\/\-]+|pemberitahuan|undangan)$/i.test(text)) {
      detectedTitle = text;
      titleIndex = p.index;
      break;
    }
  }

  // Fallback judul: cari baris singkat diawali 'SURAT '
  if (!detectedTitle) {
    for (const p of nonBlank) {
      const text = p.fullText.trim();
      if (text.length <= 75 && text.toUpperCase().startsWith('SURAT ') && !/^(demikian)/i.test(text)) {
        detectedTitle = text;
        titleIndex = p.index;
        break;
      }
    }
  }

  // 3. Deteksi Nomor Surat
  let detectedNomor = '';
  let nomorIndex = -1;

  for (const p of nonBlank) {
    const text = p.fullText.trim();
    if (text.length <= 70 && /^(nomor|no)\s*:\s*[0-9a-z\/\.\-]+/i.test(text)) {
      detectedNomor = text;
      nomorIndex = p.index;
      break;
    }
  }

  // 4. Deteksi Tanda Tangan & Tanggal di Bagian Bawah
  let detectedDate = '';
  let sigDateIndex = -1;
  let detectedSigner = '';
  let sigNameIndex = -1;
  let detectedSigTitle = 'Kepala Sekolah';
  let sigTitleIndex = -1;
  let detectedNip = '-';
  let sigNipIndex = -1;

  for (let i = nonBlank.length - 1; i >= 0; i--) {
    const p = nonBlank[i];
    const text = p.fullText.trim();

    if (!detectedNip || detectedNip === '-') {
      if (/^(nip|nuptk)\s*[\.:]/i.test(text)) {
        detectedNip = text;
        sigNipIndex = p.index;
        continue;
      }
    }
    if (!detectedSigner) {
      if (text.includes('S.Pd') || text.includes('M.Pd') || text.includes('Drs.') || text.includes('Dian Nurhadian') || (p.isBold && text.length <= 40 && !text.includes(':'))) {
        detectedSigner = text;
        sigNameIndex = p.index;
        continue;
      }
    }
    if (!detectedSigTitle || detectedSigTitle === 'Kepala Sekolah') {
      if (/^(kepala\s+sekolah|wali\s+kelas|guru\s+bk|ketua|pembina)/i.test(text)) {
        detectedSigTitle = text;
        sigTitleIndex = p.index;
        continue;
      }
    }
    if (!detectedDate) {
      if (/(garut|cibiuk|bandung|jakarta|surabaya|[a-z]+),\s*[0-9]{1,2}\s+[a-z]+\s+[0-9]{4}/i.test(text)) {
        detectedDate = text;
        sigDateIndex = p.index;
        continue;
      }
    }
  }

  // 5. Kategorisasi struktur setiap paragraf
  const kopParagraphs = [];
  const metaRows = [];
  const recipientLines = [];
  const identityRows = [];
  const bodyParagraphs = [];
  let closingParagraph = null;

  for (const p of nonBlank) {
    const text = p.fullText.trim();

    // Cek apakah ini Kop (sebelum Title, berisi nama dinas/sekolah/alamat)
    if (titleIndex !== -1 && p.index < titleIndex) {
      if (/pemerintah|dinas|sekolah|smk|yayasan|alamat|curug|cibiuk|garut|npsn|email/i.test(text)) {
        p.role = 'kop';
        kopParagraphs.push({ index: p.index, text });
        continue;
      }
    }

    if (p.index === titleIndex) {
      p.role = 'title';
      continue;
    }

    if (p.index === nomorIndex) {
      p.role = 'nomor';
      continue;
    }

    if (p.index === sigDateIndex || p.index === sigTitleIndex || p.index === sigNameIndex || p.index === sigNipIndex) {
      p.role = 'signature';
      continue;
    }

    if (/^demikian\s+/i.test(text)) {
      p.role = 'closing';
      closingParagraph = { index: p.index, text };
      continue;
    }

    // Cek apakah baris Sifat / Lampiran / Perihal / Hal di bagian kepala surat
    const metaMatch = text.match(/^(sifat|lampiran|perihal|hal)\s*:\s*(.*)$/i);
    if (metaMatch) {
      p.role = 'meta';
      metaRows.push({
        index: p.index,
        label: metaMatch[1].trim(),
        value: metaMatch[2].trim()
      });
      continue;
    }

    // Cek apakah baris Tujuan / Kepada Yth / Di Tempat (sebelum isi surat utama)
    if (bodyParagraphs.length === 0 && /^(kepada(\s+yth\.?.*)?|yth\.?\s*.*|_{3,}|di|tempat|di\s+tempat)$/i.test(text)) {
      p.role = 'recipient';
      recipientLines.push(text);
      continue;
    }

    // Cek apakah baris rincian identitas / jadwal (e.g. Nama : ..., Hari/Tanggal : ..., dll)
    const colonMatch = text.match(/^([^:]{2,30})\s*:\s*(.+)$/);
    if (colonMatch || p.isTable) {
      p.role = 'identity';
      if (colonMatch) {
        identityRows.push({
          index: p.index,
          label: colonMatch[1].trim(),
          value: colonMatch[2].trim()
        });
      } else {
        identityRows.push({
          index: p.index,
          label: text,
          value: ''
        });
      }
      continue;
    }

    // Paragraf narasi biasa (pembuka / isi)
    p.role = 'body';
    bodyParagraphs.push({
      index: p.index,
      text,
      alignment: p.alignment,
      isBold: p.isBold,
      isUnderline: p.isUnderline
    });
  }

  return {
    rawBytes: docxData,
    totalParagraphs: allParagraphs.length,
    editableParagraphs: nonBlank.map(p => ({
      index: p.index,
      originalText: p.fullText,
      currentText: p.fullText,
      styleName: p.styleName,
      hasSectPr: p.hasSectPr,
      alignment: p.alignment,
      isBold: p.isBold,
      isUnderline: p.isUnderline,
      isTable: p.isTable,
      role: p.role || 'body'
    })),
    meta: {
      title: detectedTitle || 'Dokumen Word Terunggah',
      nomor: detectedNomor || '-',
      date: detectedDate || '14 November 2026',
      signer: detectedSigner || 'Dian Nurhadian, S.Pd.'
    },
    structured: {
      isAlHawari,
      kopParagraphs,
      title: { text: detectedTitle || 'SURAT RESMI', index: titleIndex },
      nomor: { text: detectedNomor || '', index: nomorIndex },
      metaRows,
      recipientLines,
      identityRows,
      bodyParagraphs,
      closing: closingParagraph,
      signature: {
        date: detectedDate || 'Garut, 14 November 2026',
        dateIndex: sigDateIndex,
        title: detectedSigTitle || 'Kepala Sekolah',
        titleIndex: sigTitleIndex,
        signer: detectedSigner || 'Dian Nurhadian, S.Pd.',
        signerIndex: sigNameIndex,
        nip: detectedNip || '-',
        nipIndex: sigNipIndex
      }
    }
  };
}

/**
 * Menerapkan perubahan teks dari canvas ke file .docx asli menggunakan patchUploadedDocx
 * sehingga 100% tata letak, kop asli, font, margin, dan styles tidak rusak.
 * @param {Blob|Uint8Array|ArrayBuffer} originalBytes 
 * @param {Array<{ index: number, originalText: string, currentText: string }>} edits 
 * @param {object} [options]
 * @returns {Promise<Uint8Array>}
 */
export async function applyCanvasEditsToDocx(originalBytes, edits = [], options = {}) {
  const operations = [];

  for (const edit of edits) {
    // Hanya lakukan replace jika teks benar-benar diubah
    if (edit.currentText !== undefined && edit.currentText !== edit.originalText) {
      operations.push({
        action: 'replace',
        anchor: { type: 'index', value: edit.index },
        text: edit.currentText
      });
    }
  }

  // Jika tidak ada perubahan, kembalikan data asli
  if (operations.length === 0) {
    let raw;
    if (originalBytes instanceof Uint8Array) raw = originalBytes;
    else if (originalBytes instanceof ArrayBuffer) raw = new Uint8Array(originalBytes);
    else raw = new Uint8Array(await new Response(originalBytes).arrayBuffer());

    if (options.returnFullResult) {
      return {
        patchedData: raw,
        report: { safetyPassed: true, executedChanges: [], warnings: [] },
        originalXml: '',
        patchedXml: ''
      };
    }
    return raw;
  }

  const result = await patchUploadedDocx(originalBytes, operations, {
    trackChanges: options.trackChanges || false,
    author: options.author || 'AI Document Builder',
    outputType: options.outputType || 'uint8array'
  });

  if (options.returnFullResult) {
    return result;
  }

  return result.patchedData;
}

/**
 * Mengonversi berkas docx yang diunggah / diparsing menjadi Document Model web murni
 * Mengambil teksnya, merapikan struktur menjadi format surat dinas sekolah, dan memuatnya langsung ke canvas
 * @param {object} parsedDocx Hasil dari parseDocxForEditing
 * @param {object} [schoolProfile]
 * @param {object} [options]
 * @returns {object} Document Model resmi web
 */
export function convertDocxToDocumentModel(parsedDocx, schoolProfile = null, options = {}) {
  const meta = parsedDocx.meta || {};
  const structured = parsedDocx.structured || {};

  const title = (structured.title?.text || meta.title || 'SURAT PANGGILAN ORANG TUA SISWA').trim();
  const rawNomor = (structured.nomor?.text || meta.nomor || '').trim();
  const perihalFromMeta = (structured.metaRows || []).find(r => /^(perihal|hal)$/i.test(r.label))?.value || '';
  const smartNomor = formatSmartNomorSurat({
    rawNomor,
    contextText: `${title} ${perihalFromMeta}`
  });

  const defaultMetaRows = (structured.metaRows && structured.metaRows.length > 0)
    ? structured.metaRows.map(r => ({ label: r.label, value: r.value }))
    : [
        { label: 'Sifat', value: 'Penting' },
        { label: 'Perihal', value: 'Undangan Rapat' }
      ];

  const defaultRecipientLines = (structured.recipientLines && structured.recipientLines.length > 0)
    ? structured.recipientLines
    : [
        'Kepada',
        'Yth. ...............',
        '__________________',
        'Di',
        'Tempat'
      ];

  const blocks = [
    {
      type: 'letterKop',
      school: schoolProfile
    },
    {
      type: 'letterTitle',
      title: title.toUpperCase(),
      nomor: `Nomor: ${smartNomor}`,
      metaRecipient: {
        metaRows: defaultMetaRows,
        recipientLines: defaultRecipientLines
      }
    }
  ];

  // Paragraf Pembuka
  if (structured.bodyParagraphs && structured.bodyParagraphs.length > 0) {
    blocks.push({
      type: 'paragraph',
      style: 'letterBody',
      runs: [{ text: structured.bodyParagraphs[0].text }]
    });
  }

  // Tabel Identitas / Jadwal (jika ada)
  if (structured.identityRows && structured.identityRows.length > 0) {
    blocks.push({
      type: 'identityTable',
      rows: structured.identityRows.map(r => ({
        label: r.label,
        value: r.value.startsWith(':') ? r.value : `: ${r.value}`
      }))
    });
  }

  // Paragraf Maksud / Isi berikutnya
  if (structured.bodyParagraphs && structured.bodyParagraphs.length > 1) {
    for (let i = 1; i < structured.bodyParagraphs.length; i++) {
      blocks.push({
        type: 'paragraph',
        style: 'letterBody',
        runs: [{ text: structured.bodyParagraphs[i].text }]
      });
    }
  }

  // Paragraf Penutup
  if (structured.closing) {
    blocks.push({
      type: 'paragraph',
      style: 'letterBody',
      runs: [{ text: structured.closing.text }]
    });
  }

  // Blok Tanda Tangan
  const sig = structured.signature || {};
  blocks.push({
    type: 'letterSignature',
    city: sig.city || 'Garut',
    date: sig.date || meta.date || '14 November 2026',
    title: sig.title || 'Kepala Sekolah',
    name: sig.signer || schoolProfile?.principalName || 'Dian Nurhadian, S.Pd.',
    nip: sig.nip && sig.nip !== '-' ? sig.nip : (schoolProfile?.principalNip || '-')
  });

  return {
    meta: {
      title,
      author: schoolProfile?.schoolName || 'SMK IT Al-Hawari',
      year: new Date().getFullYear().toString()
    },
    profile: options.paperSize === 'f4' ? 'suratDinasF4' : 'suratDinas',
    paperSize: options.paperSize || 'a4',
    orientation: options.orientation || 'portrait',
    customMargins: options.margins,
    letterData: {
      templateId: 'surat_konversi',
      schoolProfile,
      formData: {
        nomorSurat: String(smartNomor || '').replace(/^Nomor:\s*/i, ''),
        judulSurat: title
      },
      options
    },
    blocks
  };
}

