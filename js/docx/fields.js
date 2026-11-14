/**
 * fields.js - Helper pembuat Native Word Field Codes (OOXML)
 * Mendukung PAGE, NUMPAGES, TOC, SEQ (Gambar/Tabel), Bookmark & REF.
 */

import { escapeXml } from '../utils/xml.js';

/**
 * Menghasilkan blok OOXML untuk field run (fldChar: begin -> instrText -> separate -> fallback -> end)
 * @param {string} instruction 
 * @param {string} fallbackText 
 * @param {string} [rPrXml]
 * @returns {string}
 */
export function buildFieldRun(instruction, fallbackText = '1', rPrXml = '') {
  const safeInstr = escapeXml(instruction);
  const safeFallback = escapeXml(fallbackText);
  return `
    <w:r>${rPrXml}<w:fldChar w:fldCharType="begin"/></w:r>
    <w:r>${rPrXml}<w:instrText xml:space="preserve"> ${safeInstr} </w:instrText></w:r>
    <w:r>${rPrXml}<w:fldChar w:fldCharType="separate"/></w:r>
    <w:r>${rPrXml}<w:t>${safeFallback}</w:t></w:r>
    <w:r>${rPrXml}<w:fldChar w:fldCharType="end"/></w:r>
  `.trim();
}

/**
 * Menghasilkan field nomor halaman (PAGE)
 * @param {string} [rPrXml]
 * @returns {string}
 */
export function buildPageField(rPrXml = '') {
  return buildFieldRun('PAGE', '1', rPrXml);
}

/**
 * Menghasilkan field total halaman (NUMPAGES)
 * @param {string} [rPrXml]
 * @returns {string}
 */
export function buildNumPagesField(rPrXml = '') {
  return buildFieldRun('NUMPAGES', '1', rPrXml);
}

/**
 * Menghasilkan field nomor urut otomatis (SEQ) untuk gambar atau tabel
 * @param {'Gambar'|'Tabel'|string} sequenceName 
 * @param {number|string} fallbackNumber 
 * @param {string} [rPrXml]
 * @returns {string}
 */
export function buildSeqField(sequenceName, fallbackNumber = 1, rPrXml = '') {
  return buildFieldRun(`SEQ ${sequenceName} \\* ARABIC`, String(fallbackNumber), rPrXml);
}

/**
 * Menghasilkan rujukan silang (REF) ke sebuah bookmark
 * @param {string} bookmarkName 
 * @param {string} fallbackText 
 * @param {string} [rPrXml]
 * @returns {string}
 */
export function buildRefField(bookmarkName, fallbackText = '1', rPrXml = '') {
  return buildFieldRun(`REF ${bookmarkName} \\h`, fallbackText, rPrXml);
}

/**
 * Bookmark pembuka
 * @param {number|string} id 
 * @param {string} name 
 * @returns {string}
 */
export function buildBookmarkStart(id, name) {
  const safeName = escapeXml(name.replace(/\s+/g, '_'));
  return `<w:bookmarkStart w:id="${id}" w:name="${safeName}"/>`;
}

/**
 * Bookmark penutup
 * @param {number|string} id 
 * @returns {string}
 */
export function buildBookmarkEnd(id) {
  return `<w:bookmarkEnd w:id="${id}"/>`;
}

/**
 * Blok Daftar Isi Word (TOC) otomatis berdasarkan Heading 1 - Heading 3
 * @param {string} [tocTitle]
 * @returns {string} XML paragraf TOC
 */
export function buildTocBlock(tocTitle = 'DAFTAR ISI') {
  return `
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:jc w:val="center"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
        </w:rPr>
        <w:t>${escapeXml(tocTitle)}</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
      </w:pPr>
      <w:r>
        <w:fldChar w:fldCharType="begin"/>
      </w:r>
      <w:r>
        <w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType="separate"/>
      </w:r>
      <w:r>
        <w:rPr><w:i/></w:rPr>
        <w:t>[Daftar Isi otomatis akan diperbarui saat dibuka di Microsoft Word]</w:t>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType="end"/>
      </w:r>
    </w:p>
  `.trim();
}

/**
 * Blok Daftar Gambar otomatis memakai field TOC \c "Gambar"
 * @param {string} [title]
 * @returns {string}
 */
export function buildListOfFiguresBlock(title = 'DAFTAR GAMBAR') {
  return `
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:jc w:val="center"/>
      </w:pPr>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>${escapeXml(title)}</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> TOC \\c "Gambar" \\h \\z </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:rPr><w:i/></w:rPr><w:t>[Daftar Gambar otomatis]</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  `.trim();
}

/**
 * Blok Daftar Tabel otomatis memakai field TOC \c "Tabel"
 * @param {string} [title]
 * @returns {string}
 */
export function buildListOfTablesBlock(title = 'DAFTAR TABEL') {
  return `
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:jc w:val="center"/>
      </w:pPr>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>${escapeXml(title)}</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> TOC \\c "Tabel" \\h \\z </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:rPr><w:i/></w:rPr><w:t>[Daftar Tabel otomatis]</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
  `.trim();
}

