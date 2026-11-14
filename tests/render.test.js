import { test } from 'node:test';
import assert from 'node:assert/strict';
import { renderDocx, getProfile } from '../js/docx/render.js';
import { readZipPackage } from '../js/utils/zip.js';
import { samplePklDocument } from '../js/model/sample.js';

test('Renderer DOCX: Profil PKL, Skripsi, Jurnal terdefinisi dengan ukuran A4', () => {
  const pkl = getProfile('pkl');
  const skripsi = getProfile('skripsi');
  const jurnal = getProfile('jurnal');

  assert.equal(pkl.page.width, 11906);
  assert.equal(pkl.page.height, 16838);
  assert.equal(skripsi.page.width, 11906);
  assert.equal(jurnal.page.width, 11906);
});

test('Renderer DOCX: Menghasilkan paket berkas .docx lengkap (Mode A)', async () => {
  const docxUint8 = await renderDocx(samplePklDocument, 'uint8array');
  assert.ok(docxUint8 instanceof Uint8Array);
  assert.ok(docxUint8.length > 500);

  // Verifikasi isi ZIP
  const zip = await readZipPackage(docxUint8);
  
  // Pastikan struktur OOXML wajib ada
  const requiredFiles = [
    '[Content_Types].xml',
    '_rels/.rels',
    'word/_rels/document.xml.rels',
    'word/document.xml',
    'word/styles.xml',
    'word/numbering.xml',
    'word/settings.xml',
    'word/footer1.xml',
    'word/footer2.xml'
  ];

  for (const filename of requiredFiles) {
    const file = zip.file(filename);
    assert.ok(file, `File ${filename} harus ada di dalam arsip docx`);
  }

  // Periksa konten XML word/document.xml
  const documentXml = await zip.file('word/document.xml').async('text');
  assert.match(documentXml, /<w:document/);
  assert.match(documentXml, /<w:pStyle w:val="Heading1"/);
  assert.match(documentXml, /TOC/);
  assert.match(documentXml, /SEQ Gambar/);
  assert.match(documentXml, /SEQ Tabel/);
  assert.match(documentXml, /<w:tbl/);
  assert.match(documentXml, /<w:sectPr/);
});

