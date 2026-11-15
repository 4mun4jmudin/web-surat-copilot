/**
 * tests/smartCanvas.test.js - Unit & Integrasi Kanvas Word Pro (Multi-Halaman Bebas MS Word)
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { DOMParser as XmlDomParser } from '@xmldom/xmldom';
import {
  parseSmartPageHtmlToBlocks,
  buildDocModelFromSmartPagesData,
  docModelToSmartPages,
  extractSmartSnapshotFromPages,
  aiJsonToSmartDocModel,
  createEmptySmartDocument
} from '../js/ui/smartWordCanvas.js';
import { store, createEmptyDocumentModel } from '../js/state.js';
import { DEFAULT_SCHOOL_PROFILE } from '../js/letters/schoolProfile.js';
import { suratKeteranganTemplate } from '../js/letters/templates/suratKeterangan.js';
import { renderDocx } from '../js/docx/render.js';
import { readZipPackage } from '../js/utils/zip.js';

test('Word Pro: Konversi surat standar ke HTML bebas tanpa pembagian blok paragraf', () => {
  const sampleDoc = suratKeteranganTemplate.generateDocumentModel(
    {
      nomorSurat: '421.3/005/KET.SMKITA/XI/2026',
      judulSurat: 'SURAT KETERANGAN SISWA AKTIF',
      sifat: 'Penting',
      perihal: 'Keterangan Siswa Aktif',
      tujuan: 'Yth. Orang Tua/Wali Murid',
      namaSiswa: 'Ahmad Fauzi',
      nis: '24251001',
      kelas: 'XI RPL 1'
    },
    DEFAULT_SCHOOL_PROFILE,
    { paperSize: 'f4' }
  );

  const converted = docModelToSmartPages(sampleDoc, DEFAULT_SCHOOL_PROFILE);
  assert.equal(converted.pages.length, 1);
  const pageHtml = converted.pages[0].html;

  // Pastikan bebas dari pembagian kotak blok lama (.paragraph-box / .block-wrapper)
  assert.ok(!pageHtml.includes('paragraph-box'));
  assert.ok(!pageHtml.includes('block-wrapper'));
  assert.ok(pageHtml.includes('data-smart-role="title"'));
  assert.ok(pageHtml.includes('SURAT KETERANGAN SISWA AKTIF'));
  assert.ok(pageHtml.includes('Ahmad Fauzi'));
});

test('Word Pro: Ekspor multi-halaman (3 halaman bebas + tabel + list + format teks) ke .docx OOXML valid', async () => {
  const pages = [
    {
      id: 'page-1',
      isManualBreak: false,
      html: `
        <h2 class="smart-doc-title" data-smart-role="title" style="text-align: center;">SURAT KEPUTUSAN KEPALA SEKOLAH</h2>
        <p class="smart-doc-nomor" data-smart-role="nomor" style="text-align: center;">Nomor: 010/KEP.SMKITA/XI/2026</p>
        <table class="smart-meta-table" data-smart-role="meta-recipient">
          <tbody>
            <tr>
              <td>Sifat : Penting<br>Perihal : Panitia Ujian Sekolah</td>
              <td>Kepada<br>Yth. Dewan Guru<br>Di Tempat</td>
            </tr>
          </tbody>
        </table>
        <p style="text-align: justify;">Dengan ini menetapkan susunan panitia dengan ketentuan <strong>wajib hadir</strong> dan <em>tepat waktu</em> serta <u>bertanggung jawab</u>.</p>
        <ol>
          <li>Menyiapkan perangkat asesmen berbasis proyek.</li>
          <li>Mengawasi pelaksanaan evaluasi semester ganjil.</li>
        </ol>
      `
    },
    {
      id: 'page-2',
      isManualBreak: true,
      html: `
        <h2>LAMPIRAN SUSUNAN PANITIA</h2>
        <table class="smart-word-table">
          <thead>
            <tr><th>No</th><th>Nama Guru</th><th>Jabatan Panitia</th></tr>
          </thead>
          <tbody>
            <tr><td>1</td><td>Dian Nurhadian, S.Pd.</td><td>Penanggung Jawab</td></tr>
            <tr><td>2</td><td>Siti Aminah, S.Kom.</td><td>Ketua Pelaksana</td></tr>
          </tbody>
        </table>
      `
    },
    {
      id: 'page-3',
      isManualBreak: true,
      html: `
        <p style="text-align: right;">Garut, 15 November 2026</p>
        <table class="smart-sig-table" data-smart-role="signature">
          <tbody>
            <tr>
              <td class="smart-sig-spacer"></td>
              <td class="smart-sig-cell">Garut, 15 November 2026<br>Kepala Sekolah<br><br><strong>Dian Nurhadian, S.Pd.</strong><br>NIP. 198001012010011001</td>
            </tr>
          </tbody>
        </table>
      `
    }
  ];

  const docModel = buildDocModelFromSmartPagesData({
    pages,
    schoolProfile: DEFAULT_SCHOOL_PROFILE,
    letterOptions: { paperSize: 'f4', orientation: 'portrait', showKop: true }
  });

  // 3 halaman menghasilkan tepat 2 blok pageBreak di antara halaman
  const pageBreaks = docModel.blocks.filter(b => b.type === 'pageBreak');
  assert.equal(pageBreaks.length, 2);

  // Verifikasi tabel grid, list terurut, perataan kanan, dan inline runs terbaca akurat
  assert.ok(docModel.blocks.some(b => b.type === 'list' && b.ordered === true && b.items.length === 2));
  assert.ok(docModel.blocks.some(b => b.type === 'table' && b.rows.length === 3 && b.headerRow === true));
  assert.ok(docModel.blocks.some(b => b.type === 'paragraph' && b.alignment === 'right'));
  const richPara = docModel.blocks.find(b => b.type === 'paragraph' && (b.runs || []).some(r => r.bold && r.text.includes('wajib hadir')));
  assert.ok(richPara, 'Inline bold run harus terpreservasi dari <strong>');

  // Render menjadi .docx biner asli dan validasi document.xml menggunakan JSZip + XMLDOM
  const docxBytes = await renderDocx(docModel, 'uint8array');
  assert.ok(docxBytes instanceof Uint8Array);
  assert.ok(docxBytes.length > 1000);

  const zip = await readZipPackage(docxBytes);
  const docXmlStr = await zip.file('word/document.xml').async('string');

  const parser = new XmlDomParser();
  const xmlDoc = parser.parseFromString(docXmlStr, 'text/xml');
  assert.ok(xmlDoc && xmlDoc.documentElement);

  const pageBreakMatches = docXmlStr.match(/<w:br w:type="page"\/>/g) || [];
  assert.equal(pageBreakMatches.length, 2);
  assert.ok(docXmlStr.includes('<w:jc w:val="right"/>'));
  assert.ok(docXmlStr.includes('SURAT KEPUTUSAN KEPALA SEKOLAH'));
  assert.ok(docXmlStr.includes('Ketua Pelaksana'));
});

test('Word Pro: Snapshot & AI JSON converter mempertahankan struktur surat untuk Quick AI Edit dan Copilot', () => {
  const aiPayload = {
    letterTitle: 'SURAT UNDANGAN RAPAT',
    kodeJenisSurat: 'UND',
    nomorSurat: '050/UND.SMKITA/XI/2026',
    sifat: 'Penting',
    hal: 'Undangan Rapat Kenaikan Kelas',
    tujuan: 'Bapak/Ibu Dewan Guru SMKS IT Al-Hawari',
    blocks: [
      { type: 'paragraph', text: 'Assalamualaikum Wr. Wb. Dengan hormat, kami mengundang Bapak/Ibu pada:' },
      {
        type: 'identityTable',
        rows: [
          { label: 'Hari/Tanggal', value: ': Senin, 16 November 2026' },
          { label: 'Tempat', value: ': Ruang Guru Utama' }
        ]
      },
      { type: 'paragraph', text: 'Demikian surat undangan ini kami sampaikan.' }
    ],
    signer: {
      city: 'Garut',
      date: '15 November 2026',
      title: 'Kepala Sekolah',
      name: 'Dian Nurhadian, S.Pd.',
      nip: '-'
    }
  };

  const docModel = aiJsonToSmartDocModel(aiPayload, DEFAULT_SCHOOL_PROFILE, { paperSize: 'a4' });
  const converted = docModelToSmartPages(docModel, DEFAULT_SCHOOL_PROFILE);
  const snapshot = extractSmartSnapshotFromPages(converted.pages, DEFAULT_SCHOOL_PROFILE);

  assert.equal(snapshot.letterTitle, 'SURAT UNDANGAN RAPAT');
  assert.equal(snapshot.nomorSurat, '050/UND.SMKITA/XI/2026');
  assert.equal(snapshot.hideMetaRecipient, false);
  assert.equal(snapshot.sifat, 'Penting');
  assert.equal(snapshot.hal, 'Undangan Rapat Kenaikan Kelas');
  assert.ok(snapshot.blocks.some(b => b.type === 'identityTable' && b.rows.length === 2));
});

test('State Isolasi: Memperbarui smartDocument tidak mengganggu state document kanvas lama', () => {
  const originalDocBefore = JSON.stringify(store.getState().document);
  store.setState({
    smartDocument: {
      pages: [{ id: 'p1', html: '<p>Konten khusus Word Pro</p>', isManualBreak: false }],
      meta: { title: 'Uji Isolasi' }
    }
  }, false);

  const originalDocAfter = JSON.stringify(store.getState().document);
  assert.equal(originalDocAfter, originalDocBefore);
  assert.equal(store.getState().smartDocument.pages[0].id, 'p1');
});

test('Default Kosong: Word Pro (createEmptySmartDocument) dimulai benar-benar kosong tanpa teks contoh', () => {
  const blankSmart = createEmptySmartDocument(DEFAULT_SCHOOL_PROFILE, { paperSize: 'f4' });
  assert.equal(blankSmart.isBlankDraft, true);
  assert.equal(blankSmart.pages.length, 1);
  assert.equal(blankSmart.pages[0].html, '<p><br></p>');
  assert.ok(!blankSmart.pages[0].html.includes('Ahmad Fauzi'));
  assert.ok(!blankSmart.pages[0].html.includes('SURAT KETERANGAN'));
});

test('Default Kosong: Lembar Kerja (createEmptyDocumentModel) dimulai kosong dan tetap valid diekspor ke .docx', async () => {
  const blankDoc = createEmptyDocumentModel(DEFAULT_SCHOOL_PROFILE, { paperSize: 'f4' });
  assert.equal(blankDoc.isBlankDraft, true);
  assert.ok(!blankDoc.blocks.some(b => b.type === 'letterTitle'));
  assert.ok(!blankDoc.blocks.some(b => b.type === 'letterSignature'));
  const serialized = JSON.stringify(blankDoc);
  assert.ok(!serialized.includes('Ahmad Fauzi'));
  assert.ok(!serialized.includes('SURAT KETERANGAN'));

  // Pastikan dokumen kosong awal dari store juga mengekspor ZIP .docx OOXML yang valid
  const storeDefaultDoc = createEmptyDocumentModel(DEFAULT_SCHOOL_PROFILE, { showKop: false });
  const docxBytes = await renderDocx(storeDefaultDoc, 'uint8array');
  assert.ok(docxBytes instanceof Uint8Array);
  assert.ok(docxBytes.length > 500);
  const zip = await readZipPackage(docxBytes);
  const docXmlStr = await zip.file('word/document.xml').async('string');
  const parser = new XmlDomParser();
  const xmlDoc = parser.parseFromString(docXmlStr, 'text/xml');
  assert.ok(xmlDoc && xmlDoc.documentElement);
});
