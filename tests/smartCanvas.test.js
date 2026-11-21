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
import { buildSuratTugasSppdDocModel } from '../js/letters/templates/suratTugasSppd.js';
import { applySignatureModeToDocModel } from '../js/letters/verification.js';
import { computeColumnResizeWidths, pxToCm, cmToPx } from '../js/ui/tableInteraction.js';
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
    tujuan: 'Bapak/Ibu Dewan Guru SMK IT Al-Hawari',
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

test('1:1 MS Word Layout: Surat Tugas + SPPD 3 Halaman (Merge Colspan/Rowspan, Garis SPPD, 2-Kolom, QR, & Ekspor .docx)', async () => {
  const stSppdModel = buildSuratTugasSppdDocModel({
    nomorSurat: '065/ST/SMKITA/12.01/XI/2026',
    assignees: [
      { nama: 'Taufik Dwi Setyo Ramadhan, S.Pd.', jabatan: 'Guru' },
      { nama: 'Siti Laras Maemunah, S.Pd.', jabatan: 'Guru' }
    ],
    maksudTugas: 'Penjemputan Peserta Didik PKL',
    tujuanLokasi: 'DRAGON MOTION PICTURE ASIA',
    tempatTiba: 'Dragon Motion Picture',
    maksudSppd: 'Penjemputan siswa kelas 12 Multimedia ke tempat PKL',
    kendaraan: 'Mobil',
    kendaraanSingkat: 'Roda 4',
    lamaHari: '1',
    tanggalSurat: '18 November 2026'
  }, DEFAULT_SCHOOL_PROFILE, { paperSize: 'f4' });

  // 1. Pastikan 3 halaman HTML Word Pro terbentuk lengkap
  const converted = docModelToSmartPages(stSppdModel, DEFAULT_SCHOOL_PROFILE);
  assert.equal(converted.pages.length, 3);
  assert.ok(converted.pages[0].html.includes('SURAT TUGAS'));
  assert.ok(converted.pages[0].html.includes('Tiba di'));
  assert.ok(converted.pages[1].html.includes('( S P P D )'));
  assert.ok(converted.pages[1].html.includes('smart-table-sppd'));
  assert.ok(converted.pages[1].html.includes('smart-table-none'));
  assert.ok(converted.pages[1].html.includes('CATATAN:'));
  assert.ok(converted.pages[2].html.includes('K e t e r a n g a n :'));
  assert.ok(converted.pages[2].html.includes('rowspan="2"'));
  assert.ok(converted.pages[2].html.includes('colspan="2"'));
  assert.ok(converted.pages[2].html.includes('colspan="3"'));

  // 2. Round-trip Smart HTML -> DocumentModel tetap menjaga seluruh fitur tabel kompleks & bebas typo SMKS
  const roundTripped = buildDocModelFromSmartPagesData({
    pages: converted.pages,
    schoolProfile: DEFAULT_SCHOOL_PROFILE,
    letterOptions: { paperSize: 'f4', showKop: true }
  });
  const serializedRoundTrip = JSON.stringify(roundTripped);
  assert.ok(!serializedRoundTrip.includes('SMKS'), 'Ejaan harus SMK IT AL-HAWARI, bukan SMKS');
  assert.ok(serializedRoundTrip.includes('SMK IT Al-Hawari'));

  // 3. Ekspor menjadi .docx OOXML dan validasi XML Struktur MS Word 1:1
  const docxBytes = await renderDocx(roundTripped, 'uint8array');
  const zip = await readZipPackage(docxBytes);
  const docXmlStr = await zip.file('word/document.xml').async('string');
  const parser = new XmlDomParser();
  const xmlDoc = parser.parseFromString(docXmlStr, 'text/xml');
  assert.ok(xmlDoc && xmlDoc.documentElement);

  const pageBreaks = docXmlStr.match(/<w:br w:type="page"\/>/g) || [];
  assert.equal(pageBreaks.length, 2, 'Harus memiliki tepat 2 page break (3 halaman utuh)');
  assert.ok(docXmlStr.includes('<w:gridSpan w:val="2"/>'), 'Header Berangkat/Datang harus memakai colspan=2');
  assert.ok(docXmlStr.includes('<w:gridSpan w:val="3"/>'), 'Header Berangkat lagi harus memakai colspan=3');
  assert.ok(docXmlStr.includes('<w:vMerge w:val="restart"/>'), 'Header 2 baris harus memakai rowspan restart');
  assert.ok(docXmlStr.includes('<w:vMerge/>'), 'Baris ke-2 header harus memakai rowspan continuation');
  assert.ok(docXmlStr.includes('<w:left w:val="none"'), 'Tabel SPPD 8 poin harus tanpa garis tepi kiri luar');
  assert.ok(docXmlStr.includes('<w:right w:val="none"'), 'Tabel SPPD 8 poin harus tanpa garis tepi kanan luar');
  assert.ok(docXmlStr.includes('<w:trHeight w:val="2552"'), 'Baris stamp Hal 3 I harus punya minHeight 4.5cm');
  assert.ok(docXmlStr.includes('<w:trHeight w:val="2722"'), 'Baris stamp Hal 3 II harus punya minHeight 4.8cm');
  assert.ok(docXmlStr.includes('SURAT TUGAS'));
  assert.ok(docXmlStr.includes('CATATAN:'));

  // 4. Uji mode TTD QR Barcode (menanamkan media/qr.png pada Hal 1 & Hal 2 SPPD)
  const minimalPngDataUrl = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==';
  const qrModel = applySignatureModeToDocModel(roundTripped, {
    mode: 'qr',
    qrDataUrl: minimalPngDataUrl,
    verificationUrl: 'https://copilot-dokumen-alhawari.netlify.app/verify?id=test-st-sppd'
  });
  const qrDocxBytes = await renderDocx(qrModel, 'uint8array');
  const qrZip = await readZipPackage(qrDocxBytes);
  assert.ok(qrZip.file('word/media/qr.png'), 'File media/qr.png wajib tertanam pada paket .docx mode QR');
  const qrDocXml = await qrZip.file('word/document.xml').async('string');
  const drawingMatches = qrDocXml.match(/<w:drawing>/g) || [];
  assert.ok(drawingMatches.length >= 2, 'QR wajib muncul di Surat Tugas (Hal 1) maupun slot QR tabel SPPD (Hal 2)');

  // 5. Pastikan AI converter & Gudang Surat mengenal template Surat Tugas + SPPD
  const aiPresetModel = aiJsonToSmartDocModel({
    templatePreset: 'surat_tugas_sppd',
    letterTitle: 'SURAT TUGAS & SPPD',
    formData: { tujuanLokasi: 'PT MEDIA GARUT', lamaHari: '2' }
  }, DEFAULT_SCHOOL_PROFILE, { paperSize: 'f4' });
  assert.ok(aiPresetModel.blocks.some(b => b.type === 'pageBreak'));
  assert.ok(aiPresetModel.blocks.some(b => b.type === 'table' && b.borderStyle === 'sppd'));

  const vaultEntry = store.getLetterVault().find(v => v.id === 'vault-ref-004');
  assert.ok(vaultEntry, 'Gudang Surat bawaan harus menyediakan Surat Tugas & SPPD (vault-ref-004)');
  assert.equal(store.isSmartTemplateDocument(vaultEntry.document), true);
  assert.equal(store.loadLetterToCanvas('vault-ref-004'), true);
  assert.ok(store.getState().smartDocument?.pendingDocModel, 'Memuat template 3-halaman dari Gudang harus mengisi smartDocument.pendingDocModel');
});

test('Kontrol Interaktif Tabel MS Word: Hitung Ulang Lebar Kolom, Tinggi Baris (cm), & Ekspor Lebar Kustom ke .docx', async () => {
  // 1. Resize batas internal kolom 0 -> menggeser porsi antara kolom 0 dan kolom 1 dengan total tetap 100%
  const resInternal = computeColumnResizeWidths([100, 300, 200], 0, 50, { minColPx: 28 });
  assert.equal(resInternal.widthsPx[0], 150);
  assert.equal(resInternal.widthsPx[1], 250);
  assert.equal(resInternal.widthsPx[2], 200);
  assert.equal(resInternal.totalWidthPx, 600);
  const sumPct = resInternal.widthsPct.reduce((a, b) => a + b, 0);
  assert.ok(Math.abs(sumPct - 100) <= 0.2, `Total persen lebar kolom harus ~100%, actual: ${sumPct}`);

  // 2. Resize batas tepi paling kanan tabel -> mengubah lebar keseluruhan tabel secara proporsional
  const resOuter = computeColumnResizeWidths([200, 400], 1, -150, {
    isOuterRight: true,
    minTotalPx: 140,
    maxTotalPx: 600
  });
  assert.equal(resOuter.totalWidthPx, 450);

  // 3. Konversi presisi tinggi baris px <-> cm
  const fourPointFivePx = cmToPx(4.5);
  assert.equal(pxToCm(fourPointFivePx), 4.5);

  // 4. Round-trip tabel berkustom lebar kolom + lebar tabel (<100%) + identityTable colWidths ke OOXML .docx
  const customDoc = {
    meta: { title: 'Uji Tabel Resizable', author: 'SMK IT Al-Hawari', year: '2026' },
    profile: 'suratDinas',
    paperSize: 'a4',
    blocks: [
      { type: 'letterKop', school: DEFAULT_SCHOOL_PROFILE },
      {
        type: 'identityTable',
        colWidthsPct: [38, 62],
        rows: [
          { label: 'Nama Lengkap', value: ': Taufik Dwi Setyo Ramadhan, S.Pd.' }
        ]
      },
      {
        type: 'table',
        borderStyle: 'grid',
        headerRow: true,
        colWidthsPct: [20, 50, 30],
        tableWidthPct: 80,
        rows: [
          { isHeader: true, cells: ['No', 'Uraian', 'Ket'] },
          { isHeader: false, minHeightCm: 2.5, cells: ['1', 'Kolom hasil resize', 'OK'] }
        ]
      }
    ]
  };

  const smartPages = docModelToSmartPages(customDoc, DEFAULT_SCHOOL_PROFILE);
  assert.ok(smartPages.pages[0].html.includes('data-table-width-pct="80"'));
  assert.ok(smartPages.pages[0].html.includes('data-col-widths="20,50,30"'));
  assert.ok(smartPages.pages[0].html.includes('data-col-widths="38,62"'));

  const rebuilt = buildDocModelFromSmartPagesData({
    pages: smartPages.pages,
    schoolProfile: DEFAULT_SCHOOL_PROFILE,
    letterOptions: { paperSize: 'a4', showKop: true }
  });
  const rebuiltTbl = rebuilt.blocks.find(b => b.type === 'table');
  const rebuiltId = rebuilt.blocks.find(b => b.type === 'identityTable');
  assert.deepEqual(rebuiltTbl.colWidthsPct, [20, 50, 30]);
  assert.equal(rebuiltTbl.tableWidthPct, 80);
  assert.deepEqual(rebuiltId.colWidthsPct, [38, 62]);

  const docxBytes = await renderDocx(rebuilt, 'uint8array');
  const zip = await readZipPackage(docxBytes);
  const xmlStr = await zip.file('word/document.xml').async('string');
  const xmlDoc = new XmlDomParser().parseFromString(xmlStr, 'text/xml');
  assert.ok(xmlDoc && xmlDoc.documentElement);
  assert.ok(xmlStr.includes('<w:tblLayout w:type="fixed"/>'), 'Tabel dengan kolom terukur harus mengekspor w:tblLayout fixed');
  assert.ok(xmlStr.includes('<w:trHeight w:val="1418"'), 'Baris dengan tinggi 2.5cm harus mengekspor trHeight=1418 twips');
});

test('Alur Save (Update Asli Tanpa Duplikat) vs Save As (Buat Duplikat Terpisah) di Gudang Surat', () => {
  const initialCount = store.getLetterVault().length;

  // 1. Buka file Gudang Surat ke Word Pro -> activeSmartVaultId tercatat
  assert.equal(store.loadLetterToCanvas('vault-ref-004'), true);
  assert.equal(store.getActiveVaultId('smart'), 'vault-ref-004');
  const activeSmart = store.getActiveVaultEntry('smart');
  assert.equal(activeSmart.id, 'vault-ref-004');

  // 2. Simulasikan Save (Ctrl+S): memperbarui item asli in-place tanpa menambah jumlah item di Gudang
  const updated = store.updateLetterVaultEntry('vault-ref-004', {
    perihal: 'Perihal Hasil Edit Save Langsung'
  });
  assert.ok(updated);
  assert.equal(store.getLetterVault().length, initialCount, 'Save pada dokumen gudang aktif tidak boleh membuat duplikat');
  assert.equal(store.getActiveVaultEntry('smart').perihal, 'Perihal Hasil Edit Save Langsung');

  // 3. Deteksi judul kembar (pencegahan duplikasi tak sengaja)
  const dupHit = store.findLetterVaultByTitle('surat tugas & sppd penjemputan pkl (3 halaman word)');
  assert.ok(dupHit && dupHit.id === 'vault-ref-004');

  // 4. Simulasikan Save As (Ctrl+Shift+S / F12): membuat salinan independen dan memindahkan activeVaultId ke salinan
  const copyEntry = store.addLetterVaultEntry({
    title: 'Salinan - SURAT TUGAS KHUSUS',
    category: 'Surat Tugas & SPPD',
    document: updated.document
  });
  store.setActiveVaultId(copyEntry.id, 'smart');
  assert.equal(store.getLetterVault().length, initialCount + 1);
  assert.equal(store.getActiveVaultId('smart'), copyEntry.id);

  // Bersihkan item uji agar vault bawaan rapi kembali
  store.removeLetterVaultEntry(copyEntry.id);
  assert.equal(store.getActiveVaultId('smart'), null);
});

test('Word Pro: Unwrap container <div> agar tabel di dalam blok contenteditable tidak berubah menjadi teks datar', () => {
  const wrappedHtml = `
    <div>
      <p>Paragraf pembuka surat tugas</p>
      <table class="smart-word-table" data-col-widths="10,50,40">
        <thead><tr><th>No</th><th>Nama</th><th>Jabatan</th></tr></thead>
        <tbody><tr><td>1</td><td>Taufik Dwi Setyo Ramadhan, S.Pd.</td><td>Guru</td></tr></tbody>
      </table>
      <div><p>Paragraf penutup</p></div>
    </div>
  `;
  const blocks = parseSmartPageHtmlToBlocks(wrappedHtml, DEFAULT_SCHOOL_PROFILE);
  const tbl = blocks.find(b => b.type === 'table');
  assert.ok(tbl, 'Tabel di dalam wrapper <div> wajib tetap terbaca sebagai block type="table"');
  assert.deepEqual(tbl.colWidthsPct, [10, 50, 40]);
  assert.equal(blocks.filter(b => b.type === 'paragraph').length, 2);
});

test('Preservasi Spasi, &nbsp;, dan Tab saat Simpan ke Gudang -> Buka Ulang -> Ekspor .docx (Identitas Rapi)', async () => {
  const editedHtml = `
    <h2 class="smart-doc-title" data-smart-role="title" style="text-align: center;"><u>SURAT TUGAS</u></h2>
    <p class="smart-doc-nomor" data-smart-role="nomor" style="text-align: center;">Nomor: 800/013/ST.SMKITA/XI/2026</p>
    <p>Yang bertanda tangan dibawah ini:</p>
    <p>Nama&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: Dian Nurhamdani, S.Pd.</p>
    <p>Jabatan    : Kepala Sekolah</p>
    <p>Unit Kerja : SMKS IT AL-HAWARI</p>
    <p>Alamat\t: Kp. Curug RT/RW 002/006 Garut</p>
    <p>&nbsp;&nbsp;&nbsp;&nbsp;Dengan hormat, disampaikan surat tugas resmi.</p>
    <table class="smart-identity-table" data-smart-role="identity-table" data-col-widths="32,68"><colgroup><col style="width: 32%;"><col style="width: 68%;"></colgroup><tbody><tr><td>Nama   </td><td>   : Siti Aminah</td></tr></tbody></table>
  `;

  // 1. Konversi saat Klik Simpan (HTML -> DocumentModel)
  const savedDocModel = buildDocModelFromSmartPagesData({
    pages: [{ id: 'p-1', html: editedHtml, isManualBreak: false }],
    schoolProfile: DEFAULT_SCHOOL_PROFILE,
    letterOptions: { paperSize: 'f4', showKop: true }
  });

  // 2. Simpan ke Gudang Surat & Buka kembali ke Word Pro
  const entry = store.addLetterVaultEntry({
    title: 'Uji Preservasi Tab Identitas',
    category: 'Surat Tugas',
    document: savedDocModel,
    preferSmartCanvas: true
  });

  try {
    assert.equal(store.loadLetterToCanvas(entry.id), true);
    const pending = store.getState().smartDocument?.pendingDocModel;
    assert.ok(pending, 'Pending DocModel harus tersedia saat dibuka dari Gudang');

    const reopened = docModelToSmartPages(pending, DEFAULT_SCHOOL_PROFILE);
    const reopenedHtml = reopened.pages[0].html;

    // Pastikan seluruh spasi/tab tidak ikut menciut jadi spasi tunggal
    assert.ok(reopenedHtml.includes('Nama&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;: Dian Nurhamdani, S.Pd.'), '7x &nbsp; harus terjaga penuh');
    assert.ok(reopenedHtml.includes('Jabatan&nbsp;&nbsp;&nbsp; : Kepala Sekolah'), '4x spasi berturut-turut harus terjaga sebagai spasi non-collapsing');
    assert.ok(reopenedHtml.includes('Alamat\t: Kp. Curug'), 'Karakter tab literal harus tetap utuh');
    assert.ok(reopenedHtml.includes('&nbsp;&nbsp;&nbsp;&nbsp;Dengan hormat,'), 'Indentasi awal paragraf harus tetap terjaga');
    assert.ok(reopenedHtml.includes('Nama&nbsp;&nbsp; ') && reopenedHtml.includes('&nbsp;&nbsp; : Siti Aminah'), 'Spasi manual dalam identityTable harus terjaga');

    // 3. Pastikan ekspor .docx menyertakan xml:space="preserve" dan <w:tab/>
    const docxBytes = await renderDocx(pending, 'uint8array');
    const zip = await readZipPackage(docxBytes);
    const xmlStr = await zip.file('word/document.xml').async('string');
    assert.ok(xmlStr.includes('xml:space="preserve"'), 'Multiple spaces wajib memakai xml:space="preserve"');
    assert.ok(xmlStr.includes('<w:tab/>'), 'Karakter tab wajib menjadi <w:tab/> di OOXML');
  } finally {
    store.removeLetterVaultEntry(entry.id);
    store.setActiveVaultId(null, 'smart');
  }
});
