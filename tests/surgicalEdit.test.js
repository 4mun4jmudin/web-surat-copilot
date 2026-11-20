import test from 'node:test';
import assert from 'node:assert/strict';

import { parseHtmlNodes, serializeHtmlNodes } from '../js/utils/htmlAst.js';
import {
  shouldAllowLayoutRestructure,
  buildSurgicalIndexFromSmartPages,
  applySurgicalPatchesToSmartPages,
  buildSurgicalIndexFromDocModel,
  applySurgicalPatchesToDocModel
} from '../js/ai/surgicalEdit.js';
import {
  docModelToSmartPages,
  buildDocModelFromSmartPagesData
} from '../js/ui/smartWordCanvas.js';
import { buildSuratTugasSppdDocModel } from '../js/letters/templates/suratTugasSppd.js';
import { DEFAULT_SCHOOL_PROFILE } from '../js/letters/schoolProfile.js';

test('shouldAllowLayoutRestructure mengunci tata letak pada edit data & mengizinkan hanya bila diminta eksplisit', () => {
  assert.equal(shouldAllowLayoutRestructure('Ganti nama guru menjadi Ahmad Fauzi dan tujuan PT Telkom Bandung'), false);
  assert.equal(shouldAllowLayoutRestructure('Tambahkan 1 baris pegawai atas nama Rina di tabel Surat Tugas'), false);
  assert.equal(shouldAllowLayoutRestructure('Perbaiki kata dan ejaan PUEBI tanpa mengubah rincian tabel'), false);
  assert.equal(shouldAllowLayoutRestructure('Ubah tata letak tabel SPPD jadi 3 kolom dan desain ulang surat'), true);
  assert.equal(shouldAllowLayoutRestructure('Pindahkan tabel keterangan ke halaman 1'), true);
});

test('htmlAst menjaga void tag <col> dalam <colgroup> tetap datar & utuh', () => {
  const sample = '<table class="smart-word-table" data-col-widths="12,50,38"><colgroup><col style="width: 12%;"><col style="width: 50%;"><col style="width: 38%;"></colgroup><tbody><tr><td>1</td><td>A</td><td>Guru</td></tr></tbody></table>';
  const ast = parseHtmlNodes(sample);
  const out = serializeHtmlNodes(ast);
  assert.ok(out.includes('<colgroup><col style="width: 12%;"><col style="width: 50%;"><col style="width: 38%;"></colgroup>'));
});

test('Surgical Patch Word Pro menjaga 100% tata letak, lebar kolom kustom, tinggi baris, merge header, & slot QR pada SPPD 3 halaman', () => {
  const sppdModel = buildSuratTugasSppdDocModel({
    nomorSurat: '800/005/ST.SMKITA/XI/2026',
    assignees: [
      { nama: 'Taufik Dwi Setyo Ramadhan, S.Pd.', jabatan: 'Guru' },
      { nama: 'Siti Laras Maemunah, S.Pd.', jabatan: 'Guru' }
    ],
    tujuanLokasi: 'DRAGON MOTION PICTURE ASIA',
    tempatTiba: 'Dragon Motion Picture',
    maksudTugas: 'Penjemputan Peserta Didik PKL',
    tglBerangkat: '18 November 2026',
    tglKembali: '18 November 2026',
    tanggalSurat: 'Garut, 18 November 2026'
  }, DEFAULT_SCHOOL_PROFILE, { paperSize: 'f4' });

  const converted = docModelToSmartPages(sppdModel, DEFAULT_SCHOOL_PROFILE);
  assert.equal(converted.pages.length, 3);

  // Simulasikan pengguna telah menggeser lebar kolom tabel pegawai Page 1 & tinggi baris tabel Page 3
  converted.pages[0].html = converted.pages[0].html
    .replace('data-col-widths="9,51,40"', 'data-col-widths="12,50,38"')
    .replace('<col style="width: 9%;"><col style="width: 51%;"><col style="width: 40%;">', '<col style="width: 12%;"><col style="width: 50%;"><col style="width: 38%;">');
  converted.pages[2].html = converted.pages[2].html
    .replace('style="height: 4.5cm;"', 'style="height: 5.2cm;"')
    .replace('data-min-height="4.5"', 'data-min-height="5.2"');

  const indexObj = buildSurgicalIndexFromSmartPages(converted.pages, DEFAULT_SCHOOL_PROFILE);
  assert.equal(indexObj.snapshot.totalPages, 3);

  // Verifikasi bahwa kolom QR pada Page 2 ditandai terkunci di snapshot
  const p2FooterTable = indexObj.snapshot.pages[1].blocks.find((b) => b.type === 'table' && b.borderStyle === 'none');
  assert.ok(p2FooterTable, 'Tabel layout footer SPPD harus terindeks');
  assert.equal(p2FooterTable.rows[0].cells[1].hasQrSlot, true);

  const p1Table = indexObj.snapshot.pages[0].blocks.find((b) => b.type === 'table');
  const p1Sig = indexObj.snapshot.pages[0].blocks.find((b) => b.type === 'signature');
  const p2MainTable = indexObj.snapshot.pages[1].blocks.find((b) => b.type === 'table' && b.borderStyle === 'sppd');
  const p3Ket1 = indexObj.snapshot.pages[2].blocks.find((b) => b.type === 'table');
  const tibaPara = indexObj.snapshot.pages[0].blocks.find((b) => b.type === 'paragraph' && /^Tiba di/i.test(b.text));

  const patchRes = applySurgicalPatchesToSmartPages(converted.pages, {
    editScope: 'surgical',
    summaryOfChanges: 'Ubah nama pegawai, tambah 1 guru, dan ganti lokasi tujuan ke PT Telkom Bandung',
    patches: [
      { op: 'setTableCell', blockId: p1Table.blockId, row: 1, col: 1, text: 'Budi Santoso, S.Pd.' },
      { op: 'addTableRow', blockId: p1Table.blockId, cells: ['3', 'Ahmad Fauzi, S.Pd.', 'Guru'] },
      { op: 'setParagraph', blockId: tibaPara.blockId, text: 'Tiba di PT Telkom Bandung' },
      { op: 'setSignature', blockId: p1Sig.blockId, date: '25 November 2026' },
      { op: 'setTableCell', blockId: p2MainTable.blockId, row: 1, col: 1, text: '1.  Budi Santoso, S.Pd.\n2.  Siti Laras Maemunah, S.Pd.\n3.  Ahmad Fauzi, S.Pd.' },
      { op: 'setTableCell', blockId: p2MainTable.blockId, row: 3, col: 1, text: 'Dari   : SMK IT AL-HAWARI\nKe     : PT TELKOM BANDUNG\nDengan menggunakan kendaraan: Mobil' },
      { op: 'setTableCell', blockId: p2FooterTable.blockId, row: 0, col: 1, paragraphIndex: 0, text: '25 November 2026' },
      { op: 'setTableCell', blockId: p3Ket1.blockId, row: 2, col: 1, text: '25 November 2026.' }
    ]
  }, { schoolProfile: DEFAULT_SCHOOL_PROFILE, allowLayoutRestructure: false });

  assert.equal(patchRes.mode, 'surgical');
  assert.equal(patchRes.appliedCount, 8);
  assert.equal(patchRes.pages.length, 3);

  // Page 1: lebar kolom kustom tetap [12, 50, 38], baris tambah jadi 4 baris, underline "Tiba di" terjaga
  assert.ok(patchRes.pages[0].html.includes('data-col-widths="12,50,38"'));
  assert.ok(patchRes.pages[0].html.includes('<col style="width: 12%;">'));
  assert.ok(patchRes.pages[0].html.includes('Budi Santoso, S.Pd.'));
  assert.ok(patchRes.pages[0].html.includes('Ahmad Fauzi, S.Pd.'));
  assert.ok(patchRes.pages[0].html.includes('Tiba di <u>PT Telkom Bandung</u>'));
  assert.ok(patchRes.pages[0].html.includes('Garut, 25 November 2026'));

  // Page 2: borderStyle sppd & none tetap utuh, CATATAN aman, dan smart-qr-slot TIDAK terhapus!
  assert.ok(patchRes.pages[1].html.includes('smart-table-sppd'));
  assert.ok(patchRes.pages[1].html.includes('smart-table-none'));
  assert.ok(patchRes.pages[1].html.includes('PT TELKOM BANDUNG'));
  assert.ok(patchRes.pages[1].html.includes('data-qr-slot="true"'));
  assert.ok(patchRes.pages[1].html.includes('25 November 2026'));
  assert.ok(patchRes.pages[1].html.includes('CATATAN:'));

  // Page 3: tinggi baris 5.2 cm, colspan="2", rowspan="2" tetap utuh, tidak ada duplikat TTD di akhir Halaman 3
  assert.ok(patchRes.pages[2].html.includes('style="height: 5.2cm;"'));
  assert.ok(patchRes.pages[2].html.includes('colspan="2"'));
  assert.ok(patchRes.pages[2].html.includes('rowspan="2"'));
  assert.ok(patchRes.pages[2].html.includes('25 November 2026.'));
  assert.ok(!patchRes.pages[2].html.includes('data-smart-role="signature"'));

  // Pastikan konversi balik ke DocumentModel tetap valid 100% untuk ekspor .docx
  const finalModel = buildDocModelFromSmartPagesData({ pages: patchRes.pages, schoolProfile: DEFAULT_SCHOOL_PROFILE });
  const tables = finalModel.blocks.filter((b) => b.type === 'table');
  assert.equal(tables.length, 5);
  assert.deepEqual(tables[0].colWidthsPct, [12, 50, 38]);
  assert.equal(tables[0].rows.length, 4);
  assert.equal(tables[1].borderStyle, 'sppd');
  assert.equal(tables[2].borderStyle, 'none');
});

test('Reconciler fallback (format blocks lama) tetap mengunci struktur tabel SPPD & tidak menghapus kolom/merge', () => {
  const sppdModel = buildSuratTugasSppdDocModel({}, DEFAULT_SCHOOL_PROFILE, { paperSize: 'f4' });
  const converted = docModelToSmartPages(sppdModel, DEFAULT_SCHOOL_PROFILE);

  const legacyAiOutput = {
    letterTitle: 'SURAT TUGAS',
    nomorSurat: '800/099/ST.SMKITA/XI/2026',
    blocks: [
      {
        type: 'table',
        rows: [
          ['No', 'Nama', 'Jabatan'],
          ['1', 'Guru Baru Satu', 'Guru'],
          ['2', 'Guru Baru Dua', 'Waka Kurikulum']
        ]
      }
    ]
  };

  const res = applySurgicalPatchesToSmartPages(converted.pages, legacyAiOutput, {
    schoolProfile: DEFAULT_SCHOOL_PROFILE,
    allowLayoutRestructure: false
  });

  assert.equal(res.mode, 'reconciled');
  assert.ok(res.pages[0].html.includes('800/099/ST.SMKITA/XI/2026'));
  assert.ok(res.pages[0].html.includes('Guru Baru Satu'));
  assert.ok(res.pages[0].html.includes('data-col-widths="9,51,40"'));
  // Tabel SPPD Page 2 & Keterangan Page 3 tidak boleh ikut terhapus meskipun legacy output tidak menyertakannya!
  assert.ok(res.pages[1].html.includes('smart-table-sppd'));
  assert.ok(res.pages[2].html.includes('rowspan="2"'));
});

test('Surgical Patch pada DocumentModel (Lembar Kerja) menjaga tabel & pageBreak tetap utuh', () => {
  const docModel = buildSuratTugasSppdDocModel({}, DEFAULT_SCHOOL_PROFILE, { paperSize: 'f4' });
  const idx = buildSurgicalIndexFromDocModel(docModel);
  const firstTable = idx.snapshot.pages[0].blocks.find((b) => b.type === 'table');

  const res = applySurgicalPatchesToDocModel(docModel, {
    editScope: 'surgical',
    patches: [
      { op: 'setField', target: 'nomor', text: '800/077/ST.SMKITA/XI/2026' },
      { op: 'setTableCell', blockId: firstTable.blockId, row: 1, col: 1, text: 'Deni Kurniawan, S.Pd.' }
    ]
  }, { allowLayoutRestructure: false });

  assert.equal(res.mode, 'surgical');
  assert.equal(res.appliedCount, 2);
  assert.equal(res.docModel.blocks.filter((b) => b.type === 'pageBreak').length, 2);
  const tables = res.docModel.blocks.filter((b) => b.type === 'table');
  assert.equal(tables.length, 5);
  assert.equal(tables[0].rows[1][1].text, 'Deni Kurniawan, S.Pd.');
  assert.deepEqual(tables[0].colWidthsPct, [9, 51, 40]);
  assert.equal(tables[1].borderStyle, 'sppd');
  assert.equal(tables[3].rows[2].minHeightCm, 4.5);
});
