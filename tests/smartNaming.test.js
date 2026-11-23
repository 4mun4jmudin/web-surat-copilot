/**
 * tests/smartNaming.test.js - Pengujian Penamaan Surat Keluar & Gudang Surat Otomatis
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {
  extractDocumentMetadata,
  buildFormattedDocumentTitle,
  buildDocumentDownloadFilename,
  normalizeIndonesianDateText,
  enhanceDocumentMetadataWithAi
} from '../js/letters/smartNaming.js';
import { buildSuratTugasSppdDocModel } from '../js/letters/templates/suratTugasSppd.js';
import { DEFAULT_SCHOOL_PROFILE, SCHOOL_UNITS } from '../js/letters/schoolProfile.js';
import { store } from '../js/state.js';

test('SmartNaming: mengekstrak Surat Tugas + SPPD (orang ditugaskan, keperluan, tanggal) tanpa tertukar Kepsek', () => {
  const sppdDoc = buildSuratTugasSppdDocModel(
    {
      nomorSurat: '800/014/ST.SMKITA/XI/2026',
      tanggalSurat: 'Garut, 23 November 2026',
      tujuanLokasi: 'GOLDEN TULIP ESSENTIAL TANGERANG',
      maksud: 'Penjemputan Peserta Didik PKL',
      tanggalBerangkat: '23 November 2026',
      assignees: [
        { nama: 'Taufik Dwi Setyo Ramadhan, S.Pd.', jabatan: 'Guru' },
        { nama: 'Siti Laras Maemunah, S.Pd.', jabatan: 'Guru' }
      ]
    },
    DEFAULT_SCHOOL_PROFILE,
    { paperSize: 'f4' }
  );

  const meta = extractDocumentMetadata({
    docModel: sppdDoc,
    schoolProfile: DEFAULT_SCHOOL_PROFILE
  });

  assert.equal(meta.typeCode, 'ST');
  assert.equal(meta.category, 'Surat Tugas');
  assert.ok(
    meta.recipientName.includes('Taufik Dwi Setyo Ramadhan') && meta.recipientName.includes('Siti Laras Maemunah'),
    `Harus menangkap guru yang ditugaskan, bukan Kepsek atau Umum (dapat: ${meta.recipientName})`
  );
  assert.ok(!meta.recipientName.includes('Dian Nurhamdani'), 'Nama Kepsek penandatangan tidak boleh dianggap penerima tugas');
  assert.match(meta.tanggalSurat, /23 November 2026/);
  assert.ok(meta.formattedTitle.includes('Taufik Dwi Setyo Ramadhan'), 'Judul terformat harus menyertakan nama guru');
  assert.ok(meta.formattedTitle.includes('23 Nov 2026'), 'Judul terformat harus menyertakan tanggal ringkas');
});

test('SmartNaming: mengekstrak Surat Keterangan SMP Word Pro & menormalisasi judul dan nama file .docx', () => {
  const smpPages = [
    {
      id: 'p1',
      html: `
        <h2 class="smart-doc-title" style="text-align:center;">SURAT KETERANGAN PENERIMA BANTUAN</h2>
        <p class="smart-doc-nomor" style="text-align:center;">Nomor: 421.3/015/KET.SMPITA/XI/2026</p>
        <p>Yang bertanda tangan di bawah ini, Kepala SMP IT Al-Hawari menerangkan bahwa:</p>
        <table><tbody>
          <tr><td>Nama</td><td>: Ujang</td></tr>
          <tr><td>Kelas</td><td>: 10 DPIB</td></tr>
        </tbody></table>
        <p>Surat keterangan ini dibuat untuk dipergunakan sebagai syarat administrasi pengajuan bantuan/beasiswa.</p>
        <p>Garut, 22 November 2026</p>
        <p>Kepala Sekolah</p>
        <p><strong>Yangyang Budiman, S.Pd.</strong></p>
      `
    }
  ];

  const meta = extractDocumentMetadata({
    smartPages: smpPages,
    schoolProfile: { ...SCHOOL_UNITS.smp }
  });

  assert.equal(meta.typeCode, 'KET');
  assert.equal(meta.category, 'Surat Keterangan');
  assert.equal(meta.recipientName, 'Ujang');
  assert.match(meta.keperluan, /Bantuan|Beasiswa/i);
  assert.equal(meta.tanggalSurat, '22 November 2026');
  assert.ok(meta.formattedTitle.includes('Ujang') && meta.formattedTitle.includes('22 Nov 2026'));

  const outFile = buildDocumentDownloadFilename({
    kind: 'outgoing',
    sequenceNumber: 15,
    typeCode: meta.typeCode,
    title: meta.formattedTitle,
    perihal: meta.perihal,
    recipientName: meta.recipientName,
    tanggalSurat: meta.tanggalSurat,
    unitId: 'smp',
    schoolProfile: SCHOOL_UNITS.smp
  });
  assert.match(outFile, /^015_KET_Ujang_.*22-Nov-2026_SMP\.docx$/);
  assert.ok(!/[\\/:*?"<>|]/.test(outFile), 'Nama file .docx harus bebas karakter ilegal OS');

  const vaultFile = buildDocumentDownloadFilename({
    kind: 'vault',
    typeCode: meta.typeCode,
    title: meta.formattedTitle,
    perihal: meta.perihal,
    recipientName: meta.recipientName,
    tanggalSurat: meta.tanggalSurat,
    unitId: 'smk',
    schoolProfile: SCHOOL_UNITS.smk
  });
  assert.match(vaultFile, /^Surat-Keterangan.*_Ujang_.*22-Nov-2026_SMK\.docx$/);
});

test('SmartNaming: file .docx impor bergaya bebas di Gudang tetap mendapat judul, kategori & pihak rapih', () => {
  const uploadedDocx = {
    fileName: 'doc_final(2)_revisi.docx',
    paragraphs: [
      { currentText: 'SURAT UNDANGAN' },
      { currentText: 'Nomor: 005/008/UND.SMKITA/XI/2026' },
      { currentText: 'Perihal: Rapat Pleno Persiapan PKL Tahun 2026' },
      { currentText: 'Kepada Yth. Bapak/Ibu Orang Tua Wali Siswa Kelas XI' },
      { currentText: 'Garut, 20 November 2026' }
    ]
  };

  const meta = extractDocumentMetadata({
    uploadedDocx,
    schoolProfile: DEFAULT_SCHOOL_PROFILE,
    fallbackTitle: 'doc_final(2)_revisi'
  });

  assert.equal(meta.typeCode, 'UND');
  assert.equal(meta.category, 'Surat Undangan');
  assert.match(meta.perihal, /Rapat Pleno Persiapan PKL/i);
  assert.match(meta.recipientName, /Orang Tua Wali Siswa/i);
  assert.ok(!meta.formattedTitle.includes('doc_final'), 'Tidak boleh memakai nama file mentah jika isi surat terbaca');
});

test('SmartNaming: normalisasi tanggal & pemindaian AI dengan guardrail mencegah Kepsek menjadi penerima ST', async () => {
  assert.deepEqual(normalizeIndonesianDateText('2026-11-23'), {
    full: '23 November 2026',
    short: '23 Nov 2026'
  });
  assert.deepEqual(normalizeIndonesianDateText('23, November 2026.'), {
    full: '23 November 2026',
    short: '23 Nov 2026'
  });
  const formatted = buildFormattedDocumentTitle({
    typeCode: 'ST',
    jenisSurat: 'Surat Tugas',
    recipientName: 'Siti Laras Maemunah, S.Pd.',
    keperluan: 'Monitoring PKL',
    tanggalSurat: '24 November 2026'
  });
  assert.equal(formatted, 'Surat Tugas — Siti Laras Maemunah, S.Pd. — Monitoring PKL (24 Nov 2026)');

  // Mock AI aktif + respons JSON
  const prevFetch = globalThis.fetch;
  const prevVault = store.getState().apiVault;
  const prevId = store.getState().activeApiId;
  store.setState({
    apiVault: [{ id: 'mock-gem', provider: 'gemini', model: 'gemini-2.5-flash', apiKey: 'AIzaTestKey123456' }],
    activeApiId: 'mock-gem'
  }, false);

  globalThis.fetch = async () => new Response(JSON.stringify({
    candidates: [{
      content: {
        parts: [{
          text: JSON.stringify({
            typeCode: 'ST',
            jenisSurat: 'Surat Tugas & SPPD',
            recipientName: 'Dian Nurhamdani, S.Pd.', // sengaja kembalikan Kepsek -> harus ditolak guardrail sistem
            keperluan: 'Koordinasi Lomba LKS Provinsi',
            tanggalSurat: '25 November 2026'
          })
        }]
      }
    }]
  }), { status: 200, headers: { 'Content-Type': 'application/json' } });

  try {
    const aiMeta = await enhanceDocumentMetadataWithAi({
      docModel: {
        blocks: [
          { type: 'letterTitle', title: 'SURAT TUGAS', nomor: 'Nomor: 800/020/ST.SMKITA/XI/2026' },
          {
            type: 'table',
            rows: [
              [{ text: 'No' }, { text: 'Nama' }, { text: 'Jabatan' }],
              [{ text: '1' }, { text: 'Rina Marlina, S.Pd.' }, { text: 'Guru Produktif' }]
            ]
          },
          { type: 'paragraph', text: 'Untuk mengikuti Lomba LKS Provinsi.' }
        ]
      },
      schoolProfile: DEFAULT_SCHOOL_PROFILE
    });

    assert.equal(aiMeta.scannedByAi, true);
    assert.equal(aiMeta.recipientName, 'Rina Marlina, S.Pd.', 'Guardrail harus mempertahankan nama guru, tidak tertukar nama Kepsek');
    assert.equal(aiMeta.keperluan, 'Koordinasi Lomba LKS Provinsi');
    assert.equal(aiMeta.tanggalSurat, '25 November 2026');
    assert.ok(aiMeta.formattedTitle.includes('Koordinasi Lomba LKS Provinsi'));
  } finally {
    globalThis.fetch = prevFetch;
    store.setState({ apiVault: prevVault, activeApiId: prevId }, false);
  }
});

test('Responsive UI: tabel Surat Keluar & Gudang Surat memakai full-screen layout dan data-label card', () => {
  const css = fs.readFileSync('css/app.css', 'utf8');
  const outJs = fs.readFileSync('js/ui/outgoingLettersView.js', 'utf8');
  const vaultJs = fs.readFileSync('js/ui/vaultView.js', 'utf8');

  assert.match(css, /\.view-shell\.view-shell-fluid/);
  assert.match(css, /\.enterprise-table\.responsive-data-table\s+td::before/);
  assert.match(outJs, /view-shell-fluid/);
  assert.match(outJs, /responsive-data-table/);
  assert.match(outJs, /data-label="Nomor &amp; Status Surat"/);
  assert.match(vaultJs, /view-shell-fluid/);
  assert.match(vaultJs, /responsive-data-table/);
  assert.match(vaultJs, /btn-vault-normalize-all/);
});
