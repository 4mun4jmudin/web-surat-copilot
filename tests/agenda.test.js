import { test } from 'node:test';
import assert from 'node:assert/strict';
import { generateLetterNumber, toRomanMonth, formatSequenceNumber } from '../js/letters/numbering.js';
import { addAgendaEntry, removeAgendaEntry, filterAgendaList, exportAgendaToCsv } from '../js/letters/agenda.js';

test('Penomoran Otomatis: Konversi bulan Romawi dan padding nomor urut', () => {
  assert.equal(toRomanMonth(1), 'I');
  assert.equal(toRomanMonth(5), 'V');
  assert.equal(toRomanMonth(11), 'XI');
  assert.equal(toRomanMonth(12), 'XII');

  assert.equal(formatSequenceNumber(1, 3), '001');
  assert.equal(formatSequenceNumber(85, 3), '085');
  assert.equal(formatSequenceNumber(125, 3), '125');
});

test('Penomoran Otomatis: Menghasilkan nomor surat sesuai pola format dinamis', () => {
  const result = generateLetterNumber({
    pattern: '{kode}/{urut}/SMKITA/{bulanRomawi}/{tahun}',
    counter: 85,
    templateId: 'surat_keterangan',
    date: new Date(2026, 10, 14) // 14 November 2026 (bulan 11 / index 10)
  });

  assert.equal(result, '421.3/085/SMKITA/XI/2026');
});

test('Buku Agenda: Menambahkan, memfilter, dan menghapus entri agenda surat keluar', () => {
  let agenda = [];

  // 1. Tambah entri pertama
  const res1 = addAgendaEntry(agenda, {
    nomorSurat: '421.3/001/SMKN.1/XI/2026',
    tanggalSurat: '14 November 2026',
    jenisSurat: 'Surat Keterangan',
    perihal: 'Keterangan Siswa Aktif Beasiswa',
    tujuan: 'Ahmad Fauzi (XII RPL 2)',
    penandatangan: 'Drs. H. Eko Wahyudi, M.Pd.'
  });
  agenda = res1.updatedList;
  assert.equal(agenda.length, 1);
  assert.equal(agenda[0].noUrut, 1);
  assert.equal(agenda[0].tujuan, 'Ahmad Fauzi (XII RPL 2)');

  // 2. Tambah entri kedua
  const res2 = addAgendaEntry(agenda, {
    nomorSurat: '800/002/SMKN.1/XI/2026',
    tanggalSurat: '15 November 2026',
    jenisSurat: 'Surat Tugas',
    perihal: 'Tugas Pelatihan Kurikulum',
    tujuan: 'Budi Raharjo S.Pd',
    penandatangan: 'Drs. H. Eko Wahyudi, M.Pd.'
  });
  agenda = res2.updatedList;
  assert.equal(agenda.length, 2);

  // 3. Filter pencarian
  const filteredFauzi = filterAgendaList(agenda, 'Fauzi');
  assert.equal(filteredFauzi.length, 1);
  assert.equal(filteredFauzi[0].tujuan, 'Ahmad Fauzi (XII RPL 2)');

  const filteredTugas = filterAgendaList(agenda, 'Tugas');
  assert.equal(filteredTugas.length, 1);

  // 4. Hapus entri
  const targetId = res1.newItem.id;
  const afterDelete = removeAgendaEntry(agenda, targetId);
  assert.equal(afterDelete.length, 1);
  assert.equal(afterDelete[0].nomorSurat, '800/002/SMKN.1/XI/2026');
});

test('Buku Agenda: Ekspor CSV menyertakan UTF-8 BOM dan format tabel Excel yang valid', () => {
  const sampleList = [
    {
      noUrut: 1,
      nomorSurat: '421.3/001/SMKN.1/XI/2026',
      tanggalSurat: '14 November 2026',
      jenisSurat: 'Surat Keterangan',
      perihal: 'Beasiswa, KIP Kuliah',
      tujuan: 'Siti Aminah',
      penandatangan: 'Kepala Sekolah',
      createdAt: '2026-11-14T01:00:00.000Z'
    }
  ];

  const csv = exportAgendaToCsv(sampleList, 'SMK Negeri 1 Surabaya');
  
  // Pastikan dimulai dengan UTF-8 BOM (\uFEFF)
  assert.ok(csv.startsWith('\uFEFF'), 'CSV harus memiliki UTF-8 BOM untuk kompatibilitas Excel');
  assert.match(csv, /BUKU AGENDA SURAT KELUAR - SMK NEGERI 1 SURABAYA/);
  assert.match(csv, /No\. Urut/);
  assert.match(csv, /421\.3\/001\/SMKN\.1\/XI\/2026/);
  assert.match(csv, /"Beasiswa, KIP Kuliah"/);
});
