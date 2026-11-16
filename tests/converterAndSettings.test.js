/**
 * tests/converterAndSettings.test.js - Unit Test Konversi Docx Impor, AI Document Polisher, & Pengaturan Dokumen Lengkap
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { store } from '../js/state.js';
import { convertDocxToDocumentModel } from '../js/docx/patch.js';
import { renderDocx } from '../js/docx/render.js';
import { POLISHER_SYSTEM_PROMPT, buildPolisherUserPrompt } from '../js/ai/prompts.js';

test('Konversi Teks Docx Impor: Menghasilkan Document Model Resmi Standar Sekolah', () => {
  const dummyParsedDocx = {
    meta: {
      title: 'SURAT PANGGILAN ORANG TUA SISWA',
      nomor: '0006/03/SMKITA/III/2026',
      date: '14 November 2026',
      signer: 'Dian Nurhadian, S.Pd.'
    },
    structured: {
      title: { text: 'SURAT PANGGILAN ORANG TUA SISWA' },
      nomor: { text: '0006/03/SMKITA/III/2026' },
      bodyParagraphs: [
        { text: 'Sehubungan dengan adanya evaluasi perkembangan kedisiplinan dan kehadiran siswa di sekolah, dengan ini kami mengundang kehadiran Bapak/Ibu orang tua/wali siswa pada:' },
        { text: 'Mengingat pentingnya agenda tersebut, kami sangat mengharapkan kehadiran Bapak/Ibu tepat pada waktu yang telah ditentukan.' }
      ],
      identityRows: [
        { label: 'Hari / Tanggal', value: 'Senin, 16 November 2026' },
        { label: 'Waktu', value: '09.00 WIB s/d Selesai' },
        { label: 'Tempat', value: 'Ruang Bimbingan Konseling (BK) SMK IT Al-Hawari' },
        { label: 'Bertemu Dengan', value: 'Wali Kelas & Guru BK' }
      ],
      closing: { text: 'Demikian Surat pemberitahuan ini kami sampaikan untuk dapat diketahui oleh orang tua/wali siswa, atas perhatian dan kerjasamanya diucapkan terima kasih.' },
      signature: {
        city: 'Garut',
        date: '14 November 2026',
        title: 'Kepala Sekolah',
        signer: 'Dian Nurhadian, S.Pd.',
        nip: '-'
      }
    }
  };

  const convertedDoc = convertDocxToDocumentModel(dummyParsedDocx, store.getSchoolProfile(), {
    paperSize: 'a4',
    orientation: 'portrait'
  });

  assert.ok(convertedDoc, 'Document Model harus berhasil dibuat');
  assert.equal(convertedDoc.blocks[0].type, 'letterKop', 'Blok pertama harus berupa kop surat');
  assert.equal(convertedDoc.blocks[1].type, 'letterTitle', 'Blok kedua harus berupa judul surat');
  assert.equal(convertedDoc.blocks[1].title, 'SURAT PANGGILAN ORANG TUA SISWA');
  assert.equal(convertedDoc.blocks[1].nomor, 'Nomor: 0006/03/SMKITA/III/2026');

  // Cek tabel identitas / jadwal
  const tableBlock = convertedDoc.blocks.find(b => b.type === 'identityTable');
  assert.ok(tableBlock, 'Harus memiliki blok tabel jadwal');
  assert.equal(tableBlock.rows.length, 4, 'Harus memiliki 4 baris rincian jadwal');
  assert.equal(tableBlock.rows[0].label, 'Hari / Tanggal');
  assert.equal(tableBlock.rows[0].value, ': Senin, 16 November 2026');

  // Cek blok tanda tangan
  const sigBlock = convertedDoc.blocks.find(b => b.type === 'letterSignature');
  assert.ok(sigBlock, 'Harus memiliki blok tanda tangan');
  assert.equal(sigBlock.name, 'Dian Nurhadian, S.Pd.');
  assert.equal(sigBlock.city, 'Garut');
});

test('Pengaturan Dokumen: Pengelolaan Ukuran Kertas, Orientasi, dan Margin', async () => {
  // 1. Simpan pengaturan baru
  store.updateLetterOptions({
    paperSize: 'legal',
    orientation: 'landscape',
    margins: { top: 2.5, bottom: 2.5, left: 3.0, right: 2.5 }
  });

  const opts = store.getLetterOptions();
  assert.equal(opts.paperSize, 'legal');
  assert.equal(opts.orientation, 'landscape');
  assert.equal(opts.margins.left, 3.0);

  // 2. Render ke docx untuk memastikan renderDocx mendukung Legal & Lanskap
  const currentDoc = store.getState().document;
  const bytes = await renderDocx(currentDoc, 'uint8array');
  assert.ok(bytes && bytes.length > 0, 'Harus berhasil merender berkas Word dengan ukuran Legal Lanskap');

  // Reset kembali ke default A4 potret dinas
  store.updateLetterOptions({
    paperSize: 'a4',
    orientation: 'portrait',
    margins: { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 }
  });
  assert.equal(store.getLetterOptions().paperSize, 'a4');
  assert.equal(store.getLetterOptions().orientation, 'portrait');
});

test('AI Document Polisher: Prompt Sistem dan Pembentukan User Prompt', () => {
  assert.ok(POLISHER_SYSTEM_PROMPT.includes('AI Document Polisher'), 'Prompt sistem harus terdefinisi');
  assert.ok(POLISHER_SYSTEM_PROMPT.includes('identityTable'), 'Skema output harus memuat identityTable');

  const prompt = buildPolisherUserPrompt({
    messyText: 'surat panggilan ortu siswa ahmad fauzi bolos 3 hari garut 14 nov 2026',
    schoolName: 'SMK IT Al-Hawari',
    principalName: 'Dian Nurhadian, S.Pd.',
    dateStr: '14 November 2026'
  });

  assert.ok(prompt.includes('ahmad fauzi'), 'Prompt pengguna harus menyertakan teks berantakan');
  assert.ok(prompt.includes('SMK IT Al-Hawari'), 'Prompt pengguna harus menyertakan nama sekolah');
});
