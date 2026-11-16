/**
 * tests/verificationQr.test.js - Unit & Integration Tests for Surat Keluar Digital QR Verification (Option B)
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { DOMParser as XmlDomParser } from '@xmldom/xmldom';
import {
  buildVerificationId,
  generateManageToken,
  buildVerificationUrl,
  generateQrPngDataUrl,
  buildSignerSelectionInfo,
  buildQrCaptionLines,
  formatSignerNipLine,
  buildTypeSpecificVerificationDetails,
  buildVerificationPayload,
  applySignatureModeToDocModel
} from '../js/letters/verification.js';
import { buildPrintQrBoxHtml } from '../js/ui/printChoiceModal.js';
import { DEFAULT_SCHOOL_PROFILE } from '../js/letters/schoolProfile.js';
import { suratKeteranganTemplate } from '../js/letters/templates/suratKeterangan.js';
import { renderDocx } from '../js/docx/render.js';
import { readZipPackage } from '../js/utils/zip.js';
import { store } from '../js/state.js';
import { handleVerifyRequest } from '../netlify/functions/verify-letter.mjs';

function createMockBlobStore() {
  const map = new Map();
  return {
    async get(key, options = {}) {
      if (!map.has(key)) return null;
      const val = map.get(key);
      return options?.type === 'json' ? JSON.parse(val) : val;
    },
    async setJSON(key, value, options = {}) {
      if (options?.onlyIfNew && map.has(key)) {
        return { modified: false, etag: 'existing' };
      }
      map.set(key, JSON.stringify(value));
      return { modified: true, etag: 'etag-1' };
    }
  };
}

test('QR PNG Generator menghasilkan data URL PNG valid di Node/Browser tanpa canvas', () => {
  const vId = buildVerificationId('KET', 2026);
  assert.match(vId, /^SMKIT-2026-KET-[A-Z0-9]{8}$/);

  const url = buildVerificationUrl(vId, 'https://copilot-dokumen-alhawari.netlify.app');
  assert.equal(url, `https://copilot-dokumen-alhawari.netlify.app/?v=${vId}`);

  const dataUrl = generateQrPngDataUrl(url);
  assert.ok(dataUrl.startsWith('data:image/png;base64,'));

  const bytes = Buffer.from(dataUrl.split(',')[1], 'base64');
  // Check PNG 8-byte signature
  assert.deepEqual([...bytes.subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10]);
  assert.equal(bytes.subarray(12, 16).toString('ascii'), 'IHDR');
  assert.ok(bytes.includes(Buffer.from('IDAT', 'ascii')));
  assert.ok(bytes.includes(Buffer.from('IEND', 'ascii')));
});

test('Pemilihan Pejabat Pengesahan mendukung Kepala Sekolah, Tata Usaha (TU), dan Kepsek & TU', () => {
  const customSchool = {
    ...DEFAULT_SCHOOL_PROFILE,
    name: 'SMK IT Al-Hawari',
    principalName: 'Dian Nurhadian, S.Pd.',
    principalTitle: 'Kepala Sekolah',
    tuName: 'Hendra Wijaya, S.Kom.',
    tuTitle: 'Kepala Tata Usaha'
  };

  const kepsek = buildSignerSelectionInfo(customSchool, 'kepala_sekolah');
  assert.equal(kepsek.signerRole, 'kepala_sekolah');
  assert.equal(kepsek.signerName, 'Dian Nurhadian, S.Pd.');
  assert.match(kepsek.verificationStatement, /atas nama Kepala Sekolah/i);

  const tu = buildSignerSelectionInfo(customSchool, 'tata_usaha');
  assert.equal(tu.signerRole, 'tata_usaha');
  assert.equal(tu.signerName, 'Hendra Wijaya, S.Kom.');
  assert.equal(tu.signerTitle, 'Kepala Tata Usaha');
  assert.match(tu.verificationStatement, /atas nama Tata Usaha \(TU\)/i);

  const both = buildSignerSelectionInfo(customSchool, 'kepsek_dan_tu');
  assert.equal(both.signerRole, 'kepsek_dan_tu');
  assert.match(both.signerName, /Dian Nurhadian/);
  assert.match(both.signerName, /Hendra Wijaya/);
  assert.match(both.verificationStatement, /Kepala Sekolah.*Tata Usaha \(TU\)/i);

  // Caption QR di bawah barcode hanya "Disahkan secara elektronik" (tanpa duplikasi nama / a.n.), dan NIP kosong tampil "NIP. -"
  assert.equal(formatSignerNipLine('-'), 'NIP. -');
  assert.equal(formatSignerNipLine(''), 'NIP. -');
  assert.equal(formatSignerNipLine('198001012010011001'), 'NIP. 198001012010011001');

  const qrCapKepsek = buildQrCaptionLines({ signerRole: 'kepala_sekolah', school: customSchool });
  assert.equal(qrCapKepsek.line1, 'Disahkan secara elektronik');
  assert.equal(qrCapKepsek.line2, '');

  const qrCapTu = buildQrCaptionLines({ signerRole: 'tata_usaha', school: customSchool });
  assert.equal(qrCapTu.line1, 'Disahkan secara elektronik');
  assert.equal(qrCapTu.line2, '');

  const htmlKepsek = buildPrintQrBoxHtml({
    qrDataUrl: 'data:image/png;base64,AA==',
    signerRole: 'kepala_sekolah',
    school: customSchool
  });
  assert.ok(htmlKepsek.includes('Disahkan secara elektronik'));
  assert.ok(!htmlKepsek.includes('Dian Nurhadian'), 'Nama penanda tangan tidak boleh duplikat di bawah QR');
  assert.ok(!htmlKepsek.includes('a.n.'), 'Caption QR cetak tidak boleh lagi memakai teks a.n.');

  const htmlTu = buildPrintQrBoxHtml({
    qrDataUrl: 'data:image/png;base64,AA==',
    signerRole: 'tata_usaha',
    school: customSchool
  });
  assert.ok(htmlTu.includes('Disahkan secara elektronik'));
  assert.ok(!htmlTu.includes('a.n.'), 'Caption QR TU tidak boleh menampilkan a.n.');
  assert.ok(!htmlTu.includes('Hendra Wijaya') && !htmlTu.includes('Dian Nurhadian'), 'Caption QR TU tidak boleh menampilkan nama di bawah Disahkan secara elektronik');
});

test('Detil Verifikasi menyesuaikan jenis surat (KET, UND, KEP, ST)', () => {
  const ketDetails = buildTypeSpecificVerificationDetails('KET', {
    nomorSurat: '421.3/001/KET.SMKITA/XI/2026',
    title: 'SURAT KETERANGAN SISWA AKTIF',
    perihal: 'Keterangan Aktif Sekolah',
    recipientName: 'Ahmad Fauzi',
    tanggalSurat: '15 November 2026'
  }, DEFAULT_SCHOOL_PROFILE);
  assert.ok(ketDetails.some(d => d.label === 'Subjek Keterangan' && d.value === 'Ahmad Fauzi'));

  const stDetails = buildTypeSpecificVerificationDetails('ST', {
    nomorSurat: '421.3/002/ST.SMKITA/XI/2026',
    title: 'SURAT TUGAS',
    perihal: 'Pelatihan Kurikulum Merdeka',
    recipientName: 'Dewi Sartika, S.Pd.',
    tanggalSurat: '15 November 2026'
  }, DEFAULT_SCHOOL_PROFILE);
  assert.ok(stDetails.some(d => d.label === 'Pihak Ditugaskan / Tujuan' && d.value === 'Dewi Sartika, S.Pd.'));

  const kepDetails = buildTypeSpecificVerificationDetails('KEP', {
    nomorSurat: '421.3/003/KEP.SMKITA/XI/2026',
    title: 'SURAT KEPUTUSAN KEPALA SEKOLAH',
    perihal: 'Pembagian Tugas Mengajar'
  }, DEFAULT_SCHOOL_PROFILE);
  assert.ok(kepDetails.some(d => d.label === 'Status Ketetapan'));
});

test('Render .docx Mode QR menyematkan media/qr.png & rIdQr, sedangkan Mode Manual menjaga ruang TTD pulpen', async () => {
  const baseDoc = suratKeteranganTemplate.generateDocumentModel(
    {
      nomorSurat: '421.3/010/KET.SMKITA/XI/2026',
      tanggalSurat: '15 November 2026',
      namaSiswa: 'Siti Aminah',
      keperluan: 'Pendaftaran Beasiswa'
    },
    DEFAULT_SCHOOL_PROFILE,
    { paperSize: 'f4' }
  );

  const payload = buildVerificationPayload({
    entry: {
      sequenceNumber: 10,
      year: 2026,
      typeCode: 'KET',
      nomorSurat: '421.3/010/KET.SMKITA/XI/2026',
      title: 'SURAT KETERANGAN',
      recipientName: 'Siti Aminah'
    },
    school: DEFAULT_SCHOOL_PROFILE,
    signerRole: 'tata_usaha'
  });
  const qrDataUrl = generateQrPngDataUrl(payload.verificationUrl);

  // 1a. Uji Mode QR sebagai TU (tidak ada nama / a.n. di bawah Disahkan secara elektronik)
  const tuSignerInfo = buildSignerSelectionInfo(DEFAULT_SCHOOL_PROFILE, 'tata_usaha');
  const qrDocModel = applySignatureModeToDocModel(baseDoc, {
    mode: 'qr',
    qrDataUrl,
    verificationUrl: payload.verificationUrl,
    signerRole: 'tata_usaha',
    signerInfo: tuSignerInfo,
    school: DEFAULT_SCHOOL_PROFILE
  });

  // Model asli harus tetap manual (non-destruktif)
  const origSig = baseDoc.blocks.find(b => b.type === 'letterSignature');
  assert.notEqual(origSig.signatureMode, 'qr');

  const qrBytes = await renderDocx(qrDocModel, 'uint8array');
  const qrZip = await readZipPackage(qrBytes);
  assert.ok(qrZip.file('word/media/qr.png'), 'word/media/qr.png harus ada saat mode QR');

  const qrRelsXml = await qrZip.file('word/_rels/document.xml.rels').async('text');
  assert.ok(qrRelsXml.includes('Id="rIdQr"') && qrRelsXml.includes('media/qr.png'));

  const qrDocXml = await qrZip.file('word/document.xml').async('text');
  assert.ok(qrDocXml.includes('r:embed="rIdQr"'));
  assert.ok(qrDocXml.includes('cx="864000" cy="864000"'), 'Ukuran QR .docx harus minimal 2.4x2.4 cm (864000 EMU)');
  assert.ok(qrDocXml.includes('Disahkan secara elektronik'));
  assert.ok(!qrDocXml.includes('a.n.'), 'DOKUMEN .docx tidak boleh mencantumkan a.n.');
  assert.ok(qrDocXml.includes('NIP. -'), 'Jika belum memiliki NIP, baris di bawah nama harus menampilkan NIP. -');
  assert.ok(!qrDocXml.includes('Kepala SMKS IT Al-Hawari') && !qrDocXml.includes('Kepala SMK IT Al-Hawari'), 'Keterangan Kepala SMK IT Al-Hawari tidak boleh muncul di bawah nama TTD');
  assert.equal((qrDocXml.match(/166534/g) || []).length, 1, 'Caption QR hanya 1 baris hijau tanpa duplikasi nama');

  const parsedQrXml = new XmlDomParser().parseFromString(qrDocXml, 'text/xml');
  assert.equal(parsedQrXml.documentElement.localName, 'document');

  // 1b. Uji Mode QR sebagai Kepala Sekolah (tanpa duplikasi nama di caption QR & tetap menampilkan NIP. -)
  const kepsekSignerInfo = buildSignerSelectionInfo(DEFAULT_SCHOOL_PROFILE, 'kepala_sekolah');
  const qrKepsekDocModel = applySignatureModeToDocModel(baseDoc, {
    mode: 'qr',
    qrDataUrl,
    verificationUrl: payload.verificationUrl,
    signerRole: 'kepala_sekolah',
    signerInfo: kepsekSignerInfo,
    school: DEFAULT_SCHOOL_PROFILE
  });
  const qrKepsekBytes = await renderDocx(qrKepsekDocModel, 'uint8array');
  const qrKepsekZip = await readZipPackage(qrKepsekBytes);
  const qrKepsekXml = await qrKepsekZip.file('word/document.xml').async('text');
  assert.ok(qrKepsekXml.includes('Disahkan secara elektronik'));
  assert.ok(qrKepsekXml.includes('NIP. -'));
  assert.ok(!qrKepsekXml.includes('Kepala SMKS IT Al-Hawari') && !qrKepsekXml.includes('Kepala SMK IT Al-Hawari'));
  assert.equal((qrKepsekXml.match(/166534/g) || []).length, 1, 'Mode Kepala Sekolah cukup 1 baris hijau Disahkan secara elektronik agar tidak duplikat nama');
  assert.ok(!qrKepsekXml.includes('a.n.'));

  // 2. Uji Mode Manual Pulpen (kosongan)
  const manualDocModel = applySignatureModeToDocModel(qrDocModel, { mode: 'manual' });
  const manualBytes = await renderDocx(manualDocModel, 'uint8array');
  const manualZip = await readZipPackage(manualBytes);
  assert.equal(manualZip.file('word/media/qr.png'), null, 'word/media/qr.png tidak boleh ada saat mode manual pulpen');

  const manualRelsXml = await manualZip.file('word/_rels/document.xml.rels').async('text');
  assert.ok(!manualRelsXml.includes('rIdQr'));

  const manualDocXml = await manualZip.file('word/document.xml').async('text');
  assert.ok(!manualDocXml.includes('rIdQr'));
  assert.ok(manualDocXml.includes('w:before="720"'), 'Ruang kosong TTD manual pulpen harus dipertahankan');
});

test('Store & Netlify Function /api/verify: POST SAH -> GET -> PATCH revoke/sahkan ulang + Clone bersih', async () => {
  const mockStore = createMockBlobStore();
  const factory = () => mockStore;

  const manageToken = generateManageToken();
  const payload = buildVerificationPayload({
    entry: {
      sequenceNumber: 1,
      year: 2026,
      typeCode: 'UND',
      nomorSurat: '421.3/001/UND.SMKITA/XI/2026',
      title: 'UNDANGAN RAPAT WALI MURID',
      perihal: 'Rapat Evaluasi Semester Ganjil',
      recipientName: 'Orang Tua / Wali Siswa Kelas XII',
      tanggalSurat: '15 November 2026'
    },
    school: { ...DEFAULT_SCHOOL_PROFILE, tuName: 'Hendra Wijaya, S.Kom.' },
    signerRole: 'kepsek_dan_tu',
    manageToken
  });

  // 1. POST buat record SAH
  const postRes = await handleVerifyRequest(
    new Request('https://example.netlify.app/verify-letter', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    }),
    {},
    factory
  );
  assert.equal(postRes.status, 201);
  const postJson = await postRes.json();
  assert.equal(postJson.ok, true);
  assert.equal(postJson.verification.status, 'sah');
  assert.equal(postJson.verification.manageHash, undefined, 'manageHash rahasia tidak boleh bocor ke response publik');

  // 2. GET publik ?v=...
  const getRes = await handleVerifyRequest(
    new Request(`https://example.netlify.app/verify-letter?v=${encodeURIComponent(payload.verificationId)}`, {
      method: 'GET'
    }),
    {},
    factory
  );
  assert.equal(getRes.status, 200);
  const getJson = await getRes.json();
  assert.equal(getJson.verification.status, 'sah');
  assert.equal(getJson.verification.signerRole, 'kepsek_dan_tu');
  assert.match(getJson.verification.verificationStatement, /Kepala Sekolah.*Tata Usaha/i);

  // 3. PATCH dengan token salah harus 403
  const badPatch = await handleVerifyRequest(
    new Request('https://example.netlify.app/verify-letter', {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ verificationId: payload.verificationId, manageToken: 'salah-token', status: 'revoked' })
    }),
    {},
    factory
  );
  assert.equal(badPatch.status, 403);

  // 4. PATCH cabut pengesahan dengan token valid -> 200 & status revoked
  const revokePatch = await handleVerifyRequest(
    new Request('https://example.netlify.app/verify-letter', {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        verificationId: payload.verificationId,
        manageToken,
        status: 'revoked',
        reason: 'Revisi jadwal rapat'
      })
    }),
    {},
    factory
  );
  assert.equal(revokePatch.status, 200);
  const revokedJson = await revokePatch.json();
  assert.equal(revokedJson.verification.status, 'revoked');
  assert.equal(revokedJson.verification.revokedReason, 'Revisi jadwal rapat');

  // 5. GET tidak ditemukan -> 404
  const notFoundRes = await handleVerifyRequest(
    new Request('https://example.netlify.app/verify-letter?v=TIDAK-ADA-999', { method: 'GET' }),
    {},
    factory
  );
  assert.equal(notFoundRes.status, 404);

  // 6. Store persistence & clone cleanup
  const docWithQr = applySignatureModeToDocModel(
    suratKeteranganTemplate.generateDocumentModel({}, DEFAULT_SCHOOL_PROFILE),
    { mode: 'qr', qrDataUrl: generateQrPngDataUrl(payload.verificationUrl) }
  );
  const saved = store.addOutgoingLetter({
    ...payload,
    isVerifiedSah: true,
    verificationStatus: 'sah',
    qrDataUrl: generateQrPngDataUrl(payload.verificationUrl),
    document: docWithQr,
    skipAutoAdvanceNumber: true
  });
  assert.equal(saved.isVerifiedSah, true);

  store.cloneOutgoingToNewDraft(saved.id);
  const clonedStateDoc = store.getState().document;
  assert.equal(clonedStateDoc.verification, undefined, 'Draf hasil salin tidak boleh membawa status SAH lama');
  const clonedSig = clonedStateDoc.blocks.find(b => b.type === 'letterSignature');
  assert.equal(clonedSig.signatureMode, undefined, 'Draf hasil salin tidak boleh membawa QR surat lama');
});
