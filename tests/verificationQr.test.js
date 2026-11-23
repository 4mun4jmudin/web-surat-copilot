/**
 * tests/verificationQr.test.js - Unit & Integration Tests for Surat Keluar Digital QR Verification (Option B)
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
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
import { renderReadOnlyDocumentHtml } from '../js/ui/outgoingLettersView.js';
import {
  computePrintableDimsMm,
  normalizeQrPlacement,
  findPreferredSignatureQrPlacement,
  calibratePlacementFromDom,
  buildFloatingQrElementHtml,
  buildFloatingQrContextMenuHtml,
  mountInteractiveFloatingQr
} from '../js/ui/qrFloatingHelper.js';
import { attachInteractiveTables } from '../js/ui/tableInteraction.js';
import { buildSuratTugasSppdDocModel } from '../js/letters/templates/suratTugasSppd.js';
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

test('Pratinjau 1:1 Surat Keluar & Floating QR (Depan/Belakang Teks, Anti Halaman Kosong, Show/Hide Kanvas)', () => {
  const sppdDoc = buildSuratTugasSppdDocModel(
    {
      nomorSurat: '800/006/ST.SMKITA/XI/2026',
      tujuanLokasi: 'DRAGON MOTION PICTURE ASIA'
    },
    DEFAULT_SCHOOL_PROFILE,
    { paperSize: 'f4' }
  );

  // 1. Default placement Surat Tugas menempatkan QR pada halaman 0 (halaman Surat Tugas)
  const pref = findPreferredSignatureQrPlacement(sppdDoc, 3);
  assert.equal(pref.pageIndex, 0);

  // 2. Uji layer 'behind' (Di Belakang Teks) & koordinat kustom + hide di kanvas
  const customPlacement = normalizeQrPlacement({
    pageIndex: 0,
    xPct: 62.5,
    yPct: 71.2,
    sizePx: 96,
    layer: 'behind',
    visibleOnCanvas: false,
    showCaption: true,
    customPositioned: true
  }, 3);
  assert.equal(customPlacement.layer, 'behind');
  assert.equal(customPlacement.visibleOnCanvas, false);

  const behindHtml = buildFloatingQrElementHtml({
    qrDataUrl: 'data:image/png;base64,AA==',
    placement: customPlacement,
    interactive: true,
    pageCount: 3
  });
  assert.ok(behindHtml.includes('layer-behind'));
  assert.ok(behindHtml.includes('left: 62.5%; top: 71.2%; width: 96px;'));
  assert.ok(!behindHtml.includes('floating-qr-toolbar'), 'Kanvas barcode harus bersih tanpa toolbar menempel');

  const ctxHtml = buildFloatingQrContextMenuHtml({
    placement: customPlacement,
    pageCount: 3
  });
  assert.ok(ctxHtml.includes('Di Belakang Teks'));
  assert.ok(ctxHtml.includes('data-qr-action="hide-canvas"'));
  assert.ok(ctxHtml.includes('data-qr-action="move-page"'));

  // 3. Printable dimensions mencegah halaman ekstra kosong akibat overlay
  const f4Dims = computePrintableDimsMm({
    paperSize: 'f4',
    orientation: 'portrait',
    margins: { top: 2, bottom: 2, left: 2.5, right: 2 }
  });
  assert.equal(f4Dims.sheetHeightMm, 330);
  assert.equal(f4Dims.printableHeightMm, 288);

  // 4. Pratinjau Surat Keluar wajib mempertahankan ke-3 halaman, tabel SPPD, & floating QR meski visibleOnCanvas === false
  sppdDoc.qrPlacement = customPlacement;
  const savedOut = store.addOutgoingLetter({
    sequenceNumber: 6,
    year: 2026,
    typeCode: 'ST',
    nomorSurat: '800/006/ST.SMKITA/XI/2026',
    title: 'SURAT TUGAS',
    recipientName: 'Guru PKL',
    isVerifiedSah: true,
    verificationStatus: 'sah',
    qrDataUrl: 'data:image/png;base64,AA==',
    document: sppdDoc,
    skipAutoAdvanceNumber: true
  });

  const previewHtml = renderReadOnlyDocumentHtml(savedOut, { signatureMode: 'auto', isForPrint: false });
  assert.equal((previewHtml.match(/class="smart-paper-sheet /g) || []).length, 3, 'Pratinjau Surat Keluar wajib menampilkan tepat 3 halaman kertas 1:1');
  assert.ok(previewHtml.includes('smart-word-table'), 'Tabel pegawai dan keterangan wajib tampil di pratinjau Surat Keluar');
  assert.ok(previewHtml.includes('smart-table-sppd'), 'Tabel SPPD wajib tampil utuh di pratinjau Surat Keluar');
  assert.ok(previewHtml.includes('data-floating-qr="true"'), 'Barcode QR wajib tetap muncul di Surat Keluar walau disembunyikan di kanvas');
  assert.ok(previewHtml.includes('layer-behind'), 'Posisi lapisan Belakang Teks terbawa ke Surat Keluar');

  const printManualHtml = renderReadOnlyDocumentHtml(savedOut, { signatureMode: 'manual', isForPrint: true });
  assert.ok(!printManualHtml.includes('data-floating-qr="true"'), 'Mode cetak TTD manual tidak boleh merender Barcode QR');

  // Persistensi setActiveQrPlacement sinkron ke arsip outgoing terkait
  store.setState({ smartDocument: { pages: [{ id: 'p1', html: '<p>tes</p>' }], verification: { outgoingId: savedOut.id } } }, false);
  store.setActiveQrPlacement({ ...customPlacement, layer: 'front', xPct: 70 }, 'smart');
  const updatedOut = store.getOutgoingLetters().find(x => x.id === savedOut.id);
  assert.equal(updatedOut.qrPlacement.layer, 'front');
  assert.equal(updatedOut.qrPlacement.xPct, 70);

  store.removeOutgoingLetter(savedOut.id);
});

function createMiniDomEnv() {
  const createEvent = (type, props = {}) => {
    const ev = {
      type,
      bubbles: props.bubbles !== false,
      cancelable: props.cancelable !== false,
      defaultPrevented: false,
      cancelBubble: false,
      target: null,
      currentTarget: null,
      button: props.button ?? 0,
      buttons: props.buttons ?? 1,
      clientX: props.clientX ?? 0,
      clientY: props.clientY ?? 0,
      key: props.key ?? '',
      preventDefault() { ev.defaultPrevented = true; },
      stopPropagation() { ev.cancelBubble = true; },
      stopImmediatePropagation() { ev.cancelBubble = true; }
    };
    return ev;
  };

  class MiniElement {
    constructor(tagName = 'div') {
      this.tagName = tagName.toUpperCase();
      this.parentNode = null;
      this.children = [];
      this.attributes = {};
      this.dataset = {};
      this.style = {};
      this._listeners = {};
      this._rect = { left: 0, top: 0, right: 600, bottom: 900, width: 600, height: 900 };
      this.offsetWidth = 600;
      const self = this;
      this.classList = {
        _set: new Set(),
        add(...cls) { cls.forEach(c => c && this._set.add(c)); self._syncClassAttr(); },
        remove(...cls) { cls.forEach(c => this._set.delete(c)); self._syncClassAttr(); },
        contains(c) { return this._set.has(c); },
        toggle(c, force) {
          if (force === true) { this._set.add(c); }
          else if (force === false) { this._set.delete(c); }
          else if (this._set.has(c)) { this._set.delete(c); }
          else { this._set.add(c); }
          self._syncClassAttr();
        }
      };
    }
    _syncClassAttr() {
      this.attributes.class = [...this.classList._set].join(' ');
    }
    set className(val) {
      this.classList._set = new Set(String(val || '').split(/\s+/).filter(Boolean));
      this._syncClassAttr();
    }
    get className() {
      return [...this.classList._set].join(' ');
    }
    setAttribute(k, v) {
      const val = String(v ?? '');
      this.attributes[k] = val;
      if (k === 'class') this.className = val;
      if (k.startsWith('data-')) {
        const camel = k.slice(5).replace(/-([a-z])/g, (_, ch) => ch.toUpperCase());
        this.dataset[camel] = val;
      }
    }
    getAttribute(k) {
      if (k === 'class') return this.className;
      return this.attributes[k] ?? null;
    }
    appendChild(child) {
      if (child.parentNode) {
        const idx = child.parentNode.children.indexOf(child);
        if (idx >= 0) child.parentNode.children.splice(idx, 1);
      }
      child.parentNode = this;
      this.children.push(child);
      return child;
    }
    remove() {
      if (this.parentNode) {
        const idx = this.parentNode.children.indexOf(this);
        if (idx >= 0) this.parentNode.children.splice(idx, 1);
        this.parentNode = null;
      }
    }
    set innerHTML(htmlStr) {
      this.children = [];
      const root = parseMiniHtml(String(htmlStr || ''));
      [...root.children].forEach(c => this.appendChild(c));
    }
    get innerHTML() {
      return this.children.map(c => c.outerHTML || '').join('');
    }
    matchesSelectorPart(part) {
      const s = part.trim();
      if (!s) return false;
      const tagMatch = s.match(/^[a-zA-Z0-9_-]+/);
      let rest = s;
      if (tagMatch) {
        if (this.tagName !== tagMatch[0].toUpperCase()) return false;
        rest = rest.slice(tagMatch[0].length);
      }
      while (rest.length > 0) {
        if (rest.startsWith('.')) {
          const m = rest.match(/^\.([a-zA-Z0-9_-]+)/);
          if (!m || !this.classList.contains(m[1])) return false;
          rest = rest.slice(m[0].length);
        } else if (rest.startsWith('[')) {
          const m = rest.match(/^\[([a-zA-Z0-9_-]+)(?:="([^"]*)")?\]/);
          if (!m) return false;
          const attrVal = this.getAttribute(m[1]);
          if (attrVal === null) return false;
          if (m[2] !== undefined && attrVal !== m[2]) return false;
          rest = rest.slice(m[0].length);
        } else {
          return false;
        }
      }
      return true;
    }
    matches(selector) {
      return String(selector || '').split(',').some((commaSel) => {
        const parts = commaSel.trim().split(/\s+/).filter(Boolean);
        if (parts.length === 0) return false;
        if (!this.matchesSelectorPart(parts[parts.length - 1])) return false;
        let cur = this.parentNode;
        for (let i = parts.length - 2; i >= 0; i--) {
          while (cur && !cur.matchesSelectorPart(parts[i])) cur = cur.parentNode;
          if (!cur) return false;
          cur = cur.parentNode;
        }
        return true;
      });
    }
    closest(selector) {
      let cur = this;
      while (cur) {
        if (cur.matches && cur.matches(selector)) return cur;
        cur = cur.parentNode;
      }
      return null;
    }
    contains(other) {
      let cur = other;
      while (cur) {
        if (cur === this) return true;
        cur = cur.parentNode;
      }
      return false;
    }
    querySelectorAll(selector) {
      const out = [];
      const walk = (node) => {
        for (const ch of node.children) {
          if (ch.matches(selector)) out.push(ch);
          walk(ch);
        }
      };
      walk(this);
      return out;
    }
    querySelector(selector) {
      return this.querySelectorAll(selector)[0] || null;
    }
    addEventListener(type, fn, opts = false) {
      const capture = typeof opts === 'boolean' ? opts : Boolean(opts?.capture);
      if (!this._listeners[type]) this._listeners[type] = [];
      this._listeners[type].push({ fn, capture });
    }
    removeEventListener(type, fn, opts = false) {
      const capture = typeof opts === 'boolean' ? opts : Boolean(opts?.capture);
      if (!this._listeners[type]) return;
      this._listeners[type] = this._listeners[type].filter(l => l.fn !== fn || l.capture !== capture);
    }
    dispatchEvent(ev) {
      if (!ev.target) ev.target = this;
      const chain = [];
      let cur = this;
      while (cur) { chain.push(cur); cur = cur.parentNode; }
      const fullPath = [win, ...chain.reverse()];
      for (const node of fullPath) {
        if (ev.cancelBubble) break;
        const list = (node._listeners[ev.type] || []).filter(l => l.capture);
        for (const l of list) {
          ev.currentTarget = node;
          l.fn.call(node, ev);
          if (ev.cancelBubble) break;
        }
      }
      if (!ev.cancelBubble && ev.bubbles) {
        for (let i = fullPath.length - 1; i >= 0; i--) {
          if (ev.cancelBubble) break;
          const node = fullPath[i];
          const list = (node._listeners[ev.type] || []).filter(l => !l.capture);
          for (const l of list) {
            ev.currentTarget = node;
            l.fn.call(node, ev);
            if (ev.cancelBubble) break;
          }
        }
      }
      return !ev.defaultPrevented;
    }
    getBoundingClientRect() {
      return { ...this._rect };
    }
  }

  function parseMiniHtml(html) {
    const root = new MiniElement('div');
    const stack = [root];
    const tokenRe = /<\/([a-zA-Z0-9]+)\s*>|<([a-zA-Z0-9]+)([^>]*?)(\/?)>/g;
    let m;
    while ((m = tokenRe.exec(html))) {
      if (m[1]) {
        if (stack.length > 1) stack.pop();
      } else {
        const tag = m[2].toLowerCase();
        const attrsStr = m[3] || '';
        const selfClose = Boolean(m[4]) || tag === 'img';
        const el = new MiniElement(tag);
        const attrRe = /([a-zA-Z0-9_-]+)="([^"]*)"/g;
        let am;
        while ((am = attrRe.exec(attrsStr))) {
          el.setAttribute(am[1], am[2]);
        }
        stack[stack.length - 1].appendChild(el);
        if (!selfClose) stack.push(el);
      }
    }
    return root;
  }

  const docBody = new MiniElement('body');
  const docEl = new MiniElement('document');
  docEl.body = docBody;
  docEl.createElement = (tag) => new MiniElement(tag);
  docEl.createRange = () => ({ selectNodeContents() {} });

  const win = new MiniElement('window');
  win.appendChild(docEl);
  docEl.appendChild(docBody);
  win.getSelection = () => ({ removeAllRanges() {}, addRange() {} });

  return { win, docEl, docBody, MiniElement, createEvent };
}

test('Interaksi Floating QR (Drag Bersih, Resize, Menu Klik Kanan MS Word, Lintas Halaman, Sembunyi) & Isolasi Border Tabel', () => {
  const prevWin = globalThis.window;
  const prevDoc = globalThis.document;
  const prevGCS = globalThis.getComputedStyle;
  const env = createMiniDomEnv();
  globalThis.window = env.win;
  globalThis.document = env.docEl;
  globalThis.getComputedStyle = () => ({ position: 'relative' });

  try {
    // 1. Pastikan CSS bersih tanpa toolbar menempel permanen & tidak memerangkap visual layer
    const cssText = fs.readFileSync(path.resolve('css/app.css'), 'utf8');
    const boxBehindRule = cssText.match(/\.floating-qr-box\.layer-behind\s*\{([^}]*)\}/)?.[1] || '';
    assert.ok(!/z-index\s*:/.test(boxBehindRule), '.floating-qr-box.layer-behind tidak boleh memiliki z-index agar layer tetap rapi');
    assert.match(cssText, /\.floating-qr-box\.layer-behind\s+\.floating-qr-visual\s*\{[^}]*z-index:\s*2;/);
    assert.match(cssText, /\.floating-qr-box\.layer-front\s+\.floating-qr-visual\s*\{[^}]*z-index:\s*70;/);
    assert.match(cssText, /\.fqr-context-menu\s*\{[^}]*position:\s*fixed;[^}]*z-index:\s*99999;/);
    assert.match(cssText, /\.floating-qr-resize\s*\{[^}]*z-index:\s*90;/);

    // 2. Bangun kanvas desk dengan tabel SPPD di bawah overlay QR
    const desk = new env.MiniElement('div');
    env.docBody.appendChild(desk);
    const page1 = new env.MiniElement('div');
    page1.className = 'smart-paper-sheet';
    const page2 = new env.MiniElement('div');
    page2.className = 'smart-paper-sheet';
    const tbl = new env.MiniElement('table');
    tbl._rect = { left: 40, top: 80, right: 560, bottom: 520, width: 520, height: 440 };
    const tr1 = new env.MiniElement('tr');
    const tr2 = new env.MiniElement('tr');
    const td1 = new env.MiniElement('td');
    td1._rect = { left: 40, top: 80, right: 300, bottom: 200, width: 260, height: 120 };
    const td2 = new env.MiniElement('td');
    td2._rect = { left: 300, top: 80, right: 560, bottom: 200, width: 260, height: 120 };
    const td3 = new env.MiniElement('td');
    td3.setAttribute('colspan', '2');
    td3._rect = { left: 40, top: 200, right: 560, bottom: 520, width: 520, height: 320 };
    tr1.appendChild(td1);
    tr1.appendChild(td2);
    tr2.appendChild(td3);
    tbl.appendChild(tr1);
    tbl.appendChild(tr2);
    page1.appendChild(tbl);

    const stage1 = new env.MiniElement('div');
    stage1.className = 'floating-qr-stage';
    stage1._rect = { left: 50, top: 40, right: 550, bottom: 740, width: 500, height: 700 };
    stage1.offsetWidth = 625; // Simulasi zoom 80% (500 / 625 = 0.8)
    page1.appendChild(stage1);

    const stage2 = new env.MiniElement('div');
    stage2.className = 'floating-qr-stage';
    stage2._rect = { left: 50, top: 800, right: 550, bottom: 1500, width: 500, height: 700 };
    stage2.offsetWidth = 625;
    page2.appendChild(stage2);

    desk.appendChild(page1);
    desk.appendChild(page2);

    const tableCtrl = attachInteractiveTables({
      deskEl: desk,
      tableSelector: 'table',
      onMutate: () => { throw new Error('Tabel tidak boleh ikut terseret saat klik/drag QR'); }
    });

    let latestPlacement = null;
    mountInteractiveFloatingQr({
      stages: [stage1, stage2],
      verification: { isVerifiedSah: true, qrDataUrl: 'data:image/png;base64,AA==' },
      rawPlacement: { pageIndex: 0, xPct: 50, yPct: 30, sizePx: 86, layer: 'front', visibleOnCanvas: true, showCaption: true, customPositioned: true },
      onPlacementChange: (p) => { latestPlacement = p; }
    });

    const getBox = (st) => st.querySelector('[data-floating-qr="true"]');
    const getCtxMenu = () => env.docBody.querySelector('.fqr-context-menu');

    let box = getBox(stage1);
    assert.ok(box, 'Floating QR harus terpasang pada stage halaman 1');
    assert.equal(box.querySelector('.floating-qr-toolbar'), null, 'Tidak boleh ada bar tombol muncul permanen bersama barcode');
    box._rect = { left: 300, top: 200, right: 386, bottom: 300, width: 86, height: 100 };

    // A. Hover dulu ke batas baris tabel (x = 350, y = 200 -> armedHit aktif), lalu pindah ke QR di titik x = 350, y = 200:
    //    tableInteraction wajib melepas armedHit, membersihkan cursor resize, dan tidak membajak drag QR!
    const visualEl = box.querySelector('.floating-qr-visual');
    tbl.dispatchEvent(env.createEvent('mousemove', { clientX: 350, clientY: 200 }));
    assert.equal(env.docBody.classList.contains('cursor-row-resize'), true, 'Pra-syarat: batas tabel asli memicu cursor-row-resize');

    visualEl.dispatchEvent(env.createEvent('mousemove', { clientX: 350, clientY: 200 }));
    assert.equal(env.docBody.classList.contains('cursor-row-resize'), false, 'Hover QR di atas garis tabel wajib membersihkan row-resize');
    assert.equal(env.docBody.classList.contains('cursor-col-resize'), false, 'Hover QR di atas garis kolom tabel wajib membersihkan col-resize');

    const mdVisualEv = env.createEvent('mousedown', { clientX: 350, clientY: 200 });
    visualEl.dispatchEvent(mdVisualEv);
    assert.equal(env.docBody.classList.contains('word-table-dragging'), false, 'Klik QR di atas garis tabel tidak boleh membajak drag tabel');
    env.win.dispatchEvent(env.createEvent('mouseup', { clientX: 350, clientY: 200 }));

    // B. Klik kanan (contextmenu) pada QR membuka menu konteks profesional di document.body
    assert.equal(getCtxMenu(), null, 'Menu konteks belum boleh muncul sebelum klik kanan');
    box.dispatchEvent(env.createEvent('contextmenu', { clientX: 340, clientY: 240 }));
    let ctxMenu = getCtxMenu();
    assert.ok(ctxMenu, 'Klik kanan pada barcode QR wajib membuka menu konteks .fqr-context-menu');

    // Switch ke Belakang Teks lewat menu konteks
    ctxMenu.querySelector('[data-qr-action="set-layer"][data-layer="behind"]').dispatchEvent(env.createEvent('click'));
    assert.equal(latestPlacement.layer, 'behind');
    assert.equal(getCtxMenu(), null, 'Menu konteks otomatis menutup setelah memilih aksi');

    // C. Saat mode Belakang Teks (pointer-events: none), klik kanan di koordinat di atas QR (di sel td3 di balik QR) tetap membuka menu QR!
    box = getBox(stage1);
    box._rect = { left: 300, top: 200, right: 386, bottom: 300, width: 86, height: 100 };
    td3.dispatchEvent(env.createEvent('contextmenu', { clientX: 340, clientY: 240 }));
    ctxMenu = getCtxMenu();
    assert.ok(ctxMenu, 'Klik kanan di area barcode meski dalam mode Belakang Teks wajib tetap membuka menu konteks');

    // Ubah ke mode Depan Teks lagi
    ctxMenu.querySelector('[data-qr-action="set-layer"][data-layer="front"]').dispatchEvent(env.createEvent('click'));
    assert.equal(latestPlacement.layer, 'front');

    // D. Klik kanan -> Toggle Caption, Pindah ke H2, & Ubah Ukuran Preset Besar (112px)
    box = getBox(stage1);
    box.dispatchEvent(env.createEvent('contextmenu', { clientX: 340, clientY: 240 }));
    ctxMenu = getCtxMenu();
    ctxMenu.querySelector('[data-qr-action="set-size"][data-size="112"]').dispatchEvent(env.createEvent('click'));
    assert.equal(latestPlacement.sizePx, 112);

    ctxMenu.querySelector('[data-qr-action="toggle-caption"]').dispatchEvent(env.createEvent('click'));
    assert.equal(latestPlacement.showCaption, false);

    box = getBox(stage1);
    box.dispatchEvent(env.createEvent('contextmenu', { clientX: 340, clientY: 240 }));
    ctxMenu = getCtxMenu();
    ctxMenu.querySelector('[data-qr-action="move-page"][data-page="1"]').dispatchEvent(env.createEvent('click'));
    assert.equal(latestPlacement.pageIndex, 1);
    assert.ok(getBox(stage2), 'QR harus berpindah ke stage halaman 2 setelah memilih H2');
    assert.equal(getBox(stage1), null);

    // E. Drag langsung pada visual barcode QR (pointerdown -> pointermove -> pointerup) antar halaman
    box = getBox(stage2);
    box._rect = { left: 200, top: 900, right: 312, bottom: 1012, width: 112, height: 112 };
    const qrDragZone = box.querySelector('[data-qr-drag="true"]');
    qrDragZone.dispatchEvent(env.createEvent('pointerdown', { clientX: 240, clientY: 940 }));
    env.win.dispatchEvent(env.createEvent('pointermove', { clientX: 160, clientY: 120 }));
    env.win.dispatchEvent(env.createEvent('pointerup', { clientX: 160, clientY: 120 }));
    assert.equal(latestPlacement.pageIndex, 0, 'Drag ke area atas harus memindahkan QR kembali ke halaman 1');
    assert.ok(latestPlacement.yPct < 20, 'Koordinat Y hasil drag harus diperbarui');

    // F. Resize sudut memperhitungkan zoom dinamis (0.8 => dx 40px layar = +50px ukuran QR: 112 + 50 = 162px)
    box = getBox(stage1);
    const resizeBtn = box.querySelector('[data-qr-resize="true"]');
    resizeBtn.dispatchEvent(env.createEvent('mousedown', { clientX: 300, clientY: 300 }));
    env.win.dispatchEvent(env.createEvent('mousemove', { clientX: 340, clientY: 340 }));
    env.win.dispatchEvent(env.createEvent('mouseup', { clientX: 340, clientY: 340 }));
    assert.equal(latestPlacement.sizePx, 162, 'Resize pada skala zoom 80% harus akurat diskalakan');

    // G. Klik kanan -> Sembunyikan di Kanvas
    box = getBox(stage1);
    box.dispatchEvent(env.createEvent('contextmenu', { clientX: 200, clientY: 150 }));
    ctxMenu = getCtxMenu();
    ctxMenu.querySelector('[data-qr-action="hide-canvas"]').dispatchEvent(env.createEvent('click'));
    assert.equal(latestPlacement.visibleOnCanvas, false);
    assert.equal(getBox(stage1), null);

    tableCtrl.destroy();
  } finally {
    globalThis.window = prevWin;
    globalThis.document = prevDoc;
    globalThis.getComputedStyle = prevGCS;
  }
});

test('Pratinjau Surat Keluar fallback otomatis saat smartPagesSnapshot kosong/rusak dan mendukung fallback uploadedDocx', () => {
  const sppdDoc = buildSuratTugasSppdDocModel(
    {
      nomorSurat: '800/014/ST.SMKITA/XI/2026',
      tujuanLokasi: 'Golden Tulip Essential Tangerang'
    },
    DEFAULT_SCHOOL_PROFILE,
    { paperSize: 'f4' }
  );

  // 1. Jika smartPagesSnapshot punya halaman kosong namun documentSnapshot valid -> wajib fallback ke documentSnapshot
  const emptySnapItem = {
    id: 'out-empty-snap-reg',
    sequenceNumber: 14,
    typeCode: 'ST',
    nomorSurat: '800/014/ST.SMKITA/XI/2026',
    title: 'SURAT TUGAS',
    recipientName: 'Umum',
    tanggalSurat: '23 November 2026',
    isVerifiedSah: true,
    verificationStatus: 'sah',
    qrDataUrl: 'data:image/png;base64,AA==',
    documentSnapshot: sppdDoc,
    smartPagesSnapshot: {
      pages: [{ id: 'p1', html: '' }, { id: 'p2', html: '<p><br></p>' }],
      letterOptions: { paperSize: 'f4', orientation: 'portrait', showKop: true }
    }
  };
  const recoveredHtml = renderReadOnlyDocumentHtml(emptySnapItem, { signatureMode: 'auto', isForPrint: false });
  assert.ok(recoveredHtml.includes('SURAT TUGAS'), 'Pratinjau tidak boleh putih kosong ketika smartPagesSnapshot kosong');
  assert.ok(recoveredHtml.includes('Golden Tulip Essential Tangerang'), 'Isi documentSnapshot harus dirender saat snapshot halaman kosong');

  // 2. Jika hanya ada uploadedDocx.paragraphs (tanpa smartPagesSnapshot & tanpa blocks) -> tetap tampil teks surat
  const docxOnlyItem = {
    id: 'out-docx-only-reg',
    sequenceNumber: 15,
    typeCode: 'KET',
    nomorSurat: '421.3/015/KET.SMPITA/XI/2026',
    title: 'SURAT KETERANGAN',
    recipientName: 'Siswa SMP',
    tanggalSurat: '23 November 2026',
    uploadedDocx: {
      paragraphs: [
        { text: 'SURAT KETERANGAN AKTIF' },
        { text: 'Nomor: 421.3/015/KET.SMPITA/XI/2026' },
        { text: 'Kepala SMP IT Al-Hawari menerangkan bahwa siswa terkait aktif.' }
      ]
    }
  };
  const docxFallbackHtml = renderReadOnlyDocumentHtml(docxOnlyItem, { signatureMode: 'manual', isForPrint: false });
  assert.ok(docxFallbackHtml.includes('SURAT KETERANGAN AKTIF'));
  assert.ok(docxFallbackHtml.includes('Kepala SMP IT Al-Hawari menerangkan'));
});

test('Deteksi Otomatis Area Tanda Tangan QR menolak tabel Kendaraan/SPPD Hal 3 & tabel Identitas, namun menghormati customPositioned', () => {
  // 1. Model bebas 3 halaman: Hal 0 tabel identitas (Nama: Yangyang Budiman / Jabatan: Kepala Sekolah),
  //    Hal 1 blok TTD bebas Kepala Sekolah + Yangyang Budiman, Hal 2 tabel Keterangan I/II (Mobil / Umum)
  const customThreePageDoc = {
    blocks: [
      {
        type: 'paragraph',
        runs: [{ text: 'Yang bertanda tangan di bawah ini, Kepala SMP IT Al-Hawari menerangkan bahwa:' }]
      },
      {
        type: 'table',
        rows: [
          [{ text: 'Nama' }, { text: ': Yangyang Budiman, S.Pd.' }],
          [{ text: 'Jabatan' }, { text: ': Kepala Sekolah' }]
        ]
      },
      { type: 'pageBreak' },
      { type: 'paragraph', text: 'Demikian surat tugas ini dibuat untuk dilaksanakan.' },
      { type: 'paragraph', text: 'Kepala SMP IT Al-Hawari', alignment: 'right' },
      { type: 'paragraph', text: 'Yangyang Budiman, S.Pd.', alignment: 'right' },
      { type: 'pageBreak' },
      {
        type: 'table',
        rows: [
          [{ text: 'Tempat kedudukan pegawai' }, { text: 'Dengan Kendaraan' }, { text: 'Tanda Tangan' }],
          [{ text: 'SMKS IT Al-Hawari' }, { text: 'Mobil / Umum' }, { text: '' }]
        ]
      }
    ]
  };

  const modelPlacement = findPreferredSignatureQrPlacement(
    customThreePageDoc,
    3,
    { principalName: 'Yangyang Budiman, S.Pd.' }
  );
  assert.equal(
    modelPlacement.pageIndex,
    1,
    'Deteksi otomatis wajib memilih Halaman 2 (indeks 1) tempat TTD Kepala Sekolah, bukan Hal 1 identitas atau Hal 3 Mobil/Umum'
  );

  // 2. Uji DOM calibratePlacementFromDom ketika rawPlacement awal salah di halaman 3 (index 2, Mobil / Umum) dengan customPositioned: false
  const env = createMiniDomEnv();
  const makeMiniStage = (pageIdx, htmlStr) => {
    const sheet = new env.MiniElement('div');
    sheet.className = 'smart-paper-sheet';
    const page = new env.MiniElement('div');
    page.className = 'smart-editable-page';
    page.innerHTML = htmlStr;
    sheet.appendChild(page);

    const stage = new env.MiniElement('div');
    stage.className = 'floating-qr-stage';
    stage.dataset.qrPage = String(pageIdx);
    stage._rect = { left: 40, top: 40 + pageIdx * 900, right: 640, bottom: 840 + pageIdx * 900, width: 600, height: 800 };
    stage.offsetWidth = 600;
    sheet.appendChild(stage);
    env.docBody.appendChild(sheet);
    return stage;
  };

  const st0 = makeMiniStage(0, `
    <p></p>
    <table><tbody><tr><td></td><td></td></tr></tbody></table>
    <table><tbody><tr><td></td><td></td></tr></tbody></table>
  `);
  const p0Tables = st0.parentNode.querySelectorAll('table');
  const idTable = p0Tables[0];
  const idTds = idTable.querySelectorAll('td');
  idTds[0].textContent = 'Nama';
  idTds[1].textContent = ': Dian Nurhamdani, S.Pd.';
  idTds.forEach((td) => { td._rect = { left: 80, top: 200, right: 300, bottom: 230, width: 220, height: 30 }; td.parentElement = td.parentNode; });
  idTable.textContent = 'Nama : Dian Nurhamdani, S.Pd.';

  const sigTable = p0Tables[1];
  const p0Tds = sigTable.querySelectorAll('td');
  p0Tds[0].textContent = '';
  p0Tds[0]._rect = { left: 60, top: 440, right: 350, bottom: 560, width: 290, height: 120 };
  p0Tds[0].parentElement = p0Tds[0].parentNode;
  const sigTd = p0Tds[1];
  sigTd.textContent = 'Garut, 23 November 2026 Kepala SMKS IT Al-Hawari Dian Nurhamdani, S.Pd.';
  sigTd.childNodes = [
    { nodeType: 3, nodeValue: 'Garut, 23 November 2026', parentElement: { tagName: 'TD', getBoundingClientRect: () => ({ left: 380, top: 445, right: 590, bottom: 463, width: 210, height: 18 }) } },
    { nodeType: 3, nodeValue: 'Kepala SMKS IT Al-Hawari', parentElement: { tagName: 'TD', getBoundingClientRect: () => ({ left: 380, top: 465, right: 590, bottom: 483, width: 210, height: 18 }) } },
    { nodeType: 3, nodeValue: 'Dian Nurhamdani, S.Pd.', parentElement: { tagName: 'STRONG', getBoundingClientRect: () => ({ left: 380, top: 535, right: 590, bottom: 553, width: 210, height: 18 }) } }
  ];
  sigTd._rect = { left: 360, top: 440, right: 610, bottom: 560, width: 250, height: 120 };
  sigTd.parentElement = sigTd.parentNode;
  sigTable.textContent = sigTd.textContent;

  const st1 = makeMiniStage(1, '<p></p>');
  const st2 = makeMiniStage(2, '<table><tbody><tr><td></td><td></td><td></td></tr></tbody></table>');
  const p2Table = st2.parentNode.querySelector('table');
  const p2Tds = st2.parentNode.querySelectorAll('td');
  ['Dengan Kendaraan', 'Mobil / Umum', 'Tanda Tangan'].forEach((label, i) => {
    p2Tds[i].textContent = label;
    p2Tds[i].parentElement = p2Tds[i].parentNode;
    p2Tds[i]._rect = { left: 60 + i * 160, top: 2100, right: 200 + i * 160, bottom: 2300, width: 140, height: 200 };
  });
  p2Table.textContent = 'Dengan Kendaraan Mobil / Umum Tanda Tangan';

  const recalibrated = calibratePlacementFromDom(
    [st0, st1, st2],
    { pageIndex: 2, xPct: 68, yPct: 64, sizePx: 86, customPositioned: false },
    86,
    { school: { principalName: 'Dian Nurhamdani, S.Pd.' } }
  );
  assert.equal(recalibrated.pageIndex, 0, 'Kalibrasi otomatis wajib menarik QR dari halaman 3 (Mobil/Umum) ke sel TTD halaman 1');

  // Jika user sudah menggeser manual (customPositioned: true), posisi manual TIDAK boleh ditimpa
  const manualPinned = calibratePlacementFromDom(
    [st0, st1, st2],
    { pageIndex: 2, xPct: 55, yPct: 45, sizePx: 86, customPositioned: true },
    86,
    { school: { principalName: 'Dian Nurhamdani, S.Pd.' } }
  );
  assert.equal(manualPinned.pageIndex, 2);
  assert.equal(manualPinned.xPct, 55);
});
