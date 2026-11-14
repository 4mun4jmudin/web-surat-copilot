/**
 * tests/vaultAndPatcher.test.js - Unit Test Gudang Surat (Letter Vault), Acuan AI, & Non-Destructive Patcher
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { store } from '../js/state.js';
import { parseDocxForEditing, applyCanvasEditsToDocx } from '../js/docx/patch.js';
import { renderDocx } from '../js/docx/render.js';
import { samplePklDocument } from '../js/model/sample.js';

test('Gudang Surat: Tersedia arsip bawaan resmi dan mendukung operasi CRUD', () => {
  const initialVault = store.getLetterVault();
  assert.ok(initialVault.length >= 3, 'Harus tersedia minimal 3 arsip surat bawaan');

  const skEntry = initialVault.find(l => l.category === 'Surat Keterangan');
  assert.ok(skEntry, 'Surat Keterangan bawaan harus ada');
  assert.ok(skEntry.isAiReference, 'Surat bawaan aktif sebagai acuan AI secara default');

  // 1. Tambah arsip baru ke Gudang Surat
  const newEntry = store.addLetterVaultEntry({
    title: 'Surat Dispensasi Mengikuti LKS Wilayah',
    nomorSurat: '421.3 / 112 / SMKITA / XI / 2026',
    tanggalSurat: '14 November 2026',
    category: 'Surat Dispensasi',
    targetPerson: 'Rizky Ramadhan',
    perihal: 'Dispensasi Lomba Kompetensi Siswa',
    summary: 'Dispensasi siswa mengikuti LKS bidang Graphic Design Technology selama 4 hari.',
    isAiReference: true,
    source: 'created'
  });

  assert.ok(newEntry.id, 'Entri baru harus memiliki ID unik');
  const afterAdd = store.getLetterVault();
  assert.equal(afterAdd.length, initialVault.length + 1);

  // 2. Perbarui arsip
  store.updateLetterVaultEntry(newEntry.id, {
    perihal: 'Dispensasi Lomba Kompetensi Siswa (Tingkat Provinsi)'
  });
  const updated = store.getLetterVault().find(l => l.id === newEntry.id);
  assert.equal(updated.perihal, 'Dispensasi Lomba Kompetensi Siswa (Tingkat Provinsi)');

  // 3. Toggle status acuan AI (Knowledge Base)
  const isNowRef = store.toggleLetterAiReference(newEntry.id);
  assert.equal(isNowRef, false, 'Setelah toggle, harus menjadi false');
  assert.equal(store.getLetterVault().find(l => l.id === newEntry.id).isAiReference, false);

  // Toggle kembali menjadi true
  store.toggleLetterAiReference(newEntry.id);
  assert.equal(store.getLetterVault().find(l => l.id === newEntry.id).isAiReference, true);

  // 4. Pengambilan acuan AI (Knowledge Base Few-Shot)
  const aiRefs = store.getAiReferences(5);
  assert.ok(aiRefs.length >= 1, 'Harus ada minimal 1 acuan AI aktif');
  assert.ok(aiRefs.every(r => r.isAiReference === true), 'Semua referensi harus berstatus isAiReference: true');

  // 5. Muat surat ke Lembar Kerja Canvas
  const loadSuccess = store.loadLetterToCanvas(newEntry.id);
  assert.equal(loadSuccess, true, 'Harus berhasil dimuat ke kanvas');

  // 6. Hapus arsip
  store.removeLetterVaultEntry(newEntry.id);
  const afterDelete = store.getLetterVault();
  assert.equal(afterDelete.length, initialVault.length);
  assert.equal(afterDelete.some(l => l.id === newEntry.id), false);
});

test('Preservasi Format Berkas Impor: Patcher memodifikasi teks tanpa merusak sectPr atau zip asli', async () => {
  // Buat berkas docx valid dari samplePklDocument
  const originalBytes = await renderDocx(samplePklDocument, 'uint8array');
  assert.ok(originalBytes.length > 0, 'Harus berhasil merender berkas docx awal');

  // 1. Parse berkas untuk pengeditan
  const parsed = await parseDocxForEditing(originalBytes);
  assert.ok(parsed.editableParagraphs.length > 0, 'Harus mengekstrak paragraf non-kosong');
  assert.ok(parsed.totalParagraphs >= parsed.editableParagraphs.length);

  // Ambil paragraf pertama yang akan diedit
  const targetPara = parsed.editableParagraphs[0];
  const oldText = targetPara.originalText;
  const newText = 'PENDAHULUAN DAN LATAR BELAKANG TERBARU HASIL EDITAN KANVAS';

  // 2. Terapkan perubahan teks menggunakan applyCanvasEditsToDocx
  const edits = [
    {
      index: targetPara.index,
      originalText: oldText,
      currentText: newText
    }
  ];

  const patchedResult = await applyCanvasEditsToDocx(originalBytes, edits, { 
    outputType: 'uint8array',
    returnFullResult: true 
  });
  assert.ok(patchedResult.patchedData, 'Harus menghasilkan data byte docx ter-patch');
  assert.equal(patchedResult.report.safetyPassed, true, 'Safety audit harus lolos (sectPr tidak rusak)');
  assert.equal(patchedResult.report.executedChanges.length, 1, 'Harus mencatat 1 perubahan replace');
  assert.ok(patchedResult.patchedXml.includes(newText), 'XML yang di-patch harus mengandung teks baru');

  // 3. Simpan dan bersihkan status dokumen impor di state
  store.setActiveUploadedDocx({
    rawBytes: originalBytes,
    fileName: 'surat_lama_sekolah.docx',
    paragraphs: edits
  });
  assert.equal(store.getState().activeUploadedDocx?.fileName, 'surat_lama_sekolah.docx');

  store.clearActiveUploadedDocx();
  assert.equal(store.getState().activeUploadedDocx, null);
});
