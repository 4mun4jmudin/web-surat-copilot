/**
 * tests/vaultAndPatcher.test.js - Unit Test Gudang Surat (Letter Vault), Acuan AI, & Non-Destructive Patcher
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { store } from '../js/state.js';
import { parseDocxForEditing, applyCanvasEditsToDocx } from '../js/docx/patch.js';
import { renderDocx } from '../js/docx/render.js';
import { samplePklDocument } from '../js/model/sample.js';
import { buildSuratTugasSppdDocModel } from '../js/letters/templates/suratTugasSppd.js';
import {
  extractDocumentBlueprintForAi,
  selectRelevantAiReferences,
  summarizeKnowledgeBaseStats,
  describeAiMatchFeedback
} from '../js/ai/knowledgeBase.js';
import { buildCopilotUserPrompt } from '../js/ai/prompts.js';
import { convertAiBlocksToDocumentBlocks } from '../js/ai/client.js';

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

test('AI Knowledge Base: Ekstraksi cetak biru lengkap (DocumentModel & SPPD & .docx) serta prioritas Master Terkunci ⭐', () => {
  const sppdDoc = buildSuratTugasSppdDocModel(
    {
      nomorSurat: '065/ST/SMKITA/12.01/IX/2025',
      tempatPKL: 'DRAGON MOTION PICTURE ASIA'
    },
    store.getSchoolProfile()
  );

  // 1. Ekstraksi dari DocumentModel kompleks multi-halaman (SPPD)
  const sppdBp = extractDocumentBlueprintForAi({
    id: 'test-sppd',
    title: 'Surat Tugas Penjemputan PKL + SPPD',
    category: 'Surat Tugas',
    perihal: 'Penjemputan PKL',
    document: sppdDoc
  });
  assert.ok(sppdBp, 'Blueprint SPPD harus terekstrak');
  assert.equal(sppdBp.sourceType, 'documentModel');
  assert.equal(sppdBp.hideMetaRecipient, true);
  assert.ok(sppdBp.blocksStructure.some(b => b.type === 'pageBreak'), 'Harus menangkap blok pageBreak');
  const sppdTable = sppdBp.blocksStructure.find(b => b.type === 'table' && b.borderStyle === 'sppd');
  assert.ok(sppdTable, 'Harus menangkap blok table SPPD beserta borderStyle');
  assert.ok(Array.isArray(sppdTable.colWidthsPct) && sppdTable.colWidthsPct.length === 2);

  // 2. Ekstraksi dari uploadedDocx.structured + edited paragraphs
  const docxBp = extractDocumentBlueprintForAi({
    id: 'test-docx',
    title: 'Surat Undangan Rapat Komite',
    category: 'Surat Undangan',
    uploadedDocx: {
      paragraphs: [{ index: 5, currentText: 'Kalimat pembuka terbaru yang telah diedit user.' }],
      structured: {
        title: { text: 'SURAT UNDANGAN', index: 1 },
        nomor: { text: 'Nomor: 012/UND/SMKITA/XI/2026', index: 2 },
        metaRows: [{ label: 'Perihal', value: 'Rapat Komite Sekolah' }],
        recipientLines: ['Kepada', 'Yth. Orang Tua Siswa'],
        bodyParagraphs: [{ index: 5, text: 'Kalimat pembuka lama.' }],
        identityRows: [{ label: 'Hari', value: ': Senin' }],
        signature: { city: 'Garut', title: 'Kepala Sekolah', signer: 'Dian Nurhadian, S.Pd.' }
      }
    }
  });
  assert.equal(docxBp.sourceType, 'uploadedDocx');
  assert.equal(docxBp.blocksStructure[0].text, 'Kalimat pembuka terbaru yang telah diedit user.');

  // 3. Prioritas Master Terkunci ⭐ ketika relevan vs Mode Fresh bila tidak ada yang cocok
  const miniVault = [
    {
      id: 'doc-biasa',
      title: 'Surat Tugas Upacara',
      category: 'Surat Tugas',
      perihal: 'Tugas Upacara Bendera',
      summary: 'Surat tugas guru pendamping upacara.',
      isAiReference: true,
      isMasterTemplate: false,
      isLockedTemplate: false,
      document: sppdDoc
    },
    {
      id: 'doc-master',
      title: 'Surat Tugas Penjemputan PKL + SPPD Master',
      category: 'Surat Tugas',
      perihal: 'Penjemputan Peserta Didik PKL',
      summary: 'Template 3 halaman Surat Tugas + SPPD perjalanan dinas.',
      isAiReference: true,
      isMasterTemplate: true,
      isLockedTemplate: true,
      document: sppdDoc
    }
  ];

  const matched = selectRelevantAiReferences('Buatkan surat tugas dan SPPD guru ke tempat PKL', miniVault, { limit: 2 });
  assert.equal(matched.matchMode, 'matched');
  assert.equal(matched.topReference.id, 'doc-master', 'Master Terkunci ⭐ yang relevan harus selalu menjadi peringkat #1');
  assert.ok(matched.topReference.blueprint?.blocksStructure?.length > 0);
  assert.ok(describeAiMatchFeedback(matched).includes('⭐ Master Terkunci'));

  const fresh = selectRelevantAiReferences('Buatkan berita acara serah terima barang laboratorium inventaris', miniVault, { limit: 2 });
  assert.equal(fresh.matchMode, 'fresh', 'Prompt yang tidak ada contoh sejenis di Gudang harus masuk Mode Buat Baru (fresh)');
  assert.equal(fresh.references.length, 0);
  assert.ok(describeAiMatchFeedback(fresh).includes('dari nol'));

  const stats = summarizeKnowledgeBaseStats(miniVault);
  assert.deepEqual(stats, { totalVault: 2, totalActive: 2, masterCount: 1 });
});

test('Gudang Surat Lock & Prompt Builder: Master Terkunci terlindung dari hapus dan injeksi cetak biru ke prompt AI', () => {
  const created = store.addLetterVaultEntry({
    title: 'Template Panggilan Ortu Bagus',
    nomorSurat: '099/PGL/SMKITA/XI/2026',
    tanggalSurat: '18 November 2026',
    category: 'Surat Panggilan',
    targetPerson: 'Orang Tua Siswa',
    perihal: 'Panggilan Pembinaan Siswa',
    summary: 'Surat panggilan orang tua/wali resmi.',
    isAiReference: true
  });

  // Kunci sebagai Master Template ⭐
  const locked = store.toggleMasterTemplate(created.id);
  assert.equal(locked.isMasterTemplate, true);
  assert.equal(locked.isLockedTemplate, true);

  // Upaya hapus biasa wajib ditolak ketika terkunci
  const blockedDelete = store.removeLetterVaultEntry(created.id);
  assert.equal(blockedDelete, false, 'Dokumen ⭐ Master Terkunci tidak boleh terhapus sebelum dibuka kuncinya');
  assert.ok(store.getLetterVault().some(l => l.id === created.id));

  // Cek integrasi smart selection dari store -> buildCopilotUserPrompt
  const kbSelection = store.getSmartAiReferencesForPrompt('Buatkan surat panggilan orang tua siswa pembinaan', 2);
  assert.equal(kbSelection.matchMode, 'matched');
  const promptText = buildCopilotUserPrompt({
    userPrompt: 'Buatkan surat panggilan orang tua siswa pembinaan',
    references: kbSelection.references,
    matchMode: kbSelection.matchMode,
    totalKnowledgeCount: kbSelection.totalKnowledgeCount
  });
  assert.ok(promptText.includes('⭐ MASTER TEMPLATE TERKUNCI'), 'Prompt harus menandai Master Template prioritas');

  const freshPrompt = buildCopilotUserPrompt({
    userPrompt: 'Buatkan surat balasan kerja sama magang luar negeri baru',
    references: [],
    matchMode: 'fresh',
    totalKnowledgeCount: kbSelection.totalKnowledgeCount
  });
  assert.ok(freshPrompt.includes('BUAT BARU DARI NOL'), 'Prompt mode fresh harus memberi tahu AI membuat baru dari awal');

  // Cek helper konversi blok AI bersama
  const docBlocks = convertAiBlocksToDocumentBlocks([
    { type: 'paragraph', text: 'Pembuka surat', bold: true },
    { type: 'table', borderStyle: 'sppd', rows: [['1.', 'Pejabat']] },
    { type: 'pageBreak' }
  ]);
  assert.equal(docBlocks.length, 3);
  assert.equal(docBlocks[1].borderStyle, 'sppd');
  assert.equal(docBlocks[2].type, 'pageBreak');

  // Buka kunci dan bersihkan entri uji
  const unlocked = store.toggleMasterTemplate(created.id);
  assert.equal(unlocked.isMasterTemplate, false);
  assert.equal(store.removeLetterVaultEntry(created.id), true);
});
