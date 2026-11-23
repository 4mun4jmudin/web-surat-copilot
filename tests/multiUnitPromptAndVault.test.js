import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  SCHOOL_UNITS,
  detectTargetUnitFromText,
  resolveSchoolProfileForUnit,
  adaptDocumentModelForUnit,
  adaptTextForSchoolUnit
} from '../js/letters/schoolProfile.js';
import { DEFAULT_KOP_DATA_URL, DEFAULT_SMP_KOP_DATA_URL } from '../js/letters/defaultKop.js';
import { formatSmartNomorSurat } from '../js/letters/numbering.js';
import { buildChatCopilotUserPrompt, normalizeAiChatResponse } from '../js/ai/chatEngine.js';
import { callAiChatTurn } from '../js/ai/client.js';
import { buildSignerSelectionInfo, buildVerificationPayload } from '../js/letters/verification.js';
import { store, DEFAULT_LETTER_VAULT } from '../js/state.js';

test('Deteksi Otomatis Unit SMP vs SMK & 2 Pejabat Kepala Sekolah', () => {
  assert.equal(detectTargetUnitFromText('Buatkan surat keterangan siswa aktif untuk SMP'), 'smp');
  assert.equal(detectTargetUnitFromText('Buat surat tugas PKL untuk siswa XII DKV SMK'), 'smk');
  assert.equal(detectTargetUnitFromText('Surat undangan rapat ditandatangani oleh Yangyang Budiman, S.Pd.'), 'smp');
  assert.equal(detectTargetUnitFromText('Surat tugas kepala sekolah Dian Nurhamdani, S.Pd.'), 'smk');
  // Frasa target eksplisit mengalahkan penyebutan unit sumber sebelumnya
  assert.equal(detectTargetUnitFromText('Ubah template surat SMK ini untuk SMP pakai kop dan kepsek SMP'), 'smp');
  assert.equal(detectTargetUnitFromText('Sesuaikan surat SMP tersebut ke unit SMK'), 'smk');
  assert.equal(detectTargetUnitFromText('Buat surat undangan rapat guru besok', 'smk'), 'smk');
});

test('Penomoran & Adaptor Dokumen Gudang Surat ke SMP vs SMK + Kepsek Otomatis', () => {
  assert.equal(
    formatSmartNomorSurat({
      rawNomor: '421.3/085/KET.SMKITA/XI/2026',
      forcedTypeCode: 'KET',
      schoolCode: 'SMPITA'
    }),
    '421.2/085/KET.SMPITA/XI/2026'
  );

  // Uji adaptasi template SKB (vault-ref-001) & Surat Tugas SPPD 3 halaman (vault-ref-004)
  const skbVault = DEFAULT_LETTER_VAULT.find(v => v.id === 'vault-ref-001');
  const sppdVault = DEFAULT_LETTER_VAULT.find(v => v.id === 'vault-ref-004');
  assert.ok(skbVault?.document);
  assert.ok(sppdVault?.document);

  const smpSkb = adaptDocumentModelForUnit(skbVault.document, 'smp');
  assert.equal(smpSkb.targetUnitId, 'smp');
  const smpKop = smpSkb.blocks.find(b => b.type === 'letterKop');
  const smpTitle = smpSkb.blocks.find(b => b.type === 'letterTitle');
  const smpSig = smpSkb.blocks.find(b => b.type === 'letterSignature');
  assert.equal(smpKop.school.npsn, '20279898');
  assert.equal(smpKop.image, DEFAULT_SMP_KOP_DATA_URL);
  assert.ok(smpTitle.nomor.includes('421.2') && smpTitle.nomor.includes('SMPITA'));
  assert.equal(smpSig.name, 'Yangyang Budiman, S.Pd.');

  const smkSkb = adaptDocumentModelForUnit(smpSkb, 'smk');
  assert.equal(smkSkb.targetUnitId, 'smk');
  const smkKop = smkSkb.blocks.find(b => b.type === 'letterKop');
  const smkTitle = smkSkb.blocks.find(b => b.type === 'letterTitle');
  const smkSig = smkSkb.blocks.find(b => b.type === 'letterSignature');
  assert.equal(smkKop.school.npsn, '69758993');
  assert.equal(smkKop.image, DEFAULT_KOP_DATA_URL);
  assert.ok(smkTitle.nomor.includes('421.3') && smkTitle.nomor.includes('SMKITA'));
  assert.equal(smkSig.name, 'Dian Nurhamdani, S.Pd.');

  // Verifikasi QR Payload memakai pejabat yang terpilih otomatis untuk SMP
  const smpSigner = buildSignerSelectionInfo(smpSkb.letterData.schoolProfile, 'kepala_sekolah');
  assert.equal(smpSigner.signerName, 'Yangyang Budiman, S.Pd.');
  const payload = buildVerificationPayload({
    entry: {
      nomorSurat: smpTitle.nomor.replace(/^Nomor:\s*/i, ''),
      title: 'SURAT KETERANGAN SISWA AKTIF',
      recipientName: 'Ahmad Fauzi',
      typeCode: 'KET',
      year: '2026',
      documentSnapshot: smpSkb
    },
    school: smpSkb.letterData.schoolProfile,
    signerRole: 'kepala_sekolah'
  });
  assert.equal(payload.schoolName, 'SMP IT AL-HAWARI - CIBIUK');
  assert.equal(payload.signerName, 'Yangyang Budiman, S.Pd.');

  // Uji adaptasi SPPD 3 halaman ke SMP mengganti teks SMK & Kepsek di dalam sel tabel
  const smpSppd = adaptDocumentModelForUnit(sppdVault.document, 'smp');
  assert.equal(smpSppd.targetUnitId, 'smp');
  assert.equal(smpSppd.letterData.schoolProfile.principalName, 'Yangyang Budiman, S.Pd.');
  const serializedSppd = JSON.stringify(smpSppd.blocks);
  assert.ok(serializedSppd.includes('Yangyang Budiman, S.Pd.'));
  assert.ok(serializedSppd.includes('SMPITA'));
  assert.equal(/Dian Nurh/i.test(serializedSppd), false);
});

test('StateStore.loadLetterToCanvas Memilih Unit SMP/SMK & Menjaga Konteks Dokumen Aktif', () => {
  // Workspace default SMK, tapi user membuka file Gudang untuk SMP
  store.setState({
    activeUnitId: 'smk',
    schoolProfile: resolveSchoolProfileForUnit('smk')
  }, false);

  const loadedSmp = store.loadLetterToCanvas('vault-ref-001', { targetUnitId: 'smp' });
  assert.equal(loadedSmp, true);
  const docAfterSmp = store.getState().document;
  assert.equal(docAfterSmp.targetUnitId, 'smp');
  const sigSmp = docAfterSmp.blocks.find(b => b.type === 'letterSignature');
  assert.equal(sigSmp.name, 'Yangyang Budiman, S.Pd.');
  const effectiveSmp = store.getEffectiveSchoolForDocument(docAfterSmp);
  assert.equal(effectiveSmp.unitId, 'smp');
  assert.equal(effectiveSmp.principalName, 'Yangyang Budiman, S.Pd.');

  // Sekarang buka file SPPD dari Gudang untuk SMK -> kembali otomatis Dian Nurhamdani, S.Pd.
  const loadedSppdSmk = store.loadLetterToCanvas('vault-ref-004', { targetUnitId: 'smk' });
  assert.equal(loadedSppdSmk, true);
  const smartState = store.getState().smartDocument;
  assert.equal(smartState.targetUnitId, 'smk');
  assert.equal(smartState.schoolProfile.principalName, 'Dian Nurhamdani, S.Pd.');
});

test('Prompt Builder & AI Turn Detection Mengunci Kepsek SMP vs SMK Sesuai Perintah AI', async () => {
  const smpPrompt = buildChatCopilotUserPrompt({
    userMessage: 'Buatkan surat undangan rapat guru SMP',
    schoolName: SCHOOL_UNITS.smp.schoolName,
    principalName: SCHOOL_UNITS.smp.principalName,
    schoolCode: 'SMPITA',
    unitId: 'smp'
  });
  assert.ok(smpPrompt.includes('Yangyang Budiman, S.Pd.'));
  assert.ok(smpPrompt.includes('SMPITA'));
  assert.ok(smpPrompt.includes('UNIT TARGET TERDETEKSI: SMP'));

  // Simulasi LLM yang keliru mengisi Kepsek SMK pada perintah SMP -> normalisasi mengoreksi deterministik
  const correctedSmp = normalizeAiChatResponse({
    intent: 'create_letter',
    reply: 'Surat undangan untuk SMP sudah dibuat.',
    letterDraft: {
      title: 'Undangan Rapat SMP',
      letterTitle: 'SURAT UNDANGAN RAPAT',
      kodeJenisSurat: 'UND',
      nomorSurat: '005/010/UND.SMKITA/XI/2026',
      blocks: [{ type: 'paragraph', text: 'Dengan hormat...' }],
      signer: { title: 'Kepala Sekolah', name: 'Dian Nurhamdani, S.Pd.' }
    }
  }, {
    userMessage: 'Buatkan surat undangan rapat untuk SMP',
    targetUnitId: 'smp',
    targetSchoolProfile: resolveSchoolProfileForUnit('smp')
  });
  assert.equal(correctedSmp.letterDraft.targetUnit, 'smp');
  assert.equal(correctedSmp.letterDraft.signer.name, 'Yangyang Budiman, S.Pd.');
  assert.equal(correctedSmp.letterDraft.nomorSurat, '005/010/UND.SMPITA/XI/2026');

  // Uji integrasi callAiChatTurn (mock fetch & API vault lokal) dari workspace SMK meminta surat SMP
  store.setState({
    activeUnitId: 'smk',
    schoolProfile: resolveSchoolProfileForUnit('smk')
  }, false);

  const prevFetch = globalThis.fetch;
  let capturedApiBodyText = '';
  globalThis.fetch = async (_url, init = {}) => {
    capturedApiBodyText = String(init.body || '');
    return new Response(JSON.stringify({
      candidates: [{
        content: {
          parts: [{
            text: JSON.stringify({
              intent: 'create_letter',
              reply: 'Berikut surat keterangan SMP yang telah dibuat.',
              checklist: ['Menggunakan Kop SMP IT Al-Hawari', 'TTD Kepala Sekolah Yangyang Budiman, S.Pd.'],
              letterDraft: {
                targetUnit: 'smp',
                title: 'Surat Keterangan SMP',
                letterTitle: 'SURAT KETERANGAN SISWA AKTIF',
                kodeJenisSurat: 'KET',
                nomorSurat: '421.2/001/KET.SMPITA/XI/2026',
                blocks: [{ type: 'paragraph', text: 'Kepala SMP IT Al-Hawari menerangkan bahwa siswa tersebut aktif.' }],
                signer: { title: 'Kepala Sekolah', name: 'Dian Nurhamdani, S.Pd.' }
              }
            })
          }]
        }
      }]
    }), { status: 200, headers: { 'Content-Type': 'application/json' } });
  };

  try {
    const aiRes = await callAiChatTurn({
      userMessage: 'Tolong buatkan surat keterangan aktif untuk siswa kelas 8 SMP',
      preferredApiId: 'gemini-auto',
      executionMode: 'direct_apply'
    });
    assert.equal(aiRes.targetUnitId, 'smp');
    assert.equal(aiRes.schoolProfile.principalName, 'Yangyang Budiman, S.Pd.');
    assert.ok(capturedApiBodyText.includes('Yangyang Budiman, S.Pd.'));
    assert.ok(capturedApiBodyText.includes('SMPITA'));
    assert.equal(aiRes.letterDraft.signer.name, 'Yangyang Budiman, S.Pd.');
    assert.equal(adaptTextForSchoolUnit('Kepala SMK IT AL-HAWARI', 'smp'), 'Kepala SMP IT AL-HAWARI - CIBIUK');
  } finally {
    globalThis.fetch = prevFetch;
  }
});
