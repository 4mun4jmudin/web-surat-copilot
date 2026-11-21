/**
 * tests/proFeatures.test.js
 * Unit Test 6 Fitur Pro & Alur Resubmit Revisi Kepsek -> Surgical AI
 */

import test from 'node:test';
import assert from 'node:assert/strict';

import {
  AGENDA_MONTH_LABELS,
  extractDateYearMonth,
  filterLettersByPeriod,
  collectAvailableAgendaYears,
  formatPeriodLabel,
  exportOutgoingLettersCsv,
  exportIncomingLettersCsv,
  buildPrintableAgendaHtml
} from '../js/letters/agenda.js';
import { store } from '../js/state.js';
import {
  buildSurgicalIndexFromSmartPages,
  applySurgicalPatchesToSmartPages,
  buildSurgicalIndexFromDocModel,
  applySurgicalPatchesToDocModel
} from '../js/ai/surgicalEdit.js';
import { docModelToSmartPages } from '../js/ui/smartWordCanvas.js';
import { buildSuratTugasSppdDocModel } from '../js/letters/templates/suratTugasSppd.js';
import { DEFAULT_SCHOOL_PROFILE } from '../js/letters/schoolProfile.js';
import { handleWorkspaceRequest } from '../netlify/functions/workspace.mjs';

function createMockBlobStore() {
  const data = new Map();
  return {
    async get(key, opts = {}) {
      if (!data.has(key)) return null;
      const val = data.get(key);
      if (opts.type === 'json') return JSON.parse(JSON.stringify(val));
      return typeof val === 'string' ? val : JSON.stringify(val);
    },
    async setJSON(key, value, opts = {}) {
      if (opts.onlyIfNew && data.has(key)) return { modified: false };
      data.set(key, JSON.parse(JSON.stringify(value)));
      return { modified: true, etag: String(Date.now()) };
    }
  };
}

const TEST_ENV = {
  AUTH_SESSION_SECRET: 'test-session-secret-64chars-abcdefghijklmnopqrstuvwxyz123456',
  INITIAL_ADMIN_PASSWORD: 'TestAdmin#Pass1',
  INITIAL_KEPSEK_PASSWORD: 'TestKepsek#Pass2',
  INITIAL_TU_PASSWORD: 'TestTu#Pass3',
  GAS_WEBHOOK_URL: 'https://script.google.com/macros/s/MOCK_GAS_EXEC/exec',
  GAS_WEBHOOK_SECRET: 'mock-gas-secret-token',
  DRIVE_ROOT_FOLDER_ID: '19nFdoJ4VOEgFmUe_heAFxoFEt-9U9Crw'
};

async function callWorkspace({ wsStore, verifyStore, method = 'GET', action = '', query = null, token = '', body = null, fetchImpl }) {
  const headers = {};
  if (token) headers['Authorization'] = `Bearer ${token}`;
  if (body) headers['Content-Type'] = 'application/json';
  const params = new URLSearchParams({ action });
  if (query && typeof query === 'object') {
    for (const [k, v] of Object.entries(query)) {
      if (v !== undefined && v !== null) params.set(k, String(v));
    }
  }
  const req = new Request(`https://copilot-dokumen-alhawari.netlify.app/api/workspace?${params.toString()}`, {
    method,
    headers,
    body: body ? JSON.stringify(body) : undefined
  });
  return handleWorkspaceRequest(req, {
    workspaceStoreFactory: () => wsStore,
    verificationStoreFactory: () => verifyStore,
    env: TEST_ENV,
    fetchImpl
  });
}

test('Agenda & Rekap Resmi: Ekstraksi bulan/tahun, filter periode, ekspor CSV BOM, dan HTML rekap cetak', () => {
  assert.equal(AGENDA_MONTH_LABELS.length, 12);
  assert.deepEqual(extractDateYearMonth('2026-11-20'), { year: '2026', month: '11' });
  assert.deepEqual(extractDateYearMonth('2026-03-05T10:20:00Z'), { year: '2026', month: '03' });
  assert.deepEqual(extractDateYearMonth('Garut, 18 November 2026'), { year: '2026', month: '11' });
  assert.deepEqual(extractDateYearMonth('14 Januari 2025'), { year: '2025', month: '01' });
  assert.equal(extractDateYearMonth(''), null);

  const sampleItems = [
    { id: '1', nomorSurat: '800/001/ST.SMKITA/XI/2026', typeCode: 'ST', title: 'Surat Tugas PKL', perihal: 'PKL', recipientName: 'Rina, S.Pd.', tanggalSurat: '18 November 2026', sequenceNumber: 1, approvedByName: 'Kepala Sekolah', qrStatus: 'active' },
    { id: '2', nomorSurat: '421/002/SK.SMKITA/I/2025', typeCode: 'SK', title: 'Surat Keterangan Siswa', perihal: 'Aktif Sekolah', recipientName: 'Budi', tanggalSurat: '14 Januari 2025', sequenceNumber: 2 }
  ];

  const nov2026 = filterLettersByPeriod(sampleItems, { month: '11', year: '2026', getDateValue: x => x.tanggalSurat });
  assert.equal(nov2026.length, 1);
  assert.equal(nov2026[0].id, '1');

  const all2025 = filterLettersByPeriod(sampleItems, { month: 'all', year: '2025', getDateValue: x => x.tanggalSurat });
  assert.equal(all2025.length, 1);
  assert.equal(all2025[0].id, '2');

  const years = collectAvailableAgendaYears(sampleItems, x => x.tanggalSurat);
  assert.ok(years.includes('2026') && years.includes('2025'));
  assert.equal(formatPeriodLabel('11', '2026'), 'November · Tahun 2026');
  assert.equal(formatPeriodLabel('all', 'all'), 'Semua Bulan · Semua Tahun');

  const outCsv = exportOutgoingLettersCsv(sampleItems, { schoolName: 'SMKS IT Al-Hawari', periodLabel: 'Tahun 2026' });
  assert.ok(outCsv.startsWith('\uFEFF'), 'CSV harus diawali UTF-8 BOM untuk Microsoft Excel');
  assert.ok(outCsv.includes('800/001/ST.SMKITA/XI/2026'));
  assert.ok(outCsv.includes('Surat Tugas PKL'));

  const inCsv = exportIncomingLettersCsv([
    {
      kodeSurat: 'SM-001/XI/2026',
      nomorSuratAsli: '099/DISDIK/2026',
      namaSurat: 'Undangan Sosialisasi, Penting',
      pengirim: 'Dinas Pendidikan',
      kategori: 'Undangan',
      tanggalSurat: '2026-11-19',
      driveSynced: true,
      driveUrl: 'https://drive.google.com/file/d/abc/view'
    }
  ], { schoolName: 'SMKS IT Al-Hawari', periodLabel: 'November · Tahun 2026' });
  assert.ok(inCsv.startsWith('\uFEFF'));
  assert.ok(inCsv.includes('SM-001/XI/2026'));
  assert.ok(inCsv.includes('"Undangan Sosialisasi, Penting"'));

  const printHtmlOut = buildPrintableAgendaHtml({
    kind: 'outgoing',
    school: { name: 'SMKS IT AL-HAWARI', address: 'Garut', principalName: 'Dian Nurhamdani, S.Pd.' },
    periodLabel: 'November · Tahun 2026',
    items: sampleItems
  });
  assert.ok(printHtmlOut.includes('BUKU AGENDA SURAT KELUAR'));
  assert.ok(printHtmlOut.includes('800/001/ST.SMKITA/XI/2026'));

  const printHtmlIn = buildPrintableAgendaHtml({
    kind: 'incoming',
    school: { name: 'SMKS IT AL-HAWARI' },
    periodLabel: 'Semua Periode',
    items: [{ kodeSurat: 'SM-002/XI/2026', namaSurat: '<script>alert(1)</script> Undangan', pengirim: 'Kemdikbud', tanggalSurat: '2026-11-20' }]
  });
  assert.ok(printHtmlIn.includes('BUKU AGENDA SURAT MASUK'));
  assert.ok(!printHtmlIn.includes('<script>alert(1)</script>'), 'HTML rekap harus mengamankan XSS');
});

test('Version History Gudang Surat (Max 10, Dedupe, Restore) & State Pending Revisi Kepsek', () => {
  const initialLen = store.getLetterVault().length;
  const entry = store.addLetterVaultEntry({
    title: 'Draf Uji Versi v1',
    nomorSurat: '001/TEST/XI/2026',
    perihal: 'Uji Riwayat Versi',
    summary: 'Ringkasan v1',
    isAiReference: false,
    source: 'created',
    document: { version: 1, blocks: [{ type: 'paragraph', text: 'Konten v1' }] }
  });

  try {
    assert.ok(Array.isArray(entry.versionHistory) && entry.versionHistory.length === 1);
    const v1Id = entry.versionHistory[0].versionId;

    // Update dengan dokumen identik -> tidak menduplikasi snapshot kecuali forceVersionSnapshot
    store.updateLetterVaultEntry(entry.id, { summary: 'Ringkasan v1 edit' });
    let current = store.getLetterVault().find(l => l.id === entry.id);
    assert.equal(current.versionHistory.length, 1);

    // Tambah 11 versi dokumen berbeda -> harus terpotong maksimal 10 snapshot
    for (let i = 2; i <= 12; i++) {
      store.updateLetterVaultEntry(
        entry.id,
        {
          title: `Draf Uji Versi v${i}`,
          document: { version: 1, blocks: [{ type: 'paragraph', text: `Konten v${i}` }] }
        },
        { versionNote: `Revisi ke-${i}` }
      );
    }

    current = store.getLetterVault().find(l => l.id === entry.id);
    assert.equal(current.versionHistory.length, 10, 'Riwayat versi maksimal 10 snapshot terbaru');
    assert.equal(current.title, 'Draf Uji Versi v12');

    // Pulihkan versi lama (v5)
    const targetV5 = current.versionHistory.find(v => v.title === 'Draf Uji Versi v5');
    assert.ok(targetV5, 'Versi v5 harus tersedia di dalam 10 snapshot terbaru');
    const restored = store.restoreLetterVaultVersion(entry.id, targetV5.versionId);
    assert.ok(restored);
    assert.equal(restored.title, 'Draf Uji Versi v5');
    assert.equal(restored.document.blocks[0].text, 'Konten v5');
    assert.ok(restored.versionHistory.length <= 10);

    // v1 sudah tergeser karena batas 10 versi; v5 & snapshot pemulihan tetap aman
    assert.equal(current.versionHistory.some(v => v.versionId === v1Id), false);

    // State pendingKepsekRevision
    store.clearPendingKepsekRevision();
    assert.equal(store.getPendingKepsekRevision(), null);
    store.setPendingKepsekRevision({
      accId: 'acc_test_1',
      reason: 'Perbaiki tanggal ke 25 November 2026',
      autoRunAi: true
    });
    const rev = store.getPendingKepsekRevision();
    assert.equal(rev.accId, 'acc_test_1');
    assert.equal(rev.autoRunAi, true);
    store.clearPendingKepsekRevision();
    assert.equal(store.getPendingKepsekRevision(), null);
  } finally {
    store.removeLetterVaultEntry(entry.id);
    assert.equal(store.getLetterVault().length, initialLen);
  }
});

test('Surgical AI Compare Log: SmartPages dan DocModel menghasilkan rincian changes (Before -> After)', () => {
  const sppdModel = buildSuratTugasSppdDocModel({}, DEFAULT_SCHOOL_PROFILE, { paperSize: 'f4' });
  const converted = docModelToSmartPages(sppdModel, DEFAULT_SCHOOL_PROFILE);
  const idxSmart = buildSurgicalIndexFromSmartPages(converted.pages, DEFAULT_SCHOOL_PROFILE);
  const firstTableSmart = idxSmart.snapshot.pages[0].blocks.find(b => b.type === 'table');

  const smartRes = applySurgicalPatchesToSmartPages(converted.pages, {
    editScope: 'surgical',
    patches: [
      { op: 'setTableCell', blockId: firstTableSmart.blockId, row: 1, col: 1, text: 'Ahmad Fauzi, S.Pd.' }
    ]
  }, { schoolProfile: DEFAULT_SCHOOL_PROFILE, allowLayoutRestructure: false });

  assert.equal(smartRes.appliedCount, 1);
  assert.ok(Array.isArray(smartRes.changes) && smartRes.changes.length === 1);
  assert.equal(smartRes.changes[0].kind, 'tableCell');
  assert.ok(smartRes.changes[0].after.includes('Ahmad Fauzi'));

  const idxDoc = buildSurgicalIndexFromDocModel(sppdModel);
  const firstTableDoc = idxDoc.snapshot.pages[0].blocks.find(b => b.type === 'table');
  const docRes = applySurgicalPatchesToDocModel(sppdModel, {
    editScope: 'surgical',
    patches: [
      { op: 'setField', target: 'nomor', text: '800/888/ST.SMKITA/XI/2026' },
      { op: 'setTableCell', blockId: firstTableDoc.blockId, row: 1, col: 1, text: 'Rina Marlina, S.Pd.' }
    ]
  }, { allowLayoutRestructure: false });

  assert.equal(docRes.appliedCount, 2);
  assert.ok(Array.isArray(docRes.changes) && docRes.changes.length === 2);
  assert.equal(docRes.changes[0].kind, 'field');
  assert.equal(docRes.changes[1].kind, 'tableCell');
});

test('Resubmit ACC setelah Revisi Kepsek memperbarui documentSnapshot & metadata terbaru di Cloud', async () => {
  const wsStore = createMockBlobStore();
  const verifyStore = createMockBlobStore();
  const mockFetch = async () => new Response(JSON.stringify({ ok: true }), { status: 200, headers: { 'Content-Type': 'application/json' } });

  const tuLogin = await callWorkspace({
    wsStore, verifyStore, method: 'POST', action: 'login',
    body: { username: 'tu', password: TEST_ENV.INITIAL_TU_PASSWORD }
  });
  assert.equal(tuLogin.status, 200);
  const tuToken = (await tuLogin.json()).token;

  const ksLogin = await callWorkspace({
    wsStore, verifyStore, method: 'POST', action: 'login',
    body: { username: 'kepsek', password: TEST_ENV.INITIAL_KEPSEK_PASSWORD }
  });
  assert.equal(ksLogin.status, 200);
  const ksToken = (await ksLogin.json()).token;

  // 1. TU mencatat Surat Keluar dan mengajukan ACC awal
  const req1 = await callWorkspace({
    wsStore, verifyStore, method: 'POST', action: 'record_outgoing', token: tuToken, fetchImpl: mockFetch,
    body: {
      authorizeSah: false,
      requestAcc: true,
      signerRole: 'kepala_sekolah',
      requestNote: 'Mohon ACC surat tugas',
      entry: {
        typeCode: 'ST',
        year: '2026',
        nomorSurat: '800/010/ST.SMKITA/XI/2026',
        title: 'Surat Tugas Uji v1',
        recipientName: 'Ujang',
        perihal: 'PKL',
        tanggalSurat: 'Garut, 18 November 2026',
        documentSnapshot: { blocks: [{ type: 'paragraph', text: 'Versi Lama' }] }
      }
    }
  });
  assert.equal(req1.status, 201);
  const req1Json = await req1.json();
  const outgoingId = req1Json.item.id;

  // 2. Kepala Sekolah menolak dengan catatan revisi
  const rejRes = await callWorkspace({
    wsStore, verifyStore, method: 'POST', action: 'reject_acc', token: ksToken, fetchImpl: mockFetch,
    body: { outgoingId, reason: 'Ganti tanggal surat ke 25 November 2026' }
  });
  assert.equal(rejRes.status, 200);
  const rejJson = await rejRes.json();
  assert.equal(rejJson.item.accStatus, 'rejected');

  // 3. TU mengajukan ulang ACC dengan documentSnapshot yang sudah diperbaiki AI Surgical
  const req2 = await callWorkspace({
    wsStore, verifyStore, method: 'POST', action: 'request_acc', token: tuToken, fetchImpl: mockFetch,
    body: {
      outgoingId,
      signerRole: 'kepala_sekolah',
      requestNote: 'Sudah direvisi sesuai catatan Kepsek',
      title: 'Surat Tugas Uji v2 (Revisi)',
      nomorSurat: '800/010/ST.SMKITA/XI/2026',
      recipientName: 'Ujang, S.Pd.',
      perihal: 'PKL Revisi',
      tanggalSurat: 'Garut, 25 November 2026',
      summary: 'Tanggal telah diperbarui menjadi 25 November 2026',
      documentSnapshot: { blocks: [{ type: 'paragraph', text: 'Versi Revisi 25 November 2026' }] }
    }
  });
  assert.equal(req2.status, 200);
  const req2Json = await req2.json();
  assert.equal(req2Json.item.verificationStatus, 'pending_acc');
  assert.equal(req2Json.item.accStatus, 'pending');
  assert.equal(req2Json.item.rejectedReason, '');
  assert.equal(req2Json.item.rejectedAt, null);
  assert.equal(req2Json.item.title, 'Surat Tugas Uji v2 (Revisi)');
  assert.equal(req2Json.item.tanggalSurat, 'Garut, 25 November 2026');
  assert.equal(req2Json.item.documentSnapshot.blocks[0].text, 'Versi Revisi 25 November 2026');
});
