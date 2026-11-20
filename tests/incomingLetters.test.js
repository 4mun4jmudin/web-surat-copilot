import test from 'node:test';
import assert from 'node:assert/strict';

import { handleWorkspaceRequest } from '../netlify/functions/workspace.mjs';
import {
  APPS_SCRIPT_FULL_BRIDGE_CODE,
  APPS_SCRIPT_INCOMING_SNIPPET,
  detectIncomingFileKind,
  formatIncomingCode,
  getNextIncomingSequence,
  prepareIncomingFilePayload,
  scoreIncomingMatch,
  sortIncomingLetters
} from '../js/letters/incoming.js';
import { store } from '../js/state.js';

function createMockBlobStore() {
  const data = new Map();
  return {
    async get(key, opts = {}) {
      if (!data.has(key)) return null;
      const val = data.get(key);
      if (opts.type === 'json') {
        return JSON.parse(JSON.stringify(val));
      }
      return typeof val === 'string' ? val : JSON.stringify(val);
    },
    async setJSON(key, value, opts = {}) {
      if (opts.onlyIfNew && data.has(key)) {
        return { modified: false };
      }
      data.set(key, JSON.parse(JSON.stringify(value)));
      return { modified: true, etag: String(Date.now()) };
    },
    async delete(key) {
      data.delete(key);
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

test('Helper Surat Masuk: kode urut, deteksi jenis berkas (Gambar/PDF/Foto), pencarian typo-tolerant, & snippet GAS', () => {
  const code = formatIncomingCode({ sequenceNumber: 4, date: new Date('2026-11-20T08:00:00Z') });
  assert.equal(code, 'SM-004/XI/2026');

  const nextSeq = getNextIncomingSequence([
    { year: '2026', sequenceNumber: 1 },
    { year: '2026', sequenceNumber: 5 },
    { year: '2025', sequenceNumber: 20 }
  ], 2026);
  assert.equal(nextSeq, 6);

  assert.equal(detectIncomingFileKind('undangan.pdf', 'application/pdf'), 'pdf');
  assert.equal(detectIncomingFileKind('scan.png', 'image/png', 'upload_image'), 'image');
  assert.equal(detectIncomingFileKind('kamera.jpg', 'image/jpeg', 'camera_photo'), 'photo');

  const items = [
    { kodeSurat: 'SM-001/XI/2026', namaSurat: 'Undangan Rapat Dinas Cabdin', pengirim: 'Cabang Dinas Pendidikan Garut', fileKind: 'pdf' },
    { kodeSurat: 'SM-002/XI/2026', namaSurat: 'Pemberitahuan Uji Kompetensi', pengirim: 'BNSP Pusat', fileKind: 'photo' }
  ];
  assert.ok(scoreIncomingMatch(items[0], 'SM-001') > 0);
  assert.ok(scoreIncomingMatch(items[0], 'undngan cabdn') > 0, 'Pencarian harus mendukung typo');
  assert.equal(scoreIncomingMatch(items[1], 'SM-001'), 0);

  const sortedByCode = sortIncomingLetters([items[1], items[0]], 'code_asc');
  assert.equal(sortedByCode[0].kodeSurat, 'SM-001/XI/2026');
  assert.ok(APPS_SCRIPT_INCOMING_SNIPPET.includes("action === 'archive_incoming'"));
  assert.ok(APPS_SCRIPT_INCOMING_SNIPPET.includes('handleIncomingArchive_'));
  assert.ok(APPS_SCRIPT_FULL_BRIDGE_CODE.includes("action === 'archive_outgoing'") && APPS_SCRIPT_FULL_BRIDGE_CODE.includes("action === 'archive_incoming'"));
});

test('Backend Surat Masuk: Upload PDF/Gambar/Foto langsung simpan ke Drive + Bootstrap + Poll + Retry Drive + Hapus RBAC', async () => {
  const wsStore = createMockBlobStore();
  const verifyStore = createMockBlobStore();
  const gasRequests = [];
  let gasMode = 'ok';

  const mockFetch = async (_url, init) => {
    const parsed = JSON.parse(init.body);
    gasRequests.push(parsed);
    if (gasMode === 'missing_action') {
      return new Response(JSON.stringify({ ok: false, error: 'Aksi tidak dikenali: archive_incoming' }), { status: 200 });
    }
    return new Response(JSON.stringify({
      ok: true,
      driveFileId: `drv_in_${gasRequests.length}`,
      driveUrl: `https://drive.google.com/file/d/drv_in_${gasRequests.length}/view`,
      driveDownloadUrl: `https://drive.google.com/uc?export=download&id=drv_in_${gasRequests.length}`,
      rootFolderUrl: 'https://drive.google.com/drive/folders/19nFdoJ4VOEgFmUe_heAFxoFEt-9U9Crw',
      sheetUrl: 'https://docs.google.com/spreadsheets/d/SHEET_MASTER/edit',
      syncedAt: new Date().toISOString()
    }), { status: 200 });
  };

  const loginTu = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'login',
    body: { username: 'tu', password: TEST_ENV.INITIAL_TU_PASSWORD }
  });
  const tuToken = (await loginTu.json()).token;

  const loginAdmin = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'login',
    body: { username: 'admin', password: TEST_ENV.INITIAL_ADMIN_PASSWORD }
  });
  const adminToken = (await loginAdmin.json()).token;

  // 1. Upload Foto Langsung / Gambar oleh TU -> langsung tersimpan ke Google Drive
  const recImgRes = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'record_incoming',
    token: tuToken,
    body: {
      entry: {
        kodeSurat: 'SM-001/XI/2026',
        namaSurat: 'Surat Undangan Koordinasi Kepala SMK',
        pengirim: 'Dinas Pendidikan Jabar',
        nomorSuratAsli: '005/442/DISDIK/XI/2026',
        tanggalSurat: '2026-11-20',
        kategori: 'Undangan',
        captureSource: 'camera_photo',
        fileName: 'SM-001_foto.jpg',
        mimeType: 'image/jpeg',
        previewThumbDataUrl: 'data:image/jpeg;base64,/9j/4AAQ'
      },
      fileBase64: Buffer.from('FAKE_JPEG_BINARY_DATA').toString('base64')
    },
    fetchImpl: mockFetch
  });
  assert.equal(recImgRes.status, 201);
  const recImgJson = await recImgRes.json();
  assert.equal(recImgJson.ok, true);
  assert.equal(recImgJson.item.kodeSurat, 'SM-001/XI/2026');
  assert.equal(recImgJson.item.fileKind, 'photo');
  assert.equal(recImgJson.item.driveSynced, true);
  assert.equal(recImgJson.item.driveFileId, 'drv_in_1');
  assert.match(recImgJson.item.driveUrl, /drive\.google\.com/);
  assert.equal(gasRequests[0].action, 'archive_incoming');
  assert.equal(gasRequests[0].mimeType, 'image/jpeg');

  // 2. Upload PDF dengan Bridge belum update (simulasi fallback aman + Sync Ulang Drive kemudian)
  gasMode = 'missing_action';
  const recPdfRes = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'record_incoming',
    token: tuToken,
    body: {
      entry: {
        namaSurat: 'Edaran Pelaksanaan PAS Ganjil 2026',
        pengirim: 'MKKS Kabupaten Garut',
        captureSource: 'upload_pdf',
        fileName: 'edaran_pas.pdf',
        mimeType: 'application/pdf'
      },
      fileBase64: Buffer.from('%PDF-1.4 MOCK PDF CONTENT').toString('base64')
    },
    fetchImpl: mockFetch
  });
  assert.equal(recPdfRes.status, 201);
  const recPdfJson = await recPdfRes.json();
  assert.equal(recPdfJson.item.fileKind, 'pdf');
  assert.equal(recPdfJson.item.driveSynced, false, 'Tidak boleh menandai tersinkron Drive jika GAS menolak archive_incoming');
  assert.equal(recPdfJson.driveSync.needGasUpdate, true, 'Harus memberi tanda needGasUpdate agar UI membuka panduan Code.gs');
  assert.ok(recPdfJson.item.kodeSurat.startsWith('SM-002/'));

  const retryNeedUpdateRes = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'sync_incoming_drive',
    token: adminToken,
    body: { incomingId: recPdfJson.item.id },
    fetchImpl: mockFetch
  });
  assert.equal(retryNeedUpdateRes.status, 502);
  const retryNeedUpdateJson = await retryNeedUpdateRes.json();
  assert.equal(retryNeedUpdateJson.needGasUpdate, true);

  // 3. Cek Bootstrap & Poll Updates menyertakan daftar Surat Masuk
  const bootRes = await callWorkspace({ wsStore, verifyStore, method: 'GET', action: 'bootstrap', token: adminToken });
  const bootJson = await bootRes.json();
  assert.equal(bootJson.workspace.incomingLetters.length, 2);

  const pollRes = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'GET',
    action: 'poll_updates',
    query: { since: 1 },
    token: tuToken
  });
  const pollJson = await pollRes.json();
  assert.equal(pollJson.changed, true);
  assert.equal(pollJson.workspace.incomingLetters.length, 2);

  // 4. Pratinjau file biner dari cloud blob bekerja untuk perangkat lain (Kepsek/Admin)
  const fileGetRes = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'GET',
    action: 'get_incoming_file',
    query: { id: recPdfJson.item.id },
    token: adminToken
  });
  assert.equal(fileGetRes.status, 200);
  const fileGetJson = await fileGetRes.json();
  assert.equal(fileGetJson.mimeType, 'application/pdf');
  assert.equal(Buffer.from(fileGetJson.fileBase64, 'base64').toString('utf8'), '%PDF-1.4 MOCK PDF CONTENT');

  // 5. Setelah script Drive siap, Sync Ulang Drive berhasil menautkan file PDF ke Drive
  gasMode = 'ok';
  const retryRes = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'sync_incoming_drive',
    token: adminToken,
    body: { incomingId: recPdfJson.item.id },
    fetchImpl: mockFetch
  });
  assert.equal(retryRes.status, 200);
  const retryJson = await retryRes.json();
  assert.equal(retryJson.item.driveSynced, true);
  assert.ok(retryJson.item.driveUrl.includes('drv_in_'));

  // 6. Store applyRemoteBootstrap & applyLiveWorkspaceDelta menyinkronkan daftar ke state frontend
  store.applyRemoteBootstrap(bootJson);
  assert.equal(store.getIncomingLetters().length, 2);
  store.applyLiveWorkspaceDelta({
    ok: true,
    changed: true,
    rev: 999999,
    workspace: { incomingLetters: retryJson.incomingLetters }
  });
  const pdfFromStore = store.getIncomingLetters().find(x => x.id === recPdfJson.item.id);
  assert.equal(pdfFromStore.driveSynced, true);

  // 7. RBAC: Jika Admin kunci izin hapus Surat Masuk untuk TU, TU ditolak (403); Admin tetap bisa menghapus & memanggil delete_incoming
  await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'update_permissions',
    token: adminToken,
    body: { permissions: { allowTuDeleteIncoming: false } }
  });
  const delDenied = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'delete_incoming',
    token: tuToken,
    body: { incomingId: recImgJson.item.id },
    fetchImpl: mockFetch
  });
  assert.equal(delDenied.status, 403);

  const delAdmin = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'delete_incoming',
    token: adminToken,
    body: { incomingId: recImgJson.item.id },
    fetchImpl: mockFetch
  });
  assert.equal(delAdmin.status, 200);
  const delJson = await delAdmin.json();
  assert.equal(delJson.remaining, 1);
  assert.ok(gasRequests.some(r => r.action === 'delete_incoming' && r.driveFileId === 'drv_in_1'));

  const missingAfterDel = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'GET',
    action: 'get_incoming_file',
    query: { id: recImgJson.item.id },
    token: adminToken
  });
  assert.equal(missingAfterDel.status, 404);
});
