import test from 'node:test';
import assert from 'node:assert/strict';

import { handleWorkspaceRequest } from '../netlify/functions/workspace.mjs';
import { handleVerifyRequest } from '../netlify/functions/verify-letter.mjs';
import {
  canApproveOrRejectAcc,
  canDeleteOutgoingItem,
  canManageApiVault,
  canRevokeOutgoingItem,
  getPendingAccLetters,
  getSelectableSignerRoles,
  isCanvasLockedForUser,
  resolveDefaultAllowedSignerRole
} from '../js/auth/permissions.js';

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

async function loginAs(wsStore, verifyStore, username, password) {
  const res = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'login',
    body: { username, password }
  });
  const json = await res.json();
  return { status: res.status, json };
}

test('RBAC & Delegasi Pengesahan: TU terkunci sampai diizinkan Kepala Sekolah atau Admin', async () => {
  const wsStore = createMockBlobStore();
  const verifyStore = createMockBlobStore();
  const driveCalls = [];

  const mockFetch = async (url, init) => {
    const parsed = JSON.parse(init.body);
    driveCalls.push({ url, parsed });
    return new Response(JSON.stringify({
      ok: true,
      driveDocxFileId: `drv_docx_${driveCalls.length}`,
      driveDocxUrl: `https://drive.google.com/file/d/drv_docx_${driveCalls.length}/view`,
      driveJsonFileId: `drv_json_${driveCalls.length}`,
      driveJsonUrl: `https://drive.google.com/file/d/drv_json_${driveCalls.length}/view`,
      rootFolderUrl: 'https://drive.google.com/drive/folders/19nFdoJ4VOEgFmUe_heAFxoFEt-9U9Crw',
      sheetUrl: 'https://docs.google.com/spreadsheets/d/mock_sheet/edit',
      syncedAt: new Date().toISOString()
    }), { status: 200 });
  };

  // 1. Wrong password rejected
  const badLogin = await loginAs(wsStore, verifyStore, 'tu', 'salah');
  assert.equal(badLogin.status, 401);

  // 2. Valid logins for Admin, Kepsek, TU
  const adminLogin = await loginAs(wsStore, verifyStore, 'admin', TEST_ENV.INITIAL_ADMIN_PASSWORD);
  assert.equal(adminLogin.status, 200);
  assert.equal(adminLogin.json.user.role, 'admin');

  const kepsekLogin = await loginAs(wsStore, verifyStore, 'kepsek', TEST_ENV.INITIAL_KEPSEK_PASSWORD);
  assert.equal(kepsekLogin.status, 200);
  assert.equal(kepsekLogin.json.user.role, 'kepsek');

  const tuLogin = await loginAs(wsStore, verifyStore, 'tu', TEST_ENV.INITIAL_TU_PASSWORD);
  assert.equal(tuLogin.status, 200);
  assert.equal(tuLogin.json.user.role, 'tu');
  assert.equal(tuLogin.json.permissions.allowTuSignPrincipal, false);

  // 3. TU mencoba sahkan atas nama Kepala Sekolah saat izin FALSE -> 403 Forbidden
  const deniedTuRes = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'record_outgoing',
    token: tuLogin.json.token,
    fetchImpl: mockFetch,
    body: {
      authorizeSah: true,
      signerRole: 'kepala_sekolah',
      entry: {
        year: 2026,
        typeCode: 'KET',
        nomorSurat: '421.3/001/KET.SMKITA/XI/2026',
        title: 'SURAT KETERANGAN',
        recipientName: 'Siswa Uji'
      },
      verificationPayload: {
        verificationId: 'SMKIT-2026-KET-TEST1111',
        manageToken: 'tok_secret_123456789012',
        signerRole: 'kepala_sekolah',
        signerName: 'Dian Nurhadian, S.Pd.'
      }
    }
  });
  assert.equal(deniedTuRes.status, 403);

  // 4. TU sahkan atas nama Tata Usaha -> 201 Created + nomor #001 + tersimpan ke Drive
  const allowedTuRes = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'record_outgoing',
    token: tuLogin.json.token,
    fetchImpl: mockFetch,
    body: {
      authorizeSah: true,
      signerRole: 'tata_usaha',
      docxBase64: Buffer.from('mock-docx-bytes-1').toString('base64'),
      entry: {
        year: 2026,
        typeCode: 'TU',
        nomorSurat: '421.3/001/TU.SMKITA/XI/2026',
        title: 'SURAT TATA USAHA',
        recipientName: 'Wali Kelas XII'
      },
      verificationPayload: {
        verificationId: 'SMKIT-2026-TU-TEST2222',
        manageToken: 'tok_secret_abcdef123456',
        signerRole: 'tata_usaha',
        signerName: 'Staf Tata Usaha (TU)'
      }
    }
  });
  assert.equal(allowedTuRes.status, 201);
  const allowedTuJson = await allowedTuRes.json();
  assert.equal(allowedTuJson.item.sequenceNumber, 1);
  assert.equal(allowedTuJson.item.driveSynced, true);
  assert.match(allowedTuJson.item.driveDocxUrl, /drive\.google\.com/);
  assert.equal(driveCalls.length, 1);
  assert.equal(driveCalls[0].parsed.action, 'archive_outgoing');

  // 5. Kepala Sekolah aktifkan izin delegasi pengesahan Kepala Sekolah untuk TU
  const permRes = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'update_permissions',
    token: kepsekLogin.json.token,
    body: {
      permissions: {
        allowTuSignPrincipal: true,
        allowTuSignKepsekDanTu: true,
        note: 'Kepsek sedang dinas luar, TU diizinkan mengesahkan QR a.n. Kepsek'
      }
    }
  });
  assert.equal(permRes.status, 200);
  const permJson = await permRes.json();
  assert.equal(permJson.permissions.allowTuSignPrincipal, true);
  assert.equal(permJson.permissions.updatedBy, 'kepsek');

  // 6. Sekarang TU boleh mengesahkan a.n. Kepala Sekolah, otomatis dapat nomor urut #002
  const delegatedTuRes = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'record_outgoing',
    token: tuLogin.json.token,
    fetchImpl: mockFetch,
    body: {
      authorizeSah: true,
      signerRole: 'kepala_sekolah',
      docxBase64: Buffer.from('mock-docx-bytes-2').toString('base64'),
      entry: {
        year: 2026,
        typeCode: 'KET',
        nomorSurat: '421.3/001/KET.SMKITA/XI/2026', // Sengaja kirim 001 untuk uji atomik klaim -> harus jadi 002!
        title: 'SURAT KETERANGAN AKTIF',
        recipientName: 'Ahmad Fauzi'
      },
      verificationPayload: {
        verificationId: 'SMKIT-2026-KET-DELEG333',
        manageToken: 'tok_secret_deleg123456',
        signerRole: 'kepala_sekolah',
        signerName: 'Dian Nurhadian, S.Pd.'
      }
    }
  });
  assert.equal(delegatedTuRes.status, 201);
  const delJson = await delegatedTuRes.json();
  assert.equal(delJson.item.sequenceNumber, 2);
  assert.equal(delJson.item.nomorSurat, '421.3/002/KET.SMKITA/XI/2026');
  assert.equal(delJson.item.delegatedByKepsek, true);

  // 7. Pastikan QR yang dibuat dari workspace langsung terbaca di endpoint publik /verify-letter
  const publicVerify = await handleVerifyRequest(
    new Request('https://copilot-dokumen-alhawari.netlify.app/verify-letter?v=SMKIT-2026-KET-DELEG333'),
    {},
    () => verifyStore
  );
  assert.equal(publicVerify.status, 200);
  const pubJson = await publicVerify.json();
  assert.equal(pubJson.verification.nomorSurat, '421.3/002/KET.SMKITA/XI/2026');
  assert.equal(pubJson.verification.delegatedByKepsek, true);
});

test('Kontrol Akun Admin & Sinkronisasi Gudang API Multi-Provider Tanpa Batas', async () => {
  const wsStore = createMockBlobStore();
  const verifyStore = createMockBlobStore();

  const adminLogin = await loginAs(wsStore, verifyStore, 'admin', TEST_ENV.INITIAL_ADMIN_PASSWORD);
  const kepsekLogin = await loginAs(wsStore, verifyStore, 'kepsek', TEST_ENV.INITIAL_KEPSEK_PASSWORD);
  const tuLogin = await loginAs(wsStore, verifyStore, 'tu', TEST_ENV.INITIAL_TU_PASSWORD);

  // Kepsek & TU dilarang mengelola akun pengguna (khusus Admin)
  const kepsekCreateTry = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'manage_users',
    token: kepsekLogin.json.token,
    body: { op: 'create', username: 'operator_x', role: 'tu', password: 'Password123' }
  });
  assert.equal(kepsekCreateTry.status, 403);

  // Admin membuat akun TU tambahan ('tu_operator2')
  const createUserRes = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'manage_users',
    token: adminLogin.json.token,
    body: {
      op: 'create',
      username: 'tu_operator2',
      displayName: 'Staf TU Operator 2',
      role: 'tu',
      password: 'Tu2@Password'
    }
  });
  assert.equal(createUserRes.status, 201);

  const tu2Login = await loginAs(wsStore, verifyStore, 'tu_operator2', 'Tu2@Password');
  assert.equal(tu2Login.status, 200);
  assert.equal(tu2Login.json.user.role, 'tu');

  // Admin menonaktifkan tu_operator2 -> token langsung ditolak
  const disableRes = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'manage_users',
    token: adminLogin.json.token,
    body: {
      op: 'update',
      targetUserId: tu2Login.json.user.id,
      active: false
    }
  });
  assert.equal(disableRes.status, 200);

  const rejectedSession = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'GET',
    action: 'session',
    token: tu2Login.json.token
  });
  assert.equal(rejectedSession.status, 401);

  // Admin menyimpan 6 API AI ke server -> TU tetap bisa membaca dan menambah API ke-7, ke-8 dst tanpa kehilangan API sebelumnya
  const sixApis = Array.from({ length: 6 }, (_, i) => ({
    id: `custom-api-${i + 1}`,
    label: `API Key Sekolah #${i + 1}`,
    provider: i % 2 === 0 ? 'gemini' : 'groq',
    model: i % 2 === 0 ? 'gemini-3.5-flash-lite' : 'openai/gpt-oss-120b',
    apiKey: `KEY_SECRET_VAL_${i + 1}`,
    baseUrl: 'https://example.api',
    isActive: i === 2,
    updatedAt: Date.now()
  }));

  const saveApiAdmin = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'save_state',
    token: adminLogin.json.token,
    body: { apiVault: sixApis, activeApiId: 'custom-api-3' }
  });
  assert.equal(saveApiAdmin.status, 200);

  // TU menambah 2 API tambahan lagi (merge non-destruktif!)
  const twoMoreApis = [
    {
      id: 'custom-api-7',
      label: 'API Cadangan TU #7',
      provider: 'deepseek',
      model: 'deepseek-chat',
      apiKey: 'KEY_TU_7',
      baseUrl: 'https://api.deepseek.com/v1',
      isActive: true,
      updatedAt: Date.now() + 100
    },
    {
      id: 'custom-api-8',
      label: 'API Cadangan TU #8',
      provider: 'openai',
      model: 'gpt-4o-mini',
      apiKey: 'KEY_TU_8',
      baseUrl: 'https://api.openai.com/v1',
      isActive: false,
      updatedAt: Date.now() + 100
    }
  ];

  const saveApiTu = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'save_state',
    token: tuLogin.json.token,
    body: { apiVault: twoMoreApis, activeApiId: 'custom-api-7' }
  });
  assert.equal(saveApiTu.status, 200);

  // Bootstrap di perangkat baru mengembalikan ke-8 API lengkap
  const bootRes = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'GET',
    action: 'bootstrap',
    token: kepsekLogin.json.token
  });
  assert.equal(bootRes.status, 200);
  const bootJson = await bootRes.json();
  assert.equal(bootJson.workspace.apiVault.items.length, 8);
  assert.equal(bootJson.workspace.apiVault.activeApiId, 'custom-api-7');
});

test('Helper RBAC Frontend: Mengunci opsi Kepsek untuk TU kecuali delegasi disetujui', () => {
  const tuUser = { id: 'u-tu', username: 'tu', role: 'tu', active: true };
  const kepsekUser = { id: 'u-kepsek', username: 'kepsek', role: 'kepsek', active: true };
  const adminUser = { id: 'u-admin', username: 'admin', role: 'admin', active: true };

  // Delegasi OFF
  const lockedOpts = getSelectableSignerRoles(tuUser, { allowTuSignPrincipal: false, allowTuSignKepsekDanTu: false });
  assert.equal(lockedOpts.find(o => o.id === 'kepala_sekolah').allowed, false);
  assert.equal(lockedOpts.find(o => o.id === 'kepsek_dan_tu').allowed, false);
  assert.equal(lockedOpts.find(o => o.id === 'tata_usaha').allowed, true);
  assert.equal(resolveDefaultAllowedSignerRole(tuUser, { allowTuSignPrincipal: false }, 'kepala_sekolah'), 'tata_usaha');
  assert.equal(canDeleteOutgoingItem(tuUser, { allowTuDeleteOutgoing: false }), false);
  assert.equal(canRevokeOutgoingItem(tuUser, { allowTuRevokePrincipalQr: false }, { signerRole: 'kepala_sekolah' }), false);
  assert.equal(canRevokeOutgoingItem(tuUser, { allowTuRevokePrincipalQr: false }, { signerRole: 'tata_usaha' }), true);

  // Delegasi ON oleh Kepsek/Admin
  const unlockedOpts = getSelectableSignerRoles(tuUser, { allowTuSignPrincipal: true, allowTuSignKepsekDanTu: true });
  assert.equal(unlockedOpts.find(o => o.id === 'kepala_sekolah').allowed, true);
  assert.equal(unlockedOpts.find(o => o.id === 'kepala_sekolah').delegated, true);
  assert.equal(resolveDefaultAllowedSignerRole(tuUser, { allowTuSignPrincipal: true }, 'kepala_sekolah'), 'kepala_sekolah');

  // Admin & Kepsek selalu bebas
  assert.equal(canDeleteOutgoingItem(adminUser, { allowTuDeleteOutgoing: false }), true);
  assert.equal(canDeleteOutgoingItem(kepsekUser, { allowTuDeleteOutgoing: false }), true);
  assert.equal(canManageApiVault(adminUser, { allowTuManageApiVault: false }), true);
  assert.equal(canManageApiVault(tuUser, { allowTuManageApiVault: false }), false);
});

test('Alur Antrean ACC Kepala Sekolah, Live SPA Polling, & Penerbitan QR + DOCX Server', async () => {
  const wsStore = createMockBlobStore();
  const verifyStore = createMockBlobStore();
  const gasCalls = [];
  const mockFetch = async (_url, init) => {
    const body = JSON.parse(init.body || '{}');
    gasCalls.push(body);
    return new Response(JSON.stringify({
      ok: true,
      driveDocxFileId: 'drv-docx-99',
      driveDocxUrl: 'https://drive.google.com/file/d/drv-docx-99/view'
    }), { status: 200, headers: { 'Content-Type': 'application/json' } });
  };

  const tuLogin = await loginAs(wsStore, verifyStore, 'tu', TEST_ENV.INITIAL_TU_PASSWORD);
  const kepsekLogin = await loginAs(wsStore, verifyStore, 'kepsek', TEST_ENV.INITIAL_KEPSEK_PASSWORD);

  // 1. Poll awal Kepsek mendapat rev = 1
  const bootKepsek = await callWorkspace({
    wsStore, verifyStore, method: 'GET', action: 'bootstrap', token: kepsekLogin.json.token
  });
  const bootKepsekJson = await bootKepsek.json();
  const rev0 = bootKepsekJson.syncRev;
  assert.ok(rev0 >= 1);

  // 2. TU mengajukan permintaan ACC Kepala Sekolah saat delegasi masih OFF
  const reqAccRes = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'record_outgoing',
    token: tuLogin.json.token,
    fetchImpl: mockFetch,
    body: {
      authorizeSah: false,
      requestAcc: true,
      signerRole: 'kepala_sekolah',
      requestNote: 'Mohon ACC surat keterangan siswa aktif',
      entry: {
        typeCode: 'KET',
        year: '2026',
        nomorSurat: '421.3/001/KET.SMKITA/XI/2026',
        title: 'SURAT KETERANGAN SISWA AKTIF',
        recipientName: 'Siti Aminah',
        perihal: 'Keterangan Siswa Aktif',
        tanggalSurat: '16 November 2026',
        documentSnapshot: {
          meta: { title: 'SURAT KETERANGAN SISWA AKTIF' },
          blocks: [
            { type: 'letterTitle', title: 'SURAT KETERANGAN', nomor: 'Nomor: 421.3/001/KET.SMKITA/XI/2026' },
            { type: 'paragraph', style: 'letterBody', runs: [{ text: 'Menerangkan siswa aktif.' }] },
            { type: 'letterSignature', city: 'Garut', date: '16 November 2026', title: 'Kepala Sekolah', name: 'Dian Nurhadian, S.Pd.', nip: '-' }
          ]
        }
      }
    }
  });
  assert.equal(reqAccRes.status, 201);
  const reqAccJson = await reqAccRes.json();
  assert.equal(reqAccJson.item.verificationStatus, 'pending_acc');
  assert.equal(reqAccJson.item.accStatus, 'pending');
  assert.equal(reqAccJson.item.qrDataUrl || '', '');
  const itemId = reqAccJson.item.id;
  const rev1 = reqAccJson.syncRev;
  assert.ok(rev1 > rev0);

  // 3. Poll update Kepsek sejak rev0 langsung menangkap antrean pending_acc
  const pollKepsek = await callWorkspace({
    wsStore, verifyStore, method: 'GET', action: 'poll_updates', query: { since: rev0 }, token: kepsekLogin.json.token
  });
  const pollKepsekJson = await pollKepsek.json();
  assert.equal(pollKepsekJson.changed, true);
  assert.equal(pollKepsekJson.pendingAccCount, 1);
  assert.equal(pollKepsekJson.workspace.outgoingLetters[0].verificationStatus, 'pending_acc');

  // 4. Poll ulang dengan rev1 (tidak ada perubahan) mengembalikan changed: false ringan
  const pollNoChange = await callWorkspace({
    wsStore, verifyStore, method: 'GET', action: 'poll_updates', query: { since: rev1 }, token: kepsekLogin.json.token
  });
  assert.equal((await pollNoChange.json()).changed, false);

  // 5. TU tidak boleh meng-ACC dirinya sendiri (403)
  const tuTryApprove = await callWorkspace({
    wsStore, verifyStore, method: 'POST', action: 'approve_acc', token: tuLogin.json.token, body: { outgoingId: itemId }
  });
  assert.equal(tuTryApprove.status, 403);

  // 6. Kepala Sekolah meng-ACC -> QR PNG dibuat, /verify-letter tersimpan, DOCX QR di-upload ke GAS Drive
  const approveRes = await callWorkspace({
    wsStore,
    verifyStore,
    method: 'POST',
    action: 'approve_acc',
    token: kepsekLogin.json.token,
    fetchImpl: mockFetch,
    body: {
      outgoingId: itemId,
      signerRole: 'kepala_sekolah',
      verificationOrigin: 'https://copilot-dokumen-alhawari.netlify.app'
    }
  });
  assert.equal(approveRes.status, 200);
  const approveJson = await approveRes.json();
  assert.equal(approveJson.item.verificationStatus, 'sah');
  assert.equal(approveJson.item.accStatus, 'approved');
  assert.equal(approveJson.item.isVerifiedSah, true);
  assert.ok(approveJson.item.qrDataUrl.startsWith('data:image/png;base64,'));
  assert.equal(approveJson.item.driveDocxFileId, 'drv-docx-99');
  assert.ok(gasCalls.some(c => c.action === 'archive_outgoing' && typeof c.docxBase64 === 'string' && c.docxBase64.length > 50));

  // Verifikasi QR publik dapat langsung discan HP via /verify-letter
  const verifyRes = await handleVerifyRequest(
    new Request(`https://copilot-dokumen-alhawari.netlify.app/verify-letter?v=${encodeURIComponent(approveJson.item.verificationId)}`),
    {},
    () => verifyStore
  );
  assert.equal(verifyRes.status, 200);
  assert.equal((await verifyRes.json()).verification.status, 'sah');

  // 7. TU poll_updates dengan rev1 langsung menerima status SAH tanpa refresh
  const tuLivePoll = await callWorkspace({
    wsStore, verifyStore, method: 'GET', action: 'poll_updates', query: { since: rev1 }, token: tuLogin.json.token
  });
  const tuLiveJson = await tuLivePoll.json();
  assert.equal(tuLiveJson.changed, true);
  assert.equal(tuLiveJson.pendingAccCount, 0);
  assert.equal(tuLiveJson.workspace.outgoingLetters[0].verificationStatus, 'sah');
  assert.ok(tuLiveJson.workspace.outgoingLetters[0].qrDataUrl.startsWith('data:image/png;base64,'));

  // 8. Uji Tolak ACC & Ajukan Ulang + Batalkan ACC pada surat kedua
  const secondRec = await callWorkspace({
    wsStore, verifyStore, method: 'POST', action: 'record_outgoing', token: tuLogin.json.token, fetchImpl: mockFetch,
    body: {
      authorizeSah: false,
      requestAcc: false,
      entry: { typeCode: 'KET', year: '2026', nomorSurat: '421.3/002/KET.SMKITA/XI/2026', title: 'SURAT KEDUA', recipientName: 'Budi', tanggalSurat: '16 November 2026' }
    }
  });
  const secondId = (await secondRec.json()).item.id;

  const reqSecondAcc = await callWorkspace({
    wsStore, verifyStore, method: 'POST', action: 'request_acc', token: tuLogin.json.token, fetchImpl: mockFetch,
    body: { outgoingId: secondId, signerRole: 'kepala_sekolah', requestNote: 'Cek ulang lampiran' }
  });
  assert.equal(reqSecondAcc.status, 200);
  assert.equal((await reqSecondAcc.json()).item.verificationStatus, 'pending_acc');

  const rejectRes = await callWorkspace({
    wsStore, verifyStore, method: 'POST', action: 'reject_acc', token: kepsekLogin.json.token, fetchImpl: mockFetch,
    body: { outgoingId: secondId, reason: 'Perbaiki ejaan nama penerima' }
  });
  assert.equal(rejectRes.status, 200);
  const rejectJson = await rejectRes.json();
  assert.equal(rejectJson.item.verificationStatus, 'rejected');
  assert.equal(rejectJson.item.rejectedReason, 'Perbaiki ejaan nama penerima');

  // Ajukan ulang lalu batalkan
  await callWorkspace({
    wsStore, verifyStore, method: 'POST', action: 'request_acc', token: tuLogin.json.token, fetchImpl: mockFetch,
    body: { outgoingId: secondId, signerRole: 'kepala_sekolah' }
  });
  const cancelRes = await callWorkspace({
    wsStore, verifyStore, method: 'POST', action: 'cancel_acc', token: tuLogin.json.token,
    body: { outgoingId: secondId }
  });
  assert.equal(cancelRes.status, 200);
  assert.equal((await cancelRes.json()).item.verificationStatus, 'unverified');
});

test('Helper Kunci Kanvas HP Kepala Sekolah & Filter Antrean ACC', () => {
  const kepsekUser = { id: 'u-k', username: 'kepsek', role: 'kepsek', active: true };
  const adminUser = { id: 'u-a', username: 'admin', role: 'admin', active: true };
  const tuUser = { id: 'u-t', username: 'tu', role: 'tu', active: true };

  // Kepsek di HP (viewport <= 860 atau mobile phone UA) wajib terkunci kanvasnya
  assert.equal(isCanvasLockedForUser(kepsekUser, { innerWidth: 390, userAgent: 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) Mobile' }), true);
  assert.equal(isCanvasLockedForUser(kepsekUser, { innerWidth: 412, userAgent: 'Mozilla/5.0 (Linux; Android 14) Mobile Safari/537.36' }), true);
  // Kepsek di Desktop / Laptop / PC bebas mengedit kanvas
  assert.equal(isCanvasLockedForUser(kepsekUser, { innerWidth: 1366, userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }), false);
  // Admin & TU tidak pernah terkunci di ukuran layar mana pun
  assert.equal(isCanvasLockedForUser(adminUser, { innerWidth: 390, userAgent: 'Mozilla/5.0 Mobile' }), false);
  assert.equal(isCanvasLockedForUser(tuUser, { innerWidth: 390, userAgent: 'Mozilla/5.0 Mobile' }), false);

  assert.equal(canApproveOrRejectAcc(kepsekUser), true);
  assert.equal(canApproveOrRejectAcc(adminUser), true);
  assert.equal(canApproveOrRejectAcc(tuUser), false);

  const pending = getPendingAccLetters([
    { id: '1', verificationStatus: 'pending_acc' },
    { id: '2', verificationStatus: 'sah' },
    { id: '3', verificationStatus: 'unverified', accStatus: 'pending' },
    { id: '4', verificationStatus: 'rejected' }
  ]);
  assert.deepEqual(pending.map(p => p.id), ['1', '3']);
});

test('Multi-Unit Yayasan (SMP & SMK) & Privasi Ruang Kerja + Chat AI per Akun', async () => {
  const wsStore = createMockBlobStore();
  const verifyStore = createMockBlobStore();
  const mockFetch = async () => new Response(JSON.stringify({ ok: true }), { status: 200 });

  // 1. Login Admin Yayasan, Kepsek SMP, TU SMP, dan TU SMK
  const adminLogin = await loginAs(wsStore, verifyStore, 'admin', TEST_ENV.INITIAL_ADMIN_PASSWORD);
  const tuSmkLogin = await loginAs(wsStore, verifyStore, 'tu', TEST_ENV.INITIAL_TU_PASSWORD);
  const tuSmpLogin = await loginAs(wsStore, verifyStore, 'tu_smp', 'TuSmpAlHawari#2026');
  const kepsekSmpLogin = await loginAs(wsStore, verifyStore, 'kepsek_smp', 'KepsekSmpAlHawari#2026');

  assert.equal(adminLogin.json.user.unitId, 'all');
  assert.equal(tuSmkLogin.json.user.unitId, 'smk');
  assert.equal(tuSmpLogin.status, 200);
  assert.equal(tuSmpLogin.json.user.unitId, 'smp');
  assert.equal(kepsekSmpLogin.status, 200);
  assert.equal(kepsekSmpLogin.json.user.unitId, 'smp');

  // 2. Bootstrap SMP mengembalikan profil resmi SMP IT Al-Hawari - Cibiuk (NPSN 20279898 & Yangyang Budiman, S.Pd.)
  const bootSmp = await callWorkspace({
    wsStore, verifyStore, method: 'GET', action: 'bootstrap', token: tuSmpLogin.json.token
  });
  const bootSmpJson = await bootSmp.json();
  assert.equal(bootSmpJson.unitId, 'smp');
  assert.equal(bootSmpJson.workspace.schoolProfile.npsn, '20279898');
  assert.equal(bootSmpJson.workspace.schoolProfile.principalName, 'Yangyang Budiman, S.Pd.');
  assert.ok(bootSmpJson.workspace.schoolProfile.numberingPattern.includes('SMPITA'));

  // 3. Staf SMP dilarang mengakses unit SMK (403 Forbidden) dan sebaliknya
  const crossUnitTry = await callWorkspace({
    wsStore, verifyStore, method: 'GET', action: 'bootstrap', query: { unit: 'smk' }, token: tuSmpLogin.json.token
  });
  assert.equal(crossUnitTry.status, 403);
  const crossUnitTry2 = await callWorkspace({
    wsStore, verifyStore, method: 'GET', action: 'bootstrap', query: { unit: 'smp' }, token: tuSmkLogin.json.token
  });
  assert.equal(crossUnitTry2.status, 403);

  // 4. Penomoran & Antrean ACC SMP terpisah total dari SMK (keduanya mulai urut #001)
  const smkOutRes = await callWorkspace({
    wsStore, verifyStore, method: 'POST', action: 'record_outgoing', token: tuSmkLogin.json.token, fetchImpl: mockFetch,
    body: {
      authorizeSah: false, requestAcc: true,
      entry: { year: '2026', typeCode: 'KET', nomorSurat: '421.3/001/KET.SMKITA/XI/2026', title: 'SURAT SMK 1', recipientName: 'Siswa SMK' }
    }
  });
  assert.equal((await smkOutRes.json()).item.sequenceNumber, 1);

  const smpOutRes = await callWorkspace({
    wsStore, verifyStore, method: 'POST', action: 'record_outgoing', token: tuSmpLogin.json.token, fetchImpl: mockFetch,
    body: {
      authorizeSah: false, requestAcc: true,
      entry: { year: '2026', typeCode: 'KET', nomorSurat: '421.2/001/KET.SMPITA/XI/2026', title: 'SURAT SMP 1', recipientName: 'Siswa SMP' }
    }
  });
  const smpOutJson = await smpOutRes.json();
  assert.equal(smpOutRes.status, 201);
  assert.equal(smpOutJson.item.sequenceNumber, 1);
  assert.equal(smpOutJson.item.unitId, 'smp');

  // Kepsek SMP menyetujui ACC -> verifikasi QR tercatat atas nama SMP IT AL-HAWARI - CIBIUK
  const approveSmpRes = await callWorkspace({
    wsStore, verifyStore, method: 'POST', action: 'approve_acc', token: kepsekSmpLogin.json.token, fetchImpl: mockFetch,
    body: { outgoingId: smpOutJson.item.id, signerRole: 'kepala_sekolah' }
  });
  const approveSmpJson = await approveSmpRes.json();
  assert.equal(approveSmpRes.status, 200);
  assert.equal(approveSmpJson.item.verificationStatus, 'sah');
  assert.equal(approveSmpJson.verification.schoolName, 'SMP IT AL-HAWARI - CIBIUK');

  // Admin Yayasan dapat beralih unit SMK vs SMP dengan data terpisah
  const adminBootSmp = await (await callWorkspace({
    wsStore, verifyStore, method: 'GET', action: 'bootstrap', query: { unit: 'smp' }, token: adminLogin.json.token
  })).json();
  assert.equal(adminBootSmp.unitId, 'smp');
  assert.equal(adminBootSmp.workspace.outgoingLetters.length, 1);
  assert.equal(adminBootSmp.workspace.outgoingLetters[0].title, 'SURAT SMP 1');

  // 5. Ruang Kerja & Chat AI bersifat PRIVAT per akun (TU SMP vs Kepsek SMP vs Admin tidak bisa saling intip)
  const tuPrivSave = await callWorkspace({
    wsStore, verifyStore, method: 'POST', action: 'save_private_workspace', token: tuSmpLogin.json.token,
    body: {
      document: { meta: { title: 'Draf Privat Yoga TU SMP' }, blocks: [{ type: 'paragraph', runs: [{ text: 'Rahasia draf TU SMP' }] }] },
      chatSessions: [{ id: 'chat_smp_tu_1', title: 'Percakapan Privat TU SMP', messages: [{ role: 'user', text: 'Susun surat rahasia SMP' }] }],
      activeChatId: 'chat_smp_tu_1'
    }
  });
  assert.equal(tuPrivSave.status, 200);

  // Akun Kepsek SMP & Admin memeriksa private workspace & bootstrap -> tidak ada chat/draf milik tu_smp
  const kepsekPrivGet = await (await callWorkspace({
    wsStore, verifyStore, method: 'GET', action: 'get_private_workspace', token: kepsekSmpLogin.json.token
  })).json();
  assert.equal(kepsekPrivGet.privateWorkspace, null);

  const adminBootCheckSmp = await (await callWorkspace({
    wsStore, verifyStore, method: 'GET', action: 'bootstrap', query: { unit: 'smp' }, token: adminLogin.json.token
  })).json();
  assert.equal(adminBootCheckSmp.privateWorkspace, null);

  // Poll lite=1 mengembalikan HANYA data ACC ringan (tanpa workspace private atau blob berat)
  const litePoll = await (await callWorkspace({
    wsStore, verifyStore, method: 'GET', action: 'poll_updates', query: { unit: 'smp', since: 0, lite: '1' }, token: kepsekSmpLogin.json.token
  })).json();
  assert.equal(litePoll.lite, true);
  assert.equal(litePoll.workspace.schoolProfile, undefined);
  assert.equal(litePoll.privateWorkspace, undefined);
  assert.equal(litePoll.workspace.outgoingLetters.length, 1);

  // TU SMP sendiri memanggil bootstrap -> ruang kerja & chat AI miliknya pulih utuh
  const tuOwnBoot = await (await callWorkspace({
    wsStore, verifyStore, method: 'GET', action: 'bootstrap', token: tuSmpLogin.json.token
  })).json();
  assert.equal(tuOwnBoot.privateWorkspace.activeChatId, 'chat_smp_tu_1');
  assert.equal(tuOwnBoot.privateWorkspace.chatSessions[0].title, 'Percakapan Privat TU SMP');
});
