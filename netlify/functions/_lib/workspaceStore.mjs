/**
 * netlify/functions/_lib/workspaceStore.mjs
 * Layanan Backend Terpusat untuk Autentikasi RBAC (Admin / Kepsek / TU),
 * Delegasi Pengesahan Legal Kepala Sekolah -> TU, Sinkronisasi Workspace Multi-Perangkat
 * (Surat Keluar, Gudang Surat, Profil/Kop, Gudang API tak terbatas), & Jembatan Google Drive.
 */

import crypto from 'node:crypto';
import { getStore } from '@netlify/blobs';

export const WORKSPACE_STORE_NAME = 'school-workspace';
export const VERIFICATION_STORE_NAME = 'letter-verifications';

export const BLOB_KEYS = Object.freeze({
  USERS: 'auth_users',
  PERMISSIONS: 'auth_permissions',
  DRIVE_CONFIG: 'drive_config',
  SERVER_SECRET: 'server_hmac_secret',
  SCHOOL_PROFILE: 'data_school_profile',
  KOP_GALLERY: 'data_kop_gallery',
  LETTER_VAULT: 'data_letter_vault',
  API_VAULT: 'data_api_vault',
  AGENDA_BOOK: 'data_agenda_book',
  OUTGOING_LIST: 'data_outgoing_list',
  DOC_DRAFTS: 'data_doc_drafts',
  SYNC_META: 'sync_meta'
});

export const DEFAULT_PERMISSIONS = Object.freeze({
  allowTuSignPrincipal: false,
  allowTuSignKepsekDanTu: false,
  allowTuRevokePrincipalQr: false,
  allowTuManageApiVault: true,
  allowTuDeleteOutgoing: false,
  updatedAt: '',
  updatedBy: 'system',
  updatedByRole: 'admin',
  note: 'Pengesahan atas nama Kepala Sekolah dikunci khusus Kepala Sekolah/Admin kecuali izin delegasi diaktifkan.'
});

const TOKEN_TTL_MS = 24 * 60 * 60 * 1000;

export function jsonResponse(payload, status = 200) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store, max-age=0'
    }
  });
}

export function defaultWorkspaceStoreFactory() {
  return getStore({ name: WORKSPACE_STORE_NAME, consistency: 'strong' });
}

export function defaultVerificationStoreFactory() {
  return getStore({ name: VERIFICATION_STORE_NAME, consistency: 'strong' });
}

function getEnvValue(env, key, fallback = '') {
  if (env && typeof env === 'object' && env[key]) return String(env[key]);
  if (typeof globalThis.Netlify !== 'undefined' && globalThis.Netlify?.env?.get) {
    const v = globalThis.Netlify.env.get(key);
    if (v) return String(v);
  }
  if (typeof process !== 'undefined' && process.env && process.env[key]) {
    return String(process.env[key]);
  }
  return fallback;
}

export function hashPasswordWithSalt(password, existingSalt = '') {
  const salt = existingSalt || crypto.randomBytes(16).toString('hex');
  const derived = crypto.scryptSync(String(password || ''), salt, 32).toString('hex');
  return { salt, hash: derived };
}

export function verifyPasswordHash(password, salt, expectedHashHex) {
  if (!password || !salt || !expectedHashHex) return false;
  try {
    const derivedBuf = crypto.scryptSync(String(password), String(salt), 32);
    const expectedBuf = Buffer.from(String(expectedHashHex), 'hex');
    if (derivedBuf.length !== expectedBuf.length) return false;
    return crypto.timingSafeEqual(derivedBuf, expectedBuf);
  } catch {
    return false;
  }
}

export function sha256Hex(value) {
  return crypto.createHash('sha256').update(String(value || '')).digest('hex');
}

function base64UrlEncode(input) {
  return Buffer.from(input).toString('base64url');
}

function base64UrlDecode(input) {
  return Buffer.from(String(input || ''), 'base64url').toString('utf8');
}

export async function resolveServerHmacSecret(wsStore, env) {
  const envSecret = getEnvValue(env, 'AUTH_SESSION_SECRET', '');
  if (envSecret && envSecret.length >= 16) return envSecret;

  const stored = await wsStore.get(BLOB_KEYS.SERVER_SECRET, { type: 'json' });
  if (stored?.secret) return stored.secret;

  const generated = crypto.randomBytes(32).toString('hex');
  await wsStore.setJSON(BLOB_KEYS.SERVER_SECRET, { secret: generated, createdAt: new Date().toISOString() }, { onlyIfNew: true });
  const reloaded = await wsStore.get(BLOB_KEYS.SERVER_SECRET, { type: 'json' });
  return reloaded?.secret || generated;
}

export function signSessionToken(user, secret) {
  const now = Date.now();
  const payload = {
    sub: user.id,
    username: user.username,
    role: user.role,
    iat: now,
    exp: now + TOKEN_TTL_MS
  };
  const encoded = base64UrlEncode(JSON.stringify(payload));
  const sig = crypto.createHmac('sha256', String(secret)).update(encoded).digest('base64url');
  return `${encoded}.${sig}`;
}

export function verifySignedToken(token, secret) {
  const raw = String(token || '').trim();
  const parts = raw.split('.');
  if (parts.length !== 2) return null;
  const [encoded, incomingSig] = parts;
  if (!encoded || !incomingSig) return null;

  const expectedSig = crypto.createHmac('sha256', String(secret)).update(encoded).digest('base64url');
  const sigBuf = Buffer.from(incomingSig);
  const expBuf = Buffer.from(expectedSig);
  if (sigBuf.length !== expBuf.length || !crypto.timingSafeEqual(sigBuf, expBuf)) {
    return null;
  }

  try {
    const payload = JSON.parse(base64UrlDecode(encoded));
    if (!payload?.sub || !payload?.exp || Date.now() > Number(payload.exp)) {
      return null;
    }
    return payload;
  } catch {
    return null;
  }
}

export function sanitizePublicUser(user) {
  if (!user) return null;
  return {
    id: user.id,
    username: user.username,
    displayName: user.displayName,
    role: user.role,
    active: user.active !== false,
    isDefaultAccount: Boolean(user.isDefaultAccount),
    createdAt: user.createdAt || '',
    updatedAt: user.updatedAt || '',
    lastLoginAt: user.lastLoginAt || ''
  };
}

export async function ensureInitialUsers(wsStore, env) {
  const existing = await wsStore.get(BLOB_KEYS.USERS, { type: 'json' });
  if (Array.isArray(existing) && existing.length > 0) {
    return existing;
  }

  const nowIso = new Date().toISOString();
  const adminPass = getEnvValue(env, 'INITIAL_ADMIN_PASSWORD', 'AdminAlHawari#2026');
  const kepsekPass = getEnvValue(env, 'INITIAL_KEPSEK_PASSWORD', 'KepsekAlHawari#2026');
  const tuPass = getEnvValue(env, 'INITIAL_TU_PASSWORD', 'TuAlHawari#2026');

  const specs = [
    { id: 'usr_admin_1', username: 'admin', displayName: 'Administrator Sistem', role: 'admin', password: adminPass },
    { id: 'usr_kepsek_1', username: 'kepsek', displayName: 'Dian Nurhadian, S.Pd. (Kepala Sekolah)', role: 'kepsek', password: kepsekPass },
    { id: 'usr_tu_1', username: 'tu', displayName: 'Staf Tata Usaha (TU)', role: 'tu', password: tuPass }
  ];

  const users = specs.map(spec => {
    const { salt, hash } = hashPasswordWithSalt(spec.password);
    return {
      id: spec.id,
      username: spec.username,
      displayName: spec.displayName,
      role: spec.role,
      active: true,
      isDefaultAccount: true,
      passwordSalt: salt,
      passwordHash: hash,
      createdAt: nowIso,
      updatedAt: nowIso,
      lastLoginAt: ''
    };
  });

  await wsStore.setJSON(BLOB_KEYS.USERS, users, { onlyIfNew: true });
  return (await wsStore.get(BLOB_KEYS.USERS, { type: 'json' })) || users;
}

export async function getEffectivePermissions(wsStore) {
  const stored = await wsStore.get(BLOB_KEYS.PERMISSIONS, { type: 'json' });
  return {
    ...DEFAULT_PERMISSIONS,
    ...(stored && typeof stored === 'object' ? stored : {})
  };
}

export function canUserUseSignerRole(user, permissions, requestedRole = 'kepala_sekolah') {
  if (!user || user.active === false) {
    return { allowed: false, reason: 'Akun pengguna tidak aktif.' };
  }
  const normalized = ['kepala_sekolah', 'tata_usaha', 'kepsek_dan_tu'].includes(requestedRole)
    ? requestedRole
    : 'kepala_sekolah';

  if (user.role === 'admin' || user.role === 'kepsek') {
    return { allowed: true, delegated: false };
  }

  if (user.role === 'tu') {
    if (normalized === 'tata_usaha') {
      return { allowed: true, delegated: false };
    }
    if (normalized === 'kepala_sekolah') {
      return permissions?.allowTuSignPrincipal === true
        ? { allowed: true, delegated: true }
        : {
            allowed: false,
            reason: 'Hak pengesahan atas nama Kepala Sekolah saat ini dikunci. Minta persetujuan Kepala Sekolah atau Admin pada menu Hak Akses.'
          };
    }
    if (normalized === 'kepsek_dan_tu') {
      return (permissions?.allowTuSignKepsekDanTu === true || permissions?.allowTuSignPrincipal === true)
        ? { allowed: true, delegated: true }
        : {
            allowed: false,
            reason: 'Hak pengesahan bersama Kepala Sekolah & TU belum diizinkan oleh Kepala Sekolah atau Admin.'
          };
    }
  }

  return { allowed: false, reason: 'Peran pengguna tidak dikenali.' };
}

export function canUserRevokeLetter(user, permissions, targetItem) {
  if (!user || user.active === false) return false;
  if (user.role === 'admin' || user.role === 'kepsek') return true;
  if (user.role === 'tu') {
    if (targetItem?.signerRole === 'tata_usaha') return true;
    return permissions?.allowTuRevokePrincipalQr === true;
  }
  return false;
}

export function mergeApiVaultData(existingVault = null, incomingItems = [], incomingActiveId = '') {
  const existingItems = Array.isArray(existingVault?.items) ? existingVault.items : [];
  const byId = new Map();

  for (const item of existingItems) {
    if (item?.id) byId.set(item.id, { ...item });
  }
  for (const item of (Array.isArray(incomingItems) ? incomingItems : [])) {
    if (!item?.id) continue;
    const prev = byId.get(item.id);
    if (!prev) {
      byId.set(item.id, { ...item });
    } else {
      const prevTs = Number(prev.updatedAt || prev.createdAt || 0);
      const nextTs = Number(item.updatedAt || item.createdAt || 0);
      if (nextTs >= prevTs) {
        byId.set(item.id, { ...prev, ...item });
      }
    }
  }

  const mergedItems = Array.from(byId.values());
  const resolvedActiveId = incomingActiveId
    || mergedItems.find(i => i.isActive)?.id
    || existingVault?.activeApiId
    || mergedItems[0]?.id
    || 'gemini-auto';

  return {
    activeApiId: resolvedActiveId,
    items: mergedItems.map(item => ({
      ...item,
      isActive: item.id === resolvedActiveId
    })),
    updatedAt: new Date().toISOString()
  };
}

export function mergeCollectionById(existingList = [], incomingList = [], deletedIds = []) {
  const deleteSet = new Set(Array.isArray(deletedIds) ? deletedIds : []);
  const byId = new Map();
  for (const item of (Array.isArray(existingList) ? existingList : [])) {
    if (item?.id && !deleteSet.has(item.id)) byId.set(item.id, item);
  }
  for (const item of (Array.isArray(incomingList) ? incomingList : [])) {
    if (item?.id && !deleteSet.has(item.id)) byId.set(item.id, item);
  }
  return Array.from(byId.values());
}

export async function claimNextOutgoingSequence(wsStore, outgoingList, year, preferredSeq = 0) {
  const targetYear = String(year || new Date().getFullYear());
  const sameYear = (Array.isArray(outgoingList) ? outgoingList : [])
    .filter(item => String(item.year || '') === targetYear);
  const currentMax = sameYear.reduce((max, item) => Math.max(max, Number(item.sequenceNumber) || 0), 0);
  let candidate = Math.max(currentMax + 1, Number(preferredSeq) || 0);

  for (let attempt = 0; attempt < 25; attempt += 1) {
    const claimKey = `seq_claim_${targetYear}_${String(candidate).padStart(4, '0')}`;
    const res = await wsStore.setJSON(claimKey, {
      year: targetYear,
      sequenceNumber: candidate,
      claimedAt: new Date().toISOString()
    }, { onlyIfNew: true });

    if (!res || res.modified !== false) {
      return candidate;
    }
    candidate += 1;
  }
  return candidate;
}

export async function getDriveConfigSafe(wsStore, env) {
  const stored = (await wsStore.get(BLOB_KEYS.DRIVE_CONFIG, { type: 'json' })) || {};
  const envUrl = getEnvValue(env, 'GAS_WEBHOOK_URL', '');
  const envSecret = getEnvValue(env, 'GAS_WEBHOOK_SECRET', '');
  const envRootId = getEnvValue(env, 'DRIVE_ROOT_FOLDER_ID', '19nFdoJ4VOEgFmUe_heAFxoFEt-9U9Crw');

  const webhookUrl = String(envUrl || stored.webhookUrl || '').trim();
  const webhookSecret = String(envSecret || stored.webhookSecret || '').trim();
  const rootFolderId = String(stored.rootFolderId || envRootId || '19nFdoJ4VOEgFmUe_heAFxoFEt-9U9Crw').trim();

  return {
    internal: {
      webhookUrl,
      webhookSecret,
      rootFolderId
    },
    publicStatus: {
      configured: Boolean(webhookUrl && webhookSecret),
      rootFolderId,
      rootFolderUrl: stored.rootFolderUrl || `https://drive.google.com/drive/folders/${rootFolderId}`,
      sheetUrl: stored.sheetUrl || '',
      lastSyncAt: stored.lastSyncAt || '',
      lastError: stored.lastError || '',
      accountEmail: 'tuancraf1@gmail.com'
    }
  };
}

export async function callGasBridge(wsStore, env, action, payload = {}, fetchImpl = globalThis.fetch) {
  const { internal } = await getDriveConfigSafe(wsStore, env);
  if (!internal.webhookUrl || !internal.webhookSecret) {
    return {
      ok: false,
      unconfigured: true,
      error: 'Webhook Apps Script Google Drive belum dikonfigurasi pada server.'
    };
  }
  if (typeof fetchImpl !== 'function') {
    return { ok: false, error: 'Fetch client tidak tersedia di runtime.' };
  }

  const controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
  const timer = controller ? setTimeout(() => controller.abort(), 18000) : null;

  try {
    const res = await fetchImpl(internal.webhookUrl, {
      method: 'POST',
      redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({
        secret: internal.webhookSecret,
        action,
        rootFolderId: internal.rootFolderId,
        ...payload
      }),
      signal: controller?.signal
    });
    const text = await res.text();
    let json = {};
    try { json = text ? JSON.parse(text) : {}; } catch { json = { raw: text.slice(0, 300) }; }

    if (!res.ok || json.ok === false) {
      const errMsg = json.error || `Google Drive Bridge error (HTTP ${res.status}).`;
      await saveDriveStatusUpdate(wsStore, { lastError: errMsg });
      return { ok: false, error: errMsg, data: json };
    }

    await saveDriveStatusUpdate(wsStore, {
      rootFolderId: json.rootFolderId || internal.rootFolderId,
      rootFolderUrl: json.rootFolderUrl || '',
      sheetUrl: json.sheetUrl || '',
      lastSyncAt: json.syncedAt || new Date().toISOString(),
      lastError: ''
    });

    return { ok: true, data: json };
  } catch (err) {
    const errMsg = err?.name === 'AbortError'
      ? 'Koneksi ke Google Apps Script Drive timeout (>18 detik).'
      : `Gagal menghubungi jembatan Google Drive (${err?.message || 'network error'}).`;
    await saveDriveStatusUpdate(wsStore, { lastError: errMsg });
    return { ok: false, error: errMsg };
  } finally {
    if (timer) clearTimeout(timer);
  }
}

async function saveDriveStatusUpdate(wsStore, patch = {}) {
  try {
    const cur = (await wsStore.get(BLOB_KEYS.DRIVE_CONFIG, { type: 'json' })) || {};
    const next = { ...cur };
    if (patch.rootFolderId) next.rootFolderId = patch.rootFolderId;
    if (patch.rootFolderUrl) next.rootFolderUrl = patch.rootFolderUrl;
    if (patch.sheetUrl) next.sheetUrl = patch.sheetUrl;
    if (patch.lastSyncAt !== undefined) next.lastSyncAt = patch.lastSyncAt;
    if (patch.lastError !== undefined) next.lastError = patch.lastError;
    await wsStore.setJSON(BLOB_KEYS.DRIVE_CONFIG, next);
  } catch {}
}

export async function getSyncMeta(wsStore) {
  const meta = await wsStore.get(BLOB_KEYS.SYNC_META, { type: 'json' });
  if (meta && typeof meta === 'object') {
    return {
      rev: Number(meta.rev) || 1,
      updatedAt: meta.updatedAt || '',
      lastEvent: meta.lastEvent || 'init',
      pendingAccCount: Number(meta.pendingAccCount) || 0
    };
  }
  return { rev: 1, updatedAt: '', lastEvent: 'init', pendingAccCount: 0 };
}

export async function bumpSyncMeta(wsStore, lastEvent = 'update', outgoingList = null) {
  const prev = await getSyncMeta(wsStore);
  const nextRev = Math.max(Date.now(), prev.rev + 1);
  let pendingAccCount = prev.pendingAccCount;
  if (Array.isArray(outgoingList)) {
    pendingAccCount = outgoingList.filter(i => i?.verificationStatus === 'pending_acc' || i?.accStatus === 'pending').length;
  }
  const meta = {
    rev: nextRev,
    updatedAt: new Date().toISOString(),
    lastEvent: String(lastEvent || 'update'),
    pendingAccCount
  };
  await wsStore.setJSON(BLOB_KEYS.SYNC_META, meta);
  return meta;
}

export async function authenticateRequest(req, wsStore, env) {
  const authHeader = req.headers.get('authorization') || '';
  const bearer = authHeader.toLowerCase().startsWith('bearer ')
    ? authHeader.slice(7).trim()
    : '';
  if (!bearer) return null;

  const secret = await resolveServerHmacSecret(wsStore, env);
  const payload = verifySignedToken(bearer, secret);
  if (!payload) return null;

  const users = await ensureInitialUsers(wsStore, env);
  const user = users.find(u => u.id === payload.sub);
  if (!user || user.active === false) return null;

  return user;
}
