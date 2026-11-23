/**
 * js/sync/workspaceClient.js - Klien API Server Netlify & Bridge Google Drive
 * Menangani autentikasi token sesi, sinkronisasi workspace multi-perangkat,
 * pengesahan QR, kontrol delegasi, dan unggah arsip .docx ke Google Drive.
 */

const TOKEN_STORAGE_KEY = 'ai_docbuilder_auth_token';
const ACTIVE_UNIT_STORAGE_KEY = 'ai_docbuilder_active_unit';

let unauthorizedCallback = null;
let inMemoryActiveUnitId = '';

export function onWorkspaceUnauthorized(cb) {
  unauthorizedCallback = typeof cb === 'function' ? cb : null;
}

export function getActiveWorkspaceUnitId() {
  if (inMemoryActiveUnitId) return inMemoryActiveUnitId;
  if (typeof localStorage === 'undefined') return 'smk';
  try {
    const val = String(localStorage.getItem(ACTIVE_UNIT_STORAGE_KEY) || '').trim().toLowerCase();
    return val === 'smp' ? 'smp' : 'smk';
  } catch {
    return 'smk';
  }
}

export function setActiveWorkspaceUnitId(unitId = 'smk') {
  const clean = String(unitId || '').trim().toLowerCase() === 'smp' ? 'smp' : 'smk';
  inMemoryActiveUnitId = clean;
  if (typeof localStorage === 'undefined') return;
  try {
    localStorage.setItem(ACTIVE_UNIT_STORAGE_KEY, clean);
  } catch {}
}

export function clearActiveWorkspaceUnitId() {
  inMemoryActiveUnitId = '';
  if (typeof localStorage === 'undefined') return;
  try {
    localStorage.removeItem(ACTIVE_UNIT_STORAGE_KEY);
  } catch {}
}

export function getStoredAuthToken() {
  if (typeof localStorage === 'undefined') return '';
  try {
    return localStorage.getItem(TOKEN_STORAGE_KEY) || '';
  } catch {
    return '';
  }
}

export function setStoredAuthToken(token) {
  if (typeof localStorage === 'undefined') return;
  try {
    if (token) {
      localStorage.setItem(TOKEN_STORAGE_KEY, token);
    } else {
      localStorage.removeItem(TOKEN_STORAGE_KEY);
    }
  } catch {}
}

export function clearStoredAuthToken() {
  setStoredAuthToken('');
}

export async function blobToBase64(blob) {
  if (!blob) return '';
  if (typeof blob.arrayBuffer === 'function') {
    const buf = await blob.arrayBuffer();
    const bytes = new Uint8Array(buf);
    let binary = '';
    const CHUNK = 0x8000;
    for (let i = 0; i < bytes.length; i += CHUNK) {
      binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
    }
    return btoa(binary);
  }
  return '';
}

async function requestWorkspace(action, { method = 'GET', query = null, body = null, requireAuth = true, timeoutMs = 20000, unitId = '' } = {}) {
  if (typeof fetch !== 'function') {
    return { ok: false, status: 0, error: 'Fetch tidak tersedia.' };
  }
  const token = getStoredAuthToken();
  if (requireAuth && !token) {
    return { ok: false, status: 401, unauthorized: true, error: 'Belum masuk akun.' };
  }

  const controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
  const timer = controller ? setTimeout(() => controller.abort(), timeoutMs) : null;

  try {
    const headers = {};
    if (token) headers['Authorization'] = `Bearer ${token}`;
    if (body !== null) headers['Content-Type'] = 'application/json';

    const params = new URLSearchParams({ action });
    if (requireAuth && unitId !== false) {
      const effectiveUnit = String(unitId || getActiveWorkspaceUnitId() || '').trim().toLowerCase();
      if (effectiveUnit === 'smp' || effectiveUnit === 'smk') {
        params.set('unit', effectiveUnit);
      }
    }
    if (query && typeof query === 'object') {
      for (const [k, v] of Object.entries(query)) {
        if (v !== undefined && v !== null) params.set(k, String(v));
      }
    }

    const res = await fetch(`/api/workspace?${params.toString()}`, {
      method,
      headers,
      cache: 'no-store',
      body: body !== null ? JSON.stringify(body) : undefined,
      signal: controller?.signal
    });

    const text = await res.text();
    let json = {};
    try { json = text ? JSON.parse(text) : {}; } catch { json = {}; }

    if (res.status === 401 && requireAuth) {
      clearStoredAuthToken();
      if (unauthorizedCallback) unauthorizedCallback(json.error || 'Sesi kedaluwarsa.');
    }

    if (!res.ok) {
      return {
        ok: false,
        status: res.status,
        unauthorized: res.status === 401,
        forbidden: res.status === 403,
        error: json.error || `Gagal terhubung ke server (HTTP ${res.status}).`,
        data: json
      };
    }

    return { ok: true, status: res.status, ...json };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      offline: true,
      error: err?.name === 'AbortError'
        ? 'Permintaan ke server timeout.'
        : `Tidak dapat menjangkau server (${err?.message || 'offline'}).`
    };
  } finally {
    if (timer) clearTimeout(timer);
  }
}

export function checkWorkspaceHealth() {
  return requestWorkspace('health', { method: 'GET', requireAuth: false, timeoutMs: 6000 });
}

export function loginWorkspace({ username, password }) {
  return requestWorkspace('login', {
    method: 'POST',
    requireAuth: false,
    body: { username, password }
  });
}

export function fetchSessionInfo() {
  return requestWorkspace('session', { method: 'GET' });
}

export function fetchWorkspaceBootstrap(unitId = '') {
  return requestWorkspace('bootstrap', { method: 'GET', timeoutMs: 25000, ...(unitId ? { unitId } : {}) });
}

export function fetchPrivateWorkspaceCloud(unitId = '') {
  return requestWorkspace('get_private_workspace', { method: 'GET', timeoutMs: 20000, ...(unitId ? { unitId } : {}) });
}

export function savePrivateWorkspaceCloud(privateWorkspace, unitId = '') {
  return requestWorkspace('save_private_workspace', {
    method: 'POST',
    body: { privateWorkspace },
    timeoutMs: 25000,
    ...(unitId ? { unitId } : {})
  });
}

export function saveWorkspaceStatePatch(patch) {
  return requestWorkspace('save_state', { method: 'POST', body: patch, timeoutMs: 25000 });
}

export function migrateLocalStateCloud(localData) {
  return requestWorkspace('migrate_local', { method: 'POST', body: localData, timeoutMs: 30000 });
}

export function recordOutgoingCloud(payload, unitId = '') {
  return requestWorkspace('record_outgoing', {
    method: 'POST',
    body: payload,
    timeoutMs: 30000,
    ...(unitId ? { unitId } : {})
  });
}

export function authorizeOutgoingCloud(payload, unitId = '') {
  return requestWorkspace('authorize_outgoing', {
    method: 'POST',
    body: payload,
    timeoutMs: 25000,
    ...(unitId ? { unitId } : {})
  });
}

export function revokeOutgoingCloud(payload) {
  return requestWorkspace('revoke_outgoing', { method: 'POST', body: payload, timeoutMs: 20000 });
}

export function deleteOutgoingCloud(outgoingId) {
  return requestWorkspace('delete_outgoing', { method: 'POST', body: { outgoingId } });
}

export function updatePermissionsCloud(permissions) {
  return requestWorkspace('update_permissions', { method: 'POST', body: { permissions } });
}

export function manageUsersCloud(payload) {
  return requestWorkspace('manage_users', { method: 'POST', body: payload });
}

export function changeOwnPasswordCloud({ oldPassword, newPassword }) {
  return requestWorkspace('change_password', { method: 'POST', body: { oldPassword, newPassword } });
}

export function configureDriveCloud(payload) {
  return requestWorkspace('configure_drive', { method: 'POST', body: payload, timeoutMs: 25000 });
}

export function syncDriveCloud(payload = {}) {
  return requestWorkspace('sync_drive', { method: 'POST', body: payload, timeoutMs: 35000 });
}

export function pollWorkspaceUpdates(sinceRev = 0, { lite = true, unitId = '' } = {}) {
  return requestWorkspace('poll_updates', {
    method: 'GET',
    query: {
      since: Number(sinceRev) || 0,
      ...(lite ? { lite: '1' } : {})
    },
    ...(unitId ? { unitId } : {}),
    timeoutMs: 12000
  });
}

export function requestAccCloud(payload, unitId = '') {
  return requestWorkspace('request_acc', {
    method: 'POST',
    body: payload,
    timeoutMs: 25000,
    ...(unitId ? { unitId } : {})
  });
}

export function approveAccCloud(payload) {
  return requestWorkspace('approve_acc', { method: 'POST', body: payload, timeoutMs: 35000 });
}

export function rejectAccCloud(payload) {
  return requestWorkspace('reject_acc', { method: 'POST', body: payload, timeoutMs: 20000 });
}

export function cancelAccCloud(outgoingId) {
  return requestWorkspace('cancel_acc', { method: 'POST', body: { outgoingId }, timeoutMs: 15000 });
}

export function recordIncomingCloud(payload) {
  return requestWorkspace('record_incoming', { method: 'POST', body: payload, timeoutMs: 45000 });
}

export function fetchIncomingFileCloud(incomingId) {
  return requestWorkspace('get_incoming_file', { method: 'GET', query: { id: incomingId }, timeoutMs: 25000 });
}

export function syncIncomingDriveCloud(incomingId, fileBase64 = '') {
  return requestWorkspace('sync_incoming_drive', {
    method: 'POST',
    body: { incomingId, ...(fileBase64 ? { fileBase64 } : {}) },
    timeoutMs: 40000
  });
}

export function deleteIncomingCloud(incomingId) {
  return requestWorkspace('delete_incoming', { method: 'POST', body: { incomingId }, timeoutMs: 20000 });
}
