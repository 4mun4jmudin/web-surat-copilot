/**
 * netlify/functions/workspace.mjs
 * API Terpusat Multi-Pengguna (Admin / Kepsek / TU) untuk Sinkronisasi Cloud,
 * Kontrol Delegasi Hak Pengesahan Legal, Gudang API AI Tanpa Batas, & Bridge Google Drive.
 */

import crypto from 'node:crypto';
import {
  BLOB_KEYS,
  authenticateRequest,
  bumpSyncMeta,
  callGasBridge,
  canUserRevokeLetter,
  canUserUseSignerRole,
  claimNextOutgoingSequence,
  defaultVerificationStoreFactory,
  defaultWorkspaceStoreFactory,
  ensureInitialUsers,
  getDriveConfigSafe,
  getEffectivePermissions,
  getSyncMeta,
  hashPasswordWithSalt,
  jsonResponse,
  mergeApiVaultData,
  mergeCollectionById,
  resolveServerHmacSecret,
  sanitizePublicUser,
  sha256Hex,
  signSessionToken,
  verifyPasswordHash
} from './_lib/workspaceStore.mjs';
import {
  applySignatureModeToDocModel,
  buildSignerSelectionInfo,
  buildVerificationId,
  buildVerificationPayload,
  buildVerificationUrl,
  generateQrPngDataUrl
} from '../../js/letters/verification.js';
import { renderDocx } from '../../js/docx/render.js';

export const config = {
  path: ['/api/workspace', '/api/workspace/:action']
};

function sortOutgoingList(list = []) {
  return [...list].sort((a, b) => {
    const yDiff = Number(b.year || 0) - Number(a.year || 0);
    if (yDiff !== 0) return yDiff;
    return (Number(b.sequenceNumber) || 0) - (Number(a.sequenceNumber) || 0);
  });
}

function syncNomorSequence(nomorSurat = '', claimedSeq = 1) {
  const padded = String(Math.max(1, Number(claimedSeq) || 1)).padStart(3, '0');
  const raw = String(nomorSurat || '').trim();
  if (!raw) return padded;
  // Ganti segmen 2-5 digit di antara garis miring (mis. 421.3/001/KET...)
  if (/\/\s*\d{2,5}\s*\//.test(raw)) {
    return raw.replace(/\/\s*\d{2,5}\s*\//, `/${padded}/`);
  }
  return raw;
}

function buildPublicVerificationRecord(payload = {}, actorUser = null, delegated = false) {
  const nowIso = new Date().toISOString();
  return {
    verificationId: String(payload.verificationId || '').trim(),
    status: 'sah',
    sequenceNumber: Number(payload.sequenceNumber) || 0,
    year: String(payload.year || ''),
    typeCode: String(payload.typeCode || 'DINAS').toUpperCase().slice(0, 20),
    typeLabel: String(payload.typeLabel || 'Surat Dinas Sekolah').slice(0, 120),
    nomorSurat: String(payload.nomorSurat || '-').slice(0, 120),
    title: String(payload.title || 'SURAT DINAS').slice(0, 180),
    perihal: String(payload.perihal || payload.title || '-').slice(0, 220),
    recipientName: String(payload.recipientName || '-').slice(0, 180),
    tanggalSurat: String(payload.tanggalSurat || '-').slice(0, 80),
    schoolName: String(payload.schoolName || 'SMK IT AL-HAWARI').slice(0, 160),
    schoolNpsn: String(payload.schoolNpsn || '').slice(0, 40),
    schoolAddress: String(payload.schoolAddress || '').slice(0, 240),
    signerRole: String(payload.signerRole || 'kepala_sekolah').slice(0, 40),
    signerTitle: String(payload.signerTitle || 'Kepala Sekolah').slice(0, 120),
    signerName: String(payload.signerName || 'Kepala Sekolah').slice(0, 140),
    signerNip: String(payload.signerNip || '-').slice(0, 60),
    verificationStatement: String(payload.verificationStatement || '').slice(0, 400),
    details: Array.isArray(payload.details)
      ? payload.details.slice(0, 14).map(d => ({
          label: String(d?.label || '').slice(0, 80),
          value: String(d?.value || '').slice(0, 200)
        }))
      : [],
    summary: String(payload.summary || '').slice(0, 320),
    issuedAt: nowIso,
    updatedAt: nowIso,
    revokedAt: null,
    revokedReason: '',
    authorizedByUser: actorUser?.username || '',
    authorizedByRole: actorUser?.role || '',
    delegatedByKepsek: Boolean(delegated),
    manageHash: sha256Hex(payload.manageToken || '')
  };
}

function stripVerificationSecrets(record) {
  if (!record || typeof record !== 'object') return null;
  const { manageHash, ...rest } = record;
  return rest;
}

async function handleLogin(body, wsStore, env) {
  const username = String(body?.username || '').trim().toLowerCase();
  const password = String(body?.password || '');
  if (!username || !password) {
    return jsonResponse({ ok: false, error: 'Username dan password/PIN wajib diisi.' }, 400);
  }

  const users = await ensureInitialUsers(wsStore, env);
  const idx = users.findIndex(u => String(u.username || '').toLowerCase() === username);
  const user = idx >= 0 ? users[idx] : null;
  if (!user || !verifyPasswordHash(password, user.passwordSalt, user.passwordHash)) {
    return jsonResponse({ ok: false, error: 'Username atau password salah.' }, 401);
  }
  if (user.active === false) {
    return jsonResponse({ ok: false, error: 'Akun ini dinonaktifkan oleh Administrator.' }, 403);
  }

  users[idx] = { ...user, lastLoginAt: new Date().toISOString() };
  await wsStore.setJSON(BLOB_KEYS.USERS, users);

  const secret = await resolveServerHmacSecret(wsStore, env);
  const token = signSessionToken(users[idx], secret);
  const permissions = await getEffectivePermissions(wsStore);
  const { publicStatus } = await getDriveConfigSafe(wsStore, env);

  return jsonResponse({
    ok: true,
    token,
    user: sanitizePublicUser(users[idx]),
    permissions,
    driveStatus: publicStatus
  }, 200);
}

async function handleBootstrap(user, wsStore, env) {
  const [
    schoolProfile,
    kopGallery,
    letterVault,
    apiVault,
    agendaBook,
    outgoingList,
    permissions,
    users,
    driveCfg,
    syncMeta
  ] = await Promise.all([
    wsStore.get(BLOB_KEYS.SCHOOL_PROFILE, { type: 'json' }),
    wsStore.get(BLOB_KEYS.KOP_GALLERY, { type: 'json' }),
    wsStore.get(BLOB_KEYS.LETTER_VAULT, { type: 'json' }),
    wsStore.get(BLOB_KEYS.API_VAULT, { type: 'json' }),
    wsStore.get(BLOB_KEYS.AGENDA_BOOK, { type: 'json' }),
    wsStore.get(BLOB_KEYS.OUTGOING_LIST, { type: 'json' }),
    getEffectivePermissions(wsStore),
    user.role === 'admin' ? ensureInitialUsers(wsStore, env) : Promise.resolve(null),
    getDriveConfigSafe(wsStore, env),
    getSyncMeta(wsStore)
  ]);

  const activeUsersSummary = user.role === 'admin'
    ? (users || []).map(sanitizePublicUser)
    : undefined;

  const sortedOutgoing = Array.isArray(outgoingList) ? sortOutgoingList(outgoingList) : null;
  const pendingAccCount = Array.isArray(sortedOutgoing)
    ? sortedOutgoing.filter(i => i?.verificationStatus === 'pending_acc' || i?.accStatus === 'pending').length
    : syncMeta.pendingAccCount;

  return jsonResponse({
    ok: true,
    user: sanitizePublicUser(user),
    permissions,
    driveStatus: driveCfg.publicStatus,
    syncRev: syncMeta.rev,
    pendingAccCount,
    users: activeUsersSummary,
    workspace: {
      schoolProfile: schoolProfile || null,
      kopGallery: Array.isArray(kopGallery) ? kopGallery : null,
      letterVault: Array.isArray(letterVault) ? letterVault : null,
      apiVault: apiVault?.items ? apiVault : null,
      agendaBook: Array.isArray(agendaBook) ? agendaBook : null,
      outgoingLetters: sortedOutgoing
    },
    syncedAt: new Date().toISOString()
  }, 200);
}

async function handleSaveState(user, body, wsStore, env, fetchImpl) {
  const permissions = await getEffectivePermissions(wsStore);
  const savedKeys = [];

  if (body.schoolProfile && typeof body.schoolProfile === 'object') {
    await wsStore.setJSON(BLOB_KEYS.SCHOOL_PROFILE, {
      ...body.schoolProfile,
      updatedAt: new Date().toISOString(),
      updatedBy: user.username
    });
    savedKeys.push('schoolProfile');
  }

  if (Array.isArray(body.kopGallery)) {
    await wsStore.setJSON(BLOB_KEYS.KOP_GALLERY, body.kopGallery);
    savedKeys.push('kopGallery');

    if (body.syncDriveItem && body.syncDriveItem.type === 'kop') {
      await callGasBridge(wsStore, env, 'archive_kop', {
        item: body.syncDriveItem.item,
        actorUsername: user.username
      }, fetchImpl);
    }
  }

  if (Array.isArray(body.letterVault)) {
    const existing = (await wsStore.get(BLOB_KEYS.LETTER_VAULT, { type: 'json' })) || [];
    const merged = body.replaceLetterVault
      ? body.letterVault
      : mergeCollectionById(existing, body.letterVault, body.deletedIds || []);
    await wsStore.setJSON(BLOB_KEYS.LETTER_VAULT, merged);
    savedKeys.push('letterVault');

    if (body.syncDriveItem && body.syncDriveItem.type === 'vault' && body.syncDriveItem.item) {
      const gasRes = await callGasBridge(wsStore, env, 'archive_vault', {
        item: body.syncDriveItem.item,
        docxBase64: body.syncDriveItem.docxBase64 || '',
        actorUsername: user.username
      }, fetchImpl);
      if (gasRes.ok) {
        const updatedVault = merged.map(v => (
          v.id === body.syncDriveItem.item.id
            ? { ...v, driveFileId: gasRes.data.driveFileId, driveUrl: gasRes.data.driveUrl, driveSynced: true }
            : v
        ));
        await wsStore.setJSON(BLOB_KEYS.LETTER_VAULT, updatedVault);
      }
    }
  }

  if (Array.isArray(body.apiVault)) {
    if (user.role === 'tu' && permissions.allowTuManageApiVault === false) {
      return jsonResponse({ ok: false, error: 'Akses pengelolaan Gudang API oleh TU dinonaktifkan oleh Admin.' }, 403);
    }
    const existingApi = await wsStore.get(BLOB_KEYS.API_VAULT, { type: 'json' });
    let nextVault;
    if (body.replaceApiVault) {
      const nextActiveId = body.activeApiId || body.apiVault.find(i => i.isActive)?.id || body.apiVault[0]?.id || 'gemini-auto';
      nextVault = {
        activeApiId: nextActiveId,
        items: body.apiVault.map(item => ({ ...item, isActive: item.id === nextActiveId })),
        updatedAt: new Date().toISOString()
      };
    } else {
      nextVault = mergeApiVaultData(existingApi, body.apiVault, body.activeApiId || '');
    }
    await wsStore.setJSON(BLOB_KEYS.API_VAULT, nextVault);
    savedKeys.push('apiVault');
  }

  if (Array.isArray(body.agendaBook)) {
    await wsStore.setJSON(BLOB_KEYS.AGENDA_BOOK, body.agendaBook);
    savedKeys.push('agendaBook');
  }

  let syncMeta = null;
  if (savedKeys.length > 0) {
    syncMeta = await bumpSyncMeta(wsStore, `save:${savedKeys.join(',')}`);
  }

  return jsonResponse({
    ok: true,
    savedKeys,
    syncRev: syncMeta?.rev,
    syncedAt: new Date().toISOString()
  }, 200);
}

async function handleMigrateLocal(user, body, wsStore) {
  const existingOutgoing = (await wsStore.get(BLOB_KEYS.OUTGOING_LIST, { type: 'json' })) || [];
  const incomingOutgoing = Array.isArray(body.outgoingLetters) ? body.outgoingLetters : [];
  const mergedOutgoing = sortOutgoingList(mergeCollectionById(existingOutgoing, incomingOutgoing));
  if (incomingOutgoing.length > 0) {
    await wsStore.setJSON(BLOB_KEYS.OUTGOING_LIST, mergedOutgoing);
  }

  const existingVault = (await wsStore.get(BLOB_KEYS.LETTER_VAULT, { type: 'json' })) || [];
  const incomingVault = Array.isArray(body.letterVault) ? body.letterVault : [];
  const mergedVault = mergeCollectionById(existingVault, incomingVault);
  if (incomingVault.length > 0) {
    await wsStore.setJSON(BLOB_KEYS.LETTER_VAULT, mergedVault);
  }

  const existingKop = (await wsStore.get(BLOB_KEYS.KOP_GALLERY, { type: 'json' })) || [];
  const incomingKop = Array.isArray(body.kopGallery) ? body.kopGallery : [];
  const mergedKop = mergeCollectionById(existingKop, incomingKop);
  if (incomingKop.length > 0) {
    await wsStore.setJSON(BLOB_KEYS.KOP_GALLERY, mergedKop);
  }

  const existingApi = await wsStore.get(BLOB_KEYS.API_VAULT, { type: 'json' });
  const incomingApi = Array.isArray(body.apiVault) ? body.apiVault : [];
  const mergedApi = incomingApi.length > 0
    ? mergeApiVaultData(existingApi, incomingApi, body.activeApiId || '')
    : existingApi;
  if (incomingApi.length > 0) {
    await wsStore.setJSON(BLOB_KEYS.API_VAULT, mergedApi);
  }

  const existingProfile = await wsStore.get(BLOB_KEYS.SCHOOL_PROFILE, { type: 'json' });
  if (!existingProfile && body.schoolProfile) {
    await wsStore.setJSON(BLOB_KEYS.SCHOOL_PROFILE, {
      ...body.schoolProfile,
      migratedBy: user.username,
      updatedAt: new Date().toISOString()
    });
  }

  const syncMeta = await bumpSyncMeta(wsStore, 'migrate_local', mergedOutgoing);

  return jsonResponse({
    ok: true,
    syncRev: syncMeta.rev,
    counts: {
      outgoingLetters: mergedOutgoing.length,
      letterVault: mergedVault.length,
      kopGallery: mergedKop.length,
      apiVault: mergedApi?.items?.length || 0
    }
  }, 200);
}

async function handleRecordOutgoing(user, body, wsStore, verifyStore, env, fetchImpl) {
  const permissions = await getEffectivePermissions(wsStore);
  const authorizeSah = Boolean(body.authorizeSah);
  const requestAcc = Boolean(body.requestAcc) && !authorizeSah;
  const requestedSignerRole = String(body.signerRole || 'kepala_sekolah').trim();
  const requestNote = String(body.requestNote || '').trim().slice(0, 240);

  let roleCheck = { allowed: true, delegated: false };
  if (authorizeSah) {
    roleCheck = canUserUseSignerRole(user, permissions, requestedSignerRole);
    if (!roleCheck.allowed) {
      return jsonResponse({ ok: false, error: roleCheck.reason }, 403);
    }
  }

  const entry = body.entry || {};
  const year = String(entry.year || new Date().getFullYear());
  const existingList = (await wsStore.get(BLOB_KEYS.OUTGOING_LIST, { type: 'json' })) || [];
  const claimedSeq = await claimNextOutgoingSequence(wsStore, existingList, year, Number(entry.sequenceNumber) || 0);
  const finalNomor = syncNomorSequence(entry.nomorSurat, claimedSeq);

  const snapshot = entry.documentSnapshot ? JSON.parse(JSON.stringify(entry.documentSnapshot)) : null;
  if (snapshot && Array.isArray(snapshot.blocks)) {
    const titleBlk = snapshot.blocks.find(b => b?.type === 'letterTitle');
    if (titleBlk) titleBlk.nomor = `Nomor: ${finalNomor}`;
  }

  const nowIso = new Date().toISOString();
  let cloudSynced = false;
  let verificationRecord = null;
  const payload = body.verificationPayload ? { ...body.verificationPayload } : null;

  if (authorizeSah && payload?.verificationId) {
    payload.sequenceNumber = claimedSeq;
    payload.nomorSurat = finalNomor;
    if (Array.isArray(payload.details)) {
      payload.details = payload.details.map(d => (
        d?.label === 'Nomor Surat Resmi' ? { ...d, value: finalNomor } : d
      ));
    }
    const pubRec = buildPublicVerificationRecord(payload, user, roleCheck.delegated);
    await verifyStore.setJSON(pubRec.verificationId, pubRec, { onlyIfNew: true });
    verificationRecord = stripVerificationSecrets(pubRec);
    cloudSynced = true;
  }

  const finalItem = {
    ...entry,
    id: entry.id || `out-${Date.now()}-${crypto.randomBytes(3).toString('hex')}`,
    sequenceNumber: claimedSeq,
    year,
    nomorSurat: finalNomor,
    documentSnapshot: snapshot,
    isLocked: true,
    signerRole: requestedSignerRole || entry.signerRole || 'kepala_sekolah',
    verificationStatus: authorizeSah ? 'sah' : (requestAcc ? 'pending_acc' : (entry.verificationStatus || 'unverified')),
    accStatus: authorizeSah ? 'approved' : (requestAcc ? 'pending' : (entry.accStatus || '')),
    isVerifiedSah: Boolean(authorizeSah),
    recordedBy: user.username,
    recordedByRole: user.role,
    requestedBy: requestAcc ? user.username : (entry.requestedBy || ''),
    requestedByName: requestAcc ? (user.displayName || user.username) : (entry.requestedByName || ''),
    requestedByRole: requestAcc ? user.role : (entry.requestedByRole || ''),
    requestedAt: requestAcc ? nowIso : (entry.requestedAt || null),
    requestNote: requestAcc ? requestNote : (entry.requestNote || ''),
    authorizedBy: authorizeSah ? user.username : '',
    authorizedByRole: authorizeSah ? user.role : '',
    delegatedByKepsek: Boolean(roleCheck.delegated),
    delegatedNote: roleCheck.delegated ? `Disahkan oleh TU (${user.displayName}) atas persetujuan delegasi Kepala Sekolah/Admin` : '',
    cloudSynced: authorizeSah ? cloudSynced : true,
    driveSynced: false,
    driveDocxFileId: '',
    driveDocxUrl: '',
    driveJsonFileId: '',
    driveJsonUrl: '',
    recordedAt: entry.recordedAt || nowIso
  };

  const gasRes = await callGasBridge(wsStore, env, 'archive_outgoing', {
    item: finalItem,
    docxBase64: body.docxBase64 || '',
    actorUsername: user.username
  }, fetchImpl);

  if (gasRes.ok && gasRes.data) {
    finalItem.driveSynced = true;
    finalItem.driveDocxFileId = gasRes.data.driveDocxFileId || '';
    finalItem.driveDocxUrl = gasRes.data.driveDocxUrl || '';
    finalItem.driveJsonFileId = gasRes.data.driveJsonFileId || '';
    finalItem.driveJsonUrl = gasRes.data.driveJsonUrl || '';
  }

  const nextList = sortOutgoingList([finalItem, ...existingList.filter(i => i.id !== finalItem.id)]);
  await wsStore.setJSON(BLOB_KEYS.OUTGOING_LIST, nextList);
  const syncMeta = await bumpSyncMeta(
    wsStore,
    requestAcc ? 'acc_requested' : (authorizeSah ? 'outgoing_sah' : 'outgoing_recorded'),
    nextList
  );

  return jsonResponse({
    ok: true,
    item: finalItem,
    verification: verificationRecord,
    syncRev: syncMeta.rev,
    driveSync: gasRes.ok ? { ok: true, url: finalItem.driveDocxUrl || finalItem.driveJsonUrl } : { ok: false, error: gasRes.error || '' }
  }, 201);
}

async function handleAuthorizeOutgoing(user, body, wsStore, verifyStore, env, fetchImpl) {
  const permissions = await getEffectivePermissions(wsStore);
  const itemId = String(body.outgoingId || '').trim();
  const requestedRole = String(body.signerRole || 'kepala_sekolah').trim();

  const roleCheck = canUserUseSignerRole(user, permissions, requestedRole);
  if (!roleCheck.allowed) {
    return jsonResponse({ ok: false, error: roleCheck.reason }, 403);
  }

  const existingList = (await wsStore.get(BLOB_KEYS.OUTGOING_LIST, { type: 'json' })) || [];
  const target = existingList.find(i => i.id === itemId);
  if (!target) {
    return jsonResponse({ ok: false, error: 'Catatan Surat Keluar tidak ditemukan.' }, 404);
  }

  const payload = body.verificationPayload || {};
  if (!payload.verificationId) {
    return jsonResponse({ ok: false, error: 'Payload verifikasi QR tidak valid.' }, 400);
  }

  const pubRec = buildPublicVerificationRecord(payload, user, roleCheck.delegated);
  await verifyStore.setJSON(pubRec.verificationId, pubRec);

  const updatedItem = {
    ...target,
    signerRole: payload.signerRole || requestedRole,
    signerTitle: payload.signerTitle || target.signerTitle,
    signerName: payload.signerName || target.signerName,
    signerNip: payload.signerNip ?? target.signerNip,
    verificationStatus: 'sah',
    accStatus: 'approved',
    isVerifiedSah: true,
    verificationId: payload.verificationId,
    verificationUrl: payload.verificationUrl || target.verificationUrl,
    qrDataUrl: body.qrDataUrl || target.qrDataUrl,
    manageToken: payload.manageToken || target.manageToken,
    verificationStatement: payload.verificationStatement || target.verificationStatement,
    verificationDetails: payload.details || target.verificationDetails,
    cloudSynced: true,
    authorizedBy: user.username,
    authorizedByRole: user.role,
    delegatedByKepsek: Boolean(roleCheck.delegated),
    authorizedAt: new Date().toISOString(),
    rejectedAt: null,
    rejectedReason: '',
    revokedAt: null,
    revokedReason: ''
  };

  const nextList = sortOutgoingList(existingList.map(i => (i.id === itemId ? updatedItem : i)));
  await wsStore.setJSON(BLOB_KEYS.OUTGOING_LIST, nextList);
  const syncMeta = await bumpSyncMeta(wsStore, 'outgoing_authorized', nextList);

  const gasRes = await callGasBridge(wsStore, env, 'update_outgoing_status', {
    item: updatedItem,
    statusAction: 'SAH_QR',
    actorUsername: user.username,
    note: roleCheck.delegated ? 'Pengesahan delegasi TU' : 'Pengesahan resmi'
  }, fetchImpl);

  return jsonResponse({
    ok: true,
    item: updatedItem,
    verification: stripVerificationSecrets(pubRec),
    syncRev: syncMeta.rev,
    driveSync: gasRes
  }, 200);
}

async function handleRevokeOutgoing(user, body, wsStore, verifyStore, env, fetchImpl) {
  const permissions = await getEffectivePermissions(wsStore);
  const itemId = String(body.outgoingId || '').trim();
  const nextStatus = body.status === 'sah' ? 'sah' : 'revoked';
  const reason = String(body.reason || '').trim();

  const existingList = (await wsStore.get(BLOB_KEYS.OUTGOING_LIST, { type: 'json' })) || [];
  const target = existingList.find(i => i.id === itemId);
  if (!target) {
    return jsonResponse({ ok: false, error: 'Catatan Surat Keluar tidak ditemukan.' }, 404);
  }

  if (!canUserRevokeLetter(user, permissions, target)) {
    return jsonResponse({ ok: false, error: 'Anda tidak memiliki izin mencabut / mengubah status pengesahan surat ini.' }, 403);
  }

  if (target.verificationId) {
    const verExisting = await verifyStore.get(target.verificationId, { type: 'json' });
    if (verExisting) {
      const nowIso = new Date().toISOString();
      await verifyStore.setJSON(target.verificationId, {
        ...verExisting,
        status: nextStatus,
        updatedAt: nowIso,
        revokedAt: nextStatus === 'revoked' ? nowIso : null,
        revokedReason: nextStatus === 'revoked' ? (reason || 'Dicabut oleh pejabat berwenang.') : '',
        updatedByUser: user.username
      });
    }
  }

  const nowIso = new Date().toISOString();
  const updatedItem = {
    ...target,
    verificationStatus: nextStatus,
    isVerifiedSah: nextStatus === 'sah',
    revokedAt: nextStatus === 'revoked' ? nowIso : null,
    revokedReason: nextStatus === 'revoked' ? (reason || 'Dicabut oleh pejabat berwenang.') : '',
    cloudSynced: true
  };

  const nextList = sortOutgoingList(existingList.map(i => (i.id === itemId ? updatedItem : i)));
  await wsStore.setJSON(BLOB_KEYS.OUTGOING_LIST, nextList);
  const syncMeta = await bumpSyncMeta(wsStore, nextStatus === 'revoked' ? 'outgoing_revoked' : 'outgoing_restored', nextList);

  await callGasBridge(wsStore, env, 'update_outgoing_status', {
    item: updatedItem,
    statusAction: nextStatus === 'revoked' ? 'CABUT_SAH' : 'SAHKAN_ULANG',
    actorUsername: user.username,
    note: updatedItem.revokedReason
  }, fetchImpl);

  return jsonResponse({ ok: true, item: updatedItem, syncRev: syncMeta.rev }, 200);
}

async function handleDeleteOutgoing(user, body, wsStore) {
  const permissions = await getEffectivePermissions(wsStore);
  const canDel = user.role === 'admin' || user.role === 'kepsek' || (user.role === 'tu' && permissions.allowTuDeleteOutgoing === true);
  if (!canDel) {
    return jsonResponse({ ok: false, error: 'Hanya Admin atau Kepala Sekolah yang dapat menghapus catatan Surat Keluar terkunci.' }, 403);
  }
  const itemId = String(body.outgoingId || '').trim();
  const existingList = (await wsStore.get(BLOB_KEYS.OUTGOING_LIST, { type: 'json' })) || [];
  const nextList = existingList.filter(i => i.id !== itemId);
  await wsStore.setJSON(BLOB_KEYS.OUTGOING_LIST, nextList);
  const syncMeta = await bumpSyncMeta(wsStore, 'outgoing_deleted', nextList);
  return jsonResponse({ ok: true, remaining: nextList.length, syncRev: syncMeta.rev }, 200);
}

async function handleUpdatePermissions(user, body, wsStore) {
  if (user.role !== 'admin' && user.role !== 'kepsek') {
    return jsonResponse({ ok: false, error: 'Hanya Kepala Sekolah atau Administrator yang dapat mengatur persetujuan hak akses.' }, 403);
  }

  const current = await getEffectivePermissions(wsStore);
  const patch = body.permissions || {};

  const next = {
    ...current,
    allowTuSignPrincipal: patch.allowTuSignPrincipal !== undefined ? Boolean(patch.allowTuSignPrincipal) : current.allowTuSignPrincipal,
    allowTuSignKepsekDanTu: patch.allowTuSignKepsekDanTu !== undefined ? Boolean(patch.allowTuSignKepsekDanTu) : current.allowTuSignKepsekDanTu,
    note: patch.note !== undefined ? String(patch.note).slice(0, 240) : current.note,
    updatedAt: new Date().toISOString(),
    updatedBy: user.username,
    updatedByRole: user.role
  };

  if (user.role === 'admin') {
    if (patch.allowTuRevokePrincipalQr !== undefined) next.allowTuRevokePrincipalQr = Boolean(patch.allowTuRevokePrincipalQr);
    if (patch.allowTuManageApiVault !== undefined) next.allowTuManageApiVault = Boolean(patch.allowTuManageApiVault);
    if (patch.allowTuDeleteOutgoing !== undefined) next.allowTuDeleteOutgoing = Boolean(patch.allowTuDeleteOutgoing);
  }

  await wsStore.setJSON(BLOB_KEYS.PERMISSIONS, next);
  const syncMeta = await bumpSyncMeta(wsStore, 'permissions_updated');
  return jsonResponse({ ok: true, permissions: next, syncRev: syncMeta.rev }, 200);
}

async function handleManageUsers(user, body, wsStore, env) {
  if (user.role !== 'admin') {
    return jsonResponse({ ok: false, error: 'Hanya Administrator yang memiliki hak penuh mengelola akun pengguna.' }, 403);
  }

  const users = await ensureInitialUsers(wsStore, env);
  const op = String(body.op || '').trim();
  const nowIso = new Date().toISOString();

  if (op === 'list') {
    return jsonResponse({ ok: true, users: users.map(sanitizePublicUser) }, 200);
  }

  if (op === 'create') {
    const username = String(body.username || '').trim().toLowerCase().replace(/[^a-z0-9._-]/g, '');
    const displayName = String(body.displayName || '').trim();
    const role = ['admin', 'kepsek', 'tu'].includes(body.role) ? body.role : 'tu';
    const password = String(body.password || '');

    if (!username || username.length < 3 || password.length < 6) {
      return jsonResponse({ ok: false, error: 'Username minimal 3 huruf & password/PIN minimal 6 karakter.' }, 400);
    }
    if (users.some(u => String(u.username || '').toLowerCase() === username)) {
      return jsonResponse({ ok: false, error: `Username "${username}" sudah digunakan.` }, 409);
    }

    const { salt, hash } = hashPasswordWithSalt(password);
    const created = {
      id: `usr_${Date.now()}_${crypto.randomBytes(2).toString('hex')}`,
      username,
      displayName: displayName || username,
      role,
      active: true,
      isDefaultAccount: false,
      passwordSalt: salt,
      passwordHash: hash,
      createdAt: nowIso,
      updatedAt: nowIso,
      createdBy: user.username
    };
    const nextUsers = [...users, created];
    await wsStore.setJSON(BLOB_KEYS.USERS, nextUsers);
    return jsonResponse({ ok: true, user: sanitizePublicUser(created), users: nextUsers.map(sanitizePublicUser) }, 201);
  }

  const targetId = String(body.targetUserId || '').trim();
  const idx = users.findIndex(u => u.id === targetId);
  if (idx < 0) {
    return jsonResponse({ ok: false, error: 'Akun pengguna tidak ditemukan.' }, 404);
  }
  const target = users[idx];

  if (op === 'update') {
    const nextRole = ['admin', 'kepsek', 'tu'].includes(body.role) ? body.role : target.role;
    const nextActive = body.active !== undefined ? Boolean(body.active) : (target.active !== false);

    // Pastikan selalu ada minimal 1 Admin aktif
    if (target.role === 'admin' && (nextRole !== 'admin' || !nextActive)) {
      const remainingAdmins = users.filter(u => u.id !== target.id && u.role === 'admin' && u.active !== false);
      if (remainingAdmins.length === 0) {
        return jsonResponse({ ok: false, error: 'Minimal harus ada 1 akun Administrator yang aktif.' }, 400);
      }
    }

    users[idx] = {
      ...target,
      displayName: body.displayName !== undefined ? String(body.displayName).trim() : target.displayName,
      role: nextRole,
      active: nextActive,
      updatedAt: nowIso
    };
    if (body.newPassword && String(body.newPassword).length >= 6) {
      const { salt, hash } = hashPasswordWithSalt(String(body.newPassword));
      users[idx].passwordSalt = salt;
      users[idx].passwordHash = hash;
      users[idx].isDefaultAccount = false;
    }

    await wsStore.setJSON(BLOB_KEYS.USERS, users);
    return jsonResponse({ ok: true, user: sanitizePublicUser(users[idx]), users: users.map(sanitizePublicUser) }, 200);
  }

  if (op === 'delete') {
    if (target.id === user.id) {
      return jsonResponse({ ok: false, error: 'Anda tidak dapat menghapus akun yang sedang dipakai login.' }, 400);
    }
    if (target.role === 'admin') {
      const otherAdmins = users.filter(u => u.id !== target.id && u.role === 'admin' && u.active !== false);
      if (otherAdmins.length === 0) {
        return jsonResponse({ ok: false, error: 'Tidak bisa menghapus satu-satunya akun Administrator.' }, 400);
      }
    }
    const nextUsers = users.filter(u => u.id !== targetId);
    await wsStore.setJSON(BLOB_KEYS.USERS, nextUsers);
    return jsonResponse({ ok: true, users: nextUsers.map(sanitizePublicUser) }, 200);
  }

  return jsonResponse({ ok: false, error: `Operasi akun tidak dikenali: ${op}` }, 400);
}

async function handleChangeOwnPassword(user, body, wsStore, env) {
  const oldPassword = String(body.oldPassword || '');
  const newPassword = String(body.newPassword || '');
  if (!oldPassword || newPassword.length < 6) {
    return jsonResponse({ ok: false, error: 'Password lama wajib diisi dan password baru minimal 6 karakter.' }, 400);
  }
  const users = await ensureInitialUsers(wsStore, env);
  const idx = users.findIndex(u => u.id === user.id);
  if (idx < 0 || !verifyPasswordHash(oldPassword, users[idx].passwordSalt, users[idx].passwordHash)) {
    return jsonResponse({ ok: false, error: 'Password lama tidak sesuai.' }, 403);
  }
  const { salt, hash } = hashPasswordWithSalt(newPassword);
  users[idx] = {
    ...users[idx],
    passwordSalt: salt,
    passwordHash: hash,
    isDefaultAccount: false,
    updatedAt: new Date().toISOString()
  };
  await wsStore.setJSON(BLOB_KEYS.USERS, users);
  return jsonResponse({ ok: true, message: 'Password berhasil diperbarui.' }, 200);
}

async function handleConfigureDrive(user, body, wsStore, env, fetchImpl) {
  if (user.role !== 'admin') {
    return jsonResponse({ ok: false, error: 'Hanya Administrator yang dapat mengatur konfigurasi Google Drive Bridge.' }, 403);
  }

  const current = (await wsStore.get(BLOB_KEYS.DRIVE_CONFIG, { type: 'json' })) || {};
  const next = {
    ...current,
    ...(body.webhookUrl !== undefined ? { webhookUrl: String(body.webhookUrl).trim() } : {}),
    ...(body.webhookSecret !== undefined && String(body.webhookSecret).trim() ? { webhookSecret: String(body.webhookSecret).trim() } : {}),
    ...(body.rootFolderId ? { rootFolderId: String(body.rootFolderId).trim() } : {})
  };
  await wsStore.setJSON(BLOB_KEYS.DRIVE_CONFIG, next);

  const testRes = await callGasBridge(wsStore, env, 'ping', {}, fetchImpl);
  const { publicStatus } = await getDriveConfigSafe(wsStore, env);
  return jsonResponse({
    ok: testRes.ok,
    driveStatus: publicStatus,
    error: testRes.ok ? '' : testRes.error
  }, testRes.ok ? 200 : 502);
}

async function handleSyncDrive(user, body, wsStore, env, fetchImpl) {
  const [outgoingList, letterVault, kopGallery, schoolProfile, apiVault] = await Promise.all([
    wsStore.get(BLOB_KEYS.OUTGOING_LIST, { type: 'json' }),
    wsStore.get(BLOB_KEYS.LETTER_VAULT, { type: 'json' }),
    wsStore.get(BLOB_KEYS.KOP_GALLERY, { type: 'json' }),
    wsStore.get(BLOB_KEYS.SCHOOL_PROFILE, { type: 'json' }),
    wsStore.get(BLOB_KEYS.API_VAULT, { type: 'json' })
  ]);

  if (body.mode === 'single_outgoing' && body.outgoingId) {
    const list = Array.isArray(outgoingList) ? outgoingList : [];
    const target = list.find(i => i.id === body.outgoingId);
    if (!target) return jsonResponse({ ok: false, error: 'Surat tidak ditemukan.' }, 404);
    const gasRes = await callGasBridge(wsStore, env, 'archive_outgoing', {
      item: target,
      docxBase64: body.docxBase64 || '',
      actorUsername: user.username
    }, fetchImpl);
    if (!gasRes.ok) {
      return jsonResponse({ ok: false, error: gasRes.error || 'Gagal kirim ke Drive.' }, 502);
    }
    const updated = list.map(i => (
      i.id === target.id
        ? {
            ...i,
            driveSynced: true,
            driveDocxFileId: gasRes.data.driveDocxFileId || i.driveDocxFileId,
            driveDocxUrl: gasRes.data.driveDocxUrl || i.driveDocxUrl,
            driveJsonFileId: gasRes.data.driveJsonFileId || i.driveJsonFileId,
            driveJsonUrl: gasRes.data.driveJsonUrl || i.driveJsonUrl
          }
        : i
    ));
    await wsStore.setJSON(BLOB_KEYS.OUTGOING_LIST, updated);
    const { publicStatus } = await getDriveConfigSafe(wsStore, env);
    return jsonResponse({ ok: true, item: updated.find(i => i.id === target.id), driveStatus: publicStatus }, 200);
  }

  const backupSnapshot = {
    exportedAt: new Date().toISOString(),
    exportedBy: user.username,
    schoolProfile: schoolProfile || null,
    outgoingLetters: (Array.isArray(outgoingList) ? outgoingList : []).map(({ manageToken, qrDataUrl, ...rest }) => rest),
    letterVault: Array.isArray(letterVault) ? letterVault : [],
    kopGalleryCount: Array.isArray(kopGallery) ? kopGallery.length : 0,
    apiVaultCount: Array.isArray(apiVault?.items) ? apiVault.items.length : 0
  };

  const backupRes = await callGasBridge(wsStore, env, 'backup_snapshot', {
    snapshot: backupSnapshot,
    actorUsername: user.username
  }, fetchImpl);
  const { publicStatus } = await getDriveConfigSafe(wsStore, env);

  return jsonResponse({
    ok: backupRes.ok,
    driveStatus: publicStatus,
    backupFileUrl: backupRes.data?.backupFileUrl || '',
    sheetUrl: backupRes.data?.sheetUrl || '',
    error: backupRes.ok ? '' : backupRes.error
  }, backupRes.ok ? 200 : 502);
}

/**
 * Polling ringan untuk pembaruan SPA real-time tanpa refresh halaman.
 */
async function handlePollUpdates(user, url, wsStore, env) {
  const since = Number(url.searchParams.get('since') || 0);
  const meta = await getSyncMeta(wsStore);

  if (since > 0 && meta.rev <= since) {
    return jsonResponse({
      ok: true,
      changed: false,
      rev: meta.rev,
      lastEvent: meta.lastEvent,
      pendingAccCount: meta.pendingAccCount
    }, 200);
  }

  const [
    permissions,
    outgoingList,
    driveCfg,
    schoolProfile,
    kopGallery,
    letterVault,
    apiVault
  ] = await Promise.all([
    getEffectivePermissions(wsStore),
    wsStore.get(BLOB_KEYS.OUTGOING_LIST, { type: 'json' }),
    getDriveConfigSafe(wsStore, env),
    wsStore.get(BLOB_KEYS.SCHOOL_PROFILE, { type: 'json' }),
    wsStore.get(BLOB_KEYS.KOP_GALLERY, { type: 'json' }),
    wsStore.get(BLOB_KEYS.LETTER_VAULT, { type: 'json' }),
    wsStore.get(BLOB_KEYS.API_VAULT, { type: 'json' })
  ]);

  const sortedOutgoing = Array.isArray(outgoingList) ? sortOutgoingList(outgoingList) : [];
  const pendingAccCount = sortedOutgoing.filter(
    i => i?.verificationStatus === 'pending_acc' || i?.accStatus === 'pending'
  ).length;

  return jsonResponse({
    ok: true,
    changed: true,
    rev: meta.rev,
    lastEvent: meta.lastEvent,
    pendingAccCount,
    permissions,
    driveStatus: driveCfg.publicStatus,
    workspace: {
      schoolProfile: schoolProfile || null,
      kopGallery: Array.isArray(kopGallery) ? kopGallery : null,
      letterVault: Array.isArray(letterVault) ? letterVault : null,
      apiVault: apiVault?.items ? apiVault : null,
      outgoingLetters: sortedOutgoing
    },
    syncedAt: new Date().toISOString()
  }, 200);
}

/**
 * Mengajukan permintaan ACC Kepala Sekolah untuk surat keluar yang belum disahkan QR.
 */
async function handleRequestAcc(user, body, wsStore, env, fetchImpl) {
  const itemId = String(body.outgoingId || '').trim();
  const requestedRole = String(body.signerRole || 'kepala_sekolah').trim();
  const requestNote = String(body.requestNote || '').trim().slice(0, 240);

  const [existingList, schoolProfile] = await Promise.all([
    wsStore.get(BLOB_KEYS.OUTGOING_LIST, { type: 'json' }),
    wsStore.get(BLOB_KEYS.SCHOOL_PROFILE, { type: 'json' })
  ]);
  const list = Array.isArray(existingList) ? existingList : [];
  const target = list.find(i => i.id === itemId);
  if (!target) {
    return jsonResponse({ ok: false, error: 'Catatan Surat Keluar tidak ditemukan.' }, 404);
  }
  if (target.verificationStatus === 'sah') {
    return jsonResponse({ ok: false, error: 'Surat ini sudah berstatus SAH dan memiliki QR verifikasi.' }, 400);
  }

  const signerInfo = buildSignerSelectionInfo(schoolProfile || {}, requestedRole);
  const nowIso = new Date().toISOString();
  const updatedItem = {
    ...target,
    signerRole: signerInfo.signerRole,
    signerTitle: signerInfo.signerTitle || target.signerTitle,
    signerName: signerInfo.signerName || target.signerName,
    signerNip: signerInfo.signerNip ?? target.signerNip,
    verificationStatus: 'pending_acc',
    accStatus: 'pending',
    isVerifiedSah: false,
    requestedBy: user.username,
    requestedByName: user.displayName || user.username,
    requestedByRole: user.role,
    requestedAt: nowIso,
    requestNote,
    rejectedAt: null,
    rejectedReason: '',
    revokedAt: null,
    revokedReason: '',
    cloudSynced: true
  };

  const nextList = sortOutgoingList(list.map(i => (i.id === itemId ? updatedItem : i)));
  await wsStore.setJSON(BLOB_KEYS.OUTGOING_LIST, nextList);
  const syncMeta = await bumpSyncMeta(wsStore, 'acc_requested', nextList);

  const gasRes = await callGasBridge(wsStore, env, 'update_outgoing_status', {
    item: updatedItem,
    statusAction: 'MINTA_ACC',
    actorUsername: user.username,
    note: requestNote || 'Menunggu ACC Kepala Sekolah'
  }, fetchImpl);

  return jsonResponse({
    ok: true,
    item: updatedItem,
    syncRev: syncMeta.rev,
    driveSync: gasRes
  }, 200);
}

/**
 * Kepala Sekolah atau Admin meng-ACC antrean surat, menerbitkan barcode QR & arsip Drive.
 */
async function handleApproveAcc(user, body, wsStore, verifyStore, env, fetchImpl) {
  if (user.role !== 'kepsek' && user.role !== 'admin') {
    return jsonResponse({ ok: false, error: 'Hanya Kepala Sekolah atau Administrator yang dapat meng-ACC dan menerbitkan Barcode QR.' }, 403);
  }

  const itemId = String(body.outgoingId || '').trim();
  const [existingList, storedSchool] = await Promise.all([
    wsStore.get(BLOB_KEYS.OUTGOING_LIST, { type: 'json' }),
    wsStore.get(BLOB_KEYS.SCHOOL_PROFILE, { type: 'json' })
  ]);
  const list = Array.isArray(existingList) ? existingList : [];
  const target = list.find(i => i.id === itemId);
  if (!target) {
    return jsonResponse({ ok: false, error: 'Surat antrean ACC tidak ditemukan.' }, 404);
  }

  const school = storedSchool || body.schoolProfile || {};
  const finalSignerRole = String(body.signerRole || target.signerRole || 'kepala_sekolah').trim();

  const payload = body.verificationPayload?.verificationId
    ? { ...body.verificationPayload }
    : (() => {
        const vId = target.verificationId || buildVerificationId(target.typeCode || 'TU', target.year);
        const vUrl = target.verificationUrl || buildVerificationUrl(vId, body.verificationOrigin || undefined);
        return buildVerificationPayload({
          entry: target,
          school,
          signerRole: finalSignerRole,
          verificationId: vId,
          manageToken: target.manageToken || undefined,
          verificationUrl: vUrl
        });
      })();

  const qrDataUrl = body.qrDataUrl || generateQrPngDataUrl(payload.verificationUrl);
  const pubRec = buildPublicVerificationRecord(payload, user, false);
  await verifyStore.setJSON(pubRec.verificationId, pubRec);

  let updatedSnapshot = target.documentSnapshot ? JSON.parse(JSON.stringify(target.documentSnapshot)) : null;
  let docxBase64 = String(body.docxBase64 || '');

  if (updatedSnapshot && Array.isArray(updatedSnapshot.blocks) && updatedSnapshot.blocks.length > 0) {
    updatedSnapshot = applySignatureModeToDocModel(updatedSnapshot, {
      mode: 'qr',
      qrDataUrl,
      verificationUrl: payload.verificationUrl,
      signerRole: payload.signerRole,
      school,
      signerInfo: payload,
      syncSignerFields: true
    });
    if (!docxBase64) {
      try {
        const bytes = await renderDocx(updatedSnapshot, 'uint8array');
        docxBase64 = Buffer.from(bytes).toString('base64');
      } catch (err) {
        console.warn('[approve_acc] Gagal merender ulang DOCX QR di server:', err?.message || err);
      }
    }
  }

  const nowIso = new Date().toISOString();
  const updatedItem = {
    ...target,
    signerRole: payload.signerRole || finalSignerRole,
    signerTitle: payload.signerTitle || target.signerTitle,
    signerName: payload.signerName || target.signerName,
    signerNip: payload.signerNip ?? target.signerNip,
    verificationStatus: 'sah',
    accStatus: 'approved',
    isVerifiedSah: true,
    verificationId: payload.verificationId,
    verificationUrl: payload.verificationUrl,
    qrDataUrl,
    manageToken: payload.manageToken || target.manageToken || '',
    verificationStatement: payload.verificationStatement || '',
    verificationDetails: payload.details || [],
    documentSnapshot: updatedSnapshot,
    approvedBy: user.username,
    approvedByName: user.displayName || user.username,
    approvedByRole: user.role,
    approvedAt: nowIso,
    authorizedBy: user.username,
    authorizedByRole: user.role,
    authorizedAt: nowIso,
    delegatedByKepsek: false,
    rejectedAt: null,
    rejectedReason: '',
    revokedAt: null,
    revokedReason: '',
    cloudSynced: true
  };

  let gasRes;
  if (docxBase64) {
    gasRes = await callGasBridge(wsStore, env, 'archive_outgoing', {
      item: updatedItem,
      docxBase64,
      actorUsername: user.username
    }, fetchImpl);
    if (gasRes.ok && gasRes.data) {
      updatedItem.driveSynced = true;
      updatedItem.driveDocxFileId = gasRes.data.driveDocxFileId || updatedItem.driveDocxFileId || '';
      updatedItem.driveDocxUrl = gasRes.data.driveDocxUrl || updatedItem.driveDocxUrl || '';
      updatedItem.driveJsonFileId = gasRes.data.driveJsonFileId || updatedItem.driveJsonFileId || '';
      updatedItem.driveJsonUrl = gasRes.data.driveJsonUrl || updatedItem.driveJsonUrl || '';
    }
  } else {
    gasRes = await callGasBridge(wsStore, env, 'update_outgoing_status', {
      item: updatedItem,
      statusAction: 'ACC_KEPSEK_SAH',
      actorUsername: user.username,
      note: `Di-ACC resmi oleh ${user.displayName || user.username}`
    }, fetchImpl);
  }

  const nextList = sortOutgoingList(list.map(i => (i.id === itemId ? updatedItem : i)));
  await wsStore.setJSON(BLOB_KEYS.OUTGOING_LIST, nextList);
  const syncMeta = await bumpSyncMeta(wsStore, 'acc_approved', nextList);

  return jsonResponse({
    ok: true,
    item: updatedItem,
    verification: stripVerificationSecrets(pubRec),
    syncRev: syncMeta.rev,
    driveSync: gasRes
  }, 200);
}

/**
 * Kepala Sekolah atau Admin menolak permintaan ACC agar TU dapat merevisi surat.
 */
async function handleRejectAcc(user, body, wsStore, env, fetchImpl) {
  if (user.role !== 'kepsek' && user.role !== 'admin') {
    return jsonResponse({ ok: false, error: 'Hanya Kepala Sekolah atau Administrator yang dapat menolak pengajuan ACC.' }, 403);
  }

  const itemId = String(body.outgoingId || '').trim();
  const reason = String(body.reason || '').trim().slice(0, 300) || 'Perlu revisi oleh TU sebelum pengesahan Kepala Sekolah.';
  const existingList = (await wsStore.get(BLOB_KEYS.OUTGOING_LIST, { type: 'json' })) || [];
  const target = existingList.find(i => i.id === itemId);
  if (!target) {
    return jsonResponse({ ok: false, error: 'Surat antrean ACC tidak ditemukan.' }, 404);
  }

  const nowIso = new Date().toISOString();
  const updatedItem = {
    ...target,
    verificationStatus: 'rejected',
    accStatus: 'rejected',
    isVerifiedSah: false,
    rejectedBy: user.username,
    rejectedByName: user.displayName || user.username,
    rejectedByRole: user.role,
    rejectedAt: nowIso,
    rejectedReason: reason,
    cloudSynced: true
  };

  const nextList = sortOutgoingList(existingList.map(i => (i.id === itemId ? updatedItem : i)));
  await wsStore.setJSON(BLOB_KEYS.OUTGOING_LIST, nextList);
  const syncMeta = await bumpSyncMeta(wsStore, 'acc_rejected', nextList);

  const gasRes = await callGasBridge(wsStore, env, 'update_outgoing_status', {
    item: updatedItem,
    statusAction: 'TOLAK_ACC',
    actorUsername: user.username,
    note: reason
  }, fetchImpl);

  return jsonResponse({
    ok: true,
    item: updatedItem,
    syncRev: syncMeta.rev,
    driveSync: gasRes
  }, 200);
}

/**
 * Membatalkan pengajuan ACC yang masih menunggu (kembali ke status non-QR).
 */
async function handleCancelAcc(user, body, wsStore) {
  const itemId = String(body.outgoingId || '').trim();
  const existingList = (await wsStore.get(BLOB_KEYS.OUTGOING_LIST, { type: 'json' })) || [];
  const target = existingList.find(i => i.id === itemId);
  if (!target) {
    return jsonResponse({ ok: false, error: 'Catatan Surat Keluar tidak ditemukan.' }, 404);
  }

  const isOwner = target.requestedBy === user.username || target.recordedBy === user.username;
  if (!isOwner && user.role !== 'admin' && user.role !== 'kepsek') {
    return jsonResponse({ ok: false, error: 'Anda tidak berwenang membatalkan pengajuan ACC surat ini.' }, 403);
  }
  if (target.verificationStatus !== 'pending_acc' && target.accStatus !== 'pending') {
    return jsonResponse({ ok: false, error: 'Pengajuan ACC surat ini sudah diproses atau tidak dalam status menunggu.' }, 400);
  }

  const updatedItem = {
    ...target,
    verificationStatus: 'unverified',
    accStatus: 'cancelled',
    isVerifiedSah: false,
    cancelledAt: new Date().toISOString(),
    cancelledBy: user.username,
    cloudSynced: true
  };

  const nextList = sortOutgoingList(existingList.map(i => (i.id === itemId ? updatedItem : i)));
  await wsStore.setJSON(BLOB_KEYS.OUTGOING_LIST, nextList);
  const syncMeta = await bumpSyncMeta(wsStore, 'acc_cancelled', nextList);

  return jsonResponse({ ok: true, item: updatedItem, syncRev: syncMeta.rev }, 200);
}

/**
 * Router utama yang diekspor untuk runtime Netlify dan dapat diuji langsung di Node.js unit test.
 */
export async function handleWorkspaceRequest(req, options = {}) {
  const wsStore = (options.workspaceStoreFactory || defaultWorkspaceStoreFactory)();
  const verifyStore = (options.verificationStoreFactory || defaultVerificationStoreFactory)();
  const env = options.env || null;
  const fetchImpl = options.fetchImpl || globalThis.fetch;

  const url = new URL(req.url);
  const method = (req.method || 'GET').toUpperCase();
  const action = (
    options.context?.params?.action ||
    url.searchParams.get('action') ||
    ''
  ).trim();

  if (method === 'GET' && action === 'health') {
    const { publicStatus } = await getDriveConfigSafe(wsStore, env);
    return jsonResponse({ ok: true, service: 'workspace', driveConfigured: publicStatus.configured }, 200);
  }

  if (method === 'POST' && action === 'login') {
    let body = {};
    try { body = await req.json(); } catch { return jsonResponse({ ok: false, error: 'Payload JSON login tidak valid.' }, 400); }
    return handleLogin(body, wsStore, env);
  }

  const user = await authenticateRequest(req, wsStore, env);
  if (!user) {
    return jsonResponse({ ok: false, error: 'Sesi login tidak valid atau telah berakhir. Silakan masuk kembali.' }, 401);
  }

  if (method === 'GET') {
    if (action === 'session') {
      const permissions = await getEffectivePermissions(wsStore);
      const { publicStatus } = await getDriveConfigSafe(wsStore, env);
      return jsonResponse({
        ok: true,
        user: sanitizePublicUser(user),
        permissions,
        driveStatus: publicStatus
      }, 200);
    }
    if (action === 'bootstrap') {
      return handleBootstrap(user, wsStore, env);
    }
    if (action === 'poll_updates') {
      return handlePollUpdates(user, url, wsStore, env);
    }
    return jsonResponse({ ok: false, error: `Aksi GET tidak dikenali: ${action}` }, 400);
  }

  if (method === 'POST') {
    let body = {};
    try { body = await req.json(); } catch { return jsonResponse({ ok: false, error: 'Payload JSON tidak valid.' }, 400); }

    switch (action) {
      case 'save_state':
        return handleSaveState(user, body, wsStore, env, fetchImpl);
      case 'migrate_local':
        return handleMigrateLocal(user, body, wsStore);
      case 'record_outgoing':
        return handleRecordOutgoing(user, body, wsStore, verifyStore, env, fetchImpl);
      case 'authorize_outgoing':
        return handleAuthorizeOutgoing(user, body, wsStore, verifyStore, env, fetchImpl);
      case 'request_acc':
        return handleRequestAcc(user, body, wsStore, env, fetchImpl);
      case 'approve_acc':
        return handleApproveAcc(user, body, wsStore, verifyStore, env, fetchImpl);
      case 'reject_acc':
        return handleRejectAcc(user, body, wsStore, env, fetchImpl);
      case 'cancel_acc':
        return handleCancelAcc(user, body, wsStore);
      case 'revoke_outgoing':
        return handleRevokeOutgoing(user, body, wsStore, verifyStore, env, fetchImpl);
      case 'delete_outgoing':
        return handleDeleteOutgoing(user, body, wsStore);
      case 'update_permissions':
        return handleUpdatePermissions(user, body, wsStore);
      case 'manage_users':
        return handleManageUsers(user, body, wsStore, env);
      case 'change_password':
        return handleChangeOwnPassword(user, body, wsStore, env);
      case 'configure_drive':
        return handleConfigureDrive(user, body, wsStore, env, fetchImpl);
      case 'sync_drive':
        return handleSyncDrive(user, body, wsStore, env, fetchImpl);
      default:
        return jsonResponse({ ok: false, error: `Aksi POST tidak dikenali: ${action}` }, 400);
    }
  }

  return jsonResponse({ ok: false, error: `Metode ${method} tidak didukung.` }, 405);
}

export default async function handler(req, context) {
  try {
    return await handleWorkspaceRequest(req, { context });
  } catch (err) {
    console.error('[workspace] Error:', err);
    return jsonResponse({ ok: false, error: `Gangguan server workspace: ${err?.message || 'internal error'}` }, 500);
  }
}
