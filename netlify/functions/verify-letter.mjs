/**
 * netlify/functions/verify-letter.mjs
 * Netlify Function + Netlify Blobs untuk registrasi, verifikasi publik QR Code,
 * dan pencabutan/pengesahan ulang Surat Keluar Resmi.
 */

import { getStore } from '@netlify/blobs';

const STORE_NAME = 'letter-verifications';
const VALID_ID_RE = /^[A-Za-z0-9_-]{6,80}$/;

export const config = {
  path: ['/verify-letter', '/verify-letter/:id']
};

function jsonResponse(payload, status = 200) {
  return new Response(JSON.stringify(payload), {
    status,
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store, max-age=0'
    }
  });
}

async function sha256Hex(value) {
  const bytes = new TextEncoder().encode(String(value || ''));
  const digest = await crypto.subtle.digest('SHA-256', bytes);
  return Array.from(new Uint8Array(digest))
    .map(b => b.toString(16).padStart(2, '0'))
    .join('');
}

function sanitizePublicRecord(record) {
  if (!record || typeof record !== 'object') return null;
  const { manageHash, ...publicData } = record;
  return publicData;
}

function defaultBlobStoreFactory() {
  return getStore({ name: STORE_NAME, consistency: 'strong' });
}

/**
 * Handler utama yang juga dapat diuji langsung pada unit test dengan store mock
 */
export async function handleVerifyRequest(req, context = {}, storeFactory = defaultBlobStoreFactory) {
  const method = (req.method || 'GET').toUpperCase();
  const url = new URL(req.url);
  const store = storeFactory();

  // Ekstrak ID dari param rute (/api/verify/:id) atau query string (?v= / ?id=)
  const paramId = (
    context?.params?.id ||
    url.searchParams.get('v') ||
    url.searchParams.get('id') ||
    ''
  ).trim();

  if (method === 'GET') {
    if (!paramId || !VALID_ID_RE.test(paramId)) {
      return jsonResponse({ ok: false, error: 'Kode verifikasi surat tidak valid.' }, 400);
    }

    const record = await store.get(paramId, { type: 'json' });
    if (!record) {
      return jsonResponse({ ok: false, status: 'not_found', error: 'Data pengesahan surat tidak ditemukan.' }, 404);
    }

    return jsonResponse({ ok: true, verification: sanitizePublicRecord(record) }, 200);
  }

  if (method === 'POST') {
    let body;
    try {
      body = await req.json();
    } catch {
      return jsonResponse({ ok: false, error: 'Payload JSON tidak valid.' }, 400);
    }

    const verificationId = String(body?.verificationId || '').trim();
    const manageToken = String(body?.manageToken || '').trim();
    if (!VALID_ID_RE.test(verificationId) || manageToken.length < 12) {
      return jsonResponse({ ok: false, error: 'ID verifikasi atau token pengelolaan tidak valid.' }, 400);
    }

    const nowIso = new Date().toISOString();
    const manageHash = await sha256Hex(manageToken);

    const record = {
      verificationId,
      status: 'sah',
      sequenceNumber: Number(body.sequenceNumber) || 0,
      year: String(body.year || ''),
      typeCode: String(body.typeCode || 'DINAS').toUpperCase().slice(0, 20),
      typeLabel: String(body.typeLabel || 'Surat Dinas Sekolah').slice(0, 120),
      nomorSurat: String(body.nomorSurat || '-').slice(0, 120),
      title: String(body.title || 'SURAT DINAS').slice(0, 180),
      perihal: String(body.perihal || body.title || '-').slice(0, 220),
      recipientName: String(body.recipientName || '-').slice(0, 180),
      tanggalSurat: String(body.tanggalSurat || '-').slice(0, 80),
      schoolName: String(body.schoolName || 'SMK IT AL-HAWARI').slice(0, 160),
      schoolNpsn: String(body.schoolNpsn || '').slice(0, 40),
      schoolAddress: String(body.schoolAddress || '').slice(0, 240),
      signerRole: String(body.signerRole || 'kepala_sekolah').slice(0, 40),
      signerTitle: String(body.signerTitle || 'Kepala Sekolah').slice(0, 120),
      signerName: String(body.signerName || 'Kepala Sekolah').slice(0, 140),
      signerNip: String(body.signerNip || '-').slice(0, 60),
      verificationStatement: String(body.verificationStatement || '').slice(0, 400),
      details: Array.isArray(body.details)
        ? body.details.slice(0, 14).map(d => ({
            label: String(d?.label || '').slice(0, 80),
            value: String(d?.value || '').slice(0, 200)
          }))
        : [],
      summary: String(body.summary || '').slice(0, 320),
      issuedAt: nowIso,
      updatedAt: nowIso,
      revokedAt: null,
      revokedReason: '',
      manageHash
    };

    const writeResult = await store.setJSON(verificationId, record, { onlyIfNew: true });
    if (writeResult && writeResult.modified === false) {
      return jsonResponse({ ok: false, error: 'Kode verifikasi sudah terdaftar sebelumnya.' }, 409);
    }

    return jsonResponse({ ok: true, verification: sanitizePublicRecord(record) }, 201);
  }

  if (method === 'PATCH') {
    let body;
    try {
      body = await req.json();
    } catch {
      return jsonResponse({ ok: false, error: 'Payload JSON tidak valid.' }, 400);
    }

    const verificationId = String(body?.verificationId || paramId || '').trim();
    const manageToken = String(body?.manageToken || '').trim();
    const nextStatus = body?.status === 'revoked' ? 'revoked' : (body?.status === 'sah' ? 'sah' : '');

    if (!VALID_ID_RE.test(verificationId) || !manageToken || !nextStatus) {
      return jsonResponse({ ok: false, error: 'Parameter pembaruan status verifikasi tidak lengkap.' }, 400);
    }

    const existing = await store.get(verificationId, { type: 'json' });
    if (!existing) {
      return jsonResponse({ ok: false, error: 'Data pengesahan surat tidak ditemukan di server.' }, 404);
    }

    const incomingHash = await sha256Hex(manageToken);
    if (!existing.manageHash || incomingHash !== existing.manageHash) {
      return jsonResponse({ ok: false, error: 'Token otorisasi tidak cocok. Perubahan ditolak.' }, 403);
    }

    const nowIso = new Date().toISOString();
    const updated = {
      ...existing,
      status: nextStatus,
      updatedAt: nowIso,
      revokedAt: nextStatus === 'revoked' ? nowIso : null,
      revokedReason: nextStatus === 'revoked'
        ? String(body.reason || 'Dibatalkan / dicabut oleh admin persuratan sekolah.').slice(0, 220)
        : ''
    };

    await store.setJSON(verificationId, updated);
    return jsonResponse({ ok: true, verification: sanitizePublicRecord(updated) }, 200);
  }

  return jsonResponse({ ok: false, error: `Metode ${method} tidak didukung.` }, 405);
}

export default async function handler(req, context) {
  try {
    return await handleVerifyRequest(req, context);
  } catch (err) {
    console.error('[verify-letter] Error:', err);
    return jsonResponse({ ok: false, error: 'Terjadi gangguan pada layanan verifikasi surat.' }, 500);
  }
}
