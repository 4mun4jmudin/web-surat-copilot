/**
 * Digital Verification & QR Barcode Engine for Surat Keluar (Option B).
 *
 * Generates crisp monochrome PNG QR codes in both Browser and Node.js
 * using vendor/qrcode.mjs + a pure-JS PNG encoder (stored DEFLATE + CRC32),
 * builds type-specific verification payloads for Netlify Blobs + Functions,
 * and formats signer authorization statements (Kepala Sekolah / Tata Usaha / Bersama).
 */

import { qrcode } from '../../vendor/qrcode.mjs';
import { stringToBytes as utf8StringToBytes } from '../../vendor/qrcode_UTF8.mjs';
import { LETTER_TYPE_CODES } from './numbering.js';

export const DEFAULT_VERIFICATION_ORIGIN = 'https://copilot-dokumen-alhawari.netlify.app';

export const SIGNER_ROLE_OPTIONS = Object.freeze([
  {
    id: 'kepala_sekolah',
    label: 'Kepala Sekolah',
    shortLabel: 'Kepala Sekolah',
    description: 'Pengesahan atas nama Kepala Sekolah'
  },
  {
    id: 'tata_usaha',
    label: 'Tata Usaha (TU)',
    shortLabel: 'Tata Usaha (TU)',
    description: 'Pengesahan atas nama Kepala / Staf Tata Usaha (TU)'
  },
  {
    id: 'kepsek_dan_tu',
    label: 'Kepala Sekolah & TU',
    shortLabel: 'Kepsek & TU',
    description: 'Pengesahan resmi bersama Kepala Sekolah & Tata Usaha (TU)'
  }
]);

const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.length; i += 1) {
    crc = CRC32_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function adler32(bytes) {
  let s1 = 1;
  let s2 = 0;
  const MOD = 65521;
  for (let i = 0; i < bytes.length; i += 1) {
    s1 = (s1 + bytes[i]) % MOD;
    s2 = (s2 + s1) % MOD;
  }
  return ((s2 << 16) | s1) >>> 0;
}

function writeUint32BE(arr, offset, value) {
  arr[offset] = (value >>> 24) & 0xff;
  arr[offset + 1] = (value >>> 16) & 0xff;
  arr[offset + 2] = (value >>> 8) & 0xff;
  arr[offset + 3] = value & 0xff;
}

function makePngChunk(type, data) {
  const typeBytes = new TextEncoder().encode(type);
  const chunk = new Uint8Array(12 + data.length);
  writeUint32BE(chunk, 0, data.length);
  chunk.set(typeBytes, 4);
  chunk.set(data, 8);
  const crcInput = new Uint8Array(4 + data.length);
  crcInput.set(typeBytes, 0);
  crcInput.set(data, 4);
  writeUint32BE(chunk, 8 + data.length, crc32(crcInput));
  return chunk;
}

function wrapZlibStored(raw) {
  const MAX_BLOCK = 65535;
  const blockCount = Math.max(1, Math.ceil(raw.length / MAX_BLOCK));
  const out = new Uint8Array(2 + raw.length + blockCount * 5 + 4);
  let pos = 0;

  // Zlib header (CMF=0x78 DEFLATE 32K, FLG=0x01 FCHECK valid)
  out[pos++] = 0x78;
  out[pos++] = 0x01;

  for (let offset = 0, idx = 0; idx < blockCount; idx += 1) {
    const remaining = raw.length - offset;
    const len = Math.min(MAX_BLOCK, remaining);
    const isFinal = idx === blockCount - 1 ? 1 : 0;
    const nlen = (~len) & 0xffff;

    out[pos++] = isFinal;
    out[pos++] = len & 0xff;
    out[pos++] = (len >>> 8) & 0xff;
    out[pos++] = nlen & 0xff;
    out[pos++] = (nlen >>> 8) & 0xff;

    if (len > 0) {
      out.set(raw.subarray(offset, offset + len), pos);
      pos += len;
      offset += len;
    }
  }

  writeUint32BE(out, pos, adler32(raw));
  return out;
}

function bytesToBase64(bytes) {
  if (typeof Buffer !== 'undefined') {
    return Buffer.from(bytes).toString('base64');
  }
  let binary = '';
  const CHUNK = 0x8000;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(i, i + CHUNK));
  }
  return btoa(binary);
}

/**
 * Encode an 8-bit grayscale matrix (0 = black, 255 = white) into a valid PNG Uint8Array.
 */
function encodeGrayscalePngBytes(width, height, getPixelGray) {
  const rowSize = 1 + width;
  const raw = new Uint8Array(rowSize * height);
  for (let y = 0; y < height; y += 1) {
    const rowStart = y * rowSize;
    raw[rowStart] = 0; // filter type 0 (None)
    for (let x = 0; x < width; x += 1) {
      raw[rowStart + 1 + x] = getPixelGray(x, y);
    }
  }

  const ihdr = new Uint8Array(13);
  writeUint32BE(ihdr, 0, width);
  writeUint32BE(ihdr, 4, height);
  ihdr[8] = 8; // bit depth 8
  ihdr[9] = 0; // color type 0: grayscale (compatible with Word/Chrome/Android/iOS)
  ihdr[10] = 0;
  ihdr[11] = 0;
  ihdr[12] = 0;

  const signature = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
  const ihdrChunk = makePngChunk('IHDR', ihdr);
  const idatChunk = makePngChunk('IDAT', wrapZlibStored(raw));
  const iendChunk = makePngChunk('IEND', new Uint8Array(0));

  const totalLen = signature.length + ihdrChunk.length + idatChunk.length + iendChunk.length;
  const png = new Uint8Array(totalLen);
  let cursor = 0;
  for (const part of [signature, ihdrChunk, idatChunk, iendChunk]) {
    png.set(part, cursor);
    cursor += part.length;
  }
  return png;
}

/**
 * Generate a crisp QR Code PNG Data URL (`data:image/png;base64,...`).
 */
export function generateQrPngDataUrl(text, options = {}) {
  const payload = String(text || '').trim();
  if (!payload) {
    throw new Error('Konten QR code tidak boleh kosong.');
  }

  const errorCorrectionLevel = options.errorCorrectionLevel || 'M';
  const moduleSize = Math.max(3, Math.min(12, Number(options.moduleSize) || 6));
  const marginModules = Math.max(1, Math.min(6, Number(options.marginModules) || 2));

  // Ensure UTF-8 byte support is wired to qrcode-generator
  qrcode.stringToBytes = utf8StringToBytes;

  const qr = qrcode(0, errorCorrectionLevel);
  qr.addData(payload);
  qr.make();

  const count = qr.getModuleCount();
  const totalModules = count + marginModules * 2;
  const size = totalModules * moduleSize;

  const pngBytes = encodeGrayscalePngBytes(size, size, (px, py) => {
    const modX = Math.floor(px / moduleSize) - marginModules;
    const modY = Math.floor(py / moduleSize) - marginModules;
    if (modX < 0 || modY < 0 || modX >= count || modY >= count) {
      return 255;
    }
    return qr.isDark(modY, modX) ? 17 : 255;
  });

  return `data:image/png;base64,${bytesToBase64(pngBytes)}`;
}

function randomAlphaNum(length = 8) {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let result = '';
  const gCrypto = typeof globalThis !== 'undefined' ? globalThis.crypto : null;
  if (gCrypto?.getRandomValues) {
    const buf = new Uint8Array(length);
    gCrypto.getRandomValues(buf);
    for (let i = 0; i < length; i += 1) {
      result += alphabet[buf[i] % alphabet.length];
    }
    return result;
  }
  for (let i = 0; i < length; i += 1) {
    result += alphabet[Math.floor(Math.random() * alphabet.length)];
  }
  return result;
}

/**
 * Build a readable, URL-safe verification ID, e.g. SMKIT-2026-KET-A7K9M2P4.
 */
export function buildVerificationId(typeCode = 'TU', year = new Date().getFullYear()) {
  const safeCode = String(typeCode || 'TU').toUpperCase().replace(/[^A-Z0-9]/g, '').slice(0, 6) || 'TU';
  const safeYear = String(year || new Date().getFullYear()).replace(/[^0-9]/g, '').slice(0, 4) || String(new Date().getFullYear());
  return `SMKIT-${safeYear}-${safeCode}-${randomAlphaNum(8)}`;
}

/**
 * Generate a high-entropy secret manageToken stored in the school's local archive
 * so only the originating admin can revoke/re-authorize the record in Netlify Blobs.
 */
export function generateManageToken() {
  return `tok_${randomAlphaNum(24)}_${Date.now().toString(36)}`;
}

/**
 * Resolve the public origin that phone cameras can reach when scanning the printed QR code.
 */
export function resolveVerificationBaseUrl(overrideOrigin) {
  if (overrideOrigin && /^https?:\/\//i.test(overrideOrigin)) {
    return overrideOrigin.replace(/\/+$/, '');
  }
  if (typeof window !== 'undefined' && window.location?.origin) {
    const origin = window.location.origin;
    const hostname = window.location.hostname || '';
    const isLocalHost = hostname === 'localhost' || hostname === '127.0.0.1' || hostname.startsWith('192.168.');
    if (origin.startsWith('https://') && !isLocalHost) {
      return origin.replace(/\/+$/, '');
    }
  }
  return DEFAULT_VERIFICATION_ORIGIN;
}

export function buildVerificationUrl(verificationId, baseUrl) {
  const origin = resolveVerificationBaseUrl(baseUrl);
  return `${origin}/?v=${encodeURIComponent(String(verificationId || '').trim())}`;
}

/**
 * Normalize & resolve the official authorization signer info and legal verification statement
 * according to selected authority: 'kepala_sekolah' | 'tata_usaha' | 'kepsek_dan_tu'.
 */
export function buildSignerSelectionInfo(school = {}, signerRole = 'kepala_sekolah', overrides = {}) {
  const normalizedRole = ['kepala_sekolah', 'tata_usaha', 'kepsek_dan_tu'].includes(signerRole)
    ? signerRole
    : 'kepala_sekolah';

  const schoolName = school?.name || 'SMK IT Al-Hawari';
  const principalTitle = school?.principalTitle || 'Kepala Sekolah';
  const principalName = school?.principalName || 'Dian Nurhadian, S.Pd.';
  const principalNip = school?.principalNip || '-';
  const principalRank = school?.principalRank || `Kepala ${schoolName}`;

  const tuTitle = school?.tuTitle || 'Kepala Tata Usaha';
  const tuName = school?.tuName || 'Staf Tata Usaha (TU)';
  const tuNip = school?.tuNip || '-';
  const tuRank = school?.tuRank || `Tata Usaha ${schoolName}`;

  if (normalizedRole === 'tata_usaha') {
    const signerName = String(overrides.signerName || tuName).trim();
    const signerTitle = String(overrides.signerTitle || tuTitle).trim();
    const signerNip = overrides.signerNip !== undefined ? String(overrides.signerNip) : tuNip;
    const signerRank = String(overrides.signerRank || tuRank).trim();
    return {
      signerRole: 'tata_usaha',
      roleLabel: 'Tata Usaha (TU)',
      signerTitle,
      signerName,
      signerNip,
      signerRank,
      principalName,
      principalTitle,
      tuName,
      tuTitle,
      verificationStatement: `Surat resmi ini telah diverifikasi dan dinyatakan SAH atas nama Tata Usaha (TU) (${signerName}) pada ${schoolName}.`
    };
  }

  if (normalizedRole === 'kepsek_dan_tu') {
    const signerName = String(overrides.signerName || `${principalName} & ${tuName}`).trim();
    const signerTitle = String(overrides.signerTitle || 'Kepala Sekolah / Tata Usaha (TU)').trim();
    const signerNip = overrides.signerNip !== undefined
      ? String(overrides.signerNip)
      : (principalNip && principalNip !== '-' ? principalNip : tuNip);
    const signerRank = String(overrides.signerRank || `Kepala Sekolah & Tata Usaha ${schoolName}`).trim();
    return {
      signerRole: 'kepsek_dan_tu',
      roleLabel: 'Kepala Sekolah & Tata Usaha (TU)',
      signerTitle,
      signerName,
      signerNip,
      signerRank,
      principalName,
      principalTitle,
      tuName,
      tuTitle,
      verificationStatement: `Surat resmi ini telah diverifikasi dan dinyatakan SAH atas nama Kepala Sekolah (${principalName}) dan Tata Usaha (TU) (${tuName}) pada ${schoolName}.`
    };
  }

  const signerName = String(overrides.signerName || principalName).trim();
  const signerTitle = String(overrides.signerTitle || principalTitle).trim();
  const signerNip = overrides.signerNip !== undefined ? String(overrides.signerNip) : principalNip;
  const signerRank = String(overrides.signerRank || principalRank).trim();
  return {
    signerRole: 'kepala_sekolah',
    roleLabel: 'Kepala Sekolah',
    signerTitle,
    signerName,
    signerNip,
    signerRank,
    principalName,
    principalTitle,
    tuName,
    tuTitle,
    verificationStatement: `Surat resmi ini telah diverifikasi dan dinyatakan SAH atas nama Kepala Sekolah (${signerName}) pada ${schoolName}.`
  };
}

/**
 * Format NIP line shown beneath the signer name:
 * always shows "NIP. -" when no NIP number is configured.
 */
export function formatSignerNipLine(nip) {
  const raw = String(nip ?? '').trim();
  if (!raw || raw === '-') return 'NIP. -';
  return /^NIP\b/i.test(raw) ? raw : `NIP. ${raw}`;
}

/**
 * Resolve the short caption rendered directly underneath the QR code:
 * single clean line "Disahkan secara elektronik" (signer name is already shown below).
 */
export function buildQrCaptionLines() {
  return {
    line1: 'Disahkan secara elektronik',
    line2: '',
    fullText: 'Disahkan secara elektronik'
  };
}

export function getTypeLabel(typeCode = 'TU') {
  const code = String(typeCode || 'TU').toUpperCase().trim();
  return LETTER_TYPE_CODES[code] || `Surat Dinas (${code})`;
}

function summarizeSnapshotText(snapshot) {
  if (!snapshot) return '';
  if (typeof snapshot.rawHtml === 'string' && snapshot.rawHtml.trim()) {
    return snapshot.rawHtml
      .replace(/<style[\s\S]*?<\/style>/gi, ' ')
      .replace(/<[^>]+>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 360);
  }
  if (Array.isArray(snapshot.blocks)) {
    const chunks = [];
    for (const blk of snapshot.blocks) {
      if (blk.type === 'paragraph') {
        if (Array.isArray(blk.runs)) {
          const text = blk.runs.map(r => r?.text || '').join('');
          if (text) chunks.push(text);
        } else if (blk.data?.text) {
          chunks.push(blk.data.text);
        } else if (blk.text) {
          chunks.push(blk.text);
        }
      } else if (blk.type === 'header' && blk.data?.text) {
        chunks.push(blk.data.text);
      } else if (blk.type === 'letterTitle' && blk.title) {
        chunks.push(blk.title);
      } else if (blk.type === 'list' && Array.isArray(blk.data?.items)) {
        chunks.push(blk.data.items.join('; '));
      }
      if (chunks.join(' ').length > 320) break;
    }
    return chunks.join(' ').replace(/\s+/g, ' ').trim().slice(0, 360);
  }
  return '';
}

const TYPE_AUTHENTICATION_NOTES = Object.freeze({
  KET: 'Dokumen Keterangan Resmi Sekolah (verifikasi status/identitas/keadaan yang diterangkan).',
  UND: 'Surat Undangan Rapat / Kegiatan Resmi Sekolah kepada pihak terkait.',
  KEP: 'Surat Keputusan (SK) resmi yang ditetapkan dan mengikat di lingkungan satuan pendidikan.',
  ST: 'Surat Tugas resmi kedinasan bagi pendidik, tenaga kependidikan, atau delegasi sekolah.',
  PGL: 'Surat Panggilan Resmi kepada orang tua/wali atau peserta didik untuk koordinasi sekolah.',
  DISP: 'Surat Izin / Dispensasi resmi mengikuti kegiatan akademik atau non-akademik.',
  REK: 'Surat Rekomendasi / Pengantar resmi dari satuan pendidikan.',
  PKL: 'Surat resmi terkait Praktik Kerja Lapangan (PKL) / kemitraan DUDI.',
  BA: 'Berita Acara resmi penyelenggaraan kegiatan / penyerahan / rapat.',
  PBR: 'Surat Pemberitahuan / Edaran resmi kepada warga sekolah atau instansi.',
  PRM: 'Surat Permohonan resmi institusi sekolah kepada pihak eksternal/terkait.',
  KUR: 'Surat kedinasan bidang Kurikulum & Evaluasi Pembelajaran.',
  KES: 'Surat kedinasan bidang Kesiswaan & Pembinaan Siswa.',
  TU: 'Surat Dinas Umum Ketatausahaan Sekolah.'
});

/**
 * Build structured detail rows rendered on the public ?v=... verification page
 * tailored to the exact typeCode of the outgoing letter.
 */
export function buildTypeSpecificVerificationDetails(typeCode = 'TU', entry = {}, school = {}) {
  const code = String(typeCode || 'TU').toUpperCase().trim() || 'TU';
  const typeLabel = getTypeLabel(code);
  const schoolName = school?.name || 'SMK IT Al-Hawari';
  const recipient = entry.recipientName || '-';
  const perihal = entry.perihal || entry.title || '-';
  const nomor = entry.nomorSurat || '-';
  const tanggal = entry.tanggalSurat || '-';

  const details = [
    { label: 'Jenis Dokumen', value: `${typeLabel} (Kode: ${code})` },
    { label: 'Nomor Surat Resmi', value: nomor },
    { label: 'Judul Surat', value: entry.title || typeLabel },
    { label: 'Perihal / Maksud', value: perihal },
    { label: 'Tujuan / Penerima', value: recipient },
    { label: 'Tanggal Surat', value: tanggal }
  ];

  if (code === 'ST') {
    details.push({ label: 'Pihak Ditugaskan / Tujuan', value: recipient });
  } else if (code === 'KEP') {
    details.push({ label: 'Status Ketetapan', value: `Keputusan Resmi Berlaku di ${schoolName}` });
  } else if (code === 'UND') {
    details.push({ label: 'Undangan Kepada', value: recipient });
  } else if (code === 'KET') {
    details.push({ label: 'Subjek Keterangan', value: recipient !== '-' ? recipient : perihal });
  } else if (code === 'PKL') {
    details.push({ label: 'Mitra / Lokasi PKL', value: recipient });
  } else if (code === 'PGL') {
    details.push({ label: 'Pihak Dipanggil', value: recipient });
  }

  details.push({
    label: 'Sifat Autentikasi',
    value: TYPE_AUTHENTICATION_NOTES[code] || TYPE_AUTHENTICATION_NOTES.TU
  });

  const summary = String(entry.summary || '').trim() || summarizeSnapshotText(entry.documentSnapshot);
  if (summary) {
    details.push({ label: 'Ringkasan Isi', value: summary.slice(0, 360) });
  }

  return details;
}

/**
 * Construct the full verification record payload for POST /api/verify and local storage.
 */
export function buildVerificationPayload({
  entry = {},
  school = {},
  signerRole = 'kepala_sekolah',
  verificationId,
  manageToken,
  verificationUrl
}) {
  const code = String(entry.typeCode || 'TU').toUpperCase().trim() || 'TU';
  const vId = verificationId || entry.verificationId || buildVerificationId(code, entry.year);
  const vUrl = verificationUrl || entry.verificationUrl || buildVerificationUrl(vId);
  const token = manageToken || entry.manageToken || generateManageToken();
  const signerInfo = buildSignerSelectionInfo(school, signerRole, {
    signerName: entry.customSignerName,
    signerTitle: entry.customSignerTitle,
    signerNip: entry.customSignerNip
  });
  const typeLabel = getTypeLabel(code);
  const details = buildTypeSpecificVerificationDetails(code, entry, school);
  const summary = String(entry.summary || '').trim() || summarizeSnapshotText(entry.documentSnapshot);

  return {
    verificationId: vId,
    verificationUrl: vUrl,
    manageToken: token,
    status: 'sah',
    sequenceNumber: Number(entry.sequenceNumber) || 1,
    year: Number(entry.year) || new Date().getFullYear(),
    typeCode: code,
    typeLabel,
    nomorSurat: String(entry.nomorSurat || '-').trim(),
    title: String(entry.title || typeLabel).trim(),
    perihal: String(entry.perihal || entry.title || '-').trim(),
    recipientName: String(entry.recipientName || '-').trim(),
    tanggalSurat: String(entry.tanggalSurat || '-').trim(),
    schoolName: school?.name || 'SMK IT Al-Hawari',
    schoolNpsn: school?.npsn || '',
    schoolAddress: [school?.address, school?.district, school?.regency].filter(Boolean).join(', '),
    signerRole: signerInfo.signerRole,
    signerTitle: signerInfo.signerTitle,
    signerName: signerInfo.signerName,
    signerNip: signerInfo.signerNip,
    signerRank: signerInfo.signerRank,
    verificationStatement: signerInfo.verificationStatement,
    details,
    summary
  };
}

/**
 * Clone a docModel and update its letterSignature block for either QR or Manual wet-ink mode.
 * Non-destructive to the caller's live draft object.
 */
export function applySignatureModeToDocModel(docModel, options = {}) {
  if (!docModel) return docModel;
  const cloned = JSON.parse(JSON.stringify(docModel));
  const mode = options.mode === 'qr' && options.qrDataUrl ? 'qr' : 'manual';
  const resolvedRole = options.signerRole || options.signerInfo?.signerRole || 'kepala_sekolah';
  const defaultCaptions = buildQrCaptionLines({
    signerRole: resolvedRole,
    school: options.school || {},
    principalName: options.principalName || options.signerInfo?.principalName || '',
    signerName: options.signerName || options.signerInfo?.signerName || ''
  });

  if (Array.isArray(cloned.blocks)) {
    for (const block of cloned.blocks) {
      if (block?.type === 'letterSignature') {
        block.signatureMode = mode;
        block.qrImage = mode === 'qr' ? options.qrDataUrl : '';
        block.verificationUrl = mode === 'qr' ? (options.verificationUrl || '') : '';
        block.qrCaption = mode === 'qr'
          ? (options.qrCaption || defaultCaptions.line1)
          : '';
        block.qrSubCaption = mode === 'qr'
          ? (options.qrSubCaption !== undefined ? options.qrSubCaption : defaultCaptions.line2)
          : '';
        if (options.signerInfo && options.syncSignerFields) {
          block.title = options.signerInfo.signerTitle || block.title;
          block.name = options.signerInfo.signerName || block.name;
          block.nip = options.signerInfo.signerNip ?? block.nip;
          block.rank = options.signerInfo.signerRank || block.rank;
        }
      }
    }
  }
  return cloned;
}

/**
 * Register a newly authorized SAH outgoing letter on Netlify Functions + Blobs (/api/verify).
 */
export async function registerVerificationInCloud(payload) {
  const controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
  const timer = controller ? setTimeout(() => controller.abort(), 12000) : null;

  try {
    const res = await fetch('/verify-letter', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
      signal: controller?.signal
    });
    const text = await res.text();
    let json = {};
    try { json = text ? JSON.parse(text) : {}; } catch { json = {}; }

    if (!res.ok) {
      return {
        ok: false,
        status: res.status,
        error: json.error || `Gagal menyimpan pengesahan ke cloud (HTTP ${res.status}).`
      };
    }
    return { ok: true, status: res.status, verification: json.verification || null };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: err?.name === 'AbortError'
        ? 'Koneksi ke server verifikasi timeout (tersimpan lokal, coba Sinkron Ulang nanti).'
        : `Server verifikasi cloud tidak terjangkau (${err?.message || 'offline'}).`
    };
  } finally {
    if (timer) clearTimeout(timer);
  }
}

/**
 * Update status (sah <-> revoked) in Netlify Blobs (/api/verify).
 */
export async function updateCloudVerificationStatus({ verificationId, manageToken, status, reason = '' }) {
  const controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
  const timer = controller ? setTimeout(() => controller.abort(), 12000) : null;

  try {
    const res = await fetch('/verify-letter', {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ verificationId, manageToken, status, reason }),
      signal: controller?.signal
    });
    const text = await res.text();
    let json = {};
    try { json = text ? JSON.parse(text) : {}; } catch { json = {}; }

    if (!res.ok) {
      return {
        ok: false,
        status: res.status,
        error: json.error || `Gagal memperbarui status verifikasi (HTTP ${res.status}).`
      };
    }
    return { ok: true, status: res.status, verification: json.verification || null };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      error: `Gagal terhubung ke server verifikasi (${err?.message || 'offline'}).`
    };
  } finally {
    if (timer) clearTimeout(timer);
  }
}

/**
 * Fetch public verification details for a scanned ?v=<verificationId>.
 */
export async function fetchCloudVerification(verificationId) {
  const id = String(verificationId || '').trim();
  if (!id) {
    return { ok: false, status: 400, error: 'Kode verifikasi tidak ditemukan pada URL.' };
  }

  const controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
  const timer = controller ? setTimeout(() => controller.abort(), 12000) : null;

  try {
    const res = await fetch(`/verify-letter?v=${encodeURIComponent(id)}`, {
      method: 'GET',
      cache: 'no-store',
      signal: controller?.signal
    });
    const text = await res.text();
    let json = {};
    try { json = text ? JSON.parse(text) : {}; } catch { json = {}; }

    if (!res.ok) {
      return {
        ok: false,
        status: res.status,
        notFound: res.status === 404,
        error: json.error || `Verifikasi tidak ditemukan (HTTP ${res.status}).`
      };
    }
    return { ok: true, status: res.status, verification: json.verification || null };
  } catch (err) {
    return {
      ok: false,
      status: 0,
      offline: true,
      error: `Tidak dapat menjangkau server verifikasi cloud (${err?.message || 'offline'}).`
    };
  } finally {
    if (timer) clearTimeout(timer);
  }
}
