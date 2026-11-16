/**
 * Public QR Digital Verification Portal View (?v=<verificationId>)
 *
 * Displayed cleanly on mobile/desktop when someone scans the printed QR Code
 * on an authorized SMK IT Al-Hawari outgoing letter.
 * Fetches live verification status (SAH / REVOKED / NOT FOUND) from Netlify Functions + Blobs,
 * and presents structured details tailored to the exact letter typeCode plus
 * the chosen authorizing authority statement (Kepala Sekolah / Tata Usaha / Keduanya).
 */

import { store } from '../state.js';
import { DEFAULT_SCHOOL_PROFILE } from '../letters/schoolProfile.js';
import { fetchCloudVerification, getTypeLabel } from '../letters/verification.js';

function esc(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatDateTimeId(isoString) {
  if (!isoString) return '-';
  try {
    const d = new Date(isoString);
    if (Number.isNaN(d.getTime())) return String(isoString);
    return d.toLocaleString('id-ID', {
      day: 'numeric',
      month: 'long',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    }) + ' WIB';
  } catch {
    return String(isoString);
  }
}

function getSignerRoleBadge(role) {
  if (role === 'tata_usaha') return 'Disahkan Atas Nama Tata Usaha (TU)';
  if (role === 'kepsek_dan_tu') return 'Disahkan Atas Nama Kepala Sekolah & Tata Usaha (TU)';
  return 'Disahkan Atas Nama Kepala Sekolah';
}

function buildVerificationCardHtml(v, { isLocalPreview = false } = {}) {
  const isSah = v.status === 'sah';
  const isRevoked = v.status === 'revoked';
  const logoSrc = DEFAULT_SCHOOL_PROFILE.logo || store.getState().schoolProfile?.logo || '';
  const typeLabel = v.typeLabel || getTypeLabel(v.typeCode);
  const details = Array.isArray(v.details) && v.details.length > 0
    ? v.details
    : [
        { label: 'Jenis Surat', value: `${typeLabel} (${v.typeCode || 'TU'})` },
        { label: 'Nomor Surat Resmi', value: v.nomorSurat || '-' },
        { label: 'Judul Surat', value: v.title || '-' },
        { label: 'Perihal', value: v.perihal || '-' },
        { label: 'Tujuan / Penerima', value: v.recipientName || '-' },
        { label: 'Tanggal Surat', value: v.tanggalSurat || '-' }
      ];

  const statusHtml = isSah
    ? `
      <div class="verify-status-hero verify-status-sah">
        <div class="verify-status-icon">✓</div>
        <div>
          <div class="verify-status-kicker">VERIFIKASI DIGITAL BERHASIL</div>
          <h2>DOKUMEN ASLI &amp; DINYATAKAN SAH</h2>
          <p>Surat keluar ini terdaftar resmi dalam sistem persuratan ${esc(v.schoolName || 'SMK IT Al-Hawari')}.</p>
        </div>
      </div>
    `
    : isRevoked
    ? `
      <div class="verify-status-hero verify-status-revoked">
        <div class="verify-status-icon">⛔</div>
        <div>
          <div class="verify-status-kicker">PERINGATAN KEABSAHAN</div>
          <h2>PENGESAHAN TELAH DICABUT</h2>
          <p>Surat ini sebelumnya terdaftar, namun status pengesahannya telah dibatalkan pada ${esc(formatDateTimeId(v.revokedAt))}${v.revokedReason ? ` (${esc(v.revokedReason)})` : ''}.</p>
        </div>
      </div>
    `
    : `
      <div class="verify-status-hero verify-status-warn">
        <div class="verify-status-icon">⚠️</div>
        <div>
          <div class="verify-status-kicker">STATUS LOKAL</div>
          <h2>BELUM DISINKRONKAN KE CLOUD</h2>
          <p>Data surat ini baru tersimpan pada perangkat lokal dan belum tersinkron ke server verifikasi publik.</p>
        </div>
      </div>
    `;

  return `
    <div class="verify-card">
      <div class="verify-school-banner">
        ${logoSrc ? `<img src="${esc(logoSrc)}" alt="Logo Sekolah" class="verify-school-logo" />` : '<div class="verify-school-logo-fallback">🏫</div>'}
        <div class="verify-school-meta">
          <div class="verify-portal-tag">PORTAL VERIFIKASI SURAT RESMI</div>
          <h1>${esc(v.schoolName || 'SMK IT Al-Hawari')}</h1>
          <p>${v.schoolNpsn ? `NPSN: ${esc(v.schoolNpsn)} • ` : ''}${esc(v.schoolAddress || 'Cibiuk, Kabupaten Garut')}</p>
        </div>
      </div>

      ${isLocalPreview ? `
        <div class="verify-local-notice">
          ℹ️ Pratinjau dari Arsip Lokal Perangkat (pastikan tombol <strong>Sinkron Cloud</strong> sudah aktif sebelum surat diedarkan agar dapat dipindai dari HP mana saja).
        </div>
      ` : ''}

      ${statusHtml}

      <div class="verify-section">
        <div class="verify-section-header">
          <span>Pernyataan Pengesahan Pejabat</span>
          <span class="verify-role-pill">${esc(getSignerRoleBadge(v.signerRole))}</span>
        </div>
        <div class="verify-authority-box">
          <p class="verify-statement-text">${esc(v.verificationStatement || `Surat ini disahkan atas nama ${v.signerTitle || 'Kepala Sekolah'} (${v.signerName || '-'}).`)}</p>
          <div class="verify-signer-grid">
            <div>
              <span class="verify-field-label">Pejabat Pengesahan</span>
              <strong>${esc(v.signerName || '-')}</strong>
            </div>
            <div>
              <span class="verify-field-label">Jabatan Resmi</span>
              <strong>${esc(v.signerTitle || 'Kepala Sekolah')}</strong>
            </div>
            <div>
              <span class="verify-field-label">NIP / NUPTK</span>
              <strong>${esc(v.signerNip && v.signerNip !== '-' ? v.signerNip : '-')}</strong>
            </div>
            <div>
              <span class="verify-field-label">Waktu Pengesahan</span>
              <strong>${esc(formatDateTimeId(v.issuedAt || v.authorizedAt))}</strong>
            </div>
          </div>
        </div>
      </div>

      <div class="verify-section">
        <div class="verify-section-header">
          <span>Detil Informasi Dokumen</span>
          <span class="verify-type-pill">${esc(v.typeCode || 'TU')} • ${esc(typeLabel)}</span>
        </div>
        <div class="verify-details-list">
          ${details.map(row => `
            <div class="verify-detail-row">
              <div class="verify-detail-label">${esc(row.label)}</div>
              <div class="verify-detail-value">${esc(row.value)}</div>
            </div>
          `).join('')}
          <div class="verify-detail-row">
            <div class="verify-detail-label">Kode Registrasi QR</div>
            <div class="verify-detail-value verify-mono">${esc(v.verificationId)}</div>
          </div>
        </div>
      </div>

      <div class="verify-footer">
        <span>Sistem Persuratan Digital ${esc(v.schoolName || 'SMK IT Al-Hawari')} • Terverifikasi melalui Barcode QR</span>
        <a href="${esc(window.location.pathname || '/')}" class="verify-admin-link">Masuk ke Portal Admin TU →</a>
      </div>
    </div>
  `;
}

export async function renderQrVerificationView(container, verificationId, { localFallbackOnly = false, isInlinePreview = false, onClosePreview = null } = {}) {
  if (!isInlinePreview) {
    document.body.classList.add('verify-mode');
  }
  const cleanId = String(verificationId || '').trim();

  container.innerHTML = `
    <div class="${isInlinePreview ? 'verify-inline-root' : 'verify-scroll-root'}">
      <div class="verify-page-shell">
        <div class="verify-loading-card">
          <div class="verify-spinner"></div>
          <h3>Memeriksa Keaslian Dokumen...</h3>
          <p>Menghubungi catatan resmi Barcode QR <code>${esc(cleanId)}</code></p>
        </div>
      </div>
    </div>
  `;

  const shell = container.querySelector('.verify-page-shell');

  if (!cleanId) {
    shell.innerHTML = `
      <div class="verify-card">
        <div class="verify-status-hero verify-status-invalid">
          <div class="verify-status-icon">❌</div>
          <div>
            <div class="verify-status-kicker">KODE TIDAK VALID</div>
            <h2>KODE VERIFIKASI KOSONG</h2>
            <p>Parameter verifikasi <code>?v=</code> tidak ditemukan pada tautan QR yang dipindai.</p>
          </div>
        </div>
        <div class="verify-footer">
          <a href="${esc(window.location.pathname || '/')}" class="verify-admin-link">Buka Aplikasi Persuratan →</a>
        </div>
      </div>
    `;
    return;
  }

  let cloudResult = { ok: false };
  if (!localFallbackOnly) {
    cloudResult = await fetchCloudVerification(cleanId);
  }

  if (cloudResult.ok && cloudResult.verification) {
    shell.innerHTML = buildVerificationCardHtml(cloudResult.verification, { isLocalPreview: false });
    return;
  }

  // Fallback check on local device IndexedDB (wait briefly if DB still loading)
  let localEntry = store.getOutgoingLetterByVerificationId(cleanId);
  if (!localEntry) {
    await new Promise(r => setTimeout(r, 450));
    localEntry = store.getOutgoingLetterByVerificationId(cleanId);
  }

  if (localEntry) {
    const school = store.getState().schoolProfile || DEFAULT_SCHOOL_PROFILE;
    const localRecord = {
      verificationId: localEntry.verificationId,
      status: localEntry.cloudSynced ? (localEntry.verificationStatus || 'sah') : (localEntry.verificationStatus === 'revoked' ? 'revoked' : 'local_only'),
      typeCode: localEntry.typeCode,
      typeLabel: getTypeLabel(localEntry.typeCode),
      nomorSurat: localEntry.nomorSurat,
      title: localEntry.title,
      perihal: localEntry.perihal,
      recipientName: localEntry.recipientName,
      tanggalSurat: localEntry.tanggalSurat,
      schoolName: school.name,
      schoolNpsn: school.npsn,
      schoolAddress: [school.address, school.district, school.regency].filter(Boolean).join(', '),
      signerRole: localEntry.signerRole,
      signerTitle: localEntry.signerTitle,
      signerName: localEntry.signerName,
      signerNip: localEntry.signerNip,
      verificationStatement: localEntry.verificationStatement,
      details: localEntry.verificationDetails,
      issuedAt: localEntry.authorizedAt || localEntry.recordedAt,
      revokedAt: localEntry.revokedAt,
      revokedReason: localEntry.revokedReason
    };
    shell.innerHTML = buildVerificationCardHtml(localRecord, { isLocalPreview: !localEntry.cloudSynced || localFallbackOnly });
    return;
  }

  const isOfflineOrServerUnavailable = cloudResult.offline || cloudResult.status === 0 || cloudResult.status >= 500;
  shell.innerHTML = `
    <div class="verify-card">
      <div class="verify-school-banner">
        <div class="verify-school-meta">
          <div class="verify-portal-tag">PORTAL VERIFIKASI SURAT RESMI</div>
          <h1>${esc(DEFAULT_SCHOOL_PROFILE.name)}</h1>
        </div>
      </div>
      <div class="verify-status-hero ${isOfflineOrServerUnavailable ? 'verify-status-warn' : 'verify-status-invalid'}">
        <div class="verify-status-icon">${isOfflineOrServerUnavailable ? '📡' : '❌'}</div>
        <div>
          <div class="verify-status-kicker">${isOfflineOrServerUnavailable ? 'GANGGUAN KONEKSI' : 'VERIFIKASI DITOLAK'}</div>
          <h2>${isOfflineOrServerUnavailable ? 'SERVER VERIFIKASI TIDAK TERJANGKAU' : 'DOKUMEN TIDAK TERDAFTAR / TIDAK SAH'}</h2>
          <p>${isOfflineOrServerUnavailable
            ? esc(cloudResult.error || 'Tidak dapat terhubung ke server verifikasi. Pastikan perangkat terhubung ke internet.')
            : `Kode Barcode QR <strong>${esc(cleanId)}</strong> tidak ditemukan dalam arsip Surat Keluar resmi sekolah.`}</p>
        </div>
      </div>
      <div class="verify-footer">
        <span>Kode dicek: <code class="verify-mono">${esc(cleanId)}</code></span>
        <a href="${esc(window.location.pathname || '/')}" class="verify-admin-link">Buka Sistem Persuratan →</a>
      </div>
    </div>
  `;
}
