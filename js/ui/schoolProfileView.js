/**
 * schoolProfileView.js - Formulir Pengaturan Profil Sekolah & Gudang Kop Surat
 * Tersimpan permanen di IndexedDB.
 */

import { store } from '../state.js';
import { DEFAULT_SCHOOL_PROFILE } from '../letters/schoolProfile.js';
import { DEFAULT_KOP_DATA_URL } from '../letters/defaultKop.js';
import { showToast } from './notification.js';

let activeSchoolSubTab = 'profil';

const SVG_LAYERS = (size = 30) => `<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="#156c3b" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 12 12 17 22 12"/><polyline points="2 17 12 22 22 17"/></svg>`;
const SVG_BLDG_15 = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="2" width="16" height="20" rx="2"/><path d="M9 22v-4h6v4"/><path d="M8 6h.01M16 6h.01M12 6h.01M12 10h.01M12 14h.01M16 10h.01M16 14h.01M8 10h.01M8 14h.01"/></svg>`;
const SVG_INST_15 = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 21h18"/><path d="M5 21V7l8-4v18"/><path d="M19 21V11l-6-4"/><path d="M9 9v.01M9 12v.01M9 15v.01M9 18v.01"/></svg>`;
const SVG_PHONE_15 = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 22 16.92z"/></svg>`;
const SVG_USER_15 = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>`;
const SVG_HASH_15 = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="4" y1="9" x2="20" y2="9"/><line x1="4" y1="15" x2="20" y2="15"/><line x1="10" y1="3" x2="8" y2="21"/><line x1="16" y1="3" x2="14" y2="21"/></svg>`;
const SVG_BARCODE_15 = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 5v14M8 5v14M12 5v14M17 5v14M21 5v14"/></svg>`;
const SVG_AWARD_15 = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="8" r="6"/><path d="M15.477 12.89 17 22l-5-3-5 3 1.523-9.11"/></svg>`;
const SVG_DOC_15 = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>`;
const SVG_CAP_15 = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 10v6M2 10l10-5 10 5-10 5z"/><path d="M6 12v5c0 2 2 3 6 3s6-1 6-3v-5"/></svg>`;
const SVG_PIN_15 = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>`;
const SVG_MAIL_15 = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="4" width="20" height="16" rx="2"/><polyline points="22,6 12,13 2,6"/></svg>`;
const SVG_GLOBE_16 = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>`;
const SVG_STAR_15 = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#156c3b" stroke-width="2"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>`;
const SVG_IMAGE_15 = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#156c3b" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg>`;
const SVG_GEAR_15 = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#156c3b" stroke-width="2"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 1 1-4 0v-.09a1.65 1.65 0 0 0-1-1.51 1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 1 1 0-4h.09a1.65 1.65 0 0 0 1.51-1 1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 1 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 1 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>`;

function iconInput({ label, name, value, icon, type = 'text', placeholder = '', full = false }) {
  return `
    <div class="form-group" style="margin-bottom: 0; ${full ? 'grid-column: 1 / -1;' : ''}">
      <label style="font-size: 12px; font-weight: 600; color: #334155; margin-bottom: 5px; display: block;">${label}</label>
      <div class="icon-input-wrap">
        <span class="icon-input-leading">${icon}</span>
        <input type="${type}" name="${name}" class="icon-input-control" value="${escapeHtml(value || '')}" placeholder="${escapeHtml(placeholder)}">
      </div>
    </div>
  `;
}

function iconSelect({ label, name, value, icon, options }) {
  const hasCurrent = options.some(opt => opt === value);
  const allOptions = hasCurrent || !value ? options : [value, ...options];
  return `
    <div class="form-group" style="margin-bottom: 0;">
      <label style="font-size: 12px; font-weight: 600; color: #334155; margin-bottom: 5px; display: block;">${label}</label>
      <div class="icon-input-wrap">
        <span class="icon-input-leading">${icon}</span>
        <select name="${name}" class="icon-input-control" style="cursor: pointer;">
          ${allOptions.map(opt => `<option value="${escapeHtml(opt)}" ${opt === value ? 'selected' : ''}>${escapeHtml(opt)}</option>`).join('')}
        </select>
      </div>
    </div>
  `;
}

export function renderSchoolProfileView(container) {
  if (!container) return;

  const school = store.getState().schoolProfile || DEFAULT_SCHOOL_PROFILE;
  const kopGallery = store.getKopGallery();
  const activeKopUrl = school.kopImage || kopGallery[0]?.dataUrl || DEFAULT_KOP_DATA_URL;

  const accreditationOptions = [
    'Swasta / Akreditasi B',
    'Swasta / Akreditasi A',
    'Swasta / Akreditasi C',
    'Negeri / Akreditasi A',
    'Negeri / Akreditasi B'
  ];
  const educationOptions = [
    'SMK (Sekolah Menengah Kejuruan)',
    'SMA (Sekolah Menengah Atas)',
    'MA (Madrasah Aliyah)',
    'SMP (Sekolah Menengah Pertama)'
  ];

  const isVisibleForTab = (tabsCsv) => {
    const allowed = tabsCsv.split(' ');
    return allowed.includes(activeSchoolSubTab) ? '' : 'hidden';
  };

  container.innerHTML = `
    <div class="view-atmosphere">
      <div class="view-shell school-profile-container">

        <!-- HEADER -->
        <div class="view-page-header" style="margin-bottom: 16px;">
          <div class="view-header-lead" style="align-items: center;">
            ${SVG_LAYERS(32)}
            <span class="view-header-divider"></span>
            <div>
              <h2 class="view-title">Profil Sekolah &amp; Pengaturan Kop Surat</h2>
              <p class="view-subtitle">
                Kelola identitas sekolah, instansi pembina, kontak, pejabat, serta format dan tampilan kop surat yang digunakan pada seluruh dokumen dan surat.
              </p>
            </div>
          </div>
        </div>

        <!-- SUB-TAB PILLS -->
        <div style="display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 18px;">
          <button type="button" class="pill-tab btn-school-subtab ${activeSchoolSubTab === 'profil' ? 'active' : ''}" data-subtab="profil">
            ${SVG_BLDG_15} <span>Profil Sekolah</span>
          </button>
          <button type="button" class="pill-tab btn-school-subtab ${activeSchoolSubTab === 'instansi' ? 'active' : ''}" data-subtab="instansi">
            ${SVG_INST_15} <span>Instansi Pembina</span>
          </button>
          <button type="button" class="pill-tab btn-school-subtab ${activeSchoolSubTab === 'kontak' ? 'active' : ''}" data-subtab="kontak">
            ${SVG_PHONE_15} <span>Kontak &amp; Alamat</span>
          </button>
          <button type="button" class="pill-tab btn-school-subtab ${activeSchoolSubTab === 'pejabat' ? 'active' : ''}" data-subtab="pejabat">
            ${SVG_USER_15} <span>Pejabat Penandatangan</span>
          </button>
          <button type="button" class="pill-tab btn-school-subtab ${activeSchoolSubTab === 'penomoran' ? 'active' : ''}" data-subtab="penomoran">
            ${SVG_HASH_15} <span>Format Penomoran</span>
          </button>
        </div>

        <form id="form-school-profile">
          <div class="school-two-col">

            <!-- KOLOM KIRI: FORM DATA SEKOLAH -->
            <div style="display: flex; flex-direction: column; gap: 16px; min-width: 0;">

              <!-- KARTU 1: IDENTITAS SEKOLAH -->
              <div class="ui-card school-subtab-card ${isVisibleForTab('profil')}" style="padding: 18px 20px;">
                <h3 style="font-size: 16px; font-weight: 700; color: #0f172a; display: flex; align-items: center; gap: 8px; margin: 0 0 14px 0;">
                  <span style="color: #156c3b; display: inline-flex;">${SVG_INST_15}</span>
                  <span>Identitas Sekolah</span>
                </h3>

                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap: 12px 14px;">
                  ${iconInput({ label: 'Nama Resmi Sekolah *', name: 'schoolName', value: school.schoolName || 'SMKS IT AL-HAWARI', icon: SVG_BLDG_15 })}
                  ${iconInput({ label: 'NPSN', name: 'npsn', value: school.npsn || '69758993', icon: SVG_BARCODE_15 })}
                  ${iconInput({ label: 'Dinas Pendidikan', name: 'department', value: school.department || 'DINAS PENDIDIKAN', icon: SVG_LAYERS(15) })}
                  ${iconSelect({ label: 'Status &amp; Akreditasi', name: 'accreditation', value: school.accreditation || 'Swasta / Akreditasi B', icon: SVG_AWARD_15, options: accreditationOptions })}
                  ${iconInput({ label: 'SK Operasional &amp; Tanggal', name: 'skOperasional', value: school.skOperasional || '425.11/1136-DISDIK (30 April 2013)', icon: SVG_DOC_15 })}
                  ${iconSelect({ label: 'Jenjang Pendidikan', name: 'educationLevel', value: school.educationLevel || 'SMK (Sekolah Menengah Kejuruan)', icon: SVG_CAP_15, options: educationOptions })}
                </div>

                <!-- KOMPETENSI KEAHLIAN -->
                <div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 10px; padding: 12px 16px; margin-top: 16px;">
                  <div style="font-size: 13px; font-weight: 700; color: #166534; display: flex; align-items: center; gap: 6px; margin-bottom: 6px;">
                    ${SVG_STAR_15}
                    <span style="color: #cbd5e1;">|</span>
                    <span>Kompetensi Keahlian (Jurusan Resmi SMKS IT Al-Hawari)</span>
                  </div>
                  <ul style="margin: 0 0 0 18px; padding: 0; font-size: 12.5px; color: #1e293b; line-height: 1.65;">
                    <li><strong>Multimedia (MM)</strong> <span style="color: #15803d;">•</span> <span style="color: #475569;">Desain Grafis, Animasi, Fotografi, Videografi, Media Digital</span></li>
                    <li><strong>Agribisnis Tanaman Pangan dan Hortikultura (ATPH)</strong> <span style="color: #15803d;">•</span> <span style="color: #475569;">Budi Daya Pangan, Sayuran, Hortikultura Modern</span></li>
                    <li><strong>Desain Pemodelan dan Informasi Bangunan (DPIB)</strong> <span style="color: #15803d;">•</span> <span style="color: #475569;">Gambar Konstruksi, Arsitektur, RAB Bangunan</span></li>
                  </ul>
                </div>
              </div>

              <!-- KARTU 2: ALAMAT & KONTAK SEKOLAH -->
              <div class="ui-card school-subtab-card ${isVisibleForTab('profil kontak')}" style="padding: 18px 20px;">
                <h3 style="font-size: 16px; font-weight: 700; color: #0f172a; display: flex; align-items: center; gap: 8px; margin: 0 0 14px 0;">
                  <span style="color: #156c3b; display: inline-flex;">${SVG_GLOBE_16}</span>
                  <span>Alamat &amp; Kontak Sekolah</span>
                </h3>

                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap: 12px 14px;">
                  ${iconInput({ label: 'Alamat Lengkap', name: 'address', value: school.address || 'Kp. Curug RT 02 / RW 06, Desa Cipareuan, Kec. Cibiuk', icon: SVG_PIN_15, full: true })}
                  ${iconInput({ label: 'Kota / Kabupaten', name: 'city', value: school.city || 'Garut', icon: SVG_BLDG_15 })}
                  ${iconInput({ label: 'Kode Pos', name: 'postalCode', value: school.postalCode || '44193', icon: SVG_HASH_15 })}
                </div>

                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(175px, 1fr)); gap: 12px 14px; margin-top: 12px;">
                  ${iconInput({ label: 'No. Telepon / WhatsApp', name: 'phone', value: school.phone || '+62 895-3266-12155', icon: SVG_PHONE_15 })}
                  ${iconInput({ label: 'Email Resmi', name: 'email', value: school.email || 'smkitalhawari@gmail.com', icon: SVG_MAIL_15, type: 'email' })}
                  ${iconInput({ label: 'Website', name: 'website', value: school.website || 'al-hawari.sch.id', icon: SVG_GLOBE_16 })}
                </div>
              </div>

              <!-- KARTU 3: INSTANSI PEMBINA & YAYASAN -->
              <div class="ui-card school-subtab-card ${isVisibleForTab('instansi')}" style="padding: 18px 20px;">
                <h3 style="font-size: 16px; font-weight: 700; color: #0f172a; display: flex; align-items: center; gap: 8px; margin: 0 0 14px 0;">
                  <span style="color: #156c3b; display: inline-flex;">${SVG_INST_15}</span>
                  <span>Instansi Pembina &amp; Yayasan Naungan</span>
                </h3>

                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap: 12px 14px;">
                  ${iconInput({ label: 'Pemerintah Daerah / Naungan', name: 'government', value: school.government || 'PEMERINTAH DAERAH PROVINSI JAWA BARAT', icon: SVG_INST_15 })}
                  ${iconInput({ label: 'Yayasan Penyelenggara', name: 'foundation', value: school.foundation || 'Yayasan Nunuk Kartini Al-Irsyad', icon: SVG_AWARD_15 })}
                  ${iconInput({ label: 'Cabang Dinas (Opsional)', name: 'subOffice', value: school.subOffice || 'CABANG DINAS PENDIDIKAN WILAYAH XI', icon: SVG_LAYERS(15) })}
                  ${iconInput({ label: 'Provinsi', name: 'province', value: school.province || 'Jawa Barat', icon: SVG_PIN_15 })}
                </div>
              </div>

              <!-- KARTU 4: PEJABAT PENANDATANGAN -->
              <div class="ui-card school-subtab-card ${isVisibleForTab('pejabat')}" style="padding: 18px 20px;">
                <h3 style="font-size: 16px; font-weight: 700; color: #0f172a; display: flex; align-items: center; gap: 8px; margin: 0 0 14px 0;">
                  <span style="color: #156c3b; display: inline-flex;">${SVG_USER_15}</span>
                  <span>Pejabat Penandatangan &amp; Pengesahan (Kepala Sekolah &amp; TU)</span>
                </h3>

                <div style="font-size: 12px; font-weight: 700; color: #15803d; margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.03em;">1. Kepala Sekolah</div>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap: 12px 14px; margin-bottom: 16px;">
                  ${iconInput({ label: 'Nama Lengkap &amp; Gelar', name: 'principalName', value: school.principalName || 'Dian Nurhadian, S.Pd.', icon: SVG_USER_15 })}
                  ${iconInput({ label: "NIP / NUPTK (Isi '-' jika swasta)", name: 'principalNip', value: school.principalNip || '-', icon: SVG_BARCODE_15 })}
                  ${iconInput({ label: 'Pangkat / Keterangan', name: 'principalRank', value: school.principalRank || 'Kepala SMKS IT Al-Hawari', icon: SVG_AWARD_15 })}
                  ${iconInput({ label: 'Jabatan Resmi', name: 'principalTitle', value: school.principalTitle || 'Kepala Sekolah', icon: SVG_CAP_15 })}
                </div>

                <div style="border-top: 1px dashed #e2e8f0; padding-top: 14px; font-size: 12px; font-weight: 700; color: #0369a1; margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.03em;">2. Tata Usaha (TU) — Opsi Pengesahan Barcode QR</div>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap: 12px 14px;">
                  ${iconInput({ label: 'Nama Kepala / Staf TU', name: 'tuName', value: school.tuName || 'Staf Tata Usaha (TU)', icon: SVG_USER_15 })}
                  ${iconInput({ label: "NIP / NUPTK TU (Isi '-' jika swasta)", name: 'tuNip', value: school.tuNip || '-', icon: SVG_BARCODE_15 })}
                  ${iconInput({ label: 'Keterangan Unit TU', name: 'tuRank', value: school.tuRank || 'Tata Usaha SMKS IT Al-Hawari', icon: SVG_AWARD_15 })}
                  ${iconInput({ label: 'Jabatan Resmi TU', name: 'tuTitle', value: school.tuTitle || 'Kepala Tata Usaha', icon: SVG_CAP_15 })}
                </div>
              </div>

              <!-- KARTU 5: POLA PENOMORAN & KODE JENIS SURAT -->
              <div class="ui-card school-subtab-card ${isVisibleForTab('penomoran')}" style="padding: 18px 20px;">
                <h3 style="font-size: 16px; font-weight: 700; color: #0f172a; display: flex; align-items: center; gap: 8px; margin: 0 0 14px 0;">
                  <span style="color: #156c3b; display: inline-flex;">${SVG_HASH_15}</span>
                  <span>Format Penomoran &amp; Kode Jenis Surat</span>
                </h3>

                ${iconInput({ label: 'Pola Format Nomor Surat Otomatis', name: 'numberingPattern', value: school.numberingPattern || '421.3/{urut}/SMKITA/{bulanRomawi}/{tahun}', icon: SVG_HASH_15, full: true })}
                <div style="font-size: 11.5px; color: #64748b; margin-top: 6px;">
                  Placeholder tersedia: <code>{urut}</code>, <code>{kode}</code>, <code>{jenis}</code>, <code>{bulanRomawi}</code>, <code>{tahun}</code>
                </div>

                <div style="margin-top: 12px; background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; padding: 12px 14px; font-size: 12px; color: #14532d; line-height: 1.65;">
                  <strong>Kode Jenis Surat Otomatis (AI Copilot &amp; Lembar Kerja):</strong><br>
                  • <code>UND</code> (Undangan) &nbsp;|&nbsp;
                  • <code>TU</code> (Tata Usaha) &nbsp;|&nbsp;
                  • <code>KEP</code> (Kepala Sekolah / SK) &nbsp;|&nbsp;
                  • <code>KET</code> (Keterangan)<br>
                  • <code>ST</code> (Surat Tugas) &nbsp;|&nbsp;
                  • <code>PGL</code> (Panggilan Ortu) &nbsp;|&nbsp;
                  • <code>PBR</code> (Pemberitahuan) &nbsp;|&nbsp;
                  • <code>PRM</code> (Permohonan)<br>
                  • <code>DISP</code> (Dispensasi) &nbsp;|&nbsp;
                  • <code>BA</code> (Berita Acara) &nbsp;|&nbsp;
                  • <code>PKL</code> (Prakerin) &nbsp;|&nbsp;
                  • <code>KUR</code> (Kurikulum) &nbsp;|&nbsp;
                  • <code>KES</code> (Kesiswaan)
                </div>
              </div>

            </div>

            <!-- KOLOM KANAN: PENGATURAN KOP SURAT -->
            <div class="ui-card" style="padding: 18px 20px; position: sticky; top: 18px;">
              <div style="display: flex; align-items: flex-start; gap: 10px; margin-bottom: 12px;">
                <span style="margin-top: 2px;">${SVG_LAYERS(20)}</span>
                <div>
                  <h3 style="font-size: 16px; font-weight: 700; color: #0f172a; margin: 0;">Pengaturan Kop Surat</h3>
                  <p style="font-size: 12px; color: #64748b; margin: 2px 0 0 0;">
                    Atur tampilan, pilih template, dan kelola data kop surat sekolah.
                  </p>
                </div>
              </div>

              <!-- PRATINJAU KOP AKTIF -->
              <div style="border: 1px solid #cbd5e1; border-radius: 9px; background: #ffffff; padding: 8px; margin-bottom: 14px; box-shadow: inset 0 1px 2px rgba(0,0,0,0.03);">
                <img id="preview-active-kop-img" src="${activeKopUrl}" alt="Pratinjau Kop Surat Aktif" style="width: 100%; max-height: 105px; object-fit: contain; display: block;">
              </div>

              <!-- PILIH TEMPLATE KOP SURAT -->
              <div style="margin-bottom: 14px;">
                <div style="display: flex; align-items: center; gap: 6px; font-size: 12.5px; font-weight: 700; color: #0f172a;">
                  ${SVG_DOC_15}
                  <span>Pilih Template Kop Surat</span>
                </div>
                <div class="kop-template-grid">
                  ${kopGallery.map(k => {
                    const isActive = (school.kopImage || '') === k.dataUrl || (!school.kopImage && k.isDefault);
                    return `
                      <button type="button" class="kop-template-tile btn-apply-school-kop ${isActive ? 'is-active' : ''}" data-id="${k.id}" title="${escapeHtml(k.name)}">
                        ${isActive ? '<span class="kop-check-badge">✓</span>' : ''}
                        <img src="${k.dataUrl}" alt="${escapeHtml(k.name)}" style="width: 100%; height: 42px; object-fit: contain; display: block;">
                        <div style="font-size: 9.5px; font-weight: 600; color: #475569; margin-top: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">
                          ${escapeHtml(k.category || k.name)}
                        </div>
                      </button>
                    `;
                  }).join('')}
                </div>
              </div>

              <!-- OPSI TAMPILAN -->
              <div style="margin-bottom: 16px;">
                <div style="display: flex; align-items: center; gap: 6px; font-size: 12.5px; font-weight: 700; color: #0f172a; margin-bottom: 8px;">
                  ${SVG_GEAR_15}
                  <span>Opsi Tampilan</span>
                </div>
                <div style="display: flex; flex-direction: column; gap: 7px;">
                  <label class="toggle-switch-row">
                    <span style="display: flex; align-items: center; gap: 8px;">${SVG_IMAGE_15} <span>Tampilkan logo kiri (Pemerintah/Dinas)</span></span>
                    <span class="emerald-switch">
                      <input type="checkbox" name="showKopLeftLogo" ${school.showKopLeftLogo !== false ? 'checked' : ''}>
                      <span class="emerald-switch-slider"></span>
                    </span>
                  </label>
                  <label class="toggle-switch-row">
                    <span style="display: flex; align-items: center; gap: 8px;">${SVG_IMAGE_15} <span>Tampilkan logo kanan (Sekolah)</span></span>
                    <span class="emerald-switch">
                      <input type="checkbox" name="showKopRightLogo" ${school.showKopRightLogo !== false ? 'checked' : ''}>
                      <span class="emerald-switch-slider"></span>
                    </span>
                  </label>
                  <label class="toggle-switch-row">
                    <span style="display: flex; align-items: center; gap: 8px;">${SVG_IMAGE_15} <span>Tampilkan alamat, email, dan NPSN</span></span>
                    <span class="emerald-switch">
                      <input type="checkbox" name="showKopContact" ${school.showKopContact !== false ? 'checked' : ''}>
                      <span class="emerald-switch-slider"></span>
                    </span>
                  </label>
                </div>
              </div>

              <!-- AKSI FOOTER -->
              <div style="display: grid; grid-template-columns: 1fr 1.25fr; gap: 10px; padding-top: 4px;">
                <button type="button" id="btn-reset-school-default" class="btn btn-secondary" style="height: 40px; font-weight: 600; font-size: 12.5px; border-radius: 8px;">
                  ↺ Reset ke Default
                </button>
                <button type="submit" id="btn-save-school-profile" class="btn btn-primary" style="height: 40px; font-weight: 700; font-size: 12.5px; border-radius: 8px;">
                  💾 Simpan Perubahan
                </button>
              </div>
            </div>

          </div>
        </form>

      </div>
    </div>
  `;

  // Sub-tab switching without losing unsaved input values in DOM!
  container.querySelectorAll('.btn-school-subtab').forEach(btn => {
    btn.addEventListener('click', () => {
      activeSchoolSubTab = btn.dataset.subtab || 'profil';
      renderSchoolProfileView(container);
    });
  });

  // Bind pilih kop dari galeri
  container.querySelectorAll('.btn-apply-school-kop').forEach(btn => {
    btn.addEventListener('click', () => {
      const id = btn.dataset.id;
      const kop = store.getKopGallery().find(k => k.id === id);
      if (kop) {
        store.updateSchoolProfile({ kopImage: kop.dataUrl });
        showToast(`Kop aktif diatur ke: ${kop.name}`, 'success');
        renderSchoolProfileView(container);
      }
    });
  });

  // Reset ke default
  const btnReset = container.querySelector('#btn-reset-school-default');
  if (btnReset) {
    btnReset.addEventListener('click', () => {
      if (confirm('Kembalikan data profil sekolah dan kop surat ke pengaturan awal bawaan SMKS IT Al-Hawari?')) {
        const defaultCopy = JSON.parse(JSON.stringify(DEFAULT_SCHOOL_PROFILE));
        defaultCopy.kopImage = DEFAULT_KOP_DATA_URL;
        defaultCopy.showKopLeftLogo = true;
        defaultCopy.showKopRightLogo = true;
        defaultCopy.showKopContact = true;
        store.updateSchoolProfile(defaultCopy);
        renderSchoolProfileView(container);
        showToast('Profil sekolah dikembalikan ke pengaturan default.', 'info');
      }
    });
  }

  // Form submit handling
  const form = container.querySelector('#form-school-profile');
  form.onsubmit = (e) => {
    e.preventDefault();
    const formData = new FormData(form);
    const updated = {};
    for (const [key, val] of formData.entries()) {
      if (key === 'showKopLeftLogo' || key === 'showKopRightLogo' || key === 'showKopContact') continue;
      updated[key] = typeof val === 'string' ? val.trim() : val;
    }
    updated.showKopLeftLogo = !!form.querySelector('input[name="showKopLeftLogo"]')?.checked;
    updated.showKopRightLogo = !!form.querySelector('input[name="showKopRightLogo"]')?.checked;
    updated.showKopContact = !!form.querySelector('input[name="showKopContact"]')?.checked;

    store.updateSchoolProfile(updated);
    renderSchoolProfileView(container);
    showToast('Profil sekolah dan kop surat berhasil disimpan permanen.', 'success');
  };
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str || '';
  return div.innerHTML;
}
