/**
 * schoolProfileView.js - Formulir Pengaturan Profil Sekolah & Gudang Kop Surat
 * Tersimpan permanen di IndexedDB.
 */

import { store } from '../state.js';
import { DEFAULT_SCHOOL_PROFILE, getDefaultSchoolProfileForUnit } from '../letters/schoolProfile.js';
import { DEFAULT_KOP_DATA_URL, getDefaultKopGalleryForUnit } from '../letters/defaultKop.js';
import { showToast } from './notification.js';
import { uiIcon } from './icons.js';

let activeSchoolSubTab = 'profil';

const SVG_LAYERS = (size = 24) => uiIcon('layers', size);
const SVG_BLDG_15 = uiIcon('building', 15);
const SVG_INST_15 = uiIcon('institution', 15);
const SVG_PHONE_15 = uiIcon('phone', 15);
const SVG_USER_15 = uiIcon('user', 15);
const SVG_HASH_15 = uiIcon('hash', 15);
const SVG_BARCODE_15 = uiIcon('barcode', 15);
const SVG_AWARD_15 = uiIcon('award', 15);
const SVG_DOC_15 = uiIcon('docClean', 15);
const SVG_CAP_15 = uiIcon('school', 15);
const SVG_PIN_15 = uiIcon('pin', 15);
const SVG_MAIL_15 = uiIcon('mail', 15);
const SVG_GLOBE_16 = uiIcon('globe', 16);
const SVG_STAR_15 = uiIcon('star', 15);
const SVG_IMAGE_15 = uiIcon('gallery', 15);
const SVG_GEAR_15 = uiIcon('settings', 15);

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

  const activeUnitId = store.getActiveUnitId();
  const unitDefault = getDefaultSchoolProfileForUnit(activeUnitId);
  const school = store.getState().schoolProfile || unitDefault;
  const isSmp = activeUnitId === 'smp' || school.unitId === 'smp';
  const kopGallery = store.getKopGallery();
  const activeKopUrl = school.kopImage || kopGallery[0]?.dataUrl || getDefaultKopGalleryForUnit(activeUnitId)[0]?.dataUrl || DEFAULT_KOP_DATA_URL;

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
            <div class="view-header-badge">${SVG_LAYERS(24)}</div>
            <div>
              <h2 class="view-title">Profil Sekolah &amp; Pengaturan Kop Surat</h2>
              <p class="view-subtitle">
                Kelola identitas sekolah, instansi pembina, kontak, pejabat, serta format dan tampilan kop surat yang digunakan pada seluruh dokumen dan surat.
              </p>
            </div>
          </div>
        </div>

        <!-- SUB-TAB PILLS -->
        <div class="filter-pills-bar" style="justify-content: flex-start; margin-bottom: 18px;">
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
                  ${iconInput({ label: 'Nama Resmi Sekolah *', name: 'schoolName', value: school.schoolName || unitDefault.schoolName, icon: SVG_BLDG_15 })}
                  ${iconInput({ label: 'NPSN', name: 'npsn', value: school.npsn || unitDefault.npsn, icon: SVG_BARCODE_15 })}
                  ${iconInput({ label: 'Dinas Pendidikan', name: 'department', value: school.department || unitDefault.department, icon: SVG_LAYERS(15) })}
                  ${iconSelect({ label: 'Status &amp; Akreditasi', name: 'accreditation', value: school.accreditation || unitDefault.accreditation, icon: SVG_AWARD_15, options: accreditationOptions })}
                  ${iconInput({ label: 'SK Operasional &amp; Tanggal', name: 'skOperasional', value: school.skOperasional || unitDefault.skOperasional, icon: SVG_DOC_15 })}
                  ${iconSelect({ label: 'Jenjang Pendidikan', name: 'educationLevel', value: school.educationLevel || unitDefault.educationLevel, icon: SVG_CAP_15, options: educationOptions })}
                </div>

                <!-- KOMPETENSI KEAHLIAN / JENJANG UNIT -->
                <div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 10px; padding: 12px 16px; margin-top: 16px;">
                  <div style="font-size: 13px; font-weight: 700; color: #166534; display: flex; align-items: center; gap: 6px; margin-bottom: 6px;">
                    ${SVG_STAR_15}
                    <span style="color: #cbd5e1;">|</span>
                    <span>${isSmp ? 'Program &amp; Rombel Resmi (SMP IT Al-Hawari - Cibiuk)' : 'Kompetensi Keahlian (Jurusan Resmi SMK IT Al-Hawari)'}</span>
                  </div>
                  ${isSmp ? `
                    <ul style="margin: 0 0 0 18px; padding: 0; font-size: 12.5px; color: #1e293b; line-height: 1.65;">
                      <li><strong>Jenjang SMP (Kelas VII – IX)</strong> <span style="color: #15803d;">•</span> <span style="color: #475569;">Kurikulum Nasional terintegrasi pembiasaan nilai-nilai Islam &amp; tahfiz</span></li>
                      <li><strong>Administrasi &amp; NPSN 20279898</strong> <span style="color: #15803d;">•</span> <span style="color: #475569;">SK Operasional 425.11/794-DISDIK (16 Mei 2019), Yayasan Nunuk Kartini Al-Irsyad (NPYP AA2743)</span></li>
                    </ul>
                  ` : `
                    <ul style="margin: 0 0 0 18px; padding: 0; font-size: 12.5px; color: #1e293b; line-height: 1.65;">
                      <li><strong>Desain Komunikasi Visual (DKV) / Multimedia (MM)</strong> <span style="color: #15803d;">•</span> <span style="color: #475569;">Desain Grafis, Ilustrasi, Fotografi, Videografi, Media Digital</span></li>
                      <li><strong>Agribisnis Tanaman Pangan dan Hortikultura (ATPH)</strong> <span style="color: #15803d;">•</span> <span style="color: #475569;">Budi Daya Pangan, Sayuran, Hortikultura Modern</span></li>
                      <li><strong>Desain Pemodelan dan Informasi Bangunan (DPIB)</strong> <span style="color: #15803d;">•</span> <span style="color: #475569;">Gambar Konstruksi, Arsitektur, RAB Bangunan</span></li>
                    </ul>
                  `}
                </div>
              </div>

              <!-- KARTU 2: ALAMAT & KONTAK SEKOLAH -->
              <div class="ui-card school-subtab-card ${isVisibleForTab('profil kontak')}" style="padding: 18px 20px;">
                <h3 style="font-size: 16px; font-weight: 700; color: #0f172a; display: flex; align-items: center; gap: 8px; margin: 0 0 14px 0;">
                  <span style="color: #156c3b; display: inline-flex;">${SVG_GLOBE_16}</span>
                  <span>Alamat &amp; Kontak Sekolah</span>
                </h3>

                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap: 12px 14px;">
                  ${iconInput({ label: 'Alamat Lengkap', name: 'address', value: school.address || unitDefault.address, icon: SVG_PIN_15, full: true })}
                  ${iconInput({ label: 'Kota / Kabupaten', name: 'city', value: school.city || unitDefault.city, icon: SVG_BLDG_15 })}
                  ${iconInput({ label: 'Kode Pos', name: 'postalCode', value: school.postalCode || unitDefault.postalCode, icon: SVG_HASH_15 })}
                </div>

                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(175px, 1fr)); gap: 12px 14px; margin-top: 12px;">
                  ${iconInput({ label: 'No. Telepon / WhatsApp', name: 'phone', value: school.phone || unitDefault.phone, icon: SVG_PHONE_15 })}
                  ${iconInput({ label: 'Email Resmi', name: 'email', value: school.email || unitDefault.email, icon: SVG_MAIL_15, type: 'email' })}
                  ${iconInput({ label: 'Website', name: 'website', value: school.website || unitDefault.website, icon: SVG_GLOBE_16 })}
                </div>
              </div>

              <!-- KARTU 3: INSTANSI PEMBINA & YAYASAN -->
              <div class="ui-card school-subtab-card ${isVisibleForTab('instansi')}" style="padding: 18px 20px;">
                <h3 style="font-size: 16px; font-weight: 700; color: #0f172a; display: flex; align-items: center; gap: 8px; margin: 0 0 14px 0;">
                  <span style="color: #156c3b; display: inline-flex;">${SVG_INST_15}</span>
                  <span>Instansi Pembina &amp; Yayasan Naungan</span>
                </h3>

                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap: 12px 14px;">
                  ${iconInput({ label: 'Pemerintah Daerah / Naungan', name: 'government', value: school.government || unitDefault.government, icon: SVG_INST_15 })}
                  ${iconInput({ label: 'Yayasan Penyelenggara', name: 'foundation', value: school.foundation || unitDefault.foundation, icon: SVG_AWARD_15 })}
                  ${iconInput({ label: 'Cabang Dinas / Kecamatan (Opsional)', name: 'subOffice', value: school.subOffice ?? unitDefault.subOffice, icon: SVG_LAYERS(15) })}
                  ${iconInput({ label: 'Provinsi', name: 'province', value: school.province || unitDefault.province, icon: SVG_PIN_15 })}
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
                  ${iconInput({ label: 'Nama Lengkap &amp; Gelar', name: 'principalName', value: school.principalName || unitDefault.principalName, icon: SVG_USER_15 })}
                  ${iconInput({ label: "NIP / NUPTK (Isi '-' jika swasta)", name: 'principalNip', value: school.principalNip || '-', icon: SVG_BARCODE_15 })}
                  ${iconInput({ label: 'Pangkat / Keterangan', name: 'principalRank', value: school.principalRank || unitDefault.principalRank, icon: SVG_AWARD_15 })}
                  ${iconInput({ label: 'Jabatan Resmi', name: 'principalTitle', value: school.principalTitle || unitDefault.principalTitle, icon: SVG_CAP_15 })}
                </div>

                <div style="border-top: 1px dashed #e2e8f0; padding-top: 14px; font-size: 12px; font-weight: 700; color: #0369a1; margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.03em;">2. Tata Usaha (TU) — Opsi Pengesahan Barcode QR</div>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); gap: 12px 14px;">
                  ${iconInput({ label: 'Nama Kepala / Staf TU', name: 'tuName', value: school.tuName || unitDefault.tuName, icon: SVG_USER_15 })}
                  ${iconInput({ label: "NIP / NUPTK TU (Isi '-' jika swasta)", name: 'tuNip', value: school.tuNip || '-', icon: SVG_BARCODE_15 })}
                  ${iconInput({ label: 'Keterangan Unit TU', name: 'tuRank', value: school.tuRank || unitDefault.tuRank, icon: SVG_AWARD_15 })}
                  ${iconInput({ label: 'Jabatan Resmi TU', name: 'tuTitle', value: school.tuTitle || unitDefault.tuTitle, icon: SVG_CAP_15 })}
                </div>
              </div>

              <!-- KARTU 5: POLA PENOMORAN & KODE JENIS SURAT -->
              <div class="ui-card school-subtab-card ${isVisibleForTab('penomoran')}" style="padding: 18px 20px;">
                <h3 style="font-size: 16px; font-weight: 700; color: #0f172a; display: flex; align-items: center; gap: 8px; margin: 0 0 14px 0;">
                  <span style="color: #156c3b; display: inline-flex;">${SVG_HASH_15}</span>
                  <span>Format Penomoran &amp; Kode Jenis Surat</span>
                </h3>

                ${iconInput({ label: 'Pola Format Nomor Surat Otomatis', name: 'numberingPattern', value: school.numberingPattern || unitDefault.numberingPattern, icon: SVG_HASH_15, full: true })}
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
                <span class="view-section-icon">${SVG_LAYERS(18)}</span>
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
                        ${isActive ? `<span class="kop-check-badge">${uiIcon('check', 11)}</span>` : ''}
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
                  ${uiIcon('reset', 14)}
                  <span>Reset ke Default</span>
                </button>
                <button type="submit" id="btn-save-school-profile" class="btn btn-primary" style="height: 40px; font-weight: 700; font-size: 12.5px; border-radius: 8px;">
                  ${uiIcon('save', 14)}
                  <span>Simpan Perubahan</span>
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

  // Reset ke default unit aktif (SMP atau SMK)
  const btnReset = container.querySelector('#btn-reset-school-default');
  if (btnReset) {
    btnReset.addEventListener('click', () => {
      const unitLabel = isSmp ? 'SMP IT Al-Hawari - Cibiuk' : 'SMK IT Al-Hawari';
      if (confirm(`Kembalikan data profil sekolah dan kop surat ke pengaturan awal bawaan ${unitLabel}?`)) {
        const defaultCopy = getDefaultSchoolProfileForUnit(activeUnitId);
        defaultCopy.kopImage = getDefaultKopGalleryForUnit(activeUnitId)[0]?.dataUrl || DEFAULT_KOP_DATA_URL;
        defaultCopy.showKopLeftLogo = true;
        defaultCopy.showKopRightLogo = true;
        defaultCopy.showKopContact = true;
        store.updateSchoolProfile(defaultCopy);
        renderSchoolProfileView(container);
        showToast(`Profil ${unitLabel} dikembalikan ke pengaturan default.`, 'info');
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
