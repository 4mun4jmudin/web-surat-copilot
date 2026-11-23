/**
 * js/ui/accessControlView.js - Panel Kontrol Admin, Persetujuan Delegasi Kepala Sekolah,
 * Manajemen Pengguna, dan Koneksi Google Drive (tuancraf1@gmail.com).
 */

import { store } from '../state.js';
import {
  canManageDelegation,
  canManageUsers,
  getRoleLabel,
  getUnitDisplayName,
  isAdminUser
} from '../auth/permissions.js';
import {
  changeOwnPasswordCloud,
  configureDriveCloud,
  fetchWorkspaceBootstrap,
  manageUsersCloud,
  migrateLocalStateCloud,
  syncDriveCloud,
  updatePermissionsCloud
} from '../sync/workspaceClient.js';
import { showToast } from './notification.js';
import { APPS_SCRIPT_INCOMING_SNIPPET } from '../letters/incoming.js';
import { uiIcon } from './icons.js';

function escapeHtml(value) {
  const div = document.createElement('div');
  div.textContent = String(value ?? '');
  return div.innerHTML;
}

export function renderAccessControlView(container) {
  if (!container) return;

  const state = store.getState();
  const user = state.authUser;
  const activeUnitId = store.getActiveUnitId();
  const activeUnitName = getUnitDisplayName(activeUnitId);
  const perms = state.authPermissions || {};
  const drive = state.driveStatus || {};
  const managedUsers = Array.isArray(state.managedUsers) ? state.managedUsers : [];
  const isAdmin = isAdminUser(user);
  const canDelegate = canManageDelegation(user);

  const roleBadgeClass = user?.role === 'admin'
    ? 'status-chip-purple'
    : user?.role === 'kepsek'
      ? 'status-chip-green'
      : 'status-chip-blue';

  container.innerHTML = `
    <div class="view-atmosphere">
      <div class="view-shell-narrow access-control-container">

        <!-- HEADER -->
        <div class="view-page-header" style="margin-bottom: 18px;">
          <div class="view-header-lead">
            <div class="view-header-badge">
              ${uiIcon('shieldCheck', 22)}
            </div>
            <div>
              <h2 class="view-title">
                <span>Hak Akses, Akun Pengguna &amp; Google Drive</span>
                <span class="status-chip status-chip-green">${escapeHtml(activeUnitName)}</span>
                <span class="status-chip ${roleBadgeClass}">${escapeHtml(getRoleLabel(user?.role))}</span>
              </h2>
              <p class="view-subtitle">
                Pengaturan delegasi pengesahan surat legal unit <strong>${escapeHtml(activeUnitName)}</strong>, kontrol akun multi-unit Yayasan (SMP &amp; SMK), ruang kerja pribadi terenkapsulasi per akun, serta backup Google Drive.
              </p>
            </div>
          </div>
          <div class="view-header-actions">
            <button type="button" id="btn-refresh-access-bootstrap" class="btn btn-sm btn-secondary">
              ${uiIcon('refresh', 13)}
              <span>Muat Ulang Status Cloud</span>
            </button>
          </div>
        </div>

        <!-- KARTU 1: PERSETUJUAN DELEGASI PENGESAHAN KEPALA SEKOLAH -> TU -->
        <div class="ui-card" style="padding: 20px; margin-bottom: 18px;">
          <div class="view-section-head">
            <div class="view-section-title-group" style="align-items:flex-start;">
              <span class="view-section-icon">${uiIcon('penTool', 16)}</span>
              <div>
                <h3 style="margin:0; font-size:16px; color:#0f172a;">Persetujuan Pengesahan Legal Kepala Sekolah untuk TU</h3>
                <p style="margin:4px 0 0; font-size:12.5px; color:#475569;">
                  Mengatur apakah staf <strong>Tata Usaha (TU)</strong> diperbolehkan menerbitkan pengesahan Barcode QR atas nama <strong>Kepala Sekolah</strong>. Pengaturan ini dapat dibuka/tutup oleh <strong>Kepala Sekolah</strong> maupun <strong>Administrator</strong>.
                </p>
              </div>
            </div>
            <span class="status-chip ${perms.allowTuSignPrincipal ? 'status-chip-green' : 'status-chip-amber'}">
              ${perms.allowTuSignPrincipal ? `${uiIcon('check', 12)}<span>Delegasi a.n. Kepsek AKTIF</span>` : `${uiIcon('lock', 12)}<span>Delegasi a.n. Kepsek TERKUNCI</span>`}
            </span>
          </div>

          <form id="form-delegation-permissions" style="display:grid; gap:12px;">
            <label class="access-toggle-card">
              <input type="checkbox" id="chk-perm-tu-principal" ${perms.allowTuSignPrincipal ? 'checked' : ''} ${canDelegate ? '' : 'disabled'} />
              <div>
                <strong>Izinkan TU Mengesahkan Surat Resmi &amp; Barcode QR atas nama Kepala Sekolah</strong>
                <small>Bila aktif, pilihan <em>Kepala Sekolah</em> pada modal Sahkan QR Surat Keluar terbuka bagi akun TU dan tercatat jejak delegasinya.</small>
              </div>
            </label>

            <label class="access-toggle-card">
              <input type="checkbox" id="chk-perm-tu-joint" ${perms.allowTuSignKepsekDanTu ? 'checked' : ''} ${canDelegate ? '' : 'disabled'} />
              <div>
                <strong>Izinkan TU Mengesahkan Barcode QR Bersama (Kepala Sekolah &amp; TU)</strong>
                <small>Membuka opsi pengesahan bersama antara Kepala Sekolah dan Kepala TU untuk akun TU.</small>
              </div>
            </label>

            ${isAdmin ? `
              <div style="border-top: 1px dashed #cbd5e1; padding-top: 12px; margin-top: 2px;">
                <div style="font-size:12px; font-weight:700; color:#475569; text-transform:uppercase; margin-bottom:8px;">Kontrol Eksklusif Administrator untuk Fitur TU</div>
                <div style="display:grid; gap:10px;">
                  <label class="access-toggle-card">
                    <input type="checkbox" id="chk-perm-tu-api" ${perms.allowTuManageApiVault !== false ? 'checked' : ''} />
                    <div>
                      <strong>Izinkan TU Mengelola &amp; Menambah Kunci AI di Gudang API</strong>
                      <small>TU dapat menambahkan API Key AI tanpa batas serta mengganti model AI aktif.</small>
                    </div>
                  </label>
                  <label class="access-toggle-card">
                    <input type="checkbox" id="chk-perm-tu-revoke" ${perms.allowTuRevokePrincipalQr ? 'checked' : ''} />
                    <div>
                      <strong>Izinkan TU Mencabut QR Surat yang Disahkan atas nama Kepala Sekolah</strong>
                      <small>Secara default hanya Administrator &amp; Kepala Sekolah yang berhak mencabut legalitas QR Kepsek.</small>
                    </div>
                  </label>
                  <label class="access-toggle-card">
                    <input type="checkbox" id="chk-perm-tu-delete" ${perms.allowTuDeleteOutgoing ? 'checked' : ''} />
                    <div>
                      <strong>Izinkan TU Menghapus Catatan Surat Keluar Terkunci (Alt+D)</strong>
                      <small>Nonaktifkan agar arsip surat keluar permanen hanya dapat dihapus oleh Admin atau Kepala Sekolah.</small>
                    </div>
                  </label>
                  <label class="access-toggle-card">
                    <input type="checkbox" id="chk-perm-tu-delete-in" ${perms.allowTuDeleteIncoming !== false ? 'checked' : ''} />
                    <div>
                      <strong>Izinkan TU Menghapus Catatan Surat Masuk</strong>
                      <small>Nonaktifkan agar arsip Surat Masuk hanya dapat dihapus oleh Admin atau Kepala Sekolah.</small>
                    </div>
                  </label>
                </div>
              </div>
            ` : ''}

            <div class="form-group" style="margin:0;">
              <label style="font-size:12px; font-weight:600; color:#334155;">Catatan Persetujuan / Keterangan Izin:</label>
              <input type="text" id="input-perm-note" class="form-control" value="${escapeHtml(perms.note || '')}" placeholder="Contoh: Diizinkan selama Kepala Sekolah dinas luar kota..." ${canDelegate ? '' : 'disabled'} />
            </div>

            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
              <div style="font-size:11.5px; color:#64748b;">
                ${perms.updatedAt ? `Pembaruan terakhir: <strong>${escapeHtml(perms.updatedBy || '-')}</strong> (${escapeHtml(getRoleLabel(perms.updatedByRole))}) • ${new Date(perms.updatedAt).toLocaleString('id-ID')}` : 'Menggunakan kebijakan awal sistem'}
              </div>
              ${canDelegate ? `<button type="submit" class="btn btn-primary">${uiIcon('save', 14)}<span>Simpan Persetujuan Hak Akses</span></button>` : ''}
            </div>
          </form>
        </div>

        <!-- KARTU 2: KHUSUS ADMINISTRATOR — MANAJEMEN AKUN PENGGUNA -->
        ${isAdmin ? `
          <div class="ui-card" style="padding: 20px; margin-bottom: 18px;">
            <div class="view-section-head">
              <div class="view-section-title-group" style="align-items:flex-start;">
                <span class="view-section-icon">${uiIcon('users', 16)}</span>
                <div>
                  <h3 style="margin:0; font-size:16px; color:#0f172a;">Manajemen Akun Pengguna (Kontrol Administrator)</h3>
                  <p style="margin:4px 0 0; font-size:12.5px; color:#475569;">
                    Tambah, ubah peran (Administrator / Kepala Sekolah / TU), nonaktifkan akun, atau atur ulang password/PIN pengguna.
                  </p>
                </div>
              </div>
              <button type="button" id="btn-open-add-user" class="btn btn-primary btn-sm">${uiIcon('plus', 13)}<span>Tambah Akun Pengguna</span></button>
            </div>

            <div style="overflow-x:auto; border: 1px solid #e2e8f0; border-radius: 10px;">
              <table class="enterprise-table">
                <thead>
                  <tr>
                    <th>Username</th>
                    <th>Nama Tampilan</th>
                    <th>Unit Sekolah</th>
                    <th>Peran</th>
                    <th>Status</th>
                    <th>Login Terakhir</th>
                    <th style="text-align:right;">Tindakan</th>
                  </tr>
                </thead>
                <tbody>
                  ${managedUsers.map(u => {
                    const isSelf = u.id === user?.id;
                    const rChip = u.role === 'admin' ? 'status-chip-purple' : u.role === 'kepsek' ? 'status-chip-green' : 'status-chip-blue';
                    const uScope = String(u.unitId || 'smk').toLowerCase();
                    const uUnitLabel = uScope === 'all' ? 'Yayasan (SMP & SMK)' : (uScope === 'smp' ? 'SMP IT Al-Hawari' : 'SMK IT Al-Hawari');
                    const uUnitChip = uScope === 'all' ? 'status-chip-purple' : (uScope === 'smp' ? 'status-chip-blue' : 'status-chip-green');
                    return `
                      <tr>
                        <td style="font-family:Consolas,monospace; font-weight:700;">${escapeHtml(u.username)}${isSelf ? ' <span style="font-size:10px;color:#15803d;">(Anda)</span>' : ''}</td>
                        <td>${escapeHtml(u.displayName)}</td>
                        <td><span class="status-chip ${uUnitChip}" style="font-size:11px;">${escapeHtml(uUnitLabel)}</span></td>
                        <td><span class="status-chip ${rChip}" style="font-size:11px;">${escapeHtml(getRoleLabel(u.role))}</span></td>
                        <td>
                          <span class="status-chip ${u.active !== false ? 'status-chip-green' : 'status-chip-red'}" style="font-size:11px;">
                            ${u.active !== false ? `${uiIcon('check', 11)}<span>Aktif</span>` : `${uiIcon('ban', 11)}<span>Nonaktif</span>`}
                          </span>
                        </td>
                        <td style="font-size:11.5px; color:#64748b;">
                          ${u.lastLoginAt ? new Date(u.lastLoginAt).toLocaleString('id-ID') : 'Belum pernah'}
                        </td>
                        <td style="text-align:right; white-space:nowrap;">
                          <div class="table-action-group">
                            <button type="button" class="btn btn-sm btn-secondary btn-edit-user" data-id="${escapeHtml(u.id)}">${uiIcon('edit', 13)}<span>Edit / Password</span></button>
                            <button type="button" class="btn btn-sm btn-secondary btn-toggle-user-active" data-id="${escapeHtml(u.id)}">${u.active !== false ? 'Nonaktifkan' : 'Aktifkan'}</button>
                            ${!isSelf ? `<button type="button" class="btn btn-sm btn-danger btn-square-sm btn-delete-user" data-id="${escapeHtml(u.id)}" title="Hapus Akun">${uiIcon('trash', 13)}</button>` : ''}
                          </div>
                        </td>
                      </tr>
                    `;
                  }).join('')}
                </tbody>
              </table>
            </div>
          </div>
        ` : ''}

        <!-- KARTU 3: GOOGLE DRIVE & MIGRASI DATA -->
        <div class="ui-card" style="padding: 20px; margin-bottom: 18px;">
          <div class="view-section-head">
            <div class="view-section-title-group" style="align-items:flex-start;">
              <span class="view-section-icon">${uiIcon('folder', 16)}</span>
              <div>
                <h3 style="margin:0; font-size:16px; color:#0f172a;">Penyimpanan &amp; Backup Google Drive (${escapeHtml(drive.accountEmail || 'tuancraf1@gmail.com')})</h3>
                <p style="margin:4px 0 0; font-size:12.5px; color:#475569;">
                  Arsip fisik <code>.docx</code> Surat Keluar, template Gudang Surat, aset Kop, dan Google Sheet master otomatis disinkronkan ke folder sekolah Anda.
                </p>
              </div>
            </div>
            <span class="status-chip ${drive.configured ? 'status-chip-green' : 'status-chip-amber'}">
              ${drive.configured ? `${uiIcon('cloudCheck', 12)}<span>Bridge Drive Terhubung</span>` : `${uiIcon('clock', 12)}<span>Menunggu Webhook Apps Script</span>`}
            </span>
          </div>

          <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:12px 14px; font-size:12.5px; display:grid; gap:6px; margin-bottom:14px;">
            <div><strong>Folder Root Drive:</strong> <a href="${escapeHtml(drive.rootFolderUrl || `https://drive.google.com/drive/folders/${drive.rootFolderId}`)}" target="_blank" rel="noopener" style="display:inline-flex;align-items:center;gap:4px;color:#156c3b;font-weight:600;text-decoration:none;"><span>Persuratan Yayasan Al-Hawari — ${escapeHtml(activeUnitName)} (${escapeHtml(drive.rootFolderId)})</span>${uiIcon('externalLink', 12)}</a></div>
            ${drive.sheetUrl ? `<div><strong>Google Sheet Master:</strong> <a href="${escapeHtml(drive.sheetUrl)}" target="_blank" rel="noopener" style="display:inline-flex;align-items:center;gap:4px;color:#156c3b;font-weight:600;text-decoration:none;"><span>Buka DB_Persuratan_AlHawari</span>${uiIcon('externalLink', 12)}</a></div>` : ''}
            <div><strong>Sinkron Terakhir:</strong> ${drive.lastSyncAt ? new Date(drive.lastSyncAt).toLocaleString('id-ID') : 'Belum pernah'}</div>
            ${drive.lastError ? `<div style="color:#b91c1c;"><strong>Catatan Terakhir:</strong> ${escapeHtml(drive.lastError)}</div>` : ''}
          </div>

          <div style="display:flex; gap:8px; flex-wrap:wrap; margin-bottom:${isAdmin ? '14px' : '0'};">
            <button type="button" id="btn-sync-drive-backup" class="btn btn-primary btn-sm">${uiIcon('cloudUpload', 14)}<span>Backup &amp; Sync ke Google Drive Sekarang</span></button>
            <button type="button" id="btn-migrate-local-to-cloud" class="btn btn-secondary btn-sm">${uiIcon('upload', 14)}<span>Unggah Data Lokal Browser Ini ke Server</span></button>
          </div>

          ${isAdmin ? `
            <details style="margin-top:8px; border-top:1px dashed #cbd5e1; padding-top:12px;">
              <summary style="cursor:pointer; font-size:12.5px; font-weight:700; color:#334155; display:inline-flex; align-items:center; gap:6px;">
                ${uiIcon('settings', 14)}
                <span>Pengaturan Bridge Apps Script Google Drive (Khusus Admin)</span>
              </summary>
              <form id="form-configure-drive" style="display:grid; gap:10px; margin-top:12px;">
                <div class="form-group" style="margin:0;">
                  <label style="font-size:12px;">URL Web App Google Apps Script (<code>/exec</code>)</label>
                  <input type="url" id="input-drive-webhook-url" class="form-control" placeholder="https://script.google.com/macros/s/.../exec" />
                </div>
                <div class="form-group" style="margin:0;">
                  <label style="font-size:12px;">Secret Bridge Apps Script (<code>GAS_WEBHOOK_SECRET</code>)</label>
                  <input type="password" id="input-drive-webhook-secret" class="form-control" placeholder="Kosongkan jika sudah terpasang di Environment Variable Netlify" />
                </div>
                <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;">
                  <button type="button" id="btn-copy-gas-incoming-access" class="btn btn-sm btn-secondary">${uiIcon('clipboard', 13)}<span>Salin Handler Surat Masuk (archive_incoming) ke Code.gs</span></button>
                  <button type="submit" class="btn btn-sm btn-secondary">${uiIcon('plug', 13)}<span>Simpan &amp; Uji Koneksi Drive</span></button>
                </div>
              </form>
            </details>
          ` : ''}
        </div>

        <!-- KARTU 4: UBAH PASSWORD AKUN SENDIRI -->
        <div class="ui-card" style="padding: 20px;">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom: 12px;">
            <span class="view-section-icon">${uiIcon('key', 16)}</span>
            <h3 style="margin:0; font-size:15.5px; color:#0f172a;">Ubah Password / PIN Akun Anda (${escapeHtml(user?.username || '-')})</h3>
          </div>
          <form id="form-change-own-password" style="display:grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap:10px; align-items:end;">
            <div class="form-group" style="margin:0;">
              <label style="font-size:12px;">Password Lama</label>
              <input type="password" id="input-own-old-pass" class="form-control" required />
            </div>
            <div class="form-group" style="margin:0;">
              <label style="font-size:12px;">Password Baru (min. 6 karakter)</label>
              <input type="password" id="input-own-new-pass" class="form-control" minlength="6" required />
            </div>
            <button type="submit" class="btn btn-secondary" style="height:38px;">${uiIcon('check', 13)}<span>Perbarui Password</span></button>
          </form>
        </div>

      </div>
    </div>

    <!-- MODAL TAMBAH / EDIT AKUN PENGGUNA (ADMIN) -->
    <div id="modal-admin-user" class="modal-overlay hidden">
      <div class="modal-card" style="max-width:460px; width:92%;">
        <div style="display:flex; justify-content:space-between; align-items:center; padding-bottom:10px; border-bottom:1px solid #e2e8f0; margin-bottom:14px;">
          <h3 id="modal-admin-user-title" style="margin:0; font-size:16px;">Tambah Akun Pengguna</h3>
          <button type="button" id="btn-close-admin-user" class="btn btn-sm btn-secondary btn-square-sm" aria-label="Tutup">${uiIcon('x', 15)}</button>
        </div>
        <form id="form-admin-user" style="display:grid; gap:12px;">
          <input type="hidden" id="input-admin-user-id" value="" />
          <div class="form-group" style="margin:0;">
            <label style="font-size:12px;">Username</label>
            <input type="text" id="input-admin-user-username" class="form-control" placeholder="mis: tu_2" required />
          </div>
          <div class="form-group" style="margin:0;">
            <label style="font-size:12px;">Nama Lengkap / Jabatan</label>
            <input type="text" id="input-admin-user-display" class="form-control" placeholder="mis: Operator TU 2" required />
          </div>
          <div class="form-group" style="margin:0;">
            <label style="font-size:12px;">Unit Sekolah</label>
            <select id="select-admin-user-unit" class="form-control">
              <option value="smk">SMK IT Al-Hawari</option>
              <option value="smp">SMP IT Al-Hawari - Cibiuk</option>
              <option value="all">Yayasan / Semua Unit (Khusus Admin)</option>
            </select>
          </div>
          <div class="form-group" style="margin:0;">
            <label style="font-size:12px;">Peran Akses</label>
            <select id="select-admin-user-role" class="form-control">
              <option value="tu">Tata Usaha (TU)</option>
              <option value="kepsek">Kepala Sekolah</option>
              <option value="admin">Administrator (Akses Penuh)</option>
            </select>
          </div>
          <div class="form-group" style="margin:0;">
            <label style="font-size:12px;" id="label-admin-user-pass">Password / PIN (minimal 6 karakter)</label>
            <input type="password" id="input-admin-user-pass" class="form-control" placeholder="Masukkan password..." />
          </div>
          <div style="display:flex; justify-content:flex-end; gap:8px; margin-top:6px;">
            <button type="button" id="btn-cancel-admin-user" class="btn btn-secondary">Batal</button>
            <button type="submit" class="btn btn-primary">Simpan Akun</button>
          </div>
        </form>
      </div>
    </div>
  `;

  bindAccessControlEvents(container);
}

function bindAccessControlEvents(container) {
  const isAdmin = isAdminUser(store.getState().authUser);

  container.querySelector('#btn-refresh-access-bootstrap')?.addEventListener('click', async () => {
    const res = await fetchWorkspaceBootstrap();
    if (res.ok) {
      store.applyRemoteBootstrap(res);
      showToast('Status hak akses, akun, & Drive diperbarui dari server.', 'success');
      renderAccessControlView(container);
    } else {
      showToast(res.error || 'Gagal memuat ulang dari server.', 'error');
    }
  });

  // 1. Form Delegasi Persetujuan Pengesahan Kepsek -> TU
  container.querySelector('#form-delegation-permissions')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const patch = {
      allowTuSignPrincipal: Boolean(container.querySelector('#chk-perm-tu-principal')?.checked),
      allowTuSignKepsekDanTu: Boolean(container.querySelector('#chk-perm-tu-joint')?.checked),
      note: container.querySelector('#input-perm-note')?.value?.trim() || ''
    };
    if (isAdmin) {
      patch.allowTuManageApiVault = Boolean(container.querySelector('#chk-perm-tu-api')?.checked);
      patch.allowTuRevokePrincipalQr = Boolean(container.querySelector('#chk-perm-tu-revoke')?.checked);
      patch.allowTuDeleteOutgoing = Boolean(container.querySelector('#chk-perm-tu-delete')?.checked);
      patch.allowTuDeleteIncoming = Boolean(container.querySelector('#chk-perm-tu-delete-in')?.checked);
    }

    const res = await updatePermissionsCloud(patch);
    if (res.ok && res.permissions) {
      store.setAuthSessionData({ permissions: res.permissions });
      showToast('✓ Persetujuan hak akses berhasil disimpan ke server!', 'success');
      renderAccessControlView(container);
    } else {
      showToast(res.error || 'Gagal menyimpan hak akses.', 'error');
    }
  });

  // 2. Backup & Sync Drive + Migrasi Lokal
  container.querySelector('#btn-sync-drive-backup')?.addEventListener('click', async (e) => {
    const btn = e.currentTarget;
    btn.disabled = true;
    btn.innerHTML = `${uiIcon('loader', 14, 'spin-icon')}<span>Menyinkronkan ke Drive...</span>`;
    try {
      const res = await syncDriveCloud({ mode: 'full_backup' });
      if (res.driveStatus) store.setAuthSessionData({ driveStatus: res.driveStatus });
      if (res.ok) {
        showToast('Backup lengkap & indeks Google Sheet berhasil disinkronkan ke Google Drive!', 'success');
        renderAccessControlView(container);
      } else {
        showToast(res.error || 'Gagal menyinkronkan ke Google Drive.', 'warning');
      }
    } finally {
      if (document.body.contains(btn)) {
        btn.disabled = false;
        btn.innerHTML = `${uiIcon('cloudUpload', 14)}<span>Backup &amp; Sync ke Google Drive Sekarang</span>`;
      }
    }
  });

  container.querySelector('#btn-migrate-local-to-cloud')?.addEventListener('click', async () => {
    const st = store.getState();
    const res = await migrateLocalStateCloud({
      outgoingLetters: st.outgoingLetters || [],
      letterVault: st.letterVault || [],
      kopGallery: st.kopGallery || [],
      apiVault: st.apiVault || [],
      activeApiId: st.activeApiId || '',
      schoolProfile: st.schoolProfile || null
    });
    if (res.ok) {
      const boot = await fetchWorkspaceBootstrap();
      if (boot.ok) store.applyRemoteBootstrap(boot);
      showToast(`Data lokal digabungkan ke server (${res.counts?.outgoingLetters || 0} surat keluar, ${res.counts?.apiVault || 0} API AI).`, 'success');
      renderAccessControlView(container);
    } else {
      showToast(res.error || 'Gagal migrasi data lokal.', 'error');
    }
  });

  // 3. Configure Drive Bridge (Admin)
  container.querySelector('#btn-copy-gas-incoming-access')?.addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(APPS_SCRIPT_INCOMING_SNIPPET);
      showToast('Kode handler Google Apps Script Surat Masuk (archive_incoming) berhasil disalin!', 'success');
    } catch {
      showToast('Buka tab Surat Masuk → Script Drive untuk menyalin kode manual.', 'info');
    }
  });

  container.querySelector('#form-configure-drive')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const webhookUrl = container.querySelector('#input-drive-webhook-url')?.value?.trim() || '';
    const webhookSecret = container.querySelector('#input-drive-webhook-secret')?.value?.trim() || '';
    const res = await configureDriveCloud({
      ...(webhookUrl ? { webhookUrl } : {}),
      ...(webhookSecret ? { webhookSecret } : {})
    });
    if (res.driveStatus) store.setAuthSessionData({ driveStatus: res.driveStatus });
    if (res.ok) {
      showToast('✓ Koneksi Google Drive Apps Script berhasil terhubung!', 'success');
      renderAccessControlView(container);
    } else {
      showToast(res.error || 'Bridge Drive belum merespons.', 'error');
    }
  });

  // 4. Ubah Password Sendiri
  container.querySelector('#form-change-own-password')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const oldPassword = container.querySelector('#input-own-old-pass')?.value || '';
    const newPassword = container.querySelector('#input-own-new-pass')?.value || '';
    const res = await changeOwnPasswordCloud({ oldPassword, newPassword });
    if (res.ok) {
      showToast('✓ Password akun Anda berhasil diperbarui.', 'success');
      e.currentTarget.reset();
    } else {
      showToast(res.error || 'Gagal memperbarui password.', 'error');
    }
  });

  // 5. Manajemen Akun (Admin)
  if (!isAdmin) return;

  const userModal = container.querySelector('#modal-admin-user');
  const userForm = container.querySelector('#form-admin-user');
  const idInput = container.querySelector('#input-admin-user-id');
  const usernameInput = container.querySelector('#input-admin-user-username');
  const displayInput = container.querySelector('#input-admin-user-display');
  const unitSelect = container.querySelector('#select-admin-user-unit');
  const roleSelect = container.querySelector('#select-admin-user-role');
  const passInput = container.querySelector('#input-admin-user-pass');
  const passLabel = container.querySelector('#label-admin-user-pass');
  const modalTitle = container.querySelector('#modal-admin-user-title');

  const closeUserModal = () => userModal?.classList.add('hidden');
  container.querySelector('#btn-close-admin-user')?.addEventListener('click', closeUserModal);
  container.querySelector('#btn-cancel-admin-user')?.addEventListener('click', closeUserModal);

  container.querySelector('#btn-open-add-user')?.addEventListener('click', () => {
    modalTitle.textContent = 'Tambah Akun Pengguna Baru';
    idInput.value = '';
    usernameInput.value = '';
    usernameInput.disabled = false;
    displayInput.value = '';
    if (unitSelect) unitSelect.value = store.getActiveUnitId() || 'smk';
    roleSelect.value = 'tu';
    passInput.value = '';
    passInput.required = true;
    passLabel.textContent = 'Password / PIN Awal (minimal 6 karakter)';
    userModal.classList.remove('hidden');
  });

  container.querySelectorAll('.btn-edit-user').forEach(btn => {
    btn.addEventListener('click', () => {
      const target = (store.getState().managedUsers || []).find(u => u.id === btn.dataset.id);
      if (!target) return;
      modalTitle.textContent = `Edit Akun: ${target.username}`;
      idInput.value = target.id;
      usernameInput.value = target.username;
      usernameInput.disabled = true;
      displayInput.value = target.displayName || '';
      if (unitSelect) unitSelect.value = target.unitId || (target.role === 'admin' ? 'all' : 'smk');
      roleSelect.value = target.role || 'tu';
      passInput.value = '';
      passInput.required = false;
      passLabel.textContent = 'Password Baru (kosongkan jika tidak ingin ganti)';
      userModal.classList.remove('hidden');
    });
  });

  userForm?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const targetId = idInput.value.trim();
    const payload = targetId
      ? {
          op: 'update',
          targetUserId: targetId,
          displayName: displayInput.value.trim(),
          assignedUnitId: unitSelect?.value || 'smk',
          role: roleSelect.value,
          ...(passInput.value.trim() ? { newPassword: passInput.value } : {})
        }
      : {
          op: 'create',
          username: usernameInput.value.trim(),
          displayName: displayInput.value.trim(),
          assignedUnitId: unitSelect?.value || 'smk',
          role: roleSelect.value,
          password: passInput.value
        };

    const res = await manageUsersCloud(payload);
    if (res.ok && Array.isArray(res.users)) {
      store.setAuthSessionData({ users: res.users });
      closeUserModal();
      showToast(targetId ? 'Akun pengguna berhasil diperbarui.' : 'Akun pengguna baru berhasil dibuat!', 'success');
      renderAccessControlView(container);
    } else {
      showToast(res.error || 'Gagal menyimpan akun pengguna.', 'error');
    }
  });

  container.querySelectorAll('.btn-toggle-user-active').forEach(btn => {
    btn.addEventListener('click', async () => {
      const target = (store.getState().managedUsers || []).find(u => u.id === btn.dataset.id);
      if (!target) return;
      const res = await manageUsersCloud({
        op: 'update',
        targetUserId: target.id,
        active: target.active === false
      });
      if (res.ok && Array.isArray(res.users)) {
        store.setAuthSessionData({ users: res.users });
        showToast(`Status akun ${target.username} diperbarui.`, 'success');
        renderAccessControlView(container);
      } else {
        showToast(res.error || 'Gagal mengubah status akun.', 'error');
      }
    });
  });

  container.querySelectorAll('.btn-delete-user').forEach(btn => {
    btn.addEventListener('click', async () => {
      const target = (store.getState().managedUsers || []).find(u => u.id === btn.dataset.id);
      if (!target) return;
      if (!confirm(`Hapus akun "${target.username}" (${target.displayName})?`)) return;
      const res = await manageUsersCloud({ op: 'delete', targetUserId: target.id });
      if (res.ok && Array.isArray(res.users)) {
        store.setAuthSessionData({ users: res.users });
        showToast(`Akun ${target.username} dihapus.`, 'info');
        renderAccessControlView(container);
      } else {
        showToast(res.error || 'Gagal menghapus akun.', 'error');
      }
    });
  });
}
