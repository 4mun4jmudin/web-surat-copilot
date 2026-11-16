/**
 * js/ui/accessControlView.js - Panel Kontrol Admin, Persetujuan Delegasi Kepala Sekolah,
 * Manajemen Pengguna, dan Koneksi Google Drive (tuancraf1@gmail.com).
 */

import { store } from '../state.js';
import {
  canManageDelegation,
  canManageUsers,
  getRoleLabel,
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

function escapeHtml(value) {
  const div = document.createElement('div');
  div.textContent = String(value ?? '');
  return div.innerHTML;
}

export function renderAccessControlView(container) {
  if (!container) return;

  const state = store.getState();
  const user = state.authUser;
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
            <div class="view-header-badge">🛡️</div>
            <div>
              <h2 class="view-title">
                <span>Hak Akses, Akun Pengguna &amp; Google Drive</span>
                <span class="status-chip ${roleBadgeClass}">${escapeHtml(getRoleLabel(user?.role))}</span>
              </h2>
              <p class="view-subtitle">
                Pengaturan persetujuan delegasi pengesahan surat legal Kepala Sekolah kepada TU, kontrol akun Administrator, serta sinkronisasi arsip ke Google Drive (<code>tuancraf1@gmail.com</code>).
              </p>
            </div>
          </div>
          <button type="button" id="btn-refresh-access-bootstrap" class="btn btn-sm btn-secondary">🔄 Muat Ulang Status Cloud</button>
        </div>

        <!-- KARTU 1: PERSETUJUAN DELEGASI PENGESAHAN KEPALA SEKOLAH -> TU -->
        <div class="ui-card" style="padding: 20px; margin-bottom: 18px;">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:12px; flex-wrap:wrap; margin-bottom:14px;">
            <div>
              <h3 style="margin:0; font-size:16px; color:#0f172a;">✍️ Persetujuan Pengesahan Legal Kepala Sekolah untuk TU</h3>
              <p style="margin:4px 0 0; font-size:12.5px; color:#475569;">
                Mengatur apakah staf <strong>Tata Usaha (TU)</strong> diperbolehkan menerbitkan pengesahan Barcode QR atas nama <strong>Kepala Sekolah</strong>. Pengaturan ini dapat dibuka/tutup oleh <strong>Kepala Sekolah</strong> maupun <strong>Administrator</strong>.
              </p>
            </div>
            <span class="status-chip ${perms.allowTuSignPrincipal ? 'status-chip-green' : 'status-chip-amber'}">
              ${perms.allowTuSignPrincipal ? '✓ Delegasi a.n. Kepsek AKTIF' : '🔒 Delegasi a.n. Kepsek TERKUNCI'}
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
              ${canDelegate ? `<button type="submit" class="btn btn-primary">💾 Simpan Persetujuan Hak Akses</button>` : ''}
            </div>
          </form>
        </div>

        <!-- KARTU 2: KHUSUS ADMINISTRATOR — MANAJEMEN AKUN PENGGUNA -->
        ${isAdmin ? `
          <div class="ui-card" style="padding: 20px; margin-bottom: 18px;">
            <div style="display:flex; justify-content:space-between; align-items:center; gap:12px; flex-wrap:wrap; margin-bottom:14px;">
              <div>
                <h3 style="margin:0; font-size:16px; color:#0f172a;">👥 Manajemen Akun Pengguna (Kontrol Administrator)</h3>
                <p style="margin:4px 0 0; font-size:12.5px; color:#475569;">
                  Tambah, ubah peran (Administrator / Kepala Sekolah / TU), nonaktifkan akun, atau atur ulang password/PIN pengguna.
                </p>
              </div>
              <button type="button" id="btn-open-add-user" class="btn btn-primary btn-sm">+ Tambah Akun Pengguna</button>
            </div>

            <div style="overflow-x:auto;">
              <table style="width:100%; border-collapse:collapse; font-size:13px;">
                <thead>
                  <tr style="background:#f8fafc; border-bottom:1px solid #e2e8f0; text-align:left; color:#334155; font-size:12px;">
                    <th style="padding:10px 12px;">Username</th>
                    <th style="padding:10px 12px;">Nama Tampilan</th>
                    <th style="padding:10px 12px;">Peran</th>
                    <th style="padding:10px 12px;">Status</th>
                    <th style="padding:10px 12px;">Login Terakhir</th>
                    <th style="padding:10px 12px; text-align:right;">Tindakan</th>
                  </tr>
                </thead>
                <tbody>
                  ${managedUsers.map(u => {
                    const isSelf = u.id === user?.id;
                    const rChip = u.role === 'admin' ? 'status-chip-purple' : u.role === 'kepsek' ? 'status-chip-green' : 'status-chip-blue';
                    return `
                      <tr style="border-bottom:1px solid #f1f5f9;">
                        <td style="padding:10px 12px; font-family:Consolas,monospace; font-weight:700;">${escapeHtml(u.username)}${isSelf ? ' <span style="font-size:10px;color:#15803d;">(Anda)</span>' : ''}</td>
                        <td style="padding:10px 12px;">${escapeHtml(u.displayName)}</td>
                        <td style="padding:10px 12px;"><span class="status-chip ${rChip}" style="font-size:11px;">${escapeHtml(getRoleLabel(u.role))}</span></td>
                        <td style="padding:10px 12px;">
                          <span class="status-chip ${u.active !== false ? 'status-chip-green' : 'status-chip-red'}" style="font-size:11px;">
                            ${u.active !== false ? '✓ Aktif' : '⛔ Nonaktif'}
                          </span>
                        </td>
                        <td style="padding:10px 12px; font-size:11.5px; color:#64748b;">
                          ${u.lastLoginAt ? new Date(u.lastLoginAt).toLocaleString('id-ID') : 'Belum pernah'}
                        </td>
                        <td style="padding:10px 12px; text-align:right; white-space:nowrap;">
                          <button type="button" class="btn btn-sm btn-secondary btn-edit-user" data-id="${escapeHtml(u.id)}">✏️ Edit / Password</button>
                          <button type="button" class="btn btn-sm btn-secondary btn-toggle-user-active" data-id="${escapeHtml(u.id)}">${u.active !== false ? 'Nonaktifkan' : 'Aktifkan'}</button>
                          ${!isSelf ? `<button type="button" class="btn btn-sm btn-danger btn-delete-user" data-id="${escapeHtml(u.id)}">🗑️</button>` : ''}
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
          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:12px; flex-wrap:wrap; margin-bottom:12px;">
            <div>
              <h3 style="margin:0; font-size:16px; color:#0f172a;">📁 Penyimpanan &amp; Backup Google Drive (${escapeHtml(drive.accountEmail || 'tuancraf1@gmail.com')})</h3>
              <p style="margin:4px 0 0; font-size:12.5px; color:#475569;">
                Arsip fisik <code>.docx</code> Surat Keluar, template Gudang Surat, aset Kop, dan Google Sheet master otomatis disinkronkan ke folder sekolah Anda.
              </p>
            </div>
            <span class="status-chip ${drive.configured ? 'status-chip-green' : 'status-chip-amber'}">
              ${drive.configured ? '☁️ Bridge Drive Terhubung' : '🕒 Menunggu Webhook Apps Script'}
            </span>
          </div>

          <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:12px 14px; font-size:12.5px; display:grid; gap:6px; margin-bottom:14px;">
            <div><strong>Folder Root Drive:</strong> <a href="${escapeHtml(drive.rootFolderUrl || `https://drive.google.com/drive/folders/${drive.rootFolderId}`)}" target="_blank" rel="noopener">Persuratan SMKS IT Al-Hawari (${escapeHtml(drive.rootFolderId)}) ↗</a></div>
            ${drive.sheetUrl ? `<div><strong>Google Sheet Master:</strong> <a href="${escapeHtml(drive.sheetUrl)}" target="_blank" rel="noopener">Buka DB_Persuratan_SMKS_IT_AlHawari ↗</a></div>` : ''}
            <div><strong>Sinkron Terakhir:</strong> ${drive.lastSyncAt ? new Date(drive.lastSyncAt).toLocaleString('id-ID') : 'Belum pernah'}</div>
            ${drive.lastError ? `<div style="color:#b91c1c;"><strong>Catatan Terakhir:</strong> ${escapeHtml(drive.lastError)}</div>` : ''}
          </div>

          <div style="display:flex; gap:8px; flex-wrap:wrap; margin-bottom:${isAdmin ? '14px' : '0'};">
            <button type="button" id="btn-sync-drive-backup" class="btn btn-primary btn-sm">☁️ Backup &amp; Sync ke Google Drive Sekarang</button>
            <button type="button" id="btn-migrate-local-to-cloud" class="btn btn-secondary btn-sm">📤 Unggah Data Lokal Browser Ini ke Server</button>
          </div>

          ${isAdmin ? `
            <details style="margin-top:8px; border-top:1px dashed #cbd5e1; padding-top:12px;">
              <summary style="cursor:pointer; font-size:12.5px; font-weight:700; color:#334155;">⚙️ Pengaturan Bridge Apps Script Google Drive (Khusus Admin)</summary>
              <form id="form-configure-drive" style="display:grid; gap:10px; margin-top:12px;">
                <div class="form-group" style="margin:0;">
                  <label style="font-size:12px;">URL Web App Google Apps Script (<code>/exec</code>)</label>
                  <input type="url" id="input-drive-webhook-url" class="form-control" placeholder="https://script.google.com/macros/s/.../exec" />
                </div>
                <div class="form-group" style="margin:0;">
                  <label style="font-size:12px;">Secret Bridge Apps Script (<code>GAS_WEBHOOK_SECRET</code>)</label>
                  <input type="password" id="input-drive-webhook-secret" class="form-control" placeholder="Kosongkan jika sudah terpasang di Environment Variable Netlify" />
                </div>
                <div style="display:flex; justify-content:flex-end;">
                  <button type="submit" class="btn btn-sm btn-secondary">🔌 Simpan &amp; Uji Koneksi Drive</button>
                </div>
              </form>
            </details>
          ` : ''}
        </div>

        <!-- KARTU 4: UBAH PASSWORD AKUN SENDIRI -->
        <div class="ui-card" style="padding: 20px;">
          <h3 style="margin:0 0 8px; font-size:15.5px; color:#0f172a;">🔑 Ubah Password / PIN Akun Anda (${escapeHtml(user?.username || '-')})</h3>
          <form id="form-change-own-password" style="display:grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap:10px; align-items:end;">
            <div class="form-group" style="margin:0;">
              <label style="font-size:12px;">Password Lama</label>
              <input type="password" id="input-own-old-pass" class="form-control" required />
            </div>
            <div class="form-group" style="margin:0;">
              <label style="font-size:12px;">Password Baru (min. 6 karakter)</label>
              <input type="password" id="input-own-new-pass" class="form-control" minlength="6" required />
            </div>
            <button type="submit" class="btn btn-secondary" style="height:38px;">Perbarui Password</button>
          </form>
        </div>

      </div>
    </div>

    <!-- MODAL TAMBAH / EDIT AKUN PENGGUNA (ADMIN) -->
    <div id="modal-admin-user" class="modal-overlay hidden">
      <div class="modal-card" style="max-width:460px; width:92%;">
        <div style="display:flex; justify-content:space-between; align-items:center; padding-bottom:10px; border-bottom:1px solid #e2e8f0; margin-bottom:14px;">
          <h3 id="modal-admin-user-title" style="margin:0; font-size:16px;">Tambah Akun Pengguna</h3>
          <button type="button" id="btn-close-admin-user" class="btn btn-sm btn-secondary">&times;</button>
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
    btn.textContent = '⏳ Menyinkronkan ke Drive...';
    try {
      const res = await syncDriveCloud({ mode: 'full_backup' });
      if (res.driveStatus) store.setAuthSessionData({ driveStatus: res.driveStatus });
      if (res.ok) {
        showToast('✓ Backup lengkap & indeks Google Sheet berhasil disinkronkan ke Google Drive!', 'success');
        renderAccessControlView(container);
      } else {
        showToast(res.error || 'Gagal menyinkronkan ke Google Drive.', 'warning');
      }
    } finally {
      if (document.body.contains(btn)) {
        btn.disabled = false;
        btn.textContent = '☁️ Backup & Sync ke Google Drive Sekarang';
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
      showToast(`✓ Data lokal digabungkan ke server (${res.counts?.outgoingLetters || 0} surat keluar, ${res.counts?.apiVault || 0} API AI).`, 'success');
      renderAccessControlView(container);
    } else {
      showToast(res.error || 'Gagal migrasi data lokal.', 'error');
    }
  });

  // 3. Configure Drive Bridge (Admin)
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
          role: roleSelect.value,
          ...(passInput.value.trim() ? { newPassword: passInput.value } : {})
        }
      : {
          op: 'create',
          username: usernameInput.value.trim(),
          displayName: displayInput.value.trim(),
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
