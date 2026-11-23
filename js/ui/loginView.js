/**
 * js/ui/loginView.js - Layar Autentikasi Staf TU, Kepala Sekolah, & Administrator
 */

import { store } from '../state.js';
import {
  clearActiveWorkspaceUnitId,
  fetchWorkspaceBootstrap,
  loginWorkspace,
  setActiveWorkspaceUnitId,
  setStoredAuthToken
} from '../sync/workspaceClient.js';
import { showToast } from './notification.js';
import { uiIcon } from './icons.js';

function escapeHtml(value) {
  const div = document.createElement('div');
  div.textContent = String(value ?? '');
  return div.innerHTML;
}

export function renderLoginView(rootContainer, { onAuthenticated, statusMessage = '' } = {}) {
  if (!rootContainer) return;

  const existingOverlay = document.getElementById('auth-login-gate');
  if (existingOverlay) existingOverlay.remove();

  const overlay = document.createElement('div');
  overlay.id = 'auth-login-gate';
  overlay.className = 'auth-gate-backdrop';

  overlay.innerHTML = `
    <div class="auth-gate-card">
      <div class="auth-gate-brand">
        <img src="img/logo.png" alt="Logo Al-Hawari" class="auth-gate-logo" />
        <div>
          <div class="auth-gate-title">AI DocBuilder Persuratan</div>
          <div class="auth-gate-subtitle">Yayasan Nunuk Kartini Al-Irsyad • SMP &amp; SMK IT Al-Hawari</div>
        </div>
      </div>

      <div class="auth-role-hints" id="auth-role-hints">
        <button type="button" class="auth-role-pill active" data-username="admin">
          ${uiIcon('shieldCheck', 13)}<span>Admin Yayasan</span>
        </button>
        <button type="button" class="auth-role-pill" data-username="kepsek">
          ${uiIcon('school', 13)}<span>Kepsek SMK</span>
        </button>
        <button type="button" class="auth-role-pill" data-username="tu">
          ${uiIcon('docClean', 13)}<span>TU SMK</span>
        </button>
        <button type="button" class="auth-role-pill" data-username="kepsek_smp">
          ${uiIcon('school', 13)}<span>Kepsek SMP</span>
        </button>
        <button type="button" class="auth-role-pill" data-username="tu_smp">
          ${uiIcon('docClean', 13)}<span>TU SMP</span>
        </button>
      </div>

      ${statusMessage ? `<div class="auth-alert auth-alert-warn">${escapeHtml(statusMessage)}</div>` : ''}
      <div id="auth-login-error" class="auth-alert auth-alert-error hidden"></div>

      <form id="form-workspace-login" class="auth-login-form">
        <div class="form-group">
          <label for="login-username">Username Akun</label>
          <input type="text" id="login-username" class="form-control" value="admin" placeholder="Masukkan username (admin / kepsek / tu / kepsek_smp / tu_smp)" autocomplete="username" required />
        </div>

        <div class="form-group">
          <label for="login-password">Password / PIN Akses</label>
          <div style="position: relative;">
            <input type="password" id="login-password" class="form-control" placeholder="Masukkan password akun..." autocomplete="current-password" required style="padding-right: 76px;" />
            <button type="button" id="btn-toggle-login-pass" class="auth-pass-toggle">Tampilkan</button>
          </div>
        </div>

        <button type="submit" id="btn-login-submit" class="btn btn-primary" style="width: 100%; padding: 11px 16px; font-size: 14px; font-weight: 700; display:inline-flex; align-items:center; justify-content:center; gap:8px;">
          ${uiIcon('key', 15)}<span>Masuk ke Ruang Kerja Akun</span>
        </button>
      </form>

      <div class="auth-gate-footer">
        <div>${uiIcon('lock', 12)} Setiap akun &amp; unit sekolah (SMP / SMK) memiliki ruang kerja &amp; percakapan AI terpisah.</div>
        <div>Scan QR Verifikasi keaslian surat tetap terbuka publik lewat tautan <code>?v=...</code></div>
      </div>
    </div>
  `;

  document.body.appendChild(overlay);

  const form = overlay.querySelector('#form-workspace-login');
  const usernameInput = overlay.querySelector('#login-username');
  const passwordInput = overlay.querySelector('#login-password');
  const toggleBtn = overlay.querySelector('#btn-toggle-login-pass');
  const errorBox = overlay.querySelector('#auth-login-error');
  const submitBtn = overlay.querySelector('#btn-login-submit');
  const pills = overlay.querySelectorAll('.auth-role-pill');

  pills.forEach(pill => {
    pill.addEventListener('click', () => {
      pills.forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
      usernameInput.value = pill.dataset.username || '';
      passwordInput.focus();
    });
  });

  toggleBtn.addEventListener('click', () => {
    const isPass = passwordInput.type === 'password';
    passwordInput.type = isPass ? 'text' : 'password';
    toggleBtn.textContent = isPass ? 'Sembunyikan' : 'Tampilkan';
  });

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const username = usernameInput.value.trim();
    const password = passwordInput.value;
    if (!username || !password) return;

    errorBox.classList.add('hidden');
    submitBtn.disabled = true;
    submitBtn.textContent = 'Memverifikasi akun & menyinkronkan ruang kerja...';

    try {
      clearActiveWorkspaceUnitId();
      const loginRes = await loginWorkspace({ username, password });
      if (!loginRes.ok || !loginRes.token) {
        errorBox.textContent = loginRes.error || 'Gagal masuk. Periksa kembali username dan password.';
        errorBox.classList.remove('hidden');
        return;
      }

      const initialUnitId = loginRes.unitId || (loginRes.user?.unitId === 'smp' ? 'smp' : 'smk');
      setStoredAuthToken(loginRes.token);
      setActiveWorkspaceUnitId(initialUnitId);
      store.setAuthSessionData({
        user: loginRes.user,
        unitId: initialUnitId,
        permissions: loginRes.permissions,
        driveStatus: loginRes.driveStatus
      });

      const bootRes = await fetchWorkspaceBootstrap(initialUnitId);
      if (bootRes.ok) {
        store.applyRemoteBootstrap(bootRes);
      }

      overlay.remove();
      showToast(`Selamat datang, ${loginRes.user?.displayName || loginRes.user?.username}!`, 'success');
      if (typeof onAuthenticated === 'function') {
        onAuthenticated(bootRes.ok ? bootRes : loginRes);
      }
    } finally {
      if (document.body.contains(submitBtn)) {
        submitBtn.disabled = false;
        submitBtn.innerHTML = `${uiIcon('key', 15)}<span>Masuk ke Ruang Kerja Akun</span>`;
      }
    }
  });

  setTimeout(() => passwordInput.focus(), 60);
}

export function hideLoginView() {
  document.getElementById('auth-login-gate')?.remove();
}
