/**
 * main.js - Bootstrap Aplikasi Word Processor & AI Copilot Persuratan Sekolah
 */

import { store } from './state.js';
import {
  renderWordCanvas,
  triggerWordCanvasPrintWithChoice,
  triggerWordCanvasDownloadWithChoice
} from './ui/wordCanvas.js';
import {
  renderSmartWordCanvas,
  getActiveSmartDocModel,
  triggerSmartCanvasPrintWithChoice,
  triggerSmartCanvasDownloadWithChoice
} from './ui/smartWordCanvas.js';
import { renderOutgoingLettersView } from './ui/outgoingLettersView.js';
import { renderVaultView } from './ui/vaultView.js';
import { renderSchoolProfileView } from './ui/schoolProfileView.js';
import { renderSettingsView } from './ui/settingsView.js';
import { renderAccessControlView } from './ui/accessControlView.js';
import { renderAccApprovalView } from './ui/accApprovalView.js';
import {
  renderLiveAccTray,
  recordCompletedAccForTray,
  resetAccTrayNotifications
} from './ui/accQueueWidget.js';
import { initLiveSync, startLiveSync, stopLiveSync } from './sync/liveSync.js';
import { renderLoginView, hideLoginView } from './ui/loginView.js';
import { renderQrVerificationView } from './ui/verifyView.js';
import { renderDocx } from './docx/render.js';
import { downloadFile } from './utils/download.js';
import { showToast } from './ui/notification.js';
import { runSafeWindowPrint } from './ui/printChoiceModal.js';
import {
  getPendingAccLetters,
  getRoleLabel,
  isCanvasLockedForUser
} from './auth/permissions.js';
import {
  clearStoredAuthToken,
  fetchWorkspaceBootstrap,
  getStoredAuthToken,
  onWorkspaceUnauthorized
} from './sync/workspaceClient.js';
import { initPwa } from './pwa.js';

const CACHED_USER_KEY = 'ai_docbuilder_cached_user';

function saveCachedUserSummary(user) {
  if (typeof localStorage === 'undefined') return;
  try {
    if (user) {
      localStorage.setItem(CACHED_USER_KEY, JSON.stringify({
        id: user.id,
        username: user.username,
        displayName: user.displayName,
        role: user.role,
        active: user.active !== false
      }));
    } else {
      localStorage.removeItem(CACHED_USER_KEY);
    }
  } catch {}
}

function getCachedUserSummary() {
  if (typeof localStorage === 'undefined') return null;
  try {
    const raw = localStorage.getItem(CACHED_USER_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

document.addEventListener('DOMContentLoaded', () => {
  initApp();
});

async function initApp() {
  const mainContainer = document.getElementById('main-content-view');
  const params = new URLSearchParams(window.location.search);
  const verifyParam = params.get('v');

  // Rute Publik QR Verifikasi Surat (?v=<verificationId>) — langsung tampilkan portal verifikasi tanpa editor/login
  if (verifyParam !== null) {
    renderQrVerificationView(mainContainer, verifyParam);
    return;
  }

  let currentTab = 'canvas';
  const appHeaderEl = document.getElementById('app-header');
  const navTabs = document.querySelectorAll('.nav-tab');
  const userPillBtn = document.getElementById('header-user-pill');
  const userRoleBadge = document.getElementById('header-user-role-badge');
  const syncDotEl = document.getElementById('header-sync-dot');
  const logoutBtn = document.getElementById('header-btn-logout');
  const accNavBadgeEl = document.getElementById('nav-acc-badge');
  const mobileNotifBtn = document.getElementById('btn-mobile-notif');
  const mobileMenuBtn = document.getElementById('btn-mobile-menu');
  const mobileMoreTabsBtn = document.getElementById('btn-mobile-more-tabs');
  const mobileMenuBadgeEl = document.getElementById('mobile-menu-badge');
  const mobileUserChipEl = document.getElementById('mobile-user-chip');
  const mobileUserChipLabelEl = document.getElementById('mobile-user-chip-label');
  const mobileSyncDotEl = document.getElementById('mobile-sync-dot');
  const bottomNavEl = document.getElementById('mobile-bottom-nav');
  const bottomHomeLabelEl = document.getElementById('bottom-home-label');
  const bottomFabLabelEl = document.getElementById('bottom-fab-label');
  const bottomAccBadgeEl = document.getElementById('bottom-acc-badge');
  const mobileFabBtn = document.getElementById('btn-mobile-fab');
  const mobileFabSheet = document.getElementById('mobile-fab-sheet');
  let lastMobileLockState = isCanvasLockedForUser(store.getState().authUser);

  function setMobileMenuOpen(open) {
    if (!appHeaderEl) return;
    const isOpen = Boolean(open);
    appHeaderEl.classList.toggle('is-mobile-menu-open', isOpen);
    if (mobileMenuBtn) {
      mobileMenuBtn.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
    }
    if (mobileMoreTabsBtn) {
      mobileMoreTabsBtn.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
      const span = mobileMoreTabsBtn.querySelector('span');
      if (span) span.textContent = isOpen ? 'Tutup ▴' : 'Lainnya ▾';
    }
  }

  function closeMobileFabSheet() {
    if (!mobileFabSheet) return;
    mobileFabSheet.classList.add('hidden');
    mobileFabSheet.setAttribute('aria-hidden', 'true');
    mobileFabSheet.innerHTML = '';
  }

  function openMobileFabSheet() {
    if (!mobileFabSheet) return;
    const stNow = store.getState();
    const user = stNow.authUser;
    const pendingCount = getPendingAccLetters(stNow.outgoingLetters || []).length;
    const canvasLocked = isCanvasLockedForUser(user);

    const actions = canvasLocked || user?.role === 'kepsek'
      ? [
          {
            id: 'sheet-acc',
            icon: '🛡️',
            title: pendingCount > 0 ? `Tinjau Antrean ACC (${pendingCount} Surat)` : 'Buka Antrean ACC Kepala Sekolah',
            sub: 'Periksa isi surat & terbitkan Barcode QR resmi',
            highlight: pendingCount > 0,
            onClick: () => activateNavTab('acc-queue')
          },
          {
            id: 'sheet-delegasi',
            icon: '🔐',
            title: 'Setelan Delegasi Kepsek → TU',
            sub: 'Aktifkan/kunci izin tanda tangan a.n. Kepsek oleh TU',
            onClick: () => activateNavTab('access')
          },
          {
            id: 'sheet-outgoing',
            icon: '📤',
            title: 'Riwayat & Arsip Surat Keluar',
            sub: 'Cetak QR, unduh .docx, atau lihat arsip Google Drive',
            onClick: () => activateNavTab('outgoing')
          },
          {
            id: 'sheet-school',
            icon: '🏛️',
            title: 'Profil Sekolah & Galeri Kop',
            sub: 'Lihat identitas sekolah dan format kop surat resmi',
            onClick: () => activateNavTab('school')
          }
        ]
      : [
          {
            id: 'sheet-canvas',
            icon: '📄',
            title: 'Buat Surat Baru (Lembar Kerja)',
            sub: 'Draft surat resmi SMK dengan kop & penomoran otomatis',
            highlight: true,
            onClick: () => activateNavTab('canvas')
          },
          {
            id: 'sheet-wordpro',
            icon: '✨',
            title: 'Buka Word Pro Multi-Halaman',
            sub: 'Editor surat bebas & bantuan AI Copilot Persuratan',
            onClick: () => activateNavTab('smart-canvas')
          },
          {
            id: 'sheet-acc',
            icon: '⏳',
            title: pendingCount > 0 ? `Pantau Antrean ACC (${pendingCount} Menunggu)` : 'Status Antrean ACC Kepsek',
            sub: 'Cek persetujuan Kepala Sekolah secara real-time tanpa refresh',
            onClick: () => activateNavTab('acc-queue')
          },
          {
            id: 'sheet-vault',
            icon: '📚',
            title: 'Buka Gudang Template Surat',
            sub: 'Gunakan ulang template surat sekolah siap pakai',
            onClick: () => activateNavTab('vault')
          }
        ];

    mobileFabSheet.innerHTML = `
      <div class="mobile-fab-sheet-card" role="dialog" aria-label="Menu Aksi Cepat">
        <div class="mobile-fab-sheet-handle"></div>
        <div class="mobile-fab-sheet-header">
          <div>
            <strong>${user?.role === 'kepsek' ? 'Menu Eksekutif Kepala Sekolah' : 'Aksi Cepat Persuratan'}</strong>
            <small>${user ? `${getRoleLabel(user.role)} (${user.username})` : ''}</small>
          </div>
          <button type="button" id="btn-close-fab-sheet" class="btn btn-sm btn-secondary">✕</button>
        </div>
        <div class="mobile-fab-sheet-list">
          ${actions.map(act => `
            <button type="button" class="mobile-fab-sheet-item ${act.highlight ? 'is-highlight' : ''}" data-sheet-act="${act.id}">
              <span class="mobile-fab-sheet-icon">${act.icon}</span>
              <span class="mobile-fab-sheet-text">
                <strong>${act.title}</strong>
                <small>${act.sub}</small>
              </span>
              <span class="mobile-fab-sheet-arrow">›</span>
            </button>
          `).join('')}
        </div>
      </div>
    `;
    mobileFabSheet.classList.remove('hidden');
    mobileFabSheet.setAttribute('aria-hidden', 'false');

    mobileFabSheet.querySelector('#btn-close-fab-sheet')?.addEventListener('click', closeMobileFabSheet);
    mobileFabSheet.onclick = (e) => {
      if (e.target === mobileFabSheet) closeMobileFabSheet();
    };
    actions.forEach(act => {
      mobileFabSheet.querySelector(`[data-sheet-act="${act.id}"]`)?.addEventListener('click', () => {
        closeMobileFabSheet();
        act.onClick();
      });
    });
  }

  mobileNotifBtn?.addEventListener('click', () => {
    setMobileMenuOpen(false);
    closeMobileFabSheet();
    activateNavTab('acc-queue');
  });

  mobileFabBtn?.addEventListener('click', () => {
    setMobileMenuOpen(false);
    if (mobileFabSheet && !mobileFabSheet.classList.contains('hidden')) {
      closeMobileFabSheet();
    } else {
      openMobileFabSheet();
    }
  });

  bottomNavEl?.querySelectorAll('[data-bottom-slot]').forEach(btn => {
    btn.addEventListener('click', () => {
      closeMobileFabSheet();
      const slot = btn.dataset.bottomSlot;
      const stNow = store.getState();
      const isKepsek = stNow.authUser?.role === 'kepsek';
      if (slot === 'home') {
        activateNavTab(isKepsek ? 'acc-queue' : 'canvas');
      } else if (slot === 'acc') {
        activateNavTab('acc-queue');
      } else if (slot === 'outgoing') {
        activateNavTab('outgoing');
      } else if (slot === 'settings') {
        activateNavTab('access');
      }
    });
  });

  mobileMenuBtn?.addEventListener('click', (e) => {
    e.stopPropagation();
    closeMobileFabSheet();
    const isOpen = appHeaderEl?.classList.contains('is-mobile-menu-open');
    setMobileMenuOpen(!isOpen);
  });

  mobileMoreTabsBtn?.addEventListener('click', (e) => {
    e.stopPropagation();
    closeMobileFabSheet();
    const isOpen = appHeaderEl?.classList.contains('is-mobile-menu-open');
    setMobileMenuOpen(!isOpen);
  });

  mobileUserChipEl?.addEventListener('click', () => {
    setMobileMenuOpen(false);
    closeMobileFabSheet();
    activateNavTab('access');
  });

  mainContainer?.addEventListener('click', () => {
    if (appHeaderEl?.classList.contains('is-mobile-menu-open')) {
      setMobileMenuOpen(false);
    }
  });

  function renderMobileCanvasLockPanel(container, targetTab) {
    const label = targetTab === 'smart-canvas' ? 'Word Pro (Kanvas Lanjutan)' : 'Lembar Kerja Surat';
    const pendingCount = getPendingAccLetters(store.getState().outgoingLetters).length;
    container.innerHTML = `
      <div class="view-atmosphere">
        <div class="kepsek-mobile-lock-shell">
          <div class="kepsek-mobile-lock-card">
            <div class="kepsek-lock-icon-wrap">
              <span class="kepsek-lock-badge">🔒</span>
              <span>💻</span>
            </div>
            <h2>Akses ${label} Terkunci di Handphone</h2>
            <p class="kepsek-lock-desc">
              Untuk akun <strong>Kepala Sekolah</strong>, fitur pengeditan kanvas dokumen surat
              <strong>hanya dapat dibuka melalui Komputer Desktop, Laptop, atau PC</strong> guna menjaga tata letak kertas A4/F4 dan tanda tangan resmi tetap presisi.
            </p>
            <div class="kepsek-lock-info-box">
              Dari Handphone ini Anda tetap memiliki akses penuh untuk:
              meninjau &amp; meng-ACC antrean surat QR, melihat arsip Surat Keluar, dan mengatur Setelan Delegasi/Hak Akses.
            </div>
            <div class="kepsek-lock-actions">
              <button type="button" id="btn-lock-go-acc" class="btn btn-primary ${pendingCount > 0 ? 'acc-glow-btn' : ''}">
                🛡️ Buka Antrean ACC Kepala Sekolah ${pendingCount > 0 ? `(${pendingCount} Perlu ACC)` : ''}
              </button>
              <div class="kepsek-lock-sub-links">
                <button type="button" class="btn btn-sm btn-secondary" data-jump-tab="outgoing">📤 Surat Keluar</button>
                <button type="button" class="btn btn-sm btn-secondary" data-jump-tab="access">🔐 Setelan &amp; Akses</button>
                <button type="button" class="btn btn-sm btn-secondary" data-jump-tab="school">🏛️ Profil Sekolah</button>
                <button type="button" class="btn btn-sm btn-secondary" data-jump-tab="settings">🔗 Gudang API</button>
              </div>
            </div>
          </div>
        </div>
      </div>
    `;
    const btnGoAcc = container.querySelector('#btn-lock-go-acc');
    if (btnGoAcc) btnGoAcc.onclick = () => activateNavTab('acc-queue');
    container.querySelectorAll('[data-jump-tab]').forEach(btn => {
      btn.onclick = () => activateNavTab(btn.dataset.jumpTab);
    });
  }

  function updateBottomNavActiveState(tabName) {
    if (mobileMoreTabsBtn) {
      const isTu = store.getState().authUser?.role === 'tu';
      const primaryTopTabs = isTu ? ['canvas', 'acc-queue', 'outgoing'] : ['acc-queue', 'outgoing', 'vault'];
      mobileMoreTabsBtn.classList.toggle('has-active-child', !primaryTopTabs.includes(tabName));
    }
    if (!bottomNavEl) return;
    bottomNavEl.querySelectorAll('[data-bottom-slot]').forEach(btn => {
      const slot = btn.dataset.bottomSlot;
      let isActive = false;
      if (slot === 'home') {
        isActive = tabName === 'canvas' || tabName === 'smart-canvas';
      } else if (slot === 'acc') {
        isActive = tabName === 'acc-queue';
      } else if (slot === 'outgoing') {
        isActive = tabName === 'outgoing';
      } else if (slot === 'settings') {
        isActive = tabName === 'access' || tabName === 'settings' || tabName === 'school' || tabName === 'vault';
      }
      btn.classList.toggle('active', isActive);
    });
  }

  function activateNavTab(tabName) {
    currentTab = tabName;
    setMobileMenuOpen(false);
    closeMobileFabSheet();
    let activeTabBtn = null;
    navTabs.forEach(t => {
      t.classList.remove('active', 'btn-primary', 'btn-secondary');
      if (t.dataset.tab === tabName) {
        t.classList.add('active');
        activeTabBtn = t;
      }
    });
    updateBottomNavActiveState(tabName);
    if (activeTabBtn && typeof activeTabBtn.scrollIntoView === 'function') {
      try {
        activeTabBtn.scrollIntoView({ behavior: 'smooth', inline: 'center', block: 'nearest' });
      } catch {}
    }
    renderTab(tabName, mainContainer);
  }

  navTabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const targetTab = tab.dataset.tab;
      activateNavTab(targetTab);
    });
  });

  function renderTab(tabName, container) {
    container.innerHTML = '';
    const canvasLocked = isCanvasLockedForUser(store.getState().authUser);
    switch (tabName) {
      case 'canvas':
        if (canvasLocked) {
          renderMobileCanvasLockPanel(container, 'canvas');
        } else {
          renderWordCanvas(container);
        }
        break;
      case 'smart-canvas':
        if (canvasLocked) {
          renderMobileCanvasLockPanel(container, 'smart-canvas');
        } else {
          renderSmartWordCanvas(container);
        }
        break;
      case 'acc-queue':
        renderAccApprovalView(container, { onNavigateTab: activateNavTab });
        break;
      case 'outgoing':
        renderOutgoingLettersView(container);
        break;
      case 'vault':
        renderVaultView(container);
        break;
      case 'school':
        renderSchoolProfileView(container);
        break;
      case 'settings':
        renderSettingsView(container);
        break;
      case 'access':
        renderAccessControlView(container);
        break;
      default:
        if (canvasLocked) {
          renderAccApprovalView(container, { onNavigateTab: activateNavTab });
        } else {
          renderWordCanvas(container);
        }
    }
    renderLiveAccTray({ onNavigateToAccQueue: () => activateNavTab('acc-queue') });
  }

  function updateHeaderSessionUi(state) {
    const user = state.authUser;
    const isKepsek = user?.role === 'kepsek';
    const canvasLocked = isCanvasLockedForUser(user);
    const pendingList = getPendingAccLetters(state.outgoingLetters || []);

    document.body.classList.toggle('role-kepsek', Boolean(isKepsek));
    document.body.classList.toggle('role-tu', user?.role === 'tu');
    document.body.classList.toggle('kepsek-mobile-locked', Boolean(canvasLocked));
    document.body.classList.toggle('is-logged-in', Boolean(user));

    if (bottomHomeLabelEl) {
      bottomHomeLabelEl.textContent = isKepsek ? 'Beranda' : 'Kanvas';
    }
    if (bottomFabLabelEl) {
      bottomFabLabelEl.textContent = isKepsek ? 'Aksi Cepat' : 'Buat Surat';
    }
    updateBottomNavActiveState(currentTab);

    navTabs.forEach(t => {
      if (t.dataset.tab === 'canvas' || t.dataset.tab === 'smart-canvas') {
        t.classList.toggle('is-locked-mobile', Boolean(canvasLocked));
        if (canvasLocked) {
          t.title = 'Terkunci di Handphone — Gunakan Komputer Desktop/Laptop/PC';
        }
      }
    });

    if (accNavBadgeEl) {
      if (user && pendingList.length > 0) {
        accNavBadgeEl.textContent = String(pendingList.length);
        accNavBadgeEl.classList.remove('hidden');
        accNavBadgeEl.classList.toggle('is-glowing', user.role === 'kepsek' || user.role === 'admin');
        accNavBadgeEl.classList.toggle('is-waiting', user.role === 'tu');
      } else {
        accNavBadgeEl.classList.add('hidden');
        accNavBadgeEl.classList.remove('is-glowing', 'is-waiting');
      }
    }

    if (mobileMenuBadgeEl) {
      if (user && pendingList.length > 0) {
        mobileMenuBadgeEl.textContent = String(pendingList.length);
        mobileMenuBadgeEl.classList.remove('hidden');
        mobileMenuBadgeEl.classList.toggle('is-glowing', user.role === 'kepsek' || user.role === 'admin');
      } else {
        mobileMenuBadgeEl.classList.add('hidden');
        mobileMenuBadgeEl.classList.remove('is-glowing');
      }
    }

    if (bottomAccBadgeEl) {
      if (user && pendingList.length > 0) {
        bottomAccBadgeEl.textContent = String(pendingList.length);
        bottomAccBadgeEl.classList.remove('hidden');
      } else {
        bottomAccBadgeEl.classList.add('hidden');
      }
    }

    if (!user) {
      userPillBtn?.classList.add('hidden');
      mobileNotifBtn?.classList.add('hidden');
      mobileUserChipEl?.classList.add('hidden');
      logoutBtn?.classList.add('hidden');
      renderLiveAccTray({ onNavigateToAccQueue: () => activateNavTab('acc-queue') });
      return;
    }
    const roleIcon = user.role === 'admin' ? '🛡️' : user.role === 'kepsek' ? '🎓' : '📋';
    if (userRoleBadge) {
      userRoleBadge.textContent = `${roleIcon} ${getRoleLabel(user.role)} (${user.username})`;
    }
    if (mobileUserChipLabelEl) {
      mobileUserChipLabelEl.textContent = `${roleIcon} ${user.username}`;
    }
    if (syncDotEl) {
      if (state.cloudSyncState === 'syncing') {
        syncDotEl.textContent = '⏳';
        syncDotEl.title = 'Menyinkronkan ke server...';
      } else if (state.cloudSyncState === 'error') {
        syncDotEl.textContent = '⚠️';
        syncDotEl.title = 'Sinkron cloud tertunda (tersimpan lokal)';
      } else if (state.driveStatus?.configured) {
        syncDotEl.textContent = '☁️ Drive';
        syncDotEl.title = 'Tersinkron ke Cloud & Google Drive';
      } else {
        syncDotEl.textContent = '☁️';
        syncDotEl.title = 'Tersinkron ke Server Cloud';
      }
    }
    if (mobileSyncDotEl) {
      mobileSyncDotEl.textContent = state.cloudSyncState === 'syncing'
        ? '⏳'
        : state.cloudSyncState === 'error'
        ? '⚠️'
        : '☁️';
    }
    userPillBtn?.classList.remove('hidden');
    mobileNotifBtn?.classList.remove('hidden');
    mobileUserChipEl?.classList.remove('hidden');
    logoutBtn?.classList.remove('hidden');
    saveCachedUserSummary(user);
    renderLiveAccTray({ onNavigateToAccQueue: () => activateNavTab('acc-queue') });
  }

  store.subscribe(updateHeaderSessionUi);

  userPillBtn?.addEventListener('click', () => {
    setMobileMenuOpen(false);
    activateNavTab('access');
  });

  const requireLoginPrompt = (msg = '') => {
    setMobileMenuOpen(false);
    renderLoginView(mainContainer, {
      statusMessage: msg,
      onAuthenticated: () => {
        openInitialTab();
      }
    });
  };

  logoutBtn?.addEventListener('click', () => {
    setMobileMenuOpen(false);
    if (!confirm('Keluar dari akun aktif sekarang?')) return;
    stopLiveSync();
    resetAccTrayNotifications();
    saveCachedUserSummary(null);
    store.clearAuthState();
    mainContainer.innerHTML = '';
    requireLoginPrompt('Anda telah keluar. Silakan login kembali dengan akun Administrator, Kepala Sekolah, atau TU.');
  });

  onWorkspaceUnauthorized((errMsg) => {
    stopLiveSync();
    resetAccTrayNotifications();
    saveCachedUserSummary(null);
    store.clearAuthState();
    requireLoginPrompt(errMsg || 'Sesi Anda telah berakhir. Silakan login kembali.');
  });

  // Live SPA Polling Update Tanpa Refresh Halaman (Non-Destruktif terhadap Kanvas)
  initLiveSync({
    onDeltaApplied: ({ transitions } = {}) => {
      const user = store.getState().authUser;
      const canApprove = user?.role === 'kepsek' || user?.role === 'admin';

      for (const item of (transitions?.approved || [])) {
        recordCompletedAccForTray(item, 'approved');
        showToast(`🎉 Surat #${String(item.sequenceNumber || 0).padStart(3, '0')} (${item.nomorSurat}) telah DI-ACC Kepala Sekolah! Barcode QR siap dicetak.`, 'success');
      }
      for (const item of (transitions?.rejected || [])) {
        recordCompletedAccForTray(item, 'rejected');
        showToast(`❌ Pengajuan ACC Surat #${String(item.sequenceNumber || 0).padStart(3, '0')} ditolak Kepala Sekolah.`, 'warning');
      }
      for (const item of (transitions?.newRequests || [])) {
        if (canApprove) {
          showToast(`🔔 Permintaan ACC baru: ${item.title} (#${String(item.sequenceNumber || 0).padStart(3, '0')})`, 'info');
        }
      }

      const activeInput = document.activeElement && ['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement.tagName);
      const openModal = Boolean(mainContainer.querySelector('.modal-overlay:not(.hidden)'));
      if (!activeInput && !openModal && (currentTab === 'acc-queue' || currentTab === 'outgoing' || currentTab === 'access')) {
        renderTab(currentTab, mainContainer);
      } else {
        renderLiveAccTray({ onNavigateToAccQueue: () => activateNavTab('acc-queue') });
      }
    }
  });

  window.addEventListener('resize', () => {
    const user = store.getState().authUser;
    const lockedNow = isCanvasLockedForUser(user);
    if (lockedNow !== lastMobileLockState) {
      lastMobileLockState = lockedNow;
      updateHeaderSessionUi(store.getState());
      if (currentTab === 'canvas' || currentTab === 'smart-canvas') {
        renderTab(currentTab, mainContainer);
      }
    }
  });

  // 2. Header Quick Download (Mendukung Pilihan QR vs TTD Manual pada Kanvas & Word Pro)
  document.getElementById('header-btn-download').addEventListener('click', async () => {
    if (isCanvasLockedForUser(store.getState().authUser) && (currentTab === 'canvas' || currentTab === 'smart-canvas')) {
      showToast('Kanvas editor terkunci di Handphone. Unduh dokumen melalui Antrean ACC atau Surat Keluar.', 'info');
      return;
    }
    if (currentTab === 'smart-canvas') {
      triggerSmartCanvasDownloadWithChoice();
      return;
    }
    if (currentTab === 'canvas' || mainContainer.querySelector('#editable-sig-date')) {
      triggerWordCanvasDownloadWithChoice(mainContainer);
      return;
    }

    const docModel = store.getState().document || getActiveSmartDocModel();
    showToast('Menyusun berkas Word .docx resmi...', 'info');
    try {
      const blob = await renderDocx(docModel, 'blob');
      const safeTitle = (docModel.meta?.title || 'surat_dinas')
        .replace(/[^a-zA-Z0-9_\-]/g, '_')
        .slice(0, 45);
      downloadFile(blob, `${safeTitle}.docx`);
      showToast('Berkas .docx berhasil diunduh.', 'success');
    } catch (err) {
      showToast(`Gagal mengunduh: ${err.message}`, 'error');
    }
  });

  // 3. Header Quick Print (Pilihan Barcode QR Pengesahan vs Kosongan TTD Manual Pulpen)
  document.getElementById('header-btn-print').addEventListener('click', () => {
    if (isCanvasLockedForUser(store.getState().authUser) && (currentTab === 'canvas' || currentTab === 'smart-canvas')) {
      showToast('Kanvas editor terkunci di Handphone. Cetak dokumen melalui Antrean ACC atau Surat Keluar.', 'info');
      return;
    }
    if (currentTab === 'smart-canvas') {
      triggerSmartCanvasPrintWithChoice();
      return;
    }
    if (currentTab === 'canvas' || mainContainer.querySelector('#editable-sig-date')) {
      triggerWordCanvasPrintWithChoice(mainContainer);
      return;
    }
    runSafeWindowPrint();
  });

  function openInitialTab() {
    hideLoginView();
    const stNow = store.getState();
    const requestedTab = new URLSearchParams(window.location.search).get('tab');
    const isKepsek = stNow.authUser?.role === 'kepsek';
    const target = requestedTab && document.querySelector(`.nav-tab[data-tab="${CSS.escape(requestedTab)}"]`)
      ? requestedTab
      : isKepsek
      ? 'acc-queue'
      : 'canvas';
    lastMobileLockState = isCanvasLockedForUser(stNow.authUser);
    updateHeaderSessionUi(stNow);
    activateNavTab(target);
    startLiveSync(stNow.syncRev);
  }

  // Inisialisasi PWA (Service Worker & Install App Prompt)
  initPwa();

  const existingToken = getStoredAuthToken();
  if (!existingToken) {
    requireLoginPrompt();
    return;
  }

  // Verifikasi & tarik data workspace dari server
  const cachedSummary = getCachedUserSummary();
  if (cachedSummary) {
    store.setAuthSessionData({ user: cachedSummary });
  }
  const bootRes = await fetchWorkspaceBootstrap();
  if (bootRes.ok) {
    store.applyRemoteBootstrap(bootRes);
    openInitialTab();
  } else if (bootRes.offline && cachedSummary) {
    showToast('Mode offline aktif menggunakan data tersimpan lokal. Sinkron ulang saat online.', 'warning');
    openInitialTab();
  } else {
    clearStoredAuthToken();
    saveCachedUserSummary(null);
    store.clearAuthState();
    requireLoginPrompt(bootRes.error || 'Sesi login berakhir, silakan masuk kembali.');
  }
}
