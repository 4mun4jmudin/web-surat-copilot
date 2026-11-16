/**
 * accQueueWidget.js - Panel Mengambang Non-Blocking Antrean ACC Real-Time
 * - Untuk TU: menampilkan spinner berputar terus-menerus selama menunggu ACC Kepala Sekolah
 *   tanpa memblokir pembuatan/pengeditan surat lainnya, bertahan walau web di-refresh,
 *   dan langsung menampilkan tombol Cetak QR saat di-ACC tanpa refresh halaman.
 * - Untuk Kepala Sekolah / Admin: menampilkan notifikasi menyala saat ada surat masuk butuh ACC.
 */

import { store } from '../state.js';
import { getPendingAccLetters } from '../auth/permissions.js';
import { promptOutgoingActionWithChoice } from './outgoingLettersView.js';
import { showToast } from './notification.js';

let trayCollapsed = false;
const recentlyCompletedMap = new Map(); // id -> { outcome: 'approved'|'rejected', item }
const dismissedCompletedIds = new Set();

function escHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export function recordCompletedAccForTray(item, outcome = 'approved') {
  if (!item?.id) return;
  dismissedCompletedIds.delete(item.id);
  recentlyCompletedMap.set(item.id, { outcome, item, at: Date.now() });
}

export function resetAccTrayNotifications() {
  recentlyCompletedMap.clear();
  dismissedCompletedIds.clear();
  trayCollapsed = false;
}

export function renderLiveAccTray({ onNavigateToAccQueue } = {}) {
  const host = typeof document !== 'undefined' ? document.getElementById('live-acc-tray') : null;
  if (!host) return;

  const st = store.getState();
  const user = st.authUser;
  if (!user) {
    host.innerHTML = '';
    return;
  }

  const allOutgoing = st.outgoingLetters || [];
  const pending = getPendingAccLetters(allOutgoing);

  // Sinkronkan status item recentlyCompletedMap dengan data terbaru di store
  for (const [id, meta] of recentlyCompletedMap.entries()) {
    const fresh = allOutgoing.find(x => x.id === id);
    if (fresh) {
      meta.item = fresh;
      if (fresh.verificationStatus === 'sah') meta.outcome = 'approved';
      else if (fresh.verificationStatus === 'rejected') meta.outcome = 'rejected';
    }
  }

  const completedNotices = Array.from(recentlyCompletedMap.values())
    .filter(x => !dismissedCompletedIds.has(x.item?.id))
    .slice(0, 4);

  // Untuk Kepala Sekolah / Admin: tampilkan pill notifikasi menyala jika ada antrean & tidak sedang di tab Antrean ACC
  if (user.role === 'kepsek' || user.role === 'admin') {
    if (pending.length === 0 || st.activeTab === 'acc-queue') {
      host.innerHTML = '';
      return;
    }
    host.innerHTML = `
      <button type="button" id="btn-acc-tray-jump" class="acc-tray-kepsek-pill acc-glow-pill" title="Buka Antrean ACC Kepala Sekolah">
        <span class="acc-glow-dot"></span>
        <span>🛡️ <strong>${pending.length} Surat</strong> Menunggu ACC Anda — Tinjau Sekarang</span>
      </button>
    `;
    const jumpBtn = host.querySelector('#btn-acc-tray-jump');
    if (jumpBtn && typeof onNavigateToAccQueue === 'function') {
      jumpBtn.onclick = () => onNavigateToAccQueue();
    }
    return;
  }

  // Untuk TU: tampilkan panel antrean non-blocking jika ada yang sedang menunggu ACC atau baru saja di-ACC/ditolak
  if (pending.length === 0 && completedNotices.length === 0) {
    host.innerHTML = '';
    return;
  }

  if (trayCollapsed) {
    host.innerHTML = `
      <button type="button" id="btn-acc-tray-expand" class="acc-tray-collapsed-btn ${pending.length > 0 ? 'is-waiting' : 'is-ready'}">
        ${pending.length > 0
          ? `<span class="acc-live-spinner"></span> <span>Menunggu ACC Kepsek (${pending.length})</span>`
          : `<span>✅ Update ACC Siap (${completedNotices.length})</span>`
        }
      </button>
    `;
    host.querySelector('#btn-acc-tray-expand').onclick = () => {
      trayCollapsed = false;
      renderLiveAccTray({ onNavigateToAccQueue });
    };
    return;
  }

  host.innerHTML = `
    <div class="acc-tray-card" role="status" aria-live="polite">
      <div class="acc-tray-header">
        <div class="acc-tray-title">
          ${pending.length > 0
            ? '<span class="acc-live-spinner"></span><span>Menunggu ACC Kepala Sekolah</span>'
            : '<span>✅ Status ACC Kepala Sekolah</span>'
          }
          <span class="acc-tray-count-pill">${pending.length} antre</span>
        </div>
        <div class="acc-tray-header-actions">
          <button type="button" id="btn-acc-tray-open-tab" class="acc-tray-mini-btn" title="Buka halaman Antrean ACC">Antrean</button>
          <button type="button" id="btn-acc-tray-min" class="acc-tray-mini-btn" title="Kecilkan panel">▾</button>
        </div>
      </div>

      <div class="acc-tray-body">
        <div class="acc-tray-hint">
          ✨ Anda tetap bebas membuat &amp; mengedit surat lain. Status ACC otomatis terbarui tanpa refresh.
        </div>

        ${completedNotices.map(entry => {
          const it = entry.item;
          const isApproved = entry.outcome === 'approved';
          return `
            <div class="acc-tray-item ${isApproved ? 'is-approved' : 'is-rejected'}">
              <div class="acc-tray-item-top">
                <strong>${isApproved ? '✅ DI-ACC KEPSEK (QR Terbit!)' : '❌ Ditolak Kepsek'}</strong>
                <button type="button" class="acc-tray-dismiss-btn" data-id="${escHtml(it.id)}" title="Tutup info ini">✕</button>
              </div>
              <div class="acc-tray-item-title">${escHtml(it.title)} • #${String(it.sequenceNumber || 0).padStart(3, '0')}</div>
              <div class="acc-tray-item-meta">${escHtml(it.nomorSurat)}</div>
              ${isApproved ? `
                <div class="acc-tray-item-actions">
                  <button type="button" class="btn btn-sm btn-primary btn-tray-print" data-id="${escHtml(it.id)}">🖨️ Cetak QR</button>
                  <button type="button" class="btn btn-sm btn-success btn-tray-docx" data-id="${escHtml(it.id)}">⬇️ .docx QR</button>
                </div>
              ` : `
                <div class="acc-tray-reject-note">${escHtml(it.rejectedReason || 'Perlu revisi sebelum diajukan ulang.')}</div>
              `}
            </div>
          `;
        }).join('')}

        ${pending.map(it => `
          <div class="acc-tray-item is-pending">
            <div class="acc-tray-item-top">
              <span class="acc-tray-pending-tag"><span class="acc-live-spinner acc-spinner-sm"></span> Menunggu ACC Kepala Sekolah...</span>
              <span class="acc-tray-seq">#${String(it.sequenceNumber || 0).padStart(3, '0')}</span>
            </div>
            <div class="acc-tray-item-title">${escHtml(it.title || 'Surat Keluar')}</div>
            <div class="acc-tray-item-meta">${escHtml(it.nomorSurat)} • ${escHtml(it.recipientName || '-')}</div>
          </div>
        `).join('')}
      </div>
    </div>
  `;

  const btnMin = host.querySelector('#btn-acc-tray-min');
  if (btnMin) {
    btnMin.onclick = () => {
      trayCollapsed = true;
      renderLiveAccTray({ onNavigateToAccQueue });
    };
  }
  const btnOpenTab = host.querySelector('#btn-acc-tray-open-tab');
  if (btnOpenTab && typeof onNavigateToAccQueue === 'function') {
    btnOpenTab.onclick = () => onNavigateToAccQueue();
  }
  host.querySelectorAll('.acc-tray-dismiss-btn').forEach(btn => {
    btn.onclick = () => {
      dismissedCompletedIds.add(btn.dataset.id);
      renderLiveAccTray({ onNavigateToAccQueue });
    };
  });
  host.querySelectorAll('.btn-tray-print').forEach(btn => {
    btn.onclick = () => {
      const item = store.getOutgoingLetters().find(x => x.id === btn.dataset.id);
      if (!item) {
        showToast('Data surat belum ditemukan.', 'warning');
        return;
      }
      promptOutgoingActionWithChoice(item, 'print', document.getElementById('main-content-view'));
    };
  });
  host.querySelectorAll('.btn-tray-docx').forEach(btn => {
    btn.onclick = () => {
      const item = store.getOutgoingLetters().find(x => x.id === btn.dataset.id);
      if (!item) {
        showToast('Data surat belum ditemukan.', 'warning');
        return;
      }
      promptOutgoingActionWithChoice(item, 'docx', document.getElementById('main-content-view'));
    };
  });
}
