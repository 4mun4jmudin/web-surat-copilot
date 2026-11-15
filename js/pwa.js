/**
 * pwa.js - Registrasi Service Worker, Deteksi Mode Offline, dan Tombol Pasang Aplikasi (Install Prompt)
 */

import { showToast } from './ui/notification.js';

let deferredInstallPrompt = null;

function isRunningStandalone() {
  return window.matchMedia?.('(display-mode: standalone)').matches || window.navigator.standalone === true;
}

function isIosSafari() {
  const ua = window.navigator.userAgent || '';
  const isIos = /iPad|iPhone|iPod/i.test(ua);
  const isSafari = /Safari/i.test(ua) && !/CriOS|FxiOS|OPiOS|EdgiOS/i.test(ua);
  return isIos && isSafari;
}

export function initPwa() {
  const installBtn = document.getElementById('header-btn-install');

  const updateInstallButtonVisibility = () => {
    if (!installBtn) return;
    if (isRunningStandalone()) {
      installBtn.classList.add('hidden');
      return;
    }
    if (deferredInstallPrompt || isIosSafari()) {
      installBtn.classList.remove('hidden');
    } else {
      installBtn.classList.add('hidden');
    }
  };

  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredInstallPrompt = e;
    updateInstallButtonVisibility();
  });

  window.addEventListener('appinstalled', () => {
    deferredInstallPrompt = null;
    updateInstallButtonVisibility();
    showToast('AI DocBuilder berhasil dipasang sebagai aplikasi di perangkat Anda!', 'success');
  });

  if (installBtn) {
    installBtn.addEventListener('click', async () => {
      if (deferredInstallPrompt) {
        try {
          deferredInstallPrompt.prompt();
          const { outcome } = await deferredInstallPrompt.userChoice;
          deferredInstallPrompt = null;
          updateInstallButtonVisibility();
          if (outcome === 'accepted') {
            showToast('Memasang aplikasi ke layar utama / desktop...', 'success');
          }
        } catch (err) {
          showToast('Silakan pasang melalui ikon Install di bilah alamat browser.', 'info');
        }
        return;
      }

      if (isIosSafari()) {
        showToast('Di Safari iOS: ketuk ikon Bagikan (Share) lalu pilih "Tambah ke Layar Utama".', 'info');
      }
    });
  }

  updateInstallButtonVisibility();

  // Notifikasi ringan status jaringan (fitur lokal & .docx tetap aktif offline)
  window.addEventListener('offline', () => {
    showToast('Mode Offline aktif: Editor surat, Gudang Surat & unduh .docx tetap bisa digunakan.', 'info');
  });
  window.addEventListener('online', () => {
    showToast('Koneksi internet tersambung kembali. AI Copilot siap digunakan.', 'success');
  });

  // Registrasi Service Worker hanya pada protokol aman (https / localhost)
  const isSupportedOrigin =
    'serviceWorker' in navigator &&
    (window.location.protocol === 'https:' ||
      window.location.hostname === 'localhost' ||
      window.location.hostname === '127.0.0.1');

  if (!isSupportedOrigin) return;

  const registerSw = async () => {
    try {
      const registration = await navigator.serviceWorker.register('./sw.js', { scope: './' });

      registration.addEventListener('updatefound', () => {
        const newWorker = registration.installing;
        if (!newWorker) return;
        newWorker.addEventListener('statechange', () => {
          if (newWorker.state === 'installed' && navigator.serviceWorker.controller) {
            newWorker.postMessage({ type: 'SKIP_WAITING' });
            showToast('Pembaruan aplikasi tersedia dan akan diterapkan saat muat ulang berikutnya.', 'info');
          }
        });
      });
    } catch (err) {
      console.warn('[PWA] Gagal mendaftarkan Service Worker:', err);
    }
  };

  if (document.readyState === 'complete') {
    registerSw();
  } else {
    window.addEventListener('load', registerSw);
  }
}
