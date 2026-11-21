/**
 * sw.js - Service Worker PWA untuk AI DocBuilder Persuratan Sekolah
 * Menyediakan caching offline untuk antarmuka, aset Word (.docx) generator, dan pustaka lokal.
 */

const CACHE_NAME = 'ai-docbuilder-pwa-v31-20261121';

const PRECACHE_ASSETS = [
  './',
  './index.html',
  './manifest.webmanifest',
  './css/app.css',
  './img/logo.png',
  './img/icon-192.png',
  './img/icon-512.png',
  './img/icon-512-maskable.png',
  './img/Calm Green Wave Gradient Background.png',
  './img/AI_DocBuilder_Mobile_UI_Assets/branding/brand_logo_original.png',
  './img/AI_DocBuilder_Mobile_UI_Assets/backgrounds/01_header_emerald.png',
  './img/AI_DocBuilder_Mobile_UI_Assets/backgrounds/02_page_background.png',
  './img/AI_DocBuilder_Mobile_UI_Assets/illustrations/01_document_qr_approved.svg',
  './img/AI_DocBuilder_Mobile_UI_Assets/illustrations/02_empty_queue_approved.svg',
  './img/AI_DocBuilder_Mobile_UI_Assets/illustrations/03_mobile_qr_approved.svg',
  './img/AI_DocBuilder_Mobile_UI_Assets/icons/icon_archive.svg',
  './img/AI_DocBuilder_Mobile_UI_Assets/icons/icon_approval.svg',
  './img/AI_DocBuilder_Mobile_UI_Assets/icons/icon_file.svg',
  './img/AI_DocBuilder_Mobile_UI_Assets/icons/icon_scan.svg',
  './img/AI_DocBuilder_Mobile_UI_Assets/icons/icon_history.svg',
  './vendor/jszip.min.js',
  './vendor/docx-preview.min.js',
  './vendor/docx-preview.min.mjs',
  './vendor/qrcode.mjs',
  './vendor/qrcode_UTF8.mjs',
  './js/main.js',
  './js/pwa.js',
  './js/state.js',
  './js/auth/permissions.js',
  './js/sync/workspaceClient.js',
  './js/sync/liveSync.js',
  './js/ai/client.js',
  './js/ai/knowledgeBase.js',
  './js/ai/prompts.js',
  './js/ai/surgicalEdit.js',
  './js/checks/consistency.js',
  './js/docx/profiles/jurnal.js',
  './js/docx/profiles/pkl.js',
  './js/docx/profiles/skripsi.js',
  './js/docx/profiles/suratDinas.js',
  './js/docx/fields.js',
  './js/docx/numbering.js',
  './js/docx/patch.js',
  './js/docx/render.js',
  './js/docx/styles.js',
  './js/letters/templates/suratKeterangan.js',
  './js/letters/templates/suratTugasSppd.js',
  './js/letters/agenda.js',
  './js/letters/builder.js',
  './js/letters/defaultKop.js',
  './js/letters/incoming.js',
  './js/letters/numbering.js',
  './js/letters/schoolProfile.js',
  './js/letters/verification.js',
  './js/model/sample.js',
  './js/model/schema.js',
  './js/ui/accApprovalView.js',
  './js/ui/accQueueWidget.js',
  './js/ui/accessControlView.js',
  './js/ui/aiView.js',
  './js/ui/consistencyView.js',
  './js/ui/editor.js',
  './js/ui/icons.js',
  './js/ui/letterFormView.js',
  './js/ui/loginView.js',
  './js/ui/notification.js',
  './js/ui/incomingLettersView.js',
  './js/ui/outgoingLettersView.js',
  './js/ui/patchView.js',
  './js/ui/preview.js',
  './js/ui/printChoiceModal.js',
  './js/ui/qrFloatingHelper.js',
  './js/ui/schoolProfileView.js',
  './js/ui/settingsView.js',
  './js/ui/vaultView.js',
  './js/ui/vaultSaveHelper.js',
  './js/ui/verifyView.js',
  './js/ui/wordCanvas.js',
  './js/ui/smartWordCanvas.js',
  './js/ui/tableInteraction.js',
  './js/utils/download.js',
  './js/utils/htmlAst.js',
  './js/utils/xml.js',
  './js/utils/zip.js'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(async (cache) => {
      await Promise.all(
        PRECACHE_ASSETS.map(async (assetPath) => {
          try {
            const url = encodeURI(assetPath);
            const res = await fetch(url, { cache: 'reload' });
            if (res && res.ok) {
              await cache.put(url, res);
            }
          } catch (err) {
            console.warn('[PWA] Gagal cache awal aset:', assetPath, err);
          }
        })
      );
    }).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((oldKey) => caches.delete(oldKey))
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;
  if (url.pathname.startsWith('/api/') || url.pathname.startsWith('/verify-letter') || url.pathname.includes('/functions/')) return;

  // Navigasi dokumen HTML: Network-First dengan fallback cache index.html untuk mode offline
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req)
        .then((res) => {
          if (res && res.ok) {
            const clone = res.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put('./index.html', clone)).catch(() => {});
          }
          return res;
        })
        .catch(async () => {
          const cache = await caches.open(CACHE_NAME);
          return (await cache.match('./index.html')) || (await cache.match('./'));
        })
    );
    return;
  }

  // Aset berat/statis (gambar, ikon, vendor): Stale-While-Revalidate
  const isHeavyStatic = url.pathname.includes('/img/') || url.pathname.includes('/vendor/') || url.pathname.endsWith('manifest.webmanifest');
  if (isHeavyStatic) {
    event.respondWith(
      caches.open(CACHE_NAME).then(async (cache) => {
        const cached = await cache.match(req, { ignoreSearch: true });
        const networkPromise = fetch(req)
          .then((networkRes) => {
            if (networkRes && networkRes.ok && networkRes.type === 'basic') {
              const cacheKey = new URL(url.pathname, self.location.origin).href;
              cache.put(cacheKey, networkRes.clone()).catch(() => {});
            }
            return networkRes;
          })
          .catch(() => null);

        if (cached) {
          event.waitUntil(networkPromise);
          return cached;
        }
        return (await networkPromise) || new Response('Offline', { status: 503 });
      })
    );
    return;
  }

  // Kode aplikasi (.js, .css): Network-First agar pembaruan cepat tampil, fallback ke cache saat offline
  event.respondWith(
    fetch(req)
      .then((networkRes) => {
        if (networkRes && networkRes.ok && networkRes.type === 'basic') {
          const clone = networkRes.clone();
          const cacheKey = new URL(url.pathname, self.location.origin).href;
          caches.open(CACHE_NAME).then((cache) => cache.put(cacheKey, clone)).catch(() => {});
        }
        return networkRes;
      })
      .catch(async () => {
        const cache = await caches.open(CACHE_NAME);
        return (await cache.match(req, { ignoreSearch: true })) || new Response('Offline', { status: 503 });
      })
  );
});
