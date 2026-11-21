/**
 * accApprovalView.js - Dashboard Prioritas Antrean ACC Kepala Sekolah & Tracking TU
 * Menggunakan aset grafis Pro Mobile resmi dari /img/AI_DocBuilder_Mobile_UI_Assets/
 * dengan tata letak kartu bersih, ilustrasi status, tab tersegmen, dan Live SPA sync.
 */

import { store } from '../state.js';
import {
  canApproveOrRejectAcc,
  isCanvasLockedForUser
} from '../auth/permissions.js';
import {
  approveAccCloud,
  cancelAccCloud,
  rejectAccCloud,
  requestAccCloud
} from '../sync/workspaceClient.js';
import { forcePollNow, setLastKnownSyncRev } from '../sync/liveSync.js';
import { showToast } from './notification.js';
import {
  downloadOutgoingItem,
  printOutgoingItem,
  promptOutgoingActionWithChoice,
  renderReadOnlyDocumentHtml
} from './outgoingLettersView.js';
import { docModelToSmartPages } from './smartWordCanvas.js';

let activeFilterTab = 'pending'; // 'pending' | 'approved' | 'rejected' | 'all'
let accSearchQuery = '';
let accSortDesc = true;
let mobileInfoDismissed = false;

function escHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatShortDateTime(isoStr) {
  if (!isoStr) return '-';
  try {
    const d = new Date(isoStr);
    if (isNaN(d.getTime())) return '-';
    return d.toLocaleString('id-ID', {
      day: '2-digit',
      month: 'short',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
  } catch {
    return '-';
  }
}

function isItemPending(item) {
  return item?.verificationStatus === 'pending_acc' || item?.accStatus === 'pending';
}

function isItemApproved(item) {
  return item?.verificationStatus === 'sah' || item?.accStatus === 'approved' || item?.isVerifiedSah;
}

function isItemRejected(item) {
  return item?.verificationStatus === 'rejected' || item?.accStatus === 'rejected';
}

function buildFilteredAccItems(allOutgoing = []) {
  let list = allOutgoing.filter(item => {
    if (activeFilterTab === 'pending') return isItemPending(item);
    if (activeFilterTab === 'approved') return isItemApproved(item) && !isItemPending(item);
    if (activeFilterTab === 'rejected') return isItemRejected(item) && !isItemPending(item);
    return isItemPending(item) || isItemApproved(item) || isItemRejected(item) || item?.accStatus === 'cancelled';
  });

  const q = accSearchQuery.trim().toLowerCase();
  if (q) {
    list = list.filter(item =>
      String(item.title || '').toLowerCase().includes(q) ||
      String(item.nomorSurat || '').toLowerCase().includes(q) ||
      String(item.recipientName || '').toLowerCase().includes(q) ||
      String(item.perihal || '').toLowerCase().includes(q) ||
      String(item.requestedByName || item.recordedBy || '').toLowerCase().includes(q)
    );
  }

  return [...list].sort((a, b) => {
    const seqA = Number(a.sequenceNumber) || 0;
    const seqB = Number(b.sequenceNumber) || 0;
    const yA = Number(a.year || 0);
    const yB = Number(b.year || 0);
    const diff = yA !== yB ? yA - yB : seqA - seqB;
    return accSortDesc ? -diff : diff;
  });
}

function resolveDocTileTone(item) {
  const code = String(item?.typeCode || '').toUpperCase();
  if (isItemRejected(item)) return 'tone-red';
  if (code === 'UND' || code === 'PGL') return 'tone-blue';
  if (code === 'ST' || code === 'KEP') return 'tone-purple';
  return 'tone-green';
}

function renderAccCardsListHtml(items, { canApprove, allPendingCount }) {
  if (items.length === 0) {
    const emptyTitle = accSearchQuery
      ? 'Tidak Ada Surat yang Cocok'
      : activeFilterTab === 'pending'
      ? 'Belum Ada Surat Menunggu ACC Saat Ini'
      : activeFilterTab === 'approved'
      ? 'Belum Ada Surat Disetujui QR'
      : activeFilterTab === 'rejected'
      ? 'Tidak Ada Surat Ditolak'
      : 'Belum Ada Pengajuan ACC';

    const emptyMsg = accSearchQuery
      ? 'Coba kata kunci nomor surat, judul, atau nama penerima lain.'
      : activeFilterTab === 'pending'
      ? (canApprove
          ? 'Semua pengajuan pengesahan QR dari Tata Usaha sudah selesai diproses. Kartu menyala akan langsung muncul otomatis di sini saat ada surat baru diajukan.'
          : 'Saat Anda meminta ACC pengesahan QR Kepala Sekolah, kartu status berputar akan otomatis tampil di sini tanpa perlu refresh.')
      : 'Belum ada catatan pada kategori ini.';

    return `
      <div class="acc-pro-empty-card">
        <img
          src="img/AI_DocBuilder_Mobile_UI_Assets/illustrations/02_empty_queue_approved.svg"
          alt="Ilustrasi Antrean ACC"
          class="acc-empty-artwork"
        />
        <strong>${emptyTitle}</strong>
        <p>${emptyMsg}</p>
        ${activeFilterTab !== 'pending' && allPendingCount > 0 ? `
          <button type="button" id="btn-acc-jump-pending-tab" class="btn btn-sm btn-primary" style="margin-top: 10px;">
            Lihat ${allPendingCount} Surat Menunggu ACC
          </button>
        ` : ''}
      </div>
    `;
  }

  const roleLabels = {
    kepala_sekolah: 'Kepala Sekolah',
    kepsek_dan_tu: 'Kepala Sekolah & TU',
    tata_usaha: 'Tata Usaha (TU)'
  };

  return `
    <div class="acc-pro-cards-list">
      ${items.map(item => {
        const pending = isItemPending(item);
        const approved = isItemApproved(item) && !pending;
        const rejected = isItemRejected(item) && !pending;
        const toneClass = resolveDocTileTone(item);
        const shortType = String(item.typeCode || 'SURAT').toUpperCase().slice(0, 5);
        const cardStateClass = pending
          ? (canApprove ? 'is-glowing' : 'is-tu-waiting')
          : approved
          ? 'is-approved'
          : rejected
          ? 'is-rejected'
          : '';

        const statusPillHtml = pending
          ? (canApprove
              ? '<span class="acc-pro-status status-glow"><span class="acc-glow-dot"></span> Perlu ACC</span>'
              : '<span class="acc-pro-status status-waiting"><span class="acc-live-spinner acc-spinner-sm"></span> Menunggu ACC</span>')
          : approved
          ? '<span class="acc-pro-status status-approved">✓ Disetujui</span>'
          : rejected
          ? '<span class="acc-pro-status status-rejected">✕ Ditolak</span>'
          : '<span class="acc-pro-status status-neutral">Manual</span>';

        return `
          <article class="acc-pro-card ${cardStateClass}" data-acc-id="${escHtml(item.id)}">
            <div class="acc-pro-card-main">
              <div class="acc-pro-doc-tile acc-doc-code-badge ${toneClass}" aria-hidden="true">
                <span class="acc-doc-fold"></span>
                <span class="acc-doc-code-text">${escHtml(shortType)}</span>
              </div>
              <div class="acc-pro-card-info">
                <div class="acc-pro-card-title-row">
                  <h4 class="acc-pro-card-title">${escHtml(item.title || 'SURAT DINAS')}</h4>
                  ${statusPillHtml}
                </div>
                <div class="acc-pro-card-nomor">No: ${escHtml(item.nomorSurat || '-')}</div>
                <div class="acc-pro-card-meta">
                  <span>👤 ${escHtml(item.requestedByName || item.recordedBy || 'Tata Usaha')}</span>
                  <span>📅 ${escHtml(formatShortDateTime(item.requestedAt || item.approvedAt || item.recordedAt))}</span>
                  <span>🎯 ${escHtml(item.recipientName || '-')}</span>
                  <span>✍️ ${escHtml(roleLabels[item.signerRole] || 'Kepala Sekolah')}</span>
                </div>
              </div>
            </div>

            ${item.requestNote ? `<div class="acc-card-note">📝 Catatan TU: ${escHtml(item.requestNote)}</div>` : ''}
            ${rejected && item.rejectedReason ? `<div class="acc-card-note acc-note-danger">❌ Alasan Penolakan: ${escHtml(item.rejectedReason)}</div>` : ''}

            <div class="acc-pro-card-actions">
              <button type="button" class="btn btn-sm btn-secondary btn-acc-preview" data-id="${escHtml(item.id)}">👁️ Pratinjau</button>
              ${pending && canApprove ? `
                <button type="button" class="btn btn-sm btn-primary acc-glow-btn btn-acc-approve" data-id="${escHtml(item.id)}">✅ ACC &amp; Terbitkan QR</button>
                <button type="button" class="btn btn-sm btn-danger btn-acc-reject" data-id="${escHtml(item.id)}">❌ Tolak</button>
              ` : pending ? `
                <button type="button" class="btn btn-sm btn-secondary btn-acc-print-manual" data-id="${escHtml(item.id)}">✍️ TTD Manual</button>
                <button type="button" class="btn btn-sm btn-secondary btn-acc-cancel" data-id="${escHtml(item.id)}">Batalkan</button>
              ` : approved ? `
                <button type="button" class="btn btn-sm btn-primary btn-acc-print-qr" data-id="${escHtml(item.id)}">🖨️ Cetak QR</button>
                <button type="button" class="btn btn-sm btn-success btn-acc-docx-qr" data-id="${escHtml(item.id)}">⬇️ .docx QR</button>
              ` : rejected ? `
                ${!isCanvasLockedForUser(store.getState().authUser) && item.documentSnapshot ? `
                  <button type="button" class="btn btn-sm btn-primary btn-acc-ai-fix" data-id="${escHtml(item.id)}">🤖 Perbaiki AI (Surgical)</button>
                  <button type="button" class="btn btn-sm btn-secondary btn-acc-open-edit" data-id="${escHtml(item.id)}">✏️ Edit di Kanvas</button>
                ` : ''}
                <button type="button" class="btn btn-sm btn-secondary btn-acc-resubmit" data-id="${escHtml(item.id)}">📨 Ajukan Ulang</button>
              ` : ''}
            </div>
          </article>
        `;
      }).join('')}
    </div>
  `;
}

export function renderAccApprovalView(container, { onNavigateTab } = {}) {
  if (!container) return;
  const st = store.getState();
  const user = st.authUser;
  const perms = st.authPermissions || {};
  const allOutgoing = st.outgoingLetters || [];
  const pendingList = allOutgoing.filter(isItemPending);
  const approvedList = allOutgoing.filter(i => isItemApproved(i) && !isItemPending(i));
  const rejectedList = allOutgoing.filter(i => isItemRejected(i) && !isItemPending(i));
  const recentResolved = allOutgoing.filter(i => !isItemPending(i) && (isItemApproved(i) || isItemRejected(i))).slice(0, 3);
  const canApprove = canApproveOrRejectAcc(user);
  const canvasLocked = isCanvasLockedForUser(user);
  const visibleItems = buildFilteredAccItems(allOutgoing);

  container.innerHTML = `
    <div class="view-atmosphere">
      <div class="view-shell acc-approval-shell">

        <!-- HERO CARD DENGAN ILUSTRASI RESMI -->
        <div class="acc-pro-hero-card">
          <div class="acc-pro-hero-grid">
            <div class="acc-pro-hero-text">
              <div class="acc-hero-kicker">MANAJEMEN PERSETUJUAN</div>
              <h2 class="acc-hero-title">Antrean ACC &amp; Pengesahan QR Kepala Sekolah</h2>
              <p class="acc-hero-sub">
                ${canApprove
                  ? 'Daftar surat yang menunggu pengesahan dan tanda tangan digital (Barcode QR) oleh Kepala Sekolah. Tersinkron otomatis ke TU tanpa refresh.'
                  : 'Pantau pengajuan Barcode QR Anda secara live tanpa refresh halaman sambil tetap bebas membuat surat lain.'}
              </p>
            </div>
            <div class="acc-pro-hero-art-wrap" aria-hidden="true">
              <img
                src="img/AI_DocBuilder_Mobile_UI_Assets/illustrations/01_document_qr_approved.svg"
                alt="Pengesahan QR Resmi"
                class="acc-hero-artwork"
              />
            </div>
          </div>

          <!-- 3 BADGE STATUS LIVE -->
          <div class="acc-pro-status-strip">
            <span class="acc-live-badge"><span class="acc-live-dot"></span> Sinkron Live SPA</span>
            <span class="acc-pill-badge ${canvasLocked ? 'is-amber' : 'is-green'}">
              ${canvasLocked ? '🔒 Kanvas Terkunci (HP)' : '🛡️ Akses Penuh'}
            </span>
            <span class="acc-pill-badge is-green">📱 Akses Aktif</span>
          </div>

          <!-- 3 STAT CARDS (MENUNGGU ACC, DISETUJUI, TOTAL) -->
          <div class="acc-pro-kpi-row">
            <div class="acc-pro-kpi kpi-green">
              <div class="acc-pro-kpi-icon" aria-hidden="true">
                <img src="img/AI_DocBuilder_Mobile_UI_Assets/icons/icon_archive.svg" alt="" width="16" height="16" />
              </div>
              <div>
                <div class="acc-pro-kpi-label">Menunggu ACC</div>
                <div class="acc-pro-kpi-val"><strong>${pendingList.length}</strong> <span>Surat</span></div>
              </div>
            </div>
            <div class="acc-pro-kpi kpi-blue">
              <div class="acc-pro-kpi-icon" aria-hidden="true">
                <img src="img/AI_DocBuilder_Mobile_UI_Assets/icons/icon_approval.svg" alt="" width="16" height="16" />
              </div>
              <div>
                <div class="acc-pro-kpi-label">Disetujui QR</div>
                <div class="acc-pro-kpi-val"><strong>${approvedList.length}</strong> <span>Surat</span></div>
              </div>
            </div>
            <div class="acc-pro-kpi kpi-purple">
              <div class="acc-pro-kpi-icon" aria-hidden="true">
                <img src="img/AI_DocBuilder_Mobile_UI_Assets/icons/icon_file.svg" alt="" width="16" height="16" />
              </div>
              <div>
                <div class="acc-pro-kpi-label">Total Surat</div>
                <div class="acc-pro-kpi-val"><strong>${allOutgoing.length}</strong> <span>Surat</span></div>
              </div>
            </div>
          </div>

          <!-- BARIS TOMBOL AKSI UTAMA -->
          <div class="acc-pro-cta-row">
            <button type="button" id="btn-acc-primary-cta" class="btn btn-primary acc-pro-btn-main ${canApprove && pendingList.length > 0 ? 'acc-glow-btn' : ''}">
              <img src="img/AI_DocBuilder_Mobile_UI_Assets/icons/icon_scan.svg" alt="" width="15" height="15" class="acc-btn-svg-white" />
              <span>${canApprove
                ? (pendingList.length > 0 ? `ACC Teratas &amp; QR (${pendingList.length})` : 'Scan &amp; ACC QR')
                : 'Buat Surat Baru'}</span>
            </button>
            <button type="button" id="btn-acc-sync-now" class="btn btn-secondary acc-pro-btn-sec">
              <span>⟳ Cek Update</span>
            </button>
            <button type="button" id="btn-acc-secondary-cta" class="btn btn-secondary acc-pro-btn-sec">
              <span>${canApprove ? '⚙️ Setelan' : '📤 Surat Keluar'}</span>
            </button>
          </div>
        </div>

        <!-- PANEL DAFTAR ANTREAN ACC -->
        <section class="acc-pro-panel">
          <div class="acc-pro-panel-header">
            <div class="acc-pro-panel-title-wrap">
              <img src="img/AI_DocBuilder_Mobile_UI_Assets/icons/icon_file.svg" alt="" width="18" height="18" class="acc-icon-emerald" />
              <h3>${activeFilterTab === 'pending' ? 'Daftar Antrean ACC' : activeFilterTab === 'approved' ? 'Daftar Surat Disetujui QR' : activeFilterTab === 'rejected' ? 'Daftar Surat Ditolak' : 'Semua Riwayat Antrean'}</h3>
            </div>
            <div class="acc-pro-panel-right">
              <span class="acc-count-pill-green">${visibleItems.length} Surat</span>
              <button type="button" id="btn-acc-sort-toggle" class="btn btn-sm btn-secondary acc-sort-btn" title="Ubah urutan terbaru / terlama">
                ${accSortDesc ? '↓ Terbaru' : '↑ Terlama'}
              </button>
            </div>
          </div>

          <!-- TAB FILTER STATUS -->
          <div class="acc-segmented-tabs" role="tablist">
            <button type="button" class="acc-seg-tab ${activeFilterTab === 'pending' ? 'active' : ''}" data-acc-tab="pending">
              Menunggu (${pendingList.length})
            </button>
            <button type="button" class="acc-seg-tab ${activeFilterTab === 'approved' ? 'active' : ''}" data-acc-tab="approved">
              Disetujui (${approvedList.length})
            </button>
            <button type="button" class="acc-seg-tab ${activeFilterTab === 'rejected' ? 'active' : ''}" data-acc-tab="rejected">
              Ditolak (${rejectedList.length})
            </button>
            <button type="button" class="acc-seg-tab ${activeFilterTab === 'all' ? 'active' : ''}" data-acc-tab="all">
              Semua
            </button>
          </div>

          <!-- PENCARIAN CEPAT -->
          <div class="acc-search-toolbar">
            <div class="acc-search-input-wrap">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2">
                <circle cx="11" cy="11" r="8" />
                <line x1="21" y1="21" x2="16.65" y2="16.65" />
              </svg>
              <input type="text" id="input-acc-search" value="${escHtml(accSearchQuery)}" placeholder="Cari nomor, judul, atau penerima..." autocomplete="off" />
            </div>
          </div>

          <!-- DAFTAR KARTU -->
          <div id="acc-cards-list-host">
            ${renderAccCardsListHtml(visibleItems, { canApprove, allPendingCount: pendingList.length })}
          </div>
        </section>

        <!-- RIWAYAT KEPUTUSAN TERAKHIR (RINGKAS) -->
        ${activeFilterTab === 'pending' ? `
          <section class="acc-pro-panel" style="margin-top: 12px;">
            <div class="acc-pro-panel-header">
              <div class="acc-pro-panel-title-wrap">
                <img src="img/AI_DocBuilder_Mobile_UI_Assets/icons/icon_history.svg" alt="" width="18" height="18" class="acc-icon-emerald" />
                <h3>Riwayat Keputusan Terakhir</h3>
              </div>
              <button type="button" id="btn-acc-see-history" class="acc-link-all-btn">Lihat Semua ›</button>
            </div>
            <div style="margin-bottom: 10px;">
              <span class="acc-pill-badge ${perms.allowTuSignPrincipal ? 'is-green' : 'is-amber'}">
                ${perms.allowTuSignPrincipal ? '🔓 Delegasi Kepsek ke TU: Aktif' : '🔒 Delegasi Kepsek ke TU: Terkunci (Mode Antrean ACC)'}
              </span>
            </div>
            ${recentResolved.length === 0 ? `
              <div style="font-size: 12.5px; color: #64748b; padding: 6px 2px;">
                Belum ada keputusan ACC (Disetujui / Ditolak) sebelumnya.
              </div>
            ` : `
              <div class="acc-pro-cards-list">
                ${recentResolved.map(item => {
                  const approved = isItemApproved(item);
                  const tone = resolveDocTileTone(item);
                  const shortCode = String(item.typeCode || 'SURAT').toUpperCase().slice(0, 5);
                  return `
                    <div class="acc-pro-card acc-history-mini-card" data-acc-id="${escHtml(item.id)}">
                      <div class="acc-pro-card-main">
                        <div class="acc-pro-doc-tile acc-doc-code-badge ${tone}" aria-hidden="true">
                          <span class="acc-doc-fold"></span>
                          <span class="acc-doc-code-text">${escHtml(shortCode)}</span>
                        </div>
                        <div class="acc-pro-card-info">
                          <div class="acc-pro-card-title-row">
                            <h4 class="acc-pro-card-title">${escHtml(item.title)}</h4>
                            <span class="acc-pro-status ${approved ? 'status-approved' : 'status-rejected'}">
                              ${approved ? '✓ Disetujui' : '✕ Ditolak'}
                            </span>
                          </div>
                          <div class="acc-pro-card-nomor">No: ${escHtml(item.nomorSurat)}</div>
                          <div class="acc-pro-card-meta">
                            <span>📅 ${escHtml(formatShortDateTime(item.approvedAt || item.rejectedAt || item.recordedAt))}</span>
                            <span>🎯 ${escHtml(item.recipientName || '-')}</span>
                          </div>
                        </div>
                      </div>
                    </div>
                  `;
                }).join('')}
              </div>
            `}
          </section>
        ` : ''}

        <!-- BANNER INFO AKSES MOBILE / KUNCI KANVAS PC -->
        ${!mobileInfoDismissed ? `
          <div class="acc-pro-footer-banner" id="acc-mobile-info-banner">
            <div class="acc-pro-footer-banner-text">
              <div class="acc-pro-footer-kicker">✨ ${canvasLocked ? 'Akses Kepala Sekolah di HP' : 'Akses ACC Kapan Saja, di Mana Saja'}</div>
              <p>
                ${canvasLocked
                  ? 'Kelola persetujuan QR, lihat Surat Keluar, dan atur Setelan Delegasi langsung dari HP. Edit kanvas tetap dikunci khusus Komputer / Laptop / PC.'
                  : 'Kelola permintaan & persetujuan surat QR secara aman, real-time, dan tersinkron ke Google Drive.'}
              </p>
            </div>
            <img
              src="img/AI_DocBuilder_Mobile_UI_Assets/illustrations/03_mobile_qr_approved.svg"
              alt="Verifikasi QR Mobile"
              class="acc-footer-banner-art"
            />
            <button type="button" id="btn-dismiss-mobile-banner" class="acc-pro-dismiss-btn" title="Tutup info">✕</button>
          </div>
        ` : ''}

      </div>
    </div>

    <!-- MODAL PRATINJAU SURAT ANTREAN ACC -->
    <div id="modal-acc-preview" class="modal-overlay hidden">
      <div class="modal-card modal-card-large">
        <div class="modal-header" style="display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #e2e8f0;padding-bottom:12px;margin-bottom:12px;">
          <div>
            <h3 id="modal-acc-preview-title" style="margin:0;font-size:16px;">Pratinjau Surat</h3>
            <p id="modal-acc-preview-sub" style="margin:3px 0 0 0;font-size:12px;color:#64748b;"></p>
          </div>
          <button type="button" id="btn-close-acc-preview" class="btn btn-sm btn-secondary">✕ Tutup</button>
        </div>
        <div id="modal-acc-preview-body" style="max-height:70dvh;overflow-y:auto;background:#e2e8f0;padding:14px;border-radius:10px;"></div>
      </div>
    </div>

    <!-- MODAL PENOLAKAN ACC -->
    <div id="modal-acc-reject" class="modal-overlay hidden">
      <div class="modal-card" style="max-width:460px;">
        <div class="modal-header">
          <h3>❌ Tolak Permintaan ACC Surat</h3>
        </div>
        <p id="reject-acc-target-info" style="font-size:12.5px;color:#475569;margin:8px 0 12px 0;"></p>
        <div class="form-group">
          <label for="input-acc-reject-reason">Alasan Penolakan / Catatan Revisi untuk TU:</label>
          <textarea id="input-acc-reject-reason" class="form-control" rows="3" placeholder="Misal: Perbaiki tanggal pelaksanaan atau nama penerima terlebih dahulu."></textarea>
        </div>
        <div style="display:flex;justify-content:flex-end;gap:8px;margin-top:14px;">
          <button type="button" id="btn-cancel-acc-reject" class="btn btn-secondary">Batal</button>
          <button type="button" id="btn-confirm-acc-reject" class="btn btn-danger">Kirim Penolakan ke TU</button>
        </div>
      </div>
    </div>
  `;

  const refreshListOnly = () => {
    const host = container.querySelector('#acc-cards-list-host');
    if (!host) return;
    const latestAll = store.getState().outgoingLetters || [];
    const itemsNow = buildFilteredAccItems(latestAll);
    const pendingTotal = latestAll.filter(isItemPending).length;
    host.innerHTML = renderAccCardsListHtml(itemsNow, { canApprove, allPendingCount: pendingTotal });
    bindAccCardActions(container, { onNavigateTab });
  };

  const btnPrimary = container.querySelector('#btn-acc-primary-cta');
  if (btnPrimary) {
    btnPrimary.onclick = () => {
      if (canApprove) {
        if (activeFilterTab !== 'pending') {
          activeFilterTab = 'pending';
          renderAccApprovalView(container, { onNavigateTab });
          return;
        }
        const firstCard = container.querySelector('.acc-pro-card.is-glowing');
        if (firstCard) {
          firstCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
          firstCard.querySelector('.btn-acc-preview')?.click();
        } else {
          showToast('Belum ada surat menunggu ACC saat ini.', 'info');
        }
      } else if (typeof onNavigateTab === 'function') {
        onNavigateTab('canvas');
      }
    };
  }

  const btnSec = container.querySelector('#btn-acc-secondary-cta');
  if (btnSec && typeof onNavigateTab === 'function') {
    btnSec.onclick = () => onNavigateTab(canApprove ? 'access' : 'outgoing');
  }

  const btnSeeHistory = container.querySelector('#btn-acc-see-history');
  if (btnSeeHistory) {
    btnSeeHistory.onclick = () => {
      activeFilterTab = 'approved';
      renderAccApprovalView(container, { onNavigateTab });
    };
  }

  const btnSync = container.querySelector('#btn-acc-sync-now');
  if (btnSync) {
    btnSync.onclick = async () => {
      btnSync.disabled = true;
      const changed = await forcePollNow();
      btnSync.disabled = false;
      renderAccApprovalView(container, { onNavigateTab });
      showToast(changed ? '✅ Data antrean ACC diperbarui dari server.' : '✓ Data antrean sudah paling baru.', 'info');
    };
  }

  const btnDismissBanner = container.querySelector('#btn-dismiss-mobile-banner');
  if (btnDismissBanner) {
    btnDismissBanner.onclick = () => {
      mobileInfoDismissed = true;
      container.querySelector('#acc-mobile-info-banner')?.remove();
    };
  }

  container.querySelectorAll('.acc-seg-tab').forEach(tabBtn => {
    tabBtn.onclick = () => {
      activeFilterTab = tabBtn.dataset.accTab || 'pending';
      renderAccApprovalView(container, { onNavigateTab });
    };
  });

  const searchEl = container.querySelector('#input-acc-search');
  if (searchEl) {
    searchEl.addEventListener('input', () => {
      accSearchQuery = searchEl.value;
      refreshListOnly();
    });
  }

  const sortBtn = container.querySelector('#btn-acc-sort-toggle');
  if (sortBtn) {
    sortBtn.onclick = () => {
      accSortDesc = !accSortDesc;
      renderAccApprovalView(container, { onNavigateTab });
    };
  }

  const previewModal = container.querySelector('#modal-acc-preview');
  const closePreview = container.querySelector('#btn-close-acc-preview');
  if (closePreview && previewModal) {
    closePreview.onclick = () => previewModal.classList.add('hidden');
  }

  const rejectModal = container.querySelector('#modal-acc-reject');
  const btnCancelReject = container.querySelector('#btn-cancel-acc-reject');
  if (btnCancelReject && rejectModal) {
    btnCancelReject.onclick = () => rejectModal.classList.add('hidden');
  }

  bindAccCardActions(container, { onNavigateTab });
}

function bindAccCardActions(container, { onNavigateTab } = {}) {
  const previewModal = container.querySelector('#modal-acc-preview');
  const rejectModal = container.querySelector('#modal-acc-reject');
  const rejectInfo = container.querySelector('#reject-acc-target-info');
  const rejectInput = container.querySelector('#input-acc-reject-reason');
  const btnConfirmReject = container.querySelector('#btn-confirm-acc-reject');

  const btnJumpPending = container.querySelector('#btn-acc-jump-pending-tab');
  if (btnJumpPending) {
    btnJumpPending.onclick = () => {
      activeFilterTab = 'pending';
      renderAccApprovalView(container, { onNavigateTab });
    };
  }

  container.querySelectorAll('.btn-acc-preview').forEach(btn => {
    btn.onclick = () => {
      const item = store.getOutgoingLetters().find(x => x.id === btn.dataset.id);
      if (!item || !previewModal) return;
      container.querySelector('#modal-acc-preview-title').textContent = `${item.title} (#${String(item.sequenceNumber || 0).padStart(3, '0')})`;
      container.querySelector('#modal-acc-preview-sub').textContent = `${item.nomorSurat} • ${item.recipientName}`;
      container.querySelector('#modal-acc-preview-body').innerHTML = renderReadOnlyDocumentHtml(item);
      previewModal.classList.remove('hidden');
    };
  });

  container.querySelectorAll('.btn-acc-approve').forEach(btn => {
    btn.onclick = async () => {
      const item = store.getOutgoingLetters().find(x => x.id === btn.dataset.id);
      if (!item) return;
      btn.disabled = true;
      btn.textContent = '⟳ Meng-ACC & QR...';
      const res = await approveAccCloud({
        outgoingId: item.id,
        signerRole: item.signerRole || 'kepala_sekolah',
        verificationOrigin: window.location.origin
      });
      if (!res.ok || !res.item) {
        btn.disabled = false;
        btn.textContent = '✅ ACC & Terbitkan QR';
        showToast(res.error || 'Gagal menyetujui ACC.', 'error');
        return;
      }
      store.upsertOutgoingLetterFromCloud(res.item);
      if (res.syncRev) setLastKnownSyncRev(res.syncRev);
      showToast(`✅ Surat #${String(item.sequenceNumber || 0).padStart(3, '0')} di-ACC! Barcode QR aktif & tersinkron ke TU.`, 'success');
      renderAccApprovalView(container, { onNavigateTab });
    };
  });

  container.querySelectorAll('.btn-acc-reject').forEach(btn => {
    btn.onclick = () => {
      const item = store.getOutgoingLetters().find(x => x.id === btn.dataset.id);
      if (!item || !rejectModal) return;
      if (rejectInfo) rejectInfo.textContent = `#${String(item.sequenceNumber || 0).padStart(3, '0')} • ${item.nomorSurat} (${item.title})`;
      if (rejectInput) rejectInput.value = '';
      rejectModal.classList.remove('hidden');
      if (btnConfirmReject) {
        btnConfirmReject.onclick = async () => {
          btnConfirmReject.disabled = true;
          btnConfirmReject.textContent = 'Mengirim...';
          const reason = rejectInput?.value?.trim() || 'Perlu revisi sebelum pengesahan Kepala Sekolah.';
          const res = await rejectAccCloud({ outgoingId: item.id, reason });
          btnConfirmReject.disabled = false;
          btnConfirmReject.textContent = 'Kirim Penolakan ke TU';
          if (!res.ok || !res.item) {
            showToast(res.error || 'Gagal menolak pengajuan ACC.', 'error');
            return;
          }
          store.upsertOutgoingLetterFromCloud(res.item);
          if (res.syncRev) setLastKnownSyncRev(res.syncRev);
          rejectModal.classList.add('hidden');
          showToast('❌ Pengajuan ACC ditolak dan catatan revisi dikirim ke TU.', 'info');
          renderAccApprovalView(container, { onNavigateTab });
        };
      }
    };
  });

  container.querySelectorAll('.btn-acc-cancel').forEach(btn => {
    btn.onclick = async () => {
      btn.disabled = true;
      btn.textContent = 'Membatalkan...';
      const res = await cancelAccCloud(btn.dataset.id);
      if (!res.ok || !res.item) {
        btn.disabled = false;
        btn.textContent = 'Batalkan';
        showToast(res.error || 'Gagal membatalkan pengajuan ACC.', 'error');
        return;
      }
      store.upsertOutgoingLetterFromCloud(res.item);
      if (res.syncRev) setLastKnownSyncRev(res.syncRev);
      showToast('Pengajuan ACC dibatalkan.', 'info');
      renderAccApprovalView(container, { onNavigateTab });
    };
  });

  container.querySelectorAll('.btn-acc-print-manual').forEach(btn => {
    btn.onclick = () => {
      const item = store.getOutgoingLetters().find(x => x.id === btn.dataset.id);
      if (item) printOutgoingItem(item, { mode: 'manual' });
    };
  });

  container.querySelectorAll('.btn-acc-print-qr').forEach(btn => {
    btn.onclick = () => {
      const item = store.getOutgoingLetters().find(x => x.id === btn.dataset.id);
      if (item) promptOutgoingActionWithChoice(item, 'print', container);
    };
  });

  container.querySelectorAll('.btn-acc-docx-qr').forEach(btn => {
    btn.onclick = () => {
      const item = store.getOutgoingLetters().find(x => x.id === btn.dataset.id);
      if (item) downloadOutgoingItem(item, { mode: 'qr' });
    };
  });

  const openRejectedInEditor = (item, { autoRunAi = false } = {}) => {
    const snap = item?.documentSnapshot;
    if (!snap) {
      showToast('Snapshot isi surat tidak tersedia untuk diedit.', 'warning');
      return;
    }
    const clonedSnap = JSON.parse(JSON.stringify(snap));
    store.setPendingKepsekRevision({
      outgoingId: item.id,
      nomorSurat: item.nomorSurat || '-',
      title: item.title || 'Surat Dinas',
      reason: item.rejectedReason || 'Perlu penyesuaian data sebelum pengesahan QR.',
      rejectedByName: item.rejectedByName || 'Kepala Sekolah',
      rejectedAt: item.rejectedAt || '',
      signerRole: item.signerRole || 'kepala_sekolah',
      autoRunAi
    });

    const isSmartDoc = Boolean(
      Array.isArray(clonedSnap.pages) ||
      String(clonedSnap?.letterData?.templateId || '').startsWith('smart_')
    );
    if (isSmartDoc) {
      const smartPages = Array.isArray(clonedSnap.pages)
        ? clonedSnap.pages
        : docModelToSmartPages(clonedSnap, store.getSchoolProfile()).pages;
      store.setState({
        smartDocument: {
          pages: smartPages,
          meta: clonedSnap.meta || { title: item.title, nomorSurat: item.nomorSurat }
        },
        activeSmartVaultId: null
      }, false);
      onNavigateTab?.('smart-canvas');
    } else {
      store.setState({
        document: clonedSnap,
        docxViewMode: 'canvas',
        activeVaultId: null
      }, false);
      onNavigateTab?.('canvas');
    }
  };

  container.querySelectorAll('.btn-acc-ai-fix').forEach(btn => {
    btn.onclick = () => {
      const item = store.getOutgoingLetters().find(x => x.id === btn.dataset.id);
      if (item) openRejectedInEditor(item, { autoRunAi: true });
    };
  });

  container.querySelectorAll('.btn-acc-open-edit').forEach(btn => {
    btn.onclick = () => {
      const item = store.getOutgoingLetters().find(x => x.id === btn.dataset.id);
      if (item) openRejectedInEditor(item, { autoRunAi: false });
    };
  });

  container.querySelectorAll('.btn-acc-resubmit').forEach(btn => {
    btn.onclick = async () => {
      const item = store.getOutgoingLetters().find(x => x.id === btn.dataset.id);
      if (!item) return;
      btn.disabled = true;
      btn.textContent = 'Mengajukan...';
      const res = await requestAccCloud({
        outgoingId: item.id,
        signerRole: item.signerRole || 'kepala_sekolah',
        requestNote: 'Pengajuan ulang ACC Kepala Sekolah'
      });
      if (!res.ok || !res.item) {
        btn.disabled = false;
        btn.textContent = '📨 Ajukan Ulang';
        showToast(res.error || 'Gagal mengajukan ulang ACC.', 'error');
        return;
      }
      store.upsertOutgoingLetterFromCloud(res.item);
      if (res.syncRev) setLastKnownSyncRev(res.syncRev);
      showToast('📨 Pengajuan ulang dikirim ke antrean ACC Kepala Sekolah!', 'success');
      renderAccApprovalView(container, { onNavigateTab });
    };
  });
}
