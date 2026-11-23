/**
 * js/ui/incomingLettersView.js - Dashboard & Buku Arsip Surat Masuk
 * Mendukung Upload Gambar, Upload Dokumen PDF, dan Foto Langsung dari Kamera,
 * tersimpan langsung ke Google Drive (Bridge Apps Script) & terindeks berdasarkan Kode & Nama Surat.
 */

import { store } from '../state.js';
import { showToast } from './notification.js';
import { canDeleteIncomingItem } from '../auth/permissions.js';
import {
  APPS_SCRIPT_FULL_BRIDGE_CODE,
  APPS_SCRIPT_INCOMING_SNIPPET,
  INCOMING_CATEGORIES,
  deriveDefaultTitleFromFilename,
  formatFileSizeLabel,
  prepareIncomingFilePayload,
  scoreIncomingMatch,
  sortIncomingLetters
} from '../letters/incoming.js';
import {
  deleteIncomingCloud,
  fetchIncomingFileCloud,
  getStoredAuthToken,
  recordIncomingCloud,
  syncIncomingDriveCloud
} from '../sync/workspaceClient.js';
import { setLastKnownSyncRev } from '../sync/liveSync.js';
import { scanIncomingLetterWithAi } from '../ai/client.js';
import { downloadFile } from '../utils/download.js';
import {
  AGENDA_MONTH_LABELS,
  filterLettersByPeriod,
  collectAvailableAgendaYears,
  formatPeriodLabel,
  exportIncomingLettersCsv,
  buildPrintableAgendaHtml
} from '../letters/agenda.js';
import { uiIcon } from './icons.js';

let incomingSearchQuery = '';
let incomingKindFilter = 'all';
let incomingDriveFilter = 'all';
let incomingMonthFilter = 'all';
let incomingYearFilter = 'all';
let incomingSortMode = 'newest';
let activeCameraStream = null;

function escapeHtml(value) {
  const div = document.createElement('div');
  div.textContent = String(value ?? '');
  return div.innerHTML;
}

function stopActiveCameraStream() {
  if (activeCameraStream && typeof activeCameraStream.getTracks === 'function') {
    try {
      activeCameraStream.getTracks().forEach(track => {
        try { track.stop(); } catch {}
      });
    } catch {}
  }
  activeCameraStream = null;
}

function base64ToUint8Array(base64Str = '') {
  const binary = atob(base64Str);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function formatReadableDate(val = '') {
  const s = String(val || '').trim();
  if (!s) return '-';
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const d = new Date(s.length === 10 ? `${s}T00:00:00` : s);
    if (!Number.isNaN(d.getTime())) {
      return d.toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' });
    }
  }
  return s;
}

function getKindMeta(item = {}) {
  const k = item.fileKind || (String(item.mimeType || '').includes('pdf') ? 'pdf' : 'image');
  if (k === 'pdf') {
    return { label: 'PDF', icon: uiIcon('filePdf', 12), rowIcon: uiIcon('filePdf', 20), chipClass: 'incoming-chip-pdf' };
  }
  if (k === 'photo' || item.captureSource === 'camera_photo') {
    return { label: 'Foto Kamera', icon: uiIcon('camera', 12), rowIcon: uiIcon('camera', 20), chipClass: 'incoming-chip-photo' };
  }
  return { label: 'Gambar', icon: uiIcon('gallery', 12), rowIcon: uiIcon('gallery', 20), chipClass: 'incoming-chip-img' };
}

function getFilteredIncomingLetters() {
  const list = store.getIncomingLetters();
  const q = incomingSearchQuery.trim();

  const periodFiltered = filterLettersByPeriod(list, {
    month: incomingMonthFilter,
    year: incomingYearFilter,
    getDateValue: (item) => item.tanggalSurat || item.receivedAt || item.createdAt
  });

  let filtered = periodFiltered.filter(item => {
    if (incomingKindFilter !== 'all') {
      const k = item.fileKind || (item.captureSource === 'camera_photo' ? 'photo' : 'image');
      if (k !== incomingKindFilter) return false;
    }
    if (incomingDriveFilter === 'synced' && !item.driveSynced) return false;
    if (incomingDriveFilter === 'pending' && item.driveSynced) return false;
    if (!q) return true;
    return scoreIncomingMatch(item, q) > 0;
  });

  if (q) {
    filtered = filtered
      .map(item => ({ item, score: scoreIncomingMatch(item, q) }))
      .sort((a, b) => b.score - a.score)
      .map(x => x.item);
    return filtered;
  }

  return sortIncomingLetters(filtered, incomingSortMode);
}

export function renderIncomingLettersView(container) {
  if (!container) return;
  stopActiveCameraStream();

  const state = store.getState();
  const allIncoming = store.getIncomingLetters();
  const drive = state.driveStatus || {};
  const nextCode = store.getNextIncomingCodeInfo();

  const syncedCount = allIncoming.filter(i => i.driveSynced).length;
  const pendingDriveCount = allIncoming.length - syncedCount;
  const pdfCount = allIncoming.filter(i => i.fileKind === 'pdf').length;
  const photoImgCount = allIncoming.length - pdfCount;

  const now = new Date();
  const curMonth = now.getMonth();
  const curYear = now.getFullYear();
  const monthCount = allIncoming.filter(item => {
    const d = item.receivedAt || item.createdAt ? new Date(item.receivedAt || item.createdAt) : null;
    return d && !Number.isNaN(d.getTime()) && d.getMonth() === curMonth && d.getFullYear() === curYear;
  }).length;

  const yearsInUse = collectAvailableAgendaYears(allIncoming, (item) => item.tanggalSurat || item.receivedAt || item.createdAt);
  const SVG_INBOX = uiIcon('inbox', 24);
  const SVG_SEARCH = uiIcon('search', 15);

  container.innerHTML = `
    <div class="view-atmosphere">
      <div class="view-shell incoming-view-container">

        <!-- HEADER DASHBOARD SURAT MASUK -->
        <div class="view-page-header" style="margin-bottom: 18px;">
          <div class="view-header-lead">
            <div class="view-header-badge">${SVG_INBOX}</div>
            <div>
              <h2 class="view-title">
                <span>Arsip &amp; Daftar Surat Masuk</span>
                <span class="status-chip ${drive.configured ? 'status-chip-green' : 'status-chip-amber'}">
                  ${drive.configured ? `${uiIcon('cloudCheck', 12)}<span>Sync Google Drive Aktif</span>` : `${uiIcon('clock', 12)}<span>Bridge Drive Menunggu Webhook</span>`}
                </span>
              </h2>
              <p class="view-subtitle" style="max-width: 740px;">
                Unggah <strong>Gambar</strong>, dokumen <strong>PDF</strong>, atau <strong>Foto Langsung</strong> dari kamera. Berkas otomatis diarsipkan ke folder <strong>Google Drive</strong> sekolah dan tercatat lengkap dengan <strong>Kode Surat</strong> &amp; <strong>Nama Surat</strong> pada daftar ini.
              </p>
            </div>
          </div>

          <div class="view-header-actions">
            <button type="button" id="btn-in-open-upload" class="btn btn-primary" style="font-size: 13px; padding: 9px 15px; border-radius: 8px;">
              ${uiIcon('upload', 14)}
              <span>Upload Gambar / PDF</span>
            </button>
            <button type="button" id="btn-in-open-camera" class="btn btn-blue-cta" style="font-size: 13px; padding: 9px 15px; border-radius: 8px;">
              ${uiIcon('camera', 14)}
              <span>Foto Langsung</span>
            </button>
            ${pendingDriveCount > 0 ? `
              <button type="button" id="btn-in-sync-all-drive" class="btn btn-sm btn-secondary" style="border-color:#93c5fd; background:#eff6ff; color:#1d4ed8; font-weight:700;" title="Kirim ulang seluruh surat yang masih menunggu ke Google Drive">
                ${uiIcon('refresh', 13)}
                <span>Kirim Semua (${pendingDriveCount}) ke Drive</span>
              </button>
            ` : ''}
            <button type="button" id="btn-in-export-csv" class="btn btn-sm btn-secondary" title="Unduh Buku Agenda Surat Masuk ke Excel/CSV">
              ${uiIcon('spreadsheet', 13)}
              <span>Ekspor Excel/CSV</span>
            </button>
            <button type="button" id="btn-in-print-rekap" class="btn btn-sm btn-secondary" title="Cetak Rekap Buku Agenda Surat Masuk resmi">
              ${uiIcon('printer', 13)}
              <span>Cetak Rekap</span>
            </button>
            <button type="button" id="btn-in-open-gas-guide" class="btn btn-sm btn-secondary" title="Panduan &amp; Kode Google Apps Script (archive_incoming)">
              ${uiIcon('settings', 13)}
              <span>Script Drive</span>
            </button>
          </div>
        </div>

        <!-- KARTU STATISTIK DASHBOARD -->
        <div class="kpi-grid-4">
          <div class="kpi-card accent-left-green">
            <div class="kpi-icon-box kpi-icon-green">
              ${uiIcon('inbox', 18)}
            </div>
            <div style="flex: 1; min-width: 0;">
              <div style="font-size: 12px; color: #475569; font-weight: 600;">Total Surat Masuk</div>
              <div class="kpi-value" style="font-size: 26px; font-weight: 800; color: #0f172a; line-height: 1.2; margin-top: 3px;">${allIncoming.length}</div>
              <div style="font-size: 11px; color: #64748b; margin-top: 3px;">Kode berikutnya: <strong>${escapeHtml(nextCode.kodeSurat)}</strong></div>
            </div>
          </div>

          <div class="kpi-card accent-left-blue">
            <div class="kpi-icon-box kpi-icon-blue">
              ${uiIcon('cloud', 18)}
            </div>
            <div style="flex: 1; min-width: 0;">
              <div style="font-size: 12px; color: #475569; font-weight: 600;">Tersimpan di Google Drive</div>
              <div class="kpi-value" style="font-size: 26px; font-weight: 800; color: #1d4ed8; line-height: 1.2; margin-top: 3px;">${syncedCount}</div>
              <div style="font-size: 11px; color: #64748b; margin-top: 3px;">
                <a href="${escapeHtml(drive.rootFolderUrl || `https://drive.google.com/drive/folders/${drive.rootFolderId || '19nFdoJ4VOEgFmUe_heAFxoFEt-9U9Crw'}`)}" target="_blank" rel="noopener" style="color:#1d4ed8; font-weight:600; text-decoration:none; display:inline-flex; align-items:center; gap:4px;"><span>Buka Folder Drive</span>${uiIcon('externalLink', 11)}</a>
              </div>
            </div>
          </div>

          <div class="kpi-card accent-left-amber">
            <div class="kpi-icon-box kpi-icon-amber">
              ${uiIcon('calendar', 18)}
            </div>
            <div style="flex: 1; min-width: 0;">
              <div style="font-size: 12px; color: #475569; font-weight: 600;">Surat Masuk Bulan Ini</div>
              <div class="kpi-value" style="font-size: 26px; font-weight: 800; color: #b45309; line-height: 1.2; margin-top: 3px;">${monthCount}</div>
              <div style="font-size: 11px; color: #64748b; margin-top: 3px;">Periode bulan berjalan</div>
            </div>
          </div>

          <div class="kpi-card accent-left-purple">
            <div class="kpi-icon-box kpi-icon-purple">
              ${uiIcon('folder', 18)}
            </div>
            <div style="flex: 1; min-width: 0;">
              <div style="font-size: 12px; color: #475569; font-weight: 600;">Komposisi Berkas</div>
              <div class="kpi-value" style="font-size: 22px; font-weight: 800; color: #6d28d9; line-height: 1.25; margin-top: 3px;">${photoImgCount} <span style="font-size:12px;font-weight:600;color:#64748b;">Foto/Img</span> · ${pdfCount} <span style="font-size:12px;font-weight:600;color:#64748b;">PDF</span></div>
              <div style="font-size: 11px; color: #64748b; margin-top: 3px;">Pratinjau &amp; unduhan siap pakai</div>
            </div>
          </div>
        </div>

        <!-- FILTER & PENCARIAN -->
        <div class="filter-bar-card">
          <div class="icon-input-wrap" style="flex: 1; min-width: 240px;">
            <span class="icon-input-leading" style="color:#64748b;">${SVG_SEARCH}</span>
            <input type="text" id="input-incoming-search" class="icon-input-control"
              value="${escapeHtml(incomingSearchQuery)}"
              placeholder="Cari kode surat, nama surat, pengirim, nomor asli, atau catatan...">
          </div>
          <div class="icon-input-wrap" style="width: 185px;">
            <span class="icon-input-leading">${uiIcon('paperclip', 14)}</span>
            <select id="select-incoming-kind" class="icon-input-control" style="cursor: pointer;">
              <option value="all" ${incomingKindFilter === 'all' ? 'selected' : ''}>Semua Jenis Berkas</option>
              <option value="photo" ${incomingKindFilter === 'photo' ? 'selected' : ''}>Foto Kamera</option>
              <option value="image" ${incomingKindFilter === 'image' ? 'selected' : ''}>Gambar (JPG/PNG)</option>
              <option value="pdf" ${incomingKindFilter === 'pdf' ? 'selected' : ''}>Dokumen PDF</option>
            </select>
          </div>
          <div class="icon-input-wrap" style="width: 195px;">
            <span class="icon-input-leading">${uiIcon('cloud', 14)}</span>
            <select id="select-incoming-drive" class="icon-input-control" style="cursor: pointer;">
              <option value="all" ${incomingDriveFilter === 'all' ? 'selected' : ''}>Semua Status Drive</option>
              <option value="synced" ${incomingDriveFilter === 'synced' ? 'selected' : ''}>Tersimpan di Drive</option>
              <option value="pending" ${incomingDriveFilter === 'pending' ? 'selected' : ''}>Belum Sync Drive</option>
            </select>
          </div>
          <div class="icon-input-wrap filter-ctrl-sm">
            <span class="icon-input-leading">${uiIcon('calendar', 14)}</span>
            <select id="select-incoming-month" class="icon-input-control" style="cursor: pointer;">
              <option value="all" ${incomingMonthFilter === 'all' ? 'selected' : ''}>Semua Bulan</option>
              ${AGENDA_MONTH_LABELS.map((mName, idx) => {
                const mVal = String(idx + 1).padStart(2, '0');
                return `<option value="${mVal}" ${incomingMonthFilter === mVal ? 'selected' : ''}>${mName}</option>`;
              }).join('')}
            </select>
          </div>
          <div class="icon-input-wrap filter-ctrl-xs">
            <span class="icon-input-leading">${uiIcon('hash', 14)}</span>
            <select id="select-incoming-year" class="icon-input-control" style="cursor: pointer;">
              <option value="all" ${incomingYearFilter === 'all' ? 'selected' : ''}>Semua Tahun</option>
              ${yearsInUse.map(y => `<option value="${escapeHtml(y)}" ${incomingYearFilter === y ? 'selected' : ''}>${escapeHtml(y)}</option>`).join('')}
            </select>
          </div>
          <div class="icon-input-wrap" style="width: 205px;">
            <span class="icon-input-leading">${uiIcon('sort', 14)}</span>
            <select id="select-incoming-sort" class="icon-input-control" style="cursor: pointer;">
              <option value="newest" ${incomingSortMode === 'newest' ? 'selected' : ''}>Urut: Terbaru Diterima</option>
              <option value="oldest" ${incomingSortMode === 'oldest' ? 'selected' : ''}>Urut: Terlama</option>
              <option value="code_asc" ${incomingSortMode === 'code_asc' ? 'selected' : ''}>Urut: Kode Surat (A - Z)</option>
              <option value="name_asc" ${incomingSortMode === 'name_asc' ? 'selected' : ''}>Urut: Nama Surat (A - Z)</option>
            </select>
          </div>
        </div>

        <!-- AREA DAFTAR SURAT MASUK -->
        <div id="incoming-list-area"></div>

      </div>
    </div>

    <!-- MODAL INPUT / UPLOAD / FOTO LANGSUNG SURAT MASUK -->
    <div id="modal-add-incoming" class="modal-overlay hidden">
      <div class="modal-card" style="max-width: 760px; width: 95%; max-height: 92vh; display: flex; flex-direction: column;">
        <div class="modal-header" style="display: flex; justify-content: space-between; align-items: center; padding-bottom: 10px; border-bottom: 1px solid #e2e8f0;">
          <div>
            <h3 style="margin: 0; font-size: 16px; color: #0f172a;">${uiIcon('inbox', 17)}<span>Tambah &amp; Arsipkan Surat Masuk ke Drive</span></h3>
            <div style="font-size: 12px; color: #64748b; margin-top: 2px;">Pilih berkas Gambar, PDF, atau langsung Foto menggunakan kamera perangkat</div>
          </div>
          <button type="button" id="btn-close-add-incoming" class="btn btn-sm btn-secondary btn-square-sm" aria-label="Tutup">${uiIcon('x', 15)}</button>
        </div>

        <form id="form-add-incoming" style="flex: 1; overflow-y: auto; padding: 14px 2px; display: grid; gap: 14px;">
          <!-- PILIH SUMBER BERKAS (3 MODE) -->
          <div class="incoming-source-tabs">
            <button type="button" class="incoming-src-btn active" data-in-src="upload_image">${uiIcon('gallery', 14)}<span>Upload Gambar</span></button>
            <button type="button" class="incoming-src-btn" data-in-src="upload_pdf">${uiIcon('filePdf', 14)}<span>Upload Dokumen PDF</span></button>
            <button type="button" class="incoming-src-btn" data-in-src="camera_photo">${uiIcon('camera', 14)}<span>Foto Langsung Kamera</span></button>
          </div>

          <input type="file" id="input-in-file-image" accept="image/*" class="hidden" />
          <input type="file" id="input-in-file-pdf" accept="application/pdf,.pdf" class="hidden" />
          <input type="file" id="input-in-file-cam-native" accept="image/*" capture="environment" class="hidden" />

          <!-- PANEL FILE / KAMERA & PREVIEW -->
          <div id="in-upload-dropzone" class="incoming-dropzone">
            <div id="in-dropzone-idle" class="incoming-dropzone-idle">
              <div style="color:#156c3b; display:inline-flex;" id="in-drop-icon">${uiIcon('gallery', 30)}</div>
              <div style="font-weight: 700; color: #0f172a; font-size: 13.5px;" id="in-drop-title">Pilih atau tarik berkas Gambar surat ke sini</div>
              <div style="font-size: 11.5px; color: #64748b;" id="in-drop-sub">Format JPG, PNG, WEBP — otomatis dioptimalkan agar tajam &amp; ringan masuk Drive</div>
              <button type="button" id="btn-in-pick-file" class="btn btn-sm btn-primary" style="margin-top: 6px;">${uiIcon('folder', 13)}<span>Pilih Berkas dari Perangkat</span></button>
            </div>

            <div id="in-camera-stage" class="incoming-camera-stage hidden">
              <video id="in-camera-video" autoplay playsinline muted class="incoming-camera-video"></video>
              <div id="in-camera-fallback-msg" class="hidden" style="padding: 14px; text-align: center; font-size: 12px; color: #475569;">
                Kamera web belum diizinkan atau tidak terdeteksi. Gunakan tombol <strong>Buka Kamera HP / Pilih Foto</strong> di bawah.
              </div>
              <div class="incoming-camera-controls">
                <button type="button" id="btn-in-snap-photo" class="btn btn-primary btn-sm">${uiIcon('camera', 13)}<span>Jepret Foto Sekarang</span></button>
                <button type="button" id="btn-in-switch-cam" class="btn btn-secondary btn-sm">${uiIcon('refresh', 13)}<span>Ganti Kamera</span></button>
                <button type="button" id="btn-in-native-cam" class="btn btn-secondary btn-sm">${uiIcon('smartphone', 13)}<span>Buka Kamera Native HP</span></button>
              </div>
            </div>

            <div id="in-selected-preview" class="incoming-selected-preview hidden">
              <div id="in-preview-thumb-wrap" class="incoming-preview-thumb-wrap"></div>
              <div style="flex: 1; min-width: 0;">
                <div id="in-preview-filename" style="font-weight: 700; color: #0f172a; font-size: 13px; word-break: break-word;">-</div>
                <div id="in-preview-filemeta" style="font-size: 11.5px; color: #475569; margin-top: 2px;">-</div>
                <div id="in-ocr-status" style="font-size: 11.5px; color: #15803d; margin-top: 3px; font-weight: 600;" class="hidden"></div>
              </div>
              <div style="display: flex; gap: 6px; flex-wrap: wrap;">
                <button type="button" id="btn-in-ai-scan" class="btn btn-sm btn-primary" title="Baca isi foto/PDF dengan AI OCR &amp; isi otomatis">${uiIcon('sparkle', 13)}<span>Scan AI (OCR)</span></button>
                <button type="button" id="btn-in-retake-change" class="btn btn-sm btn-secondary">${uiIcon('refresh', 13)}<span>Ganti / Foto Ulang</span></button>
              </div>
            </div>
          </div>

          <!-- METADATA KODE & NAMA SURAT -->
          <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 10px;">
            <div class="form-group" style="margin: 0;">
              <label style="font-size: 12px; font-weight: 700;">Kode Surat Masuk *</label>
              <input type="text" id="input-in-kode" class="form-control" required placeholder="Contoh: SM-001/XI/2026" />
            </div>
            <div class="form-group" style="margin: 0;">
              <label style="font-size: 12px; font-weight: 700;">Nama / Judul Surat Masuk *</label>
              <input type="text" id="input-in-nama" class="form-control" required placeholder="Contoh: Undangan Rapat Dinas Pendidikan" />
            </div>
          </div>

          <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px;">
            <div class="form-group" style="margin: 0;">
              <label style="font-size: 12px; font-weight: 600;">Asal / Pengirim Surat</label>
              <input type="text" id="input-in-pengirim" class="form-control" placeholder="Contoh: Cabang Dinas Wilayah XI Garut" />
            </div>
            <div class="form-group" style="margin: 0;">
              <label style="font-size: 12px; font-weight: 600;">Nomor Surat Asli (di Kertas/PDF)</label>
              <input type="text" id="input-in-nomor-asli" class="form-control" placeholder="Contoh: 421/102/DISDIK/2026 (opsional)" />
            </div>
          </div>

          <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 10px;">
            <div class="form-group" style="margin: 0;">
              <label style="font-size: 12px; font-weight: 600;">Tanggal Surat / Diterima</label>
              <input type="date" id="input-in-tanggal" class="form-control" />
            </div>
            <div class="form-group" style="margin: 0;">
              <label style="font-size: 12px; font-weight: 600;">Kategori Surat</label>
              <select id="select-in-kategori" class="form-control">
                ${INCOMING_CATEGORIES.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('')}
              </select>
            </div>
          </div>

          <div class="form-group" style="margin: 0;">
            <label style="font-size: 12px; font-weight: 600;">Catatan / Disposisi Singkat (Opsional)</label>
            <input type="text" id="input-in-catatan" class="form-control" placeholder="Contoh: Diteruskan ke Wakasek Kurikulum untuk tindak lanjut" />
          </div>

          <div style="display: flex; justify-content: flex-end; gap: 8px; border-top: 1px solid #e2e8f0; padding-top: 12px;">
            <button type="button" id="btn-cancel-add-incoming" class="btn btn-secondary">Batal</button>
            <button type="submit" id="btn-submit-add-incoming" class="btn btn-primary" style="font-weight: 700;">
              ${uiIcon('cloudUpload', 14)}
              <span>Simpan &amp; Arsipkan ke Google Drive</span>
            </button>
          </div>
        </form>
      </div>
    </div>

    <!-- MODAL PRATINJAU BERKAS SURAT MASUK (GAMBAR / FOTO / PDF) -->
    <div id="modal-incoming-preview" class="modal-overlay hidden">
      <div class="modal-card" style="max-width: 920px; width: 95%; max-height: 92vh; display: flex; flex-direction: column;">
        <div class="modal-header" style="display: flex; justify-content: space-between; align-items: center; padding-bottom: 10px; border-bottom: 1px solid #e2e8f0;">
          <div>
            <h3 id="modal-in-prev-title" style="margin: 0; font-size: 16px; color: #0f172a;">Pratinjau Surat Masuk</h3>
            <div id="modal-in-prev-sub" style="font-size: 12px; color: #64748b; margin-top: 2px;"></div>
          </div>
          <button type="button" id="btn-close-in-prev" class="btn btn-sm btn-secondary">${uiIcon('x', 13)}<span>Tutup</span></button>
        </div>
        <div id="modal-in-prev-body" style="flex: 1; overflow-y: auto; padding: 14px; background: #f1f5f9; border-radius: 6px; margin: 12px 0; display: flex; flex-direction: column; align-items: center; justify-content: center; min-height: 360px;"></div>
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px; padding-top: 8px; border-top: 1px solid #e2e8f0;">
          <div id="modal-in-prev-drive-status" style="font-size: 12px; color: #475569;"></div>
          <div style="display: flex; gap: 8px; flex-wrap: wrap;">
            <a id="btn-in-prev-open-drive" href="#" target="_blank" rel="noopener" class="btn btn-sm btn-secondary hidden">${uiIcon('externalLink', 12)}<span>Buka di Google Drive</span></a>
            <button type="button" id="btn-in-prev-download" class="btn btn-sm btn-primary">${uiIcon('download', 13)}<span>Unduh Berkas</span></button>
          </div>
        </div>
      </div>
    </div>

    <!-- MODAL KONFIRMASI HAPUS SURAT MASUK -->
    <div id="modal-confirm-delete-in" class="modal-overlay hidden">
      <div class="modal-card" style="max-width: 440px; width: 92%;">
        <div class="modal-header" style="display: flex; justify-content: space-between; align-items: center; padding-bottom: 10px; border-bottom: 1px solid #e2e8f0;">
          <h3 style="margin: 0; font-size: 16px; color: #b91c1c;">${uiIcon('trash', 16)}<span>Konfirmasi Hapus Surat Masuk</span></h3>
          <button type="button" id="btn-close-del-in" class="btn btn-sm btn-secondary btn-square-sm" aria-label="Tutup">${uiIcon('x', 14)}</button>
        </div>
        <div style="padding: 14px 0; font-size: 13px; color: #334155;">
          <p style="margin: 0 0 10px 0;">Anda yakin ingin menghapus catatan Surat Masuk berikut dari daftar?</p>
          <div id="del-in-detail-box" style="background: #fef2f2; border: 1px solid #fecaca; border-radius: 6px; padding: 10px 12px; font-size: 12.5px; color: #7f1d1d;"></div>
        </div>
        <div style="display: flex; justify-content: flex-end; gap: 8px; border-top: 1px solid #e2e8f0; padding-top: 12px;">
          <button type="button" id="btn-cancel-del-in" class="btn btn-secondary">Batal</button>
          <button type="button" id="btn-confirm-del-in" class="btn btn-danger" style="font-weight: 700;">${uiIcon('trash', 13)}<span>Hapus Surat Masuk</span></button>
        </div>
      </div>
    </div>

    <!-- MODAL PANDUAN UPDATE GOOGLE APPS SCRIPT (CODE.GS) -->
    <div id="modal-in-gas-guide" class="modal-overlay hidden">
      <div class="modal-card" style="max-width: 780px; width: 95%; max-height: 92vh; display: flex; flex-direction: column;">
        <div class="modal-header" style="display: flex; justify-content: space-between; align-items: center; padding-bottom: 10px; border-bottom: 1px solid #e2e8f0;">
          <div>
            <h3 style="margin: 0; font-size: 16px; color: #0f172a;">${uiIcon('settings', 16)}<span>Update Google Apps Script (<code>Code.gs</code>) untuk Surat Masuk</span></h3>
            <div style="font-size: 12px; color: #64748b; margin-top: 2px;">Diperlukan satu kali agar <code>script.google.com</code> mengenali perintah <code>archive_incoming</code> &amp; menyimpan file ke folder <strong>03_Surat_Masuk</strong></div>
          </div>
          <button type="button" id="btn-close-gas-guide" class="btn btn-sm btn-secondary btn-square-sm" aria-label="Tutup">${uiIcon('x', 14)}</button>
        </div>
        <div style="flex: 1; overflow-y: auto; padding: 12px 0; font-size: 12.5px; color: #334155; display: grid; gap: 10px;">
          <div id="in-gas-reason-banner" class="hidden" style="background: #fffbeb; border: 1px solid #fcd34d; color: #92400e; border-radius: 8px; padding: 10px 12px; font-size: 12px; line-height: 1.45;">
            <strong style="display:inline-flex;align-items:center;gap:6px;">${uiIcon('alertTriangle', 14)}<span>Kenapa muncul pesan &ldquo;Aksi tidak dikenali: archive_incoming&rdquo;?</span></strong><br>
            Web App Google Apps Script Anda masih memakai kode lama (hanya ada Surat Keluar/Gudang/Kop). Berkas Surat Masuk Anda <strong>sudah aman di Cloud Server</strong> &mdash; silakan ikuti 3 langkah di bawah agar masuk ke Google Drive:
          </div>
          <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px 12px; font-size: 12px; line-height: 1.5;">
            <strong>Langkah Update di <code>script.google.com</code> (1 Menit):</strong>
            <ol style="margin: 4px 0 0 18px; padding: 0;">
              <li>Buka <a href="https://script.google.com/" target="_blank" rel="noopener" style="color:#15803d;font-weight:700;">script.google.com</a> (login akun <code>tuancraf1@gmail.com</code>) dan buka project Bridge Drive Anda.</li>
              <li>Salin kode dari tab <strong>Sisipkan ke Code.gs Lama</strong> di bawah (paling aman karena Secret Anda tidak berubah). Tempelkan 2 baris <code>if (action === 'archive_incoming')</code> tepat sebelum pesan <em>Aksi tidak dikenali</em>, dan fungsi pembantunya di paling bawah file.</li>
              <li>Klik tombol biru <strong>Deploy &rarr; Kelola deployment &rarr; ikon Pensil (Edit) &rarr; pilih Versi baru &rarr; Terapkan</strong>. Setelah selesai, klik tombol <strong>Kirim ke Drive</strong> pada daftar Surat Masuk!</li>
            </ol>
          </div>
          <div style="display: flex; gap: 6px; flex-wrap: wrap;">
            <button type="button" class="btn btn-sm btn-primary gas-tab-btn active" data-gas-tab="snippet">${uiIcon('puzzle', 13)}<span>Sisipkan ke Code.gs Lama (Direkomendasikan)</span></button>
            <button type="button" class="btn btn-sm btn-secondary gas-tab-btn" data-gas-tab="full">${uiIcon('code', 13)}<span>Kode Code.gs Lengkap (All-in-One)</span></button>
          </div>
          <pre id="pre-gas-incoming-code" style="margin: 0; padding: 12px; background: #0f172a; color: #e2e8f0; border-radius: 8px; font-size: 11px; line-height: 1.45; overflow-x: auto; max-height: 300px;">${escapeHtml(APPS_SCRIPT_INCOMING_SNIPPET)}</pre>
        </div>
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px; border-top: 1px solid #e2e8f0; padding-top: 10px;">
          <a href="https://script.google.com/" target="_blank" rel="noopener" class="btn btn-secondary btn-sm">${uiIcon('globe', 13)}<span>Buka script.google.com</span>${uiIcon('externalLink', 11)}</a>
          <div style="display: flex; gap: 8px;">
            <button type="button" id="btn-copy-gas-incoming" class="btn btn-primary btn-sm">${uiIcon('clipboard', 13)}<span>Salin Kode Aktif</span></button>
            <button type="button" id="btn-done-gas-guide" class="btn btn-secondary btn-sm">Tutup</button>
          </div>
        </div>
      </div>
    </div>
  `;

  updateIncomingListInPlace(container);
  bindIncomingViewEvents(container);
}

function updateIncomingListInPlace(container) {
  const listArea = container.querySelector('#incoming-list-area');
  if (!listArea) return;

  const filtered = getFilteredIncomingLetters();
  const totalCount = store.getIncomingLetters().length;

  if (totalCount === 0) {
    listArea.innerHTML = `
      <div class="ui-card" style="padding: 46px 24px; text-align: center; display: grid; gap: 12px; justify-items: center;">
        <div class="view-header-badge" style="width: 58px; height: 58px; border-radius: 16px;">${uiIcon('inbox', 26)}</div>
        <div style="font-size: 16px; font-weight: 800; color: #0f172a;">Belum Ada Surat Masuk Tersimpan</div>
        <div style="font-size: 13px; color: #475569; max-width: 520px;">
          Mulai arsipkan surat masuk sekolah dengan mengunggah <strong>Gambar</strong>, file <strong>PDF</strong>, atau langsung <strong>Foto dengan Kamera</strong>. Dokumen akan langsung diunggah ke Google Drive dan kodenya muncul dalam daftar.
        </div>
        <div style="display: flex; gap: 8px; flex-wrap: wrap; justify-content: center; margin-top: 4px;">
          <button type="button" class="btn btn-primary btn-sm" id="btn-empty-in-upload">${uiIcon('upload', 13)}<span>Upload Gambar / PDF</span></button>
          <button type="button" class="btn btn-blue-cta btn-sm" id="btn-empty-in-cam">${uiIcon('camera', 13)}<span>Foto Langsung</span></button>
        </div>
      </div>
    `;
    listArea.querySelector('#btn-empty-in-upload')?.addEventListener('click', () => container.querySelector('#btn-in-open-upload')?.click());
    listArea.querySelector('#btn-empty-in-cam')?.addEventListener('click', () => container.querySelector('#btn-in-open-camera')?.click());
    return;
  }

  if (filtered.length === 0) {
    listArea.innerHTML = `
      <div class="ui-card" style="padding: 36px 20px; text-align: center; color: #64748b; font-size: 13.5px;">
        Tidak ada Surat Masuk yang cocok dengan pencarian/filter saat ini.
      </div>
    `;
    return;
  }

  const state = store.getState();
  const canDelete = canDeleteIncomingItem(state.authUser, state.authPermissions);

  listArea.innerHTML = `
    <div class="ui-card" style="overflow: hidden; padding: 0;">
      <div style="overflow-x: auto;">
        <table class="enterprise-table incoming-list-table">
          <thead>
            <tr>
              <th style="width: 68px;">Berkas</th>
              <th style="min-width: 155px;">Kode Surat</th>
              <th style="min-width: 240px;">Nama Surat &amp; Pengirim</th>
              <th style="min-width: 115px;">Tanggal</th>
              <th style="min-width: 170px;">Status Google Drive</th>
              <th style="text-align: right; min-width: 195px;">Tindakan</th>
            </tr>
          </thead>
          <tbody>
            ${filtered.map(item => {
              const kind = getKindMeta(item);
              const thumbHtml = item.previewThumbDataUrl
                ? `<img src="${escapeHtml(item.previewThumbDataUrl)}" alt="${escapeHtml(item.namaSurat)}" class="incoming-row-thumb" />`
                : `<div class="incoming-row-icon ${kind.chipClass}">${kind.rowIcon}</div>`;

              const driveHtml = item.driveSynced
                ? `
                  <div style="display:flex; flex-direction:column; gap:4px; align-items:flex-start;">
                    <span class="status-chip status-chip-green" style="font-size:11px; padding:2px 8px;">${uiIcon('cloudCheck', 12)}<span>Tersimpan di Drive</span></span>
                    ${item.driveUrl ? `<a href="${escapeHtml(item.driveUrl)}" target="_blank" rel="noopener" style="font-size:11px; color:#15803d; font-weight:600; text-decoration:none; display:inline-flex; align-items:center; gap:3px;"><span>Buka di Drive</span>${uiIcon('externalLink', 11)}</a>` : ''}
                  </div>
                `
                : `
                  <div style="display:flex; flex-direction:column; gap:4px; align-items:flex-start;">
                    <span class="status-chip status-chip-amber" style="font-size:11px; padding:2px 8px;" title="${escapeHtml(item.driveError || 'Belum tersinkron ke Google Drive')}">${uiIcon('clock', 12)}<span>Cloud Lokal / Belum Drive</span></span>
                    <button type="button" class="btn btn-sm btn-secondary btn-in-sync-drive" data-id="${escapeHtml(item.id)}" style="font-size:10.5px; padding:3px 8px;">${uiIcon('refresh', 11)}<span>Kirim ke Drive</span></button>
                  </div>
                `;

              return `
                <tr class="incoming-list-row">
                  <td>
                    <button type="button" class="incoming-thumb-trigger btn-in-preview" data-id="${escapeHtml(item.id)}" title="Klik untuk pratinjau berkas">
                      ${thumbHtml}
                    </button>
                  </td>
                  <td>
                    <div style="font-family: Consolas, monospace; font-weight: 800; color: #0f172a; font-size: 13px;">${escapeHtml(item.kodeSurat || '-')}</div>
                    <div style="margin-top: 4px;">
                      <span class="incoming-kind-pill ${kind.chipClass}">${kind.icon}<span>${escapeHtml(kind.label)}</span></span>
                    </div>
                  </td>
                  <td>
                    <div style="font-weight: 700; color: #0f172a; font-size: 13.5px; line-height: 1.35;">${escapeHtml(item.namaSurat)}</div>
                    <div style="font-size: 12px; color: #475569; margin-top: 3px;">
                      ${item.pengirim ? `Dari: <strong>${escapeHtml(item.pengirim)}</strong>` : 'Pengirim belum diisi'}
                      ${item.nomorSuratAsli ? ` · No. Asli: <span style="font-family:Consolas,monospace;">${escapeHtml(item.nomorSuratAsli)}</span>` : ''}
                    </div>
                    <div style="font-size: 11px; color: #64748b; margin-top: 2px;">
                      ${escapeHtml(item.kategori || 'Lainnya')} · ${escapeHtml(item.fileName || '-')} (${formatFileSizeLabel(item.fileSizeBytes)})
                      ${item.catatan ? ` · <em>${escapeHtml(item.catatan)}</em>` : ''}
                    </div>
                  </td>
                  <td style="font-size: 12.5px; color: #334155;">
                    <div>${escapeHtml(formatReadableDate(item.tanggalSurat || item.receivedAt))}</div>
                    <div style="font-size: 10.5px; color: #94a3b8; margin-top: 2px;">oleh ${escapeHtml(item.createdByName || item.createdBy || '-')}</div>
                  </td>
                  <td>${driveHtml}</td>
                  <td style="text-align: right; white-space: nowrap;">
                    <div class="table-action-group">
                      <button type="button" class="btn btn-sm btn-secondary btn-in-preview" data-id="${escapeHtml(item.id)}" title="Lihat Pratinjau Berkas">${uiIcon('eye', 13)}<span>Lihat</span></button>
                      <button type="button" class="btn btn-sm btn-secondary btn-square-sm btn-in-download" data-id="${escapeHtml(item.id)}" title="Unduh Berkas">${uiIcon('download', 13)}</button>
                      ${canDelete ? `<button type="button" class="btn btn-sm btn-danger btn-square-sm btn-in-delete" data-id="${escapeHtml(item.id)}" title="Hapus Surat Masuk">${uiIcon('trash', 13)}</button>` : ''}
                    </div>
                  </td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>
      </div>
    </div>
  `;

  bindIncomingRowActions(container);
}

function bindIncomingViewEvents(container) {
  const searchInput = container.querySelector('#input-incoming-search');
  const kindSelect = container.querySelector('#select-incoming-kind');
  const driveSelect = container.querySelector('#select-incoming-drive');
  const monthSelect = container.querySelector('#select-incoming-month');
  const yearSelect = container.querySelector('#select-incoming-year');
  const sortSelect = container.querySelector('#select-incoming-sort');
  const btnExportCsv = container.querySelector('#btn-in-export-csv');
  const btnPrintRekap = container.querySelector('#btn-in-print-rekap');

  searchInput?.addEventListener('input', () => {
    incomingSearchQuery = searchInput.value;
    updateIncomingListInPlace(container);
  });
  kindSelect?.addEventListener('change', () => {
    incomingKindFilter = kindSelect.value;
    updateIncomingListInPlace(container);
  });
  driveSelect?.addEventListener('change', () => {
    incomingDriveFilter = driveSelect.value;
    updateIncomingListInPlace(container);
  });
  monthSelect?.addEventListener('change', () => {
    incomingMonthFilter = monthSelect.value;
    updateIncomingListInPlace(container);
  });
  yearSelect?.addEventListener('change', () => {
    incomingYearFilter = yearSelect.value;
    updateIncomingListInPlace(container);
  });
  sortSelect?.addEventListener('change', () => {
    incomingSortMode = sortSelect.value;
    updateIncomingListInPlace(container);
  });

  btnExportCsv?.addEventListener('click', () => {
    const filteredItems = getFilteredIncomingLetters();
    if (filteredItems.length === 0) {
      showToast('Tidak ada data Surat Masuk pada filter periode yang dipilih.', 'warning');
      return;
    }
    const periodLabel = formatPeriodLabel(incomingMonthFilter, incomingYearFilter);
    const schoolName = store.getSchoolProfile()?.schoolName || 'SMK IT Al-Hawari';
    const csv = exportIncomingLettersCsv(filteredItems, { schoolName, periodLabel });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const suffix = `${incomingYearFilter === 'all' ? 'semua' : incomingYearFilter}${incomingMonthFilter !== 'all' ? `-${incomingMonthFilter}` : ''}`;
    downloadFile(blob, `buku-agenda-surat-masuk-${suffix}.csv`);
    showToast(`✓ Buku Agenda Surat Masuk (${filteredItems.length} baris) berhasil diekspor ke CSV!`, 'success');
  });

  btnPrintRekap?.addEventListener('click', () => {
    const filteredItems = getFilteredIncomingLetters();
    const periodLabel = formatPeriodLabel(incomingMonthFilter, incomingYearFilter);
    const html = buildPrintableAgendaHtml({
      kind: 'incoming',
      school: store.getSchoolProfile(),
      periodLabel,
      items: filteredItems
    });
    const win = window.open('', '_blank', 'width=1024,height=720');
    if (!win) {
      showToast('Popup diblokir browser — izinkan popup untuk mencetak rekap Buku Agenda.', 'warning');
      return;
    }
    win.document.open();
    win.document.write(html);
    win.document.close();
    setTimeout(() => {
      try { win.focus(); win.print(); } catch {}
    }, 280);
  });

  // Modal Tambah Surat Masuk (Upload / Kamera)
  const addModal = container.querySelector('#modal-add-incoming');
  const formEl = container.querySelector('#form-add-incoming');
  const fileImgInput = container.querySelector('#input-in-file-image');
  const filePdfInput = container.querySelector('#input-in-file-pdf');
  const fileNativeCamInput = container.querySelector('#input-in-file-cam-native');

  const dropzoneEl = container.querySelector('#in-upload-dropzone');
  const idleBox = container.querySelector('#in-dropzone-idle');
  const camStage = container.querySelector('#in-camera-stage');
  const videoEl = container.querySelector('#in-camera-video');
  const camFallbackMsg = container.querySelector('#in-camera-fallback-msg');
  const selectedBox = container.querySelector('#in-selected-preview');
  const previewThumbWrap = container.querySelector('#in-preview-thumb-wrap');
  const previewNameEl = container.querySelector('#in-preview-filename');
  const previewMetaEl = container.querySelector('#in-preview-filemeta');
  const ocrStatusEl = container.querySelector('#in-ocr-status');
  const btnAiScan = container.querySelector('#btn-in-ai-scan');
  const dropIconEl = container.querySelector('#in-drop-icon');
  const dropTitleEl = container.querySelector('#in-drop-title');
  const dropSubEl = container.querySelector('#in-drop-sub');

  const kodeInput = container.querySelector('#input-in-kode');
  const namaInput = container.querySelector('#input-in-nama');
  const pengirimInput = container.querySelector('#input-in-pengirim');
  const nomorAsliInput = container.querySelector('#input-in-nomor-asli');
  const tanggalInput = container.querySelector('#input-in-tanggal');
  const kategoriSelect = container.querySelector('#select-in-kategori');
  const catatanInput = container.querySelector('#input-in-catatan');
  const submitBtn = container.querySelector('#btn-submit-add-incoming');

  let activeSourceMode = 'upload_image';
  let currentFacingMode = 'environment';
  let preparedFilePayload = null;
  let ocrAbortController = null;

  const runIncomingAiScan = async (payload, { silentError = false, overwriteExisting = false } = {}) => {
    if (!payload?.localDataUrl && !payload?.fileBase64) return;
    if (ocrAbortController) {
      try { ocrAbortController.abort(); } catch {}
    }
    ocrAbortController = new AbortController();
    const myCtrl = ocrAbortController;

    if (btnAiScan) {
      btnAiScan.disabled = true;
      btnAiScan.innerHTML = `${uiIcon('loader', 13, 'spin-icon')}<span>Memindai AI...</span>`;
    }
    if (ocrStatusEl) {
      ocrStatusEl.classList.remove('hidden');
      ocrStatusEl.style.color = '#1d4ed8';
      ocrStatusEl.textContent = 'AI OCR sedang membaca pengirim, nomor asli, tanggal & perihal surat...';
    }

    try {
      const dataUrl = payload.localDataUrl || `data:${payload.mimeType || 'application/octet-stream'};base64,${payload.fileBase64}`;
      const scanned = await scanIncomingLetterWithAi({
        dataUrl,
        mimeType: payload.mimeType,
        fileName: payload.fileName,
        signal: myCtrl.signal
      });
      if (myCtrl.signal.aborted || preparedFilePayload !== payload) return;

      const defaultFromFile = deriveDefaultTitleFromFilename(payload.fileName);
      if (scanned.namaSurat && (overwriteExisting || !namaInput.value.trim() || namaInput.value.trim() === defaultFromFile)) {
        namaInput.value = scanned.namaSurat;
      }
      if (scanned.pengirim && (overwriteExisting || !pengirimInput.value.trim())) {
        pengirimInput.value = scanned.pengirim;
      }
      if (scanned.nomorSuratAsli && (overwriteExisting || !nomorAsliInput.value.trim())) {
        nomorAsliInput.value = scanned.nomorSuratAsli;
      }
      if (scanned.tanggalSurat && (overwriteExisting || !tanggalInput.value || tanggalInput.value === new Date().toISOString().slice(0, 10))) {
        tanggalInput.value = scanned.tanggalSurat;
      }
      if (scanned.kategori && kategoriSelect) {
        kategoriSelect.value = scanned.kategori;
      }
      if (scanned.catatan && (overwriteExisting || !catatanInput.value.trim())) {
        catatanInput.value = scanned.catatan;
      }

      if (ocrStatusEl) {
        ocrStatusEl.style.color = '#15803d';
        ocrStatusEl.textContent = 'AI OCR berhasil mengisi kolom metadata. Silakan periksa & simpan!';
      }
      if (!silentError) {
        showToast('AI OCR selesai membaca isi surat masuk!', 'success');
      }
    } catch (err) {
      if (myCtrl.signal.aborted) return;
      if (ocrStatusEl) {
        ocrStatusEl.style.color = '#b45309';
        ocrStatusEl.textContent = silentError
          ? 'Pemindaian AI dilewati (silakan isi manual atau klik Scan AI).'
          : `Gagal Scan AI: ${err.message}`;
      }
      if (!silentError) {
        showToast(`Gagal Scan AI: ${err.message}`, 'warning');
      }
    } finally {
      if (btnAiScan) {
        btnAiScan.disabled = false;
        btnAiScan.innerHTML = `${uiIcon('sparkle', 13)}<span>Scan AI (OCR)</span>`;
      }
    }
  };

  btnAiScan?.addEventListener('click', () => {
    if (preparedFilePayload) {
      runIncomingAiScan(preparedFilePayload, { silentError: false, overwriteExisting: true });
    }
  });

  const setSelectedFileUi = (payload) => {
    preparedFilePayload = payload;
    stopActiveCameraStream();
    idleBox?.classList.add('hidden');
    camStage?.classList.add('hidden');
    selectedBox?.classList.remove('hidden');

    previewNameEl.textContent = payload.fileName;
    previewMetaEl.textContent = `${payload.fileKind === 'pdf' ? 'Dokumen PDF' : payload.fileKind === 'photo' ? 'Foto Kamera' : 'Gambar'} · ${formatFileSizeLabel(payload.fileSizeBytes)}`;

    if (payload.fileKind === 'pdf') {
      previewThumbWrap.innerHTML = `<div class="incoming-row-icon incoming-chip-pdf" style="width:52px;height:52px;">${uiIcon('filePdf', 24)}</div>`;
    } else {
      const thumbSrc = payload.previewThumbDataUrl || payload.localDataUrl;
      previewThumbWrap.innerHTML = `<img src="${escapeHtml(thumbSrc)}" alt="Pratinjau" style="width:56px;height:56px;object-fit:cover;border-radius:8px;border:1px solid #cbd5e1;" />`;
    }

    if (!namaInput.value.trim()) {
      const guessed = deriveDefaultTitleFromFilename(payload.fileName);
      if (guessed && payload.fileKind !== 'photo') {
        namaInput.value = guessed;
      }
    }

    runIncomingAiScan(payload, { silentError: true, overwriteExisting: false });
  };

  const startLiveCamera = async () => {
    stopActiveCameraStream();
    idleBox?.classList.add('hidden');
    selectedBox?.classList.add('hidden');
    camStage?.classList.remove('hidden');
    camFallbackMsg?.classList.add('hidden');
    videoEl?.classList.remove('hidden');

    if (!navigator?.mediaDevices?.getUserMedia) {
      videoEl?.classList.add('hidden');
      camFallbackMsg?.classList.remove('hidden');
      return;
    }

    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: { ideal: currentFacingMode }, width: { ideal: 1600 }, height: { ideal: 1200 } },
        audio: false
      });
      activeCameraStream = stream;
      if (videoEl) {
        videoEl.srcObject = stream;
        await videoEl.play().catch(() => {});
      }
    } catch {
      videoEl?.classList.add('hidden');
      camFallbackMsg?.classList.remove('hidden');
    }
  };

  const switchSourceMode = (mode) => {
    activeSourceMode = mode;
    container.querySelectorAll('.incoming-src-btn').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.inSrc === mode);
    });

    if (preparedFilePayload) {
      return;
    }

    if (mode === 'camera_photo') {
      startLiveCamera();
      return;
    }

    stopActiveCameraStream();
    camStage?.classList.add('hidden');
    selectedBox?.classList.add('hidden');
    idleBox?.classList.remove('hidden');

    if (mode === 'upload_pdf') {
      dropIconEl.innerHTML = uiIcon('filePdf', 30);
      dropTitleEl.textContent = 'Pilih atau tarik berkas Dokumen PDF surat ke sini';
      dropSubEl.textContent = 'Format .pdf (maksimal 3,8 MB) — disimpan langsung ke Google Drive';
    } else {
      dropIconEl.innerHTML = uiIcon('gallery', 30);
      dropTitleEl.textContent = 'Pilih atau tarik berkas Gambar scan/foto surat ke sini';
      dropSubEl.textContent = 'Format JPG, PNG, WEBP — otomatis dioptimalkan agar tajam & ringan masuk Drive';
    }
  };

  const openAddModal = (initialMode = 'upload_image') => {
    preparedFilePayload = null;
    formEl?.reset();
    const nextCode = store.getNextIncomingCodeInfo();
    if (kodeInput) kodeInput.value = nextCode.kodeSurat;
    if (tanggalInput) tanggalInput.value = new Date().toISOString().slice(0, 10);
    selectedBox?.classList.add('hidden');
    addModal?.classList.remove('hidden');
    switchSourceMode(initialMode);
  };

  const closeAddModal = () => {
    if (ocrAbortController) {
      try { ocrAbortController.abort(); } catch {}
    }
    ocrStatusEl?.classList.add('hidden');
    stopActiveCameraStream();
    addModal?.classList.add('hidden');
  };

  container.querySelector('#btn-in-open-upload')?.addEventListener('click', () => openAddModal('upload_image'));
  container.querySelector('#btn-in-open-camera')?.addEventListener('click', () => openAddModal('camera_photo'));
  container.querySelector('#btn-close-add-incoming')?.addEventListener('click', closeAddModal);
  container.querySelector('#btn-cancel-add-incoming')?.addEventListener('click', closeAddModal);

  container.querySelectorAll('.incoming-src-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      preparedFilePayload = null;
      selectedBox?.classList.add('hidden');
      switchSourceMode(btn.dataset.inSrc || 'upload_image');
    });
  });

  container.querySelector('#btn-in-pick-file')?.addEventListener('click', () => {
    if (activeSourceMode === 'upload_pdf') filePdfInput?.click();
    else fileImgInput?.click();
  });

  const handleFileChosen = async (file, sourceHint) => {
    if (!file) return;
    try {
      const prepared = await prepareIncomingFilePayload(file, { sourceHint });
      setSelectedFileUi(prepared);
    } catch (err) {
      showToast(err?.message || 'Gagal memproses berkas yang dipilih.', 'error');
    }
  };

  fileImgInput?.addEventListener('change', (e) => {
    const f = e.target.files?.[0];
    e.target.value = '';
    if (f) handleFileChosen(f, 'upload_image');
  });
  filePdfInput?.addEventListener('change', (e) => {
    const f = e.target.files?.[0];
    e.target.value = '';
    if (f) handleFileChosen(f, 'upload_pdf');
  });
  fileNativeCamInput?.addEventListener('change', (e) => {
    const f = e.target.files?.[0];
    e.target.value = '';
    if (f) handleFileChosen(f, 'camera_photo');
  });

  // Drag & Drop support on dropzone
  dropzoneEl?.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzoneEl.classList.add('is-drag-over');
  });
  dropzoneEl?.addEventListener('dragleave', () => {
    dropzoneEl.classList.remove('is-drag-over');
  });
  dropzoneEl?.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzoneEl.classList.remove('is-drag-over');
    const f = e.dataTransfer?.files?.[0];
    if (!f) return;
    const isPdf = String(f.type || '').includes('pdf') || f.name.toLowerCase().endsWith('.pdf');
    switchSourceMode(isPdf ? 'upload_pdf' : 'upload_image');
    handleFileChosen(f, isPdf ? 'upload_pdf' : 'upload_image');
  });

  // Camera buttons
  container.querySelector('#btn-in-native-cam')?.addEventListener('click', () => {
    fileNativeCamInput?.click();
  });

  container.querySelector('#btn-in-switch-cam')?.addEventListener('click', () => {
    currentFacingMode = currentFacingMode === 'environment' ? 'user' : 'environment';
    startLiveCamera();
  });

  container.querySelector('#btn-in-snap-photo')?.addEventListener('click', async () => {
    if (!videoEl || !videoEl.videoWidth || !videoEl.videoHeight) {
      fileNativeCamInput?.click();
      return;
    }
    const c = document.createElement('canvas');
    c.width = videoEl.videoWidth;
    c.height = videoEl.videoHeight;
    c.getContext('2d').drawImage(videoEl, 0, 0, c.width, c.height);
    c.toBlob(async (blob) => {
      if (!blob) return;
      const kodeSafe = (kodeInput?.value || 'SM').replace(/[^a-zA-Z0-9_-]/g, '-');
      await handleFileChosen(blob, 'camera_photo');
      if (preparedFilePayload) {
        preparedFilePayload.fileName = `${kodeSafe}_foto_${Date.now()}.jpg`;
        previewNameEl.textContent = preparedFilePayload.fileName;
      }
    }, 'image/jpeg', 0.9);
  });

  container.querySelector('#btn-in-retake-change')?.addEventListener('click', () => {
    preparedFilePayload = null;
    selectedBox?.classList.add('hidden');
    switchSourceMode(activeSourceMode);
  });

  // Submit form Surat Masuk -> Simpan + Upload langsung ke Google Drive
  formEl?.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!preparedFilePayload?.fileBase64) {
      showToast('Pilih berkas Gambar/PDF atau jepret Foto terlebih dahulu.', 'warning');
      return;
    }
    const kodeSurat = kodeInput?.value?.trim() || '';
    const namaSurat = namaInput?.value?.trim() || '';
    if (!kodeSurat || !namaSurat) {
      showToast('Kode Surat Masuk dan Nama / Judul Surat wajib diisi.', 'warning');
      return;
    }

    const origBtnHtml = submitBtn.innerHTML;
    submitBtn.disabled = true;
    submitBtn.innerHTML = `${uiIcon('loader', 14, 'spin-icon')}<span>Mengunggah ke Server &amp; Google Drive...</span>`;

    try {
      const id = `in-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
      const entry = {
        id,
        kodeSurat,
        namaSurat,
        nomorSuratAsli: nomorAsliInput?.value?.trim() || '',
        pengirim: pengirimInput?.value?.trim() || '',
        tanggalSurat: tanggalInput?.value || new Date().toISOString().slice(0, 10),
        kategori: kategoriSelect?.value || 'Lainnya',
        catatan: catatanInput?.value?.trim() || '',
        captureSource: activeSourceMode,
        fileName: preparedFilePayload.fileName,
        mimeType: preparedFilePayload.mimeType,
        fileKind: preparedFilePayload.fileKind,
        fileSizeBytes: preparedFilePayload.fileSizeBytes,
        previewThumbDataUrl: preparedFilePayload.previewThumbDataUrl
      };

      const hasToken = Boolean(getStoredAuthToken());
      if (hasToken) {
        const res = await recordIncomingCloud({
          entry,
          fileBase64: preparedFilePayload.fileBase64
        });
        if (!res.ok) {
          throw new Error(res.error || 'Gagal menyimpan Surat Masuk ke server.');
        }
        if (res.syncRev) setLastKnownSyncRev(res.syncRev);
        if (res.driveStatus) store.setAuthSessionData({ driveStatus: res.driveStatus });
        const savedItem = res.item || entry;
        store.saveIncomingFileCache(savedItem.id, preparedFilePayload);
        if (Array.isArray(res.incomingLetters)) {
          store.setState({ incomingLetters: res.incomingLetters }, false);
          store.saveIncomingLetters(res.incomingLetters);
        } else {
          store.upsertIncomingLetterLocal(savedItem);
        }

        closeAddModal();
        renderIncomingLettersView(container);

        if (savedItem.driveSynced) {
          showToast(`✓ Surat Masuk "${savedItem.kodeSurat}" berhasil disimpan & terunggah ke Google Drive!`, 'success');
        } else if (res.driveSync?.needGasUpdate || /aksi tidak dikenali|archive_incoming/i.test(savedItem.driveError || '')) {
          showToast(`✓ Surat "${savedItem.kodeSurat}" tersimpan di Daftar Cloud. Update Code.gs Apps Script untuk sinkron ke Drive.`, 'warning');
          openGasGuideModal(true);
        } else {
          showToast(`✓ Surat Masuk "${savedItem.kodeSurat}" tercatat di daftar. Sinkronisasi Drive: ${savedItem.driveError || 'periksa Bridge Drive'}.`, 'warning');
        }
      } else {
        // Fallback mode lokal browser bila belum masuk akun cloud
        const localItem = {
          ...entry,
          sequenceNumber: store.getNextIncomingCodeInfo().sequenceNumber,
          year: String(new Date().getFullYear()),
          driveSynced: false,
          driveError: 'Mode offline / belum login akun cloud',
          receivedAt: new Date().toISOString(),
          createdAt: new Date().toISOString()
        };
        store.saveIncomingFileCache(localItem.id, preparedFilePayload);
        store.upsertIncomingLetterLocal(localItem);
        closeAddModal();
        renderIncomingLettersView(container);
        showToast(`Surat Masuk "${localItem.kodeSurat}" tersimpan lokal (login ke cloud untuk auto-upload Drive).`, 'info');
      }
    } catch (err) {
      showToast(err?.message || 'Terjadi kesalahan saat menyimpan Surat Masuk.', 'error');
    } finally {
      if (document.body.contains(submitBtn)) {
        submitBtn.disabled = false;
        submitBtn.innerHTML = origBtnHtml;
      }
    }
  });

  // Modal Panduan Script Google Drive
  const gasModal = container.querySelector('#modal-in-gas-guide');
  const gasReasonBanner = container.querySelector('#in-gas-reason-banner');
  const preGasCode = container.querySelector('#pre-gas-incoming-code');
  let activeGasTab = 'snippet';

  const openGasGuideModal = (fromError = false) => {
    if (gasReasonBanner) gasReasonBanner.classList.toggle('hidden', !fromError);
    gasModal?.classList.remove('hidden');
  };
  container._openGasGuideModal = openGasGuideModal;

  const closeGasModal = () => gasModal?.classList.add('hidden');
  container.querySelector('#btn-in-open-gas-guide')?.addEventListener('click', () => openGasGuideModal(false));
  container.querySelector('#btn-close-gas-guide')?.addEventListener('click', closeGasModal);
  container.querySelector('#btn-done-gas-guide')?.addEventListener('click', closeGasModal);

  container.querySelectorAll('.gas-tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      activeGasTab = btn.dataset.gasTab === 'full' ? 'full' : 'snippet';
      container.querySelectorAll('.gas-tab-btn').forEach(b => {
        const isOn = b.dataset.gasTab === activeGasTab;
        b.classList.toggle('btn-primary', isOn);
        b.classList.toggle('btn-secondary', !isOn);
        b.classList.toggle('active', isOn);
      });
      if (preGasCode) {
        preGasCode.textContent = activeGasTab === 'full' ? APPS_SCRIPT_FULL_BRIDGE_CODE : APPS_SCRIPT_INCOMING_SNIPPET;
      }
    });
  });

  container.querySelector('#btn-copy-gas-incoming')?.addEventListener('click', async () => {
    const codeText = activeGasTab === 'full' ? APPS_SCRIPT_FULL_BRIDGE_CODE : APPS_SCRIPT_INCOMING_SNIPPET;
    try {
      await navigator.clipboard.writeText(codeText);
      showToast(activeGasTab === 'full'
        ? '✓ Kode Code.gs lengkap berhasil disalin!'
        : '✓ Kode handler Surat Masuk (archive_incoming) berhasil disalin!', 'success');
    } catch {
      showToast('Pilih & salin kode secara manual dari kotak di atas.', 'info');
    }
  });

  // Bulk sync semua yang tertunda ke Google Drive
  const btnSyncAll = container.querySelector('#btn-in-sync-all-drive');
  btnSyncAll?.addEventListener('click', async () => {
    if (!getStoredAuthToken()) {
      showToast('Silakan masuk akun terlebih dahulu untuk mengirim ke Google Drive.', 'warning');
      return;
    }
    const pendingList = store.getIncomingLetters().filter(i => !i.driveSynced);
    if (pendingList.length === 0) return;

    btnSyncAll.disabled = true;
    btnSyncAll.innerHTML = `${uiIcon('loader', 13, 'spin-icon')}<span>Mengirim 0/${pendingList.length}...</span>`;
    let okCount = 0;
    let needUpdate = false;
    let lastErr = '';

    try {
      for (let idx = 0; idx < pendingList.length; idx++) {
        const target = pendingList[idx];
        btnSyncAll.innerHTML = `${uiIcon('loader', 13, 'spin-icon')}<span>Mengirim ${idx + 1}/${pendingList.length}...</span>`;
        const cached = await store.getIncomingFileCache(target.id);
        const res = await syncIncomingDriveCloud(target.id, cached?.fileBase64 || '');
        if (res.ok) {
          okCount += 1;
          if (res.syncRev) setLastKnownSyncRev(res.syncRev);
          if (res.driveStatus) store.setAuthSessionData({ driveStatus: res.driveStatus });
          if (Array.isArray(res.incomingLetters)) {
            store.setState({ incomingLetters: res.incomingLetters }, false);
            store.saveIncomingLetters(res.incomingLetters);
          } else if (res.item) {
            store.upsertIncomingLetterLocal(res.item);
          }
        } else {
          lastErr = res.error || '';
          if (res.needGasUpdate || res.data?.needGasUpdate || /aksi tidak dikenali|archive_incoming/i.test(lastErr)) {
            needUpdate = true;
            break;
          }
        }
      }

      if (needUpdate) {
        showToast('Apps Script Google Drive belum mendukung archive_incoming. Ikuti panduan singkat berikut untuk mengaktifkannya.', 'warning');
        openGasGuideModal(true);
      } else if (okCount > 0) {
        renderIncomingLettersView(container);
        showToast(`✓ ${okCount} surat masuk berhasil dikirim ke Google Drive!`, 'success');
      } else if (lastErr) {
        showToast(lastErr, 'error');
      }
    } finally {
      if (document.body.contains(btnSyncAll)) {
        btnSyncAll.disabled = false;
        btnSyncAll.innerHTML = `${uiIcon('refresh', 13)}<span>Kirim Semua (${store.getIncomingLetters().filter(i => !i.driveSynced).length}) ke Drive</span>`;
      }
    }
  });
}

async function resolveIncomingBinaryFile(item) {
  // 1. Cek cache IndexedDB perangkat terlebih dahulu agar instan
  const cached = await store.getIncomingFileCache(item.id);
  if (cached?.fileBase64) {
    return {
      fileName: cached.fileName || item.fileName,
      mimeType: cached.mimeType || item.mimeType,
      fileKind: cached.fileKind || item.fileKind,
      fileBase64: cached.fileBase64
    };
  }

  // 2. Tarik dari server cloud blob
  if (getStoredAuthToken()) {
    const res = await fetchIncomingFileCloud(item.id);
    if (res.ok && res.fileBase64) {
      store.saveIncomingFileCache(item.id, res);
      return res;
    }
  }
  return null;
}

function bindIncomingRowActions(container) {
  const prevModal = container.querySelector('#modal-incoming-preview');
  const prevTitle = container.querySelector('#modal-in-prev-title');
  const prevSub = container.querySelector('#modal-in-prev-sub');
  const prevBody = container.querySelector('#modal-in-prev-body');
  const prevDriveStatus = container.querySelector('#modal-in-prev-drive-status');
  const btnPrevDriveLink = container.querySelector('#btn-in-prev-open-drive');
  const btnPrevDownload = container.querySelector('#btn-in-prev-download');

  let activeObjectUrl = '';
  const cleanupPreviewObjUrl = () => {
    if (activeObjectUrl) {
      try { URL.revokeObjectURL(activeObjectUrl); } catch {}
      activeObjectUrl = '';
    }
  };
  const closePrevModal = () => {
    cleanupPreviewObjUrl();
    if (prevBody) prevBody.innerHTML = '';
    prevModal?.classList.add('hidden');
  };
  container.querySelector('#btn-close-in-prev')?.addEventListener('click', closePrevModal);

  const triggerDownload = (fileRec) => {
    if (!fileRec?.fileBase64) return;
    const bytes = base64ToUint8Array(fileRec.fileBase64);
    const blob = new Blob([bytes], { type: fileRec.mimeType || 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileRec.fileName || 'surat_masuk';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  };

  container.querySelectorAll('.btn-in-preview').forEach(btn => {
    btn.addEventListener('click', async () => {
      const id = btn.dataset.id;
      const item = store.getIncomingLetters().find(i => i.id === id);
      if (!item || !prevModal) return;

      cleanupPreviewObjUrl();
      prevTitle.textContent = `${item.kodeSurat} — ${item.namaSurat}`;
      prevSub.textContent = `${item.pengirim ? `Pengirim: ${item.pengirim} · ` : ''}Tanggal: ${formatReadableDate(item.tanggalSurat || item.receivedAt)} · ${item.fileName}`;
      prevBody.innerHTML = `<div style="color:#475569;font-size:13px;display:inline-flex;align-items:center;gap:8px;">${uiIcon('loader', 15, 'spin-icon')}<span>Memuat berkas pratinjau...</span></div>`;

      if (item.driveSynced && item.driveUrl) {
        prevDriveStatus.innerHTML = `<span class="status-chip status-chip-green">${uiIcon('cloudCheck', 12)}<span>Tersimpan di Google Drive</span></span>`;
        btnPrevDriveLink.href = item.driveUrl;
        btnPrevDriveLink.classList.remove('hidden');
      } else {
        prevDriveStatus.innerHTML = `<span class="status-chip status-chip-amber">${uiIcon('clock', 12)}<span>Tersimpan di Cloud Lokal (Belum masuk Drive)</span></span>`;
        btnPrevDriveLink.classList.add('hidden');
      }

      prevModal.classList.remove('hidden');

      const fileRec = await resolveIncomingBinaryFile(item);
      if (fileRec?.fileBase64) {
        const bytes = base64ToUint8Array(fileRec.fileBase64);
        const blob = new Blob([bytes], { type: fileRec.mimeType || item.mimeType });
        activeObjectUrl = URL.createObjectURL(blob);
        btnPrevDownload.disabled = false;
        btnPrevDownload.onclick = () => triggerDownload(fileRec);

        if ((item.fileKind === 'pdf') || String(fileRec.mimeType).includes('pdf')) {
          prevBody.innerHTML = `<iframe src="${activeObjectUrl}" title="Pratinjau PDF Surat Masuk" style="width:100%; height:64vh; border:1px solid #cbd5e1; border-radius:6px; background:#ffffff;"></iframe>`;
        } else {
          prevBody.innerHTML = `<img src="${activeObjectUrl}" alt="${escapeHtml(item.namaSurat)}" style="max-width:100%; max-height:66vh; object-fit:contain; border-radius:6px; box-shadow:0 4px 16px rgba(15,23,42,0.14); background:#ffffff;" />`;
        }
        return;
      }

      btnPrevDownload.disabled = !item.driveDownloadUrl;
      btnPrevDownload.onclick = () => {
        if (item.driveDownloadUrl) window.open(item.driveDownloadUrl, '_blank', 'noopener');
      };

      if (item.driveFileId) {
        prevBody.innerHTML = `
          <iframe src="https://drive.google.com/file/d/${encodeURIComponent(item.driveFileId)}/preview"
            title="Pratinjau Google Drive" style="width:100%; height:64vh; border:1px solid #cbd5e1; border-radius:6px; background:#ffffff;"></iframe>
        `;
      } else {
        prevBody.innerHTML = `<div style="color:#64748b;font-size:13px;">Pratinjau biner tidak tersedia. Buka langsung melalui Google Drive jika sudah tersinkron.</div>`;
      }
    });
  });

  container.querySelectorAll('.btn-in-download').forEach(btn => {
    btn.addEventListener('click', async () => {
      const item = store.getIncomingLetters().find(i => i.id === btn.dataset.id);
      if (!item) return;
      const fileRec = await resolveIncomingBinaryFile(item);
      if (fileRec?.fileBase64) {
        triggerDownload(fileRec);
      } else if (item.driveDownloadUrl || item.driveUrl) {
        window.open(item.driveDownloadUrl || item.driveUrl, '_blank', 'noopener');
      } else {
        showToast('Berkas unduhan belum tersedia.', 'warning');
      }
    });
  });

  container.querySelectorAll('.btn-in-sync-drive').forEach(btn => {
    btn.addEventListener('click', async () => {
      const id = btn.dataset.id;
      const item = store.getIncomingLetters().find(i => i.id === id);
      if (!item) return;
      if (!getStoredAuthToken()) {
        showToast('Silakan masuk akun terlebih dahulu untuk mengirim ke Google Drive.', 'warning');
        return;
      }
      btn.disabled = true;
      btn.innerHTML = `${uiIcon('loader', 11, 'spin-icon')}<span>Mengirim...</span>`;
      try {
        const cached = await store.getIncomingFileCache(id);
        const res = await syncIncomingDriveCloud(id, cached?.fileBase64 || '');
        if (!res.ok) {
          const errMsg = res.error || 'Gagal mengirim ke Google Drive.';
          if (res.item || res.data?.item) store.upsertIncomingLetterLocal(res.item || res.data.item);
          const isGasMissing = Boolean(res.needGasUpdate || res.data?.needGasUpdate || /aksi tidak dikenali|archive_incoming/i.test(errMsg));
          if (isGasMissing && typeof container._openGasGuideModal === 'function') {
            showToast('Code.gs Apps Script perlu ditambah handler archive_incoming. Ikuti panduan di layar.', 'warning');
            container._openGasGuideModal(true);
          } else {
            showToast(errMsg, 'error');
          }
          return;
        }
        if (res.syncRev) setLastKnownSyncRev(res.syncRev);
        if (res.driveStatus) store.setAuthSessionData({ driveStatus: res.driveStatus });
        if (Array.isArray(res.incomingLetters)) {
          store.setState({ incomingLetters: res.incomingLetters }, false);
          store.saveIncomingLetters(res.incomingLetters);
        } else if (res.item) {
          store.upsertIncomingLetterLocal(res.item);
        }
        renderIncomingLettersView(container);
        showToast(`"${item.kodeSurat}" berhasil disimpan ke Google Drive!`, 'success');
      } finally {
        if (document.body.contains(btn)) {
          btn.disabled = false;
          btn.innerHTML = `${uiIcon('refresh', 11)}<span>Kirim ke Drive</span>`;
        }
      }
    });
  });

  // Hapus Surat Masuk
  const delModal = container.querySelector('#modal-confirm-delete-in');
  const delBox = container.querySelector('#del-in-detail-box');
  const btnCancelDel = container.querySelector('#btn-cancel-del-in');
  const btnCloseDel = container.querySelector('#btn-close-del-in');
  const btnConfirmDel = container.querySelector('#btn-confirm-del-in');
  let pendingDeleteId = '';

  const closeDel = () => {
    pendingDeleteId = '';
    delModal?.classList.add('hidden');
  };
  btnCancelDel?.addEventListener('click', closeDel);
  btnCloseDel?.addEventListener('click', closeDel);

  container.querySelectorAll('.btn-in-delete').forEach(btn => {
    btn.addEventListener('click', () => {
      const item = store.getIncomingLetters().find(i => i.id === btn.dataset.id);
      if (!item || !delModal) return;
      pendingDeleteId = item.id;
      delBox.innerHTML = `
        <div><strong>Kode:</strong> ${escapeHtml(item.kodeSurat)}</div>
        <div><strong>Nama:</strong> ${escapeHtml(item.namaSurat)}</div>
        ${item.pengirim ? `<div><strong>Pengirim:</strong> ${escapeHtml(item.pengirim)}</div>` : ''}
      `;
      delModal.classList.remove('hidden');
    });
  });

  btnConfirmDel?.addEventListener('click', async () => {
    if (!pendingDeleteId) return;
    const idToRemove = pendingDeleteId;
    btnConfirmDel.disabled = true;
    try {
      if (getStoredAuthToken()) {
        const res = await deleteIncomingCloud(idToRemove);
        if (!res.ok) {
          showToast(res.error || 'Gagal menghapus Surat Masuk.', 'error');
          return;
        }
        if (res.syncRev) setLastKnownSyncRev(res.syncRev);
        store.removeIncomingLetterLocal(idToRemove);
        if (Array.isArray(res.incomingLetters)) {
          store.setState({ incomingLetters: res.incomingLetters }, false);
          store.saveIncomingLetters(res.incomingLetters);
        }
      } else {
        store.removeIncomingLetterLocal(idToRemove);
      }
      closeDel();
      renderIncomingLettersView(container);
      showToast('Arsip Surat Masuk berhasil dihapus.', 'info');
    } finally {
      if (document.body.contains(btnConfirmDel)) btnConfirmDel.disabled = false;
    }
  });
}
