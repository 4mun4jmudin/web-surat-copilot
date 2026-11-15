/**
 * outgoingLettersView.js - Dashboard & Buku Catatan Surat Keluar Resmi (Arsip Tetap / Terkunci)
 * Berbeda dengan Gudang Surat (yang berisi template/draf yang dapat diedit & digunakan kembali),
 * Surat Keluar bersifat permanen/terkunci (tidak dapat diedit), memiliki nomor urut tersusun,
 * serta menampilkan daftar surat berdasarkan atas nama penerima atau judul surat.
 */

import { store } from '../state.js';
import { renderDocx } from '../docx/render.js';
import { downloadFile } from '../utils/download.js';
import { showToast } from './notification.js';
import { LETTER_TYPE_CODES } from '../letters/numbering.js';
import { damerauLevenshtein } from './vaultView.js';

let outgoingSearchQuery = '';
let outgoingTypeFilter = 'all';
let outgoingSortMode = 'seq_desc';
let showDeleteActions = false;
let activeAltDHandler = null;

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str || '';
  return div.innerHTML;
}

/**
 * Pencarian typo-tolerant untuk Surat Keluar (berdasarkan atas nama seseorang, judul surat, nomor surat, perihal)
 */
function scoreOutgoingMatch(item, query) {
  const q = String(query || '').trim().toLowerCase();
  if (!q) return 1;

  const recipient = (item.recipientName || '').toLowerCase();
  const title = (item.title || '').toLowerCase();
  const nomor = (item.nomorSurat || '').toLowerCase();
  const perihal = (item.perihal || '').toLowerCase();
  const combined = `${recipient} ${title} ${nomor} ${perihal}`;

  if (combined.includes(q)) {
    if (recipient.includes(q) || title.includes(q)) return 100;
    return 80;
  }

  const tokens = q.split(/\s+/).filter(Boolean);
  const words = combined.split(/[\s\/.,\-_()]+/).filter(Boolean);

  let matchedTokens = 0;
  let totalScore = 0;

  for (const tok of tokens) {
    let bestTokScore = 0;
    for (const w of words) {
      if (w.includes(tok)) {
        bestTokScore = Math.max(bestTokScore, 1.0);
        break;
      }
      const maxLen = Math.max(tok.length, w.length);
      if (maxLen >= 3) {
        const dist = damerauLevenshtein(tok, w);
        const tol = tok.length <= 4 ? 1 : tok.length <= 8 ? 2 : 3;
        if (dist <= tol) {
          bestTokScore = Math.max(bestTokScore, 1 - dist / maxLen);
        }
      }
    }
    if (bestTokScore >= 0.45) {
      matchedTokens++;
      totalScore += bestTokScore;
    }
  }

  if (matchedTokens === 0) return 0;
  return (matchedTokens / tokens.length) * 50 + totalScore * 10;
}

function getFilteredAndSortedOutgoing() {
  const list = store.getOutgoingLetters();
  const q = outgoingSearchQuery.trim();

  let filtered = list.filter(item => {
    if (outgoingTypeFilter !== 'all' && (item.typeCode || '').toUpperCase() !== outgoingTypeFilter) {
      return false;
    }
    if (!q) return true;
    return scoreOutgoingMatch(item, q) > 0;
  });

  if (q) {
    filtered = filtered
      .map(item => ({ item, score: scoreOutgoingMatch(item, q) }))
      .sort((a, b) => b.score - a.score)
      .map(x => x.item);
    return filtered;
  }

  return [...filtered].sort((a, b) => {
    if (outgoingSortMode === 'seq_asc') {
      const yDiff = Number(a.year || 0) - Number(b.year || 0);
      return yDiff !== 0 ? yDiff : (Number(a.sequenceNumber) || 0) - (Number(b.sequenceNumber) || 0);
    }
    if (outgoingSortMode === 'recipient_asc') {
      return (a.recipientName || '').localeCompare(b.recipientName || '', 'id');
    }
    if (outgoingSortMode === 'title_asc') {
      return (a.title || '').localeCompare(b.title || '', 'id');
    }
    // Default: seq_desc (terbaru / nomor urut tertinggi)
    const yDiff = Number(b.year || 0) - Number(a.year || 0);
    return yDiff !== 0 ? yDiff : (Number(b.sequenceNumber) || 0) - (Number(a.sequenceNumber) || 0);
  });
}

export function renderOutgoingLettersView(container) {
  if (!container) return;
  const allOutgoing = store.getOutgoingLetters();
  const nextSeqInfo = store.getNextFormattedOutgoingNomor();
  const uniqueRecipients = new Set(
    allOutgoing
      .map(x => (x.recipientName || '').trim().toLowerCase())
      .filter(x => x && x !== '-')
  ).size;

  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();
  const thisMonthCount = allOutgoing.filter(item => {
    const d = item.recordedAt ? new Date(item.recordedAt) : null;
    return d && !isNaN(d.getTime()) && d.getMonth() === currentMonth && d.getFullYear() === currentYear;
  }).length;

  const lastSeq = allOutgoing.length > 0
    ? Math.max(...allOutgoing.map(i => Number(i.sequenceNumber) || 0))
    : 0;

  const typeCodesInUse = Array.from(new Set(allOutgoing.map(i => (i.typeCode || '').toUpperCase()).filter(Boolean)));

  const OUT_SVG_TRAY = `<svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#156c3b" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>`;
  const OUT_SVG_DOC = (size = 18) => `<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>`;
  const OUT_SVG_HASH = (size = 18) => `<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="4" y1="9" x2="20" y2="9"/><line x1="4" y1="15" x2="20" y2="15"/><line x1="10" y1="3" x2="8" y2="21"/><line x1="16" y1="3" x2="14" y2="21"/></svg>`;
  const OUT_SVG_CAL = (size = 18) => `<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>`;
  const OUT_SVG_USERS = (size = 18) => `<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>`;
  const OUT_SVG_SEARCH = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>`;
  const OUT_SVG_LAYERS = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#156c3b" stroke-width="2"><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 12 12 17 22 12"/><polyline points="2 17 12 22 22 17"/></svg>`;

  container.innerHTML = `
    <div class="view-atmosphere">
      <div class="view-shell outgoing-view-container">

        <!-- HEADER DASHBOARD SURAT KELUAR -->
        <div class="view-page-header" style="margin-bottom: 18px;">
          <div class="view-header-lead">
            <span style="margin-top: 2px;">${OUT_SVG_TRAY}</span>
            <span class="view-header-divider"></span>
            <div>
              <h2 class="view-title">
                <span>Buku Catatan &amp; Dashboard Surat Keluar</span>
                <span class="status-chip status-chip-amber">🔒 Arsip Tetap (Tidak Bisa Diedit)</span>
              </h2>
              <p class="view-subtitle" style="max-width: 730px;">
                Daftar resmi surat yang telah diterbitkan dengan <strong>nomor surat tersusun berurutan</strong>. Catatan surat keluar terkunci permanen agar integritas nomor &amp; nama penerima terjaga. (Untuk menyimpan template/draf yang masih bisa diedit ulang, gunakan tab <strong>Gudang Surat</strong>).
              </p>
            </div>
          </div>

          <div style="display: flex; gap: 8px; align-items: center; flex-shrink: 0;">
            <button id="btn-go-canvas-new-out" class="btn btn-blue-cta" style="font-size: 13px; padding: 10px 18px; border-radius: 8px;">
              + Buat &amp; Catat Surat Keluar Baru
            </button>
          </div>
        </div>

        <!-- KARTU STATISTIK DASHBOARD -->
        <div class="kpi-grid-4">
          <div class="kpi-card accent-left-green">
            <div class="kpi-icon-box kpi-icon-green">${OUT_SVG_DOC(19)}</div>
            <div style="flex: 1; min-width: 0;">
              <div style="font-size: 12px; color: #475569; font-weight: 600;">Total Surat Keluar Tercatat</div>
              <div style="font-size: 26px; font-weight: 800; color: #0f172a; line-height: 1.2; margin-top: 3px;">${allOutgoing.length}</div>
              <div style="font-size: 11px; color: #64748b; margin-top: 3px;">Arsip resmi terkunci permanen</div>
            </div>
            <div class="kpi-watermark" style="color: #16a34a;">${OUT_SVG_DOC(48)}</div>
          </div>

          <div class="kpi-card accent-left-blue">
            <div class="kpi-icon-box kpi-icon-blue">${OUT_SVG_HASH(19)}</div>
            <div style="flex: 1; min-width: 0;">
              <div style="font-size: 12px; color: #475569; font-weight: 600;">Nomor Urut Tersusun</div>
              <div style="display: flex; align-items: center; gap: 7px; flex-wrap: wrap; margin-top: 3px;">
                <span style="font-size: 24px; font-weight: 800; color: #0f172a; line-height: 1.2;">${lastSeq > 0 ? `#${String(lastSeq).padStart(3, '0')}` : '#000'}</span>
                <span class="status-chip status-chip-green" style="font-size: 10.5px; padding: 2px 8px;">Berikutnya #${String(nextSeqInfo.sequenceNumber).padStart(3, '0')} &rsaquo;</span>
              </div>
              <div style="font-size: 11px; color: #64748b; margin-top: 3px;">Otomatis berurutan per tahun (${nextSeqInfo.year})</div>
            </div>
          </div>

          <div class="kpi-card accent-left-amber">
            <div class="kpi-icon-box kpi-icon-amber">${OUT_SVG_CAL(19)}</div>
            <div style="flex: 1; min-width: 0;">
              <div style="font-size: 12px; color: #475569; font-weight: 600;">Surat Keluar Bulan Ini</div>
              <div style="font-size: 26px; font-weight: 800; color: #b45309; line-height: 1.2; margin-top: 3px;">${thisMonthCount}</div>
              <div style="font-size: 11px; color: #64748b; margin-top: 3px;">Periode bulan berjalan</div>
            </div>
            <div class="kpi-watermark" style="color: #f59e0b;">${OUT_SVG_CAL(48)}</div>
          </div>

          <div class="kpi-card accent-left-purple">
            <div class="kpi-icon-box kpi-icon-purple">${OUT_SVG_USERS(19)}</div>
            <div style="flex: 1; min-width: 0;">
              <div style="font-size: 12px; color: #475569; font-weight: 600;">Atas Nama / Penerima Unik</div>
              <div style="font-size: 26px; font-weight: 800; color: #6d28d9; line-height: 1.2; margin-top: 3px;">${uniqueRecipients}</div>
              <div style="font-size: 11px; color: #64748b; margin-top: 3px;">Siswa, Guru, atau Instansi tujuan</div>
            </div>
            <div class="kpi-watermark" style="color: #8b5cf6;">${OUT_SVG_USERS(48)}</div>
          </div>
        </div>

        <!-- FILTER & PENCARIAN ATAS NAMA / JUDUL SURAT -->
        <div class="filter-bar-card">
          <div class="icon-input-wrap" style="flex: 1; min-width: 260px;">
            <span class="icon-input-leading">${OUT_SVG_SEARCH}</span>
            <input type="text" id="input-outgoing-search" class="icon-input-control"
              value="${escapeHtml(outgoingSearchQuery)}"
              placeholder="Cari atas nama seseorang, judul surat, nomor surat, atau perihal (mendukung typo)...">
          </div>
          <div class="icon-input-wrap" style="width: 205px;">
            <span class="icon-input-leading">${OUT_SVG_LAYERS}</span>
            <select id="select-outgoing-type" class="icon-input-control" style="cursor: pointer;">
              <option value="all" ${outgoingTypeFilter === 'all' ? 'selected' : ''}>Semua Kode Surat</option>
              ${typeCodesInUse.map(code => {
                const label = LETTER_TYPE_CODES[code]?.label || code;
                return `<option value="${escapeHtml(code)}" ${outgoingTypeFilter === code ? 'selected' : ''}>${escapeHtml(label)}</option>`;
              }).join('')}
            </select>
          </div>
          <div class="icon-input-wrap" style="width: 225px;">
            <span class="icon-input-leading">⇅</span>
            <select id="select-outgoing-sort" class="icon-input-control" style="cursor: pointer;">
              <option value="seq_desc" ${outgoingSortMode === 'seq_desc' ? 'selected' : ''}>Urut: No. Terbaru (Terbesar)</option>
              <option value="seq_asc" ${outgoingSortMode === 'seq_asc' ? 'selected' : ''}>Urut: No. Awal (001 → ...)</option>
              <option value="recipient_asc" ${outgoingSortMode === 'recipient_asc' ? 'selected' : ''}>Urut: Atas Nama (A - Z)</option>
              <option value="title_asc" ${outgoingSortMode === 'title_asc' ? 'selected' : ''}>Urut: Judul Surat (A - Z)</option>
            </select>
          </div>
        </div>

        <!-- BANNER MODE HAPUS (HANYA MUNCUL KETIKA SHORTCUT ALT+D DIAKTIFKAN) -->
        <div id="outgoing-delete-mode-banner" class="${showDeleteActions ? '' : 'hidden'}" style="background: #fef2f2; border: 1px solid #fecaca; color: #991b1b; border-radius: 10px; padding: 10px 16px; margin-bottom: 14px; display: flex; justify-content: space-between; align-items: center; gap: 10px; font-size: 12.5px;">
          <div>
            <strong>⚠️ Mode Hapus Surat Keluar Aktif:</strong> Tombol eksekusi <b>Hapus</b> kini ditampilkan pada setiap baris daftar Surat Keluar.
          </div>
          <button type="button" id="btn-exit-delete-mode" class="btn btn-sm btn-secondary" style="border-color: #fca5a5; color: #991b1b;">Sembunyikan Tombol Hapus (Alt+D)</button>
        </div>

        <!-- DAFTAR SURAT KELUAR (TABEL DASHBOARD / EMPTY STATE) -->
        <div id="outgoing-list-area"></div>

      </div>
    </div>

    <!-- MODAL KONFIRMASI HAPUS SURAT KELUAR -->
    <div id="modal-confirm-delete-out" class="modal-overlay hidden">
      <div class="modal-card" style="max-width: 440px; width: 92%;">
        <div class="modal-header" style="display: flex; justify-content: space-between; align-items: center; padding-bottom: 10px; border-bottom: 1px solid #e2e8f0;">
          <h3 style="margin: 0; font-size: 16px; color: #b91c1c;">🗑️ Konfirmasi Hapus Surat Keluar</h3>
          <button type="button" id="btn-close-del-out-modal" class="btn btn-sm btn-secondary">&times;</button>
        </div>
        <div style="padding: 14px 0; font-size: 13px; color: #334155;">
          <p style="margin: 0 0 10px 0;">Anda yakin ingin menghapus catatan Surat Keluar berikut dari daftar arsip?</p>
          <div id="del-out-detail-box" style="background: #fef2f2; border: 1px solid #fecaca; border-radius: 6px; padding: 10px 12px; font-size: 12.5px; color: #7f1d1d;"></div>
        </div>
        <div style="display: flex; justify-content: flex-end; gap: 8px; border-top: 1px solid #e2e8f0; padding-top: 12px;">
          <button type="button" id="btn-cancel-del-out" class="btn btn-secondary">Batal</button>
          <button type="button" id="btn-confirm-del-out" class="btn btn-danger" style="font-weight: 700;">🗑️ Eksekusi Hapus</button>
        </div>
      </div>
    </div>

    <!-- MODAL PRATINJAU READ-ONLY SURAT KELUAR -->
    <div id="modal-outgoing-preview" class="modal-overlay hidden">
      <div class="modal-card" style="max-width: 760px; width: 92%; max-height: 88vh; display: flex; flex-direction: column;">
        <div class="modal-header" style="display: flex; justify-content: space-between; align-items: center; padding-bottom: 10px; border-bottom: 1px solid #e2e8f0;">
          <div>
            <h3 id="modal-out-title" style="margin: 0; font-size: 16px; color: #0f172a;">Pratinjau Surat Keluar</h3>
            <div id="modal-out-subtitle" style="font-size: 12px; color: #64748b; margin-top: 2px;"></div>
          </div>
          <button type="button" id="btn-close-out-modal" class="btn btn-sm btn-secondary">✕ Tutup</button>
        </div>
        <div id="modal-out-body" style="flex: 1; overflow-y: auto; padding: 18px; background: #f1f5f9; border-radius: 6px; margin: 12px 0;"></div>
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px; padding-top: 8px; border-top: 1px solid #e2e8f0;">
          <span style="font-size: 11.5px; color: #b91c1c; font-weight: 600;">
            🔒 Arsip Surat Keluar ini terkunci dan tidak dapat diedit.
          </span>
          <div style="display: flex; gap: 8px;">
            <button type="button" id="btn-modal-out-clone" class="btn btn-sm btn-secondary">📋 Salin jadi Draf Baru</button>
            <button type="button" id="btn-modal-out-download" class="btn btn-sm btn-success">⬇️ Unduh .docx</button>
          </div>
        </div>
      </div>
    </div>
  `;

  updateOutgoingListInPlace(container);
  bindOutgoingEvents(container);
}

function updateOutgoingListInPlace(container) {
  const listArea = container.querySelector('#outgoing-list-area');
  if (!listArea) return;

  const items = getFilteredAndSortedOutgoing();

  if (items.length === 0) {
    listArea.innerHTML = `
      <div class="empty-illustration-card">
        <div style="display: inline-block; margin-bottom: 12px;">
          <svg width="160" height="96" viewBox="0 0 160 96" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
            <ellipse cx="80" cy="64" rx="68" ry="22" fill="#e6f4ec"/>
            <rect x="44" y="16" width="60" height="62" rx="8" fill="#ffffff" stroke="#1e293b" stroke-width="2.2"/>
            <line x1="56" y1="30" x2="92" y2="30" stroke="#94a3b8" stroke-width="3" stroke-linecap="round"/>
            <line x1="56" y1="40" x2="86" y2="40" stroke="#cbd5e1" stroke-width="3" stroke-linecap="round"/>
            <line x1="56" y1="50" x2="78" y2="50" stroke="#cbd5e1" stroke-width="3" stroke-linecap="round"/>
            <line x1="56" y1="60" x2="70" y2="60" stroke="#cbd5e1" stroke-width="3" stroke-linecap="round"/>
            <path d="M84 58L118 42L106 76L98 63L84 58Z" fill="#156c3b" stroke="#0f172a" stroke-width="2" stroke-linejoin="round"/>
            <line x1="98" y1="63" x2="118" y2="42" stroke="#dcfce7" stroke-width="1.8" stroke-linecap="round"/>
            <line x1="120" y1="26" x2="124" y2="18" stroke="#156c3b" stroke-width="2" stroke-linecap="round"/>
            <line x1="127" y1="32" x2="133" y2="27" stroke="#156c3b" stroke-width="2" stroke-linecap="round"/>
            <line x1="128" y1="40" x2="136" y2="39" stroke="#156c3b" stroke-width="2" stroke-linecap="round"/>
          </svg>
        </div>
        <div style="font-size: 19px; font-weight: 800; color: #0f172a; margin-bottom: 8px;">
          ${outgoingSearchQuery ? 'Tidak Ada Surat Keluar yang Cocok' : 'Belum Ada Catatan Surat Keluar'}
        </div>
        <div style="font-size: 13px; color: #475569; max-width: 540px; margin: 0 auto; line-height: 1.7;">
          ${outgoingSearchQuery
            ? 'Coba kata kunci nama penerima, judul surat, atau nomor surat lainnya.'
            : 'Buat surat di Lembar Kerja, lalu klik tombol <span style="display:inline-block; background:#dcfce7; color:#15803d; font-weight:700; padding:1px 9px; border-radius:6px; margin:0 2px;">“Catat Surat Keluar”</span> di ribbon untuk mengunci arsip &amp; menyusun nomor urut surat secara resmi.'}
        </div>
      </div>
    `;
    return;
  }

  listArea.innerHTML = `
    <div class="ui-card" style="padding: 0; overflow: hidden;">
      <div style="overflow-x: auto;">
        <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
          <thead>
            <tr style="background: #f8fafc; border-bottom: 1px solid #e2e8f0; text-align: left; color: #334155; font-size: 12px;">
              <th style="padding: 12px 14px; width: 75px;">Urut</th>
              <th style="padding: 12px 14px; width: 235px;">Nomor Surat Tersusun</th>
              <th style="padding: 12px 14px;">Judul Surat &amp; Perihal</th>
              <th style="padding: 12px 14px; width: 220px;">Atas Nama / Penerima</th>
              <th style="padding: 12px 14px; width: 135px;">Tanggal Surat</th>
              <th style="padding: 12px 14px; width: 225px; text-align: right;">Aksi (Terkunci 🔒)</th>
            </tr>
          </thead>
          <tbody>
            ${items.map(item => `
              <tr style="border-bottom: 1px solid #f1f5f9; transition: background 0.12s;" class="outgoing-row">
                <td style="padding: 12px 14px; font-weight: 800; color: #156c3b;">
                  #${String(item.sequenceNumber || 0).padStart(3, '0')}
                </td>
                <td style="padding: 12px 14px;">
                  <div style="font-family: Consolas, monospace; font-size: 12px; font-weight: 700; color: #0f172a;">
                    ${escapeHtml(item.nomorSurat)}
                  </div>
                  <span class="status-chip status-chip-blue" style="margin-top: 4px; font-size: 10px; padding: 1px 7px;">
                    ${escapeHtml(item.typeCode || 'DINAS')}
                  </span>
                </td>
                <td style="padding: 12px 14px;">
                  <div style="font-weight: 700; color: #0f172a;">${escapeHtml(item.title)}</div>
                  <div style="font-size: 11.5px; color: #64748b; margin-top: 2px;">${escapeHtml(item.perihal || '-')}</div>
                </td>
                <td style="padding: 12px 14px;">
                  <span class="status-chip status-chip-purple" style="font-size: 11.5px; font-weight: 600;">
                    👤 ${escapeHtml(item.recipientName || '-')}
                  </span>
                </td>
                <td style="padding: 12px 14px; color: #475569; font-size: 12px;">
                  ${escapeHtml(item.tanggalSurat || '-')}
                </td>
                <td style="padding: 12px 14px; text-align: right; white-space: nowrap;">
                  <button type="button" class="btn btn-sm btn-secondary btn-out-view" data-id="${escapeHtml(item.id)}" title="Lihat pratinjau surat (Read-Only)">👁️ Lihat</button>
                  <button type="button" class="btn btn-sm btn-success btn-out-download" data-id="${escapeHtml(item.id)}" title="Unduh berkas .docx resmi">⬇️</button>
                  <button type="button" class="btn btn-sm btn-secondary btn-out-clone" data-id="${escapeHtml(item.id)}" title="Salin isi surat sebagai draf baru di Lembar Kerja (tanpa mengubah arsip ini)">📋 Salin</button>
                  ${showDeleteActions ? `<button type="button" class="btn btn-sm btn-danger btn-out-delete" data-id="${escapeHtml(item.id)}" title="Hapus data Surat Keluar ini">🗑️ Hapus</button>` : ''}
                </td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>
  `;

  bindOutgoingRowActions(container);
}

function bindOutgoingEvents(container) {
  const searchInput = container.querySelector('#input-outgoing-search');
  const typeSelect = container.querySelector('#select-outgoing-type');
  const sortSelect = container.querySelector('#select-outgoing-sort');
  const btnGoCanvas = container.querySelector('#btn-go-canvas-new-out');

  if (searchInput) {
    searchInput.addEventListener('input', () => {
      outgoingSearchQuery = searchInput.value;
      updateOutgoingListInPlace(container);
    });
  }

  if (typeSelect) {
    typeSelect.addEventListener('change', () => {
      outgoingTypeFilter = typeSelect.value;
      updateOutgoingListInPlace(container);
    });
  }

  if (sortSelect) {
    sortSelect.addEventListener('change', () => {
      outgoingSortMode = sortSelect.value;
      updateOutgoingListInPlace(container);
    });
  }

  if (btnGoCanvas) {
    btnGoCanvas.addEventListener('click', () => {
      const canvasTab = document.querySelector('.nav-tab[data-tab="canvas"]');
      if (canvasTab) canvasTab.click();
    });
  }

  const modal = container.querySelector('#modal-outgoing-preview');
  const btnCloseModal = container.querySelector('#btn-close-out-modal');
  if (btnCloseModal && modal) {
    btnCloseModal.addEventListener('click', () => modal.classList.add('hidden'));
  }

  const delBanner = container.querySelector('#outgoing-delete-mode-banner');
  const btnExitDelMode = container.querySelector('#btn-exit-delete-mode');

  const toggleDeleteMode = (forceState = null) => {
    showDeleteActions = forceState !== null ? Boolean(forceState) : !showDeleteActions;
    if (delBanner) {
      delBanner.classList.toggle('hidden', !showDeleteActions);
    }
    updateOutgoingListInPlace(container);
    showToast(
      showDeleteActions
        ? '🔓 Mode Hapus Surat Keluar aktif (Alt+D): Tombol Hapus kini tampil di daftar.'
        : '🔒 Mode Hapus disembunyikan kembali.',
      'info'
    );
  };

  if (btnExitDelMode) {
    btnExitDelMode.addEventListener('click', () => toggleDeleteMode(false));
  }

  // Shortcut Keyboard Alt + D untuk memunculkan/menyembunyikan tombol eksekusi Delete
  if (activeAltDHandler) {
    document.removeEventListener('keydown', activeAltDHandler);
  }
  activeAltDHandler = (e) => {
    // Pastikan tab Surat Keluar memang sedang aktif tampil di layar
    if (!document.body.contains(container) || !container.querySelector('#outgoing-list-area')) {
      document.removeEventListener('keydown', activeAltDHandler);
      activeAltDHandler = null;
      showDeleteActions = false;
      return;
    }
    if (e.altKey && !e.ctrlKey && !e.metaKey && (e.key === 'd' || e.key === 'D' || e.code === 'KeyD')) {
      e.preventDefault();
      toggleDeleteMode();
    }
  };
  document.addEventListener('keydown', activeAltDHandler);
}

function bindOutgoingRowActions(container) {
  const modal = container.querySelector('#modal-outgoing-preview');
  const modalTitle = container.querySelector('#modal-out-title');
  const modalSubtitle = container.querySelector('#modal-out-subtitle');
  const modalBody = container.querySelector('#modal-out-body');
  const btnModalDownload = container.querySelector('#btn-modal-out-download');
  const btnModalClone = container.querySelector('#btn-modal-out-clone');

  const delModal = container.querySelector('#modal-confirm-delete-out');
  const delDetailBox = container.querySelector('#del-out-detail-box');
  const btnCloseDelModal = container.querySelector('#btn-close-del-out-modal');
  const btnCancelDel = container.querySelector('#btn-cancel-del-out');
  const btnConfirmDel = container.querySelector('#btn-confirm-del-out');

  const closeDelModal = () => {
    if (delModal) delModal.classList.add('hidden');
  };
  if (btnCloseDelModal) btnCloseDelModal.onclick = closeDelModal;
  if (btnCancelDel) btnCancelDel.onclick = closeDelModal;

  container.querySelectorAll('.btn-out-delete').forEach(btn => {
    btn.addEventListener('click', () => {
      const item = store.getOutgoingLetters().find(x => x.id === btn.dataset.id);
      if (!item || !delModal) return;

      if (delDetailBox) {
        delDetailBox.innerHTML = `
          <div><strong>No. Urut:</strong> #${String(item.sequenceNumber || 0).padStart(3, '0')}</div>
          <div><strong>Nomor Surat:</strong> ${escapeHtml(item.nomorSurat)}</div>
          <div><strong>Judul:</strong> ${escapeHtml(item.title)}</div>
          <div><strong>Atas Nama:</strong> ${escapeHtml(item.recipientName)}</div>
        `;
      }
      delModal.classList.remove('hidden');

      if (btnConfirmDel) {
        btnConfirmDel.onclick = () => {
          store.removeOutgoingLetter(item.id);
          closeDelModal();
          showToast(`🗑️ Data Surat Keluar #${String(item.sequenceNumber || 0).padStart(3, '0')} (${item.nomorSurat}) dihapus.`, 'success');
          renderOutgoingLettersView(container);
        };
      }
    });
  });

  container.querySelectorAll('.btn-out-view').forEach(btn => {
    btn.addEventListener('click', () => {
      const item = store.getOutgoingLetters().find(x => x.id === btn.dataset.id);
      if (!item || !modal) return;

      modalTitle.textContent = `${item.title} (#${String(item.sequenceNumber || 0).padStart(3, '0')})`;
      modalSubtitle.textContent = `No: ${item.nomorSurat} • Atas Nama: ${item.recipientName} • Tanggal: ${item.tanggalSurat}`;
      modalBody.innerHTML = renderReadOnlyDocumentHtml(item);
      modal.dataset.activeId = item.id;
      modal.classList.remove('hidden');

      if (btnModalDownload) {
        btnModalDownload.onclick = () => downloadOutgoingItem(item);
      }
      if (btnModalClone) {
        btnModalClone.onclick = () => {
          modal.classList.add('hidden');
          cloneOutgoingItem(item);
        };
      }
    });
  });

  container.querySelectorAll('.btn-out-download').forEach(btn => {
    btn.addEventListener('click', () => {
      const item = store.getOutgoingLetters().find(x => x.id === btn.dataset.id);
      if (item) downloadOutgoingItem(item);
    });
  });

  container.querySelectorAll('.btn-out-clone').forEach(btn => {
    btn.addEventListener('click', () => {
      const item = store.getOutgoingLetters().find(x => x.id === btn.dataset.id);
      if (item) cloneOutgoingItem(item);
    });
  });
}

async function downloadOutgoingItem(item) {
  if (item.documentSnapshot) {
    showToast(`Menyusun berkas .docx untuk "${item.title}"...`, 'info');
    try {
      const blob = await renderDocx(item.documentSnapshot, 'blob');
      const safeTitle = `${item.sequenceNumber || 'out'}_${item.title || 'surat'}`
        .replace(/[^a-zA-Z0-9_\-]/g, '_')
        .slice(0, 50);
      downloadFile(blob, `${safeTitle}.docx`);
      showToast('Berkas Surat Keluar berhasil diunduh.', 'success');
      return;
    } catch (err) {
      showToast(`Gagal mengunduh: ${err.message}`, 'error');
      return;
    }
  }
  if (item.uploadedDocx && item.uploadedDocx.rawBytes) {
    showToast(`Menyusun berkas .docx untuk "${item.title}"...`, 'info');
    try {
      const { applyCanvasEditsToDocx } = await import('../docx/patch.js');
      const patchedBlob = await applyCanvasEditsToDocx(item.uploadedDocx.rawBytes, item.uploadedDocx.paragraphs || [], { outputType: 'blob' });
      const safeTitle = `${item.sequenceNumber || 'out'}_${item.title || 'surat'}`
        .replace(/[^a-zA-Z0-9_\-]/g, '_')
        .slice(0, 50);
      downloadFile(patchedBlob, `${safeTitle}.docx`);
      showToast('Berkas Surat Keluar berhasil diunduh.', 'success');
      return;
    } catch (err) {
      showToast(`Gagal mengunduh: ${err.message}`, 'error');
      return;
    }
  }
  showToast('Snapshot dokumen tidak tersedia untuk item ini.', 'warning');
}

function cloneOutgoingItem(item) {
  if (store.cloneOutgoingToNewDraft(item.id)) {
    showToast(`Isi surat "${item.title}" disalin menjadi draf baru dengan nomor urut berikutnya (arsip lama tetap terkunci).`, 'success');
    const canvasTab = document.querySelector('.nav-tab[data-tab="canvas"]');
    if (canvasTab) canvasTab.click();
  }
}

function renderReadOnlyDocumentHtml(item) {
  const doc = item.documentSnapshot;
  if (!doc || !Array.isArray(doc.blocks)) {
    return `
      <div style="background: white; padding: 28px; border-radius: 4px; font-family: 'Times New Roman', serif; line-height: 1.5;">
        <h3 style="text-align: center; margin: 0 0 4px 0;">${escapeHtml(item.title)}</h3>
        <div style="text-align: center; margin-bottom: 18px;">Nomor: ${escapeHtml(item.nomorSurat)}</div>
        <p><strong>Atas Nama / Penerima:</strong> ${escapeHtml(item.recipientName)}</p>
        <p>${escapeHtml(item.summary || '-')}</p>
      </div>
    `;
  }

  const parts = [];
  for (const b of doc.blocks) {
    if (b.type === 'letterKop' && b.image) {
      parts.push(`<div style="text-align: center; margin-bottom: 14px;"><img src="${b.image}" style="max-width: 100%; height: auto;" alt="Kop"></div>`);
    } else if (b.type === 'letterTitle') {
      parts.push(`
        <div style="text-align: center; margin-bottom: 14px;">
          <div style="font-weight: bold; text-decoration: underline; font-size: 13.5pt;">${escapeHtml(b.title)}</div>
          <div style="font-size: 11.5pt;">${escapeHtml(b.nomor)}</div>
        </div>
      `);
      if (b.metaRecipient && !b.metaRecipient.isDeleted) {
        const metaRows = b.metaRecipient.metaRows || [];
        const recLines = b.metaRecipient.recipientLines || [];
        if (metaRows.length > 0) {
          parts.push(`
            <table style="margin-bottom: 8px; border-collapse: collapse;">
              ${metaRows.map(r => `<tr><td style="width: 72px;">${escapeHtml(r.label)}</td><td style="width: 10px;">:</td><td>${escapeHtml(r.value)}</td></tr>`).join('')}
            </table>
          `);
        }
        if (recLines.length > 0) {
          parts.push(`<div style="margin-bottom: 12px;">${recLines.map(l => `<div>${escapeHtml(l)}</div>`).join('')}</div>`);
        }
      }
    } else if (b.type === 'paragraph') {
      const txt = b.runs ? b.runs.map(r => r.text).join('') : (b.text || '');
      if (txt.trim()) {
        parts.push(`<p style="text-align: justify; margin: 0 0 10px 0; line-height: 1.4;">${escapeHtml(txt)}</p>`);
      }
    } else if (b.type === 'identityTable' && Array.isArray(b.rows)) {
      parts.push(`
        <table style="width: calc(100% - 16px); margin: 6px 0 12px 16px; border-collapse: collapse;">
          ${b.rows.map(r => `<tr><td style="width: 30%; padding: 2px 4px; vertical-align: top;">${escapeHtml(r.label)}</td><td style="padding: 2px 4px; font-weight: bold;">${escapeHtml(r.value)}</td></tr>`).join('')}
        </table>
      `);
    } else if (b.type === 'letterSignature') {
      parts.push(`
        <div style="display: flex; justify-content: flex-end; margin-top: 24px;">
          <div style="width: 240px;">
            <div>${escapeHtml(b.city || 'Garut')}, ${escapeHtml(b.date || '')}</div>
            <div>${escapeHtml(b.title || 'Kepala Sekolah')},</div>
            <div style="height: 50px;"></div>
            <div style="font-weight: bold; text-decoration: underline;">${escapeHtml(b.name || '')}</div>
            <div>${b.nip && b.nip !== '-' ? `NIP. ${escapeHtml(b.nip)}` : 'Kepala SMKS IT Al-Hawari'}</div>
          </div>
        </div>
      `);
    }
  }

  return `
    <div style="background: white; padding: 28px 34px; border-radius: 4px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); font-family: 'Times New Roman', serif; font-size: 11.5pt; color: #0f172a; user-select: text;">
      ${parts.join('')}
    </div>
  `;
}
