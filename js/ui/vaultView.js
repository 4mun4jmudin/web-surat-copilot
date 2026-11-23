/**
 * vaultView.js - Antarmuka Gudang Penyimpanan Surat (Letter Vault) & Knowledge Base AI
 * Mendukung penyimpanan, pencarian instan, duplikasi, pembukaan kembali ke Lembar Kerja Word,
 * unggah berkas .docx luar dengan preservasi format, dan pengaturan referensi AI.
 */

import { store } from '../state.js';
import { parseDocxForEditing, applyCanvasEditsToDocx } from '../docx/patch.js';
import { renderDocx } from '../docx/render.js';
import { downloadFile } from '../utils/download.js';
import { showToast } from './notification.js';
import { uiIcon } from './icons.js';
import {
  SCHOOL_UNITS,
  adaptDocumentModelForUnit,
  normalizeUnitId
} from '../letters/schoolProfile.js';
import { extractDocumentMetadata, buildDocumentDownloadFilename } from '../letters/smartNaming.js';
import { buildMetaAutoScanBarHtml, wireMetaAutoScanBar } from './vaultSaveHelper.js';

let currentSearchQuery = '';
let currentCategoryFilter = 'all';
let filterOnlyAiRef = false;
let filterOnlyMaster = false;
let currentVaultSort = 'master_first';
let currentVaultViewMode = 'grid';
let selectedVaultTargetMode = 'ask'; // 'ask' | 'smp' | 'smk'

/**
 * Menghitung jarak Damerau-Levenshtein (mendukung salah ketik huruf, kurang/lebih huruf, dan huruf tertukar posisi seperti "surta" -> "surat")
 */
export function damerauLevenshtein(a = '', b = '') {
  const s = a.toLowerCase();
  const t = b.toLowerCase();
  const m = s.length;
  const n = t.length;
  if (m === 0) return n;
  if (n === 0) return m;

  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = s[i - 1] === t[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,       // penghapusan
        dp[i][j - 1] + 1,       // penyisipan
        dp[i - 1][j - 1] + cost // substitusi
      );
      // Transposisi 2 huruf bersebelahan (contoh: "surta" <-> "surat", "undnagan" <-> "undangan")
      if (i > 1 && j > 1 && s[i - 1] === t[j - 2] && s[i - 2] === t[j - 1]) {
        dp[i][j] = Math.min(dp[i][j], dp[i - 2][j - 2] + cost);
      }
    }
  }
  return dp[m][n];
}

/**
 * Kemiripan Bigram (Dice Coefficient) untuk menangani salah ketik di tengah kata panjang
 */
function bigramSimilarity(a = '', b = '') {
  const s1 = a.toLowerCase();
  const s2 = b.toLowerCase();
  if (s1 === s2) return 1;
  if (s1.length < 2 || s2.length < 2) return 0;

  const bg1 = new Map();
  for (let i = 0; i < s1.length - 1; i++) {
    const bg = s1.slice(i, i + 2);
    bg1.set(bg, (bg1.get(bg) || 0) + 1);
  }

  let intersection = 0;
  for (let i = 0; i < s2.length - 1; i++) {
    const bg = s2.slice(i, i + 2);
    const count = bg1.get(bg) || 0;
    if (count > 0) {
      bg1.set(bg, count - 1);
      intersection++;
    }
  }
  return (2 * intersection) / ((s1.length - 1) + (s2.length - 1));
}

/**
 * Mengecek apakah karakter-karakter query muncul berurutan (subsequence) seperti singkatan "ketrngn" -> "keterangan"
 */
function subsequenceRatio(q = '', target = '') {
  if (!q || !target || q.length > target.length) return 0;
  if (q[0] !== target[0]) return 0; // huruf awal harus sama agar akurat
  let qi = 0;
  for (let ti = 0; ti < target.length && qi < q.length; ti++) {
    if (q[qi] === target[ti]) qi++;
  }
  if (qi === q.length) {
    return 0.65 + 0.35 * (q.length / target.length);
  }
  return 0;
}

function normalizeSearchText(str = '') {
  return String(str || '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Menghitung skor kemiripan 1 kata kunci (token) terhadap sebuah teks target (mendukung typo)
 * Mengembalikan nilai 0.0 s/d 1.0
 */
function scoreTokenAgainstText(token, normalizedFieldText, fieldWords) {
  if (!token || !normalizedFieldText) return 0;

  // 1. Sama persis dengan salah satu kata
  for (const w of fieldWords) {
    if (w === token) return 1.0;
  }

  // 2. Awalan kata (prefix) saat pengguna sedang mengetik (misal "und" -> "undangan")
  for (const w of fieldWords) {
    if (w.startsWith(token)) {
      return 0.92 + 0.06 * (token.length / w.length);
    }
  }

  // 3. Substring langsung di dalam kalimat/nomor
  if (normalizedFieldText.includes(token)) {
    return 0.85;
  }

  // Jika baru 1 huruf dan tidak cocok di mana pun, jangan paksa typo
  if (token.length <= 1) return 0;

  // 4. Toleransi Typo (Edit Distance, Prefix Edit Distance, Bigram, Subsequence)
  let bestFuzzy = 0;
  for (const w of fieldWords) {
    if (!w) continue;

    // Bandingkan terhadap kata utuh
    const maxLen = Math.max(token.length, w.length);
    const distFull = damerauLevenshtein(token, w);
    const allowedErrors = token.length <= 3 ? 1 : token.length <= 6 ? 2 : 3;

    if (distFull <= allowedErrors) {
      const editSim = 1 - (distFull / maxLen);
      if (editSim > bestFuzzy) bestFuzzy = editSim;
    }

    // Bandingkan terhadap potongan awal kata (jika pengguna mengetik setengah kata tapi typo, mis. "keterng" -> "keterangan")
    if (w.length > token.length && token.length >= 3) {
      const prefixW = w.slice(0, token.length);
      const distPref = damerauLevenshtein(token, prefixW);
      if (distPref <= (token.length <= 4 ? 1 : 2)) {
        const prefSim = (1 - (distPref / token.length)) * 0.9;
        if (prefSim > bestFuzzy) bestFuzzy = prefSim;
      }
    }

    // Bigram & Subsequence (mis. "undngn" -> "undangan", "keterngn" -> "keterangan")
    if (token.length >= 3 && w.length >= 3) {
      const bgSim = bigramSimilarity(token, w);
      if (bgSim >= 0.45 && bgSim * 0.92 > bestFuzzy) {
        bestFuzzy = bgSim * 0.92;
      }
      const subSim = subsequenceRatio(token, w);
      if (subSim >= 0.7 && subSim * 0.88 > bestFuzzy) {
        bestFuzzy = subSim * 0.88;
      }
    }
  }

  return bestFuzzy >= 0.52 ? bestFuzzy : 0;
}

/**
 * Mencari dan mengurutkan daftar surat berdasarkan kemiripan tertinggi dengan nama surat / atribut lainnya
 */
export function searchAndRankLetters(letters = [], rawQuery = '') {
  const qNorm = normalizeSearchText(rawQuery);
  if (!qNorm) {
    return letters.map(item => ({ item, score: 100, isFuzzy: false }));
  }

  const qTokens = qNorm.split(' ').filter(Boolean);
  const scoredResults = [];

  for (const item of letters) {
    const titleNorm = normalizeSearchText(item.title || '');
    const perihalNorm = normalizeSearchText(item.perihal || '');
    const nomorNorm = normalizeSearchText(item.nomorSurat || '');
    const categoryNorm = normalizeSearchText(item.category || '');
    const personNorm = normalizeSearchText(item.targetPerson || '');
    const summaryNorm = normalizeSearchText(item.summary || '');

    const fields = [
      { text: titleNorm, words: titleNorm.split(' ').filter(Boolean), weight: 4.5, isName: true },
      { text: perihalNorm, words: perihalNorm.split(' ').filter(Boolean), weight: 2.8, isName: true },
      { text: categoryNorm, words: categoryNorm.split(' ').filter(Boolean), weight: 2.0, isName: false },
      { text: nomorNorm, words: nomorNorm.split(' ').filter(Boolean), weight: 2.0, isName: false },
      { text: personNorm, words: personNorm.split(' ').filter(Boolean), weight: 1.8, isName: false },
      { text: summaryNorm, words: summaryNorm.split(' ').filter(Boolean), weight: 1.0, isName: false }
    ];

    let matchedTokens = 0;
    let totalTokenScore = 0;
    let exactSubstringHits = 0;

    for (const qt of qTokens) {
      let bestForToken = 0;
      let isExactToken = false;

      for (const f of fields) {
        const sim = scoreTokenAgainstText(qt, f.text, f.words);
        if (sim > 0) {
          const weighted = sim * f.weight * 100;
          if (weighted > bestForToken) {
            bestForToken = weighted;
          }
          if (sim >= 0.85) {
            isExactToken = true;
          }
        }
      }

      if (bestForToken > 0) {
        matchedTokens++;
        totalTokenScore += bestForToken;
        if (isExactToken) exactSubstringHits++;
      }
    }

    // Bonus kemiripan frasa utuh terhadap Nama Surat (Title) agar yang paling mendekati nama surat selalu di urutan #1
    const fullTitleBg = bigramSimilarity(qNorm, titleNorm);
    const fullPerihalBg = bigramSimilarity(qNorm, perihalNorm);
    const phraseBonus = Math.max(fullTitleBg * 180, fullPerihalBg * 110);

    if (matchedTokens > 0 || phraseBonus >= 55) {
      const tokenCoverageRatio = matchedTokens / qTokens.length;
      const finalScore = (totalTokenScore * (0.5 + 0.5 * tokenCoverageRatio)) + phraseBonus;

      // Syarat lolos: minimal 1 kata cocok (atau kemiripan frasa nama surat cukup dekat)
      if (tokenCoverageRatio >= 0.5 || phraseBonus >= 60 || qTokens.length === 1) {
        scoredResults.push({
          item,
          score: Math.round(finalScore),
          tokenCoverageRatio,
          isFuzzy: exactSubstringHits < matchedTokens || matchedTokens === 0
        });
      }
    }
  }

  // Jika pengguna salah ketik sangat berat sehingga belum ada yang lolos ambang normal,
  // tampilkan surat yang namanya paling mendekati agar pencarian tetap lancar
  if (scoredResults.length === 0 && qNorm.length >= 2) {
    for (const item of letters) {
      const titleNorm = normalizeSearchText(item.title || '');
      const perihalNorm = normalizeSearchText(item.perihal || '');
      const sim = Math.max(bigramSimilarity(qNorm, titleNorm), bigramSimilarity(qNorm, perihalNorm));
      if (sim >= 0.22) {
        scoredResults.push({
          item,
          score: Math.round(sim * 100),
          tokenCoverageRatio: sim,
          isFuzzy: true
        });
      }
    }
  }

  scoredResults.sort((a, b) => {
    if (b.tokenCoverageRatio !== a.tokenCoverageRatio) {
      return b.tokenCoverageRatio - a.tokenCoverageRatio;
    }
    return b.score - a.score;
  });

  return scoredResults;
}

const V_SVG_LAYERS_30 = uiIcon('layers', 24);
const V_SVG_DOC = (size = 18) => uiIcon('docClean', size);
const V_SVG_STAR = (size = 18) => uiIcon('star', size);
const V_SVG_BOLT = (size = 18) => uiIcon('bolt', size);
const V_SVG_SEARCH_15 = uiIcon('search', 15);
const V_SVG_UPLOAD_15 = uiIcon('upload', 14);
const V_SVG_CAL_12 = uiIcon('calendar', 12);
const V_SVG_USER_12 = uiIcon('user', 12);
const V_SVG_DL_13 = uiIcon('download', 13);
const V_SVG_COPY_13 = uiIcon('copy', 13);
const V_SVG_TRASH_13 = uiIcon('trash', 13);

function getCategoryPillIcon(cat) {
  if (cat === 'all') return uiIcon('grid', 13);
  if (cat === 'Surat Tugas') return uiIcon('briefcase', 13);
  if (cat === 'Surat Undangan') return uiIcon('mail', 13);
  if (cat === 'Dokumen Impor') return V_SVG_DL_13;
  return V_SVG_DOC(13);
}

export function renderVaultView(container) {
  if (!container) return;

  const allLetters = store.getLetterVault();
  const aiRefLetters = allLetters.filter(l => l.isAiReference !== false || l.isMasterTemplate);
  const masterLetters = allLetters.filter(l => l.isMasterTemplate);
  const baseCategories = ['Surat Keterangan', 'Surat Tugas', 'Surat Undangan', 'Surat Panggilan', 'Surat Keputusan', 'Dokumen Impor'];
  const presentCategories = Array.from(new Set(allLetters.map(l => l.category).filter(Boolean)));
  const categories = ['all', ...Array.from(new Set([...baseCategories, ...presentCategories]))];

  container.innerHTML = `
    <div id="dropzone-docx-vault" class="view-atmosphere" style="transition: outline 0.15s ease;">
      <div class="view-shell view-shell-fluid vault-view-container">

        <!-- HEADER GUDANG SURAT -->
        <div class="view-page-header" style="margin-bottom: 18px;">
          <div class="view-header-lead">
            <div class="view-header-badge">${V_SVG_LAYERS_30}</div>
            <div>
              <h2 class="view-title">
                <span>Gudang Template &amp; Draf Surat Sekolah</span>
                <span class="status-chip status-chip-green">${uiIcon('edit', 12)}<span>Dapat Diedit &amp; Disesuaikan Kembali</span></span>
              </h2>
              <p class="view-subtitle">
                Kumpulan template, draf, dan arsip surat sekolah yang dapat digunakan kembali atau diedit ulang di <strong>Lembar Kerja</strong>, serta menjadi referensi penulisan untuk AI <strong>Copilot</strong>. <span style="color:#64748b;">(Tarik &amp; lepas file <code>.docx</code> ke halaman ini untuk impor instan).</span>
              </p>
            </div>
          </div>

          <div class="view-header-actions">
            <button type="button" id="btn-vault-normalize-all" class="btn btn-secondary" style="padding: 8px 14px; font-weight: 600; border-radius: 8px; font-size: 13px; border-color: #bbf7d0; color: #156c3b; background: #f0fdf4;" title="Rapikan penamaan file arsip gudang secara otomatis mengikuti isi dokumen">
              ${uiIcon('sparkle', 14)}
              <span>Rapikan Nama Arsip</span>
            </button>
            <input type="file" id="input-upload-docx-vault" accept=".docx" class="hidden">
            <button id="btn-upload-docx-vault" class="btn btn-secondary" style="padding: 8px 15px; font-weight: 600; border-radius: 8px; font-size: 13px;">
              ${V_SVG_UPLOAD_15}
              <span>Unggah .docx</span>
            </button>
            <button id="btn-create-new-from-vault" class="btn btn-primary" style="padding: 8px 16px; font-weight: 700; border-radius: 8px; font-size: 13px;">
              ${uiIcon('plus', 14)}
              <span>Buat Surat Baru</span>
            </button>
          </div>
        </div>

        <!-- 3 KARTU STATISTIK KPI -->
        <div class="kpi-grid-3">
          <div class="kpi-card tint-green">
            <div class="kpi-icon-box kpi-icon-green">${V_SVG_DOC(20)}</div>
            <div>
              <div style="font-size: 12.5px; font-weight: 600; color: #334155;">Total Surat Tersimpan</div>
              <div class="kpi-value" style="font-size: 26px; font-weight: 800; color: #0f172a; line-height: 1.2; margin-top: 4px;">
                ${allLetters.length} <span style="font-size: 13px; font-weight: 500; color: #475569;">dokumen</span>
              </div>
            </div>
            <div class="kpi-watermark" style="color: #15803d;">${V_SVG_DOC(54)}</div>
          </div>

          <div class="kpi-card tint-amber">
            <div class="kpi-icon-box kpi-icon-amber">${V_SVG_STAR(20)}</div>
            <div>
              <div style="font-size: 12.5px; font-weight: 600; color: #334155;">Pembelajaran AI &amp; Master Terkunci</div>
              <div class="kpi-value" style="font-size: 24px; font-weight: 800; color: #b45309; line-height: 1.2; margin-top: 4px; display:flex; align-items:center; gap:7px; flex-wrap:wrap;">
                <span>${masterLetters.length}</span>
                <span style="display:inline-flex;align-items:center;gap:4px;font-size:12.5px;font-weight:700;color:#92400e;">${uiIcon('starFilled', 13)}<span>Master Terkunci</span></span>
              </div>
              <div style="font-size: 11px; color: #78350f; margin-top: 3px;">
                ${aiRefLetters.length} dari ${allLetters.length} arsip aktif dipelajari AI (Master diprioritaskan &amp; dilindungi)
              </div>
            </div>
            <div class="kpi-watermark" style="color: #d97706;">${V_SVG_STAR(56)}</div>
          </div>

          <div class="kpi-card tint-blue">
            <div class="kpi-icon-box kpi-icon-blue">${V_SVG_BOLT(20)}</div>
            <div>
              <div style="font-size: 13px; font-weight: 700; color: #0f172a;">Pencarian Cerdas (Anti-Typo)</div>
              <div style="font-size: 14.5px; font-weight: 800; color: #1e3a8a; margin-top: 3px;">Fuzzy Match &amp; Ranking Nama</div>
              <div style="font-size: 11.5px; color: #475569; margin-top: 2px;">Menemukan surat meski terdapat kesalahan penulisan.</div>
            </div>
            <div class="kpi-watermark" style="color: #2563eb;">${V_SVG_BOLT(54)}</div>
          </div>
        </div>

        <!-- SEARCH BAR + KATEGORI DROPDOWN + SORT DROPDOWN -->
        <div class="filter-bar-card">
          <div style="position: relative; flex: 1; min-width: 260px;">
            <span class="icon-input-leading" style="color: #64748b;">${V_SVG_SEARCH_15}</span>
            <input type="text" id="input-vault-search" class="icon-input-control" style="padding-right: 34px;"
              placeholder="Ketik nama surat, perihal, nomor, atau kata kunci lainnya..." autocomplete="off" spellcheck="false" value="${escapeHtml(currentSearchQuery)}">
            <button type="button" id="btn-clear-vault-search" class="${currentSearchQuery ? '' : 'hidden'}" title="Bersihkan pencarian"
              style="position: absolute; right: 8px; top: 50%; transform: translateY(-50%); border: none; background: #e2e8f0; color: #475569; width: 22px; height: 22px; border-radius: 50%; cursor: pointer; display: inline-flex; align-items: center; justify-content: center;">${uiIcon('x', 11)}</button>
          </div>

          <div class="icon-input-wrap" style="width: 205px;">
            <span class="icon-input-leading">${getCategoryPillIcon('all')}</span>
            <select id="select-vault-category" class="icon-input-control" style="cursor: pointer;">
              ${categories.map(cat => `
                <option value="${cat}" ${currentCategoryFilter === cat ? 'selected' : ''}>${cat === 'all' ? 'Semua Kategori' : cat}</option>
              `).join('')}
            </select>
          </div>

          <div class="icon-input-wrap" style="width: 205px;">
            <span class="icon-input-leading">${uiIcon('sort', 14)}</span>
            <select id="select-vault-sort" class="icon-input-control" style="cursor: pointer;">
              <option value="master_first" ${currentVaultSort === 'master_first' ? 'selected' : ''}>Urut: Master Terkunci Dulu</option>
              <option value="newest" ${currentVaultSort === 'newest' ? 'selected' : ''}>Urut: Terbaru</option>
              <option value="oldest" ${currentVaultSort === 'oldest' ? 'selected' : ''}>Urut: Terlama</option>
              <option value="title_asc" ${currentVaultSort === 'title_asc' ? 'selected' : ''}>Urut: Judul (A - Z)</option>
              <option value="nomor_asc" ${currentVaultSort === 'nomor_asc' ? 'selected' : ''}>Urut: Nomor Surat</option>
            </select>
          </div>
        </div>

        <!-- PILL KATEGORI + TOGGLE GRID/LIST -->
        <div class="filter-pills-bar">
          <div style="display: flex; gap: 7px; flex-wrap: wrap; align-items: center;">
            ${categories.map(cat => {
              const isAct = currentCategoryFilter === cat;
              const label = cat === 'all' ? 'Semua Surat' : cat;
              return `
                <button type="button" class="pill-tab ${isAct ? 'active' : ''} btn-filter-cat" data-cat="${cat}">
                  ${getCategoryPillIcon(cat)}
                  <span>${label}</span>
                </button>
              `;
            }).join('')}

            <button type="button" class="pill-tab ${filterOnlyMaster ? 'pill-gold-active' : ''} btn-toggle-master-filter" title="Tampilkan hanya Master Template bertanda Bintang &amp; Terkunci">
              ${uiIcon('starFilled', 13)}
              <span>Master Terkunci (${masterLetters.length})</span>
            </button>

            <button type="button" class="pill-tab ${filterOnlyAiRef ? 'pill-amber-active' : ''} btn-toggle-ai-filter">
              ${uiIcon('sparkle', 13)}
              <span>${filterOnlyAiRef ? 'Acuan AI Aktif' : 'Hanya Acuan AI'}</span>
            </button>
          </div>

          <div class="vault-view-toggle-group">
            <button type="button" class="vault-view-toggle-btn btn-vault-view-mode ${currentVaultViewMode === 'grid' ? 'active' : ''}" data-mode="grid" title="Tampilan Kartu Grid">
              ${uiIcon('grid', 15)}
            </button>
            <button type="button" class="vault-view-toggle-btn btn-vault-view-mode ${currentVaultViewMode === 'list' ? 'active' : ''}" data-mode="list" title="Tampilan Daftar Tabel">
              ${uiIcon('list', 15)}
            </button>
          </div>
        </div>

        <!-- PEMILIH TARGET SEKOLAH & KEPALA SEKOLAH (SMP / SMK) -->
        <div class="vault-unit-target-bar">
          <div class="vault-unit-target-lead">
            ${uiIcon('building', 15)}
            <span><strong>Target Sekolah &amp; Pejabat Kepsek saat Buka/Edit:</strong></span>
          </div>
          <div class="vault-unit-target-pills">
            <button type="button" class="pill-tab ${selectedVaultTargetMode === 'ask' ? 'active' : ''} btn-vault-unit-mode" data-unit-mode="ask" title="Tanyakan SMP atau SMK setiap kali membuka dokumen dari Gudang">
              <span>Pilih Saat Buka (SMP / SMK)</span>
            </button>
            <button type="button" class="pill-tab ${selectedVaultTargetMode === 'smp' ? 'active' : ''} btn-vault-unit-mode" data-unit-mode="smp" title="Buka &amp; adaptasi otomatis ke SMP IT Al-Hawari (Kepsek: ${escapeHtml(SCHOOL_UNITS.smp.principalName)})">
              <span>SMP IT &bull; Kepsek: ${escapeHtml(SCHOOL_UNITS.smp.principalName)}</span>
            </button>
            <button type="button" class="pill-tab ${selectedVaultTargetMode === 'smk' ? 'active' : ''} btn-vault-unit-mode" data-unit-mode="smk" title="Buka &amp; adaptasi otomatis ke SMK IT Al-Hawari (Kepsek: ${escapeHtml(SCHOOL_UNITS.smk.principalName)})">
              <span>SMK IT &bull; Kepsek: ${escapeHtml(SCHOOL_UNITS.smk.principalName)}</span>
            </button>
          </div>
        </div>

        <!-- CONTAINER DAFTAR KARTU ARSIP SURAT (DI-UPDATE TANPA MERESET INPUT PENCARIAN) -->
        <div id="vault-cards-list-area">
          ${buildVaultCardsHtml(allLetters)}
        </div>

      </div>
    </div>
  `;

  bindVaultEvents(container);
  bindVaultCardActions(container);
}

function buildVaultMasterLockBtnHtml(letter) {
  const isMaster = Boolean(letter.isMasterTemplate);
  const titleAttr = isMaster
    ? 'Dokumen Master bagus & terkunci (Prioritas utama acuan AI & terlindung dari hapus/timpa). Klik untuk membuka kunci.'
    : 'Beri Bintang & Kunci sebagai Master Template terbaik untuk acuan prioritas AI Copilot';
  return `
    <button type="button" class="vault-master-pill ${isMaster ? 'is-master' : ''} btn-toggle-master-lock" data-id="${letter.id}" title="${titleAttr}">
      ${isMaster ? uiIcon('starFilled', 12) : uiIcon('star', 12)}
      <span>${isMaster ? 'Master Terkunci' : 'Jadikan Master'}</span>
      ${isMaster ? uiIcon('lock', 11) : ''}
    </button>
  `;
}

function buildVaultAiRefBtnHtml(letter) {
  const isRef = letter.isAiReference !== false || Boolean(letter.isMasterTemplate);
  const titleAttr = letter.isMasterTemplate
    ? 'Master Terkunci otomatis aktif sebagai acuan prioritas AI.'
    : (isRef ? 'Dipakai sebagai referensi belajar AI Copilot. Klik untuk nonaktifkan.' : 'Aktifkan agar dipelajari AI Copilot.');
  return `
    <button type="button" class="btn-toggle-ai-ref" data-id="${letter.id}" title="${titleAttr}"
      style="display:inline-flex; align-items:center; gap:4px; background: ${isRef ? '#fef3c7' : '#f8fafc'}; border: 1px solid ${isRef ? '#fcd34d' : '#cbd5e1'}; color: ${isRef ? '#b45309' : '#64748b'}; font-size: 11px; font-weight: 700; padding: 3px 9px; border-radius: 999px; cursor: pointer;">
      ${isRef ? uiIcon('sparkle', 11) : uiIcon('eyeOff', 11)}
      <span>${isRef ? 'Dipelajari AI' : 'Nonaktif AI'}</span>
    </button>
  `;
}

function buildVaultCardsHtml(allLetters) {
  const baseFiltered = allLetters.filter(item => {
    if (filterOnlyMaster && !item.isMasterTemplate) return false;
    if (filterOnlyAiRef && item.isAiReference === false && !item.isMasterTemplate) return false;
    if (currentCategoryFilter !== 'all' && item.category !== currentCategoryFilter) return false;
    return true;
  });

  const hasActiveQuery = currentSearchQuery.trim().length > 0;
  let ranked = searchAndRankLetters(baseFiltered, currentSearchQuery);

  if (!hasActiveQuery) {
    ranked = [...ranked].sort((a, b) => {
      const la = a.item;
      const lb = b.item;
      if (currentVaultSort === 'master_first') {
        const ma = la.isMasterTemplate ? 1 : 0;
        const mb = lb.isMasterTemplate ? 1 : 0;
        if (mb !== ma) return mb - ma;
        return new Date(lb.createdAt || 0) - new Date(la.createdAt || 0);
      }
      if (currentVaultSort === 'oldest') {
        return new Date(la.createdAt || 0) - new Date(lb.createdAt || 0);
      }
      if (currentVaultSort === 'title_asc') {
        return (la.title || '').localeCompare(lb.title || '', 'id');
      }
      if (currentVaultSort === 'nomor_asc') {
        return (la.nomorSurat || '').localeCompare(lb.nomorSurat || '', 'id');
      }
      return new Date(lb.createdAt || 0) - new Date(la.createdAt || 0);
    });
  }

  if (ranked.length === 0) {
    return `
      <div class="empty-illustration-card">
        <div class="view-header-badge" style="width: 60px; height: 60px; border-radius: 16px; margin: 0 auto 14px auto;">
          ${uiIcon('inbox', 28)}
        </div>
        <h4 style="color: #0f172a; font-size: 16px; font-weight: 800; margin-bottom: 6px;">Tidak Ada Surat yang Cocok</h4>
        <p style="color: #64748b; font-size: 13px; max-width: 460px; margin: 0 auto 16px auto;">
          Belum ada surat pada kategori ini yang mendekati kata kunci <b>"${escapeHtml(currentSearchQuery.trim())}"</b>.
        </p>
        <button id="btn-empty-reset" class="btn btn-secondary btn-sm">${uiIcon('reset', 13)}<span>Reset Filter Pencarian</span></button>
      </div>
    `;
  }

  const queryHeaderHtml = hasActiveQuery ? `
    <div style="margin-bottom: 12px; font-size: 12.5px; color: #334155; background: rgba(255,255,255,0.9); border: 1px solid #dbe3eb; border-radius: 10px; padding: 8px 14px; display: flex; align-items: center; justify-content: space-between;">
      <span>Menampilkan <strong>${ranked.length}</strong> surat yang paling mendekati pencarian <b>"${escapeHtml(currentSearchQuery.trim())}"</b>:</span>
    </div>
  ` : '';

  if (currentVaultViewMode === 'list') {
    return `
      ${queryHeaderHtml}
      <div class="ui-card" style="padding: 0; overflow: hidden;">
        <div class="responsive-table-wrap">
          <table class="enterprise-table responsive-data-table vault-list-table">
            <thead>
              <tr>
                <th style="width: 17%; min-width: 155px;">Kategori &amp; Status</th>
                <th style="width: 30%; min-width: 230px;">Judul &amp; Keperluan Dokumen</th>
                <th style="width: 16%; min-width: 150px;">Nomor &amp; Tanggal Surat</th>
                <th style="width: 14%; min-width: 135px;">Pihak Terkait</th>
                <th style="width: 9%; min-width: 105px; text-align: center;">Acuan AI</th>
                <th style="width: 14%; min-width: 220px; text-align: right;">Tindakan</th>
              </tr>
            </thead>
            <tbody>
              ${ranked.map(({ item: letter }) => {
                const isLocked = Boolean(letter.isLockedTemplate || letter.isMasterTemplate);
                return `
                <tr style="${letter.isMasterTemplate ? 'background: rgba(254, 249, 195, 0.35);' : ''}">
                  <td data-label="Kategori &amp; Status">
                    <div class="td-value-stack">
                      <div style="display: flex; align-items: center; gap: 5px; flex-wrap: wrap;">
                        <span class="status-chip ${letter.category === 'Dokumen Impor' ? 'status-chip-blue' : 'status-chip-green'}" style="font-size: 10.5px;">
                          ${escapeHtml(letter.category || 'Surat Dinas')}
                        </span>
                        ${buildVaultMasterLockBtnHtml(letter)}
                      </div>
                    </div>
                  </td>
                  <td data-label="Judul &amp; Keperluan">
                    <div class="td-value-stack">
                      <div style="font-weight: 700; color: #0f172a; line-height: 1.38; overflow-wrap: anywhere;">${escapeHtml(letter.title)}</div>
                      ${letter.perihal && letter.perihal !== '-' && letter.perihal !== letter.title ? `<div style="font-size: 11.5px; color: #475569; margin-top: 3px; line-height: 1.4; overflow-wrap: anywhere;">${escapeHtml(letter.perihal)}</div>` : ''}
                    </div>
                  </td>
                  <td data-label="Nomor &amp; Tanggal" style="font-size: 12px; color: #334155;">
                    <div class="td-value-stack">
                      <div style="font-family: Consolas, monospace; font-weight: 700; overflow-wrap: anywhere;">${escapeHtml(letter.nomorSurat || '-')}</div>
                      <div style="font-size: 11.5px; color: #64748b; margin-top: 3px;">${escapeHtml(letter.tanggalSurat || '-')}</div>
                    </div>
                  </td>
                  <td data-label="Pihak Terkait" style="font-size: 12px; color: #334155;">
                    <div class="td-value-stack">
                      <span style="font-weight: 600; overflow-wrap: anywhere;">${escapeHtml(letter.targetPerson || '-')}</span>
                    </div>
                  </td>
                  <td data-label="Acuan AI" style="text-align: center;">
                    ${buildVaultAiRefBtnHtml(letter)}
                  </td>
                  <td data-label="Tindakan" class="td-actions-cell" style="text-align: right;">
                    <div class="table-action-group">
                      <button class="btn btn-sm btn-primary btn-load-letter" data-id="${letter.id}">${uiIcon('docClean', 13)}<span>Buka di Word</span></button>
                      <button class="btn btn-sm btn-secondary btn-download-vault-letter" data-id="${letter.id}">${V_SVG_DL_13}<span>.docx</span></button>
                      <button class="btn btn-sm btn-secondary btn-square-sm btn-rename-vault-letter" data-id="${letter.id}" title="Rapikan / Ubah Nama Dokumen (Format Sistem / AI)">${uiIcon('edit', 13)}</button>
                      <button class="btn btn-sm btn-secondary btn-vault-history" data-id="${letter.id}" title="Lihat &amp; Pulihkan Riwayat Versi">${uiIcon('history', 13)}<span>${Array.isArray(letter.versionHistory) ? letter.versionHistory.length : 1}</span></button>
                      <button class="btn btn-sm btn-secondary btn-square-sm btn-duplicate-letter" data-id="${letter.id}" title="Duplikasi sebagai salinan baru">${V_SVG_COPY_13}</button>
                      <button class="btn btn-sm btn-secondary btn-square-sm btn-delete-vault-letter" data-id="${letter.id}" data-locked="${isLocked ? '1' : '0'}" style="color: ${isLocked ? '#94a3b8' : '#dc2626'}; opacity: ${isLocked ? '0.65' : '1'};" title="${isLocked ? 'Terkunci (Master Template). Buka kunci untuk menghapus.' : 'Hapus dari Gudang'}">${isLocked ? uiIcon('lock', 13) : V_SVG_TRASH_13}</button>
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
  }

  return `
    ${queryHeaderHtml}
    <div class="vault-cards-grid" style="display: grid; grid-template-columns: repeat(auto-fill, minmax(min(100%, 350px), 1fr)); gap: 16px; margin-bottom: 24px;">
      ${ranked.map(({ item: letter, isFuzzy }, idx) => {
        const isMaster = Boolean(letter.isMasterTemplate);
        const isLocked = Boolean(letter.isLockedTemplate || isMaster);
        const cat = letter.category || 'Surat Dinas';
        const isImport = cat === 'Dokumen Impor';
        const isUndangan = cat === 'Surat Undangan';
        const chipClass = isImport ? 'status-chip-blue' : isUndangan ? 'status-chip-purple' : 'status-chip-green';
        const thumbClass = isImport ? 'thumb-blue' : isUndangan ? 'thumb-purple' : 'thumb-green';
        const isTopMatch = hasActiveQuery && idx === 0;

        return `
          <div class="card vault-card vault-card-modern ${isMaster ? 'vault-card-master' : ''} ${isTopMatch ? 'is-top-match' : ''}">
            <div>
              <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; gap: 6px; flex-wrap: wrap;">
                <div style="display: flex; align-items: center; gap: 6px; flex-wrap: wrap;">
                  <span class="status-chip ${chipClass}" style="font-size: 11px;">
                    ${escapeHtml(cat)}
                  </span>
                  ${isTopMatch ? `<span class="status-chip status-chip-green" style="font-size: 10px;">${uiIcon('target', 11)}<span>Paling Dekat</span></span>` : (hasActiveQuery && isFuzzy ? `<span class="status-chip status-chip-blue" style="font-size: 10px;">${uiIcon('check', 11)}<span>Toleransi Typo</span></span>` : '')}
                </div>
                <div style="display: flex; align-items: center; gap: 5px; flex-wrap: wrap;">
                  ${buildVaultMasterLockBtnHtml(letter)}
                  ${buildVaultAiRefBtnHtml(letter)}
                </div>
              </div>

              <div style="display: flex; align-items: flex-start; gap: 12px; margin-bottom: 11px;">
                <div class="vault-doc-thumb ${thumbClass}">
                  ${V_SVG_DOC(22)}
                </div>
                <div style="flex: 1; min-width: 0;">
                  <strong style="font-size: 14px; font-weight: 800; color: #0f172a; display: block; line-height: 1.35; margin-bottom: 4px; overflow-wrap: anywhere;">
                    ${escapeHtml(letter.title)}
                  </strong>
                  <div style="font-size: 11.5px; font-weight: 600; color: #334155; margin-bottom: 4px; word-break: break-word;">
                    Nomor: ${escapeHtml(letter.nomorSurat || '-')}
                  </div>
                  ${letter.perihal && letter.perihal !== '-' && letter.perihal !== letter.title ? `<div style="font-size: 11.5px; color: #475569; margin-bottom: 5px; line-height: 1.35; overflow-wrap: anywhere;"><strong>Keperluan:</strong> ${escapeHtml(letter.perihal)}</div>` : ''}
                  <div style="font-size: 11.5px; color: #64748b; display: flex; align-items: center; gap: 10px; flex-wrap: wrap;">
                    <span style="display: inline-flex; align-items: center; gap: 4px;">${V_SVG_CAL_12} ${escapeHtml(letter.tanggalSurat || '-')}</span>
                    ${letter.targetPerson && letter.targetPerson !== '-' ? `<span style="display: inline-flex; align-items: center; gap: 4px; overflow-wrap: anywhere;">${V_SVG_USER_12} ${escapeHtml(letter.targetPerson)}</span>` : ''}
                  </div>
                </div>
              </div>

              <p class="vault-summary-box">
                ${escapeHtml(letter.summary || 'Dokumen persuratan resmi sekolah.')}
              </p>
            </div>

            <div style="padding-top: 6px; border-top: 1px solid #f1f5f9; display: flex; gap: 6px; align-items: center; flex-wrap: wrap;">
              <button class="btn btn-sm btn-primary btn-load-letter" data-id="${letter.id}" style="flex: 1; min-width: 110px; height: 34px; font-weight: 600; border-radius: 8px;">
                ${uiIcon('docClean', 13)}
                <span>Buka di Word</span>
              </button>
              <button class="btn btn-sm btn-secondary btn-download-vault-letter" data-id="${letter.id}" title="Unduh berkas .docx" style="height: 34px; padding: 0 10px; border-radius: 8px; font-weight: 600;">
                ${V_SVG_DL_13}
                <span>.docx</span>
              </button>
              <button class="btn btn-sm btn-secondary btn-rename-vault-letter" data-id="${letter.id}" title="Rapikan / Ubah Nama Dokumen (Format Sistem / Pindai AI)" style="height: 34px; width: 34px; padding: 0; border-radius: 8px;">
                ${uiIcon('edit', 13)}
              </button>
              <button class="btn btn-sm btn-secondary btn-vault-history" data-id="${letter.id}" title="Riwayat Versi (${Array.isArray(letter.versionHistory) ? letter.versionHistory.length : 1} versi)" style="height: 34px; padding: 0 9px; border-radius: 8px; font-weight: 600;">
                ${uiIcon('history', 13)}${Array.isArray(letter.versionHistory) && letter.versionHistory.length > 1 ? `<span>${letter.versionHistory.length}</span>` : ''}
              </button>
              <button class="btn btn-sm btn-secondary btn-duplicate-letter" data-id="${letter.id}" title="Duplikasi sebagai salinan baru (tanpa kunci)" style="height: 34px; width: 34px; padding: 0; border-radius: 8px;">
                ${V_SVG_COPY_13}
              </button>
              <button class="btn btn-sm btn-secondary btn-delete-vault-letter" data-id="${letter.id}" data-locked="${isLocked ? '1' : '0'}" title="${isLocked ? 'Terkunci sebagai Master Template. Buka kunci bintang dahulu untuk menghapus.' : 'Hapus dari Gudang'}" style="height: 34px; width: 34px; padding: 0; border-radius: 8px; color: ${isLocked ? '#94a3b8' : '#dc2626'}; opacity: ${isLocked ? '0.68' : '1'};">
                ${isLocked ? uiIcon('lock', 13) : V_SVG_TRASH_13}
              </button>
            </div>
          </div>
        `;
      }).join('')}
    </div>
  `;
}

function updateVaultCardsInPlace(container) {
  const listArea = container.querySelector('#vault-cards-list-area');
  const clearBtn = container.querySelector('#btn-clear-vault-search');
  if (clearBtn) {
    clearBtn.classList.toggle('hidden', !currentSearchQuery);
  }
  if (!listArea) return;
  const allLetters = store.getLetterVault();
  listArea.innerHTML = buildVaultCardsHtml(allLetters);
  bindVaultCardActions(container);
}

function bindVaultEvents(container) {
  // 1. Search input — update hanya daftar kartu agar fokus & posisi kursor pengetikan TIDAK PERNAH HILANG
  const searchInput = container.querySelector('#input-vault-search');
  const clearBtn = container.querySelector('#btn-clear-vault-search');

  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      currentSearchQuery = e.target.value;
      updateVaultCardsInPlace(container);
    });
  }

  if (clearBtn && searchInput) {
    clearBtn.addEventListener('click', () => {
      currentSearchQuery = '';
      searchInput.value = '';
      updateVaultCardsInPlace(container);
      searchInput.focus();
    });
  }

  // 2. Filter category (via pills & synced dropdown)
  container.querySelectorAll('.btn-filter-cat').forEach(btn => {
    btn.onclick = () => {
      currentCategoryFilter = btn.dataset.cat;
      renderVaultView(container);
    };
  });

  const catSelect = container.querySelector('#select-vault-category');
  if (catSelect) {
    catSelect.onchange = () => {
      currentCategoryFilter = catSelect.value;
      renderVaultView(container);
    };
  }

  // 2b. Sort dropdown (updates in place cleanly)
  const sortSelect = container.querySelector('#select-vault-sort');
  if (sortSelect) {
    sortSelect.onchange = () => {
      currentVaultSort = sortSelect.value;
      updateVaultCardsInPlace(container);
    };
  }

  // 2c. Grid / List view toggle
  container.querySelectorAll('.btn-vault-view-mode').forEach(btn => {
    btn.onclick = () => {
      currentVaultViewMode = btn.dataset.mode || 'grid';
      container.querySelectorAll('.btn-vault-view-mode').forEach(b => b.classList.toggle('active', b === btn));
      updateVaultCardsInPlace(container);
    };
  });

  // 3. Toggle only Master & only AI filters
  const btnToggleMasterFilter = container.querySelector('.btn-toggle-master-filter');
  if (btnToggleMasterFilter) {
    btnToggleMasterFilter.onclick = () => {
      filterOnlyMaster = !filterOnlyMaster;
      renderVaultView(container);
    };
  }

  const btnToggleAiFilter = container.querySelector('.btn-toggle-ai-filter');
  if (btnToggleAiFilter) {
    btnToggleAiFilter.onclick = () => {
      filterOnlyAiRef = !filterOnlyAiRef;
      renderVaultView(container);
    };
  }

  // 4. Pilihan Unit Sekolah & Pejabat TTD saat Buka/Edit File Gudang
  container.querySelectorAll('.btn-vault-unit-mode').forEach(btn => {
    btn.onclick = () => {
      selectedVaultTargetMode = btn.dataset.unitMode || 'ask';
      renderVaultView(container);
    };
  });

  // 5. Tombol Buat Surat Baru -> Switch ke tab canvas
  const btnCreate = container.querySelector('#btn-create-new-from-vault');
  if (btnCreate) {
    btnCreate.onclick = () => {
      const canvasTab = document.querySelector('.nav-tab[data-tab="canvas"]');
      if (canvasTab) canvasTab.click();
    };
  }

  // 5B. Rapikan penamaan semua file Gudang non-master sekaligus
  const btnNormalizeAll = container.querySelector('#btn-vault-normalize-all');
  if (btnNormalizeAll) {
    btnNormalizeAll.onclick = () => {
      const list = store.getLetterVault();
      let updatedCount = 0;
      for (const item of list) {
        if (!item || item.isLockedTemplate || item.isMasterTemplate) continue;
        const m = extractDocumentMetadata({
          docModel: item.document || null,
          uploadedDocx: item.uploadedDocx || null,
          schoolProfile: store.getEffectiveSchoolForDocument(item.document || item),
          fallbackTitle: item.title,
          fallbackNomor: item.nomorSurat,
          fallbackDate: item.tanggalSurat,
          fallbackPerihal: item.perihal,
          fallbackRecipient: item.targetPerson
        });
        if (m.formattedTitle && (m.formattedTitle !== item.title || m.perihal !== item.perihal || m.recipientName !== item.targetPerson)) {
          store.updateLetterVaultEntry(item.id, {
            title: m.formattedTitle,
            category: item.source === 'upload' && m.category ? m.category : (m.category || item.category),
            perihal: m.perihal || item.perihal,
            targetPerson: m.recipientName || item.targetPerson,
            tanggalSurat: m.tanggalSurat || item.tanggalSurat,
            summary: m.summary || item.summary
          });
          updatedCount += 1;
        }
      }
      if (updatedCount > 0) {
        renderVaultView(container);
        showToast(`✓ ${updatedCount} nama & metadata dokumen Gudang dirapikan otomatis!`, 'success');
      } else {
        showToast('Seluruh dokumen Gudang sudah memiliki penamaan yang rapih.', 'info');
      }
    };
  }

  // 6. Upload .docx handling (Drag & Drop & Button)
  const fileInput = container.querySelector('#input-upload-docx-vault');
  const dropzone = container.querySelector('#dropzone-docx-vault');
  const btnUpload = container.querySelector('#btn-upload-docx-vault');

  if (btnUpload && fileInput) {
    btnUpload.onclick = () => fileInput.click();
  }
  if (dropzone) {
    dropzone.ondragover = (e) => {
      e.preventDefault();
      dropzone.style.outline = '2px dashed #156c3b';
      dropzone.style.outlineOffset = '-8px';
    };
    dropzone.ondragleave = () => {
      dropzone.style.outline = 'none';
    };
    dropzone.ondrop = (e) => {
      e.preventDefault();
      dropzone.style.outline = 'none';
      if (e.dataTransfer.files.length > 0) {
        processUploadedDocx(e.dataTransfer.files[0], container);
      }
    };
  }

  if (fileInput) {
    fileInput.onchange = (e) => {
      const file = e.target.files[0];
      if (file) processUploadedDocx(file, container);
    };
  }
}

function bindVaultCardActions(container) {
  // Reset empty filter
  const btnReset = container.querySelector('#btn-empty-reset');
  if (btnReset) {
    btnReset.onclick = () => {
      currentSearchQuery = '';
      currentCategoryFilter = 'all';
      filterOnlyAiRef = false;
      filterOnlyMaster = false;
      renderVaultView(container);
      const input = container.querySelector('#input-vault-search');
      if (input) input.focus();
    };
  }

  // 7A. Toggle Master Template Bintang & Kunci
  container.querySelectorAll('.btn-toggle-master-lock').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      const id = btn.dataset.id;
      const updated = store.toggleMasterTemplate(id);
      if (updated?.isMasterTemplate) {
        showToast('Dokumen ditandai sebagai Master Template Terkunci & prioritas utama pembelajaran AI!', 'success');
      } else {
        showToast('Kunci Master Template dibuka.', 'info');
      }
      renderVaultView(container);
    };
  });

  // 7B. Toggle AI Reference per card
  container.querySelectorAll('.btn-toggle-ai-ref').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      const id = btn.dataset.id;
      store.toggleLetterAiReference(id);
      showToast('Status acuan pembelajaran AI diperbarui.', 'info');
      renderVaultView(container);
    };
  });

function promptUnitChoiceForVaultLetter(letter, container, onPicked) {
  if (selectedVaultTargetMode === 'smp' || selectedVaultTargetMode === 'smk') {
    onPicked(selectedVaultTargetMode);
    return;
  }

  const activeUnit = store.getActiveUnitId();
  container.querySelector('#vault-unit-choice-modal')?.remove();
  const overlay = document.createElement('div');
  overlay.id = 'vault-unit-choice-modal';
  overlay.className = 'modal-overlay';

  const units = [SCHOOL_UNITS.smp, SCHOOL_UNITS.smk];
  overlay.innerHTML = `
    <div class="modal-card" style="max-width: 620px; width: 95%;">
      <div class="modal-header">
        <h3>${uiIcon('building', 17)}<span>Pilih Sekolah &amp; Pejabat Penanda Tangan</span></h3>
        <button type="button" class="btn btn-sm btn-secondary btn-square-sm" id="btn-close-vault-unit" aria-label="Tutup">${uiIcon('x', 15)}</button>
      </div>
      <div class="modal-body" style="display: grid; gap: 12px;">
        <div style="font-size: 13px; color: #334155;">
          Dokumen <strong>${escapeHtml(letter?.title || 'Surat Gudang')}</strong> akan disesuaikan kop, kode nomor surat, serta nama Kepala Sekolah secara otomatis sesuai unit yang Anda pilih:
        </div>
        <div class="vault-unit-modal-grid">
          ${units.map(u => {
            const isActive = u.unitId === activeUnit;
            const accent = u.unitId === 'smp' ? '#1d4ed8' : '#15803d';
            const bgTint = u.unitId === 'smp' ? '#eff6ff' : '#f0fdf4';
            const borderTint = u.unitId === 'smp' ? '#bfdbfe' : '#bbf7d0';
            return `
              <button type="button" class="vault-unit-choice-card" data-pick-unit="${u.unitId}" style="border-color: ${borderTint};">
                <div style="display:flex; justify-content:space-between; align-items:center; gap:6px; margin-bottom:6px;">
                  <span style="font-size:11px; font-weight:800; padding:2px 8px; border-radius:999px; background:${bgTint}; color:${accent}; border:1px solid ${borderTint};">
                    ${u.unitShortBadge} &bull; ${u.schoolCode}
                  </span>
                  ${isActive ? `<span class="status-chip status-chip-green" style="font-size:10px;">Unit Aktif</span>` : ''}
                </div>
                <strong style="font-size:14px; color:#0f172a; display:block; margin-bottom:4px;">${escapeHtml(u.schoolName)}</strong>
                <div style="font-size:12px; color:#475569; margin-bottom:8px; line-height:1.45;">
                  <div>Kepsek: <strong style="color:#0f172a;">${escapeHtml(u.principalName)}</strong></div>
                  <div>NPSN: ${escapeHtml(u.npsn)} &bull; Pola: ${escapeHtml(u.numberingPattern.replace('{urut}', '001').replace('{bulanRomawi}', 'XI').replace('{tahun}', '2026'))}</div>
                </div>
                <div style="font-size:12px; font-weight:700; color:${accent}; display:inline-flex; align-items:center; gap:5px;">
                  ${uiIcon('docClean', 13)}
                  <span>Buka &amp; Edit untuk ${u.unitShortBadge} →</span>
                </div>
              </button>
            `;
          }).join('')}
        </div>
        <label style="display:inline-flex; align-items:center; gap:7px; font-size:12px; color:#475569; cursor:pointer;">
          <input type="checkbox" id="chk-remember-vault-unit" style="width:15px;height:15px;" />
          <span>Ingat pilihan sekolah ini selama di halaman Gudang Surat</span>
        </label>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" id="btn-cancel-vault-unit">Batal</button>
      </div>
    </div>
  `;

  container.appendChild(overlay);
  const close = () => overlay.remove();
  overlay.querySelector('#btn-close-vault-unit').onclick = close;
  overlay.querySelector('#btn-cancel-vault-unit').onclick = close;
  overlay.addEventListener('click', (ev) => {
    if (ev.target === overlay) close();
  });
  overlay.querySelectorAll('[data-pick-unit]').forEach(cardBtn => {
    cardBtn.onclick = () => {
      const pickedUnit = normalizeUnitId(cardBtn.dataset.pickUnit);
      const remember = Boolean(overlay.querySelector('#chk-remember-vault-unit')?.checked);
      if (remember) {
        selectedVaultTargetMode = pickedUnit;
      }
      close();
      onPicked(pickedUnit);
    };
  });
}

function openVaultLetterWithUnit(letter, container, targetUnitId) {
  const unitId = normalizeUnitId(targetUnitId);
  const unitInfo = SCHOOL_UNITS[unitId] || SCHOOL_UNITS.smk;
  const isSmartTarget = Boolean(letter?.preferSmartCanvas || (letter?.document && store.isSmartTemplateDocument(letter.document)));
  const success = store.loadLetterToCanvas(letter.id, { targetUnitId: unitId });
  if (success) {
    if (isSmartTarget) {
      showToast(`✓ Dimuat ke Word Pro untuk ${unitInfo.unitLabel} (Kepsek: ${unitInfo.principalName})!`, 'success');
      const smartTab = document.querySelector('.nav-tab[data-tab="smart-canvas"]');
      if (smartTab) smartTab.click();
    } else {
      showToast(`✓ Dimuat ke Lembar Kerja untuk ${unitInfo.unitLabel} (Kepsek: ${unitInfo.principalName})!`, 'success');
      const canvasTab = document.querySelector('.nav-tab[data-tab="canvas"]');
      if (canvasTab) canvasTab.click();
    }
  } else {
    showToast('Gagal memuat dokumen.', 'error');
  }
}

  // 8. Buka di Word Canvas / Word Pro (Pilih SMP atau SMK & Kepsek otomatis terpilih)
  container.querySelectorAll('.btn-load-letter').forEach(btn => {
    btn.onclick = () => {
      const id = btn.dataset.id;
      const letter = store.getLetterVault().find(l => l.id === id);
      if (!letter) return;
      promptUnitChoiceForVaultLetter(letter, container, (pickedUnit) => {
        openVaultLetterWithUnit(letter, container, pickedUnit);
      });
    };
  });

  // 9. Unduh dokumen langsung dari kartu
  container.querySelectorAll('.btn-download-vault-letter').forEach(btn => {
    btn.onclick = async () => {
      const id = btn.dataset.id;
      const letter = store.getLetterVault().find(l => l.id === id);
      if (!letter) return;

      const unitToUse = (selectedVaultTargetMode === 'smp' || selectedVaultTargetMode === 'smk')
        ? selectedVaultTargetMode
        : store.getActiveUnitId();

      showToast(`Menyiapkan berkas .docx untuk ${letter.title} (${unitToUse.toUpperCase()})...`, 'info');
      try {
        let blob;
        if (letter.uploadedDocx?.rawBytes) {
          blob = new Blob([letter.uploadedDocx.rawBytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        } else if (letter.document) {
          const adaptedDoc = adaptDocumentModelForUnit(letter.document, unitToUse, {
            activeSchoolProfile: store.getSchoolProfile(),
            kopGallery: store.getState().kopGallery || []
          });
          blob = await renderDocx(adaptedDoc, 'blob');
        } else {
          blob = await renderDocx(store.getState().document, 'blob');
        }

        const fileName = buildDocumentDownloadFilename({
          kind: 'vault',
          title: letter.title,
          perihal: letter.perihal,
          recipientName: letter.targetPerson,
          tanggalSurat: letter.tanggalSurat,
          unitId: unitToUse,
          schoolProfile: store.getSchoolProfile()
        });
        downloadFile(blob, fileName);
        showToast(`Berkas ${fileName} berhasil diunduh.`, 'success');
      } catch (err) {
        showToast(`Gagal mengunduh: ${err.message}`, 'error');
      }
    };
  });

  // 9A. Rapikan / Ubah Nama Dokumen Gudang (Format Sistem + Pindai AI)
  container.querySelectorAll('.btn-rename-vault-letter').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      const id = btn.dataset.id;
      const letter = store.getLetterVault().find(l => l.id === id);
      if (!letter) return;

      container.querySelector('#vault-rename-modal')?.remove();
      const overlay = document.createElement('div');
      overlay.id = 'vault-rename-modal';
      overlay.className = 'modal-overlay';
      const school = store.getEffectiveSchoolForDocument(letter.document || letter) || store.getSchoolProfile();
      const autoMeta = extractDocumentMetadata({
        docModel: letter.document || null,
        uploadedDocx: letter.uploadedDocx || null,
        schoolProfile: school,
        fallbackTitle: letter.title,
        fallbackNomor: letter.nomorSurat,
        fallbackDate: letter.tanggalSurat,
        fallbackPerihal: letter.perihal,
        fallbackRecipient: letter.targetPerson
      });

      overlay.innerHTML = `
        <div class="modal-card" style="max-width: 520px; width: 94%;">
          <div class="modal-header">
            <h3>${uiIcon('edit', 16)}<span>Rapikan Nama &amp; Metadata Arsip Gudang</span></h3>
            <button type="button" class="btn btn-sm btn-secondary btn-square-sm" id="btn-close-vrename">${uiIcon('x', 14)}</button>
          </div>
          <div class="modal-body" style="display:grid; gap:10px;">
            <div id="vrename-scan-slot">${buildMetaAutoScanBarHtml({ statusText: 'Pratinjau Format Sistem siap — atau klik Pindai AI' })}</div>
            <div><label class="form-label">Judul Arsip</label><input id="vr-title" class="form-input" value="${escapeHtml(letter.title || autoMeta.formattedTitle)}" /></div>
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:8px;">
              <div><label class="form-label">Kategori</label><input id="vr-cat" class="form-input" value="${escapeHtml(letter.category || autoMeta.category)}" /></div>
              <div><label class="form-label">Tanggal Surat</label><input id="vr-date" class="form-input" value="${escapeHtml(letter.tanggalSurat || autoMeta.tanggalSurat)}" /></div>
            </div>
            <div><label class="form-label">Pihak Terkait / Sasaran</label><input id="vr-person" class="form-input" value="${escapeHtml(letter.targetPerson || autoMeta.recipientName)}" /></div>
            <div><label class="form-label">Keperluan / Perihal</label><input id="vr-hal" class="form-input" value="${escapeHtml(letter.perihal || autoMeta.perihal)}" /></div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary" id="btn-cancel-vrename">Batal</button>
            <button type="button" class="btn btn-primary" id="btn-save-vrename">Simpan Perubahan Nama</button>
          </div>
        </div>
      `;
      container.appendChild(overlay);
      const cleanupScan = wireMetaAutoScanBar({
        rootEl: overlay.querySelector('.modal-body'),
        getContext: () => ({
          docModel: letter.document || null,
          uploadedDocx: letter.uploadedDocx || null,
          schoolProfile: school,
          fallbackTitle: letter.title,
          fallbackNomor: letter.nomorSurat,
          fallbackDate: letter.tanggalSurat
        }),
        fields: {
          titleEl: overlay.querySelector('#vr-title'),
          categoryEl: overlay.querySelector('#vr-cat'),
          dateEl: overlay.querySelector('#vr-date'),
          recipientEl: overlay.querySelector('#vr-person'),
          perihalEl: overlay.querySelector('#vr-hal')
        },
        forTemplate: false,
        autoAiOnOpen: false
      });
      const closeModal = () => {
        cleanupScan();
        overlay.remove();
      };
      overlay.querySelector('#btn-close-vrename').onclick = closeModal;
      overlay.querySelector('#btn-cancel-vrename').onclick = closeModal;
      overlay.addEventListener('click', (ev) => { if (ev.target === overlay) closeModal(); });
      overlay.querySelector('#btn-save-vrename').onclick = () => {
        const nextTitle = overlay.querySelector('#vr-title').value.trim() || letter.title;
        store.updateLetterVaultEntry(letter.id, {
          title: nextTitle,
          category: overlay.querySelector('#vr-cat').value.trim() || letter.category,
          tanggalSurat: overlay.querySelector('#vr-date').value.trim() || letter.tanggalSurat,
          targetPerson: overlay.querySelector('#vr-person').value.trim() || letter.targetPerson,
          perihal: overlay.querySelector('#vr-hal').value.trim() || letter.perihal
        });
        closeModal();
        renderVaultView(container);
        showToast('✓ Nama & metadata dokumen Gudang diperbarui.', 'success');
      };
    };
  });

  // 9B. Riwayat Versi Dokumen Gudang (Version History)
  container.querySelectorAll('.btn-vault-history').forEach(btn => {
    btn.onclick = () => {
      const id = btn.dataset.id;
      const letter = store.getLetterVault().find(l => l.id === id);
      if (!letter) return;
      const versions = Array.isArray(letter.versionHistory) ? letter.versionHistory : [];

      container.querySelector('#vault-version-modal')?.remove();
      const modalWrap = document.createElement('div');
      modalWrap.id = 'vault-version-modal';
      modalWrap.className = 'modal-overlay';
      modalWrap.innerHTML = `
        <div class="modal-card" style="max-width: 660px; width: 95%; max-height: 88vh; display: flex; flex-direction: column;">
          <div class="modal-header">
            <h3>${uiIcon('history', 16)}<span>Riwayat Versi: ${escapeHtml(letter.title)}</span></h3>
            <button type="button" class="btn btn-sm btn-secondary btn-square-sm" id="btn-close-v-hist" aria-label="Tutup">${uiIcon('x', 15)}</button>
          </div>
          <div class="modal-body" style="overflow-y: auto; display: grid; gap: 9px;">
            ${versions.length === 0
              ? `<div style="padding:14px; background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; font-size:13px; color:#475569;">
                   Belum ada riwayat versi sebelumnya. Setiap kali Anda menyimpan pembaruan dokumen ini di Word (Ctrl+S), versinya otomatis terekam hingga 10 revisi terakhir.
                 </div>`
              : versions.map((ver, idx) => {
                  const dtStr = ver.savedAt
                    ? new Date(ver.savedAt).toLocaleString('id-ID', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' })
                    : '-';
                  return `
                    <div style="display:flex; justify-content:space-between; align-items:center; gap:10px; padding:11px 13px; border:1px solid ${idx === 0 ? '#86efac' : '#e2e8f0'}; background:${idx === 0 ? '#f0fdf4' : '#ffffff'}; border-radius:10px;">
                      <div style="min-width:0;">
                        <div style="font-size:13px; font-weight:700; color:#0f172a; display:flex; align-items:center; gap:5px;">
                          ${idx === 0 ? `${uiIcon('checkCircle', 13)}<span>Versi Aktif Terbaru ·</span>` : `<span>Versi #${versions.length - idx} ·</span>`}<span>${escapeHtml(ver.note || 'Simpan Perubahan')}</span>
                        </div>
                        <div style="font-size:11.5px; color:#64748b; margin-top:3px; display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                          <span style="display:inline-flex;align-items:center;gap:4px;">${uiIcon('clock', 11)}${escapeHtml(dtStr)}</span>
                          <span>·</span>
                          <span style="display:inline-flex;align-items:center;gap:4px;">${uiIcon('user', 11)}${escapeHtml(ver.savedBy || 'Pengguna')}</span>
                          <span>· No: ${escapeHtml(ver.nomorSurat || '-')}</span>
                        </div>
                      </div>
                      ${idx > 0 && ver.document
                        ? `<button type="button" class="btn btn-sm btn-primary btn-restore-vault-ver" data-ver-id="${escapeHtml(ver.versionId)}">Pulihkan &amp; Buka</button>`
                        : '<span class="status-chip status-chip-green" style="font-size:10.5px;">Terkini</span>'}
                    </div>
                  `;
                }).join('')}
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary" id="btn-ok-v-hist">Tutup</button>
          </div>
        </div>
      `;
      container.appendChild(modalWrap);
      const closeHist = () => modalWrap.remove();
      modalWrap.querySelector('#btn-close-v-hist').onclick = closeHist;
      modalWrap.querySelector('#btn-ok-v-hist').onclick = closeHist;
      modalWrap.addEventListener('click', (ev) => {
        if (ev.target === modalWrap) closeHist();
      });
      modalWrap.querySelectorAll('.btn-restore-vault-ver').forEach((rBtn) => {
        rBtn.onclick = () => {
          const restored = store.restoreLetterVaultVersion(id, rBtn.dataset.verId);
          if (!restored) {
            showToast('Gagal memulihkan versi.', 'error');
            return;
          }
          closeHist();
          const pickedUnit = (selectedVaultTargetMode === 'smp' || selectedVaultTargetMode === 'smk')
            ? selectedVaultTargetMode
            : store.getActiveUnitId();
          openVaultLetterWithUnit(restored, container, pickedUnit);
        };
      });
    };
  });

  // 10. Duplikasi Surat (salinan baru tidak dikunci sebagai Master agar bebas diedit)
  container.querySelectorAll('.btn-duplicate-letter').forEach(btn => {
    btn.onclick = () => {
      const id = btn.dataset.id;
      const letter = store.getLetterVault().find(l => l.id === id);
      if (!letter) return;

      const duplicated = {
        ...JSON.parse(JSON.stringify(letter)),
        id: `letter-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        title: `(Salinan) ${letter.title}`,
        isMasterTemplate: false,
        isLockedTemplate: false,
        isAiReference: true,
        createdAt: new Date().toISOString()
      };

      store.addLetterVaultEntry(duplicated);
      showToast('Surat diduplikasi sebagai salinan baru (tidak terkunci).', 'success');
      renderVaultView(container);
    };
  });

  // 11. Hapus Surat dari Gudang (dilindungi bila berstatus ⭐ Master Terkunci)
  container.querySelectorAll('.btn-delete-vault-letter').forEach(btn => {
    btn.onclick = () => {
      const id = btn.dataset.id;
      const letter = store.getLetterVault().find(l => l.id === id);
      if (!letter) return;

      if (letter.isLockedTemplate || letter.isMasterTemplate) {
        showToast('Dokumen Master Terkunci dilindungi dari penghapusan. Klik tombol Master Terkunci untuk membuka kunci terlebih dahulu.', 'warning');
        return;
      }

      if (confirm(`Apakah Anda yakin ingin menghapus "${letter.title}" dari Gudang Surat?`)) {
        store.removeLetterVaultEntry(id);
        showToast('Surat dihapus dari gudang.', 'info');
        renderVaultView(container);
      }
    };
  });
}

/**
 * Memproses berkas .docx yang diunggah pengguna:
 * Mengekstrak paragraf & metadata, menyimpan ke Gudang Surat dengan mode preservasi format,
 * dan memuatnya ke Lembar Kerja Word.
 */
async function processUploadedDocx(file, container) {
  if (!file.name.endsWith('.docx')) {
    showToast('Harap pilih berkas dengan ekstensi .docx Microsoft Word.', 'error');
    return;
  }

  showToast(`Membaca dan memproses struktur "${file.name}"...`, 'info');

  try {
    const arrayBuffer = await file.arrayBuffer();
    const parsed = await parseDocxForEditing(arrayBuffer);

    const uploadedPayload = {
      rawBytes: new Uint8Array(arrayBuffer),
      fileName: file.name,
      parsedMeta: parsed.meta,
      structured: parsed.structured,
      paragraphs: parsed.editableParagraphs
    };
    const metaDoc = extractDocumentMetadata({
      uploadedDocx: uploadedPayload,
      schoolProfile: store.getSchoolProfile(),
      fallbackTitle: file.name.replace(/\.docx$/i, '')
    });

    // Buat entri baru di Gudang Surat dengan penamaan terstruktur otomatis
    const newEntry = store.addLetterVaultEntry({
      title: metaDoc.formattedTitle || parsed.meta.title || file.name.replace(/\.docx$/i, ''),
      nomorSurat: metaDoc.nomorSurat || parsed.meta.nomor || '-',
      tanggalSurat: metaDoc.tanggalSurat || parsed.meta.date || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
      category: metaDoc.category || 'Dokumen Impor',
      targetPerson: metaDoc.recipientName || parsed.meta.signer || '-',
      perihal: metaDoc.perihal || parsed.meta.title,
      summary: `${metaDoc.summary} (Berkas asli: "${file.name}")`,
      isAiReference: true,
      source: 'upload',
      uploadedDocx: uploadedPayload
    });

    // Otomatis atur berkas aktif di Lembar Kerja Word
    store.setActiveUploadedDocx(newEntry.uploadedDocx);

    showToast(`Berkas "${file.name}" berhasil disimpan ke Gudang Surat & siap diedit!`, 'success');
    renderVaultView(container);

    // Tanya pengguna apakah ingin langsung membukanya di Word
    if (confirm(`Berkas "${file.name}" telah tersimpan di Gudang Surat.\n\nApakah Anda ingin langsung membukanya di Lembar Kerja Word untuk mengedit?`)) {
      const canvasTab = document.querySelector('.nav-tab[data-tab="canvas"]');
      if (canvasTab) canvasTab.click();
    }
  } catch (err) {
    showToast(`Gagal memproses berkas .docx: ${err.message}`, 'error');
  }
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str || '';
  return div.innerHTML;
}
