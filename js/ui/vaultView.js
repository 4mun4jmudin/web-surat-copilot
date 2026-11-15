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

let currentSearchQuery = '';
let currentCategoryFilter = 'all';
let filterOnlyAiRef = false;
let currentVaultSort = 'newest';
let currentVaultViewMode = 'grid';

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

const V_SVG_LAYERS_30 = `<svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="#156c3b" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 12 12 17 22 12"/><polyline points="2 17 12 22 22 17"/></svg>`;
const V_SVG_DOC = (size = 18) => `<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>`;
const V_SVG_STAR = (size = 18) => `<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>`;
const V_SVG_BOLT = (size = 18) => `<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>`;
const V_SVG_SEARCH_15 = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>`;
const V_SVG_UPLOAD_15 = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>`;
const V_SVG_CAL_12 = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>`;
const V_SVG_USER_12 = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>`;
const V_SVG_DL_13 = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>`;
const V_SVG_COPY_13 = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>`;
const V_SVG_TRASH_13 = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>`;

function getCategoryPillIcon(cat) {
  if (cat === 'all') return `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>`;
  if (cat === 'Surat Tugas') return `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/></svg>`;
  if (cat === 'Surat Undangan') return `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="4" width="20" height="16" rx="2"/><polyline points="22,6 12,13 2,6"/></svg>`;
  if (cat === 'Dokumen Impor') return V_SVG_DL_13;
  return V_SVG_DOC(13);
}

export function renderVaultView(container) {
  if (!container) return;

  const allLetters = store.getLetterVault();
  const aiRefLetters = allLetters.filter(l => l.isAiReference);
  const categories = ['all', 'Surat Keterangan', 'Surat Tugas', 'Surat Undangan', 'Dokumen Impor'];

  container.innerHTML = `
    <div id="dropzone-docx-vault" class="view-atmosphere" style="transition: background 0.2s ease;">
      <div class="view-shell vault-view-container">

        <!-- HEADER GUDANG SURAT -->
        <div class="view-page-header" style="margin-bottom: 12px;">
          <div class="view-header-lead">
            <span style="margin-top: 2px;">${V_SVG_LAYERS_30}</span>
            <span class="view-header-divider"></span>
            <div>
              <h2 class="view-title">
                <span>Gudang Template &amp; Draf Surat Sekolah</span>
                <span class="status-chip status-chip-green">✏️ Dapat Diedit &amp; Disesuaikan Kembali</span>
              </h2>
              <p class="view-subtitle">
                Kumpulan template, draf, dan arsip surat sekolah yang dapat digunakan kembali atau diedit ulang di <strong>Lembar Kerja</strong>, serta menjadi referensi penulisan untuk AI <strong>Copilot (Surat Keluar)</strong>.
              </p>
            </div>
          </div>
        </div>

        <!-- ACTION BAR -->
        <div style="display: flex; gap: 10px; flex-wrap: wrap; align-items: center; margin-bottom: 18px;">
          <input type="file" id="input-upload-docx-vault" accept=".docx" class="hidden">
          <button id="btn-upload-docx-vault" class="btn btn-secondary" style="padding: 8px 16px; font-weight: 600; border-radius: 8px;">
            ${V_SVG_UPLOAD_15}
            <span>Unggah Berkas .docx / .doc</span>
          </button>
          <button id="btn-create-new-from-vault" class="btn btn-primary" style="padding: 8px 18px; font-weight: 700; border-radius: 8px;">
            + Buat Surat Baru dari Word
          </button>
          <span style="font-size: 11.5px; color: #64748b; margin-left: 4px;">(Atau seret &amp; lepas berkas Word <code>.docx</code> langsung ke halaman ini)</span>
        </div>

        <!-- 3 KARTU STATISTIK KPI -->
        <div class="kpi-grid-3">
          <div class="kpi-card tint-green">
            <div class="kpi-icon-box kpi-icon-green">${V_SVG_DOC(20)}</div>
            <div>
              <div style="font-size: 12.5px; font-weight: 600; color: #334155;">Total Surat Tersimpan</div>
              <div style="font-size: 26px; font-weight: 800; color: #0f172a; line-height: 1.2; margin-top: 4px;">
                ${allLetters.length} <span style="font-size: 13px; font-weight: 500; color: #475569;">dokumen</span>
              </div>
            </div>
            <div class="kpi-watermark" style="color: #15803d;">${V_SVG_DOC(54)}</div>
          </div>

          <div class="kpi-card tint-amber">
            <div class="kpi-icon-box kpi-icon-amber">${V_SVG_STAR(20)}</div>
            <div>
              <div style="font-size: 12.5px; font-weight: 600; color: #334155;">Acuan AI Copilot Aktif</div>
              <div style="font-size: 26px; font-weight: 800; color: #b45309; line-height: 1.2; margin-top: 4px;">
                ${aiRefLetters.length} <span style="font-size: 13px; font-weight: 500; color: #475569;">surat dijadikan referensi</span>
              </div>
            </div>
            <div class="kpi-watermark" style="color: #15803d;">${V_SVG_STAR(56)}</div>
          </div>

          <div class="kpi-card tint-blue">
            <div class="kpi-icon-box kpi-icon-blue">${V_SVG_BOLT(20)}</div>
            <div>
              <div style="font-size: 13px; font-weight: 700; color: #0f172a;">Pencarian Cerdas (Anti-Typo)</div>
              <div style="font-size: 14.5px; font-weight: 800; color: #1e3a8a; margin-top: 3px;">Fuzzy Match &amp; Ranking Nama</div>
              <div style="font-size: 11.5px; color: #475569; margin-top: 2px;">Menemukan surat meski terdapat kesalahan penulisan.</div>
            </div>
          </div>
        </div>

        <!-- SEARCH BAR + KATEGORI DROPDOWN + SORT DROPDOWN -->
        <div class="filter-bar-card">
          <div style="position: relative; flex: 1; min-width: 260px;">
            <span class="icon-input-leading" style="color: #64748b;">${V_SVG_SEARCH_15}</span>
            <input type="text" id="input-vault-search" class="icon-input-control" style="padding-right: 32px;"
              placeholder="Ketik nama surat, perihal, nomor, atau kata kunci lainnya..." autocomplete="off" spellcheck="false" value="${escapeHtml(currentSearchQuery)}">
            <button type="button" id="btn-clear-vault-search" class="${currentSearchQuery ? '' : 'hidden'}" title="Bersihkan pencarian"
              style="position: absolute; right: 8px; top: 50%; transform: translateY(-50%); border: none; background: #e2e8f0; color: #475569; width: 20px; height: 20px; border-radius: 50%; font-size: 11px; line-height: 1; cursor: pointer; display: inline-flex; align-items: center; justify-content: center;">✕</button>
          </div>

          <div class="icon-input-wrap" style="width: 195px;">
            <span class="icon-input-leading">${getCategoryPillIcon('all')}</span>
            <select id="select-vault-category" class="icon-input-control" style="cursor: pointer;">
              ${categories.map(cat => `
                <option value="${cat}" ${currentCategoryFilter === cat ? 'selected' : ''}>${cat === 'all' ? 'Semua Kategori' : cat}</option>
              `).join('')}
            </select>
          </div>

          <div class="icon-input-wrap" style="width: 185px;">
            <span class="icon-input-leading">⇅</span>
            <select id="select-vault-sort" class="icon-input-control" style="cursor: pointer;">
              <option value="newest" ${currentVaultSort === 'newest' ? 'selected' : ''}>Urut: Terbaru</option>
              <option value="oldest" ${currentVaultSort === 'oldest' ? 'selected' : ''}>Urut: Terlama</option>
              <option value="title_asc" ${currentVaultSort === 'title_asc' ? 'selected' : ''}>Urut: Judul (A - Z)</option>
              <option value="nomor_asc" ${currentVaultSort === 'nomor_asc' ? 'selected' : ''}>Urut: Nomor Surat</option>
            </select>
          </div>
        </div>

        <!-- PILL KATEGORI + TOGGLE GRID/LIST -->
        <div class="filter-pills-bar">
          <div style="display: flex; gap: 8px; flex-wrap: wrap; align-items: center;">
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

            <button type="button" class="pill-tab ${filterOnlyAiRef ? 'pill-amber-active' : ''} btn-toggle-ai-filter">
              <span>✨</span>
              <span>${filterOnlyAiRef ? 'Acuan AI Aktif' : 'Hanya Acuan AI'}</span>
            </button>
          </div>

          <div class="vault-view-toggle-group">
            <button type="button" class="vault-view-toggle-btn btn-vault-view-mode ${currentVaultViewMode === 'grid' ? 'active' : ''}" data-mode="grid" title="Tampilan Kartu Grid">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>
            </button>
            <button type="button" class="vault-view-toggle-btn btn-vault-view-mode ${currentVaultViewMode === 'list' ? 'active' : ''}" data-mode="list" title="Tampilan Daftar Tabel">
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
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

function buildVaultCardsHtml(allLetters) {
  const baseFiltered = allLetters.filter(item => {
    if (filterOnlyAiRef && !item.isAiReference) return false;
    if (currentCategoryFilter !== 'all' && item.category !== currentCategoryFilter) return false;
    return true;
  });

  const hasActiveQuery = currentSearchQuery.trim().length > 0;
  let ranked = searchAndRankLetters(baseFiltered, currentSearchQuery);

  if (!hasActiveQuery) {
    ranked = [...ranked].sort((a, b) => {
      const la = a.item;
      const lb = b.item;
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
        <span style="font-size: 38px; display: block; margin-bottom: 10px;">📭</span>
        <h4 style="color: #0f172a; font-size: 16px; font-weight: 800; margin-bottom: 6px;">Tidak Ada Surat yang Cocok</h4>
        <p style="color: #64748b; font-size: 13px; max-width: 460px; margin: 0 auto 16px auto;">
          Belum ada surat pada kategori ini yang mendekati kata kunci <b>"${escapeHtml(currentSearchQuery.trim())}"</b>.
        </p>
        <button id="btn-empty-reset" class="btn btn-secondary btn-sm">Reset Filter Pencarian</button>
      </div>
    `;
  }

  const queryHeaderHtml = hasActiveQuery ? `
    <div style="margin-bottom: 12px; font-size: 12.5px; color: #475569; display: flex; align-items: center; justify-content: space-between;">
      <span>Menampilkan <strong>${ranked.length}</strong> surat yang paling mendekati pencarian <b>"${escapeHtml(currentSearchQuery.trim())}"</b>:</span>
    </div>
  ` : '';

  if (currentVaultViewMode === 'list') {
    return `
      ${queryHeaderHtml}
      <div class="ui-card" style="overflow: hidden;">
        <div style="overflow-x: auto;">
          <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
            <thead>
              <tr style="background: #f8fafc; border-bottom: 1px solid #e2e8f0; text-align: left; color: #475569; font-size: 12px;">
                <th style="padding: 11px 14px;">Kategori &amp; Judul Surat</th>
                <th style="padding: 11px 14px;">Nomor Surat</th>
                <th style="padding: 11px 14px;">Tanggal / Pihak Terkait</th>
                <th style="padding: 11px 14px; text-align: center;">Acuan AI</th>
                <th style="padding: 11px 14px; text-align: right;">Tindakan</th>
              </tr>
            </thead>
            <tbody>
              ${ranked.map(({ item: letter }) => `
                <tr style="border-bottom: 1px solid #f1f5f9;">
                  <td style="padding: 11px 14px;">
                    <span class="status-chip ${letter.category === 'Dokumen Impor' ? 'status-chip-blue' : 'status-chip-green'}" style="font-size: 10px; margin-bottom: 4px;">
                      ${escapeHtml(letter.category || 'Surat Dinas')}
                    </span>
                    <div style="font-weight: 700; color: #0f172a;">${escapeHtml(letter.title)}</div>
                  </td>
                  <td style="padding: 11px 14px; font-weight: 600; color: #334155; font-size: 12px;">${escapeHtml(letter.nomorSurat || '-')}</td>
                  <td style="padding: 11px 14px; font-size: 12px; color: #475569;">
                    <div>${escapeHtml(letter.tanggalSurat || '-')}</div>
                    <div style="color: #64748b;">${escapeHtml(letter.targetPerson || '-')}</div>
                  </td>
                  <td style="padding: 11px 14px; text-align: center;">
                    <button class="btn-toggle-ai-ref" data-id="${letter.id}"
                      style="background: ${letter.isAiReference ? '#fef3c7' : '#f1f5f9'}; border: 1px solid ${letter.isAiReference ? '#f59e0b' : '#cbd5e1'}; color: ${letter.isAiReference ? '#b45309' : '#64748b'}; font-size: 11px; font-weight: 700; padding: 3px 9px; border-radius: 999px; cursor: pointer;">
                      ${letter.isAiReference ? '★ Acuan AI' : '☆ Jadikan'}
                    </button>
                  </td>
                  <td style="padding: 11px 14px; text-align: right; white-space: nowrap;">
                    <button class="btn btn-sm btn-primary btn-load-letter" data-id="${letter.id}">Buka di Word</button>
                    <button class="btn btn-sm btn-secondary btn-download-vault-letter" data-id="${letter.id}">.docx</button>
                    <button class="btn btn-sm btn-secondary btn-duplicate-letter" data-id="${letter.id}" title="Duplikasi">${V_SVG_COPY_13}</button>
                    <button class="btn btn-sm btn-secondary btn-delete-vault-letter" data-id="${letter.id}" style="color: #dc2626;" title="Hapus">${V_SVG_TRASH_13}</button>
                  </td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  return `
    ${queryHeaderHtml}
    <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); gap: 16px; margin-bottom: 24px;">
      ${ranked.map(({ item: letter, isFuzzy }, idx) => {
        const isRef = letter.isAiReference;
        const cat = letter.category || 'Surat Dinas';
        const isImport = cat === 'Dokumen Impor';
        const isUndangan = cat === 'Surat Undangan';
        const chipClass = isImport ? 'status-chip-blue' : isUndangan ? 'status-chip-purple' : 'status-chip-green';
        const thumbClass = isImport ? 'thumb-blue' : isUndangan ? 'thumb-purple' : 'thumb-green';
        const isTopMatch = hasActiveQuery && idx === 0;

        return `
          <div class="card vault-card vault-card-modern ${isTopMatch ? 'is-top-match' : ''}">
            <div>
              <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; gap: 6px; flex-wrap: wrap;">
                <div style="display: flex; align-items: center; gap: 6px; flex-wrap: wrap;">
                  <span class="status-chip ${chipClass}" style="font-size: 11px;">
                    ${escapeHtml(cat)}
                  </span>
                  ${isTopMatch ? `<span class="status-chip status-chip-green" style="font-size: 10px;">🎯 Paling Dekat</span>` : (hasActiveQuery && isFuzzy ? `<span class="status-chip status-chip-blue" style="font-size: 10px;">✓ Toleransi Typo</span>` : '')}
                </div>
                <button class="btn-toggle-ai-ref" data-id="${letter.id}" title="${isRef ? 'Surat ini dijadikan acuan AI. Klik untuk nonaktifkan.' : 'Jadikan surat ini sebagai acuan gaya penulisan AI Copilot.'}" 
                  style="background: ${isRef ? '#fef3c7' : '#f8fafc'}; border: 1px solid ${isRef ? '#fcd34d' : '#cbd5e1'}; color: ${isRef ? '#b45309' : '#64748b'}; font-size: 11px; font-weight: 700; padding: 3px 10px; border-radius: 999px; cursor: pointer;">
                  ${isRef ? '★ Acuan AI' : '☆ Jadikan Acuan'}
                </button>
              </div>

              <div style="display: flex; align-items: flex-start; gap: 12px; margin-bottom: 11px;">
                <div class="vault-doc-thumb ${thumbClass}">
                  ${V_SVG_DOC(22)}
                </div>
                <div style="flex: 1; min-width: 0;">
                  <strong style="font-size: 14px; font-weight: 800; color: #0f172a; display: block; line-height: 1.3; margin-bottom: 3px;">
                    ${escapeHtml(letter.title)}
                  </strong>
                  <div style="font-size: 11.5px; font-weight: 600; color: #334155; margin-bottom: 5px; word-break: break-word;">
                    Nomor: ${escapeHtml(letter.nomorSurat || '-')}
                  </div>
                  <div style="font-size: 11.5px; color: #64748b; display: flex; align-items: center; gap: 10px; flex-wrap: wrap;">
                    <span style="display: inline-flex; align-items: center; gap: 4px;">${V_SVG_CAL_12} ${escapeHtml(letter.tanggalSurat || '-')}</span>
                    ${letter.targetPerson && letter.targetPerson !== '-' ? `<span style="display: inline-flex; align-items: center; gap: 4px;">${V_SVG_USER_12} ${escapeHtml(letter.targetPerson)}</span>` : ''}
                  </div>
                </div>
              </div>

              <p class="vault-summary-box">
                ${escapeHtml(letter.summary || 'Dokumen persuratan resmi sekolah.')}
              </p>
            </div>

            <div style="padding-top: 4px; display: flex; gap: 7px; align-items: center;">
              <button class="btn btn-sm btn-primary btn-load-letter" data-id="${letter.id}" style="flex: 1; height: 34px; font-weight: 600; border-radius: 8px;">
                <span style="background: rgba(255,255,255,0.2); padding: 1px 4px; border-radius: 3px; font-size: 10px; font-weight: 800;">W</span>
                <span>Buka di Word</span>
              </button>
              <button class="btn btn-sm btn-secondary btn-download-vault-letter" data-id="${letter.id}" title="Unduh berkas .docx" style="height: 34px; padding: 0 10px; border-radius: 8px; font-weight: 600;">
                ${V_SVG_DL_13}
                <span>.docx</span>
              </button>
              <button class="btn btn-sm btn-secondary btn-duplicate-letter" data-id="${letter.id}" title="Duplikasi Surat" style="height: 34px; width: 34px; padding: 0; border-radius: 8px;">
                ${V_SVG_COPY_13}
              </button>
              <button class="btn btn-sm btn-secondary btn-delete-vault-letter" data-id="${letter.id}" title="Hapus dari Gudang" style="height: 34px; width: 34px; padding: 0; border-radius: 8px; color: #dc2626;">
                ${V_SVG_TRASH_13}
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

  // 3. Toggle only AI filter
  const btnToggleAiFilter = container.querySelector('.btn-toggle-ai-filter');
  if (btnToggleAiFilter) {
    btnToggleAiFilter.onclick = () => {
      filterOnlyAiRef = !filterOnlyAiRef;
      renderVaultView(container);
    };
  }

  // 5. Tombol Buat Surat Baru -> Switch ke tab canvas
  const btnCreate = container.querySelector('#btn-create-new-from-vault');
  if (btnCreate) {
    btnCreate.onclick = () => {
      const canvasTab = document.querySelector('.nav-tab[data-tab="canvas"]');
      if (canvasTab) canvasTab.click();
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
      renderVaultView(container);
      const input = container.querySelector('#input-vault-search');
      if (input) input.focus();
    };
  }

  // 7. Toggle AI Reference per card
  container.querySelectorAll('.btn-toggle-ai-ref').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      const id = btn.dataset.id;
      store.toggleLetterAiReference(id);
      showToast('Status acuan referensi AI diperbarui.', 'info');
      renderVaultView(container);
    };
  });

  // 8. Buka di Word Canvas
  container.querySelectorAll('.btn-load-letter').forEach(btn => {
    btn.onclick = () => {
      const id = btn.dataset.id;
      const success = store.loadLetterToCanvas(id);
      if (success) {
        showToast('Dokumen berhasil dimuat ke Lembar Kerja Word.', 'success');
        const canvasTab = document.querySelector('.nav-tab[data-tab="canvas"]');
        if (canvasTab) canvasTab.click();
      } else {
        showToast('Gagal memuat dokumen.', 'error');
      }
    };
  });

  // 9. Unduh dokumen langsung dari kartu
  container.querySelectorAll('.btn-download-vault-letter').forEach(btn => {
    btn.onclick = async () => {
      const id = btn.dataset.id;
      const letter = store.getLetterVault().find(l => l.id === id);
      if (!letter) return;

      showToast(`Menyiapkan berkas .docx untuk ${letter.title}...`, 'info');
      try {
        let blob;
        if (letter.uploadedDocx?.rawBytes) {
          // Berkas yang diupload: Gunakan rawBytes atau patch
          blob = new Blob([letter.uploadedDocx.rawBytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        } else if (letter.document) {
          blob = await renderDocx(letter.document, 'blob');
        } else {
          blob = await renderDocx(store.getState().document, 'blob');
        }

        const safeTitle = (letter.title || 'surat_dinas').replace(/[^a-zA-Z0-9_\-]/g, '_').slice(0, 45);
        downloadFile(blob, `${safeTitle}.docx`);
        showToast('Berkas .docx berhasil diunduh.', 'success');
      } catch (err) {
        showToast(`Gagal mengunduh: ${err.message}`, 'error');
      }
    };
  });

  // 10. Duplikasi Surat
  container.querySelectorAll('.btn-duplicate-letter').forEach(btn => {
    btn.onclick = () => {
      const id = btn.dataset.id;
      const letter = store.getLetterVault().find(l => l.id === id);
      if (!letter) return;

      const duplicated = {
        ...JSON.parse(JSON.stringify(letter)),
        id: `letter-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
        title: `(Salinan) ${letter.title}`,
        createdAt: new Date().toISOString()
      };

      store.addLetterVaultEntry(duplicated);
      showToast('Surat berhasil diduplikasi.', 'success');
      renderVaultView(container);
    };
  });

  // 11. Hapus Surat dari Gudang
  container.querySelectorAll('.btn-delete-vault-letter').forEach(btn => {
    btn.onclick = () => {
      const id = btn.dataset.id;
      const letter = store.getLetterVault().find(l => l.id === id);
      if (!letter) return;

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

    // Buat entri baru di Gudang Surat
    const newEntry = store.addLetterVaultEntry({
      title: parsed.meta.title || file.name.replace(/\.docx$/i, ''),
      nomorSurat: parsed.meta.nomor || '-',
      tanggalSurat: parsed.meta.date || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
      category: 'Dokumen Impor',
      targetPerson: parsed.meta.signer || '-',
      perihal: parsed.meta.title,
      summary: `Dokumen .docx resmi diunggah: "${file.name}". Berisi ${parsed.editableParagraphs.length} paragraf dengan preservasi format asli 100%.`,
      isAiReference: true,
      source: 'upload',
      uploadedDocx: {
        rawBytes: new Uint8Array(arrayBuffer),
        fileName: file.name,
        parsedMeta: parsed.meta,
        structured: parsed.structured,
        paragraphs: parsed.editableParagraphs
      }
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
