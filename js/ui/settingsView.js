/**
 * settingsView.js - Gudang API Multi-Provider & Pengaturan Sistem
 * Mengelola banyak kunci API (Google Gemini, Groq, DeepSeek, Qwen, OpenAI-Compatible)
 * dengan fitur Uji Koneksi, Aktivasi Cepat, dan Auto-Fallback.
 */

import { store, isUsingBuiltinSystemKey, getBuiltinSystemKey } from '../state.js';
import { executeAiRequest, detectActiveGeminiModel, detectActiveGroqModel, LATEST_GEMINI_DEFAULT_MODEL, LATEST_GROQ_DEFAULT_MODEL } from '../ai/client.js';
import { canManageApiVault } from '../auth/permissions.js';
import { showToast } from './notification.js';

function renderProviderBrand(provider) {
  const p = (provider || 'openai').toLowerCase();
  if (p === 'gemini') {
    return {
      name: 'Google Gemini',
      logoHtml: `
        <svg width="24" height="24" viewBox="0 0 24 24" aria-hidden="true">
          <path fill="#4285F4" d="M23.49 12.27c0-.79-.07-1.54-.19-2.27H12v4.51h6.47a5.54 5.54 0 0 1-2.4 3.63v3.01h3.88c2.27-2.09 3.54-5.17 3.54-8.88Z"/>
          <path fill="#34A853" d="M12 24c3.24 0 5.95-1.08 7.94-2.91l-3.88-3.01c-1.08.72-2.45 1.16-4.06 1.16-3.12 0-5.77-2.11-6.72-4.96H1.27v3.1A12 12 0 0 0 12 24Z"/>
          <path fill="#FBBC05" d="M5.28 14.28A7.2 7.2 0 0 1 4.9 12c0-.79.14-1.56.38-2.28V6.62H1.27A12 12 0 0 0 0 12c0 1.94.46 3.77 1.27 5.38l4.01-3.1Z"/>
          <path fill="#EA4335" d="M12 4.76c1.76 0 3.34.61 4.58 1.8l3.44-3.44C17.95 1.19 15.24 0 12 0A12 12 0 0 0 1.27 6.62l4.01 3.1C6.23 6.87 8.88 4.76 12 4.76Z"/>
        </svg>
      `,
      bg: '#ffffff'
    };
  }
  if (p === 'groq') {
    return {
      name: 'Groq',
      logoHtml: `<span style="font-weight: 800; font-size: 11px; color: #ffffff; letter-spacing: -0.02em; font-family: system-ui, sans-serif;">groq</span>`,
      bg: 'linear-gradient(135deg, #fb7122 0%, #ea580c 100%)'
    };
  }
  if (p === 'deepseek') {
    return {
      name: 'DeepSeek',
      logoHtml: `
        <svg width="24" height="24" viewBox="0 0 24 24" fill="#2563eb" aria-hidden="true">
          <path d="M19.8 11.4c-.9-.4-1.8-.3-2.6.2-.5-2.6-2.8-4.7-5.9-5.1.3-.9.9-1.7 1.7-2.2-.3-.2-.8-.3-1.4-.1-1 .3-1.8 1.1-2.1 2.1-.6 0-1.3.1-1.9.3-2.9 1-4.8 3.8-4.3 6.9.5 3.2 3.4 5.4 6.8 5.4 4.1 0 7.1-2.5 8.5-5.7.9-.2 1.8-.7 2.2-1.4-.2-.2-.6-.3-1-.4ZM9.7 13.3c-.6 0-1-.4-1-1s.4-1 1-1 1 .4 1 1-.4 1-1 1Z"/>
        </svg>
      `,
      bg: '#eff6ff'
    };
  }
  if (p === 'qwen') {
    return {
      name: 'Qwen',
      logoHtml: `
        <svg width="23" height="23" viewBox="0 0 24 24" fill="none" stroke="#6d28d9" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
          <polygon points="12 2 20.5 7 20.5 17 12 22 3.5 17 3.5 7 12 2" fill="#ede9fe"/>
          <path d="M12 7v10M7.5 9.5l9 5M16.5 9.5l-9 5"/>
        </svg>
      `,
      bg: '#f5f3ff'
    };
  }
  return {
    name: 'OpenAI / Custom',
    logoHtml: `
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#156c3b" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M12 2v4m0 12v4M4.93 4.93l2.83 2.83m8.48 8.48 2.83 2.83M2 12h4m12 0h4M4.93 19.07l2.83-2.83m8.48-8.48 2.83-2.83"/>
      </svg>
    `,
    bg: '#ecfdf5'
  };
}

const SVG_CHAIN_22 = `<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.1" stroke-linecap="round" stroke-linejoin="round"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg>`;
const SVG_CUBE_14 = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="m21 16-9 5-9-5V8l9-5 9 5v8z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/></svg>`;
const SVG_LINK_14 = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg>`;
const SVG_KEY_14 = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><circle cx="7.5" cy="15.5" r="4.5"/><path d="m21 3-9.5 9.5"/><path d="m15.5 7.5 3 3M18 5l2 2"/></svg>`;
const SVG_SIGNAL_14 = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#156c3b" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 20v-4"/><path d="M9 20v-8"/><path d="M14 20v-12"/><path d="M19 20V4"/></svg>`;
const SVG_EDIT_14 = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20h9"/><path d="M16.5 3.5a2.121 2.121 0 1 1 3 3L7 19l-4 1 1-4Z"/></svg>`;
const SVG_LOCK_11 = `<svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>`;

export function renderSettingsView(container) {
  if (!container) return;

  const state = store.getState();
  const vault = store.getApiVault();
  const activeId = state.activeApiId;
  const allowEditVault = canManageApiVault(state.authUser, state.authPermissions);

  container.innerHTML = `
    <div class="view-atmosphere">
      <div class="view-shell-narrow settings-view-container">

        <!-- HEADER GUDANG API -->
        <div class="view-page-header" style="margin-bottom: 14px;">
          <div class="view-header-lead">
            <div class="view-header-badge">${SVG_CHAIN_22}</div>
            <div>
              <h2 class="view-title">
                <span>Gudang API Multi-Provider</span>
                <span class="status-chip status-chip-green" style="font-size:11px;">☁️ Tersinkron (${vault.length} Koneksi AI)</span>
              </h2>
              <p class="view-subtitle">
                Kelola dan hubungkan berbagai layanan AI (multi-provider) tanpa batas jumlah koneksi. Kunci yang Anda tambahkan otomatis disinkronkan lintas perangkat untuk akun sekolah.
              </p>
            </div>
          </div>
        </div>

        ${!allowEditVault ? `
          <div style="background:#fef3c7;border:1px solid #fde68a;color:#92400e;border-radius:10px;padding:10px 14px;font-size:12.5px;margin-bottom:14px;">
            🔒 Pengelolaan &amp; penambahan kunci API AI untuk peran Tata Usaha sedang dibatasi oleh Administrator. Anda tetap dapat menggunakan kunci aktif untuk AI Copilot.
          </div>
        ` : ''}

        <div style="margin-bottom: 18px;">
          <button id="btn-add-vault-entry" class="btn btn-primary" style="padding: 9px 18px; font-weight: 700; font-size: 13px; border-radius: 8px;" ${allowEditVault ? '' : 'disabled'}>
            + Tambah Koneksi API Baru
          </button>
        </div>

        <!-- DAFTAR KONEKSI API DI GUDANG -->
        <div class="vault-list-container" style="display: flex; flex-direction: column; gap: 14px;">
          ${vault.map(entry => {
            const isActive = entry.id === activeId || entry.isActive;
            const usingSystemKey = isUsingBuiltinSystemKey(entry);
            const hasKey = !!(entry.apiKey && entry.apiKey.trim().length > 0);
            const brand = renderProviderBrand(entry.provider);

            let topBadgeHtml = '';
            if (isActive) {
              topBadgeHtml = `<span class="status-chip status-chip-green">✓ Terhubung &amp; Aktif</span>`;
            } else if (hasKey) {
              topBadgeHtml = `<span class="status-chip status-chip-blue">✓ Aktif (Siap Cadangan)</span>`;
            } else {
              topBadgeHtml = `<span class="status-chip status-chip-amber">🕒 Tidak Aktif / Belum Diuji</span>`;
            }

            const keyStatusPill = usingSystemKey
              ? `<span class="status-chip status-chip-green" style="padding: 1px 9px; font-size: 11px;">${SVG_LOCK_11} Bawaan Sistem (Siap Pakai)</span>`
              : hasKey
                ? `<span class="status-chip status-chip-green" style="padding: 1px 9px; font-size: 11px;">${SVG_LOCK_11} Terenkripsi (Aman)</span>
                   <span style="font-size: 11px; color: #64748b; margin-left: 4px;">(${escapeHtml(entry.apiKey.slice(0, 5))}…${escapeHtml(entry.apiKey.slice(-4))})</span>`
                : `<span class="status-chip status-chip-red" style="padding: 1px 9px; font-size: 11px;">${SVG_LOCK_11} Belum Dikonfigurasi</span>`;

            return `
              <div class="vault-card api-provider-card ${isActive ? 'is-active vault-card-active' : ''}">
                <div class="api-brand-col">
                  <div class="api-brand-logo" style="background: ${brand.bg}; ${entry.provider === 'groq' ? 'border: none;' : ''}">
                    ${brand.logoHtml}
                  </div>
                  <div style="font-size: 13px; font-weight: 700; color: #0f172a; line-height: 1.25;">
                    ${escapeHtml(brand.name)}
                  </div>
                </div>

                <div style="flex: 1; min-width: 240px;">
                  <div style="display: flex; align-items: center; gap: 10px; flex-wrap: wrap; margin-bottom: 7px;">
                    <strong style="font-size: 15.5px; color: #0f172a; font-weight: 700;">${escapeHtml(entry.label)}</strong>
                    ${topBadgeHtml}
                  </div>

                  <div style="font-size: 12.5px; color: #475569; display: flex; flex-direction: column; gap: 4px;">
                    <div style="display: flex; align-items: center; gap: 7px; flex-wrap: wrap;">
                      ${SVG_CUBE_14}
                      <span style="color: #64748b;">Model:</span>
                      <span style="color: #0f172a; font-weight: 500;">${escapeHtml(entry.model || '-')}</span>
                    </div>
                    <div style="display: flex; align-items: center; gap: 7px; flex-wrap: wrap;">
                      ${SVG_LINK_14}
                      <span style="color: #64748b;">Endpoint:</span>
                      <span style="color: #334155; word-break: break-all;">${escapeHtml(entry.baseUrl || '-')}</span>
                    </div>
                    <div style="display: flex; align-items: center; gap: 7px; flex-wrap: wrap;">
                      ${SVG_KEY_14}
                      <span style="color: #64748b;">Status Kunci:</span>
                      <span style="display: inline-flex; align-items: center; flex-wrap: wrap;">${keyStatusPill}</span>
                    </div>
                  </div>
                </div>

                <div style="display: flex; align-items: center; gap: 8px; flex-shrink: 0;">
                  <button type="button" class="btn btn-sm btn-secondary btn-test-vault" data-id="${entry.id}" style="padding: 7px 12px; font-weight: 600; border-radius: 8px;">
                    ${SVG_SIGNAL_14}
                    <span>Uji Koneksi</span>
                  </button>
                  ${allowEditVault ? `
                  <button type="button" class="btn btn-sm btn-secondary btn-edit-vault" data-id="${entry.id}" style="padding: 7px 12px; font-weight: 600; border-radius: 8px;">
                    ${SVG_EDIT_14}
                    <span>Edit</span>
                  </button>
                  <div class="api-kebab-wrap">
                    <button type="button" class="btn btn-sm btn-secondary btn-api-kebab" data-id="${entry.id}" title="Menu Tindakan"
                      style="width: 34px; height: 34px; padding: 0; border-radius: 8px; font-size: 16px; font-weight: 700;">⋮</button>
                    <div class="api-kebab-menu hidden" data-menu-id="${entry.id}">
                      ${!isActive ? `
                        <button type="button" class="api-kebab-item btn-activate-vault" data-id="${entry.id}">
                          <span style="color: #156c3b;">✓</span> Aktifkan Sebagai Utama
                        </button>
                      ` : `
                        <div style="padding: 6px 10px; font-size: 11.5px; color: #15803d; font-weight: 700;">
                          ✓ Sedang Aktif Utama
                        </div>
                      `}
                      <button type="button" class="api-kebab-item btn-kebab-edit" data-id="${entry.id}">
                        <span>✏️</span> Edit Kunci &amp; Model
                      </button>
                      ${!entry.isBuiltin ? `
                        <button type="button" class="api-kebab-item danger btn-delete-vault" data-id="${entry.id}">
                          <span>🗑️</span> Hapus Koneksi API
                        </button>
                      ` : ''}
                    </div>
                  </div>
                  ` : ''}
                </div>
              </div>
            `;
          }).join('')}
        </div>

        <!-- MODAL TAMBAH / EDIT KUNCI API -->
        <div id="modal-vault-form" class="modal-overlay hidden">
          <div class="modal-card" style="width: 100%; max-width: 520px;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; padding-bottom: 10px; border-bottom: 1px solid #e2e8f0;">
              <h3 id="modal-vault-title" style="font-size: 16px; font-weight: 700; color: #0f172a;">Tambah Koneksi API Baru</h3>
              <button id="btn-close-modal-vault" class="btn-icon" style="font-size: 20px; color: #64748b;">&times;</button>
            </div>

            <form id="form-vault-entry">
              <input type="hidden" id="input-vault-id" value="">

              <div class="form-group">
                <label>Pilih Penyedia AI (Provider):</label>
                <select id="select-vault-provider" class="form-control">
                  <option value="gemini">Google Gemini (Gemini 3.5 / 2.5 Flash-Lite)</option>
                  <option value="groq">Groq (GPT-OSS 120B - Ultra Cepat)</option>
                  <option value="deepseek">DeepSeek (DeepSeek V3 / R1)</option>
                  <option value="qwen">Qwen / Alibaba (via OpenRouter)</option>
                  <option value="openai">OpenAI / Custom Compatible</option>
                </select>
              </div>

              <div class="form-group">
                <label>Nama Label Kunci:</label>
                <input type="text" id="input-vault-label" class="form-control" placeholder="Contoh: Google Gemini (gemini-2.5-flash-lite)" required>
              </div>

              <div class="form-group">
                <label>Model ID:</label>
                <input type="text" id="input-vault-model" class="form-control" value="gemini-3.5-flash-lite" required>
                <small style="color: #64748b; font-size: 11.5px;" id="vault-model-hint">Pilih model yang didukung oleh penyedia AI.</small>
              </div>

              <div class="form-group">
                <label>Base URL Endpoint:</label>
                <input type="text" id="input-vault-baseurl" class="form-control" value="https://generativelanguage.googleapis.com/v1beta" required>
              </div>

              <div class="form-group">
                <label>API Key / Token Rahasia:</label>
                <input type="password" id="input-vault-apikey" class="form-control" placeholder="Masukkan API Key Anda..." autocomplete="off">
                <small style="color: #64748b; font-size: 11.5px;">Kunci disimpan terenkripsi hanya di perangkat browser lokal Anda (IndexedDB).</small>
              </div>

              <div class="form-group">
                <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
                  <input type="checkbox" id="check-vault-active" checked>
                  <span>Jadikan Koneksi Aktif untuk Pembuatan Surat</span>
                </label>
              </div>

              <div style="display: flex; justify-content: flex-end; gap: 8px; margin-top: 18px; padding-top: 12px; border-top: 1px solid #e2e8f0;">
                <button type="button" id="btn-cancel-modal-vault" class="btn btn-secondary">Batal</button>
                <button type="submit" class="btn btn-primary" style="font-weight: 600;">Simpan Koneksi</button>
              </div>
            </form>
          </div>
        </div>

      </div>
    </div>
  `;

  bindSettingsEvents(container);
}

function bindSettingsEvents(container) {
  const modal = container.querySelector('#modal-vault-form');
  const form = container.querySelector('#form-vault-entry');
  const providerSelect = container.querySelector('#select-vault-provider');
  const modelInput = container.querySelector('#input-vault-model');
  const baseurlInput = container.querySelector('#input-vault-baseurl');
  const labelInput = container.querySelector('#input-vault-label');
  const apiKeyInput = container.querySelector('#input-vault-apikey');
  const activeCheck = container.querySelector('#check-vault-active');
  const idInput = container.querySelector('#input-vault-id');
  const modalTitle = container.querySelector('#modal-vault-title');

  // Preset Provider Defaults
  const PROVIDER_DEFAULTS = {
    gemini: {
      label: 'Google Gemini 3.5 Flash-Lite (Instan & 15 RPM)',
      model: 'gemini-3.5-flash-lite',
      baseUrl: 'https://generativelanguage.googleapis.com/v1beta',
      hint: 'Model tercepat: gemini-3.5-flash-lite (15 RPM instan), gemini-3.1-flash-lite, atau gemini-2.5-flash-lite'
    },
    groq: {
      label: 'Groq - GPT OSS 120B (Ultra Cepat ~0.2s)',
      model: LATEST_GROQ_DEFAULT_MODEL,
      baseUrl: 'https://api.groq.com/openai/v1',
      hint: 'Model aktif Groq: openai/gpt-oss-120b atau qwen/qwen3.8-27b (~0.2 detik!)'
    },
    deepseek: {
      label: 'DeepSeek Chat (V3)',
      model: 'deepseek-chat',
      baseUrl: 'https://api.deepseek.com/v1',
      hint: 'Model rekomendasi: deepseek-chat'
    },
    qwen: {
      label: 'Qwen 2.5 72B (OpenRouter)',
      model: 'qwen/qwen-2.5-72b-instruct',
      baseUrl: 'https://openrouter.ai/api/v1',
      hint: 'Model rekomendasi: qwen/qwen-2.5-72b-instruct'
    },
    openai: {
      label: 'OpenAI GPT-4o Mini',
      model: 'gpt-4o-mini',
      baseUrl: 'https://api.openai.com/v1',
      hint: 'Dapat digunakan untuk OpenAI asli atau server lokal seperti Ollama/LMStudio'
    }
  };

  providerSelect.onchange = () => {
    const selected = providerSelect.value;
    const def = PROVIDER_DEFAULTS[selected];
    if (def) {
      if (!idInput.value) {
        labelInput.value = def.label;
      }
      modelInput.value = def.model;
      baseurlInput.value = def.baseUrl;
      container.querySelector('#vault-model-hint').textContent = def.hint;
    }
  };

  // Open Modal Tambah
  container.querySelector('#btn-add-vault-entry').onclick = () => {
    modalTitle.textContent = 'Tambah Kunci API Baru ke Gudang';
    idInput.value = '';
    labelInput.value = 'Google Gemini (Otomatis Deteksi)';
    providerSelect.value = 'gemini';
    modelInput.value = LATEST_GEMINI_DEFAULT_MODEL;
    baseurlInput.value = 'https://generativelanguage.googleapis.com/v1beta';
    apiKeyInput.value = '';
    apiKeyInput.placeholder = 'Masukkan API Key Anda...';
    activeCheck.checked = true;
    modal.classList.remove('hidden');
  };

  const closeModal = () => modal.classList.add('hidden');
  container.querySelector('#btn-close-modal-vault').onclick = closeModal;
  container.querySelector('#btn-cancel-modal-vault').onclick = closeModal;

  // Submit Form Kunci API
  form.onsubmit = (e) => {
    e.preventDefault();
    const id = idInput.value;
    const rawTypedKey = apiKeyInput.value.trim();
    const resolvedKey = rawTypedKey || (id ? getBuiltinSystemKey(id) : '') || '';
    const entryData = {
      label: labelInput.value.trim(),
      provider: providerSelect.value,
      model: modelInput.value.trim(),
      baseUrl: baseurlInput.value.trim(),
      apiKey: resolvedKey,
      isActive: activeCheck.checked
    };

    if (id) {
      store.updateApiVaultEntry(id, entryData);
      showToast('Kunci API berhasil diperbarui.', 'success');
    } else {
      store.addApiVaultEntry(entryData);
      showToast('Kunci API baru berhasil ditambahkan ke Gudang API!', 'success');
    }

    closeModal();
    renderSettingsView(container);
  };

  // Tombol Aktifkan
  container.querySelectorAll('.btn-activate-vault').forEach(btn => {
    btn.onclick = () => {
      const id = btn.dataset.id;
      store.setActiveApi(id);
      showToast('Kunci API aktif berhasil diubah!', 'success');
      renderSettingsView(container);
    };
  });

  // Tombol Edit
  container.querySelectorAll('.btn-edit-vault').forEach(btn => {
    btn.onclick = () => {
      const id = btn.dataset.id;
      const entry = store.getApiVault().find(v => v.id === id);
      if (!entry) return;

      modalTitle.textContent = `Edit Kunci API: ${entry.label}`;
      idInput.value = entry.id;
      labelInput.value = entry.label;
      providerSelect.value = entry.provider || 'gemini';
      modelInput.value = entry.model;
      baseurlInput.value = entry.baseUrl;
      if (isUsingBuiltinSystemKey(entry)) {
        apiKeyInput.value = '';
        apiKeyInput.placeholder = 'Memakai API bawaan sistem (isi jika ingin ganti pribadi)';
      } else {
        apiKeyInput.value = entry.apiKey || '';
        apiKeyInput.placeholder = 'Masukkan API Key Anda...';
      }
      activeCheck.checked = entry.id === store.getState().activeApiId || entry.isActive;

      modal.classList.remove('hidden');
    };
  });

  // Menu Kebab (⋮)
  const closeAllKebabs = () => {
    container.querySelectorAll('.api-kebab-menu').forEach(m => m.classList.add('hidden'));
  };

  container.querySelectorAll('.btn-api-kebab').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      const id = btn.dataset.id;
      const targetMenu = container.querySelector(`.api-kebab-menu[data-menu-id="${id}"]`);
      const wasHidden = targetMenu?.classList.contains('hidden');
      closeAllKebabs();
      if (targetMenu && wasHidden) {
        targetMenu.classList.remove('hidden');
      }
    };
  });

  container.querySelectorAll('.btn-kebab-edit').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      closeAllKebabs();
      const editBtn = container.querySelector(`.btn-edit-vault[data-id="${btn.dataset.id}"]`);
      if (editBtn) editBtn.click();
    };
  });

  container.addEventListener('click', (e) => {
    if (!e.target.closest('.api-kebab-wrap')) {
      closeAllKebabs();
    }
  });

  // Tombol Hapus
  container.querySelectorAll('.btn-delete-vault').forEach(btn => {
    btn.onclick = () => {
      const id = btn.dataset.id;
      if (confirm('Apakah Anda yakin ingin menghapus kunci API ini dari gudang?')) {
        store.removeApiVaultEntry(id);
        showToast('Kunci API dihapus dari gudang.', 'info');
        renderSettingsView(container);
      }
    };
  });

  // Tombol Uji Koneksi
  container.querySelectorAll('.btn-test-vault').forEach(btn => {
    btn.onclick = async () => {
      const id = btn.dataset.id;
      const entry = store.getApiVault().find(v => v.id === id);
      if (!entry) return;

      if (!entry.apiKey && entry.provider !== 'gemini') {
        showToast(`Isi API Key untuk ${entry.label} terlebih dahulu sebelum menguji koneksi.`, 'warning');
        return;
      }

      btn.disabled = true;
      btn.textContent = '⏳ Menguji...';
      try {
        let activeModelLabel = entry.model;

        // Untuk Gemini: Otomatis deteksi model aktif terlebih dahulu
        if (entry.provider === 'gemini' && entry.apiKey) {
          const detected = await detectActiveGeminiModel(entry.apiKey, entry.baseUrl);
          if (detected) {
            activeModelLabel = detected;
            entry.model = detected;
            store.updateApiVaultEntry(entry.id, {
              model: detected,
              label: `Google Gemini (${detected})`
            });
          }
        }

        // Untuk Groq: Otomatis deteksi model aktif terlebih dahulu
        if (entry.provider === 'groq' && entry.apiKey) {
          const detected = await detectActiveGroqModel(entry.apiKey, entry.baseUrl);
          if (detected) {
            activeModelLabel = detected;
            entry.model = detected;
            store.updateApiVaultEntry(entry.id, {
              model: detected,
              label: `Groq (${detected.split('/')[1] || detected})`
            });
          }
        }

        const testRes = await executeAiRequest({
          config: entry,
          prompt: 'Tuliskan satu kata "KONEKSI_BERHASIL" dalam format JSON: {"status": "KONEKSI_BERHASIL"}',
          systemPrompt: 'Kamu adalah asisten penguji koneksi API. Berikan JSON murni.'
        });

        if (testRes) {
          showToast(`✓ Koneksi Berhasil! Model aktif terdeteksi: ${activeModelLabel}`, 'success');
          renderSettingsView(container);
        } else {
          showToast(`Koneksi terhubung tetapi tidak ada data respons dari ${entry.label}.`, 'warning');
        }
      } catch (err) {
        showToast(`❌ Gagal terhubung ke ${entry.label}: ${err.message}`, 'error');
      } finally {
        btn.disabled = false;
        btn.innerHTML = `${SVG_SIGNAL_14}<span>Uji Koneksi</span>`;
      }
    };
  });
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str || '';
  return div.innerHTML;
}
