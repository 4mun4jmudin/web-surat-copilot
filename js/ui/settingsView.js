/**
 * settingsView.js - Gudang API Multi-Provider & Pengaturan Sistem
 * Mengelola banyak kunci API (Google Gemini, Groq, DeepSeek, Qwen, OpenAI-Compatible)
 * dengan fitur Uji Koneksi, Aktivasi Cepat, dan Auto-Fallback.
 */

import { store } from '../state.js';
import { executeAiRequest, detectActiveGeminiModel, detectActiveGroqModel, LATEST_GEMINI_DEFAULT_MODEL, LATEST_GROQ_DEFAULT_MODEL } from '../ai/client.js';
import { showToast } from './notification.js';

export function renderSettingsView(container) {
  if (!container) return;

  const state = store.getState();
  const vault = store.getApiVault();
  const activeId = state.activeApiId;

  container.innerHTML = `
    <div class="settings-view-container" style="max-width: 900px; margin: 24px auto; padding: 0 16px; overflow-y: auto; height: 100%;">
      
      <!-- HEADER GUDANG API -->
      <div class="card" style="margin-bottom: 20px;">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
          <div>
            <h2 style="font-size: 20px; font-weight: 700; color: var(--dark); display: flex; align-items: center; gap: 8px;">
              <span>🔑</span> Gudang API Multi-Provider
            </h2>
            <p style="color: var(--text-muted); font-size: 13px; margin-top: 4px;">
              Simpan beberapa kunci API sekaligus dari berbagai AI (Gemini, Groq, DeepSeek, Qwen, dll). Jika kuota satu habis, sistem otomatis beralih ke kunci berikutnya.
            </p>
          </div>
          <button id="btn-add-vault-entry" class="btn btn-primary">
            ➕ Tambah Kunci API Baru
          </button>
        </div>
      </div>

      <!-- DAFTAR KUNCI API DI GUDANG -->
      <div class="vault-list-container" style="display: flex; flex-direction: column; gap: 12px; margin-bottom: 24px;">
        ${vault.map(entry => {
          const isActive = entry.id === activeId || entry.isActive;
          const hasKey = !!(entry.apiKey && entry.apiKey.trim().length > 0);
          
          let providerBadge = '#1e3a8a';
          if (entry.provider === 'groq') providerBadge = '#ea580c';
          if (entry.provider === 'deepseek') providerBadge = '#0284c7';
          if (entry.provider === 'qwen') providerBadge = '#7c3aed';
          if (entry.provider === 'openai') providerBadge = '#15803d';

          return `
            <div class="card vault-card ${isActive ? 'vault-card-active' : ''}" style="border: 2px solid ${isActive ? 'var(--primary)' : 'var(--border)'}; border-radius: 8px; padding: 16px; background: white; transition: all 0.2s;">
              <div style="display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 12px;">
                <div style="flex: 1; min-width: 250px;">
                  <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 6px;">
                    <span style="background: ${providerBadge}; color: white; font-size: 10px; font-weight: 700; padding: 2px 8px; border-radius: 12px; text-transform: uppercase;">
                      ${entry.provider}
                    </span>
                    <strong style="font-size: 15px; color: var(--dark);">${escapeHtml(entry.label)}</strong>
                    ${isActive ? '<span style="background: #dcfce7; color: #166534; font-size: 11px; font-weight: 600; padding: 2px 8px; border-radius: 4px;">✓ Digunakan Saat Ini</span>' : ''}
                  </div>
                  
                  <div style="font-size: 12.5px; color: var(--text-muted); display: flex; flex-direction: column; gap: 3px;">
                    <div><strong>Model:</strong> <code>${escapeHtml(entry.model || '-')}</code></div>
                    <div><strong>Endpoint URL:</strong> <code>${escapeHtml(entry.baseUrl || '-')}</code></div>
                    <div><strong>Status Kunci:</strong> ${hasKey ? '🔑 Tersimpan (' + entry.apiKey.slice(0, 6) + '...' + entry.apiKey.slice(-4) + ')' : '<span style="color: #b45309;">⚠️ Kunci Belum Diisi</span>'}</div>
                  </div>
                </div>

                <div style="display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
                  ${!isActive ? `
                    <button class="btn btn-sm btn-secondary btn-activate-vault" data-id="${entry.id}">
                      Aktifkan
                    </button>
                  ` : ''}
                  <button class="btn btn-sm btn-secondary btn-test-vault" data-id="${entry.id}">
                    🧪 Uji Koneksi
                  </button>
                  <button class="btn btn-sm btn-secondary btn-edit-vault" data-id="${entry.id}">
                    ✏ Edit
                  </button>
                  ${!entry.isBuiltin ? `
                    <button class="btn btn-sm btn-danger btn-delete-vault" data-id="${entry.id}">
                      🗑 Hapus
                    </button>
                  ` : ''}
                </div>
              </div>
            </div>
          `;
        }).join('')}
      </div>

      <!-- MODAL TAMBAH / EDIT KUNCI API -->
      <div id="modal-vault-form" class="modal-overlay hidden" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 10000; display: flex; align-items: center; justify-content: center;">
        <div class="card" style="width: 100%; max-width: 520px; background: white; border-radius: 8px; padding: 20px; box-shadow: 0 10px 30px rgba(0,0,0,0.25);">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
            <h3 id="modal-vault-title" style="font-size: 16px; font-weight: 700;">Tambah Kunci API Baru</h3>
            <button id="btn-close-modal-vault" class="btn-icon" style="font-size: 20px;">&times;</button>
          </div>

          <form id="form-vault-entry">
            <input type="hidden" id="input-vault-id" value="">
            
            <div class="form-group">
              <label>Pilih Penyedia AI (Provider):</label>
              <select id="select-vault-provider" class="form-control">
                <option value="gemini">Google Gemini (Gemini 2.0 Flash / 1.5 Flash)</option>
                <option value="groq">Groq (Llama 3.3 70B - Sangat Cepat)</option>
                <option value="deepseek">DeepSeek (DeepSeek V3 / R1)</option>
                <option value="qwen">Qwen / Alibaba (via OpenRouter)</option>
                <option value="openai">OpenAI / Custom Compatible</option>
              </select>
            </div>

            <div class="form-group">
              <label>Nama Label Kunci:</label>
              <input type="text" id="input-vault-label" class="form-control" placeholder="Contoh: Gemini Akun Utama" required>
            </div>

            <div class="form-group">
              <label>Model ID:</label>
              <input type="text" id="input-vault-model" class="form-control" value="gemini-3.5-flash-lite" required>
              <small class="text-muted" id="vault-model-hint">Pilih model yang didukung oleh penyedia AI.</small>
            </div>

            <div class="form-group">
              <label>Base URL Endpoint:</label>
              <input type="text" id="input-vault-baseurl" class="form-control" value="https://generativelanguage.googleapis.com/v1beta" required>
            </div>

            <div class="form-group">
              <label>API Key / Token Rahasia:</label>
              <input type="password" id="input-vault-apikey" class="form-control" placeholder="Masukkan API Key Anda..." autocomplete="off">
              <small class="text-muted">Kunci disimpan aman hanya di browser lokal Anda (IndexedDB).</small>
            </div>

            <div class="form-group">
              <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
                <input type="checkbox" id="check-vault-active" checked>
                <span>Jadikan Kunci Aktif untuk Pembuatan Surat</span>
              </label>
            </div>

            <div style="display: flex; justify-content: flex-end; gap: 8px; margin-top: 18px;">
              <button type="button" id="btn-cancel-modal-vault" class="btn btn-secondary">Batal</button>
              <button type="submit" class="btn btn-primary">Simpan Kunci</button>
            </div>
          </form>
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
    const entryData = {
      label: labelInput.value.trim(),
      provider: providerSelect.value,
      model: modelInput.value.trim(),
      baseUrl: baseurlInput.value.trim(),
      apiKey: apiKeyInput.value.trim(),
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
      apiKeyInput.value = entry.apiKey || '';
      activeCheck.checked = entry.id === store.getState().activeApiId || entry.isActive;

      modal.classList.remove('hidden');
    };
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
        btn.textContent = '🧪 Uji Koneksi';
      }
    };
  });
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str || '';
  return div.innerHTML;
}
