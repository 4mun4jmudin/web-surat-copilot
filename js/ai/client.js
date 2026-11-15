/**
 * client.js - Klien API AI Multi-Provider (Gudang API)
 * Fitur:
 * 1. Otomatis Deteksi & Filter Model Stabil Google Gemini (Mengecualikan omni/preview berlimit 0).
 * 2. Rotasi Otomatis Antar-Model saat Kuota 429 / 404 (Per-Model Quota Resilience).
 * 3. Multi-Provider Gudang API: Gemini, Groq (Llama 3.3), DeepSeek (V3), Qwen, OpenAI-Compatible.
 * 4. Auto-Fallback Antar-Kunci di Gudang API saat seluruh model pada satu kunci habis kuota.
 */

import { store, getBuiltinFallbackConfigs, getBuiltinSystemKey } from '../state.js';
import { POLISHER_SYSTEM_PROMPT, buildPolisherUserPrompt, EDITOR_SYSTEM_PROMPT, buildEditorUserPrompt } from './prompts.js';
import { formatSmartNomorSurat } from '../letters/numbering.js';

// Proxy default Cloudflare Functions (produksi)
const DEFAULT_PROXY_ENDPOINT = '/api/ai';

// Daftar prioritas model Google Gemini yang stabil (GA) pada Free Tier (Sesuai AI Studio Rate Limits)
// Prioritaskan model Flash-Lite yang berkecepatan tinggi (15 RPM) dibanding Flash standar (5 RPM)
export const STABLE_GEMINI_CANDIDATES = [
  'gemini-3.5-flash-lite',
  'gemini-3.1-flash-lite',
  'gemini-flash-lite-latest',
  'gemini-flash-latest',
  'gemini-2.5-flash-lite',
  'gemini-2.5-flash',
  'gemini-3.8-flash'
];

export const LATEST_GEMINI_DEFAULT_MODEL = 'gemini-3.5-flash-lite';

// Cache in-memory dan localStorage untuk model yang terdeteksi
let cachedDetectedGeminiModel = null;
if (typeof localStorage !== 'undefined') {
  const saved = localStorage.getItem('detected_gemini_model');
  if (saved && !saved.includes('omni') && !saved.includes('2.0-flash') && !saved.includes('1.5-pro')) {
    cachedDetectedGeminiModel = saved;
  }
}

/**
 * Mendeteksi secara otomatis model Gemini stabil yang aktif dan tersedia untuk API Key yang diberikan.
 * Mengabaikan model omni/preview/experimental yang memiliki limit kuota 0 pada Free Tier.
 * @param {string} apiKey
 * @param {string} [baseUrl]
 * @returns {Promise<string>} Nama model terdeteksi (contoh: "gemini-2.5-flash")
 */
export async function detectActiveGeminiModel(apiKey, baseUrl = 'https://generativelanguage.googleapis.com/v1beta') {
  if (cachedDetectedGeminiModel) {
    return cachedDetectedGeminiModel;
  }

  if (!apiKey || !apiKey.trim()) {
    return LATEST_GEMINI_DEFAULT_MODEL;
  }

  try {
    const cleanKey = apiKey.trim();
    const endpoint = `${baseUrl.replace(/\/+$/, '')}/models?key=${cleanKey}`;
    const res = await fetch(endpoint, {
      method: 'GET',
      headers: { 'Content-Type': 'application/json' }
    });

    if (!res.ok) {
      console.warn(`[Gemini Auto-Detect] Gagal list models (${res.status}), menggunakan default ${LATEST_GEMINI_DEFAULT_MODEL}`);
      return LATEST_GEMINI_DEFAULT_MODEL;
    }

    const data = await res.json();
    const models = Array.isArray(data.models) ? data.models : [];

    // Filter model teks yang mendukung generateContent
    const textModels = models.filter(m => 
      Array.isArray(m.supportedGenerationMethods) && 
      m.supportedGenerationMethods.includes('generateContent')
    );

    if (textModels.length === 0) {
      return LATEST_GEMINI_DEFAULT_MODEL;
    }

    // Normalisasi nama-nama model yang tersedia
    const availableModelNames = textModels.map(m => (m.name || '').replace(/^models\//, '').toLowerCase());

    // 1. Cek prioritas model stabil teruji (menghindari omni/preview yang limit 0 pada free tier)
    for (const candidate of STABLE_GEMINI_CANDIDATES) {
      if (availableModelNames.includes(candidate.toLowerCase())) {
        cachedDetectedGeminiModel = candidate;
        if (typeof localStorage !== 'undefined') {
          localStorage.setItem('detected_gemini_model', candidate);
        }
        return candidate;
      }
    }

    // 2. Jika tidak ada yang cocok persis di kandidat utama, cari model flash yang aman non-omni, non-preview
    // Dahulukan varian 'lite' untuk latency sub-detik
    const safeFlashCandidates = textModels.filter(m => {
      const name = m.name.toLowerCase();
      return name.includes('flash') && 
        !name.includes('omni') && 
        !name.includes('preview') && 
        !name.includes('exp') && 
        !name.includes('vision') && 
        !name.includes('embedding');
    }).sort((a, b) => {
      const aLite = a.name.toLowerCase().includes('lite') ? 1 : 0;
      const bLite = b.name.toLowerCase().includes('lite') ? 1 : 0;
      return bLite - aLite;
    });

    if (safeFlashCandidates.length > 0) {
      const cleanName = safeFlashCandidates[0].name.replace(/^models\//, '').trim();
      cachedDetectedGeminiModel = cleanName;
      if (typeof localStorage !== 'undefined') {
        localStorage.setItem('detected_gemini_model', cleanName);
      }
      return cleanName;
    }

    return LATEST_GEMINI_DEFAULT_MODEL;
  } catch (err) {
    console.warn('[Gemini Auto-Detect] Gagal mendeteksi otomatis model Gemini:', err);
    return LATEST_GEMINI_DEFAULT_MODEL;
  }
}

/**
 * Normalisasi model ID Google Gemini untuk mencegah penggunaan model usang atau model omni berlimit 0
 */
export function normalizeGeminiModel(modelName = '') {
  const m = modelName.trim().toLowerCase().replace(/^models\//, '');
  // Cegah model usang atau model omni-preview berlimit 0 (misal: gemini-omni-flash-preview, gemini-2.0-flash)
  if (!m || m.includes('omni') || m.includes('1.5-pro') || m.includes('2.0-flash') || m.includes('2-flash') || m === 'gemini-pro') {
    return cachedDetectedGeminiModel || LATEST_GEMINI_DEFAULT_MODEL;
  }
  return m;
}

/**
 * Ekstrak model rekomendasi dari pesan error Google jika ada
 */
export function extractSuggestedGeminiModel(errorMessage = '') {
  const targetedMatch = errorMessage.match(/(?:use|to)\s+models\/([a-zA-Z0-9\.\-_]+)/i);
  if (targetedMatch && targetedMatch[1]) {
    const candidate = targetedMatch[1];
    if (!candidate.includes('omni')) {
      return candidate;
    }
  }
  return null;
}

// Prioritas model Groq yang stabil dan aktif
export const STABLE_GROQ_CANDIDATES = [
  'openai/gpt-oss-120b',
  'qwen/qwen3.8-27b',
  'openai/gpt-oss-20b'
];

export const LATEST_GROQ_DEFAULT_MODEL = 'openai/gpt-oss-120b';

let cachedDetectedGroqModel = null;

/**
 * Mendeteksi secara otomatis model Groq aktif yang tersedia untuk API Key yang diberikan.
 * @param {string} apiKey
 * @param {string} [baseUrl]
 * @returns {Promise<string>}
 */
export async function detectActiveGroqModel(apiKey, baseUrl = 'https://api.groq.com/openai/v1') {
  if (cachedDetectedGroqModel) {
    return cachedDetectedGroqModel;
  }

  if (!apiKey || !apiKey.trim()) {
    return LATEST_GROQ_DEFAULT_MODEL;
  }

  try {
    const cleanKey = apiKey.trim();
    const endpoint = `${baseUrl.replace(/\/+$/, '')}/models`;
    const res = await fetch(endpoint, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${cleanKey}`,
        'Content-Type': 'application/json'
      }
    });

    if (!res.ok) {
      return LATEST_GROQ_DEFAULT_MODEL;
    }

    const data = await res.json();
    const models = Array.isArray(data.data) ? data.data : [];

    // Filter model teks (bukan audio whisper atau prompt guard)
    const textModels = models.filter(m => {
      const id = (m.id || '').toLowerCase();
      return !id.includes('whisper') && 
        !id.includes('prompt-guard') && 
        !id.includes('safeguard') && 
        !id.includes('orpheus');
    });

    const modelIds = textModels.map(m => m.id);

    // Prioritaskan model unggulan
    for (const candidate of STABLE_GROQ_CANDIDATES) {
      if (modelIds.includes(candidate)) {
        cachedDetectedGroqModel = candidate;
        return candidate;
      }
    }

    if (modelIds.length > 0) {
      cachedDetectedGroqModel = modelIds[0];
      return modelIds[0];
    }

    return LATEST_GROQ_DEFAULT_MODEL;
  } catch (err) {
    console.warn('[Groq Auto-Detect] Gagal deteksi model Groq:', err);
    return LATEST_GROQ_DEFAULT_MODEL;
  }
}

/**
 * Normalisasi model Groq usang (llama-3.3-70b-versatile -> openai/gpt-oss-120b)
 */
export function normalizeGroqModel(modelName = '') {
  const m = modelName.trim().toLowerCase();
  if (!m || m.includes('llama-3.3') || m.includes('llama-3.1') || m.includes('mixtral')) {
    return cachedDetectedGroqModel || LATEST_GROQ_DEFAULT_MODEL;
  }
  return modelName;
}

/**
 * Membersihkan format output markdown ```json ... ```
 */
export function cleanJsonOutput(raw) {
  if (!raw) return '';
  let str = String(raw).trim();
  if (str.startsWith('```json')) {
    str = str.replace(/^```json\s*/i, '').replace(/```\s*$/, '');
  } else if (str.startsWith('```')) {
    str = str.replace(/^```\s*/, '').replace(/```\s*$/, '');
  }
  return str.trim();
}

/**
 * Mendapatkan konfigurasi AI aktif dari State Store
 */
export function getActiveAiConfig() {
  const active = store.getActiveApi();
  if (active) return active;

  return {
    id: 'gemini-default',
    label: 'Google Gemini (Otomatis)',
    provider: 'gemini',
    model: LATEST_GEMINI_DEFAULT_MODEL,
    apiKey: '',
    baseUrl: 'https://generativelanguage.googleapis.com/v1beta',
    isActive: true
  };
}

/**
 * Menyimpan konfigurasi kompatibilitas ke localStorage
 */
export function saveAiConfig({ apiKey, model, useProxy }) {
  if (typeof localStorage === 'undefined') return;
  if (apiKey !== undefined) localStorage.setItem('ai_api_key', apiKey.trim());
  if (model !== undefined) localStorage.setItem('ai_model', model.trim());
  if (useProxy !== undefined) localStorage.setItem('ai_use_proxy', String(useProxy));
}

/**
 * Melakukan panggilan API AI ke provider tunggal dengan Rotasi Model Otomatis saat 429/404/Timeout
 */
export async function executeAiRequest({
  config,
  prompt,
  systemPrompt,
  stream = false,
  onChunk,
  signal,
  timeoutMs = 12000 // Batas maksimal 12 detik per panggilan agar tidak hang/muter terus
}) {
  const provider = config.provider || 'gemini';
  const apiKey = (config.apiKey || '').trim();

  // 1. PROVIDER: GOOGLE GEMINI
  if (provider === 'gemini') {
    const baseUrl = config.baseUrl || 'https://generativelanguage.googleapis.com/v1beta';

    if (!apiKey) {
      // Coba proxy Cloudflare
      const res = await fetch(DEFAULT_PROXY_ENDPOINT, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ prompt, systemPrompt, stream, model: normalizeGeminiModel(config.model) }),
        signal
      });
      if (!res.ok) {
        const errJson = await res.json().catch(() => ({}));
        throw new Error(errJson.error || `Proxy Gemini Gagal (${res.status})`);
      }
      const data = await res.json();
      return data.text || '';
    }

    // Tentukan model awal yang aman (menghindari omni)
    let initialModel = normalizeGeminiModel(config.model);
    if (!initialModel || initialModel.includes('omni') || initialModel.includes('2.0-flash')) {
      initialModel = await detectActiveGeminiModel(apiKey, baseUrl);
    }

    // Siapkan daftar model yang akan dicoba secara berurutan jika terjadi 429, 404, atau Timeout
    const modelsToTry = [
      initialModel,
      ...STABLE_GEMINI_CANDIDATES.filter(m => m !== initialModel)
    ];

    const buildGeminiPayload = (targetModel) => {
      const generationConfig = {
        temperature: 0.2,
        responseMimeType: 'application/json'
      };
      // thinkingBudget hanya didukung seri Gemini 2.5; seri 3.5+ / -latest menolak argumen ini (400 INVALID_ARGUMENT)
      if (String(targetModel).startsWith('gemini-2.5')) {
        generationConfig.thinkingConfig = { thinkingBudget: 0 };
      }
      const nextPayload = {
        contents: [{ role: 'user', parts: [{ text: prompt }] }],
        generationConfig
      };
      if (systemPrompt) {
        nextPayload.systemInstruction = { parts: [{ text: systemPrompt }] };
      }
      return nextPayload;
    };

    let lastGeminiError = null;

    for (const targetModel of modelsToTry) {
      let timeoutId = null;
      const attemptController = new AbortController();
      timeoutId = setTimeout(() => {
        attemptController.abort(new Error(`Timeout: Model ${targetModel} tidak merespons dalam ${timeoutMs / 1000} detik.`));
      }, timeoutMs);

      const abortHandler = () => attemptController.abort();
      if (signal) {
        if (signal.aborted) throw new Error('Pembuatan surat dibatalkan oleh pengguna.');
        signal.addEventListener('abort', abortHandler, { once: true });
      }

      try {
        const payload = buildGeminiPayload(targetModel);
        const endpoint = `${baseUrl}/models/${targetModel}:${stream ? 'streamGenerateContent?alt=sse&' : 'generateContent?'}key=${apiKey}`;
        let res = await fetch(endpoint, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
          signal: attemptController.signal
        });

        // Jika Google API mengembalikan 400 karena parameter thinkingConfig tidak didukung pada model tertentu:
        if (!res.ok && res.status === 400) {
          const errPeek = await res.clone().text().catch(() => '');
          if (errPeek.includes('thinkingConfig') || errPeek.includes('thinkingBudget') || errPeek.includes('unknown') || errPeek.includes('INVALID_ARGUMENT')) {
            const fallbackPayload = {
              contents: payload.contents,
              generationConfig: {
                temperature: 0.2,
                responseMimeType: 'application/json'
              }
            };
            if (payload.systemInstruction) fallbackPayload.systemInstruction = payload.systemInstruction;
            res = await fetch(endpoint, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(fallbackPayload),
              signal: attemptController.signal
            });
          }
        }

        if (res.ok) {
          // Model ini berhasil! Simpan ke cache & perbarui entri di Gudang API jika berbeda
          cachedDetectedGeminiModel = targetModel;
          if (typeof localStorage !== 'undefined') {
            localStorage.setItem('detected_gemini_model', targetModel);
          }
          if (config.id && config.model !== targetModel) {
            store.updateApiVaultEntry(config.id, {
              model: targetModel,
              label: `Google Gemini (${targetModel})`
            });
          }

          if (stream && res.body && onChunk) {
            return readGeminiStream(res.body, onChunk);
          }

          const data = await res.json();
          return data.candidates?.[0]?.content?.parts?.[0]?.text || '';
        }

        const errText = await res.text();

        // 429 Quota Exceeded atau 404 Not Found:
        // Coba model stabil lainnya di key yang sama!
        if (res.status === 429 || res.status === 404) {
          console.warn(`[Gemini Rotasi] Model ${targetModel} mengalami status ${res.status}. Mencoba model alternatif...`);
          lastGeminiError = new Error(
            res.status === 429
              ? `Kuota Free Tier untuk model ${targetModel} telah mencapai batas harian (429: Resource Exhausted).`
              : `Model ${targetModel} tidak tersedia untuk API key ini (404).`
          );
          continue; // Lanjut ke candidate berikutnya
        }

        throw new Error(`Panggilan Gemini API Gagal (${res.status}): ${errText}`);
      } catch (reqErr) {
        if (signal && signal.aborted) {
          throw new Error('Pembuatan surat dibatalkan oleh pengguna.');
        }
        console.warn(`[Gemini Rotasi] Model ${targetModel} mengalami kendala/timeout (${reqErr.message}). Mencoba model alternatif...`);
        lastGeminiError = reqErr;
      } finally {
        if (timeoutId) clearTimeout(timeoutId);
        if (signal) signal.removeEventListener('abort', abortHandler);
      }
    }

    // Jika seluruh model stabil pada kunci ini telah habis kuota atau timeout:
    throw lastGeminiError || new Error('Seluruh model Google Gemini pada API key ini telah mencapai batas kuota atau tidak merespons.');
  }

  // 2. PROVIDER: OPENAI-COMPATIBLE (GROQ, DEEPSEEK, QWEN, OPENROUTER, OPENAI, OLLAMA)
  let endpoint = config.baseUrl || 'https://api.groq.com/openai/v1';
  if (!endpoint.endsWith('/chat/completions')) {
    endpoint = endpoint.replace(/\/+$/, '') + '/chat/completions';
  }

  let modelId = config.model;
  if (provider === 'groq') {
    modelId = normalizeGroqModel(modelId);
  } else if (provider === 'deepseek') {
    modelId = modelId || 'deepseek-chat';
  } else {
    modelId = modelId || 'gpt-4o-mini';
  }

  const messages = [];
  if (systemPrompt) {
    messages.push({ role: 'system', content: systemPrompt });
  }
  messages.push({ role: 'user', content: prompt });

  const headers = {
    'Content-Type': 'application/json'
  };
  if (apiKey) {
    headers['Authorization'] = `Bearer ${apiKey}`;
  }

  const bodyPayload = {
    model: modelId,
    messages,
    temperature: 0.2,
    response_format: { type: 'json_object' },
    stream: false
  };

  const attemptController = new AbortController();
  const timeoutId = setTimeout(() => {
    attemptController.abort(new Error(`Timeout (${timeoutMs / 1000}s) saat memanggil ${provider.toUpperCase()}`));
  }, timeoutMs);

  const abortHandler = () => attemptController.abort();
  if (signal) {
    if (signal.aborted) throw new Error('Pembuatan surat dibatalkan oleh pengguna.');
    signal.addEventListener('abort', abortHandler, { once: true });
  }

  let res;
  try {
    res = await fetch(endpoint, {
      method: 'POST',
      headers,
      body: JSON.stringify(bodyPayload),
      signal: attemptController.signal
    });

    // Self-Healing untuk Groq jika model yang dipanggil 404 (model lama terdeprekasi)
    if (!res.ok && res.status === 404 && provider === 'groq') {
      const errText = await res.text();
      if (errText.includes('model_not_found') || errText.includes('does not exist')) {
        const detected = await detectActiveGroqModel(apiKey, config.baseUrl);
        if (detected && detected !== modelId) {
          console.warn(`[Groq Self-Healing] Beralih otomatis ke model aktif Groq: ${detected}`);
          cachedDetectedGroqModel = detected;
          bodyPayload.model = detected;

          if (config.id) {
            store.updateApiVaultEntry(config.id, {
              model: detected,
              label: `Groq (${detected.split('/')[1] || detected})`
            });
          }

          res = await fetch(endpoint, {
            method: 'POST',
            headers,
            body: JSON.stringify(bodyPayload),
            signal: attemptController.signal
          });
        }
      }
    }
  } catch (reqErr) {
    if (signal && signal.aborted) {
      throw new Error('Pembuatan surat dibatalkan oleh pengguna.');
    }
    throw reqErr;
  } finally {
    clearTimeout(timeoutId);
    if (signal) signal.removeEventListener('abort', abortHandler);
  }

  if (!res.ok) {
    const errText = await res.text();
    let friendlyMessage = errText;
    try {
      const errObj = JSON.parse(errText);
      if (res.status === 402 || errObj.error?.type === 'insufficient_balance_error' || errObj.error?.message?.includes('Insufficient Balance')) {
        friendlyMessage = `Saldo/Kredit Akun ${provider.toUpperCase()} Habis ($0.00). Akun di platform ${provider}.com memerlukan isi ulang saldo (top up) atau gunakan alternatif GRATIS (Groq Llama 3.3 / Google Gemini) di Gudang API.`;
      } else if (res.status === 401) {
        friendlyMessage = `Kunci API ${provider.toUpperCase()} tidak valid atau salah ketik. Periksa kembali API Key Anda.`;
      } else if (res.status === 429) {
        friendlyMessage = `Batas Kuota / Rate Limit ${provider.toUpperCase()} tercapai (429). Coba kembali beberapa saat lagi atau gunakan kunci lain di Gudang API.`;
      } else if (errObj.error?.message) {
        friendlyMessage = errObj.error.message;
      }
    } catch (e) {}

    throw new Error(`Panggilan ${provider.toUpperCase()} Gagal (${res.status}): ${friendlyMessage}`);
  }

  const data = await res.json();
  return data.choices?.[0]?.message?.content || '';
}

/**
 * Panggilan AI Terpusat dengan Gudang API & Otomatis Fallback Antar-Kunci
 */
export async function callAiApi({ prompt, systemPrompt, stream = false, onChunk, signal, preferredApiId }) {
  const vault = store.getApiVault();
  const candidateConfigs = [];
  const seenKeys = new Set();

  const resolveConfig = (item) => {
    if (!item) return null;
    const rawKey = (item.apiKey || '').trim();
    const resolvedKey = rawKey || getBuiltinSystemKey(item.id) || '';
    return resolvedKey === rawKey ? item : { ...item, apiKey: resolvedKey };
  };

  const pushIfUsable = (item) => {
    const resolved = resolveConfig(item);
    if (!resolved) return;
    const isOllama = (resolved.provider || '').toLowerCase() === 'ollama';
    if (!resolved.apiKey && !isOllama) return;
    const dedupeKey = `${resolved.id || ''}|${resolved.provider || ''}|${resolved.apiKey || ''}`;
    if (seenKeys.has(dedupeKey)) return;
    seenKeys.add(dedupeKey);
    candidateConfigs.push(resolved);
  };

  if (preferredApiId) {
    pushIfUsable(vault.find(v => v.id === preferredApiId));
  }

  pushIfUsable(store.getActiveApi());

  // Tambahkan cadangan lain di gudang yang memiliki apiKey
  for (const item of vault) {
    pushIfUsable(item);
  }

  // Fallback otomatis ke API bawaan sistem jika pengguna belum mengatur atau menghapus kunci
  for (const builtinCfg of getBuiltinFallbackConfigs()) {
    pushIfUsable(builtinCfg);
  }

  if (candidateConfigs.length === 0) {
    candidateConfigs.push({
      id: 'fallback-gemini',
      provider: 'gemini',
      model: LATEST_GEMINI_DEFAULT_MODEL,
      apiKey: '',
      baseUrl: 'https://generativelanguage.googleapis.com/v1beta'
    });
  }

  let lastError = null;

  for (let i = 0; i < candidateConfigs.length; i++) {
    if (signal && signal.aborted) {
      throw new Error('Pembuatan surat dibatalkan oleh pengguna.');
    }
    const config = candidateConfigs[i];
    try {
      const result = await executeAiRequest({
        config,
        prompt,
        systemPrompt,
        stream,
        onChunk,
        signal
      });
      return result;
    } catch (err) {
      if (signal && signal.aborted) {
        throw new Error('Pembuatan surat dibatalkan oleh pengguna.');
      }
      lastError = err;
      console.warn(`[Gudang API] Gagal menggunakan ${config.label || config.model} (${err.message}). Mencoba cadangan berikutnya...`);
    }
  }

  throw lastError || new Error('Tidak ada API AI yang dapat dihubungi. Periksa kunci di Gudang API.');
}

/**
 * Panggilan AI Terstruktur dengan Validasi Skema & Retry Maksimal 2x
 */
export async function callAiWithValidation({ prompt, systemPrompt, validator, onChunk, signal, preferredApiId }) {
  let attempts = 0;
  const maxRetries = 2;
  let currentPrompt = prompt;
  let lastErrors = [];

  while (attempts <= maxRetries) {
    if (signal && signal.aborted) {
      throw new Error('Pembuatan surat dibatalkan oleh pengguna.');
    }
    try {
      const rawText = await callAiApi({
        prompt: currentPrompt,
        systemPrompt,
        stream: !!onChunk,
        onChunk,
        signal,
        preferredApiId
      });

      const cleaned = cleanJsonOutput(rawText);
      let parsed;
      try {
        parsed = JSON.parse(cleaned);
      } catch (jsonErr) {
        throw new Error(`Output bukan format JSON yang valid: ${jsonErr.message}`);
      }

      // Validasi terhadap validator yang diberikan
      const validationResult = validator(parsed);
      if (validationResult.valid) {
        return parsed;
      }

      lastErrors = validationResult.errors || ['Format tidak sesuai spesifikasi skema.'];
    } catch (err) {
      lastErrors = [err.message];
    }

    attempts++;
    if (attempts <= maxRetries) {
      currentPrompt = `
Permintaan sebelumnya menghasilkan data yang tidak valid:
${lastErrors.join('\n')}

Silakan perbaiki kesalahan di atas dan hasilkan JSON murni yang valid sesuai spesifikasi.
Kueri awal:
${prompt}
      `.trim();
    }
  }

  throw new Error(`AI gagal menghasilkan data yang valid setelah ${maxRetries} kali percobaan retry:\n- ${lastErrors.join('\n- ')}`);
}

/**
 * Pembaca SSE stream untuk Gemini
 */
async function readGeminiStream(streamBody, onChunk) {
  const reader = streamBody.getReader();
  const decoder = new TextDecoder('utf-8');
  let accumulatedText = '';

  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      const chunk = decoder.decode(value, { stream: true });
      const lines = chunk.split('\n');
      for (const line of lines) {
        if (line.startsWith('data: ')) {
          try {
            const dataObj = JSON.parse(line.slice(6));
            const partText = dataObj.candidates?.[0]?.content?.parts?.[0]?.text || '';
            accumulatedText += partText;
            if (onChunk) onChunk(accumulatedText, partText);
          } catch (e) {}
        }
      }
    }
  } finally {
    reader.releaseLock();
  }

  return accumulatedText;
}

/**
 * Merapihkan draf atau teks dokumen/surat yang berantakan menggunakan AI (AI Document Polisher)
 * Menata ulang tata bahasa, ejaan PUEBI, nomor surat, identitas/jadwal, dan struktur surat dinas resmi.
 * @param {string} messyText 
 * @param {object} [options]
 * @returns {Promise<object>} Document Model JSON resmi
 */
export async function polishMessyDocumentWithAi(messyText, options = {}) {
  const school = options.schoolProfile || store.getSchoolProfile();
  const dateStr = options.dateStr || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });

  const userPrompt = buildPolisherUserPrompt({
    messyText,
    schoolName: school.schoolName,
    principalName: school.principalName,
    dateStr
  });

  const rawJson = await callAiApi({
    prompt: userPrompt,
    systemPrompt: POLISHER_SYSTEM_PROMPT,
    preferredApiId: options.preferredApiId,
    signal: options.signal
  });

  const cleaned = cleanJsonOutput(rawJson);
  let parsed;
  try {
    parsed = JSON.parse(cleaned);
  } catch (err) {
    throw new Error('AI merespons tetapi format JSON perapihan tidak valid: ' + err.message);
  }

  const title = (parsed.letterTitle || parsed.title || 'SURAT DINAS SEKOLAH').trim();
  const rawNomor = (parsed.nomorSurat || '').trim();
  const smartNomor = formatSmartNomorSurat({
    rawNomor,
    contextText: `${title} ${parsed.hal || ''} ${parsed.perihal || ''} ${messyText.slice(0, 300)}`,
    forcedTypeCode: parsed.kodeJenisSurat || ''
  });

  const docModel = {
    meta: {
      title,
      author: school.schoolName || 'SMKS IT Al-Hawari',
      year: new Date().getFullYear().toString()
    },
    profile: options.paperSize === 'f4' ? 'suratDinasF4' : 'suratDinas',
    paperSize: options.paperSize || 'a4',
    orientation: options.orientation || 'portrait',
    customMargins: options.margins,
    letterData: {
      templateId: 'surat_polisher',
      schoolProfile: school,
      formData: {
        nomorSurat: smartNomor,
        judulSurat: title
      },
      options
    },
    blocks: [
      {
        type: 'letterKop',
        school
      },
      {
        type: 'letterTitle',
        title: title.toUpperCase(),
        nomor: `Nomor: ${smartNomor}`,
        metaRecipient: {
          metaRows: [
            { label: 'Sifat', value: parsed.sifat || 'Penting' },
            { label: 'Perihal', value: parsed.hal || parsed.perihal || 'Undangan Rapat' }
          ],
          recipientLines: [
            'Kepada',
            parsed.tujuan ? (parsed.tujuan.toLowerCase().startsWith('yth') ? parsed.tujuan : `Yth. ${parsed.tujuan}`) : 'Yth. ...............',
            '__________________',
            'Di',
            'Tempat'
          ]
        }
      }
    ]
  };

  if (Array.isArray(parsed.blocks) && parsed.blocks.length > 0) {
    for (const b of parsed.blocks) {
      if (b.type === 'paragraph' && b.text) {
        docModel.blocks.push({
          type: 'paragraph',
          style: 'letterBody',
          runs: [{ text: b.text.trim() }]
        });
      } else if (b.type === 'identityTable' && Array.isArray(b.rows)) {
        docModel.blocks.push({
          type: 'identityTable',
          rows: b.rows.map(r => ({
            label: r.label || '',
            value: (r.value || '').startsWith(':') ? r.value : `: ${r.value}`
          }))
        });
      }
    }
  }

  // Tanda tangan
  const sig = parsed.signer || {};
  docModel.blocks.push({
    type: 'letterSignature',
    city: sig.city || 'Garut',
    date: sig.date || parsed.tanggalSurat || dateStr,
    title: sig.title || school.principalTitle || 'Kepala Sekolah',
    name: sig.name || school.principalName || 'Dian Nurhadian, S.Pd.',
    nip: sig.nip && sig.nip !== '-' ? sig.nip : (school.principalNip || '-')
  });

  return docModel;
}

/**
 * Mengedit, menambahkan data, atau memperbaiki kata/kalimat pada dokumen yang sedang terbuka di lembar kerja
 * berdasarkan instruksi prompt pengguna (bukan membuat dari awal).
 * @param {object} currentSnapshot Snapshot isi kanvas saat ini
 * @param {string} userInstruction Instruksi penambahan/perbaikan dari pengguna
 * @param {object} [options]
 * @returns {Promise<object>} Hasil edit JSON terstruktur
 */
export async function editDocumentWithAi(currentSnapshot, userInstruction, options = {}) {
  const school = options.schoolProfile || store.getSchoolProfile();
  const dateStr = options.dateStr || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });

  const userPrompt = buildEditorUserPrompt({
    currentSnapshot,
    userInstruction,
    schoolName: school.schoolName,
    principalName: school.principalName,
    dateStr
  });

  const rawJson = await callAiApi({
    prompt: userPrompt,
    systemPrompt: EDITOR_SYSTEM_PROMPT,
    preferredApiId: options.preferredApiId,
    signal: options.signal
  });

  const cleaned = cleanJsonOutput(rawJson);
  let parsed;
  try {
    parsed = JSON.parse(cleaned);
  } catch (err) {
    throw new Error('AI merespons tetapi format JSON hasil edit tidak valid: ' + err.message);
  }

  return parsed;
}
