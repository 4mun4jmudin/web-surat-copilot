/**
 * tests/vaultAndGallery.test.js - Unit Test Gudang API & Gudang Kop Surat
 */

import test from 'node:test';
import assert from 'node:assert/strict';
import { 
  normalizeGeminiModel, 
  normalizeGroqModel,
  cleanJsonOutput, 
  extractSuggestedGeminiModel, 
  LATEST_GEMINI_DEFAULT_MODEL,
  LATEST_GROQ_DEFAULT_MODEL,
  STABLE_GEMINI_CANDIDATES
} from '../js/ai/client.js';
import {
  store,
  getBuiltinSystemKey,
  isUsingBuiltinSystemKey,
  getBuiltinFallbackConfigs
} from '../js/state.js';
import { DEFAULT_KOP_GALLERY } from '../js/letters/defaultKop.js';

test('Gudang API: Normalisasi model Gemini untuk mencegah 404 dan model omni limit 0', () => {
  // Model yang sering 404 atau omni berlimit 0 dialihkan otomatis ke model stabil tercepat (gemini-3.5-flash-lite)
  assert.equal(normalizeGeminiModel('gemini-1.5-pro'), 'gemini-3.5-flash-lite');
  assert.equal(normalizeGeminiModel('gemini-2.0-flash'), 'gemini-3.5-flash-lite');
  assert.equal(normalizeGeminiModel('gemini-pro'), 'gemini-3.5-flash-lite');
  assert.equal(normalizeGeminiModel('gemini-omni-flash-preview'), 'gemini-3.5-flash-lite');
  assert.equal(normalizeGeminiModel(''), 'gemini-3.5-flash-lite');

  // Model stabil GA dipertahankan
  assert.equal(normalizeGeminiModel('gemini-3.5-flash-lite'), 'gemini-3.5-flash-lite');
  assert.equal(normalizeGeminiModel('gemini-2.5-flash'), 'gemini-2.5-flash');
  assert.equal(normalizeGeminiModel('gemini-1.5-flash'), 'gemini-1.5-flash');
  assert.equal(normalizeGeminiModel('gemini-3.8-flash'), 'gemini-3.8-flash');
  assert.equal(normalizeGeminiModel('models/gemini-3.5-flash-lite'), 'gemini-3.5-flash-lite');

  // Pastikan kandidat pertama adalah Flash-Lite ber-RPM 15 (instan)
  assert.equal(STABLE_GEMINI_CANDIDATES[0], 'gemini-3.5-flash-lite');
  assert.equal(LATEST_GEMINI_DEFAULT_MODEL, 'gemini-3.5-flash-lite');
});

test('Gudang API: Normalisasi model Groq (llama-3.3-70b-versatile dialihkan ke openai/gpt-oss-120b)', () => {
  assert.equal(normalizeGroqModel('llama-3.3-70b-versatile'), 'openai/gpt-oss-120b');
  assert.equal(normalizeGroqModel('llama-3.1-8b-instant'), 'openai/gpt-oss-120b');
  assert.equal(normalizeGroqModel(''), 'openai/gpt-oss-120b');
  assert.equal(normalizeGroqModel('openai/gpt-oss-120b'), 'openai/gpt-oss-120b');
  assert.equal(normalizeGroqModel('qwen/qwen3.8-27b'), 'qwen/qwen3.8-27b');
});

test('Gudang API: Ekstraksi model rekomendasi dari pesan error Google (Self-Healing)', () => {
  const errMsg = 'This model models/gemini-2.0-flash is no longer available. Please update your code to use models/gemini-3.8-flash for the latest features and improvements.';
  const extracted = extractSuggestedGeminiModel(errMsg);
  assert.equal(extracted, 'gemini-3.8-flash');
});

test('Klien AI: Membersihkan pembungkus markdown JSON code fences', () => {
  const rawWithFence = '```json\n{"letterTitle": "SURAT TUGAS", "nomorSurat": "001/SMK/2026"}\n```';
  assert.equal(cleanJsonOutput(rawWithFence), '{"letterTitle": "SURAT TUGAS", "nomorSurat": "001/SMK/2026"}');

  const rawGenericFence = '```\n{"valid": true}\n```';
  assert.equal(cleanJsonOutput(rawGenericFence), '{"valid": true}');

  const plainJson = '{"letterTitle": "SURAT DINAS"}';
  assert.equal(cleanJsonOutput(plainJson), '{"letterTitle": "SURAT DINAS"}');
});

test('Gudang API: Menambah, mengaktifkan, memperbarui, dan menghapus kunci API multi-provider', () => {
  const initialCount = store.getApiVault().length;
  assert.ok(initialCount >= 4, 'Harus menyediakan preset awal (Gemini, Groq, DeepSeek, Qwen)');

  // 1. Tambah entri baru (Groq Cadangan)
  store.addApiVaultEntry({
    id: 'groq-cadangan-1',
    label: 'Groq Cadangan Kantor',
    provider: 'groq',
    model: 'llama-3.3-70b-versatile',
    apiKey: 'gsk_test1234567890',
    baseUrl: 'https://api.groq.com/openai/v1',
    isActive: true
  });

  const updatedVault = store.getApiVault();
  assert.equal(updatedVault.length, initialCount + 1);

  const activeApi = store.getActiveApi();
  assert.equal(activeApi.id, 'groq-cadangan-1');
  assert.equal(activeApi.provider, 'groq');

  // 2. Perbarui model atau label
  store.updateApiVaultEntry('groq-cadangan-1', {
    label: 'Groq Cadangan Kantor (Updated)'
  });
  const updatedEntry = store.getApiVault().find(v => v.id === 'groq-cadangan-1');
  assert.equal(updatedEntry.label, 'Groq Cadangan Kantor (Updated)');

  // 3. Hapus entri
  store.removeApiVaultEntry('groq-cadangan-1');
  const finalVault = store.getApiVault();
  assert.equal(finalVault.length, initialCount);
  assert.equal(finalVault.some(v => v.id === 'groq-cadangan-1'), false);
});

test('Gudang Kop Surat: Tersedia template bawaan dan mendukung penambahan kop kustom', () => {
  const gallery = store.getKopGallery();
  assert.ok(gallery.length >= 3, 'Koleksi kop surat bawaan harus ada minimal 3 template');

  const smkKop = gallery.find(k => k.id === 'kop-smk-dinas');
  assert.ok(smkKop, 'Template kop SMK Dinas harus tersedia');
  assert.ok(smkKop.dataUrl.startsWith('data:image/'), 'Data URL kop harus berformat gambar');

  // Tambah kop kustom
  const newKop = store.addKopGalleryEntry({
    name: 'Kop Yayasan Bina Bangsa',
    dataUrl: 'data:image/svg+xml;utf-8,<svg></svg>',
    category: 'Yayasan'
  });

  assert.ok(newKop.id);
  const afterAdd = store.getKopGallery();
  assert.equal(afterAdd.some(k => k.id === newKop.id), true);

  // Terapkan ke lembar kerja
  store.applyKopFromGallery(newKop.id);
  const currentSchool = store.getState().schoolProfile;
  assert.equal(currentSchool.kopImage, 'data:image/svg+xml;utf-8,<svg></svg>');

  // Hapus kop kustom
  store.removeKopGalleryEntry(newKop.id);
  const afterDelete = store.getKopGallery();
  assert.equal(afterDelete.some(k => k.id === newKop.id), false);
});

test('Zoom Kanvas Lembar Kerja: Membatasi nilai skala zoom antara 50% dan 200%', () => {
  store.setZoomLevel(125);
  assert.equal(store.getState().zoomLevel, 125);

  store.setZoomLevel(30); // di bawah batas minimal 50
  assert.equal(store.getState().zoomLevel, 50);

  store.setZoomLevel(250); // di atas batas maksimal 200
  assert.equal(store.getState().zoomLevel, 200);

  // Kembalikan ke 100%
  store.setZoomLevel(100);
  assert.equal(store.getState().zoomLevel, 100);
});

test('Klien AI: Penanganan pembatalan (signal abort) dan optimasi kecepatan Flash-Lite', async () => {
  const { callAiApi } = await import('../js/ai/client.js');

  const controller = new AbortController();
  controller.abort(); // Langsung batalkan

  await assert.rejects(
    async () => {
      await callAiApi({
        prompt: 'Halo',
        signal: controller.signal
      });
    },
    /dibatalkan/i,
    'Harus menolak dengan pesan pembatalan saat signal di-abort'
  );

  // Pastikan kandidat Gemini mengutamakan varian Flash Lite (15 RPM instan)
  assert.ok(STABLE_GEMINI_CANDIDATES[0].includes('lite'), 'Kandidat model pertama harus varian lite');
  assert.ok(STABLE_GEMINI_CANDIDATES[1].includes('lite'), 'Kandidat model kedua harus varian lite');
});

test('Gudang API: Kunci API AI bawaan sistem terpasang otomatis dan menjadi fallback saat pengguna belum mengisi API', async () => {
  const vault = store.getApiVault();
  const gemEntry = vault.find(v => v.id === 'gemini-auto');
  const groqEntry = vault.find(v => v.id === 'groq-fast');
  const deepseekEntry = vault.find(v => v.id === 'deepseek-v3');

  assert.ok(gemEntry?.apiKey.startsWith('AQ.'), 'Kunci bawaan Gemini harus terpasang otomatis');
  assert.ok(groqEntry?.apiKey.startsWith('gsk_'), 'Kunci bawaan Groq harus terpasang otomatis');
  assert.ok(deepseekEntry?.apiKey.startsWith('sk-'), 'Kunci bawaan DeepSeek harus terpasang otomatis');

  assert.equal(isUsingBuiltinSystemKey(gemEntry), true);
  assert.equal(isUsingBuiltinSystemKey({ id: 'gemini-auto', apiKey: 'AIzaCustomUserKey123' }), false);
  assert.equal(getBuiltinSystemKey('qwen-openrouter'), '');

  const fallbackConfigs = getBuiltinFallbackConfigs();
  assert.equal(fallbackConfigs.length, 3);
  assert.deepEqual(fallbackConfigs.map(c => c.id), ['gemini-auto', 'groq-fast', 'deepseek-v3']);

  // Uji skenario user memilih provider tanpa kunci (Qwen kosong) -> harus otomatis fallback ke API bawaan sistem
  const { callAiApi } = await import('../js/ai/client.js');
  const prevActiveId = store.getState().activeApiId;
  const origFetch = globalThis.fetch;
  const calledUrls = [];

  try {
    store.setActiveApi('qwen-openrouter');
    globalThis.fetch = async (url) => {
      calledUrls.push(String(url));
      return new Response(JSON.stringify({
        candidates: [{ content: { parts: [{ text: '{"ok":true}' }] } }]
      }), { status: 200, headers: { 'Content-Type': 'application/json' } });
    };

    const res = await callAiApi({ prompt: 'Tes', systemPrompt: 'Sys' });
    assert.equal(res, '{"ok":true}');
    assert.ok(
      calledUrls.some(u => u.includes('generativelanguage.googleapis.com') && u.includes(encodeURIComponent(gemEntry.apiKey).slice(0, 10)) || u.includes(gemEntry.apiKey)),
      'Harus otomatis menggunakan kunci bawaan sistem saat provider aktif belum memiliki API Key'
    );
  } finally {
    globalThis.fetch = origFetch;
    if (prevActiveId) store.setActiveApi(prevActiveId);
  }
});
