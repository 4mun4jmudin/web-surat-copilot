/**
 * aiView.js - Antarmuka Alur Kerja AI (Outline -> Subbab -> Document Model)
 * Dilengkapi streaming real-time, indikator progres bertahap, dan tombol batal.
 */

import { store } from '../state.js';
import { showToast } from './notification.js';
import { callAiWithValidation, callAiApi } from '../ai/client.js';
import {
  OUTLINE_SYSTEM_PROMPT,
  SECTION_SYSTEM_PROMPT,
  buildOutlineUserPrompt,
  buildSectionUserPrompt
} from '../ai/prompts.js';

let activeAbortController = null;
let currentOutline = null;

/**
 * Merender antarmuka AI Pipeline
 * @param {HTMLElement} container 
 */
export function renderAiView(container) {
  if (!container) return;

  container.innerHTML = `
    <div class="ai-view-container">
      <div class="card">
        <h3>Alur Kerja AI Document Builder</h3>
        <p class="text-muted">
          AI bekerja secara terstruktur: menentukan outline bab & subbab, mendapatkan persetujuan Anda, 
          lalu menulis isi setiap subbab secara berurutan dengan menjaga konteks kesinambungan.
        </p>

        <!-- LANGKAH 1: TOPIK DOKUMEN -->
        <div class="workflow-step active" id="step-outline-input">
          <h4>Langkah 1: Tentukan Topik & Profil Dokumen</h4>
          <div class="form-group">
            <label>Profil Dokumen:</label>
            <select id="ai-profile-select" class="form-control">
              <option value="pkl">Laporan PKL SMK (A4, TNR 12pt, Spasi 1.5, Margin 4-4-3-3)</option>
              <option value="skripsi">Skripsi / Tugas Akhir (A4, TNR 12pt, Spasi 2.0)</option>
              <option value="jurnal">Jurnal Ilmiah (A4, Arial 10pt, Spasi 1.0)</option>
            </select>
          </div>
          <div class="form-group">
            <label>Judul / Topik Utama Dokumen:</label>
            <input type="text" id="ai-topic-input" class="form-control" 
              placeholder="Contoh: Rancang Bangun Sistem Informasi Arsip Surat Digital Berbasis Web" 
              value="Implementasi Sistem Informasi Pengelolaan Persuratan Digital pada Divisi TI">
          </div>
          <div class="form-group">
            <label>Catatan / Batasan Tambahan (Opsional):</label>
            <textarea id="ai-notes-input" class="form-control" rows="2" 
              placeholder="Contoh: Fokus pada backend REST API, autentikasi JWT, dan pengujian unit."></textarea>
          </div>
          <button id="btn-generate-outline" class="btn btn-primary">
            🚀 Buat Kerangka Outline (AI)
          </button>
        </div>

        <!-- LANGKAH 2: EDIT & SETUJUI OUTLINE -->
        <div class="workflow-step hidden" id="step-outline-review">
          <h4>Langkah 2: Tinjau & Edit Kerangka Dokumen</h4>
          <p class="text-muted">Anda dapat menyesuaikan judul bab dan subbab sebelum AI mulai menulis konten lengkap.</p>
          <div id="outline-editor-area" class="outline-tree"></div>
          <div class="btn-group">
            <button id="btn-start-writing" class="btn btn-success">
              ✍ Mulai Tulis Semua Subbab
            </button>
            <button id="btn-cancel-outline" class="btn btn-secondary">
              Ubah Topik Awal
            </button>
          </div>
        </div>

        <!-- LANGKAH 3: PROGRES PENULISAN & STREAMING -->
        <div class="workflow-step hidden" id="step-writing-progress">
          <h4>Langkah 3: Proses Penulisan Otomatis Berjalan</h4>
          <div class="progress-bar-container">
            <div id="ai-progress-bar" class="progress-bar" style="width: 0%"></div>
          </div>
          <div class="progress-status" id="ai-progress-status">Menyiapkan model...</div>
          
          <div class="streaming-monitor">
            <h5>Pratinjau Streaming Generasi Konten:</h5>
            <pre id="ai-streaming-output" class="streaming-box"></pre>
          </div>

          <div class="btn-group">
            <button id="btn-abort-generation" class="btn btn-danger">
              ⛔ Batalkan Proses
            </button>
          </div>
        </div>

      </div>
    </div>
  `;

  // Bind Elemen
  const btnGenOutline = container.querySelector('#btn-generate-outline');
  const step1 = container.querySelector('#step-outline-input');
  const step2 = container.querySelector('#step-outline-review');
  const step3 = container.querySelector('#step-writing-progress');
  const outlineArea = container.querySelector('#outline-editor-area');
  const btnStartWriting = container.querySelector('#btn-start-writing');
  const btnCancelOutline = container.querySelector('#btn-cancel-outline');
  const btnAbort = container.querySelector('#btn-abort-generation');
  const progressBar = container.querySelector('#ai-progress-bar');
  const progressStatus = container.querySelector('#ai-progress-status');
  const streamingOutput = container.querySelector('#ai-streaming-output');

  // Buat Outline
  btnGenOutline.onclick = async () => {
    const topic = container.querySelector('#ai-topic-input').value.trim();
    const profile = container.querySelector('#ai-profile-select').value;
    const notes = container.querySelector('#ai-notes-input').value.trim();

    if (!topic) {
      showToast('Harap masukkan topik dokumen.', 'warning');
      return;
    }

    btnGenOutline.disabled = true;
    btnGenOutline.textContent = 'Menyusun kerangka (AI)...';

    try {
      const prompt = buildOutlineUserPrompt({ topic, profileName: profile, notes });
      const outline = await callAiWithValidation({
        prompt,
        systemPrompt: OUTLINE_SYSTEM_PROMPT,
        validator: (data) => {
          if (!data || !Array.isArray(data.chapters) || data.chapters.length === 0) {
            return { valid: false, errors: ['Outline harus memiliki array "chapters"'] };
          }
          return { valid: true, errors: [] };
        }
      });

      currentOutline = outline;
      renderOutlineEditor(outline, outlineArea);
      step1.classList.add('hidden');
      step2.classList.remove('hidden');
      showToast('Kerangka outline berhasil disusun. Silakan periksa.', 'success');

    } catch (err) {
      showToast(`Gagal menyusun outline: ${err.message}`, 'error');
    } finally {
      btnGenOutline.disabled = false;
      btnGenOutline.textContent = '🚀 Buat Kerangka Outline (AI)';
    }
  };

  btnCancelOutline.onclick = () => {
    step2.classList.add('hidden');
    step1.classList.remove('hidden');
  };

  btnAbort.onclick = () => {
    if (activeAbortController) {
      activeAbortController.abort();
      activeAbortController = null;
      showToast('Proses penulisan dibatalkan.', 'warning');
      step3.classList.add('hidden');
      step2.classList.remove('hidden');
    }
  };

  // Mulai Penulisan Tiap Subbab
  btnStartWriting.onclick = async () => {
    if (!currentOutline || !currentOutline.chapters) return;

    step2.classList.add('hidden');
    step3.classList.remove('hidden');

    activeAbortController = new AbortController();
    const signal = activeAbortController.signal;

    // Hitung total subbab
    let totalSubchapters = 0;
    currentOutline.chapters.forEach(ch => {
      totalSubchapters += (ch.subchapters || []).length;
    });

    let completedSubchapters = 0;
    let accumulatedSummary = `Dokumen: ${currentOutline.title}`;
    const generatedBlocks = [];

    // Tambahkan Judul & Bagian Awal
    generatedBlocks.push(
      { type: 'paragraph', runs: [{ text: currentOutline.title.toUpperCase(), bold: true }] },
      { type: 'sectionBreak', numberFormat: 'roman', restart: true },
      { type: 'heading', level: 1, text: 'KATA PENGANTAR' },
      { type: 'paragraph', runs: [{ text: 'Puji dan syukur kami panjatkan ke hadirat Tuhan Yang Maha Esa atas terselesaikannya laporan ini.' }] },
      { type: 'pageBreak' },
      { type: 'toc', title: 'DAFTAR ISI' },
      { type: 'pageBreak' },
      { type: 'sectionBreak', numberFormat: 'decimal', restart: true }
    );

    try {
      for (const chapter of currentOutline.chapters) {
        if (signal.aborted) break;

        // Heading Bab Utama
        generatedBlocks.push({
          type: 'heading',
          level: 1,
          text: chapter.title
        });

        for (const sub of (chapter.subchapters || [])) {
          if (signal.aborted) break;

          progressStatus.textContent = `Menulis Bab "${chapter.title}" - Subbab: "${sub}"...`;
          streamingOutput.textContent = `[Mulai menulis subbab: ${sub}]...\n`;

          const userPrompt = buildSectionUserPrompt({
            docTitle: currentOutline.title,
            chapterTitle: chapter.title,
            subchapterTitle: sub,
            contextSummary: accumulatedSummary
          });

          const sectionResult = await callAiWithValidation({
            prompt: userPrompt,
            systemPrompt: SECTION_SYSTEM_PROMPT,
            onChunk: (allText) => {
              streamingOutput.textContent = allText.slice(-600);
            },
            signal,
            validator: (data) => {
              if (!data || !Array.isArray(data.blocks)) {
                return { valid: false, errors: ['Hasil harus berupa objek dengan properti "blocks"'] };
              }
              return { valid: true, errors: [] };
            }
          });

          // Gabungkan blok subbab
          generatedBlocks.push(...sectionResult.blocks);

          completedSubchapters++;
          const pct = Math.round((completedSubchapters / totalSubchapters) * 100);
          progressBar.style.width = `${pct}%`;

          accumulatedSummary += `\n- Telah selesai: ${sub}`;
        }
      }

      // Tambahkan Daftar Pustaka Penutup
      generatedBlocks.push(
        { type: 'pageBreak' },
        { type: 'heading', level: 1, text: 'DAFTAR PUSTAKA' },
        { type: 'reference', id: 'ref_ai_1', text: 'Kementerian Pendidikan dan Kebudayaan. (2025). Panduan Penulisan Laporan Akademik dan Tugas Akhir. Jakarta.' }
      );

      // Simpan ke State Store Dokumen
      const newDoc = {
        meta: {
          title: currentOutline.title,
          author: 'AI Document Builder',
          year: new Date().getFullYear().toString()
        },
        profile: container.querySelector('#ai-profile-select').value,
        blocks: generatedBlocks
      };

      store.setState({ document: newDoc }, true);
      showToast('Seluruh bab dokumen berhasil disusun oleh AI!', 'success');

      // Pindahkan tampilan ke Editor & Pratinjau
      const tabEditorBtn = document.querySelector('[data-tab="editor"]');
      if (tabEditorBtn) tabEditorBtn.click();

    } catch (err) {
      if (err.name === 'AbortError') {
        showToast('Generasi AI dibatalkan.', 'info');
      } else {
        showToast(`Error generasi: ${err.message}`, 'error');
      }
    } finally {
      activeAbortController = null;
    }
  };
}

function renderOutlineEditor(outline, container) {
  container.innerHTML = `
    <div class="outline-title-edit">
      <label>Judul Dokumen:</label>
      <input type="text" class="form-control" id="outline-title-input" value="${escapeHtml(outline.title || '')}">
    </div>
    <div class="chapters-list">
      ${(outline.chapters || []).map((ch, chIdx) => `
        <div class="chapter-item" data-chapter="${chIdx}">
          <div class="chapter-header">
            <strong>BAB ${ch.number || chIdx + 1}:</strong>
            <input type="text" class="form-control ch-title-input" value="${escapeHtml(ch.title)}">
          </div>
          <div class="subchapters-list">
            ${(ch.subchapters || []).map((sub, sIdx) => `
              <div class="subchapter-item">
                <span>↳</span>
                <input type="text" class="form-control sub-title-input" value="${escapeHtml(sub)}">
              </div>
            `).join('')}
          </div>
        </div>
      `).join('')}
    </div>
  `;

  // Update outline saat input berubah
  const titleInput = container.querySelector('#outline-title-input');
  titleInput.onblur = () => { outline.title = titleInput.value; };

  container.querySelectorAll('.chapter-item').forEach((chEl, chIdx) => {
    const chTitle = chEl.querySelector('.ch-title-input');
    chTitle.onblur = () => { outline.chapters[chIdx].title = chTitle.value; };

    chEl.querySelectorAll('.sub-title-input').forEach((subEl, sIdx) => {
      subEl.onblur = () => { outline.chapters[chIdx].subchapters[sIdx] = subEl.value; };
    });
  });
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

