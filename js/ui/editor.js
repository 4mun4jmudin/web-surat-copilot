/**
 * editor.js - Editor Blok Visual untuk Document Model
 * Mendukung tambah, hapus, susun ulang (reorder), undo/redo, dan revisi AI per blok.
 */

import { store } from '../state.js';
import { showToast } from './notification.js';
import { callAiWithValidation } from '../ai/client.js';
import { REWRITE_SYSTEM_PROMPT, buildRewriteUserPrompt } from '../ai/prompts.js';

/**
 * Merender daftar blok ke kontainer editor
 * @param {HTMLElement} container 
 */
export function renderBlockEditor(container) {
  if (!container) return;
  const state = store.getState();
  const blocks = state.document.blocks || [];

  container.innerHTML = `
    <div class="editor-header">
      <div class="editor-title-group">
        <h3>Editor Blok Dokumen</h3>
        <span class="block-count-badge">${blocks.length} Blok</span>
      </div>
      <div class="editor-actions">
        <button id="btn-undo" class="btn btn-sm btn-secondary" ${store.canUndo() ? '' : 'disabled'}>
          ↶ Undo
        </button>
        <button id="btn-redo" class="btn btn-sm btn-secondary" ${store.canRedo() ? '' : 'disabled'}>
          ↷ Redo
        </button>
        <div class="dropdown">
          <button id="btn-add-block" class="btn btn-sm btn-primary">+ Tambah Blok</button>
          <div id="add-block-menu" class="dropdown-menu hidden">
            <button data-type="heading">Heading (Judul/Bab)</button>
            <button data-type="paragraph">Paragraf Teks</button>
            <button data-type="list">Daftar (List)</button>
            <button data-type="table">Tabel</button>
            <button data-type="figure">Gambar / Ilustrasi</button>
            <button data-type="pageBreak">Pemisah Halaman (Page Break)</button>
            <button data-type="sectionBreak">Pemisah Bagian (Section Break)</button>
          </div>
        </div>
      </div>
    </div>
    <div class="blocks-container" id="blocks-list"></div>
  `;

  // Bind Undo & Redo
  container.querySelector('#btn-undo').onclick = () => store.undo();
  container.querySelector('#btn-redo').onclick = () => store.redo();

  // Dropdown tambah blok
  const addBtn = container.querySelector('#btn-add-block');
  const addMenu = container.querySelector('#add-block-menu');
  addBtn.onclick = (e) => {
    e.stopPropagation();
    addMenu.classList.toggle('hidden');
  };

  document.addEventListener('click', () => {
    if (addMenu) addMenu.classList.add('hidden');
  });

  addMenu.querySelectorAll('button').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      const type = btn.dataset.type;
      addNewBlock(type);
      addMenu.classList.add('hidden');
    };
  });

  // Render cards
  const listEl = container.querySelector('#blocks-list');
  blocks.forEach((block, index) => {
    const card = createBlockCard(block, index, blocks.length);
    listEl.appendChild(card);
  });
}

function createBlockCard(block, index, total) {
  const card = document.createElement('div');
  card.className = `block-card block-type-${block.type}`;
  card.dataset.index = String(index);

  let typeBadge = block.type;
  if (block.type === 'heading') typeBadge = `Heading ${block.level}`;
  if (block.type === 'sectionBreak') typeBadge = `Section (${block.numberFormat || 'decimal'})`;

  card.innerHTML = `
    <div class="block-card-header">
      <span class="block-badge">${typeBadge}</span>
      <span class="block-number">#${index + 1}</span>
      <div class="block-card-tools">
        <button class="btn-icon btn-move-up" title="Geser ke Atas" ${index === 0 ? 'disabled' : ''}>▲</button>
        <button class="btn-icon btn-move-down" title="Geser ke Bawah" ${index === total - 1 ? 'disabled' : ''}>▼</button>
        ${block.type === 'paragraph' ? '<button class="btn-icon btn-ai-rewrite" title="Revisi / Tulis Ulang dengan AI">✨ AI</button>' : ''}
        <button class="btn-icon btn-delete-block" title="Hapus Blok">🗑</button>
      </div>
    </div>
    <div class="block-card-body"></div>
  `;

  const bodyEl = card.querySelector('.block-card-body');
  renderBlockBody(block, index, bodyEl);

  // Event handlers
  card.querySelector('.btn-move-up').onclick = () => store.moveBlock(index, index - 1);
  card.querySelector('.btn-move-down').onclick = () => store.moveBlock(index, index + 1);
  card.querySelector('.btn-delete-block').onclick = () => {
    if (confirm(`Hapus blok #${index + 1} (${typeBadge})?`)) {
      store.removeBlock(index);
    }
  };

  const aiBtn = card.querySelector('.btn-ai-rewrite');
  if (aiBtn) {
    aiBtn.onclick = () => openAiRewriteDialog(block, index);
  }

  return card;
}

function renderBlockBody(block, index, container) {
  switch (block.type) {
    case 'heading': {
      container.innerHTML = `
        <div class="form-row">
          <select class="input-heading-level">
            <option value="1" ${block.level === 1 ? 'selected' : ''}>Heading 1 (Bab)</option>
            <option value="2" ${block.level === 2 ? 'selected' : ''}>Heading 2 (Subbab)</option>
            <option value="3" ${block.level === 3 ? 'selected' : ''}>Heading 3 (Sub-subbab)</option>
          </select>
          <input type="text" class="input-heading-text" value="${escapeHtml(block.text || '')}" placeholder="Judul Heading...">
        </div>
      `;
      const lvlSelect = container.querySelector('.input-heading-level');
      const txtInput = container.querySelector('.input-heading-text');
      
      const update = () => {
        store.updateBlock(index, {
          ...block,
          level: parseInt(lvlSelect.value, 10),
          text: txtInput.value
        });
      };
      lvlSelect.onchange = update;
      txtInput.onblur = update;
      break;
    }

    case 'paragraph': {
      const text = block.runs ? block.runs.map(r => r.text).join('') : (block.text || '');
      container.innerHTML = `
        <textarea class="textarea-paragraph" rows="3" placeholder="Tulis paragraf...">${escapeHtml(text)}</textarea>
      `;
      const txt = container.querySelector('.textarea-paragraph');
      txt.onblur = () => {
        if (txt.value !== text) {
          store.updateBlock(index, {
            ...block,
            runs: [{ text: txt.value }]
          });
        }
      };
      break;
    }

    case 'list': {
      const itemsText = (block.items || []).join('\n');
      container.innerHTML = `
        <div class="form-row">
          <label>
            <input type="checkbox" class="checkbox-ordered" ${block.ordered ? 'checked' : ''}> Penomoran Berurutan (1, 2, 3)
          </label>
        </div>
        <textarea class="textarea-list-items" rows="3" placeholder="1 baris per butir">${escapeHtml(itemsText)}</textarea>
      `;
      const chk = container.querySelector('.checkbox-ordered');
      const txt = container.querySelector('.textarea-list-items');

      const update = () => {
        const items = txt.value.split('\n').filter(i => i.trim() !== '');
        store.updateBlock(index, {
          ...block,
          ordered: chk.checked,
          items
        });
      };
      chk.onchange = update;
      txt.onblur = update;
      break;
    }

    case 'table': {
      container.innerHTML = `
        <input type="text" class="input-table-caption" value="${escapeHtml(block.caption || '')}" placeholder="Caption Tabel...">
        <div class="table-preview-summary">
          <em>Tabel memiliki ${(block.rows || []).length} baris data.</em>
        </div>
      `;
      const capInput = container.querySelector('.input-table-caption');
      capInput.onblur = () => {
        store.updateBlock(index, { ...block, caption: capInput.value });
      };
      break;
    }

    case 'figure': {
      container.innerHTML = `
        <input type="text" class="input-figure-caption" value="${escapeHtml(block.caption || '')}" placeholder="Caption Gambar...">
      `;
      const capInput = container.querySelector('.input-figure-caption');
      capInput.onblur = () => {
        store.updateBlock(index, { ...block, caption: capInput.value });
      };
      break;
    }

    case 'sectionBreak': {
      container.innerHTML = `
        <div class="form-row">
          <label>Format Nomor Halaman:</label>
          <select class="select-sec-format">
            <option value="none" ${block.numberFormat === 'none' ? 'selected' : ''}>Tanpa Nomor (Cover)</option>
            <option value="roman" ${block.numberFormat === 'roman' ? 'selected' : ''}>Romawi Kecil (i, ii, iii)</option>
            <option value="decimal" ${block.numberFormat === 'decimal' ? 'selected' : ''}>Angka Arab (1, 2, 3)</option>
          </select>
          <label>
            <input type="checkbox" class="chk-sec-restart" ${block.restart !== false ? 'checked' : ''}> Mulai Ulang dari 1
          </label>
        </div>
      `;
      const fmtSelect = container.querySelector('.select-sec-format');
      const restartChk = container.querySelector('.chk-sec-restart');
      const update = () => {
        store.updateBlock(index, {
          ...block,
          numberFormat: fmtSelect.value,
          restart: restartChk.checked
        });
      };
      fmtSelect.onchange = update;
      restartChk.onchange = update;
      break;
    }

    default:
      container.innerHTML = `<span class="text-muted">Blok ${block.type}</span>`;
  }
}

function addNewBlock(type) {
  let newBlock = { type };
  switch (type) {
    case 'heading':
      newBlock = { type: 'heading', level: 2, text: 'Subbab Baru' };
      break;
    case 'paragraph':
      newBlock = { type: 'paragraph', runs: [{ text: 'Tulis paragraf baru di sini.' }] };
      break;
    case 'list':
      newBlock = { type: 'list', ordered: false, items: ['Butir baru'] };
      break;
    case 'table':
      newBlock = {
        type: 'table',
        caption: 'Tabel Baru',
        headerRow: true,
        rows: [['Kolom A', 'Kolom B'], ['Data 1', 'Data 2']]
      };
      break;
    case 'figure':
      newBlock = { type: 'figure', caption: 'Ilustrasi Baru', src: 'new_img' };
      break;
    case 'pageBreak':
      newBlock = { type: 'pageBreak' };
      break;
    case 'sectionBreak':
      newBlock = { type: 'sectionBreak', numberFormat: 'decimal', restart: false };
      break;
  }
  store.addBlock(newBlock);
  showToast('Blok baru berhasil ditambahkan.', 'success');
}

/**
 * Dialog AI Rewrite untuk satu blok paragraf
 */
async function openAiRewriteDialog(block, index) {
  const action = prompt(
    `Pilih tindakan AI untuk blok #${index + 1}:\n1. perbaiki (Tata bahasa baku)\n2. perpanjang (Kembangkan ide)\n3. persingkat (Ringkas padat)\n4. formalkan (Gaya akademis tinggi)\n\nKetik salah satu aksi di atas:`,
    'formalkan'
  );

  if (!action) return;

  showToast(`Sedang memproses AI Rewrite (${action})...`, 'info');

  try {
    const userPrompt = buildRewriteUserPrompt({
      action,
      currentBlock: block
    });

    const updatedBlock = await callAiWithValidation({
      prompt: userPrompt,
      systemPrompt: REWRITE_SYSTEM_PROMPT,
      validator: (data) => {
        if (!data || data.type !== 'paragraph') {
          return { valid: false, errors: ['Output harus berupa blok paragraph valid.'] };
        }
        return { valid: true, errors: [] };
      }
    });

    store.updateBlock(index, updatedBlock);
    showToast(`Blok #${index + 1} berhasil diperbarui oleh AI!`, 'success');
  } catch (err) {
    showToast(`Gagal merevisi blok: ${err.message}`, 'error');
  }
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

