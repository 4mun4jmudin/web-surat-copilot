/**
 * wordCanvas.js - Antarmuka Lembar Kerja Microsoft Word Asli & AI Copilot Persuratan
 * Fitur:
 * 1. Ribbon 2-Row lengkap standar Microsoft Word (Font & Paragraph lengkap sesuai Screenshot).
 * 2. Pratinjau skala nyata A4 & F4 dengan pengaturan margin tanpa clipping / z-index tepat.
 * 3. Zoom In/Out khusus kanvas lembar kerja, menjaga ribbon dan sidebar tetap tetap.
 * 4. Gudang Kop Surat (Galeri Kop) dengan pratinjau thumbnail dan pemilihan 1-klik.
 * 5. Integrasi Gudang API di AI Copilot dengan pemilih model real-time.
 */

import { store } from '../state.js';
import { renderDocx } from '../docx/render.js';
import { parseDocxForEditing, applyCanvasEditsToDocx, convertDocxToDocumentModel } from '../docx/patch.js';
import { polishMessyDocumentWithAi } from '../ai/client.js';
import { LETTER_TYPE_CODES, formatSmartNomorSurat } from '../letters/numbering.js';
import { downloadFile } from '../utils/download.js';
import { showToast } from './notification.js';
import { DEFAULT_KOP_DATA_URL } from '../letters/defaultKop.js';
import {
  ICON_GALLERY, ICON_EYE, ICON_EYE_OFF, ICON_UPLOAD,
  ICON_SAVE, ICON_ARCHIVE, ICON_DOC_UPLOAD, ICON_CONVERT,
  ICON_SPARKLE, ICON_WAND, ICON_DOWNLOAD, ICON_PRINTER,
  ICON_SETTINGS, ICON_RULER, ICON_SCHOOL, ICON_SHADING, ICON_DOC
} from './icons.js';

let syncTimer = null;

export function renderWordCanvas(container) {
  if (!container) return;

  // Pastikan scroll jendela utama selalu di titik (0,0) agar bagian kiri header & ribbon tidak pernah tergeser keluar layar
  if ('scrollRestoration' in history) {
    history.scrollRestoration = 'manual';
  }
  window.scrollTo(0, 0);
  document.documentElement.scrollLeft = 0;
  document.body.scrollLeft = 0;

  const state = store.getState();
  const doc = state.document;
  const paperSize = state.letterOptions?.paperSize || 'a4';
  const orientation = state.letterOptions?.orientation || 'portrait';
  const customMargins = state.letterOptions?.margins || { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 };
  const fontFamily = state.letterOptions?.fontFamily || 'Times New Roman';
  const fontSize = state.letterOptions?.fontSize || '12';
  const lineSpacing = state.letterOptions?.lineSpacing || '1.15';
  const showKop = state.letterOptions?.showKop !== false;
  const zoomLevel = state.zoomLevel || 100;
  const apiVault = store.getApiVault();
  const activeApi = store.getActiveApi();
  const vaultLetters = store.getLetterVault();
  const aiRefs = store.getAiReferences();
  const activeDocx = state.activeUploadedDocx;

  const isDocxPreviewMode = (state.docxViewMode === 'native_preview');
  let structured = activeDocx?.structured;
  if (activeDocx && !structured && activeDocx.paragraphs) {
    const paras = activeDocx.paragraphs;
    let titleP = paras.find(p => p.role === 'title') || paras.find(p => p.originalText && p.originalText.toUpperCase().startsWith('SURAT ') && p.originalText.length <= 75);
    let nomorP = paras.find(p => p.role === 'nomor') || paras.find(p => /^(nomor|no)\s*:/i.test(p.originalText || ''));
    structured = {
      isAlHawari: paras.some(p => /al-hawari/i.test(p.originalText || '')),
      title: { text: titleP ? titleP.currentText || titleP.originalText : activeDocx.parsedMeta?.title || 'SURAT RESMI', index: titleP ? titleP.index : -1 },
      nomor: { text: nomorP ? nomorP.currentText || nomorP.originalText : activeDocx.parsedMeta?.nomor || '', index: nomorP ? nomorP.index : -1 },
      identityRows: [],
      bodyParagraphs: paras.filter(p => p !== titleP && p !== nomorP && p.role !== 'kop').map(p => ({
        index: p.index,
        text: p.currentText || p.originalText,
        alignment: p.alignment || 'justify',
        isBold: p.isBold
      })),
      closing: null,
      signature: {
        date: activeDocx.parsedMeta?.date || 'Garut, 14 November 2026',
        dateIndex: -1,
        title: 'Kepala Sekolah',
        titleIndex: -1,
        signer: activeDocx.parsedMeta?.signer || 'Dian Nurhadian, S.Pd.',
        signerIndex: -1,
        nip: '-',
        nipIndex: -1
      }
    };
  }

  const docxTitleText = structured?.title?.text || activeDocx?.parsedMeta?.title || 'SURAT RESMI';
  const docxTitleIdx = structured?.title?.index ?? -1;
  const docxNomorText = structured?.nomor?.text || activeDocx?.parsedMeta?.nomor || '';
  const docxNomorIdx = structured?.nomor?.index ?? -1;
  const docxBodyList = structured?.bodyParagraphs || [];
  const docxIdentityRows = structured?.identityRows || [];
  const docxClosing = structured?.closing;
  const docxSig = structured?.signature || {
    date: 'Garut, 14 November 2026',
    dateIndex: -1,
    title: 'Kepala Sekolah',
    titleIndex: -1,
    signer: 'Dian Nurhadian, S.Pd.',
    signerIndex: -1,
    nip: '-',
    nipIndex: -1
  };

  const kopBlock = (doc.blocks || []).find(b => b.type === 'letterKop');
  const kopImageUrl = kopBlock?.image || state.schoolProfile?.kopImage || DEFAULT_KOP_DATA_URL;

  // Temukan blok-blok surat
  const titleBlock = (doc.blocks || []).find(b => b.type === 'letterTitle') || {
    title: 'SURAT KETERANGAN SISWA AKTIF',
    nomor: 'Nomor: 421.3 / 085 / SMKITA / XI / 2026'
  };
  const standaloneMetaBlock = (doc.blocks || []).find(b => b.type === 'letterMetaRecipient');
  const rawMetaRecipient = titleBlock.metaRecipient || standaloneMetaBlock || (doc.letterData?.metaRecipientDeleted ? { isDeleted: true } : {
    metaRows: [
      { label: 'Sifat', value: 'Penting' },
      { label: 'Perihal', value: 'Undangan Rapat' }
    ],
    recipientLines: [
      'Kepada',
      'Yth. ...............',
      '__________________',
      'Di',
      'Tempat'
    ]
  });
  const isMetaDeleted = Boolean(rawMetaRecipient?.isDeleted || doc.letterData?.metaRecipientDeleted);

  const docxMetaRows = (structured?.metaRows && structured.metaRows.length > 0)
    ? structured.metaRows
    : [
        { label: 'Sifat', value: 'Penting' },
        { label: 'Perihal', value: 'Undangan Rapat' }
      ];
  const docxRecipientLines = (structured?.recipientLines && structured.recipientLines.length > 0)
    ? structured.recipientLines
    : [
        'Kepada',
        'Yth. ...............',
        '__________________',
        'Di',
        'Tempat'
      ];

  const identityBlock = (doc.blocks || []).find(b => b.type === 'identityTable');
  const sigBlock = (doc.blocks || []).find(b => b.type === 'letterSignature') || {
    city: 'Garut',
    date: '14 November 2026',
    title: 'Kepala Sekolah',
    name: 'Dian Nurhadian, S.Pd.',
    nip: '-'
  };

  const bodyParagraphs = (doc.blocks || []).filter(b => b.type === 'paragraph');

  container.innerHTML = `
    <div class="word-workspace">
      
      <!-- 1. RIBBON TOOLBAR MICROSOFT WORD (2-ROW STANDARD) -->
      <div class="word-ribbon" style="display: flex; flex-wrap: wrap; align-items: stretch; justify-content: flex-start; width: 100%; padding: 4px 12px; gap: 6px; box-sizing: border-box; overflow: visible;">
        
        <!-- KELOMPOK: HALAMAN & KERTAS (UKURAN, ORIENTASI, MARGIN, FORMAT) -->
        <div class="ribbon-group" style="display: flex; flex-direction: column; justify-content: space-between; flex-shrink: 0; padding-right: 6px; border-right: 1px solid #cbd5e1;">
          <div class="ribbon-rows-wrapper">
            <div class="ribbon-row">
              <select id="ribbon-paper-size" class="ribbon-select" style="width: 64px;" title="Ukuran Kertas Baku">
                <option value="a4" ${paperSize === 'a4' ? 'selected' : ''}>A4</option>
                <option value="f4" ${paperSize === 'f4' ? 'selected' : ''}>F4</option>
                <option value="letter" ${paperSize === 'letter' ? 'selected' : ''}>Letter</option>
                <option value="legal" ${paperSize === 'legal' ? 'selected' : ''}>Legal</option>
              </select>
              <select id="ribbon-orientation" class="ribbon-select" style="width: 68px;" title="Orientasi Kertas">
                <option value="portrait" ${orientation === 'portrait' ? 'selected' : ''}>Potret</option>
                <option value="landscape" ${orientation === 'landscape' ? 'selected' : ''}>Lanskap</option>
              </select>
            </div>
            <div class="ribbon-row">
              <div class="dropdown">
                <button id="btn-margin-menu" class="ribbon-btn ribbon-btn-compact" type="button" title="Atur batas margin kertas">
                  ${ICON_RULER} <span>Margin ▾</span>
                </button>
                <div id="margin-dropdown-menu" class="dropdown-menu hidden" style="width: 290px; padding: 12px;">
                  <strong style="font-size: 12px; color: var(--dark);">Pengaturan Margin Dokumen:</strong>
                  <div style="display: flex; flex-direction: column; gap: 4px; margin: 8px 0;">
                    <button type="button" class="btn btn-sm btn-secondary btn-preset-margin" data-top="2.0" data-bottom="2.0" data-left="2.5" data-right="2.0">Standar Dinas (Kiri 2.5, Atas/Kanan/Bawah 2.0)</button>
                    <button type="button" class="btn btn-sm btn-secondary btn-preset-margin" data-top="2.54" data-bottom="2.54" data-left="2.54" data-right="2.54">Normal Word (2.54 cm di semua sisi)</button>
                    <button type="button" class="btn btn-sm btn-secondary btn-preset-margin" data-top="1.27" data-bottom="1.27" data-left="1.27" data-right="1.27">Sempit / Padat (1.27 cm)</button>
                  </div>
                  <div class="form-row" style="margin-top: 4px;">
                    <div class="col"><label style="font-size:10.5px;">Kiri (cm):</label><input type="number" step="0.1" id="input-m-left" class="form-control" value="${customMargins.left}"></div>
                    <div class="col"><label style="font-size:10.5px;">Kanan (cm):</label><input type="number" step="0.1" id="input-m-right" class="form-control" value="${customMargins.right}"></div>
                  </div>
                  <div class="form-row" style="margin-top: 4px;">
                    <div class="col"><label style="font-size:10.5px;">Atas (cm):</label><input type="number" step="0.1" id="input-m-top" class="form-control" value="${customMargins.top}"></div>
                    <div class="col"><label style="font-size:10.5px;">Bawah (cm):</label><input type="number" step="0.1" id="input-m-bottom" class="form-control" value="${customMargins.bottom}"></div>
                  </div>
                  <button type="button" id="btn-apply-custom-margin" class="btn btn-sm btn-primary" style="margin-top: 8px; width: 100%;">Terapkan Margin</button>
                </div>
              </div>
              <button id="btn-open-doc-settings" class="ribbon-btn ribbon-btn-compact" title="Buka Pengaturan Lengkap Dokumen (Ukuran, Margin, Font, Kop, dll)">
                ${ICON_SETTINGS} <span>Format</span>
              </button>
            </div>
          </div>
          <span class="ribbon-label">Halaman</span>
        </div>

        <!-- KELOMPOK: FONT (SESUAI SCREENSHOT WORD) -->
        <div class="ribbon-group">
          <div class="ribbon-rows-wrapper">
            <!-- Font Row 1: Font Family, Font Size, Grow/Shrink, Case, Clear -->
            <div class="ribbon-row">
              <select id="cmd-font-family" class="ribbon-select" style="width: 108px;" title="Jenis Huruf">
                <option value="Times New Roman" selected>Times New Roman</option>
                <option value="Arial">Arial</option>
                <option value="Calibri">Calibri</option>
                <option value="Segoe UI">Segoe UI</option>
                <option value="Georgia">Georgia</option>
                <option value="Courier New">Courier New</option>
              </select>

              <select id="cmd-font-size" class="ribbon-select" style="width: 42px;" title="Ukuran Huruf">
                <option value="8">8</option>
                <option value="9">9</option>
                <option value="10">10</option>
                <option value="11">11</option>
                <option value="12" selected>12</option>
                <option value="14">14</option>
                <option value="16">16</option>
                <option value="18">18</option>
                <option value="20">20</option>
                <option value="24">24</option>
                <option value="28">28</option>
                <option value="36">36</option>
              </select>

              <button class="ribbon-btn-icon" id="cmd-grow-font" title="Perbesar Huruf (Ctrl+Shift+>)">A<sup>▲</sup></button>
              <button class="ribbon-btn-icon" id="cmd-shrink-font" title="Perkecil Huruf (Ctrl+Shift+<)">A<sub>▼</sub></button>
              <button class="ribbon-btn-icon" id="cmd-change-case" title="Ubah Kapitalisasi (Shift+F3)">Aa</button>
              <button class="ribbon-btn-icon" id="cmd-clear-format" title="Hapus Semua Pemformatan">A🧹</button>
            </div>

            <!-- Font Row 2: Bold, Italic, Underline, Strikethrough, Sub, Sup, Highlight, Color -->
            <div class="ribbon-row">
              <button class="ribbon-btn-icon" id="cmd-bold" title="Tebal (Ctrl+B)"><b>B</b></button>
              <button class="ribbon-btn-icon" id="cmd-italic" title="Miring (Ctrl+I)"><i>I</i></button>
              <button class="ribbon-btn-icon" id="cmd-underline" title="Garis Bawah (Ctrl+U)"><u>U</u></button>
              <button class="ribbon-btn-icon" id="cmd-strike" title="Coret Tulisan (Strikethrough)"><s>abc</s></button>
              <button class="ribbon-btn-icon" id="cmd-subscript" title="Subskrip (x₂)">x₂</button>
              <button class="ribbon-btn-icon" id="cmd-superscript" title="Superskrip (x²)">x²</button>
              <span class="ribbon-divider"></span>

              <!-- Highlight Color Picker -->
              <div class="ribbon-btn-color-wrap" title="Warna Sorotan Teks (Stabilo)">
                <button class="ribbon-btn-icon" id="btn-highlight-visual">
                  <span style="background: #fef08a; padding: 0 2px; font-weight: bold; font-size: 11px;">ab</span>
                </button>
                <input type="color" id="picker-highlight-color" value="#fef08a" class="ribbon-color-picker-input">
              </div>

              <!-- Font Text Color Picker -->
              <div class="ribbon-btn-color-wrap" title="Warna Huruf (Font Color)">
                <button class="ribbon-btn-icon" id="btn-color-visual">
                  <span style="border-bottom: 3px solid #dc2626; font-weight: bold; font-size: 11px;">A</span>
                </button>
                <input type="color" id="picker-font-color" value="#dc2626" class="ribbon-color-picker-input">
              </div>
            </div>
          </div>
          <span class="ribbon-label">Font</span>
        </div>

        <!-- KELOMPOK: PARAGRAPH (SESUAI SCREENSHOT WORD) -->
        <div class="ribbon-group">
          <div class="ribbon-rows-wrapper">
            <!-- Paragraph Row 1: Bullets, Numbering, Indents, Formatting Marks -->
            <div class="ribbon-row">
              <button class="ribbon-btn-icon" id="cmd-bullet-list" title="Daftar Poin (Bullets)">•≡</button>
              <button class="ribbon-btn-icon" id="cmd-numbered-list" title="Daftar Bernomor (Numbering)">1≡</button>
              <button class="ribbon-btn-icon" id="cmd-multilevel-list" title="Daftar Bertingkat">1.1</button>
              <span class="ribbon-divider"></span>
              <button class="ribbon-btn-icon" id="cmd-outdent" title="Kurangi Inden (Geser Kiri)">⇤</button>
              <button class="ribbon-btn-icon" id="cmd-indent" title="Tambah Inden (Geser Kanan)">⇥</button>
              <button class="ribbon-btn-icon" id="cmd-toggle-pilcrow" title="Tampilkan/Sembunyikan Tanda Paragraf">¶</button>
            </div>

            <!-- Paragraph Row 2: Align Left, Center, Right, Justify, Line Spacing, Shading -->
            <div class="ribbon-row">
              <button class="ribbon-btn-icon" id="cmd-align-left" title="Rata Kiri (Ctrl+L)">≡</button>
              <button class="ribbon-btn-icon" id="cmd-align-center" title="Rata Tengah (Ctrl+E)">≍</button>
              <button class="ribbon-btn-icon" id="cmd-align-right" title="Rata Kanan (Ctrl+R)">≣</button>
              <button class="ribbon-btn-icon active" id="cmd-align-justify" title="Rata Kiri-Kanan / Justify (Ctrl+J)">≋</button>
              <span class="ribbon-divider"></span>
              
              <select id="ribbon-line-spacing" class="ribbon-select" style="width: 76px;" title="Jarak Spasi Baris">
                <option value="1.0">Spasi 1.0</option>
                <option value="1.15" selected>Spasi 1.15</option>
                <option value="1.5">Spasi 1.5</option>
                <option value="2.0">Spasi 2.0</option>
              </select>

              <!-- Shading Background Fill -->
              <div class="ribbon-btn-color-wrap" title="Warna Latar Paragraf (Shading)">
                <button class="ribbon-btn-icon" id="btn-shading-visual" title="Warna Latar Paragraf">${ICON_SHADING}</button>
                <input type="color" id="picker-shading-color" value="#f1f5f9" class="ribbon-color-picker-input">
              </div>
            </div>
          </div>
          <span class="ribbon-label">Paragraf</span>
        </div>

        <!-- KELOMPOK: KOP SURAT (IKON PROFESIONAL MURNI - TANPA TEKS) -->
        <div class="ribbon-group" title="Kelola Kop Surat Resmi Sekolah">
          <div class="ribbon-rows-wrapper">
            <div class="ribbon-row">
              <button id="btn-open-kop-gallery" class="ribbon-btn ribbon-btn-icon-sq" title="Galeri Kop Surat: Pilih template kop resmi tersimpan atau kelola koleksi">
                ${ICON_GALLERY}
                <span class="badge-count badge-floating">${store.getKopGallery().length}</span>
              </button>
              <input type="file" id="input-quick-kop-file" accept="image/png,image/jpeg,image/svg+xml" class="hidden">
              <button id="btn-quick-upload-kop" class="ribbon-btn ribbon-btn-icon-sq" title="Unggah Berkas Gambar Kop Baru (PNG/JPG/SVG)">
                ${ICON_UPLOAD}
              </button>
            </div>
            <div class="ribbon-row">
              <button id="btn-toggle-kop" class="ribbon-btn ribbon-btn-icon-sq" title="${showKop ? 'Sembunyikan Kop Surat' : 'Tampilkan Kop Surat'}">
                ${showKop ? ICON_EYE : ICON_EYE_OFF}
              </button>
              <button id="btn-quick-school-profile" class="ribbon-btn ribbon-btn-icon-sq" title="Profil Sekolah & Format Identitas Kop (SMKS IT Al-Hawari)">
                ${ICON_SCHOOL}
              </button>
            </div>
          </div>
          <span class="ribbon-label">Kop Surat</span>
        </div>

        <!-- KELOMPOK: GUDANG SURAT & ARSIP (IKON PROFESIONAL MURNI - TANPA TEKS) -->
        <div class="ribbon-group" title="Penyimpanan & Gudang Arsip Surat">
          <div class="ribbon-rows-wrapper">
            <div class="ribbon-row">
              <button id="btn-save-to-vault" class="ribbon-btn ribbon-btn-icon-sq" title="Simpan Surat ke Gudang Arsip & Otomatis Jadi Referensi AI">
                ${ICON_SAVE}
              </button>
              <button id="btn-open-vault-tab" class="ribbon-btn ribbon-btn-icon-sq" title="Buka Gudang Surat (${vaultLetters.length} arsip tersimpan)">
                ${ICON_ARCHIVE}
                <span class="badge-count badge-floating">${vaultLetters.length}</span>
              </button>
            </div>
            <div class="ribbon-row">
              <input type="file" id="input-upload-docx-word" accept=".docx" class="hidden">
              <button id="btn-upload-docx-word" class="ribbon-btn ribbon-btn-icon-sq" title="Buka / Impor Berkas Word (.docx) Luar dari Komputer">
                ${ICON_DOC_UPLOAD}
              </button>
              <button id="btn-quick-convert-doc" class="ribbon-btn ribbon-btn-icon-sq" title="Konversi Teks Berkas ke Format Resmi Sistem (Baku & Anti Berantakan)">
                ${ICON_CONVERT}
              </button>
            </div>
          </div>
          <span class="ribbon-label">Gudang</span>
        </div>

        <!-- KELOMPOK: AI COPILOT & EKSPOR -->
        <div class="ribbon-group ribbon-group-actions">
          <div class="ribbon-rows-wrapper">
            <div class="ribbon-row">
              <button id="btn-toggle-copilot" class="btn btn-sm btn-primary ribbon-btn-compact ribbon-btn-ai" title="Buka AI Copilot Persuratan Sekolah">
                ${ICON_SPARKLE} <span>Copilot</span>
              </button>
              <button id="btn-ai-polish-ribbon" class="btn btn-sm ribbon-btn-compact ribbon-btn-polish" title="Rapihkan dokumen dan format yang berantakan dengan AI">
                ${ICON_WAND} <span>Rapihkan</span>
              </button>
            </div>
            <div class="ribbon-row">
              <button id="btn-download-word" class="btn btn-sm btn-success ribbon-btn-compact ribbon-btn-action" style="flex: 1;" title="Unduh Berkas Dokumen Word (.docx)">
                ${ICON_DOWNLOAD} <span>Unduh</span>
              </button>
              <button id="btn-print-word" class="btn btn-sm btn-secondary ribbon-btn-compact ribbon-btn-action" style="flex: 1;" title="Cetak atau Simpan PDF">
                ${ICON_PRINTER} <span>Cetak</span>
              </button>
            </div>
          </div>
          <span class="ribbon-label">AI & Berkas</span>
        </div>

      </div>

      <!-- 2. DESK AREA DENGAN LEMBAR KERTAS WORD DAN ZOOM SCALING -->
      <div class="word-desk" id="word-desk-container">
        
        ${activeDocx ? `
          <!-- BANNER DOKUMEN IMPOR PRESERVASI FORMAT (TERPUSAT DI ATAS LEMBAR KERJA) -->
          <div class="imported-docx-banner no-print" id="imported-docx-banner">
            <div style="display: flex; align-items: center; gap: 10px;">
              <span style="color: #2563eb; display: flex; align-items: center;">${ICON_DOC}</span>
              <div>
                <div style="font-weight: 700; font-size: 13.5px; color: #1e40af;">
                  Berkas Terbuka: <u>${escapeHtml(activeDocx.fileName || 'Dokumen Word')}</u>
                </div>
                <div style="font-size: 11px; color: #2563eb; margin-top: 1px;">
                  Ingin format 100% rapi dan standar sistem? Klik tombol <b>"Konversi Format"</b> di samping.
                </div>
              </div>
            </div>
            
            <div style="display: flex; gap: 6px; align-items: center; flex-wrap: wrap;">
              <button type="button" id="btn-convert-to-native" class="btn btn-sm btn-primary ribbon-btn-compact" style="background: #2563eb; font-weight: 600;" title="Ekstrak teks dan bangun ulang ke format surat resmi sistem (anti berantakan)">
                ${ICON_CONVERT} <span>Konversi Format</span>
              </button>
              <button type="button" id="btn-ai-polish-banner" class="btn btn-sm ribbon-btn-compact ribbon-btn-polish" title="Rapihkan dokumen dan teks berantakan dengan AI">
                ${ICON_WAND} <span>Rapihkan AI</span>
              </button>
              <div class="btn-group" style="display: inline-flex; border: 1px solid #bfdbfe; border-radius: 4px; overflow: hidden; background: white;">
                <button type="button" id="btn-mode-canvas" class="btn btn-sm ${!isDocxPreviewMode ? 'btn-primary' : 'btn-secondary'}" style="font-size: 11px; border: none; padding: 3px 8px; cursor: pointer;">
                  Kanvas
                </button>
                <button type="button" id="btn-mode-native" class="btn btn-sm ${isDocxPreviewMode ? 'btn-primary' : 'btn-secondary'}" style="font-size: 11px; border: none; padding: 3px 8px; cursor: pointer;">
                  Asli Word
                </button>
              </div>
              <button id="btn-save-imported-to-vault" class="btn btn-sm btn-secondary ribbon-btn-compact" title="Simpan berkas impor ini ke Gudang Surat">${ICON_SAVE} <span>Simpan</span></button>
              <button id="btn-close-imported-docx" class="btn btn-sm btn-secondary ribbon-btn-compact" title="Kembali ke template surat sekolah">✕ Tutup</button>
            </div>
          </div>
        ` : ''}

        <div class="word-sheet-wrapper" id="word-sheet-wrapper" style="transform: scale(${zoomLevel / 100});">
          
          ${activeDocx && isDocxPreviewMode ? `
            <!-- PRATINJAU DOKUMEN ASLI MENGGUNAKAN DOCX-PREVIEW -->
            <div id="docx-native-preview-card" class="docx-native-preview-card">
              <div id="docx-preview-mount" style="min-height: 297mm;"></div>
            </div>
          ` : `
            <!-- LEMBAR KERTAS SKALA NYATA WORD -->
            <div class="word-paper-sheet paper-${paperSize} ${orientation === 'landscape' ? 'landscape' : ''}" id="word-sheet"
              style="
                --margin-top: ${customMargins.top}cm;
                --margin-bottom: ${customMargins.bottom}cm;
                --margin-left: ${customMargins.left}cm;
                --margin-right: ${customMargins.right}cm;
                font-family: ${fontFamily};
                font-size: ${fontSize}pt;
              ">
              
              <!-- BANNER KOP SURAT BERBASIS GAMBAR RESMI AL-HAWARI -->
              <div class="sheet-kop-container ${showKop ? '' : 'hidden'}" id="sheet-kop-container">
                <img src="${kopImageUrl}" alt="Kop Surat" id="sheet-kop-image" class="sheet-kop-image" title="Klik 'Galeri Kop Surat' di atas untuk mengganti kop sekolah Anda">
              </div>

              <!-- KONTEN SURAT YANG DAPAT DIEDIT LANGSUNG SEPERTI WORD -->
              <div class="sheet-content-editable" id="sheet-body-content">
                
                ${activeDocx ? `
                  <!-- JUDUL DAN NOMOR SURAT DOKUMEN IMPOR -->
                  <div class="sheet-section-title">
                    <h2 contenteditable="true" id="editable-doc-title" class="editable-field text-center font-title imported-field" data-anchor-index="${docxTitleIdx !== -1 ? docxTitleIdx : ''}" spellcheck="false">
                      ${escapeHtml(docxTitleText)}
                    </h2>
                    <div contenteditable="true" id="editable-doc-nomor" class="editable-field text-center font-nomor imported-field" data-anchor-index="${docxNomorIdx !== -1 ? docxNomorIdx : ''}" spellcheck="false">
                      ${escapeHtml(docxNomorText)}
                    </div>
                  </div>

                  <!-- BLOK SIFAT, PERIHAL & KEPADA YTH (DAPAT DIEDIT / DIHAPUS MANUAL DI CANVAS) -->
                  <div class="sheet-meta-recipient-wrapper" id="sheet-meta-recipient-wrapper">
                    <div class="meta-hover-actions no-print" style="display: flex; gap: 4px; align-items: center;">
                      <select id="select-quick-kode-surat" class="btn-meta-quick-action" style="border-color: #cbd5e1; background: #f8fafc; color: #1e293b;" title="Ubah Kode Nomor Surat Sesuai Jenisnya (UND, TU, KEP, KET, ST, dll)">
                        <option value="">⚙ Kode Nomor Surat...</option>
                        ${Object.values(LETTER_TYPE_CODES).map(c => `<option value="${c.code}">${c.code} — ${c.label}</option>`).join('')}
                      </select>
                      <button type="button" id="btn-clear-meta-recipient" class="btn-meta-quick-action" title="Hapus Sifat, Perihal & Kepada Yth dari surat ini">✕ Hapus Sifat/Kepada</button>
                      <button type="button" id="btn-restore-meta-recipient" class="btn-meta-quick-action btn-meta-restore hidden" title="Tampilkan kembali Sifat, Perihal & Kepada Yth">+ Sifat & Kepada Yth</button>
                    </div>
                    <div contenteditable="true" id="editable-meta-recipient" class="editable-field meta-recipient-box" spellcheck="false" title="Klik untuk mengedit, atau blok teks & tekan Backspace/Delete untuk menghapus manual">${buildMetaRecipientInnerHtml(docxMetaRows, docxRecipientLines)}</div>
                  </div>

                  <!-- PARAGRAF PEMBUKA / NARASI / ISI -->
                  <div class="imported-body-paragraphs">
                    ${docxBodyList.map(p => `
                      <div contenteditable="true" class="editable-field paragraph-box imported-field" data-anchor-index="${p.index}" style="margin-bottom: 12px; line-height: 1.35; text-align: ${p.alignment === 'center' ? 'center' : p.alignment === 'right' ? 'right' : 'justify'}; ${p.isBold ? 'font-weight: bold;' : ''}" spellcheck="false">
                        ${escapeHtml(p.currentText !== undefined ? p.currentText : (p.text || p.originalText))}
                      </div>
                    `).join('')}
                  </div>

                  <!-- TABEL RINCIAN IDENTITAS / JADWAL -->
                  ${docxIdentityRows && docxIdentityRows.length > 0 ? `
                    <div class="sheet-identity-wrapper" id="sheet-identity-wrapper">
                      <table class="identity-table-word">
                        <tbody>
                          ${docxIdentityRows.map(r => `
                            <tr data-anchor-index="${r.index}">
                              <td class="id-label" contenteditable="true" spellcheck="false">${escapeHtml(r.label)}</td>
                              <td class="id-val" contenteditable="true" spellcheck="false">${escapeHtml(r.value.startsWith(':') ? r.value : `: ${r.value}`)}</td>
                            </tr>
                          `).join('')}
                        </tbody>
                      </table>
                    </div>
                  ` : ''}

                  <!-- PARAGRAF PENUTUP -->
                  ${docxClosing ? `
                    <div contenteditable="true" id="editable-p-closing" class="editable-field paragraph-box imported-field" data-anchor-index="${docxClosing.index}" style="margin-bottom: 12px; line-height: 1.35; text-align: justify;" spellcheck="false">
                      ${escapeHtml(docxClosing.currentText !== undefined ? docxClosing.currentText : docxClosing.text)}
                    </div>
                  ` : ''}

                  <!-- BLOK TANDA TANGAN DI KANAN BAWAH -->
                  <div class="sheet-signature-wrapper">
                    <div class="sig-box">
                      <div contenteditable="true" id="editable-sig-date" class="editable-field text-left imported-field" data-anchor-index="${docxSig.dateIndex ?? ''}" spellcheck="false">
                        ${escapeHtml(docxSig.date || 'Garut, 14 November 2026')}
                      </div>
                      <div contenteditable="true" id="editable-sig-title" class="editable-field text-left imported-field" data-anchor-index="${docxSig.titleIndex ?? ''}" spellcheck="false">
                        ${escapeHtml(docxSig.title || 'Kepala Sekolah')},
                      </div>
                      <div class="sig-space-holder"></div>
                      <div contenteditable="true" id="editable-sig-name" class="editable-field text-left font-bold-underline imported-field" data-anchor-index="${docxSig.signerIndex ?? ''}" spellcheck="false">
                        ${escapeHtml(docxSig.signer || 'Dian Nurhadian, S.Pd.')}
                      </div>
                      <div contenteditable="true" id="editable-sig-nip" class="editable-field text-left font-nip imported-field" data-anchor-index="${docxSig.nipIndex ?? ''}" spellcheck="false">
                        ${escapeHtml(docxSig.nip && docxSig.nip !== '-' ? docxSig.nip : 'Kepala SMKS IT Al-Hawari')}
                      </div>
                    </div>
                  </div>
                ` : `
                  <!-- STANDAR TEMPLATE DOKUMEN -->
                  <div class="sheet-section-title">
                    <h2 contenteditable="true" id="editable-doc-title" class="editable-field text-center font-title" spellcheck="false">
                      ${escapeHtml(titleBlock.title || 'SURAT KETERANGAN SISWA AKTIF')}
                    </h2>
                    <div contenteditable="true" id="editable-doc-nomor" class="editable-field text-center font-nomor" spellcheck="false">
                      ${escapeHtml(titleBlock.nomor || 'Nomor: 421.3 / 085 / SMKITA / XI / 2026')}
                    </div>
                  </div>

                  <!-- BLOK SIFAT, PERIHAL & KEPADA YTH (DAPAT DIEDIT / DIHAPUS MANUAL DI CANVAS) -->
                  <div class="sheet-meta-recipient-wrapper ${isMetaDeleted ? 'is-empty' : ''}" id="sheet-meta-recipient-wrapper">
                    <div class="meta-hover-actions no-print" style="display: flex; gap: 4px; align-items: center;">
                      <select id="select-quick-kode-surat" class="btn-meta-quick-action" style="border-color: #cbd5e1; background: #f8fafc; color: #1e293b;" title="Ubah Kode Nomor Surat Sesuai Jenisnya (UND, TU, KEP, KET, ST, dll)">
                        <option value="">⚙ Kode Nomor Surat...</option>
                        ${Object.values(LETTER_TYPE_CODES).map(c => `<option value="${c.code}">${c.code} — ${c.label}</option>`).join('')}
                      </select>
                      <button type="button" id="btn-clear-meta-recipient" class="btn-meta-quick-action ${isMetaDeleted ? 'hidden' : ''}" title="Hapus Sifat, Perihal & Kepada Yth dari surat ini">✕ Hapus Sifat/Kepada</button>
                      <button type="button" id="btn-restore-meta-recipient" class="btn-meta-quick-action btn-meta-restore ${isMetaDeleted ? '' : 'hidden'}" title="Tampilkan kembali Sifat, Perihal & Kepada Yth">+ Sifat & Kepada Yth</button>
                    </div>
                    <div contenteditable="true" id="editable-meta-recipient" class="editable-field meta-recipient-box" spellcheck="false" title="Klik untuk mengedit, atau blok teks & tekan Backspace/Delete untuk menghapus manual">${isMetaDeleted ? '' : (rawMetaRecipient.customHtml !== undefined ? rawMetaRecipient.customHtml : buildMetaRecipientInnerHtml(rawMetaRecipient.metaRows, rawMetaRecipient.recipientLines))}</div>
                  </div>

                  <!-- PARAGRAF PEMBUKA -->
                  <div contenteditable="true" id="editable-p-opening" class="editable-field paragraph-box" spellcheck="false">
                    ${bodyParagraphs[0]?.runs ? bodyParagraphs[0].runs.map(r => r.text).join('') : (bodyParagraphs[0]?.text || 'Yang bertanda tangan di bawah ini, Kepala SMKS IT Al-Hawari, dengan ini menerangkan bahwa:')}
                  </div>

                  <!-- TABEL RINCIAN IDENTITAS -->
                  <div class="sheet-identity-wrapper" id="sheet-identity-wrapper">
                    <table class="identity-table-word">
                      <tbody>
                        ${(identityBlock?.rows || [
                          { label: 'Nama Lengkap', value: ': Ahmad Fauzi' },
                          { label: 'NIS / NISN', value: ': 12345 / 0054321987' },
                          { label: 'Kelas / Jurusan', value: ': XII MM 1 / Multimedia' },
                          { label: 'Tempat, Tanggal Lahir', value: ': Garut, 12 Mei 2008' },
                          { label: 'Jenis Kelamin', value: ': Laki-laki' }
                        ]).map((r, rIdx) => `
                          <tr>
                            <td class="id-label" contenteditable="true" spellcheck="false">${escapeHtml(r.label)}</td>
                            <td class="id-val" contenteditable="true" spellcheck="false">${escapeHtml(r.value)}</td>
                          </tr>
                        `).join('')}
                      </tbody>
                    </table>
                  </div>

                  <!-- PARAGRAF PERNYATAAN / ISI -->
                  <div contenteditable="true" id="editable-p-body" class="editable-field paragraph-box" spellcheck="false">
                    ${bodyParagraphs[1]?.runs ? bodyParagraphs[1].runs.map(r => r.text).join('') : (bodyParagraphs[1]?.text || 'Adalah benar-benar peserta didik yang tercatat aktif terdaftar pada SMKS IT Al-Hawari pada Tahun Ajaran 2026/2027, serta senantiasa menaati seluruh tata tertib dan berkelakuan baik.')}
                  </div>

                  <!-- PARAGRAF MAKSUD / KEPERLUAN -->
                  <div contenteditable="true" id="editable-p-purpose" class="editable-field paragraph-box" spellcheck="false">
                    ${bodyParagraphs[2]?.runs ? bodyParagraphs[2].runs.map(r => r.text).join('') : (bodyParagraphs[2]?.text || 'Surat keterangan ini diterbitkan dan diberikan kepada yang bersangkutan untuk keperluan: Kelengkapan berkas pendaftaran Program Beasiswa Prestasi dan Tunjangan Pendidikan.')}
                  </div>

                  <!-- PARAGRAF PENUTUP -->
                  <div contenteditable="true" id="editable-p-closing" class="editable-field paragraph-box" spellcheck="false">
                    ${bodyParagraphs[3]?.runs ? bodyParagraphs[3].runs.map(r => r.text).join('') : (bodyParagraphs[3]?.text || 'Demikian surat keterangan ini kami buat dengan sebenarnya agar dapat dipergunakan sebagaimana mestinya.')}
                  </div>

                  <!-- BLOK TANDA TANGAN DI KANAN BAWAH -->
                  <div class="sheet-signature-wrapper">
                    <div class="sig-box">
                      <div contenteditable="true" id="editable-sig-date" class="editable-field text-left" spellcheck="false">
                        ${escapeHtml(sigBlock.city || 'Garut')}, ${escapeHtml(sigBlock.date || '14 November 2026')}
                      </div>
                      <div contenteditable="true" id="editable-sig-title" class="editable-field text-left" spellcheck="false">
                        ${escapeHtml(sigBlock.title || 'Kepala Sekolah')},
                      </div>
                      <div class="sig-space-holder"></div>
                      <div contenteditable="true" id="editable-sig-name" class="editable-field text-left font-bold-underline" spellcheck="false">
                        ${escapeHtml(sigBlock.name || 'Dian Nurhadian, S.Pd.')}
                      </div>
                      <div contenteditable="true" id="editable-sig-nip" class="editable-field text-left font-nip" spellcheck="false">
                        ${sigBlock.nip && sigBlock.nip !== '-' ? `NIP. ${escapeHtml(sigBlock.nip)}` : 'Kepala SMKS IT Al-Hawari'}
                      </div>
                    </div>
                  </div>
                `}

              </div>
            </div>
          `}
        </div>
      </div>

      <!-- 3. STATUS BAR MICROSOFT WORD (DI BAWAH) DENGAN ZOOM SLIDER KHUSUS KANVAS -->
      <footer class="word-status-bar">
        <div class="status-left">
          <span>Halaman 1 dari 1</span>
          <span>|</span>
          <span id="status-word-count">128 kata</span>
          <span>|</span>
          <span>Bahasa Indonesia (PUEBI)</span>
        </div>

        <div class="status-right">
          <div class="zoom-slider-wrap">
            <button class="zoom-btn" id="btn-zoom-out" title="Perkecil Kanvas">-</button>
            <input type="range" id="input-zoom-slider" min="50" max="175" step="5" value="${zoomLevel}" title="Geser skala zoom kanvas">
            <button class="zoom-btn" id="btn-zoom-in" title="Perbesar Kanvas">+</button>
            <span id="label-zoom-val" style="min-width: 38px; font-weight: 600;">${zoomLevel}%</span>
          </div>
          <button id="btn-zoom-fit" class="btn btn-sm btn-secondary" style="padding: 1px 6px; font-size: 11px;">
            Fit Lebar
          </button>
        </div>
      </footer>

      <!-- 4. MODAL GUDANG KOP SURAT (KOP GALLERY) -->
      <div id="modal-kop-gallery" class="modal-overlay hidden" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.55); z-index: 10000; display: flex; align-items: center; justify-content: center;">
        <div class="card" style="width: 100%; max-width: 780px; max-height: 85vh; background: white; border-radius: 8px; padding: 20px; box-shadow: 0 14px 40px rgba(0,0,0,0.3); display: flex; flex-direction: column;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; border-bottom: 1px solid var(--border); padding-bottom: 10px;">
            <div>
              <h3 style="font-size: 17px; font-weight: 700; color: var(--dark); display: flex; align-items: center; gap: 8px;">
                <span>🖼</span> Gudang Kop Surat Sekolah
              </h3>
              <p style="font-size: 12px; color: var(--text-muted); margin-top: 2px;">
                Pilih kop surat yang sudah tersimpan untuk langsung diterapkan ke lembar kerja tanpa perlu mengunggah ulang.
              </p>
            </div>
            <button id="btn-close-kop-modal" class="btn-icon" style="font-size: 24px;">&times;</button>
          </div>

          <!-- KONTEN GALERI KOP -->
          <div id="kop-gallery-cards-container" style="flex: 1; overflow-y: auto; padding-right: 6px; display: flex; flex-direction: column; gap: 14px; margin-bottom: 14px;">
            <!-- Rendered by JS -->
          </div>

          <!-- FORM UNGGAH KOP BARU KE GALERI -->
          <div style="border-top: 1px solid var(--border); padding-top: 12px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px;">
            <div style="display: flex; align-items: center; gap: 8px;">
              <input type="file" id="input-modal-upload-kop" accept="image/png,image/jpeg,image/svg+xml" class="hidden">
              <button id="btn-trigger-upload-modal" class="btn btn-sm btn-primary">
                ➕ Unggah Kop Baru ke Galeri
              </button>
            </div>
            <button id="btn-close-kop-modal-footer" class="btn btn-sm btn-secondary">Tutup Galeri</button>
          </div>
        </div>
      </div>

      <!-- 4B. MODAL PENGATURAN LENGKAP DOKUMEN & HALAMAN -->
      <div id="modal-doc-settings" class="modal-overlay hidden" style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.55); z-index: 10000; display: flex; align-items: center; justify-content: center;">
        <div class="card" style="width: 100%; max-width: 620px; max-height: 90vh; background: white; border-radius: 8px; padding: 22px; box-shadow: 0 16px 45px rgba(0,0,0,0.35); display: flex; flex-direction: column; overflow-y: auto;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; border-bottom: 1px solid var(--border); padding-bottom: 10px;">
            <div>
              <h3 style="font-size: 18px; font-weight: 700; color: var(--dark); display: flex; align-items: center; gap: 8px;">
                <span>⚙️</span> Pengaturan Format & Tata Letak Dokumen
              </h3>
              <p style="font-size: 12px; color: var(--text-muted); margin-top: 2px;">
                Sesuaikan ukuran kertas, batas margin, tipografi, dan kop surat sesuai standar sekolah.
              </p>
            </div>
            <button id="btn-close-settings-modal" class="btn-icon" style="font-size: 24px;">&times;</button>
          </div>

          <!-- PENGATURAN KERTAS & ORIENTASI -->
          <div class="form-group" style="margin-bottom: 16px;">
            <label style="font-weight: 700; color: #1e293b; font-size: 13px;">1. Ukuran Kertas & Orientasi:</label>
            <div class="form-row" style="margin-top: 6px;">
              <div class="col">
                <label style="font-size: 11px;">Ukuran Kertas:</label>
                <select id="modal-set-paper-size" class="form-control">
                  <option value="a4" ${paperSize === 'a4' ? 'selected' : ''}>A4 (21.0 x 29.7 cm) - Standar Dinas</option>
                  <option value="f4" ${paperSize === 'f4' ? 'selected' : ''}>F4 / Folio (21.5 x 33.0 cm) - Ijazah/Dinas Jabar</option>
                  <option value="letter" ${paperSize === 'letter' ? 'selected' : ''}>Letter / Kuarto (21.59 x 27.94 cm)</option>
                  <option value="legal" ${paperSize === 'legal' ? 'selected' : ''}>Legal (21.59 x 35.56 cm)</option>
                </select>
              </div>
              <div class="col">
                <label style="font-size: 11px;">Orientasi:</label>
                <select id="modal-set-orientation" class="form-control">
                  <option value="portrait" ${orientation === 'portrait' ? 'selected' : ''}>📄 Potret (Vertikal)</option>
                  <option value="landscape" ${orientation === 'landscape' ? 'selected' : ''}>📑 Lanskap (Horizontal)</option>
                </select>
              </div>
            </div>
          </div>

          <!-- PENGATURAN MARGIN -->
          <div class="form-group" style="margin-bottom: 16px;">
            <div style="display: flex; justify-content: space-between; align-items: center;">
              <label style="font-weight: 700; color: #1e293b; font-size: 13px;">2. Batas Margin Kertas (cm):</label>
              <div style="display: flex; gap: 4px;">
                <button type="button" class="btn btn-sm btn-secondary btn-modal-preset-margin" data-top="2.0" data-bottom="2.0" data-left="2.5" data-right="2.0" style="font-size: 10.5px; padding: 2px 6px;">Standar Dinas</button>
                <button type="button" class="btn btn-sm btn-secondary btn-modal-preset-margin" data-top="2.54" data-bottom="2.54" data-left="2.54" data-right="2.54" style="font-size: 10.5px; padding: 2px 6px;">Normal Word</button>
                <button type="button" class="btn btn-sm btn-secondary btn-modal-preset-margin" data-top="1.27" data-bottom="1.27" data-left="1.27" data-right="1.27" style="font-size: 10.5px; padding: 2px 6px;">Sempit</button>
              </div>
            </div>
            <div class="form-row" style="margin-top: 8px;">
              <div class="col"><label style="font-size:11px;">Kiri (Left):</label><input type="number" step="0.1" id="modal-set-m-left" class="form-control" value="${customMargins.left}"></div>
              <div class="col"><label style="font-size:11px;">Kanan (Right):</label><input type="number" step="0.1" id="modal-set-m-right" class="form-control" value="${customMargins.right}"></div>
              <div class="col"><label style="font-size:11px;">Atas (Top):</label><input type="number" step="0.1" id="modal-set-m-top" class="form-control" value="${customMargins.top}"></div>
              <div class="col"><label style="font-size:11px;">Bawah (Bottom):</label><input type="number" step="0.1" id="modal-set-m-bottom" class="form-control" value="${customMargins.bottom}"></div>
            </div>
          </div>

          <!-- PENGATURAN TIPOGRAFI & SPASI -->
          <div class="form-group" style="margin-bottom: 16px;">
            <label style="font-weight: 700; color: #1e293b; font-size: 13px;">3. Tipografi & Jarak Spasi Baris:</label>
            <div class="form-row" style="margin-top: 6px;">
              <div class="col">
                <label style="font-size: 11px;">Jenis Font Utama:</label>
                <select id="modal-set-font-family" class="form-control">
                  <option value="Times New Roman" ${fontFamily === 'Times New Roman' ? 'selected' : ''}>Times New Roman (Baku)</option>
                  <option value="Arial" ${fontFamily === 'Arial' ? 'selected' : ''}>Arial</option>
                  <option value="Calibri" ${fontFamily === 'Calibri' ? 'selected' : ''}>Calibri</option>
                  <option value="Segoe UI" ${fontFamily === 'Segoe UI' ? 'selected' : ''}>Segoe UI</option>
                  <option value="Georgia" ${fontFamily === 'Georgia' ? 'selected' : ''}>Georgia</option>
                </select>
              </div>
              <div class="col">
                <label style="font-size: 11px;">Ukuran Font:</label>
                <select id="modal-set-font-size" class="form-control">
                  <option value="10" ${fontSize === '10' ? 'selected' : ''}>10 pt</option>
                  <option value="11" ${fontSize === '11' ? 'selected' : ''}>11 pt</option>
                  <option value="12" ${fontSize === '12' ? 'selected' : ''}>12 pt (Standar Dinas)</option>
                  <option value="14" ${fontSize === '14' ? 'selected' : ''}>14 pt</option>
                </select>
              </div>
              <div class="col">
                <label style="font-size: 11px;">Spasi Baris:</label>
                <select id="modal-set-line-spacing" class="form-control">
                  <option value="1.0" ${lineSpacing === '1.0' ? 'selected' : ''}>1.0 (Tunggal)</option>
                  <option value="1.15" ${lineSpacing === '1.15' ? 'selected' : ''}>1.15 (Standar Dinas)</option>
                  <option value="1.5" ${lineSpacing === '1.5' ? 'selected' : ''}>1.5 (Satu Setengah)</option>
                  <option value="2.0" ${lineSpacing === '2.0' ? 'selected' : ''}>2.0 (Ganda)</option>
                </select>
              </div>
            </div>
          </div>

          <!-- KOP SURAT -->
          <div class="form-group" style="margin-bottom: 20px;">
            <label style="font-weight: 700; color: #1e293b; font-size: 13px;">4. Kop Surat Sekolah:</label>
            <div style="display: flex; align-items: center; justify-content: space-between; margin-top: 6px; padding: 10px; background: #f8fafc; border: 1px solid var(--border); border-radius: 6px;">
              <div>
                <strong style="font-size: 12.5px; color: var(--dark);">Tampilkan Kop Surat pada Dokumen</strong>
                <p style="font-size: 11px; color: var(--text-muted); margin: 2px 0 0 0;">Jika dimatikan, halaman akan dikosongkan bagian atasnya (misal untuk kertas berkop fisik).</p>
              </div>
              <input type="checkbox" id="modal-set-show-kop" ${showKop ? 'checked' : ''} style="width: 20px; height: 20px; cursor: pointer;">
            </div>
          </div>

          <!-- TOMBOL AKSI -->
          <div style="display: flex; justify-content: flex-end; gap: 8px; border-top: 1px solid var(--border); padding-top: 14px;">
            <button type="button" id="btn-cancel-doc-settings" class="btn btn-secondary">Batal</button>
            <button type="button" id="btn-save-doc-settings" class="btn btn-primary" style="font-weight: 600;">💾 Simpan & Terapkan Pengaturan</button>
          </div>
        </div>
      </div>

      <!-- 5. PANEL AI COPILOT (DRAWER SAMPING DENGAN PEMILIH MODEL GUDANG API) -->
      <aside class="copilot-drawer hidden" id="copilot-drawer">
        <div class="copilot-header">
          <div class="copilot-title">
            <span class="copilot-sparkle">✨</span>
            <h4>AI Copilot Surat Sekolah</h4>
          </div>
          <button id="btn-close-copilot" class="btn-icon" title="Tutup">&times;</button>
        </div>

        <div class="copilot-body">
          
          <!-- PEMILIH MODEL DARI GUDANG API -->
          <div class="form-group" style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 10px; margin-bottom: 12px;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
              <label style="font-size: 11px; font-weight: 700; color: #334155;">Model AI dari Gudang API:</label>
              <a href="javascript:void(0)" id="link-open-vault-from-copilot" style="font-size: 11px; color: var(--primary); text-decoration: underline;">
                ⚙ Gudang API
              </a>
            </div>
            <select id="select-copilot-model" class="form-control" style="font-size: 12px; font-weight: 600;">
              ${apiVault.map(v => `
                <option value="${v.id}" ${v.id === activeApi?.id ? 'selected' : ''}>
                  [${v.provider.toUpperCase()}] ${escapeHtml(v.label)}
                </option>
              `).join('')}
            </select>
          </div>

          <!-- BASIS PENGETAHUAN / KNOWLEDGE BASE DARI GUDANG SURAT -->
          <div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 6px; padding: 8px 10px; margin-bottom: 12px; font-size: 11.5px; color: #166534; display: flex; align-items: center; justify-content: space-between;">
            <span>📚 <strong>Basis Pengetahuan:</strong> ${aiRefs.length} arsip surat aktif sebagai acuan gaya</span>
            <a href="javascript:void(0)" id="link-open-vault-from-copilot-kb" style="color: #15803d; font-weight: 700; text-decoration: underline;">
              Kelola Gudang
            </a>
          </div>

          <p class="copilot-desc">
            Ketik jenis surat apa pun yang ingin Anda buat dalam kalimat bebas. AI akan otomatis merangkai isi dokumen dan menampilkannya langsung di lembar kerja Word Anda.
          </p>

          <!-- CHIPS TEMPLATE CEPAT -->
          <div class="copilot-chips">
            <button class="copilot-chip" data-prompt="Buatkan Surat Undangan Rapat Dinas Pleno Dewan Guru membahas evaluasi pembelajaran (Kode UND)">✉ Undangan Rapat (UND)</button>
            <button class="copilot-chip" data-prompt="Buatkan Surat dari Tata Usaha (TU) tentang pemberitahuan pengumpulan berkas administrasi kelas (Kode TU)">🗂 Surat Tata Usaha (TU)</button>
            <button class="copilot-chip" data-prompt="Buatkan Surat Keputusan Kepala Sekolah tentang penetapan panitia kegiatan sekolah (Kode KEP)">🏛 SK Kepala Sekolah (KEP)</button>
            <button class="copilot-chip" data-prompt="Buatkan Surat Keterangan Siswa Aktif untuk keperluan beasiswa kuliah (Kode KET)">📝 Ket. Siswa Aktif (KET)</button>
            <button class="copilot-chip" data-prompt="Buatkan Surat Tugas untuk guru mengikuti bimbingan teknis kurikulum merdeka di dinas pendidikan (Kode ST)">💼 Surat Tugas (ST)</button>
            <button class="copilot-chip" data-prompt="Buatkan Surat Panggilan Orang Tua/Wali Siswa terkait pembinaan kedisiplinan siswa (Kode PGL)">📢 Panggilan Ortu (PGL)</button>
          </div>

          <div class="form-group" style="margin-top: 14px;">
            <label>Perintah Pembuatan / Revisi:</label>
            <textarea id="copilot-prompt-input" class="form-control" rows="4" 
              placeholder="Contoh: Buatkan surat tugas untuk guru SMKS IT Al-Hawari mengikuti pelatihan kurikulum merdeka di Garut tanggal 20-22 November..."></textarea>
          </div>

          <button id="btn-run-copilot" class="btn btn-primary" style="width: 100%;">
            🚀 Buat Dokumen dengan AI
          </button>

          <div id="copilot-status" class="copilot-status-area hidden">
            <div style="display: flex; align-items: center; justify-content: center; gap: 8px;">
              <div class="spinner" style="width: 16px; height: 16px; margin: 0; border-width: 2px;"></div>
              <span id="copilot-status-text" style="font-weight: 600;">AI sedang merangkai surat...</span>
              <button id="btn-cancel-copilot" type="button" class="btn btn-sm btn-secondary" style="padding: 1px 8px; font-size: 11px; margin-left: 4px; height: 22px; border-color: #cbd5e1;">Batal</button>
            </div>
            <div id="copilot-status-tip" style="margin-top: 6px; font-size: 11px; color: #475569; display: none;">
              💡 <em>Tips: Jika Gemini terasa lambat di jam sibuk, Anda dapat memilih <b>Groq (~0.2 detik)</b> di dropdown atas.</em>
            </div>
          </div>

          <hr style="margin: 16px 0;"/>
          <h5>Aksi Cepat AI Copilot:</h5>
          <div class="btn-group-vertical" style="display: flex; flex-direction: column; gap: 6px; margin-top: 8px;">
            <button id="btn-copilot-polish-messy" class="btn btn-sm" style="text-align: left; background: #fef3c7; border: 1px solid #fde68a; color: #92400e; font-weight: 600;">
              🧹 Rapihkan Dokumen yang Berantakan (AI)
            </button>
            <button id="btn-copilot-formalize" class="btn btn-sm btn-secondary" style="text-align: left;">
              ✨ Rapikan Bahasa ke Standar Dinas Baku (PUEBI)
            </button>
            <button id="btn-copilot-today-date" class="btn btn-sm btn-secondary" style="text-align: left;">
              📅 Perbarui Tanggal Surat ke Hari Ini
            </button>
          </div>
        </div>
      </aside>

    </div>
  `;

  // 1. Bind Ribbon Controls (Formatting & Page)
  bindRibbonControls(container);

  // 2. Bind Zoom Controls (Dedicated Canvas Zoom)
  bindZoomControls(container);

  // 3. Bind Kop Gallery Modal
  bindKopGalleryModal(container);

  // 4. Bind Document Settings Modal
  bindDocSettingsModal(container);

  // 5. Bind Live Canvas Editing
  bindCanvasLiveEditing(container);

  // 6. Bind Copilot Drawer
  bindCopilotDrawer(container);

  // Update status kata
  updateWordCount(container);

  // Pastikan posisi ribbon dan window selalu terikat di paling kiri (Halaman & Margin selalu tampak penuh)
  const ribbonEl = container.querySelector('.word-ribbon');
  if (ribbonEl) {
    ribbonEl.scrollLeft = 0;
  }
  if (typeof window !== 'undefined') {
    window.scrollTo(0, 0);
  }
}

/**
 * Event handler Ribbon Toolbar
 */
function bindRibbonControls(container) {
  // Cegah hilangnya fokus/seleksi teks saat tombol ribbon diklik
  const preventSelectionLoss = (e) => {
    e.preventDefault();
  };

  const bindFormatBtn = (id, command, value = null) => {
    const btn = container.querySelector(id);
    if (!btn) return;
    btn.addEventListener('mousedown', preventSelectionLoss);
    btn.addEventListener('click', () => {
      document.execCommand(command, false, value);
      scheduleSync(container);
    });
  };

  // Font formatting row 2
  bindFormatBtn('#cmd-bold', 'bold');
  bindFormatBtn('#cmd-italic', 'italic');
  bindFormatBtn('#cmd-underline', 'underline');
  bindFormatBtn('#cmd-strike', 'strikeThrough');
  bindFormatBtn('#cmd-subscript', 'subscript');
  bindFormatBtn('#cmd-superscript', 'superscript');

  // Paragraph formatting row 1
  bindFormatBtn('#cmd-bullet-list', 'insertUnorderedList');
  bindFormatBtn('#cmd-numbered-list', 'insertOrderedList');
  bindFormatBtn('#cmd-outdent', 'outdent');
  bindFormatBtn('#cmd-indent', 'indent');

  // Paragraph formatting row 2
  bindFormatBtn('#cmd-align-left', 'justifyLeft');
  bindFormatBtn('#cmd-align-center', 'justifyCenter');
  bindFormatBtn('#cmd-align-right', 'justifyRight');
  bindFormatBtn('#cmd-align-justify', 'justifyFull');

  // Font Family selector
  const fontFamilySelect = container.querySelector('#cmd-font-family');
  fontFamilySelect.addEventListener('change', () => {
    const font = fontFamilySelect.value;
    document.execCommand('fontName', false, font);
    scheduleSync(container);
  });

  // Font Size selector
  const fontSizeSelect = container.querySelector('#cmd-font-size');
  fontSizeSelect.addEventListener('change', () => {
    const size = fontSizeSelect.value;
    // Map pt to execCommand 1-7 or inline style span
    applyInlineStyleToSelection('font-size', `${size}pt`);
    scheduleSync(container);
  });

  // Grow Font (A^) & Shrink Font (Av)
  container.querySelector('#cmd-grow-font').addEventListener('mousedown', preventSelectionLoss);
  container.querySelector('#cmd-grow-font').addEventListener('click', () => {
    adjustFontSize(2);
    scheduleSync(container);
  });

  container.querySelector('#cmd-shrink-font').addEventListener('mousedown', preventSelectionLoss);
  container.querySelector('#cmd-shrink-font').addEventListener('click', () => {
    adjustFontSize(-2);
    scheduleSync(container);
  });

  // Change Case (Aa)
  container.querySelector('#cmd-change-case').addEventListener('mousedown', preventSelectionLoss);
  container.querySelector('#cmd-change-case').addEventListener('click', () => {
    toggleCaseOnSelection();
    scheduleSync(container);
  });

  // Clear Formatting (A🧹)
  container.querySelector('#cmd-clear-format').addEventListener('mousedown', preventSelectionLoss);
  container.querySelector('#cmd-clear-format').addEventListener('click', () => {
    document.execCommand('removeFormat', false, null);
    scheduleSync(container);
  });

  // Text Highlight Color Picker (Stabilo)
  const highlightInput = container.querySelector('#picker-highlight-color');
  highlightInput.addEventListener('input', () => {
    document.execCommand('hiliteColor', false, highlightInput.value);
    scheduleSync(container);
  });

  // Font Color Picker
  const fontColorInput = container.querySelector('#picker-font-color');
  fontColorInput.addEventListener('input', () => {
    document.execCommand('foreColor', false, fontColorInput.value);
    scheduleSync(container);
  });

  // Shading / Background Color Fill
  const shadingInput = container.querySelector('#picker-shading-color');
  shadingInput.addEventListener('input', () => {
    applyInlineStyleToSelection('background-color', shadingInput.value);
    scheduleSync(container);
  });

  // Line Spacing
  const spacingSelect = container.querySelector('#ribbon-line-spacing');
  spacingSelect.addEventListener('change', () => {
    const val = spacingSelect.value;
    const bodyContent = container.querySelector('#sheet-body-content');
    bodyContent.style.lineHeight = val === '1.0' ? '1.15' : (val === '1.5' ? '1.6' : (val === '2.0' ? '2.0' : '1.35'));
    showToast(`Jarak spasi baris diatur ke ${val}`, 'info');
    scheduleSync(container);
  });

  // Paper Size & Orientation
  const paperSelect = container.querySelector('#ribbon-paper-size');
  if (paperSelect) {
    paperSelect.addEventListener('change', () => {
      const size = paperSelect.value;
      store.updateLetterOptions({ paperSize: size });
      showToast(`Ukuran kertas diubah ke ${size.toUpperCase()}`, 'info');
      renderWordCanvas(container);
    });
  }

  const orientationSelect = container.querySelector('#ribbon-orientation');
  if (orientationSelect) {
    orientationSelect.addEventListener('change', () => {
      const orientation = orientationSelect.value;
      store.updateLetterOptions({ orientation });
      showToast(`Orientasi kertas diubah ke ${orientation === 'portrait' ? 'Potret' : 'Lanskap'}`, 'info');
      renderWordCanvas(container);
    });
  }

  // Margins Dropdown Menu
  const btnMarginMenu = container.querySelector('#btn-margin-menu');
  const marginMenu = container.querySelector('#margin-dropdown-menu');

  btnMarginMenu.addEventListener('click', (e) => {
    e.stopPropagation();
    marginMenu.classList.toggle('hidden');
  });

  document.addEventListener('click', (e) => {
    if (!marginMenu.contains(e.target) && e.target !== btnMarginMenu) {
      marginMenu.classList.add('hidden');
    }
  });

  container.querySelectorAll('.btn-preset-margin').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const top = parseFloat(btn.dataset.top);
      const bottom = parseFloat(btn.dataset.bottom);
      const left = parseFloat(btn.dataset.left);
      const right = parseFloat(btn.dataset.right);

      applyNewMargins(container, { top, bottom, left, right });
      marginMenu.classList.add('hidden');
    });
  });

  container.querySelector('#btn-apply-custom-margin').addEventListener('click', (e) => {
    e.stopPropagation();
    const top = parseFloat(container.querySelector('#input-m-top').value) || 2.0;
    const bottom = parseFloat(container.querySelector('#input-m-bottom').value) || 2.0;
    const left = parseFloat(container.querySelector('#input-m-left').value) || 2.5;
    const right = parseFloat(container.querySelector('#input-m-right').value) || 2.0;

    applyNewMargins(container, { top, bottom, left, right });
    marginMenu.classList.add('hidden');
  });

  // Toggle Kop Surat
  const btnToggleKop = container.querySelector('#btn-toggle-kop');
  const kopContainer = container.querySelector('#sheet-kop-container');
  btnToggleKop.addEventListener('click', () => {
    const isHidden = kopContainer.classList.toggle('hidden');
    btnToggleKop.innerHTML = isHidden ? ICON_EYE_OFF : ICON_EYE;
    btnToggleKop.title = isHidden ? 'Tampilkan Kop Surat' : 'Sembunyikan Kop Surat';
    store.updateLetterOptions({ showKop: !isHidden });
    showToast(isHidden ? 'Kop surat disembunyikan.' : 'Kop surat ditampilkan.', 'info');
    scheduleSync(container);
  });

  // Shortcut Profil & Kop Resmi Sekolah
  const btnQuickSchool = container.querySelector('#btn-quick-school-profile');
  if (btnQuickSchool) {
    btnQuickSchool.addEventListener('click', () => {
      const schoolTab = document.querySelector('.nav-tab[data-tab="school"]');
      if (schoolTab) schoolTab.click();
    });
  }

  // Quick Upload Kop Surat
  const quickKopInput = container.querySelector('#input-quick-kop-file');
  const btnQuickUpload = container.querySelector('#btn-quick-upload-kop');
  btnQuickUpload.addEventListener('click', () => quickKopInput.click());

  quickKopInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (file.size > 3 * 1024 * 1024) {
      showToast('Ukuran gambar kop melebihi 3 MB.', 'error');
      return;
    }

    const reader = new FileReader();
    reader.onload = (ev) => {
      const dataUrl = ev.target.result;
      const kopImg = container.querySelector('#sheet-kop-image');
      kopImg.src = dataUrl;
      kopContainer.classList.remove('hidden');

      // Simpan juga ke Gudang Kop Surat
      const newEntry = store.addKopGalleryEntry({
        name: file.name.replace(/\.[^.]+$/, ''),
        dataUrl,
        category: 'Unggahan Pengguna'
      });

      store.updateSchoolProfile({ kopImage: dataUrl });
      showToast(`Kop baru "${newEntry.name}" berhasil diterapkan dan disimpan di galeri!`, 'success');
      scheduleSync(container);
    };
    reader.readAsDataURL(file);
  });

  // Gudang & Arsip Surat: Simpan Surat ke Gudang
  const btnSaveToVault = container.querySelector('#btn-save-to-vault');
  if (btnSaveToVault) {
    btnSaveToVault.addEventListener('click', () => {
      const activeDocx = store.getState().activeUploadedDocx;
      if (activeDocx) {
        store.addLetterVaultEntry({
          title: activeDocx.parsedMeta?.title || activeDocx.fileName?.replace(/\.docx$/i, '') || 'Dokumen Word Impor',
          nomorSurat: activeDocx.parsedMeta?.nomor || '-',
          tanggalSurat: activeDocx.parsedMeta?.date || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
          category: 'Dokumen Impor',
          targetPerson: activeDocx.parsedMeta?.signer || '-',
          perihal: activeDocx.parsedMeta?.title,
          summary: `Dokumen .docx resmi diunggah: "${activeDocx.fileName}". Preservasi format 100%.`,
          isAiReference: true,
          source: 'upload',
          uploadedDocx: activeDocx
        });
        showToast('Dokumen tersimpan di Gudang Surat & aktif sebagai referensi AI!', 'success');
        renderWordCanvas(container);
        return;
      }

      const docModel = buildModelFromCanvas(container);
      const titleEl = container.querySelector('#editable-doc-title');
      const nomorEl = container.querySelector('#editable-doc-nomor');
      const sigDateEl = container.querySelector('#editable-sig-date');
      const pBody = container.querySelector('#editable-p-body');

      const title = titleEl ? titleEl.innerText.trim() : 'Surat Keterangan';
      const nomor = nomorEl ? nomorEl.innerText.trim() : '-';
      const dateText = sigDateEl ? sigDateEl.innerText.trim() : new Date().toLocaleDateString('id-ID');
      const summary = pBody ? pBody.innerText.trim().slice(0, 160) : 'Dokumen persuratan sekolah.';

      store.addLetterVaultEntry({
        title,
        nomorSurat: nomor,
        tanggalSurat: dateText,
        category: 'Surat Keterangan',
        targetPerson: 'Ahmad Fauzi',
        perihal: title,
        summary,
        isAiReference: true,
        source: 'created',
        document: docModel
      });
      showToast(`Surat "${title}" berhasil disimpan ke Gudang Surat & otomatis jadi referensi AI!`, 'success');
      renderWordCanvas(container);
    });
  }

  // Gudang & Arsip Surat: Buka Tab Gudang
  const btnOpenVaultTab = container.querySelector('#btn-open-vault-tab');
  if (btnOpenVaultTab) {
    btnOpenVaultTab.addEventListener('click', () => {
      const vaultTab = document.querySelector('.nav-tab[data-tab="vault"]');
      if (vaultTab) vaultTab.click();
    });
  }

  // Gudang & Arsip Surat: Buka & Unggah Berkas .docx Luar
  const btnUploadDocx = container.querySelector('#btn-upload-docx-word');
  const inputUploadDocx = container.querySelector('#input-upload-docx-word');
  if (btnUploadDocx && inputUploadDocx) {
    btnUploadDocx.addEventListener('click', () => inputUploadDocx.click());
    inputUploadDocx.addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      if (!file.name.endsWith('.docx')) {
        showToast('Harap pilih berkas dengan ekstensi .docx Microsoft Word.', 'error');
        return;
      }
      showToast(`Membaca dan memproses berkas "${file.name}"...`, 'info');
      try {
        const arrayBuffer = await file.arrayBuffer();
        const parsed = await parseDocxForEditing(arrayBuffer);
        const uploadedData = {
          rawBytes: new Uint8Array(arrayBuffer),
          fileName: file.name,
          parsedMeta: parsed.meta,
          structured: parsed.structured,
          paragraphs: parsed.editableParagraphs
        };

        // Buat ulang teks otomatis ke Document Model sistem resmi agar rapi seketika
        const convertedDoc = convertDocxToDocumentModel(
          uploadedData,
          store.getSchoolProfile(),
          store.getLetterOptions()
        );

        store.setState({
          activeUploadedDocx: uploadedData,
          document: convertedDoc,
          docxViewMode: 'canvas'
        }, true);

        showToast(`Berkas "${file.name}" berhasil diimpor & teks dirapikan ke format resmi!`, 'success');
        renderWordCanvas(container);
      } catch (err) {
        showToast(`Gagal membuka berkas .docx: ${err.message}`, 'error');
      }
    });
  }

  // Tombol Cepat Konversi Dokumen di Ribbon Gudang
  const btnQuickConvert = container.querySelector('#btn-quick-convert-doc');
  if (btnQuickConvert) {
    btnQuickConvert.addEventListener('click', () => {
      const convertBtn = container.querySelector('#btn-convert-to-native');
      if (convertBtn) {
        convertBtn.click();
      } else {
        const active = store.getState().activeUploadedDocx;
        if (active) {
          const convertedDoc = convertDocxToDocumentModel(
            active,
            store.getSchoolProfile(),
            store.getLetterOptions()
          );
          store.clearActiveUploadedDocx();
          store.setState({ document: convertedDoc }, true);
          showToast('Dokumen berhasil dikonversi & dirapihkan ke format resmi sekolah!', 'success');
          renderWordCanvas(container);
        } else {
          showToast('Buka atau impor berkas .docx luar terlebih dahulu untuk mengonversinya.', 'info');
        }
      }
    });
  }

  // Banner Dokumen Impor: Tombol Konversi Teks ke Format Resmi Sistem
  const btnConvertNative = container.querySelector('#btn-convert-to-native');
  if (btnConvertNative) {
    btnConvertNative.addEventListener('click', () => {
      const activeDocx = store.getState().activeUploadedDocx;
      if (!activeDocx) return;
      showToast('Mengonversi teks dan menyusun ulang ke format resmi sistem...', 'info');
      try {
        const convertedDoc = convertDocxToDocumentModel(
          activeDocx,
          store.getSchoolProfile(),
          store.getLetterOptions()
        );
        store.clearActiveUploadedDocx();
        store.setState({ document: convertedDoc }, true);
        showToast('Dokumen berhasil dikonversi & dirapihkan ke format resmi sekolah!', 'success');
        renderWordCanvas(container);
      } catch (err) {
        showToast(`Gagal mengonversi dokumen: ${err.message}`, 'error');
      }
    });
  }

  // Tombol Rapihkan Dokumen dengan AI (di Ribbon & Banner)
  const btnAiPolishRibbon = container.querySelector('#btn-ai-polish-ribbon');
  if (btnAiPolishRibbon) {
    btnAiPolishRibbon.addEventListener('click', () => {
      triggerAiDocumentPolishing(container);
    });
  }

  const btnAiPolishBanner = container.querySelector('#btn-ai-polish-banner');
  if (btnAiPolishBanner) {
    btnAiPolishBanner.addEventListener('click', () => {
      triggerAiDocumentPolishing(container);
    });
  }

  // Banner Dokumen Impor: Saklar Mode Tampilan (Kanvas Rapi vs Pratinjau Asli Word)
  const btnModeCanvas = container.querySelector('#btn-mode-canvas');
  const btnModeNative = container.querySelector('#btn-mode-native');

  if (btnModeCanvas) {
    btnModeCanvas.addEventListener('click', () => {
      store.setState({ docxViewMode: 'canvas' }, false);
      renderWordCanvas(container);
    });
  }

  if (btnModeNative) {
    btnModeNative.addEventListener('click', () => {
      store.setState({ docxViewMode: 'native_preview' }, false);
      renderWordCanvas(container);
    });
  }

  // Jika dalam Mode Pratinjau Asli, render dengan docx-preview
  const activeDocxState = store.getState().activeUploadedDocx;
  const isPreviewState = (store.getState().docxViewMode === 'native_preview');
  if (activeDocxState && isPreviewState) {
    const mount = container.querySelector('#docx-preview-mount');
    if (mount && typeof window !== 'undefined' && window.docx && window.docx.renderAsync && activeDocxState.rawBytes) {
      window.docx.renderAsync(activeDocxState.rawBytes, mount, null, {
        className: 'docx-preview-wrapper',
        inWrapper: true,
        ignoreWidth: false,
        ignoreHeight: false,
        breakPages: true,
        useBase64URL: true
      }).catch(err => {
        mount.innerHTML = `<div style="padding: 20px; color: #dc2626; font-size: 13px;">Gagal merender pratinjau asli: ${err.message}</div>`;
      });
    }
  }

  // Banner Dokumen Impor: Tombol Simpan ke Gudang & Tutup
  const btnSaveImported = container.querySelector('#btn-save-imported-to-vault');
  if (btnSaveImported) {
    btnSaveImported.addEventListener('click', () => {
      const activeDocx = store.getState().activeUploadedDocx;
      if (!activeDocx) return;
      store.addLetterVaultEntry({
        title: activeDocx.parsedMeta?.title || activeDocx.fileName?.replace(/\.docx$/i, '') || 'Dokumen Word Impor',
        nomorSurat: activeDocx.parsedMeta?.nomor || '-',
        tanggalSurat: activeDocx.parsedMeta?.date || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
        category: 'Dokumen Impor',
        targetPerson: activeDocx.parsedMeta?.signer || '-',
        perihal: activeDocx.parsedMeta?.title,
        summary: `Dokumen .docx resmi diunggah: "${activeDocx.fileName}". Preservasi format 100%.`,
        isAiReference: true,
        source: 'upload',
        uploadedDocx: activeDocx
      });
      showToast(`Dokumen "${activeDocx.fileName}" tersimpan di Gudang Surat & aktif sebagai referensi AI!`, 'success');
      renderWordCanvas(container);
    });
  }

  const btnCloseImported = container.querySelector('#btn-close-imported-docx');
  if (btnCloseImported) {
    btnCloseImported.addEventListener('click', () => {
      store.clearActiveUploadedDocx();
      store.setState({ docxViewMode: 'canvas' }, false);
      showToast('Kembali ke template surat standar sekolah.', 'info');
      renderWordCanvas(container);
    });
  }

  // Edit Teks Paragraf & Field Berkas Impor (Non-Destructive)
  container.querySelectorAll('.imported-field, .imported-paragraph-item').forEach(item => {
    item.addEventListener('input', () => {
      const anchorIdx = parseInt(item.dataset.anchorIndex, 10);
      const active = store.getState().activeUploadedDocx;
      if (active && active.paragraphs && !isNaN(anchorIdx)) {
        const found = active.paragraphs.find(p => p.index === anchorIdx);
        if (found) {
          found.currentText = item.innerText;
          store.setActiveUploadedDocx(active);
        }
      }
    });
  });

  // Edit Baris Tabel Identitas/Jadwal Berkas Impor
  container.querySelectorAll('.identity-table-word tr[data-anchor-index]').forEach(row => {
    row.addEventListener('input', () => {
      const anchorIdx = parseInt(row.dataset.anchorIndex, 10);
      const active = store.getState().activeUploadedDocx;
      if (active && active.paragraphs && !isNaN(anchorIdx)) {
        const found = active.paragraphs.find(p => p.index === anchorIdx);
        if (found) {
          const lbl = row.querySelector('.id-label')?.innerText.trim() || '';
          let val = row.querySelector('.id-val')?.innerText.trim() || '';
          val = val.replace(/^:\s*/, '');
          found.currentText = `${lbl} : ${val}`;
          store.setActiveUploadedDocx(active);
        }
      }
    });
  });

  // Tautan ke Gudang Surat dari AI Copilot Drawer
  const linkOpenKb = container.querySelector('#link-open-vault-from-copilot-kb');
  if (linkOpenKb) {
    linkOpenKb.addEventListener('click', () => {
      const vaultTab = document.querySelector('.nav-tab[data-tab="vault"]');
      if (vaultTab) vaultTab.click();
    });
  }

  // Unduh .docx (Mendukung Dokumen Bawaan & Berkas Impor dengan Preservasi Format)
  container.querySelector('#btn-download-word').addEventListener('click', async () => {
    const activeDocx = store.getState().activeUploadedDocx;

    // Jika sedang mengedit berkas .docx luar yang diimpor, gunakan Non-Destructive Patcher
    if (activeDocx && activeDocx.rawBytes) {
      showToast('Menerapkan editan ke berkas Word asli (preservasi format 100% utuh)...', 'info');
      try {
        const patchedBlob = await applyCanvasEditsToDocx(activeDocx.rawBytes, activeDocx.paragraphs || [], { outputType: 'blob' });
        const safeName = (activeDocx.fileName || 'dokumen_teredit.docx').replace(/\.docx$/i, '') + '_edited.docx';
        downloadFile(patchedBlob, safeName);
        showToast('Berkas .docx hasil edit (format asli 100% terjaga) berhasil diunduh.', 'success');
        return;
      } catch (err) {
        showToast(`Gagal mem-patch berkas: ${err.message}`, 'error');
        return;
      }
    }

    const docModel = buildModelFromCanvas(container);
    showToast('Menyusun berkas Word .docx resmi...', 'info');
    try {
      const blob = await renderDocx(docModel, 'blob');
      const safeTitle = (docModel.meta?.title || 'surat_dinas')
        .replace(/[^a-zA-Z0-9_\-]/g, '_')
        .slice(0, 45);
      downloadFile(blob, `${safeTitle}.docx`);
      showToast('Berkas .docx berhasil diunduh.', 'success');
    } catch (err) {
      showToast(`Gagal mengunduh: ${err.message}`, 'error');
    }
  });

  // Print PDF
  container.querySelector('#btn-print-word').addEventListener('click', () => window.print());

  // Toggle AI Copilot Drawer
  const btnToggleCopilot = container.querySelector('#btn-toggle-copilot');
  const copilotDrawer = container.querySelector('#copilot-drawer');
  btnToggleCopilot.addEventListener('click', () => {
    copilotDrawer.classList.toggle('hidden');
  });
  container.querySelector('#btn-close-copilot').addEventListener('click', () => {
    copilotDrawer.classList.add('hidden');
  });
}

/**
 * Event handler Zoom Controls (Khusus Lembar Kanvas)
 */
function bindZoomControls(container) {
  const zoomSlider = container.querySelector('#input-zoom-slider');
  const zoomLabel = container.querySelector('#label-zoom-val');
  const sheetWrapper = container.querySelector('#word-sheet-wrapper');
  const btnZoomIn = container.querySelector('#btn-zoom-in');
  const btnZoomOut = container.querySelector('#btn-zoom-out');
  const btnZoomFit = container.querySelector('#btn-zoom-fit');

  const updateZoom = (val) => {
    const clamped = Math.max(50, Math.min(180, Math.round(val)));
    sheetWrapper.style.transform = `scale(${clamped / 100})`;
    zoomSlider.value = clamped;
    zoomLabel.textContent = `${clamped}%`;
    store.setZoomLevel(clamped);
  };

  zoomSlider.addEventListener('input', () => {
    updateZoom(parseFloat(zoomSlider.value));
  });

  btnZoomIn.addEventListener('click', () => {
    updateZoom(parseFloat(zoomSlider.value) + 10);
  });

  btnZoomOut.addEventListener('click', () => {
    updateZoom(parseFloat(zoomSlider.value) - 10);
  });

  btnZoomFit.addEventListener('click', () => {
    // Hitung perbandingan lebar viewport desk vs lembar kertas
    const desk = container.querySelector('#word-desk-container');
    const availableWidth = desk.clientWidth - 60;
    const sheetWidth = 210 * 3.78; // A4 approx px (794px)
    const fitZoom = Math.round((availableWidth / sheetWidth) * 100);
    updateZoom(Math.max(60, Math.min(130, fitZoom)));
    showToast(`Zoom disesuaikan dengan lebar layar (${fitZoom}%)`, 'info');
  });
}

/**
 * Event handler Gudang Kop Surat (Kop Gallery)
 */
function bindKopGalleryModal(container) {
  const modal = container.querySelector('#modal-kop-gallery');
  const btnOpen = container.querySelector('#btn-open-kop-gallery');
  const btnCloseHeader = container.querySelector('#btn-close-kop-modal');
  const btnCloseFooter = container.querySelector('#btn-close-kop-modal-footer');
  const cardsContainer = container.querySelector('#kop-gallery-cards-container');
  const triggerUploadBtn = container.querySelector('#btn-trigger-upload-modal');
  const uploadInput = container.querySelector('#input-modal-upload-kop');

  const renderGalleryCards = () => {
    const gallery = store.getKopGallery();
    const currentKopUrl = container.querySelector('#sheet-kop-image')?.src;

    cardsContainer.innerHTML = gallery.map(item => {
      const isCurrentlyActive = item.dataUrl === currentKopUrl;
      return `
        <div class="card" style="border: 2px solid ${isCurrentlyActive ? 'var(--primary)' : 'var(--border)'}; border-radius: 6px; padding: 12px; background: white; transition: all 0.15s;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
            <div style="display: flex; align-items: center; gap: 8px;">
              <span style="background: #e2e8f0; color: #334155; font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 4px;">${escapeHtml(item.category || 'Umum')}</span>
              <strong style="font-size: 13.5px; color: var(--dark);">${escapeHtml(item.name)}</strong>
              ${isCurrentlyActive ? '<span style="background: #dcfce7; color: #166534; font-size: 10.5px; font-weight: 600; padding: 1px 6px; border-radius: 4px;">✓ Digunakan di Lembar</span>' : ''}
            </div>
            <div style="display: flex; gap: 6px;">
              <button class="btn btn-sm ${isCurrentlyActive ? 'btn-success' : 'btn-primary'} btn-select-kop-item" data-id="${item.id}">
                ${isCurrentlyActive ? '✓ Sudah Aktif' : 'Terapkan Kop Ini'}
              </button>
              ${!item.isDefault ? `
                <button class="btn btn-sm btn-danger btn-delete-kop-item" data-id="${item.id}">🗑</button>
              ` : ''}
            </div>
          </div>

          <!-- Pratinjau Banner Kop -->
          <div style="border: 1px solid #e2e8f0; border-radius: 4px; overflow: hidden; background: #fafafa; padding: 6px;">
            <img src="${item.dataUrl}" alt="${escapeHtml(item.name)}" style="width: 100%; height: auto; max-height: 85px; object-fit: contain; display: block;">
          </div>
          <p style="font-size: 11.5px; color: var(--text-muted); margin-top: 6px;">${escapeHtml(item.description || '')}</p>
        </div>
      `;
    }).join('');

    // Bind event tombol terapkan
    cardsContainer.querySelectorAll('.btn-select-kop-item').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.id;
        store.applyKopFromGallery(id);
        const kop = store.getKopGallery().find(k => k.id === id);
        if (kop) {
          const kopImg = container.querySelector('#sheet-kop-image');
          kopImg.src = kop.dataUrl;
          container.querySelector('#sheet-kop-container').classList.remove('hidden');
          showToast(`Kop surat diubah ke: ${kop.name}`, 'success');
          scheduleSync(container);
        }
        modal.classList.add('hidden');
      });
    });

    // Bind event hapus
    cardsContainer.querySelectorAll('.btn-delete-kop-item').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.id;
        if (confirm('Hapus kop surat ini dari galeri tersimpan?')) {
          store.removeKopGalleryEntry(id);
          showToast('Kop dihapus dari galeri.', 'info');
          renderGalleryCards();
        }
      });
    });
  };

  btnOpen.addEventListener('click', () => {
    renderGalleryCards();
    modal.classList.remove('hidden');
  });

  const closeModal = () => modal.classList.add('hidden');
  btnCloseHeader.addEventListener('click', closeModal);
  btnCloseFooter.addEventListener('click', closeModal);

  // Upload baru ke galeri
  triggerUploadBtn.addEventListener('click', () => uploadInput.click());
  uploadInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (file.size > 3 * 1024 * 1024) {
      showToast('Ukuran gambar kop melebihi 3 MB.', 'error');
      return;
    }

    const reader = new FileReader();
    reader.onload = (ev) => {
      const dataUrl = ev.target.result;
      const customName = prompt('Beri nama kop surat baru ini:', file.name.replace(/\.[^.]+$/, '')) || 'Kop Sekolah Baru';

      store.addKopGalleryEntry({
        name: customName,
        dataUrl,
        category: 'Kustom Pengguna'
      });

      showToast(`Kop "${customName}" berhasil ditambahkan ke galeri!`, 'success');
      renderGalleryCards();
    };
    reader.readAsDataURL(file);
  });
}

/**
 * Event handler Modal Pengaturan Dokumen Lengkap (Ukuran, Orientasi, Margin, Font, Kop)
 */
function bindDocSettingsModal(container) {
  const modal = container.querySelector('#modal-doc-settings');
  const btnOpen = container.querySelector('#btn-open-doc-settings');
  const btnClose = container.querySelector('#btn-close-settings-modal');
  const btnCancel = container.querySelector('#btn-cancel-doc-settings');
  const btnSave = container.querySelector('#btn-save-doc-settings');

  if (!modal) return;

  const closeModal = () => modal.classList.add('hidden');
  const openModal = () => modal.classList.remove('hidden');

  if (btnOpen) btnOpen.addEventListener('click', openModal);
  if (btnClose) btnClose.addEventListener('click', closeModal);
  if (btnCancel) btnCancel.addEventListener('click', closeModal);

  // Preset Margin Buttons di dalam Modal
  modal.querySelectorAll('.btn-modal-preset-margin').forEach(btn => {
    btn.addEventListener('click', () => {
      modal.querySelector('#modal-set-m-top').value = btn.dataset.top;
      modal.querySelector('#modal-set-m-bottom').value = btn.dataset.bottom;
      modal.querySelector('#modal-set-m-left').value = btn.dataset.left;
      modal.querySelector('#modal-set-m-right').value = btn.dataset.right;
    });
  });

  // Tombol Simpan Pengaturan
  if (btnSave) {
    btnSave.addEventListener('click', () => {
      const paperSize = modal.querySelector('#modal-set-paper-size').value;
      const orientation = modal.querySelector('#modal-set-orientation').value;
      const top = parseFloat(modal.querySelector('#modal-set-m-top').value) || 2.0;
      const bottom = parseFloat(modal.querySelector('#modal-set-m-bottom').value) || 2.0;
      const left = parseFloat(modal.querySelector('#modal-set-m-left').value) || 2.5;
      const right = parseFloat(modal.querySelector('#modal-set-m-right').value) || 2.0;
      const fontFamily = modal.querySelector('#modal-set-font-family').value;
      const fontSize = modal.querySelector('#modal-set-font-size').value;
      const lineSpacing = modal.querySelector('#modal-set-line-spacing').value;
      const showKop = modal.querySelector('#modal-set-show-kop').checked;

      store.updateLetterOptions({
        paperSize,
        orientation,
        margins: { top, bottom, left, right },
        fontFamily,
        fontSize,
        lineSpacing,
        showKop
      });

      closeModal();
      showToast('✓ Pengaturan format & tata letak dokumen berhasil diterapkan!', 'success');
      renderWordCanvas(container);
    });
  }
}

/**
 * Memicu proses perapihan dokumen yang berantakan dengan AI
 */
async function triggerAiDocumentPolishing(container) {
  const activeDocx = store.getState().activeUploadedDocx;
  const currentDoc = store.getState().document;

  let textToPolish = '';
  if (activeDocx && activeDocx.paragraphs) {
    textToPolish = activeDocx.paragraphs.map(p => p.currentText || p.originalText).filter(Boolean).join('\n\n');
  } else if (currentDoc) {
    const title = currentDoc.metadata?.title || currentDoc.blocks?.find(b => b.type === 'letterTitle')?.title || '';
    const nomor = currentDoc.metadata?.nomorSurat || currentDoc.blocks?.find(b => b.type === 'letterTitle')?.nomor || '';
    const paras = (currentDoc.blocks || [])
      .filter(b => b.type === 'paragraph')
      .map(b => b.runs ? b.runs.map(r => r.text).join('') : (b.text || ''))
      .filter(Boolean)
      .join('\n\n');
    const tableRows = (currentDoc.blocks || [])
      .find(b => b.type === 'identityTable')
      ?.rows?.map(r => `${r.label}: ${r.value}`)
      .join('\n') || '';
    textToPolish = `${title}\n${nomor}\n\n${paras}\n\n${tableRows}`;
  }

  if (!textToPolish || textToPolish.trim().length < 10) {
    showToast('Dokumen masih kosong atau terlalu pendek untuk dirapihkan oleh AI.', 'warning');
    return;
  }

  showToast('✨ AI sedang menganalisis & menata ulang dokumen ke format dinas baku...', 'info');

  try {
    const polishedDoc = await polishMessyDocumentWithAi(textToPolish, {
      schoolProfile: store.getSchoolProfile(),
      paperSize: store.getLetterOptions().paperSize,
      margins: store.getLetterOptions().margins,
      preferredApiId: store.getState().activeApiId
    });

    store.clearActiveUploadedDocx();
    store.setState({ document: polishedDoc }, true);
    showToast('✓ Berhasil! Dokumen yang berantakan telah dirapihkan oleh AI ke standar resmi.', 'success');
    renderWordCanvas(container);
  } catch (err) {
    showToast(`❌ Gagal merapihkan dokumen dengan AI: ${err.message}`, 'error');
  }
}

/**
 * Event handler AI Copilot
 */
function bindCopilotDrawer(container) {
  const promptInput = container.querySelector('#copilot-prompt-input');
  const btnRun = container.querySelector('#btn-run-copilot');
  const statusArea = container.querySelector('#copilot-status');
  const statusText = container.querySelector('#copilot-status-text');
  const statusTip = container.querySelector('#copilot-status-tip');
  const btnCancel = container.querySelector('#btn-cancel-copilot');
  const modelSelect = container.querySelector('#select-copilot-model');
  const linkVault = container.querySelector('#link-open-vault-from-copilot');

  let copilotAbortController = null;

  // Tombol Batal
  if (btnCancel) {
    btnCancel.addEventListener('click', () => {
      if (copilotAbortController) {
        copilotAbortController.abort();
        copilotAbortController = null;
        showToast('Pembuatan surat dibatalkan.', 'info');
      }
    });
  }

  // Shortcut ke tab pengaturan gudang API
  linkVault.addEventListener('click', () => {
    const settingsTabBtn = document.querySelector('.nav-tab[data-tab="settings"]');
    if (settingsTabBtn) settingsTabBtn.click();
  });

  // Chips click
  container.querySelectorAll('.copilot-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      promptInput.value = chip.dataset.prompt;
      promptInput.focus();
    });
  });

  // Jalankan AI Copilot
  btnRun.addEventListener('click', async () => {
    const userPrompt = promptInput.value.trim();
    if (!userPrompt) {
      showToast('Ketik perintah pembuatan surat terlebih dahulu.', 'warning');
      return;
    }

    const selectedApiId = modelSelect.value;
    const vault = store.getApiVault();
    const activeEntry = vault.find(v => v.id === selectedApiId) || store.getActiveApi();
    const modelLabel = activeEntry?.label || activeEntry?.model || 'AI';

    copilotAbortController = new AbortController();
    btnRun.disabled = true;
    statusArea.classList.remove('hidden');
    if (statusTip) statusTip.style.display = 'none';

    let elapsed = 0;
    statusText.textContent = `Menghubungkan ke ${modelLabel}... (0d)`;

    const timerInterval = setInterval(() => {
      elapsed++;
      statusText.textContent = `Merangkai draf dokumen (${modelLabel})... (${elapsed}d)`;
      if (elapsed >= 4 && statusTip) {
        statusTip.style.display = 'block';
      }
    }, 1000);

    try {
      const { callAiWithValidation } = await import('../ai/client.js');
      const { COPILOT_SYSTEM_PROMPT, buildCopilotUserPrompt } = await import('../ai/prompts.js');

      const state = store.getState();
      const prompt = buildCopilotUserPrompt({
        userPrompt,
        schoolName: state.schoolProfile?.schoolName,
        principalName: state.schoolProfile?.principalName,
        dateStr: new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
        references: store.getAiReferences(3)
      });

      const aiResult = await callAiWithValidation({
        prompt,
        systemPrompt: COPILOT_SYSTEM_PROMPT,
        preferredApiId: selectedApiId,
        signal: copilotAbortController.signal,
        validator: (data) => {
          if (!data || !data.letterTitle || !Array.isArray(data.blocks)) {
            return { valid: false, errors: ['Output AI harus menyertakan letterTitle dan blocks'] };
          }
          return { valid: true, errors: [] };
        }
      });

      // Terapkan hasil AI langsung ke lembar kerja Word!
      applyAiResultToCanvas(container, aiResult);
      showToast('Surat berhasil dibuat oleh AI Copilot!', 'success');

    } catch (err) {
      if (copilotAbortController?.signal?.aborted || err.message.includes('dibatalkan')) {
        // Abaikan atau telah dihandle toast pembatalan
      } else {
        showToast(`Gagal membuat surat: ${err.message}`, 'error');
      }
    } finally {
      clearInterval(timerInterval);
      copilotAbortController = null;
      btnRun.disabled = false;
      statusArea.classList.add('hidden');
    }
  });

  // Aksi Rapihkan Dokumen yang Berantakan (AI Polisher)
  const btnCopilotPolish = container.querySelector('#btn-copilot-polish-messy');
  if (btnCopilotPolish) {
    btnCopilotPolish.addEventListener('click', () => {
      triggerAiDocumentPolishing(container);
    });
  }

  // Aksi Rapikan Bahasa (PUEBI)
  container.querySelector('#btn-copilot-formalize').addEventListener('click', async () => {
    const currentDoc = buildModelFromCanvas(container);
    const selectedApiId = modelSelect.value;
    showToast('AI sedang menyempurnakan redaksi bahasa dinas (PUEBI)...', 'info');
    try {
      const { callAiApi, cleanJsonOutput } = await import('../ai/client.js');
      const raw = await callAiApi({
        prompt: `Rapikan redaksi kalimat dari surat dinas berikut agar menggunakan tata bahasa baku, resmi, sopan, dan sesuai PUEBI tanpa merubah fakta, nama, atau tanggal:\n${JSON.stringify(currentDoc.blocks)}`,
        systemPrompt: 'Kembalikan format JSON array blocks yang telah disempurnakan redaksinya.',
        preferredApiId: selectedApiId
      });
      const parsed = JSON.parse(cleanJsonOutput(raw));
      if (Array.isArray(parsed)) {
        currentDoc.blocks = parsed;
        store.setState({ document: currentDoc }, true);
        renderWordCanvas(container);
        showToast('Redaksi surat telah dirapikan.', 'success');
      }
    } catch (err) {
      showToast(`Gagal merapikan bahasa: ${err.message}`, 'error');
    }
  });

  // Aksi Perbarui Tanggal ke Hari Ini
  container.querySelector('#btn-copilot-today-date').addEventListener('click', () => {
    const sigDateEl = container.querySelector('#editable-sig-date');
    if (sigDateEl) {
      const city = store.getState().schoolProfile?.city || 'Garut';
      const todayFormatted = new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
      sigDateEl.innerText = `${city}, ${todayFormatted}`;
      showToast('Tanggal surat diperbarui ke hari ini.', 'success');
      scheduleSync(container);
    }
  });
}

function applyNewMargins(container, margins) {
  const sheet = container.querySelector('#word-sheet');
  sheet.style.setProperty('--margin-top', `${margins.top}cm`);
  sheet.style.setProperty('--margin-bottom', `${margins.bottom}cm`);
  sheet.style.setProperty('--margin-left', `${margins.left}cm`);
  sheet.style.setProperty('--margin-right', `${margins.right}cm`);

  container.querySelector('#btn-margin-menu').textContent = `📐 Margin: ${margins.left}cm • ${margins.top}cm ▾`;
  store.updateLetterOptions({ margins });
  showToast(`Margin lembar diatur: Kiri ${margins.left}cm, Kanan ${margins.right}cm, Atas ${margins.top}cm, Bawah ${margins.bottom}cm`, 'success');
  scheduleSync(container);
}

/**
 * Sinkronisasi pengeditan langsung di lembar kerja
 */
function bindCanvasLiveEditing(container) {
  const editables = container.querySelectorAll('[contenteditable="true"]');
  const metaWrapper = container.querySelector('#sheet-meta-recipient-wrapper');
  const metaEl = container.querySelector('#editable-meta-recipient');
  const btnClearMeta = container.querySelector('#btn-clear-meta-recipient');
  const btnRestoreMeta = container.querySelector('#btn-restore-meta-recipient');

  const syncMetaEmptyState = () => {
    if (!metaEl || !metaWrapper) return;
    // Bersihkan sub-elemen kosong bila user menghapus sebagian lewat Backspace/Delete
    const tbl = metaEl.querySelector('.meta-perihal-table');
    if (tbl && tbl.innerText.replace(/[:\s\u00a0]/g, '') === '') {
      tbl.remove();
    }
    const addr = metaEl.querySelector('.recipient-address-block');
    if (addr && addr.innerText.replace(/\s|\u00a0/g, '') === '') {
      addr.remove();
    }

    const plain = metaEl.innerText.replace(/\s|\u00a0/g, '');
    if (!plain) {
      metaEl.innerHTML = '';
      metaWrapper.classList.add('is-empty');
      if (btnClearMeta) btnClearMeta.classList.add('hidden');
      if (btnRestoreMeta) btnRestoreMeta.classList.remove('hidden');
    } else {
      metaWrapper.classList.remove('is-empty');
      if (btnClearMeta) btnClearMeta.classList.remove('hidden');
      if (btnRestoreMeta) btnRestoreMeta.classList.add('hidden');
    }
  };

  if (btnClearMeta && metaEl) {
    btnClearMeta.addEventListener('click', (e) => {
      e.preventDefault();
      metaEl.innerHTML = '';
      syncMetaEmptyState();
      scheduleSync(container);
      updateWordCount(container);
      showToast('Bagian Sifat, Perihal & Kepada Yth dihapus dari kanvas.', 'info');
    });
  }

  if (btnRestoreMeta && metaEl) {
    btnRestoreMeta.addEventListener('click', (e) => {
      e.preventDefault();
      metaEl.innerHTML = buildMetaRecipientInnerHtml(
        [
          { label: 'Sifat', value: 'Penting' },
          { label: 'Perihal', value: 'Undangan Rapat' }
        ],
        [
          'Kepada',
          'Yth. ...............',
          '__________________',
          'Di',
          'Tempat'
        ]
      );
      syncMetaEmptyState();
      scheduleSync(container);
      updateWordCount(container);
      showToast('Bagian Sifat, Perihal & Kepada Yth ditampilkan kembali.', 'success');
    });
  }

  // Dropdown Pemilih Cepat Kode Nomor Surat (UND, TU, KEP, KET, ST, PGL, dll)
  const quickKodeSelect = container.querySelector('#select-quick-kode-surat');
  if (quickKodeSelect) {
    quickKodeSelect.addEventListener('change', () => {
      const selectedCode = quickKodeSelect.value;
      if (!selectedCode) return;
      const nomorEl = container.querySelector('#editable-doc-nomor');
      if (nomorEl) {
        const currentNomor = nomorEl.innerText.trim();
        const updatedNomor = formatSmartNomorSurat({
          rawNomor: currentNomor,
          forcedTypeCode: selectedCode
        });
        nomorEl.innerText = `Nomor: ${updatedNomor}`;
        showToast(`Kode nomor surat diubah ke ${selectedCode} (${updatedNomor})`, 'success');
        scheduleSync(container);
      }
      quickKodeSelect.value = '';
    });
  }

  editables.forEach(el => {
    el.addEventListener('input', () => {
      if (el === metaEl) {
        syncMetaEmptyState();
      }
      scheduleSync(container);
      updateWordCount(container);
    });
  });
}

function updateWordCount(container) {
  const bodyText = container.querySelector('#sheet-body-content')?.innerText || '';
  const words = bodyText.trim().split(/\s+/).filter(w => w.length > 0).length;
  const countEl = container.querySelector('#status-word-count');
  if (countEl) countEl.textContent = `${words} kata`;
}

function scheduleSync(container) {
  clearTimeout(syncTimer);
  syncTimer = setTimeout(() => {
    if (container) {
      const updatedDoc = buildModelFromCanvas(container);
      store.setState({ document: updatedDoc }, false);
    }
  }, 400);
}

/**
 * Helper pemformatan teks kustom
 */
function applyInlineStyleToSelection(cssProp, cssValue) {
  const sel = window.getSelection();
  if (!sel.rangeCount) return;
  const range = sel.getRangeAt(0);

  if (range.collapsed) return;

  const span = document.createElement('span');
  span.style[cssProp] = cssValue;
  try {
    range.surroundContents(span);
  } catch (e) {
    // Selection crosses multiple blocks
    document.execCommand('styleWithCSS', false, true);
    if (cssProp === 'font-size') {
      document.execCommand('fontSize', false, '3');
    }
  }
}

function adjustFontSize(delta) {
  const sel = window.getSelection();
  if (!sel.rangeCount) return;
  const range = sel.getRangeAt(0);
  let parent = range.commonAncestorContainer;
  if (parent.nodeType === 3) parent = parent.parentNode;

  const currentSize = parseFloat(window.getComputedStyle(parent).fontSize) || 16;
  const newSizePt = Math.max(8, Math.min(72, Math.round((currentSize * 0.75) + delta)));
  applyInlineStyleToSelection('font-size', `${newSizePt}pt`);
}

function toggleCaseOnSelection() {
  const sel = window.getSelection();
  if (!sel.rangeCount) return;
  const range = sel.getRangeAt(0);
  const text = range.toString();
  if (!text) return;

  let converted;
  if (text === text.toUpperCase()) {
    converted = text.toLowerCase();
  } else if (text === text.toLowerCase()) {
    converted = text.replace(/\b\w/g, c => c.toUpperCase());
  } else {
    converted = text.toUpperCase();
  }

  document.execCommand('insertText', false, converted);
}

/**
 * Menyusun Document Model dari isi lembar kerja saat ini
 */
function buildModelFromCanvas(container) {
  const state = store.getState();
  const paperSize = state.letterOptions?.paperSize === 'f4' ? 'suratDinasF4' : 'suratDinas';
  const margins = state.letterOptions?.margins || { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 };
  const showKop = state.letterOptions?.showKop !== false;

  const titleText = container.querySelector('#editable-doc-title')?.innerText.trim() || 'SURAT DINAS';
  const nomorText = container.querySelector('#editable-doc-nomor')?.innerText.trim() || '';

  // Parse Sifat, Perihal & Kepada Yth
  const metaEl = container.querySelector('#editable-meta-recipient');
  const rawMetaText = metaEl ? metaEl.innerText.replace(/\u00a0/g, ' ').trim() : '';
  let metaRecipientData = { isDeleted: true, metaRows: [], recipientLines: [] };

  if (metaEl && rawMetaText.length > 0) {
    const metaRows = [];
    const recipientLines = [];

    const tableEl = metaEl.querySelector('.meta-perihal-table');
    if (tableEl) {
      tableEl.querySelectorAll('tr').forEach(tr => {
        const lbl = tr.querySelector('.meta-label')?.innerText.trim() || '';
        const val = tr.querySelector('.meta-val')?.innerText.trim().replace(/^:\s*/, '') || '';
        if (lbl || val) {
          metaRows.push({ label: lbl, value: val });
        }
      });
    }

    const addrEl = metaEl.querySelector('.recipient-address-block');
    if (addrEl) {
      const lines = addrEl.innerText.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
      recipientLines.push(...lines);
    }

    if (!tableEl && !addrEl) {
      const allLines = rawMetaText.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
      for (const line of allLines) {
        const m = line.match(/^(sifat|lampiran|perihal|hal)\s*:\s*(.*)$/i);
        if (m) {
          metaRows.push({ label: m[1].trim(), value: m[2].trim() });
        } else {
          recipientLines.push(line);
        }
      }
    }

    if (metaRows.length > 0 || recipientLines.length > 0) {
      metaRecipientData = {
        isDeleted: false,
        metaRows,
        recipientLines,
        customHtml: metaEl.innerHTML
      };
    }
  }

  const pOpening = container.querySelector('#editable-p-opening')?.innerText.trim() || '';
  const pBody = container.querySelector('#editable-p-body')?.innerText.trim() || '';
  const pPurpose = container.querySelector('#editable-p-purpose')?.innerText.trim() || '';
  const pClosing = container.querySelector('#editable-p-closing')?.innerText.trim() || '';

  // Parse identity table
  const identityRows = [];
  container.querySelectorAll('.identity-table-word tr').forEach(tr => {
    const label = tr.querySelector('.id-label')?.innerText.trim() || '';
    const val = tr.querySelector('.id-val')?.innerText.trim() || '';
    if (label || val) {
      identityRows.push({ label, value: val });
    }
  });

  // Parse signature
  const sigDateStr = container.querySelector('#editable-sig-date')?.innerText.trim() || '';
  const sigTitle = container.querySelector('#editable-sig-title')?.innerText.trim() || 'Kepala Sekolah';
  const sigName = container.querySelector('#editable-sig-name')?.innerText.trim() || 'Dian Nurhadian, S.Pd.';
  const sigNip = container.querySelector('#editable-sig-nip')?.innerText.trim().replace(/^NIP\.?\s*/i, '') || '';

  const kopImgEl = container.querySelector('#sheet-kop-image');
  const kopImageUrl = kopImgEl?.src || DEFAULT_KOP_DATA_URL;

  // Twips conversion (1 cm ≈ 567 twips)
  const twipMargins = {
    top: Math.round(margins.top * 567),
    bottom: Math.round(margins.bottom * 567),
    left: Math.round(margins.left * 567),
    right: Math.round(margins.right * 567)
  };

  const blocks = [];

  // 1. Kop Surat
  if (showKop) {
    blocks.push({
      type: 'letterKop',
      image: kopImageUrl,
      widthCm: 16.5,
      heightCm: 3.3
    });
  }

  // 2. Title, Nomor & Sifat/Perihal/Kepada Yth
  blocks.push({
    type: 'letterTitle',
    title: titleText,
    nomor: nomorText,
    metaRecipient: metaRecipientData.isDeleted ? null : metaRecipientData
  });

  // 3. Paragraf Pembuka
  if (pOpening) {
    blocks.push({ type: 'paragraph', style: 'letterBody', text: pOpening });
  }

  // 4. Tabel Identitas
  if (identityRows.length > 0) {
    blocks.push({ type: 'identityTable', rows: identityRows });
  }

  // 5. Paragraf Isi & Maksud
  if (pBody) {
    blocks.push({ type: 'paragraph', style: 'letterBody', text: pBody });
  }
  if (pPurpose) {
    blocks.push({ type: 'paragraph', style: 'letterBody', text: pPurpose });
  }

  // 6. Paragraf Penutup
  if (pClosing) {
    blocks.push({ type: 'paragraph', style: 'letterBody', text: pClosing });
  }

  // 7. Signature Block
  blocks.push({
    type: 'letterSignature',
    city: sigDateStr.split(',')[0]?.trim() || 'Garut',
    date: sigDateStr.split(',')[1]?.trim() || '14 November 2026',
    title: sigTitle,
    name: sigName,
    nip: sigNip
  });

  return {
    meta: {
      title: titleText,
      author: sigName,
      year: new Date().getFullYear().toString()
    },
    profile: paperSize,
    customMargins: twipMargins,
    letterData: {
      kopImage: kopImageUrl,
      metaRecipientDeleted: metaRecipientData.isDeleted
    },
    blocks
  };
}

/**
 * Helper menyusun HTML dalam blok Sifat, Perihal & Kepada Yth
 */
function buildMetaRecipientInnerHtml(metaRows = [], recipientLines = []) {
  const hasMeta = Array.isArray(metaRows) && metaRows.length > 0;
  const hasRecipient = Array.isArray(recipientLines) && recipientLines.length > 0;
  if (!hasMeta && !hasRecipient) return '';

  const metaTableHtml = hasMeta ? `
    <table class="meta-perihal-table">
      <tbody>
        ${metaRows.map(r => {
          const lbl = (r.label || '').trim();
          const val = (r.value || '').trim().replace(/^:\s*/, '');
          const isPerihal = /^(perihal|hal)$/i.test(lbl);
          return `
            <tr>
              <td class="meta-label">${escapeHtml(lbl)}</td>
              <td class="meta-colon">:</td>
              <td class="meta-val">${isPerihal ? `<b><i>${escapeHtml(val)}</i></b>` : escapeHtml(val)}</td>
            </tr>
          `;
        }).join('')}
      </tbody>
    </table>
  ` : '';

  const recipientHtml = hasRecipient ? `
    <div class="recipient-address-block">
      ${recipientLines.map(line => {
        const t = (line || '').trim();
        const isDi = /^di$/i.test(t);
        const isLine = /^_{3,}$/.test(t);
        return `<div class="${isDi ? 'recipient-di' : (isLine ? 'recipient-line' : '')}">${escapeHtml(t)}</div>`;
      }).join('')}
    </div>
  ` : '';

  return `${metaTableHtml}${recipientHtml}`;
}

/**
 * Memasukkan hasil AI Copilot langsung ke lembar kerja Word
 */
function applyAiResultToCanvas(container, aiResult) {
  // Title & Nomor (Otomatis sesuaikan kode jenis surat: UND, TU, KEP, KET, ST, PGL, dll)
  const titleEl = container.querySelector('#editable-doc-title');
  const nomorEl = container.querySelector('#editable-doc-nomor');
  const letterTitle = aiResult.letterTitle || 'SURAT DINAS';
  const smartNomor = formatSmartNomorSurat({
    rawNomor: aiResult.nomorSurat || '',
    contextText: `${letterTitle} ${aiResult.hal || ''} ${aiResult.perihal || ''} ${aiResult.title || ''}`,
    forcedTypeCode: aiResult.kodeJenisSurat || ''
  });

  if (titleEl) titleEl.innerText = letterTitle;
  if (nomorEl) nomorEl.innerText = `Nomor: ${smartNomor}`;

  // Sifat, Perihal & Kepada Yth
  const metaEl = container.querySelector('#editable-meta-recipient');
  const metaWrapper = container.querySelector('#sheet-meta-recipient-wrapper');
  if (metaEl) {
    const sifatVal = aiResult.sifat || 'Penting';
    const perihalVal = aiResult.hal || aiResult.perihal || aiResult.title || 'Undangan Rapat';
    const tujuanVal = aiResult.tujuan
      ? (aiResult.tujuan.toLowerCase().startsWith('yth') ? aiResult.tujuan : `Yth. ${aiResult.tujuan}`)
      : 'Yth. ...............';
    metaEl.innerHTML = buildMetaRecipientInnerHtml(
      [
        { label: 'Sifat', value: sifatVal },
        { label: 'Perihal', value: perihalVal }
      ],
      [
        'Kepada',
        tujuanVal,
        '__________________',
        'Di',
        'Tempat'
      ]
    );
    if (metaWrapper) metaWrapper.classList.remove('is-empty');
    const btnClear = container.querySelector('#btn-clear-meta-recipient');
    const btnRestore = container.querySelector('#btn-restore-meta-recipient');
    if (btnClear) btnClear.classList.remove('hidden');
    if (btnRestore) btnRestore.classList.add('hidden');
  }

  // Extract paragraphs & identity table
  const paragraphs = aiResult.blocks.filter(b => b.type === 'paragraph');
  const idTable = aiResult.blocks.find(b => b.type === 'identityTable');

  const pOpen = container.querySelector('#editable-p-opening');
  const pBody = container.querySelector('#editable-p-body');
  const pPurpose = container.querySelector('#editable-p-purpose');
  const pClose = container.querySelector('#editable-p-closing');

  if (pOpen) pOpen.innerText = paragraphs[0]?.text || '';
  if (pBody) pBody.innerText = paragraphs[1]?.text || '';
  if (pPurpose) pPurpose.innerText = paragraphs[2]?.text || '';
  if (pClose) pClose.innerText = paragraphs[3]?.text || (paragraphs[2]?.text ? 'Demikian surat ini dibuat agar dapat dipergunakan sebagaimana mestinya.' : '');

  // Identity table rows
  if (idTable && Array.isArray(idTable.rows)) {
    const tbody = container.querySelector('.identity-table-word tbody');
    if (tbody) {
      tbody.innerHTML = idTable.rows.map(r => `
        <tr>
          <td class="id-label" contenteditable="true" spellcheck="false">${escapeHtml(r.label)}</td>
          <td class="id-val" contenteditable="true" spellcheck="false">${escapeHtml(r.value)}</td>
        </tr>
      `).join('');
    }
  }

  // Update date
  const sigDateEl = container.querySelector('#editable-sig-date');
  if (sigDateEl && aiResult.tanggalSurat) {
    const city = store.getState().schoolProfile?.city || 'Garut';
    sigDateEl.innerText = `${city}, ${aiResult.tanggalSurat}`;
  }

  scheduleSync(container);
  updateWordCount(container);
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str || '';
  return div.innerHTML;
}
