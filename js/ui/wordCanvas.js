/**
 * wordCanvas.js - Antarmuka Lembar Kerja Microsoft Word Asli & AI Copilot Persuratan
 * Fitur:
 * 1. Ribbon 2-Row lengkap standar Microsoft Word (Font & Paragraph lengkap sesuai Screenshot).
 * 2. Pratinjau skala nyata A4 & F4 dengan pengaturan margin tanpa clipping / z-index tepat.
 * 3. Zoom In/Out khusus kanvas lembar kerja, menjaga ribbon dan sidebar tetap tetap.
 * 4. Gudang Kop Surat (Galeri Kop) dengan pratinjau thumbnail dan pemilihan 1-klik.
 * 5. Integrasi Gudang API di AI Copilot dengan pemilih model real-time.
 */

import { store, createEmptyDocumentModel, computeDefaultDesktopZoom } from '../state.js';
import { renderDocx } from '../docx/render.js';
import { parseDocxForEditing, applyCanvasEditsToDocx, convertDocxToDocumentModel } from '../docx/patch.js';
import { polishMessyDocumentWithAi, editDocumentWithAi, convertAiBlocksToDocumentBlocks } from '../ai/client.js';
import { shouldAllowLayoutRestructure, buildSurgicalIndexFromDocModel, applySurgicalPatchesToDocModel } from '../ai/surgicalEdit.js';
import { summarizeKnowledgeBaseStats, describeAiMatchFeedback } from '../ai/knowledgeBase.js';
import { renderCanvasVaultBadge, openLockedMasterSaveChoiceModal } from './vaultSaveHelper.js';
import { LETTER_TYPE_CODES, formatSmartNomorSurat } from '../letters/numbering.js';
import { applySignatureModeToDocModel, buildSignerSelectionInfo, formatSignerNipLine } from '../letters/verification.js';
import { buildSuratTugasSppdDocModel } from '../letters/templates/suratTugasSppd.js';
import { attachInteractiveTables } from './tableInteraction.js';
import {
  computePrintableDimsMm,
  mountInteractiveFloatingQr,
  updateCanvasQrToggleButton,
  normalizeQrPlacement,
  findPreferredSignatureQrPlacement
} from './qrFloatingHelper.js';
import { canSignWithRole, resolveDefaultAllowedSignerRole } from '../auth/permissions.js';
import {
  openPrintSignatureChoiceModal,
  buildSignerRoleSelectorHtml,
  authorizeAndRecordOutgoing,
  requestAccForOutgoing,
  printWithTemporaryQrOverlay,
  syncGlobalPrintPageStyle
} from './printChoiceModal.js';
import { downloadFile } from '../utils/download.js';
import { showToast } from './notification.js';
import { DEFAULT_KOP_DATA_URL } from '../letters/defaultKop.js';
import {
  ICON_GALLERY, ICON_EYE, ICON_EYE_OFF, ICON_UPLOAD,
  ICON_SAVE, ICON_SEND_OUT, ICON_ARCHIVE, ICON_DOC_UPLOAD, ICON_CONVERT,
  ICON_SPARKLE, ICON_WAND, ICON_DOWNLOAD, ICON_PRINTER,
  ICON_SETTINGS, ICON_RULER, ICON_SCHOOL, ICON_SHADING, ICON_DOC,
  ICON_ALIGN_LEFT, ICON_ALIGN_CENTER, ICON_ALIGN_RIGHT, ICON_ALIGN_JUSTIFY,
  ICON_LIST_BULLET, ICON_LIST_NUMBER, ICON_OUTDENT, ICON_INDENT, ICON_ERASER,
  ICON_TABLE, ICON_UNDO, ICON_REDO, ICON_PILCROW, ICON_GLOBE, ICON_MAXIMIZE,
  ICON_PORTRAIT, ICON_MAIL, ICON_FOLDER, ICON_BUILDING, ICON_ID_CARD,
  ICON_BRIEFCASE, ICON_USERS, ICON_CALENDAR, ICON_TARGET, ICON_PLUS,
  ICON_BOT_SPARK, ICON_OPEN_BOOK, ICON_PENCIL_NOTE, ICON_REFRESH, ICON_COPY_DOC,
  ICON_CLIPBOARD_CHECK
} from './icons.js';

let syncTimer = null;
let wordPaginateTimer = null;
let isWordPaginating = false;
let isCopilotOpenState = null;

export function renderWordCanvas(container) {
  if (!container) return;
  syncGlobalPrintPageStyle(store.getState().letterOptions);

  // Pastikan scroll jendela utama selalu di titik (0,0) agar bagian kiri header & ribbon tidak pernah tergeser keluar layar
  if ('scrollRestoration' in history) {
    history.scrollRestoration = 'manual';
  }
  window.scrollTo(0, 0);
  document.documentElement.scrollLeft = 0;
  document.body.scrollLeft = 0;

  const state = store.getState();
  // Pastikan nomor surat di lembar kerja selalu otomatis mengikuti nomor urut Surat Keluar berikutnya bila masih default atau sudah tercatat keluar
  store.syncActiveDocumentNumberToNextOutgoing(false);
  const freshState = store.getState();
  const doc = freshState.document;
  const paperSize = state.letterOptions?.paperSize || 'a4';
  const orientation = state.letterOptions?.orientation || 'portrait';
  const customMargins = state.letterOptions?.margins || { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 };
  const fontFamily = state.letterOptions?.fontFamily || 'Times New Roman';
  const fontSize = state.letterOptions?.fontSize || '12';
  const lineSpacing = state.letterOptions?.lineSpacing || '1.15';
  const showKop = state.letterOptions?.showKop !== false;
  const zoomLevel = state.zoomUserAdjusted ? (state.zoomLevel || 100) : computeDefaultDesktopZoom();
  const apiVault = store.getApiVault();
  const activeApi = store.getActiveApi();
  const vaultLetters = store.getLetterVault();
  const kbStats = summarizeKnowledgeBaseStats(vaultLetters);
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

  // Temukan blok-blok surat (tanpa mengisi dummy otomatis saat lembar masih kosong)
  const rawTitleBlock = (doc.blocks || []).find(b => b.type === 'letterTitle');
  const hasTitleBlock = Boolean(rawTitleBlock && ((rawTitleBlock.title || '').trim() || (rawTitleBlock.nomor || '').trim()));
  const standaloneMetaBlock = (doc.blocks || []).find(b => b.type === 'letterMetaRecipient');
  const hasVisibleMeta = Boolean(
    (rawTitleBlock?.metaRecipient && !rawTitleBlock.metaRecipient.isDeleted) ||
    (standaloneMetaBlock && !standaloneMetaBlock.isDeleted)
  );
  const rawSigBlock = (doc.blocks || []).find(b => b.type === 'letterSignature');
  const hasSigBlock = Boolean(rawSigBlock && ((rawSigBlock.name || '').trim() || (rawSigBlock.title || '').trim() || (rawSigBlock.date || '').trim()));
  const hasMeaningfulBody = (doc.blocks || []).some(b => {
    if (b.type === 'identityTable') return Array.isArray(b.rows) && b.rows.length > 0;
    if (b.type === 'paragraph') {
      const t = Array.isArray(b.runs) && b.runs.length > 0
        ? b.runs.map(r => r.text || '').join('')
        : (b.text || '');
      return t.trim().length > 0;
    }
    return false;
  });

  const isBlankSheet = !activeDocx && !hasTitleBlock && !hasSigBlock && !hasVisibleMeta && !hasMeaningfulBody;
  const nextAutoNomor = store.getNextFormattedOutgoingNomor({
    title: rawTitleBlock?.title || 'SURAT DINAS',
    forcedTypeCode: 'TU'
  }).nomorSurat;
  const titleBlock = hasTitleBlock ? rawTitleBlock : { title: '', nomor: '' };
  const rawMetaRecipient = rawTitleBlock?.metaRecipient || standaloneMetaBlock || ((isBlankSheet || doc.letterData?.metaRecipientDeleted) ? { isDeleted: true } : {
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
  const isMetaDeleted = Boolean(
    isBlankSheet ||
    rawMetaRecipient?.isDeleted ||
    doc.letterData?.metaRecipientDeleted ||
    (!rawMetaRecipient?.metaRows?.length && !rawMetaRecipient?.recipientLines?.length && !rawMetaRecipient?.customHtml)
  );

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
  const schoolSig = state.schoolProfile || {};
  const sigBlock = hasSigBlock ? rawSigBlock : {
    city: '',
    date: '',
    title: '',
    name: '',
    nip: ''
  };

  const bodyParagraphs = (doc.blocks || []).filter(b => b.type === 'paragraph');

  if (isCopilotOpenState === null) {
    isCopilotOpenState = typeof window !== 'undefined' ? window.innerWidth >= 1120 : true;
  }
  const isCopilotOpen = Boolean(isCopilotOpenState);

  container.innerHTML = `
    <div class="word-workspace">
      
      <!-- 1. RIBBON TOOLBAR CARD-GROUPED (SESUAI UI BARU) -->
      <div class="word-ribbon">
        
        <!-- KELOMPOK: HALAMAN -->
        <div class="ribbon-group">
          <span class="ribbon-label">Halaman</span>
          <div class="ribbon-rows-wrapper">
            <div class="ribbon-row">
              <select id="ribbon-paper-size" class="ribbon-select" style="width: 62px;" title="Ukuran Kertas Baku">
                <option value="a4" ${paperSize === 'a4' ? 'selected' : ''}>A4</option>
                <option value="f4" ${paperSize === 'f4' ? 'selected' : ''}>F4</option>
                <option value="letter" ${paperSize === 'letter' ? 'selected' : ''}>Letter</option>
                <option value="legal" ${paperSize === 'legal' ? 'selected' : ''}>Legal</option>
              </select>
              <div class="ribbon-select-icon-wrap" title="Orientasi Kertas">
                ${ICON_PORTRAIT}
                <select id="ribbon-orientation" class="ribbon-select" style="width: 72px;">
                  <option value="portrait" ${orientation === 'portrait' ? 'selected' : ''}>Potret</option>
                  <option value="landscape" ${orientation === 'landscape' ? 'selected' : ''}>Lanskap</option>
                </select>
              </div>
            </div>
            <div class="ribbon-row">
              <div class="dropdown">
                <button id="btn-margin-menu" class="ribbon-btn ribbon-btn-compact" type="button" title="Atur batas margin kertas">
                  ${ICON_RULER} <span>Margin</span>
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
              <button id="btn-open-doc-settings" class="ribbon-btn ribbon-btn-compact" type="button" title="Buka Pengaturan Lengkap Dokumen (Ukuran, Margin, Font, Kop, dll)">
                ${ICON_SETTINGS} <span>Format</span>
              </button>
            </div>
          </div>
        </div>

        <!-- KELOMPOK: FONT & TEKS -->
        <div class="ribbon-group">
          <span class="ribbon-label">Font &amp; Teks</span>
          <div class="ribbon-rows-wrapper">
            <div class="ribbon-row">
              <select id="cmd-font-family" class="ribbon-select" style="width: 132px;" title="Jenis Huruf">
                <option value="Times New Roman" ${fontFamily === 'Times New Roman' ? 'selected' : ''}>Times New Roman</option>
                <option value="Arial" ${fontFamily === 'Arial' ? 'selected' : ''}>Arial</option>
                <option value="Calibri" ${fontFamily === 'Calibri' ? 'selected' : ''}>Calibri</option>
                <option value="Segoe UI" ${fontFamily === 'Segoe UI' ? 'selected' : ''}>Segoe UI</option>
                <option value="Georgia" ${fontFamily === 'Georgia' ? 'selected' : ''}>Georgia</option>
                <option value="Courier New" ${fontFamily === 'Courier New' ? 'selected' : ''}>Courier New</option>
              </select>

              <select id="cmd-font-size" class="ribbon-select" style="width: 48px;" title="Ukuran Huruf">
                <option value="8" ${String(fontSize) === '8' ? 'selected' : ''}>8</option>
                <option value="9" ${String(fontSize) === '9' ? 'selected' : ''}>9</option>
                <option value="10" ${String(fontSize) === '10' ? 'selected' : ''}>10</option>
                <option value="11" ${String(fontSize) === '11' ? 'selected' : ''}>11</option>
                <option value="12" ${String(fontSize) === '12' ? 'selected' : ''}>12</option>
                <option value="14" ${String(fontSize) === '14' ? 'selected' : ''}>14</option>
                <option value="16" ${String(fontSize) === '16' ? 'selected' : ''}>16</option>
                <option value="18" ${String(fontSize) === '18' ? 'selected' : ''}>18</option>
                <option value="20" ${String(fontSize) === '20' ? 'selected' : ''}>20</option>
                <option value="24" ${String(fontSize) === '24' ? 'selected' : ''}>24</option>
                <option value="28" ${String(fontSize) === '28' ? 'selected' : ''}>28</option>
                <option value="36" ${String(fontSize) === '36' ? 'selected' : ''}>36</option>
              </select>

              <button class="ribbon-btn-icon" id="cmd-grow-font" type="button" title="Perbesar Huruf (Ctrl+Shift+>)"><span style="font-weight: 700; font-size: 12px;">A<sup>+</sup></span></button>
              <button class="ribbon-btn-icon" id="cmd-shrink-font" type="button" title="Perkecil Huruf (Ctrl+Shift+<)"><span style="font-weight: 700; font-size: 11px;">A<sup>−</sup></span></button>
              <button class="ribbon-btn-icon" id="cmd-change-case" type="button" title="Ubah Kapitalisasi (Shift+F3)"><span style="font-weight: 600; font-size: 11.5px;">Aa</span></button>
              <button class="ribbon-btn-icon" id="cmd-clear-format" type="button" title="Hapus Semua Pemformatan">${ICON_ERASER}</button>
            </div>

            <div class="ribbon-row">
              <button class="ribbon-btn-icon" id="cmd-bold" type="button" title="Tebal (Ctrl+B)"><b>B</b></button>
              <button class="ribbon-btn-icon" id="cmd-italic" type="button" title="Miring (Ctrl+I)"><i style="font-family: Georgia, serif;">I</i></button>
              <button class="ribbon-btn-icon" id="cmd-underline" type="button" title="Garis Bawah (Ctrl+U)"><u>U</u></button>
              <button class="ribbon-btn-icon" id="cmd-strike" type="button" title="Coret Tulisan (Strikethrough)"><s>S</s></button>
              <button class="ribbon-btn-icon" id="cmd-subscript" type="button" title="Subskrip (x₂)">x<sub>2</sub></button>
              <button class="ribbon-btn-icon" id="cmd-superscript" type="button" title="Superskrip (x²)">x<sup>2</sup></button>
              <span class="ribbon-divider"></span>

              <div class="ribbon-btn-color-wrap" title="Warna Sorotan Teks (Stabilo)">
                <button class="ribbon-btn-icon" id="btn-highlight-visual" type="button">
                  <span style="background: #fef08a; padding: 0 4px; border-radius: 3px; font-weight: 700; font-size: 11px; color: #0f172a; line-height: 1.35;">ab</span>
                </button>
                <input type="color" id="picker-highlight-color" value="#fef08a" class="ribbon-color-picker-input">
              </div>

              <div class="ribbon-btn-color-wrap" title="Warna Huruf (Font Color)">
                <button class="ribbon-btn-icon" id="btn-color-visual" type="button">
                  <span style="border-bottom: 2.5px solid #dc2626; font-weight: 700; font-size: 12px; color: #0f172a; line-height: 1.05; padding: 0 2px;">A</span>
                </button>
                <input type="color" id="picker-font-color" value="#dc2626" class="ribbon-color-picker-input">
              </div>
            </div>
          </div>
        </div>

        <!-- KELOMPOK: PARAGRAF -->
        <div class="ribbon-group">
          <span class="ribbon-label">Paragraf</span>
          <div class="ribbon-rows-wrapper">
            <div class="ribbon-row">
              <button class="ribbon-btn-icon ribbon-align-btn" id="cmd-align-left" type="button" title="Rata Kiri (Ctrl+L)">${ICON_ALIGN_LEFT}</button>
              <button class="ribbon-btn-icon ribbon-align-btn" id="cmd-align-center" type="button" title="Rata Tengah (Ctrl+E)">${ICON_ALIGN_CENTER}</button>
              <button class="ribbon-btn-icon ribbon-align-btn" id="cmd-align-right" type="button" title="Rata Kanan (Ctrl+R)">${ICON_ALIGN_RIGHT}</button>
              <button class="ribbon-btn-icon ribbon-align-btn active" id="cmd-align-justify" type="button" title="Rata Kiri-Kanan / Justify (Ctrl+J)">${ICON_ALIGN_JUSTIFY}</button>
              <span class="ribbon-divider"></span>
              <button class="ribbon-btn-icon" id="cmd-toggle-pilcrow" type="button" title="Tampilkan/Sembunyikan Tanda Paragraf">${ICON_PILCROW}</button>
              <div class="ribbon-btn-color-wrap" title="Warna Latar Paragraf (Shading)">
                <button class="ribbon-btn-icon" id="btn-shading-visual" type="button" title="Warna Latar Paragraf">${ICON_SHADING}</button>
                <input type="color" id="picker-shading-color" value="#f1f5f9" class="ribbon-color-picker-input">
              </div>
            </div>

            <div class="ribbon-row">
              <button class="ribbon-btn-icon" id="cmd-bullet-list" type="button" title="Daftar Poin (Bullets)">${ICON_LIST_BULLET}</button>
              <button class="ribbon-btn-icon" id="cmd-numbered-list" type="button" title="Daftar Bernomor (Numbering)">${ICON_LIST_NUMBER}</button>
              <button class="ribbon-btn-icon" id="cmd-outdent" type="button" title="Kurangi Inden (Geser Kiri)">${ICON_OUTDENT}</button>
              <button class="ribbon-btn-icon" id="cmd-indent" type="button" title="Tambah Inden (Geser Kanan)">${ICON_INDENT}</button>
              <select id="ribbon-line-spacing" class="ribbon-select" style="width: 90px;" title="Jarak Spasi Baris">
                <option value="1.0" ${lineSpacing === '1.0' ? 'selected' : ''}>Spasi 1.0</option>
                <option value="1.15" ${lineSpacing === '1.15' ? 'selected' : ''}>Spasi 1,15</option>
                <option value="1.5" ${lineSpacing === '1.5' ? 'selected' : ''}>Spasi 1.5</option>
                <option value="2.0" ${lineSpacing === '2.0' ? 'selected' : ''}>Spasi 2.0</option>
              </select>
            </div>
          </div>
        </div>

        <!-- KELOMPOK: SISIPKAN -->
        <div class="ribbon-group" title="Sisipkan Elemen Kop Surat &amp; Tabel">
          <span class="ribbon-label">Sisipkan</span>
          <div class="ribbon-tiles-row">
            <button id="btn-open-kop-gallery" class="ribbon-tile" type="button" title="Galeri Kop Surat: Pilih template kop resmi tersimpan">
              ${ICON_GALLERY}
              <span>Kop</span>
              <span class="badge-count badge-floating">${store.getKopGallery().length}</span>
            </button>
            <button id="btn-ribbon-insert-table" class="ribbon-tile" type="button" data-canvas-insert="table" title="Sisipkan Tabel Rincian Tanpa Garis (Nama : ...)">
              ${ICON_TABLE}
              <span>Rincian</span>
            </button>
            <button id="btn-ribbon-insert-grid-table" class="ribbon-tile" type="button" data-canvas-insert="grid-table" title="Sisipkan Tabel Grid Berbingkai (No | Nama | Jabatan)">
              ${ICON_TABLE}
              <span>Tabel Grid</span>
            </button>
            <button id="btn-ribbon-add-pagebreak" class="ribbon-tile" type="button" data-canvas-insert="pagebreak" title="Sisipkan Pindah Halaman Baru (Page Break)">
              ${ICON_PLUS}
              <span>Hal Baru</span>
            </button>
            <button id="btn-open-st-sppd-pro" class="ribbon-tile smart-ribbon-highlight" type="button" title="Buka Paket Resmi 3 Halaman Surat Tugas &amp; SPPD (1:1 MS Word) di Word Pro">
              ${ICON_TARGET}
              <span>ST + SPPD</span>
            </button>
            <input type="file" id="input-quick-kop-file" accept="image/png,image/jpeg,image/svg+xml" class="hidden">
            <button id="btn-quick-upload-kop" class="ribbon-tile" type="button" title="Unggah Berkas Gambar Kop Baru (PNG/JPG/SVG)">
              ${ICON_UPLOAD}
              <span>Unggah</span>
            </button>
            <button id="btn-toggle-kop" class="ribbon-tile ${showKop ? 'active' : ''}" type="button" title="${showKop ? 'Sembunyikan Kop Surat' : 'Tampilkan Kop Surat'}">
              ${showKop ? ICON_EYE : ICON_EYE_OFF}
              <span>Tampil</span>
            </button>
          </div>
        </div>

        <!-- KELOMPOK: DOKUMEN -->
        <div class="ribbon-group" title="Simpan, Riwayat Edit, Impor Berkas &amp; AI Copilot">
          <span class="ribbon-label">Dokumen</span>
          <div class="ribbon-tiles-row">
            <button id="btn-save-to-vault" class="ribbon-tile" type="button" title="Simpan perubahan ke dokumen aktif di Gudang Surat (Ctrl+S)">
              ${ICON_SAVE}
              <span>Simpan</span>
            </button>
            <button id="btn-save-as-vault" class="ribbon-tile" type="button" title="Save As / Duplikat sebagai dokumen baru di Gudang Surat (Ctrl+Shift+S / F12)">
              ${ICON_COPY_DOC}
              <span>Save As</span>
            </button>
            <button id="btn-refresh-canvas" class="ribbon-tile" type="button" title="Refresh / Kosongkan Kertas Kerja ke Lembaran Baru">
              ${ICON_REFRESH}
              <span>Refresh</span>
            </button>
            <button id="btn-toggle-word-qr-canvas" class="ribbon-tile hidden" type="button" title="Tampilkan / Sembunyikan Barcode QR di Kanvas (tetap tercetak saat Print &amp; Surat Keluar)">
              ${ICON_CLIPBOARD_CHECK}
              <span data-qr-btn-label="1">Tampil QR</span>
            </button>
            <button id="cmd-undo" class="ribbon-tile" type="button" title="Urungkan Perubahan (Ctrl+Z)">
              ${ICON_UNDO}
              <span>Undo</span>
            </button>
            <button id="cmd-redo" class="ribbon-tile" type="button" title="Ulangi Perubahan (Ctrl+Y)">
              ${ICON_REDO}
              <span>Redo</span>
            </button>
            <input type="file" id="input-upload-docx-word" accept=".docx" class="hidden">
            <button id="btn-upload-docx-word" class="ribbon-tile" type="button" title="Buka / Impor Berkas Word (.docx) dari Komputer">
              ${ICON_DOC_UPLOAD}
              <span>Impor</span>
            </button>
            <button id="btn-toggle-copilot" class="ribbon-tile ${isCopilotOpen ? 'active' : ''}" type="button" title="Buka / Tutup Panel AI Copilot Persuratan">
              ${ICON_SPARKLE}
              <span>Copilot</span>
            </button>
          </div>
        </div>

      </div>

      <!-- 2. STAGE UTAMA HORIZONTAL: DESK KIRI + AI COPILOT SIDEBAR KANAN -->
      <div class="word-main-stage">
        <div class="word-main-left">

          <!-- DESK AREA DENGAN LEMBAR KERTAS WORD DAN ZOOM SCALING -->
          <div class="word-desk" id="word-desk-container">

            <!-- BAR PROMPT INSTRUKSI EDIT / REVISI LANGSUNG DI ATAS LEMBAR KERJA -->
            <div class="quick-ai-edit-bar no-print" id="quick-ai-edit-bar">
              <div class="quick-ai-edit-main">
                <span class="quick-ai-badge" title="Edit, koreksi kata, atau tambah data pada surat yang sedang terbuka tanpa merombak dari awal">
                  ${ICON_SPARKLE} <span>AI Edit Surat Ini</span>
                </span>
                <input type="text" id="input-quick-ai-edit" class="quick-ai-input" autocomplete="off"
                  placeholder="Tinjau dan rapikan surat ini: perbaiki ejaan sesuai PUEBI, tingkatkan bahasa resmi, atau ketik instruksi revisi...">
                <button type="button" id="btn-run-quick-ai-edit" class="btn btn-sm btn-primary quick-ai-btn">
                  ${ICON_SPARKLE} <span>Terapkan Edit</span>
                </button>
                <button type="button" id="btn-quick-record-outgoing" class="quick-ai-outgoing-btn" title="Kunci &amp; catat surat ini ke Buku Surat Keluar resmi">
                  ${ICON_SEND_OUT} <span>Catat Surat Keluar</span>
                </button>
              </div>
              <div class="quick-ai-chips-row">
                <span style="font-size: 11px; color: #475569; font-weight: 600;">Contoh Instruksi Cepat:</span>
                <button type="button" class="quick-ai-chip" data-edit="Tambahkan baris Hari/Tanggal: Sabtu, 20 November 2026, Waktu: 08.00 WIB s.d Selesai, dan Tempat: Ruang Rapat SMK IT Al-Hawari">
                  <span style="color: #0f766e; display: inline-flex;">${ICON_CALENDAR}</span> Tambah Jadwal &amp; Tempat
                </button>
                <button type="button" class="quick-ai-chip" data-edit="Perbaiki tata bahasa, ejaan typo, dan kalimat pada surat ini agar lebih sopan dan baku sesuai PUEBI tanpa mengubah data">
                  <span style="color: #15803d; display: inline-flex;">${ICON_SPARKLE}</span> Perbaiki Kata &amp; Kalimat
                </button>
                <button type="button" class="quick-ai-chip" data-edit="Ubah tujuan Kepada Yth menjadi Yth. Bapak/Ibu Orang Tua/Wali Siswa dan sesuaikan sifat menjadi Penting">
                  <span style="color: #0284c7; display: inline-flex;">${ICON_TARGET}</span> Ubah Tujuan Kepada Yth
                </button>
                <button type="button" class="quick-ai-chip" data-edit="Tambahkan paragraf penegasan agar hadir tepat waktu 15 menit sebelum acara dimulai">
                  <span style="color: #334155; display: inline-flex;">${ICON_PLUS}</span> Tambah Kalimat
                </button>
              </div>
              <div id="quick-ai-edit-status" class="quick-ai-status hidden">
                <div class="spinner" style="width: 14px; height: 14px; margin: 0; border-width: 2px;"></div>
                <span id="quick-ai-edit-status-text">AI sedang menerapkan instruksi revisi pada dokumen Anda...</span>
              </div>
            </div>
        
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
              <button type="button" id="btn-banner-record-outgoing" class="btn btn-sm ribbon-btn-compact" style="background: #0369a1; color: white; border: none; font-weight: 600;" title="Catat surat impor ini ke Surat Keluar (terkunci &amp; bernomor urut)">📤 <span>Surat Keluar</span></button>
              <button id="btn-close-imported-docx" class="btn btn-sm btn-secondary ribbon-btn-compact" title="Kembali ke template surat sekolah">✕ Tutup</button>
            </div>
          </div>
        ` : ''}

        <!-- BANNER KONFIRMASI PASCA-CETAK (CATAT KE SURAT KELUAR) -->
        <div id="post-print-outgoing-bar" class="no-print hidden" style="width: 100%; max-width: 794px; margin: 0 auto 10px auto; background: #ecfdf5; border: 1px solid #6ee7b7; border-radius: 8px; padding: 10px 14px; display: flex; justify-content: space-between; align-items: center; gap: 10px; box-shadow: 0 2px 8px rgba(16, 185, 129, 0.12); box-sizing: border-box;">
          <div style="font-size: 12.5px; color: #065f46;">
            <strong>🖨️ Dokumen siap / selesai dicetak!</strong> Jika surat ini sudah final &amp; resmi diterbitkan, catat ke <b>Surat Keluar</b> agar nomor tersusun &amp; arsip terkunci.
          </div>
          <div style="display: flex; gap: 6px; flex-shrink: 0;">
            <button type="button" id="btn-post-print-record" class="btn btn-sm btn-success" style="font-weight: 700;">🔒 Simpan ke Surat Keluar</button>
            <button type="button" id="btn-close-post-print" class="btn btn-sm btn-secondary">Tutup</button>
          </div>
        </div>

        <div class="word-sheet-wrapper" id="word-sheet-wrapper" style="transform: scale(${zoomLevel / 100});">
          
          ${activeDocx && isDocxPreviewMode ? `
            <!-- PRATINJAU DOKUMEN ASLI MENGGUNAKAN DOCX-PREVIEW -->
            <div id="docx-native-preview-card" class="docx-native-preview-card">
              <div id="docx-preview-mount" style="min-height: 297mm;"></div>
            </div>
          ` : `
            <!-- LEMBAR KERTAS MULTI-HALAMAN SKALA NYATA WORD -->
            <div class="word-page-wrapper" data-word-page="0">
              <div class="word-page-header-bar no-print hidden">
                <span class="word-page-pill">Halaman 1 dari 1</span>
              </div>
              <div class="word-paper-sheet paper-${paperSize} ${orientation === 'landscape' ? 'landscape' : ''}" id="word-sheet" data-word-page="0"
                style="
                  --margin-top: ${customMargins.top}cm;
                  --margin-bottom: ${customMargins.bottom}cm;
                  --margin-left: ${customMargins.left}cm;
                  --margin-right: ${customMargins.right}cm;
                  --printable-h-mm: ${computePrintableDimsMm({ paperSize, orientation, margins: customMargins }).printableHeightMm}mm;
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

                  <!-- BLOK ISI SURAT DINAMIS (BEBAS TAMBAH & HAPUS TANPA CELAH KOSONG) -->
                  <div id="sheet-dynamic-blocks" class="sheet-dynamic-blocks">
                    ${renderDynamicImportedBlocksHtml(docxBodyList, docxIdentityRows, docxClosing)}
                  </div>
                  <div class="canvas-add-block-bar no-print">
                    <button type="button" class="btn-add-canvas-block" data-canvas-insert="para">+ Tambah Area Teks</button>
                    <button type="button" class="btn-add-canvas-block" data-canvas-insert="table">+ Tabel Rincian</button>
                    <button type="button" class="btn-add-canvas-block" data-canvas-insert="grid-table">+ Tabel Grid (Kolom)</button>
                    <button type="button" class="btn-add-canvas-block" data-canvas-insert="pagebreak">+ Pindah Halaman</button>
                  </div>

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
                        ${escapeHtml(formatSignerNipLine(docxSig.nip))}
                      </div>
                    </div>
                  </div>
                ` : `
                  <!-- STANDAR TEMPLATE DOKUMEN (MULAI BERSIH/KOSONG SAAT DOKUMEN BARU) -->
                  <div class="sheet-section-title ${hasTitleBlock ? '' : 'hidden'}">
                    <h2 contenteditable="true" id="editable-doc-title" class="editable-field text-center font-title" spellcheck="false" data-placeholder="JUDUL SURAT (MISAL: SURAT UNDANGAN)">${hasTitleBlock ? escapeHtml(titleBlock.title || '') : ''}</h2>
                    <div contenteditable="true" id="editable-doc-nomor" class="editable-field text-center font-nomor" spellcheck="false" data-placeholder="Nomor: ...">${hasTitleBlock ? escapeHtml(titleBlock.nomor || '') : ''}</div>
                  </div>

                  <!-- BLOK SIFAT, PERIHAL & KEPADA YTH (DAPAT DIEDIT / DIHAPUS MANUAL DI CANVAS) -->
                  <div class="sheet-meta-recipient-wrapper ${isMetaDeleted ? 'is-empty' : ''} ${isBlankSheet ? 'hidden' : ''}" id="sheet-meta-recipient-wrapper">
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

                  <!-- BLOK ISI SURAT DINAMIS (BEBAS TAMBAH & HAPUS TANPA CELAH KOSONG) -->
                  <div id="sheet-dynamic-blocks" class="sheet-dynamic-blocks">
                    ${renderDynamicStandardBlocksHtml(doc)}
                  </div>
                  <div class="canvas-add-block-bar no-print">
                    <button type="button" class="btn-add-canvas-block" data-canvas-insert="para">+ Tambah Area Teks</button>
                    <button type="button" class="btn-add-canvas-block" data-canvas-insert="table">+ Tabel Rincian</button>
                    <button type="button" class="btn-add-canvas-block" data-canvas-insert="grid-table">+ Tabel Grid (Kolom)</button>
                    <button type="button" class="btn-add-canvas-block" data-canvas-insert="pagebreak">+ Pindah Halaman</button>
                    <button type="button" class="btn-add-canvas-block ${hasTitleBlock ? 'hidden' : ''}" data-canvas-insert="title">+ Judul &amp; Nomor</button>
                    <button type="button" class="btn-add-canvas-block ${!isBlankSheet && !isMetaDeleted ? 'hidden' : ''}" data-canvas-insert="meta">+ Sifat &amp; Kepada Yth</button>
                    <button type="button" class="btn-add-canvas-block ${hasSigBlock ? 'hidden' : ''}" data-canvas-insert="sig">+ Blok TTD</button>
                  </div>

                  <!-- BLOK TANDA TANGAN DI KANAN BAWAH -->
                  <div class="sheet-signature-wrapper ${hasSigBlock ? '' : 'hidden'}">
                    <div class="sig-box">
                      <div contenteditable="true" id="editable-sig-date" class="editable-field text-left" spellcheck="false">${hasSigBlock ? `${escapeHtml(sigBlock.city || schoolSig.city || 'Garut')}, ${escapeHtml(sigBlock.date || '14 November 2026')}` : ''}</div>
                      <div contenteditable="true" id="editable-sig-title" class="editable-field text-left" spellcheck="false">${hasSigBlock ? `${escapeHtml(sigBlock.title || 'Kepala Sekolah')},` : ''}</div>
                      <div class="sig-space-holder"></div>
                      <div contenteditable="true" id="editable-sig-name" class="editable-field text-left font-bold-underline" spellcheck="false">${hasSigBlock ? escapeHtml(sigBlock.name || schoolSig.principalName || 'Dian Nurhadian, S.Pd.') : ''}</div>
                      <div contenteditable="true" id="editable-sig-nip" class="editable-field text-left font-nip" spellcheck="false">${hasSigBlock ? escapeHtml(formatSignerNipLine(sigBlock.nip || schoolSig.principalNip)) : ''}</div>
                    </div>
                  </div>
                `}

              </div>
              <div class="floating-qr-stage" data-qr-page="0"></div>
            </div>
            </div>
          `}
          </div>
        </div>

          <!-- 3. STATUS BAR MICROSOFT WORD (DI BAWAH DESK KIRI) -->
          <footer class="word-status-bar">
            <div class="status-left">
              <span id="word-status-pages" class="status-item">${ICON_DOC} <span>Halaman 1 dari 1</span></span>
              <span class="status-sep">|</span>
              <span id="status-word-count" class="status-item status-word-count">128 kata</span>
              <span class="status-sep">|</span>
              <span class="status-item">${ICON_GLOBE} <span>Bahasa Indonesia (PUEBI)</span></span>
              <span class="status-sep">|</span>
              <span id="word-status-vault-badge" class="status-item status-vault-pill">📄 Draf Baru</span>
            </div>

            <div class="status-right">
              <div class="zoom-slider-wrap">
                <button type="button" class="zoom-btn" id="btn-zoom-out" title="Perkecil Kanvas">&minus;</button>
                <input type="range" id="input-zoom-slider" min="50" max="175" step="5" value="${zoomLevel}" title="Geser skala zoom kanvas">
                <button type="button" class="zoom-btn" id="btn-zoom-in" title="Perbesar Kanvas">+</button>
                <span id="label-zoom-val" class="status-zoom-pill">${zoomLevel}%</span>
              </div>
              <button type="button" id="btn-zoom-fit" class="status-fit-btn" title="Sesuaikan Lebar Layar (Fit)">
                ${ICON_MAXIMIZE}
              </button>
            </div>
          </footer>
        </div> <!-- /.word-main-left -->

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

      <!-- 5. PANEL AI COPILOT (DOCKED KANAN DI DESKTOP, SLIDE-OVER DI LAYAR KECIL) -->
      <aside class="copilot-drawer ${isCopilotOpen ? '' : 'hidden'}" id="copilot-drawer">
        <div class="copilot-header">
          <div class="copilot-title">
            <span style="display: inline-flex;">${ICON_BOT_SPARK}</span>
            <h4>AI Copilot Surat Sekolah</h4>
          </div>
          <button id="btn-close-copilot" class="copilot-close-btn" type="button" title="Tutup Panel AI Copilot">&times;</button>
        </div>

        <div class="copilot-body">
          
          <!-- PEMILIH MODEL DARI GUDANG API -->
          <div class="copilot-model-card">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
              <label style="font-size: 11.5px; font-weight: 700; color: #1e293b;">Model AI dari Gudang API</label>
              <a href="javascript:void(0)" id="link-open-vault-from-copilot" style="display: inline-flex; align-items: center; gap: 4px; font-size: 11px; color: var(--primary); font-weight: 600; text-decoration: none;">
                ${ICON_SETTINGS} <span>Gudang API</span>
              </a>
            </div>
            <select id="select-copilot-model" class="form-control" style="font-size: 12px; font-weight: 600; background: #ffffff;">
              ${apiVault.map(v => `
                <option value="${v.id}" ${v.id === activeApi?.id ? 'selected' : ''}>
                  [${v.provider.toUpperCase()}] ${escapeHtml(v.label)}
                </option>
              `).join('')}
            </select>
          </div>

          <!-- BASIS PENGETAHUAN / KNOWLEDGE BASE DARI GUDANG SURAT -->
          <div class="copilot-kb-card">
            <div style="display: flex; align-items: flex-start; gap: 7px; min-width: 0;">
              <span style="color: #15803d; display: inline-flex; margin-top: 1px;">${ICON_OPEN_BOOK}</span>
              <span><strong>Basis Pengetahuan:</strong> ${kbStats.totalActive} arsip aktif dipelajari AI (${kbStats.masterCount} ⭐ Master Terkunci)</span>
            </div>
            <a href="javascript:void(0)" id="link-open-vault-from-copilot-kb" class="copilot-kb-btn">
              Kelola<br>Gudang
            </a>
          </div>

          <p class="copilot-desc">
            Ketik jenis surat apa pun yang ingin Anda buat dalam kalimat bebas. AI akan otomatis merangkai isi dokumen dan menampilkannya langsung di lembar kerja Word Anda.
          </p>

          <!-- CHIPS TEMPLATE CEPAT (GRID 2 KOLOM) -->
          <div class="copilot-chips">
            <button type="button" class="copilot-chip" data-prompt="Buatkan Surat Undangan Rapat Dinas Pleno Dewan Guru membahas evaluasi pembelajaran (Kode UND)">
              <span class="copilot-chip-icon" style="color: #16a34a;">${ICON_MAIL}</span><span>Undangan Rapat (UND)</span>
            </button>
            <button type="button" class="copilot-chip" data-prompt="Buatkan Surat dari Tata Usaha (TU) tentang pemberitahuan pengumpulan berkas administrasi kelas (Kode TU)">
              <span class="copilot-chip-icon" style="color: #d97706;">${ICON_FOLDER}</span><span>Surat Tata Usaha (TU)</span>
            </button>
            <button type="button" class="copilot-chip" data-prompt="Buatkan Surat Keputusan Kepala Sekolah tentang penetapan panitia kegiatan sekolah (Kode KEP)">
              <span class="copilot-chip-icon" style="color: #0284c7;">${ICON_BUILDING}</span><span>SK Kepala Sekolah (KEP)</span>
            </button>
            <button type="button" class="copilot-chip" data-prompt="Buatkan Surat Keterangan Siswa Aktif untuk keperluan beasiswa kuliah (Kode KET)">
              <span class="copilot-chip-icon" style="color: #0f766e;">${ICON_ID_CARD}</span><span>Ket. Siswa Aktif (KET)</span>
            </button>
            <button type="button" class="copilot-chip" data-prompt="Buatkan Surat Tugas Guru SMK IT Al-Hawari beserta lampiran SPPD 8 poin dan halaman Keterangan perjalanan dinas (Kode ST)">
              <span class="copilot-chip-icon" style="color: #b45309;">${ICON_BRIEFCASE}</span><span>Surat Tugas + SPPD</span>
            </button>
            <button type="button" class="copilot-chip" data-prompt="Buatkan Surat Panggilan Orang Tua/Wali Siswa terkait pembinaan kedisiplinan siswa (Kode PGL)">
              <span class="copilot-chip-icon" style="color: #334155;">${ICON_USERS}</span><span>Panggilan Ortu (PGL)</span>
            </button>
          </div>

          <div class="form-group" style="margin-top: 4px;">
            <label class="copilot-prompt-label">${ICON_PENCIL_NOTE} <span>Perintah Pembuatan Baru atau Instruksi Edit Surat Terbuka</span></label>
            <textarea id="copilot-prompt-input" class="form-control" rows="4" spellcheck="false"
              placeholder="Ketik instruksi edit untuk surat yang sedang terbuka (misal: Tambahkan baris Waktu: 09.00 WIB, ganti nama siswa, perbaiki kalimat...) ATAU perintah buat surat baru..."></textarea>
          </div>

          <div style="display: flex; flex-direction: column; gap: 7px;">
            <button type="button" id="btn-edit-open-copilot" class="btn copilot-primary-cta" title="Terapkan penambahan data, koreksi, atau perbaikan kata pada surat yang sedang terbuka di lembar kerja (bukan buat dari awal)">
              ${ICON_PENCIL_NOTE} <span>Edit / Tambah pada Surat Terbuka</span>
            </button>
            <button type="button" id="btn-run-copilot" class="btn copilot-secondary-cta" title="Buat surat baru dari awal sesuai perintah">
              ${ICON_DOC} <span>Buat Surat Baru dari Awal</span>
            </button>
          </div>

          <div id="copilot-status" class="copilot-status-area hidden">
            <div style="display: flex; align-items: center; justify-content: center; gap: 8px;">
              <div class="spinner" style="width: 16px; height: 16px; margin: 0; border-width: 2px;"></div>
              <span id="copilot-status-text" style="font-weight: 600;">AI sedang merangkai surat...</span>
              <button id="btn-cancel-copilot" type="button" class="btn btn-sm btn-secondary" style="padding: 1px 8px; font-size: 11px; margin-left: 4px; height: 22px; border-color: #cbd5e1;">Batal</button>
            </div>
            <div id="copilot-status-tip" style="margin-top: 6px; font-size: 11px; color: #475569; display: none;">
              Tips: Jika Gemini terasa lambat di jam sibuk, Anda dapat memilih <b>Groq (~0.2 detik)</b> di dropdown atas.
            </div>
          </div>

          <hr style="margin: 14px 0; border-color: #e2e8f0;"/>
          <h5 class="copilot-quick-title">${ICON_SPARKLE} <span>Aksi Cepat AI Copilot:</span></h5>
          <div class="btn-group-vertical" style="display: flex; flex-direction: column; gap: 6px;">
            <button type="button" id="btn-copilot-polish-messy" class="copilot-quick-btn copilot-quick-btn-polish">
              <span style="color: #15803d; display: inline-flex;">${ICON_SPARKLE}</span> <span>Rapihkan Dokumen yang Berantakan (AI)</span>
            </button>
            <button type="button" id="btn-copilot-formalize" class="copilot-quick-btn">
              <span style="color: #475569; display: inline-flex;">${ICON_OPEN_BOOK}</span> <span>Rapikan Bahasa ke Standar Dinas Baku (PUEBI)</span>
            </button>
            <button type="button" id="btn-copilot-today-date" class="copilot-quick-btn">
              <span style="color: #475569; display: inline-flex;">${ICON_CALENDAR}</span> <span>Perbarui Tanggal Surat ke Hari Ini</span>
            </button>
          </div>
        </div>
      </aside>
      </div> <!-- /.word-main-stage -->

      <!-- 6. MODAL KONFIRMASI CATAT SURAT KELUAR (ARSIP TETAP TERKUNCI & NOMOR BERURUTAN) -->
      <div id="modal-record-outgoing" class="modal-overlay hidden">
        <div class="modal-card" style="max-width: 540px; width: 92%;">
          <div class="modal-header" style="display: flex; justify-content: space-between; align-items: center; padding-bottom: 10px; border-bottom: 1px solid #e2e8f0;">
            <h3 style="margin: 0; font-size: 16px; color: #0f172a;">📤 Catat sebagai Surat Keluar Resmi</h3>
            <button type="button" id="btn-close-outgoing-modal" class="btn btn-sm btn-secondary">&times;</button>
          </div>
          <div class="modal-body" style="padding: 14px 0;">
            <div style="background: #fffbeb; border: 1px solid #fde68a; color: #92400e; padding: 10px 12px; border-radius: 6px; font-size: 12px; margin-bottom: 14px;">
              <strong>🔒 Perhatian:</strong> Surat yang dicatat ke <b>Surat Keluar</b> akan <b>terkunci permanen (tidak bisa diedit lagi)</b> dan dicatat dengan nomor urut tersusun di Buku Catatan Surat Keluar.
            </div>
            <div class="form-group" style="margin-bottom: 10px;">
              <label style="font-size: 12px; font-weight: 700;">Nomor Urut &amp; Nomor Surat Tersusun Otomatis:</label>
              <div style="display: flex; gap: 8px; align-items: center;">
                <span id="out-seq-badge" style="background: #dbeafe; color: #1e40af; font-weight: 800; padding: 6px 10px; border-radius: 6px; font-size: 13px;">#001</span>
                <input type="text" id="input-out-nomor" class="form-control" style="font-family: Consolas, monospace; font-weight: 700;">
              </div>
              <div style="margin-top: 6px; display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
                <label style="font-size: 11px; color: #64748b; margin: 0;">Kode Jenis Surat:</label>
                <select id="select-out-type-code" class="form-control" style="width: auto; font-size: 11.5px; padding: 3px 8px; height: auto;">
                  ${Object.values(LETTER_TYPE_CODES).map(c => `<option value="${c.code}">${c.code} — ${c.label}</option>`).join('')}
                </select>
              </div>
            </div>
            <div class="form-group" style="margin-bottom: 10px;">
              <label style="font-size: 12px; font-weight: 700;">Judul Surat:</label>
              <input type="text" id="input-out-title" class="form-control" placeholder="Misal: SURAT KETERANGAN SISWA AKTIF">
            </div>
            <div class="form-group" style="margin-bottom: 10px;">
              <label style="font-size: 12px; font-weight: 700;">Atas Nama / Tujuan Penerima:</label>
              <input type="text" id="input-out-recipient" class="form-control" placeholder="Nama siswa, guru, atau instansi penerima surat...">
            </div>
            <div class="form-group" style="margin-bottom: 10px;">
              <label style="font-size: 12px; font-weight: 700;">Perihal / Keterangan Singkat:</label>
              <input type="text" id="input-out-perihal" class="form-control" placeholder="Perihal surat...">
            </div>
            ${buildSignerRoleSelectorHtml({
              radioName: 'out-signer-role',
              checkboxId: 'chk-out-authorize-sah',
              defaultRole: doc?.verification?.signerRole || 'kepala_sekolah',
              school: state.schoolProfile
            })}
          </div>
          <div style="display: flex; justify-content: flex-end; gap: 8px; border-top: 1px solid #e2e8f0; padding-top: 12px;">
            <button type="button" id="btn-cancel-record-outgoing" class="btn btn-secondary">Batal</button>
            <button type="button" id="btn-confirm-record-outgoing" class="btn btn-primary" style="font-weight: 700;">🛡️ Kunci, Sahkan &amp; Catat</button>
          </div>
        </div>
      </div>

      <!-- 7. MODAL KONFIRMASI REFRESH / KOSONGKAN KERTAS -->
      <div id="modal-refresh-canvas" class="modal-overlay hidden">
        <div class="modal-box" style="max-width: 490px; width: 92%;">
          <div class="modal-header">
            <h3>🔄 Refresh &amp; Kosongkan Lembar Kerja</h3>
            <button type="button" id="btn-close-refresh-modal" class="btn-icon" aria-label="Tutup">&times;</button>
          </div>
          <div class="modal-body">
            <div style="background: #fffbeb; border: 1px solid #fde68a; color: #92400e; padding: 10px 12px; border-radius: 8px; font-size: 12.5px; line-height: 1.5;">
              <strong>⚠️ Dokumen masih memiliki isi:</strong> Sebelum mengosongkan kanvas menjadi lembaran baru yang bersih, pilih apakah isi surat saat ini ingin disimpan terlebih dahulu:
            </div>
            <div class="refresh-choice-list">
              <button type="button" id="btn-refresh-save-vault" class="refresh-choice-btn refresh-choice-vault">
                <div>
                  📁 Simpan ke Gudang Surat &amp; Kosongkan
                  <small>Arsip dapat dibuka/diedit kembali kapan saja dan menjadi referensi gaya AI.</small>
                </div>
                <span>&rarr;</span>
              </button>
              <button type="button" id="btn-refresh-save-outgoing" class="refresh-choice-btn refresh-choice-outgoing">
                <div>
                  📤 Simpan ke Surat Keluar &amp; Kosongkan
                  <small>Tercatat dengan nomor urut resmi &amp; terkunci di Buku Surat Keluar, lalu kanvas bersih.</small>
                </div>
                <span>&rarr;</span>
              </button>
              <button type="button" id="btn-refresh-discard" class="refresh-choice-btn refresh-choice-discard">
                <div>
                  🗑️ Kosongkan Langsung (Tanpa Simpan)
                  <small>Hapus seluruh isi di kertas kerja saat ini dan mulai lembar putih kosong baru.</small>
                </div>
                <span>&times;</span>
              </button>
            </div>
          </div>
          <div class="modal-footer">
            <button type="button" id="btn-cancel-refresh-modal" class="btn btn-secondary">Batal</button>
          </div>
        </div>
      </div>

      <div id="word-save-modal-root"></div>
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

  // 7. Bind Catat Surat Keluar Modal
  bindRecordOutgoingModal(container);

  // 8. Bind Refresh & Konfirmasi Simpan Modal
  bindRefreshCanvasModal(container);

  // Update status kata & paginasi multi-halaman awal
  updateWordCount(container);
  const kopImgForPaginate = container.querySelector('#sheet-kop-image');
  if (kopImgForPaginate && !kopImgForPaginate.complete) {
    kopImgForPaginate.addEventListener('load', () => scheduleWordPagination(container), { once: true });
  }
  requestAnimationFrame(() => {
    reflowWordCanvasPagination(container);
    refreshWordFloatingQrOverlay(container);
  });

  // Toggle Tampil / Sembunyi Barcode QR di Kanvas Lembar Kerja Word
  container.querySelector('#btn-toggle-word-qr-canvas')?.addEventListener('click', () => {
    const curDoc = store.getState().document || {};
    const stages = [...container.querySelectorAll('.word-paper-sheet .floating-qr-stage')];
    const current = curDoc.qrPlacement
      ? normalizeQrPlacement(curDoc.qrPlacement, Math.max(1, stages.length))
      : findPreferredSignatureQrPlacement(curDoc, Math.max(1, stages.length));
    const next = { ...current, visibleOnCanvas: !(current.visibleOnCanvas !== false) };
    store.setActiveQrPlacement(next, 'standard');
    refreshWordFloatingQrOverlay(container);
    showToast(
      next.visibleOnCanvas
        ? '▣ Barcode QR ditampilkan di kanvas — bebas digeser, diubah ukuran, atau diletakkan di Depan/Belakang Teks.'
        : '👁️ Barcode QR disembunyikan di kanvas (tetap muncul saat Cetak & Surat Keluar).',
      'info'
    );
  });

  const onWordVerifSync = () => {
    if (!document.body.contains(container)) {
      window.removeEventListener('docbuilder:verification-updated', onWordVerifSync);
      return;
    }
    refreshWordFloatingQrOverlay(container);
  };
  window.addEventListener('docbuilder:verification-updated', onWordVerifSync);

  // Pastikan posisi ribbon dan window selalu terikat di paling kiri (Halaman & Margin selalu tampak penuh)
  const ribbonEl = container.querySelector('.word-ribbon');
  if (ribbonEl) {
    ribbonEl.scrollLeft = 0;
  }
  if (typeof window !== 'undefined') {
    window.scrollTo(0, 0);
  }

  // Dengarkan sinkronisasi async IndexedDB (mis. saat startup memulihkan outgoingLetters)
  if (!container.__outgoingSyncSubscribed) {
    container.__outgoingSyncSubscribed = true;
    let lastOutgoingLen = store.getOutgoingLetters().length;
    store.subscribe((newState) => {
      if (!document.body.contains(container)) return;
      const newLen = (newState.outgoingLetters || []).length;
      if (newLen !== lastOutgoingLen) {
        lastOutgoingLen = newLen;
        const badgeEl = container.querySelector('#badge-outgoing-count');
        if (badgeEl) badgeEl.textContent = String(newLen);
        const nomorEl = container.querySelector('#editable-doc-nomor');
        const tb = newState.document?.blocks?.find(b => b.type === 'letterTitle');
        if (nomorEl && tb?.nomor && document.activeElement !== nomorEl) {
          nomorEl.innerText = tb.nomor;
        }
      }
    });
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

  // Paragraph formatting row 1 & History
  bindFormatBtn('#cmd-bullet-list', 'insertUnorderedList');
  bindFormatBtn('#cmd-numbered-list', 'insertOrderedList');
  bindFormatBtn('#cmd-outdent', 'outdent');
  bindFormatBtn('#cmd-indent', 'indent');
  bindFormatBtn('#cmd-undo', 'undo');
  bindFormatBtn('#cmd-redo', 'redo');

  const btnPilcrow = container.querySelector('#cmd-toggle-pilcrow');
  if (btnPilcrow) {
    btnPilcrow.addEventListener('mousedown', preventSelectionLoss);
    btnPilcrow.addEventListener('click', () => {
      const sheets = [...container.querySelectorAll('.word-paper-sheet')];
      if (sheets.length > 0) {
        const shown = !sheets[0].classList.contains('show-pilcrow-marks');
        sheets.forEach(s => s.classList.toggle('show-pilcrow-marks', shown));
        btnPilcrow.classList.toggle('active', shown);
      }
    });
  }

  // Paragraph alignment (with active emerald state sync)
  const alignButtons = [
    { id: '#cmd-align-left', cmd: 'justifyLeft' },
    { id: '#cmd-align-center', cmd: 'justifyCenter' },
    { id: '#cmd-align-right', cmd: 'justifyRight' },
    { id: '#cmd-align-justify', cmd: 'justifyFull' }
  ];
  alignButtons.forEach(({ id, cmd }) => {
    const btn = container.querySelector(id);
    if (!btn) return;
    btn.addEventListener('mousedown', preventSelectionLoss);
    btn.addEventListener('click', () => {
      document.execCommand(cmd, false, null);
      alignButtons.forEach(item => container.querySelector(item.id)?.classList.remove('active'));
      btn.classList.add('active');
      scheduleSync(container);
    });
  });

  // Font Family selector
  const fontFamilySelect = container.querySelector('#cmd-font-family');
  if (fontFamilySelect) {
    fontFamilySelect.addEventListener('change', () => {
      const font = fontFamilySelect.value;
      document.execCommand('fontName', false, font);
      scheduleSync(container);
    });
  }

  // Font Size selector
  const fontSizeSelect = container.querySelector('#cmd-font-size');
  if (fontSizeSelect) {
    fontSizeSelect.addEventListener('change', () => {
      const size = fontSizeSelect.value;
      // Map pt to execCommand 1-7 or inline style span
      applyInlineStyleToSelection('font-size', `${size}pt`);
      scheduleSync(container);
    });
  }

  // Grow Font (A+), Shrink Font (A-), Change Case (Aa), Clear Format
  const btnGrowFont = container.querySelector('#cmd-grow-font');
  if (btnGrowFont) {
    btnGrowFont.addEventListener('mousedown', preventSelectionLoss);
    btnGrowFont.addEventListener('click', () => {
      adjustFontSize(2);
      scheduleSync(container);
    });
  }

  const btnShrinkFont = container.querySelector('#cmd-shrink-font');
  if (btnShrinkFont) {
    btnShrinkFont.addEventListener('mousedown', preventSelectionLoss);
    btnShrinkFont.addEventListener('click', () => {
      adjustFontSize(-2);
      scheduleSync(container);
    });
  }

  const btnChangeCase = container.querySelector('#cmd-change-case');
  if (btnChangeCase) {
    btnChangeCase.addEventListener('mousedown', preventSelectionLoss);
    btnChangeCase.addEventListener('click', () => {
      toggleCaseOnSelection();
      scheduleSync(container);
    });
  }

  const btnClearFormat = container.querySelector('#cmd-clear-format');
  if (btnClearFormat) {
    btnClearFormat.addEventListener('mousedown', preventSelectionLoss);
    btnClearFormat.addEventListener('click', () => {
      document.execCommand('removeFormat', false, null);
      scheduleSync(container);
    });
  }

  // Text Highlight Color Picker (Stabilo)
  const highlightInput = container.querySelector('#picker-highlight-color');
  if (highlightInput) {
    highlightInput.addEventListener('input', () => {
      document.execCommand('hiliteColor', false, highlightInput.value);
      scheduleSync(container);
    });
  }

  // Font Color Picker
  const fontColorInput = container.querySelector('#picker-font-color');
  if (fontColorInput) {
    fontColorInput.addEventListener('input', () => {
      document.execCommand('foreColor', false, fontColorInput.value);
      scheduleSync(container);
    });
  }

  // Shading / Background Color Fill
  const shadingInput = container.querySelector('#picker-shading-color');
  if (shadingInput) {
    shadingInput.addEventListener('input', () => {
      applyInlineStyleToSelection('background-color', shadingInput.value);
      scheduleSync(container);
    });
  }

  // Line Spacing
  const spacingSelect = container.querySelector('#ribbon-line-spacing');
  if (spacingSelect) {
    spacingSelect.addEventListener('change', () => {
      const val = spacingSelect.value;
      const cssLine = val === '1.0' ? '1.15' : (val === '1.5' ? '1.6' : (val === '2.0' ? '2.0' : '1.35'));
      container.querySelectorAll('.sheet-content-editable').forEach(bodyEl => {
        bodyEl.style.lineHeight = cssLine;
      });
      store.updateLetterOptions({ lineSpacing: val });
      showToast(`Jarak spasi baris diatur ke ${val}`, 'info');
      scheduleSync(container);
      scheduleWordPagination(container);
    });
  }

  // Paper Size & Orientation
  const paperSelect = container.querySelector('#ribbon-paper-size');
  if (paperSelect) {
    paperSelect.addEventListener('change', () => {
      const size = paperSelect.value;
      store.updateLetterOptions({ paperSize: size });
      syncGlobalPrintPageStyle(store.getLetterOptions());
      showToast(`Ukuran kertas diubah ke ${size.toUpperCase()}`, 'info');
      renderWordCanvas(container);
    });
  }

  const orientationSelect = container.querySelector('#ribbon-orientation');
  if (orientationSelect) {
    orientationSelect.addEventListener('change', () => {
      const orientation = orientationSelect.value;
      store.updateLetterOptions({ orientation });
      syncGlobalPrintPageStyle(store.getLetterOptions());
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
  if (btnToggleKop && kopContainer) {
    btnToggleKop.addEventListener('click', () => {
      const isHidden = kopContainer.classList.toggle('hidden');
      btnToggleKop.innerHTML = `${isHidden ? ICON_EYE_OFF : ICON_EYE}<span>Tampil</span>`;
      btnToggleKop.classList.toggle('active', !isHidden);
      btnToggleKop.title = isHidden ? 'Tampilkan Kop Surat' : 'Sembunyikan Kop Surat';
      store.updateLetterOptions({ showKop: !isHidden });
      showToast(isHidden ? 'Kop surat disembunyikan.' : 'Kop surat ditampilkan.', 'info');
      scheduleSync(container);
      scheduleWordPagination(container);
    });
  }

  // Shortcut Paket Surat Tugas + SPPD (3 Halaman Resmi Word Pro)
  const btnOpenStSppdPro = container.querySelector('#btn-open-st-sppd-pro');
  if (btnOpenStSppdPro) {
    btnOpenStSppdPro.addEventListener('click', () => {
      const stNow = store.getState();
      const optsNow = store.getLetterOptions();
      const stSppdDoc = buildSuratTugasSppdDocModel({
        nomorSurat: store.getNextFormattedOutgoingNomor({ title: 'SURAT TUGAS', forcedTypeCode: 'ST' }).nomorSurat
      }, stNow.schoolProfile, {
        ...optsNow,
        paperSize: 'f4'
      });
      store.setState({
        smartDocument: {
          pages: [],
          pendingDocModel: stSppdDoc,
          meta: stSppdDoc.meta || { title: 'SURAT TUGAS & SPPD' }
        }
      }, false);
      store.setActiveVaultId(null, 'smart');
      showToast('Membuka Template Resmi Surat Tugas & SPPD (3 Halaman 1:1 MS Word) di Word Pro...', 'success');
      document.querySelector('.nav-tab[data-tab="smart-canvas"]')?.click();
    });
  }

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

  // Gudang & Arsip Surat: Simpan (Save: Ctrl+S) & Save As (Duplikat: Ctrl+Shift+S / F12)
  updateWordVaultBadge(container);
  const btnSaveToVault = container.querySelector('#btn-save-to-vault');
  if (btnSaveToVault) {
    btnSaveToVault.addEventListener('click', () => performWordSave(container, { clearAfterSave: false }));
  }
  const btnSaveAsVault = container.querySelector('#btn-save-as-vault');
  if (btnSaveAsVault) {
    btnSaveAsVault.addEventListener('click', () => openWordSaveVaultModal(container, { mode: 'save-as', clearAfterSave: false }));
  }

  const wordWorkspaceRoot = container.querySelector('.word-workspace');
  const onWordGlobalKeyDown = (e) => {
    if (!wordWorkspaceRoot || !wordWorkspaceRoot.isConnected) {
      document.removeEventListener('keydown', onWordGlobalKeyDown);
      return;
    }
    const isSaveAs = ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === 's') || e.key === 'F12';
    const isSave = (e.ctrlKey || e.metaKey) && !e.shiftKey && e.key.toLowerCase() === 's';
    if (isSaveAs) {
      e.preventDefault();
      openWordSaveVaultModal(container, { mode: 'save-as', clearAfterSave: false });
    } else if (isSave) {
      e.preventDefault();
      performWordSave(container, { clearAfterSave: false });
    }
  };
  document.addEventListener('keydown', onWordGlobalKeyDown);

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
        store.setActiveVaultId(null, 'standard');

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

  // Unduh .docx (Mendukung Dokumen Bawaan & Berkas Impor dengan Pilihan QR Pengesahan vs TTD Manual)
  const btnDownloadWord = container.querySelector('#btn-download-word');
  if (btnDownloadWord) {
    btnDownloadWord.addEventListener('click', () => triggerWordCanvasDownloadWithChoice(container));
  }

  // Print PDF (Dengan Pilihan QR Pengesahan vs Kosongan TTD Manual Pulpen)
  const btnPrintWord = container.querySelector('#btn-print-word');
  if (btnPrintWord) {
    btnPrintWord.addEventListener('click', () => triggerWordCanvasPrintWithChoice(container));
  }

  // Toggle AI Copilot Drawer (sinkron status terbuka/tutup)
  const btnToggleCopilot = container.querySelector('#btn-toggle-copilot');
  const copilotDrawer = container.querySelector('#copilot-drawer');
  if (btnToggleCopilot && copilotDrawer) {
    btnToggleCopilot.addEventListener('click', () => {
      const willHide = !copilotDrawer.classList.contains('hidden');
      copilotDrawer.classList.toggle('hidden', willHide);
      isCopilotOpenState = !willHide;
      btnToggleCopilot.classList.toggle('active', isCopilotOpenState);
    });
  }
  const btnCloseCopilot = container.querySelector('#btn-close-copilot');
  if (btnCloseCopilot && copilotDrawer) {
    btnCloseCopilot.addEventListener('click', () => {
      copilotDrawer.classList.add('hidden');
      isCopilotOpenState = false;
      if (btnToggleCopilot) btnToggleCopilot.classList.remove('active');
    });
  }
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

  const updateZoom = (val, { fromUser = true } = {}) => {
    const clamped = Math.max(50, Math.min(180, Math.round(val)));
    sheetWrapper.style.transform = `scale(${clamped / 100})`;
    zoomSlider.value = clamped;
    zoomLabel.textContent = `${clamped}%`;
    store.setZoomLevel(clamped, { fromUser });
  };

  zoomSlider.addEventListener('input', () => {
    updateZoom(parseFloat(zoomSlider.value), { fromUser: true });
  });

  btnZoomIn.addEventListener('click', () => {
    updateZoom(parseFloat(zoomSlider.value) + 10, { fromUser: true });
  });

  btnZoomOut.addEventListener('click', () => {
    updateZoom(parseFloat(zoomSlider.value) - 10, { fromUser: true });
  });

  btnZoomFit.addEventListener('click', () => {
    // Hitung perbandingan lebar viewport desk vs lembar kertas
    const desk = container.querySelector('#word-desk-container');
    const availableWidth = desk.clientWidth - 60;
    const sheetWidth = 210 * 3.78; // A4 approx px (794px)
    const fitZoom = Math.round((availableWidth / sheetWidth) * 100);
    updateZoom(Math.max(60, Math.min(130, fitZoom)), { fromUser: true });
    showToast(`Zoom disesuaikan dengan lebar layar (${fitZoom}%)`, 'info');
  });

  let resizeRaf = null;
  const onViewportResize = () => {
    if (!sheetWrapper.isConnected) {
      window.removeEventListener('resize', onViewportResize);
      return;
    }
    if (store.getState().zoomUserAdjusted) return;
    if (resizeRaf) cancelAnimationFrame(resizeRaf);
    resizeRaf = requestAnimationFrame(() => {
      const nextIdeal = computeDefaultDesktopZoom();
      if (Math.round(parseFloat(zoomSlider.value)) !== nextIdeal) {
        updateZoom(nextIdeal, { fromUser: false });
      }
    });
  };
  window.addEventListener('resize', onViewportResize);
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
  const btnEditOpen = container.querySelector('#btn-edit-open-copilot');
  const statusArea = container.querySelector('#copilot-status');
  const statusText = container.querySelector('#copilot-status-text');
  const statusTip = container.querySelector('#copilot-status-tip');
  const btnCancel = container.querySelector('#btn-cancel-copilot');
  const modelSelect = container.querySelector('#select-copilot-model');
  const linkVault = container.querySelector('#link-open-vault-from-copilot');

  // Quick AI Edit Bar di atas lembar kerja
  const quickEditInput = container.querySelector('#input-quick-ai-edit');
  const btnRunQuickEdit = container.querySelector('#btn-run-quick-ai-edit');
  const quickEditStatus = container.querySelector('#quick-ai-edit-status');
  const quickEditStatusText = container.querySelector('#quick-ai-edit-status-text');

  let copilotAbortController = null;

  const executePromptEditOnOpenDoc = async (instructionText, fromQuickBar = false) => {
    const cleanInstruction = String(instructionText || '').trim();
    if (!cleanInstruction) {
      showToast('Ketik instruksi penambahan atau perbaikan terlebih dahulu.', 'warning');
      return;
    }

    const selectedApiId = modelSelect ? modelSelect.value : store.getState().activeApiId;
    const vault = store.getApiVault();
    const activeEntry = vault.find(v => v.id === selectedApiId) || store.getActiveApi();
    const modelLabel = activeEntry?.label || activeEntry?.model || 'AI';

    copilotAbortController = new AbortController();
    if (btnRun) btnRun.disabled = true;
    if (btnEditOpen) btnEditOpen.disabled = true;
    if (btnRunQuickEdit) btnRunQuickEdit.disabled = true;

    if (fromQuickBar && quickEditStatus) {
      quickEditStatus.classList.remove('hidden');
      if (quickEditStatusText) quickEditStatusText.textContent = `Menerapkan instruksi edit dengan ${modelLabel}...`;
    } else if (statusArea) {
      statusArea.classList.remove('hidden');
      if (statusText) statusText.textContent = `Mengedit dokumen terbuka (${modelLabel})...`;
    }

    try {
      const state = store.getState();
      const hasDomCanvas = Boolean(container.querySelector('#editable-doc-title') || container.querySelector('#sheet-dynamic-blocks'));
      const currentDocModel = hasDomCanvas
        ? buildModelFromCanvas(container)
        : JSON.parse(JSON.stringify(state.document || buildModelFromCanvas(container)));
      const allowLayoutRestructure = shouldAllowLayoutRestructure(cleanInstruction);
      const hasExistingContent = !currentDocModel.isBlankDraft && (currentDocModel.blocks || []).some(
        (b) => b && (b.type === 'table' || b.type === 'identityTable' || (b.type === 'paragraph' && String(b.text || b.runs?.map(r => r.text).join('') || '').trim()))
      );

      if (hasExistingContent) {
        const indexObj = buildSurgicalIndexFromDocModel(currentDocModel);
        const editedData = await editDocumentWithAi(indexObj.snapshot, cleanInstruction, {
          schoolProfile: store.getSchoolProfile(),
          preferredApiId: selectedApiId,
          allowLayoutRestructure,
          signal: copilotAbortController.signal
        });

        const patchResult = applySurgicalPatchesToDocModel(currentDocModel, editedData, { allowLayoutRestructure });
        if (patchResult.mode === 'restructure_needed') {
          const fallbackSnap = extractCanvasSnapshot(container);
          applyEditedResultToCanvas(container, editedData, fallbackSnap);
        } else {
          const patchedDoc = {
            ...patchResult.docModel,
            ...(state.document?.verification ? { verification: state.document.verification } : {}),
            ...(state.document?.qrPlacement ? { qrPlacement: state.document.qrPlacement } : {})
          };
          store.clearActiveUploadedDocx();
          store.setState({ document: patchedDoc, docxViewMode: 'canvas' }, true);
          renderWordCanvas(container);
        }

        const baseNote = editedData.summaryOfChanges || 'Instruksi penambahan/perbaikan berhasil diterapkan!';
        const changeNote = patchResult.appliedCount > 0 && patchResult.mode !== 'restructure_needed'
          ? `${baseNote} (${patchResult.appliedCount} bagian diperbarui, tata letak & tabel tetap terjaga)`
          : baseNote;
        showToast(`✓ ${changeNote}`, 'success');
      } else {
        const currentSnapshot = extractCanvasSnapshot(container);
        const editedData = await editDocumentWithAi(currentSnapshot, cleanInstruction, {
          schoolProfile: store.getSchoolProfile(),
          preferredApiId: selectedApiId,
          allowLayoutRestructure,
          signal: copilotAbortController.signal
        });

        applyEditedResultToCanvas(container, editedData, currentSnapshot);
        const changeNote = editedData.summaryOfChanges || 'Instruksi penambahan/perbaikan berhasil diterapkan!';
        showToast(`✓ ${changeNote}`, 'success');
      }
      if (fromQuickBar && quickEditInput) {
        quickEditInput.value = '';
      }
    } catch (err) {
      if (!copilotAbortController?.signal?.aborted) {
        showToast(`Gagal mengedit dokumen: ${err.message}`, 'error');
      }
    } finally {
      copilotAbortController = null;
      if (btnRun) btnRun.disabled = false;
      if (btnEditOpen) btnEditOpen.disabled = false;
      if (btnRunQuickEdit) btnRunQuickEdit.disabled = false;
      if (quickEditStatus) quickEditStatus.classList.add('hidden');
      if (statusArea) statusArea.classList.add('hidden');
    }
  };

  if (btnRunQuickEdit && quickEditInput) {
    btnRunQuickEdit.addEventListener('click', () => {
      executePromptEditOnOpenDoc(quickEditInput.value, true);
    });
    quickEditInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        executePromptEditOnOpenDoc(quickEditInput.value, true);
      }
    });
  }

  container.querySelectorAll('.quick-ai-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      if (quickEditInput) {
        quickEditInput.value = chip.dataset.edit || '';
        quickEditInput.focus();
      }
    });
  });

  if (btnEditOpen) {
    btnEditOpen.addEventListener('click', () => {
      executePromptEditOnOpenDoc(promptInput.value, false);
    });
  }

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
      const kbSelection = store.getSmartAiReferencesForPrompt(userPrompt, 3);
      const prompt = buildCopilotUserPrompt({
        userPrompt,
        schoolName: state.schoolProfile?.schoolName,
        principalName: state.schoolProfile?.principalName,
        dateStr: new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
        references: kbSelection.references,
        matchMode: kbSelection.matchMode,
        totalKnowledgeCount: kbSelection.totalKnowledgeCount
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

      // Terapkan hasil AI langsung ke lembar kerja Word sebagai dokumen baru!
      store.setActiveVaultId(null, 'standard');
      applyAiResultToCanvas(container, aiResult);
      updateWordVaultBadge(container);
      showToast(describeAiMatchFeedback(kbSelection), 'success');

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
  container.querySelectorAll('.word-paper-sheet').forEach(sheet => {
    sheet.style.setProperty('--margin-top', `${margins.top}cm`);
    sheet.style.setProperty('--margin-bottom', `${margins.bottom}cm`);
    sheet.style.setProperty('--margin-left', `${margins.left}cm`);
    sheet.style.setProperty('--margin-right', `${margins.right}cm`);
  });

  const btnMarginMenu = container.querySelector('#btn-margin-menu');
  if (btnMarginMenu) {
    btnMarginMenu.innerHTML = `${ICON_RULER} <span>Margin (${margins.left}•${margins.top})</span>`;
  }
  store.updateLetterOptions({ margins });
  syncGlobalPrintPageStyle(store.getLetterOptions());
  showToast(`Margin lembar diatur: Kiri ${margins.left}cm, Kanan ${margins.right}cm, Atas ${margins.top}cm, Bawah ${margins.bottom}cm`, 'success');
  scheduleSync(container);
  scheduleWordPagination(container);
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
          forcedTypeCode: selectedCode,
          sequenceNumber: store.getNextOutgoingSequence()
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
      scheduleWordPagination(container);
    });
  });

  bindDynamicBlockControls(container);
}

function runsToWordSafeHtml(runs = []) {
  if (!Array.isArray(runs) || runs.length === 0) return '';
  return runs.map((r) => {
    let html = escapeHtml(r.text || '').replace(/\r?\n/g, '<br>');
    if (r.bold) html = `<b>${html}</b>`;
    if (r.italic) html = `<i>${html}</i>`;
    if (r.underline) html = `<u>${html}</u>`;
    if (r.sup) html = `<sup>${html}</sup>`;
    if (r.sub) html = `<sub>${html}</sub>`;
    return html;
  }).join('');
}

function extractWordInlineRuns(rootEl, inherited = {}) {
  const runs = [];
  if (!rootEl) return runs;

  for (const child of rootEl.childNodes || []) {
    if (child.nodeType === 3) {
      const txt = child.nodeValue || '';
      if (txt) {
        runs.push({
          text: txt,
          ...(inherited.bold ? { bold: true } : {}),
          ...(inherited.italic ? { italic: true } : {}),
          ...(inherited.underline ? { underline: true } : {}),
          ...(inherited.sup ? { sup: true } : {}),
          ...(inherited.sub ? { sub: true } : {})
        });
      }
      continue;
    }
    if (child.nodeType !== 1) continue;
    const tag = child.tagName.toLowerCase();
    if (tag === 'br') {
      runs.push({ text: '\n', ...inherited });
      continue;
    }
    const style = (child.getAttribute('style') || '').toLowerCase();
    const nextFmt = {
      bold: inherited.bold || tag === 'b' || tag === 'strong' || /font-weight\s*:\s*(bold|700|800|900)/.test(style),
      italic: inherited.italic || tag === 'i' || tag === 'em' || /font-style\s*:\s*italic/.test(style),
      underline: inherited.underline || tag === 'u' || /text-decoration[^;]*underline/.test(style),
      sup: inherited.sup || tag === 'sup',
      sub: inherited.sub || tag === 'sub'
    };
    runs.push(...extractWordInlineRuns(child, nextFmt));
  }

  const merged = [];
  for (const r of runs) {
    const prev = merged[merged.length - 1];
    if (
      prev &&
      Boolean(prev.bold) === Boolean(r.bold) &&
      Boolean(prev.italic) === Boolean(r.italic) &&
      Boolean(prev.underline) === Boolean(r.underline) &&
      Boolean(prev.sup) === Boolean(r.sup) &&
      Boolean(prev.sub) === Boolean(r.sub)
    ) {
      prev.text += r.text;
    } else {
      merged.push({ ...r });
    }
  }
  return merged;
}

/**
 * Merender blok paragraf tunggal beserta tombol aksi cepat (Tambah Teks Atas/Bawah, Sisipkan Tabel, Hapus Area)
 */
function buildParagraphBlockHtml({
  text = '',
  runs = null,
  anchorIndex = '',
  alignment = 'justify',
  isBold = false,
  isImported = false
} = {}) {
  const cleanText = String(text || '');
  const innerHtml = Array.isArray(runs) && runs.length > 0
    ? runsToWordSafeHtml(runs)
    : escapeHtml(cleanText);
  const checkStr = Array.isArray(runs) && runs.length > 0
    ? runs.map(r => r.text || '').join('')
    : cleanText;
  const isEmpty = checkStr.replace(/[\s\u00a0\u200b]/g, '').length === 0;
  const alignCss = alignment === 'center' ? 'center' : alignment === 'right' ? 'right' : alignment === 'left' ? 'left' : 'justify';
  const anchorAttr = (anchorIndex !== '' && anchorIndex !== undefined && anchorIndex !== -1)
    ? `data-anchor-index="${anchorIndex}"`
    : '';

  return `
    <div class="canvas-block-item paragraph-box-wrapper ${isEmpty ? 'is-empty-para' : ''}">
      <div class="block-hover-tools no-print">
        <button type="button" class="btn-block-tool" data-block-action="add-para-above" title="Tambah area teks di atas">+ Teks Atas</button>
        <button type="button" class="btn-block-tool" data-block-action="add-para-below" title="Tambah area teks di bawah (atau tekan Enter)">+ Teks Bawah</button>
        <button type="button" class="btn-block-tool" data-block-action="add-table-below" title="Sisipkan tabel rincian di bawah">+ Rincian</button>
        <button type="button" class="btn-block-tool btn-block-del" data-block-action="delete-block" title="Hapus area teks ini (atau tekan Backspace saat kosong)">✕ Hapus</button>
      </div>
      <div contenteditable="true" class="editable-field paragraph-box ${isImported ? 'imported-field' : ''}" ${anchorAttr} data-empty="${isEmpty ? 'true' : 'false'}" data-placeholder="Ketik teks di sini... (Tekan Backspace/Delete atau klik ✕ Hapus untuk menghapus area kosong ini)" style="line-height: 1.35; text-align: ${alignCss}; ${isBold ? 'font-weight: bold;' : ''}" spellcheck="false">${isEmpty ? '' : innerHtml}</div>
    </div>
  `;
}

/**
 * Merender blok Tabel Grid Berbingkai (multi-kolom fleksibel ala Microsoft Word)
 */
function buildCustomGridTableBlockHtml(b = {}) {
  const rows = Array.isArray(b.rows) && b.rows.length > 0
    ? b.rows
    : [
        [{ text: 'No', align: 'center' }, { text: 'Nama', align: 'center' }, { text: 'Jabatan / Keterangan', align: 'center' }],
        [{ text: '1', align: 'left' }, { text: '....................................', align: 'left' }, { text: 'Guru', align: 'center' }]
      ];
  const borderStyle = b.borderStyle || 'grid';
  const headerShading = b.headerShading !== false;
  const headerRow = b.headerRow !== false;
  const colWidths = Array.isArray(b.colWidthsPct) && b.colWidthsPct.length > 0 ? b.colWidthsPct.join(',') : '';
  const colgroupHtml = Array.isArray(b.colWidthsPct) && b.colWidthsPct.length > 0
    ? `<colgroup>${b.colWidthsPct.map(w => `<col style="width: ${w}%;">`).join('')}</colgroup>`
    : '';
  const tableWidthPct = Number(b.tableWidthPct) || 0;
  const widthAttr = (tableWidthPct >= 30 && tableWidthPct < 100)
    ? `data-table-width-pct="${tableWidthPct}" style="width: ${tableWidthPct}%;"`
    : '';

  const renderCellHtml = (c, tag, isHeadCell) => {
    const isObj = c && typeof c === 'object' && !Array.isArray(c);
    const txt = isObj
      ? (Array.isArray(c.paragraphs) ? c.paragraphs.map(p => p.text || '').join('\n') : (c.text ?? ''))
      : String(c ?? '');
    const cs = isObj && c.colspan > 1 ? ` colspan="${c.colspan}"` : '';
    const rs = isObj && c.rowspan > 1 ? ` rowspan="${c.rowspan}"` : '';
    const alignCss = isObj && c.align ? `text-align: ${c.align === 'both' ? 'justify' : c.align};` : '';
    const boldStyle = isObj && c.bold ? 'font-weight: 700;' : '';
    return `<${tag} contenteditable="true" spellcheck="false"${cs}${rs} style="${alignCss} ${boldStyle}">${escapeHtml(txt).replace(/\r?\n/g, '<br>')}</${tag}>`;
  };

  const rowsHtml = rows.map((rowEntry, idx) => {
    const isObj = rowEntry && typeof rowEntry === 'object' && !Array.isArray(rowEntry);
    const cells = isObj ? (rowEntry.cells || []) : (Array.isArray(rowEntry) ? rowEntry : []);
    const isHead = isObj && 'isHeader' in rowEntry ? Boolean(rowEntry.isHeader) : (headerRow && idx === 0);
    const minH = isObj && Number(rowEntry.minHeightCm) > 0 ? Number(rowEntry.minHeightCm) : 0;
    const trAttrs = [
      isHead ? 'data-header="true"' : '',
      minH > 0 ? `data-min-height="${minH}" style="height: ${minH}cm;"` : ''
    ].filter(Boolean).join(' ');
    const tag = isHead ? 'th' : 'td';
    return `<tr${trAttrs ? ` ${trAttrs}` : ''}>${cells.map(c => renderCellHtml(c, tag, isHead)).join('')}</tr>`;
  }).join('');

  return `
    <div class="canvas-block-item sheet-custom-table-wrapper" data-block-type="custom-table">
      <div class="block-hover-tools no-print">
        <button type="button" class="btn-block-tool" data-grid-act="add-row" title="Tambah baris ke tabel ini">+ Baris</button>
        <button type="button" class="btn-block-tool" data-grid-act="del-row" title="Hapus baris terakhir">− Baris</button>
        <button type="button" class="btn-block-tool" data-grid-act="add-col" title="Tambah kolom di sebelah kanan">+ Kolom</button>
        <button type="button" class="btn-block-tool" data-grid-act="del-col" title="Hapus kolom paling kanan">− Kolom</button>
        <button type="button" class="btn-block-tool" data-grid-act="cycle-border" title="Ganti gaya garis tabel (Penuh / SPPD / Tanpa Garis)">Gaya Garis</button>
        <button type="button" class="btn-block-tool" data-grid-act="toggle-header-bg" title="Aktif/nonaktifkan warna latar abu header">Header Abu/Putih</button>
        <button type="button" class="btn-block-tool" data-block-action="add-para-below" title="Tambah paragraf di bawah tabel">+ Teks Bawah</button>
        <button type="button" class="btn-block-tool btn-block-del" data-block-action="delete-block" title="Hapus tabel ini">✕ Hapus</button>
      </div>
      <table class="word-custom-grid-table ${borderStyle === 'sppd' ? 'grid-border-sppd' : borderStyle === 'none' ? 'grid-border-none' : ''} ${headerShading ? '' : 'grid-no-header-bg'}"
        data-border-style="${borderStyle}"
        data-header-shading="${headerShading ? 'true' : 'false'}"
        data-header-row="${headerRow ? 'true' : 'false'}"
        ${colWidths ? `data-col-widths="${colWidths}"` : ''}
        ${widthAttr}>
        ${colgroupHtml}
        <tbody>${rowsHtml}</tbody>
      </table>
    </div>
  `;
}

/**
 * Merender blok pemisah halaman manual (Page Break) di Lembar Kerja
 */
function buildPageBreakBlockHtml() {
  return `
    <div class="canvas-block-item sheet-pagebreak-wrapper" data-block-type="pagebreak">
      <div class="word-pagebreak-divider no-print">
        <span>✂ Pindah Halaman Baru (Page Break)</span>
        <button type="button" class="btn-block-tool btn-block-del" data-block-action="delete-block" title="Hapus pemisah halaman ini">✕ Hapus</button>
      </div>
    </div>
  `;
}

/**
 * Merender blok Tabel Rincian Identitas / Jadwal beserta tombol aksi baris & blok
 */
function buildIdentityTableBlockHtml(rows = [], colWidthsPct = null) {
  const safeRows = (Array.isArray(rows) && rows.length > 0) ? rows : [
    { label: 'Nama / Hari', value: ': ...........................' },
    { label: 'Keterangan / Tempat', value: ': ...........................' }
  ];
  const hasCustomW = Array.isArray(colWidthsPct) && colWidthsPct.length >= 2;
  const w0 = hasCustomW ? Number(colWidthsPct[0]) || 28 : 28;
  const w1 = hasCustomW ? Number(colWidthsPct[1]) || 72 : 72;

  return `
    <div class="canvas-block-item sheet-identity-wrapper" id="sheet-identity-wrapper">
      <div class="block-hover-tools no-print">
        <button type="button" class="btn-block-tool" data-block-action="add-table-row" title="Tambah baris rincian baru di tabel ini">+ Baris</button>
        <button type="button" class="btn-block-tool" data-block-action="add-para-above" title="Tambah area teks di atas tabel">+ Teks Atas</button>
        <button type="button" class="btn-block-tool" data-block-action="add-para-below" title="Tambah area teks di bawah tabel">+ Teks Bawah</button>
        <button type="button" class="btn-block-tool btn-block-del" data-block-action="delete-block" title="Hapus tabel rincian ini">✕ Hapus Tabel</button>
      </div>
      <table class="identity-table-word"${hasCustomW ? ` data-col-widths="${w0},${w1}"` : ''}>
        <tbody>
          ${safeRows.map(r => buildIdentityRowHtml(r, hasCustomW ? [w0, w1] : null)).join('')}
        </tbody>
      </table>
    </div>
  `;
}

function buildIdentityRowHtml(r = {}, colWidthsPct = null) {
  const rawVal = String(r.value || '').trim();
  const formattedVal = rawVal ? (rawVal.startsWith(':') ? rawVal : `: ${rawVal}`) : ': ';
  const anchorAttr = (r.index !== undefined && r.index !== -1) ? `data-anchor-index="${r.index}"` : '';
  const hasCustomW = Array.isArray(colWidthsPct) && colWidthsPct.length >= 2;
  return `
    <tr ${anchorAttr}>
      <td class="id-label" contenteditable="true" spellcheck="false"${hasCustomW ? ` style="width: ${colWidthsPct[0]}%;"` : ''}>${escapeHtml(r.label || '')}</td>
      <td class="id-val" contenteditable="true" spellcheck="false"${hasCustomW ? ` style="width: ${colWidthsPct[1]}%;"` : ''}>${escapeHtml(formattedVal)}</td>
      <td class="id-row-action-cell no-print">
        <button type="button" class="btn-row-mini" data-row-action="add-row-below" title="Tambah baris di bawah">+</button>
        <button type="button" class="btn-row-mini btn-row-del" data-row-action="delete-row" title="Hapus baris ini">✕</button>
      </td>
    </tr>
  `;
}

/**
 * Merender daftar blok dinamis untuk dokumen standar / hasil konversi / AI
 * Otomatis menyaring paragraf kosong bawaan agar tidak menimbulkan jarak renggang
 */
function renderDynamicStandardBlocksHtml(doc) {
  const rawBlocks = Array.isArray(doc?.blocks) ? doc.blocks : [];
  const bodyItems = rawBlocks.filter(b =>
    b.type === 'paragraph' || b.type === 'identityTable' || b.type === 'table' || b.type === 'pageBreak'
  );

  const htmlParts = [];
  for (const b of bodyItems) {
    if (b.type === 'pageBreak') {
      htmlParts.push(buildPageBreakBlockHtml());
    } else if (b.type === 'paragraph') {
      const text = (b.runs ? b.runs.map(r => r.text).join('') : (b.text || '')).trim();
      if (!text) continue; // Abaikan kotak kosong bawaan supaya tidak ada celah jauh
      htmlParts.push(buildParagraphBlockHtml({
        text,
        runs: b.runs || null,
        alignment: b.alignment || 'justify',
        isBold: Boolean(b.isBold)
      }));
    } else if (b.type === 'identityTable' && Array.isArray(b.rows) && b.rows.length > 0) {
      htmlParts.push(buildIdentityTableBlockHtml(b.rows, b.colWidthsPct || null));
    } else if (b.type === 'table' && Array.isArray(b.rows) && b.rows.length > 0) {
      htmlParts.push(buildCustomGridTableBlockHtml(b));
    }
  }

  // Jika dokumen masih kosong/baru, tampilkan satu area teks bersih siap ketik tanpa isi dummy
  if (htmlParts.length === 0) {
    htmlParts.push(buildParagraphBlockHtml({ text: '' }));
  }

  return htmlParts.join('');
}

/**
 * Merender daftar blok dinamis untuk dokumen .docx impor
 */
function renderDynamicImportedBlocksHtml(docxBodyList = [], docxIdentityRows = [], docxClosing = null) {
  const htmlParts = [];

  for (const p of docxBodyList) {
    const text = (p.currentText !== undefined ? p.currentText : (p.text || p.originalText || '')).trim();
    if (!text) continue; // Lewati paragraf kosong agar tidak ada jarak renggang
    htmlParts.push(buildParagraphBlockHtml({
      text,
      anchorIndex: p.index,
      alignment: p.alignment || 'justify',
      isBold: Boolean(p.isBold),
      isImported: true
    }));
  }

  if (Array.isArray(docxIdentityRows) && docxIdentityRows.length > 0) {
    htmlParts.push(buildIdentityTableBlockHtml(docxIdentityRows));
  }

  if (docxClosing) {
    const closingText = (docxClosing.currentText !== undefined ? docxClosing.currentText : (docxClosing.text || '')).trim();
    if (closingText) {
      htmlParts.push(buildParagraphBlockHtml({
        text: closingText,
        anchorIndex: docxClosing.index,
        alignment: 'justify',
        isImported: true
      }));
    }
  }

  if (htmlParts.length === 0) {
    htmlParts.push(buildParagraphBlockHtml({ text: '' }));
  }

  return htmlParts.join('');
}

/**
 * Mengumpulkan seluruh blok dinamis (.paragraph-box-wrapper & .sheet-identity-wrapper)
 * di seluruh halaman Lembar Kerja secara berurutan, sekaligus menggabungkan kembali
 * tabel/paragraf yang terbelah lintas halaman.
 */
function collectAllDynamicBlockNodes(container) {
  return [...container.querySelectorAll('.sheet-dynamic-blocks > .canvas-block-item')];
}

function collectMergedDynamicItems(container) {
  const nodes = collectAllDynamicBlockNodes(container);
  const items = [];
  for (const wrapper of nodes) {
    if (wrapper.classList.contains('sheet-pagebreak-wrapper')) {
      items.push({ type: 'pageBreak' });
      continue;
    }
    if (wrapper.classList.contains('paragraph-box-wrapper')) {
      const pBox = wrapper.querySelector('.paragraph-box');
      const text = pBox ? pBox.innerText.trim() : '';
      const runs = pBox ? extractWordInlineRuns(pBox) : [];
      const alignCss = (pBox?.style?.textAlign || '').toLowerCase();
      const alignment = alignCss === 'justify' ? 'both' : (alignCss || 'both');
      const groupId = wrapper.dataset.splitParaGroup || '';
      if (text) {
        const prev = items[items.length - 1];
        if (groupId && prev && prev.type === 'paragraph' && prev._splitParaGroup === groupId) {
          prev.text = `${prev.text} ${text}`;
          prev.runs = [...(prev.runs || [{ text: prev.text }]), { text: ' ' }, ...runs];
        } else {
          items.push({ type: 'paragraph', text, runs, alignment, _splitParaGroup: groupId });
        }
      }
    } else if (wrapper.classList.contains('sheet-identity-wrapper')) {
      const idTbl = wrapper.querySelector('.identity-table-word');
      const idColRaw = (idTbl?.dataset?.colWidths || '').trim();
      const idColWidths = idColRaw
        ? idColRaw.split(',').map(n => parseFloat(n)).filter(n => !Number.isNaN(n) && n > 0)
        : [];
      const rows = [];
      wrapper.querySelectorAll('.identity-table-word tr').forEach(tr => {
        const label = tr.querySelector('.id-label')?.innerText.trim() || '';
        const val = tr.querySelector('.id-val')?.innerText.trim() || '';
        if (label || val.replace(/^:\s*/, '')) {
          rows.push({ label, value: val });
        }
      });
      const groupId = wrapper.dataset.splitGroup || '';
      if (rows.length > 0) {
        const prev = items[items.length - 1];
        if (groupId && prev && prev.type === 'identityTable' && prev._splitGroup === groupId) {
          prev.rows.push(...rows);
        } else {
          items.push({
            type: 'identityTable',
            ...(idColWidths.length >= 2 ? { colWidthsPct: idColWidths } : {}),
            rows,
            _splitGroup: groupId
          });
        }
      }
    } else if (wrapper.classList.contains('sheet-custom-table-wrapper')) {
      const tbl = wrapper.querySelector('table.word-custom-grid-table');
      if (!tbl) continue;
      const borderStyle = tbl.dataset.borderStyle || 'grid';
      const headerShading = tbl.dataset.headerShading !== 'false';
      const headerRow = tbl.dataset.headerRow !== 'false';
      const tableWidthPct = parseFloat(tbl.dataset.tableWidthPct || '0') || 0;
      const colWidthsRaw = (tbl.dataset.colWidths || '').trim();
      const colWidthsPct = colWidthsRaw
        ? colWidthsRaw.split(',').map(n => parseFloat(n)).filter(n => !Number.isNaN(n) && n > 0)
        : [];
      const rows = [];
      let hasAnyMinH = false;
      tbl.querySelectorAll('tr').forEach((tr) => {
        const isHead = tr.dataset.header === 'true' || Boolean(tr.closest('thead'));
        const minH = parseFloat(tr.dataset.minHeight || '0') || 0;
        if (minH > 0) hasAnyMinH = true;
        const cells = [...tr.children].filter(el => el.tagName === 'TD' || el.tagName === 'TH').map((td) => {
          const cs = Math.max(1, parseInt(td.getAttribute('colspan') || '1', 10) || 1);
          const rs = Math.max(1, parseInt(td.getAttribute('rowspan') || '1', 10) || 1);
          const al = (td.style?.textAlign || '').toLowerCase();
          const alignVal = al === 'justify' ? 'both' : al;
          const bold = /700|bold/i.test(td.style?.fontWeight || '');
          const cellText = td.innerText.replace(/\u00a0/g, ' ').trim();
          if (cs > 1 || rs > 1 || alignVal || bold || cellText.includes('\n')) {
            return {
              text: cellText,
              ...(cs > 1 ? { colspan: cs } : {}),
              ...(rs > 1 ? { rowspan: rs } : {}),
              ...(alignVal ? { align: alignVal } : {}),
              ...(bold ? { bold: true } : {})
            };
          }
          return cellText;
        });
        rows.push((isHead || minH > 0) ? { isHeader: isHead, ...(minH > 0 ? { minHeightCm: minH } : {}), cells } : cells);
      });
      const groupId = wrapper.dataset.splitGroup || '';
      if (rows.length > 0) {
        const prev = items[items.length - 1];
        if (groupId && prev && prev.type === 'table' && prev._splitGroup === groupId) {
          prev.rows.push(...rows);
        } else {
          items.push({
            type: 'table',
            borderStyle,
            headerRow,
            headerShading,
            ...(colWidthsPct.length > 0 ? { colWidthsPct } : {}),
            ...(tableWidthPct >= 30 && tableWidthPct < 100 ? { tableWidthPct } : {}),
            rows,
            _splitGroup: groupId
          });
        }
      }
    }
  }
  return items.map(({ _splitGroup, _splitParaGroup, ...clean }) => clean);
}

function getActiveWordPageIndex(container) {
  const active = document.activeElement?.closest('.word-paper-sheet');
  if (active && container.contains(active)) {
    return Math.max(0, Number(active.dataset.wordPage || 0));
  }
  return 0;
}

function updateWordPageIndicators(container) {
  const wrapperEl = container?.querySelector('#word-sheet-wrapper');
  if (!wrapperEl) return;
  const pageWrappers = [...wrapperEl.querySelectorAll('.word-page-wrapper')];
  const totalPages = Math.max(1, pageWrappers.length);
  pageWrappers.forEach((pw, idx) => {
    pw.dataset.wordPage = String(idx);
    const sheet = pw.querySelector('.word-paper-sheet');
    if (sheet) sheet.dataset.wordPage = String(idx);
    const blocks = pw.querySelector('.sheet-dynamic-blocks');
    if (blocks) blocks.dataset.pageBlocks = String(idx);
    const headerBar = pw.querySelector('.word-page-header-bar');
    const pill = pw.querySelector('.word-page-pill');
    if (pill) pill.textContent = `Halaman ${idx + 1} dari ${totalPages}`;
    if (headerBar) headerBar.classList.toggle('hidden', totalPages <= 1);
  });

  const activeIdx = Math.min(getActiveWordPageIndex(container), totalPages - 1);
  const statusPagesEl = container.querySelector('#word-status-pages');
  if (statusPagesEl) {
    statusPagesEl.innerHTML = `${ICON_DOC} <span>Halaman ${activeIdx + 1} dari ${totalPages}</span>`;
  }
}

function createContinuationWordPage(container, pageIdx) {
  const baseSheet = container.querySelector('#word-sheet');
  const opts = store.getLetterOptions();
  const paperSize = opts.paperSize || 'a4';
  const orientation = opts.orientation || 'portrait';
  const margins = opts.margins || { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 };
  const fontFamily = opts.fontFamily || 'Times New Roman';
  const fontSize = opts.fontSize || '12';
  const lineSpacing = opts.lineSpacing || '1.15';
  const cssLine = lineSpacing === '1.0' ? '1.15' : (lineSpacing === '1.5' ? '1.6' : (lineSpacing === '2.0' ? '2.0' : '1.35'));
  const hasPilcrow = Boolean(baseSheet?.classList.contains('show-pilcrow-marks'));

  const dims = computePrintableDimsMm({ paperSize, orientation, margins });
  const pw = document.createElement('div');
  pw.className = 'word-page-wrapper';
  pw.dataset.wordPage = String(pageIdx);
  pw.innerHTML = `
    <div class="word-page-header-bar no-print">
      <span class="word-page-pill">Halaman ${pageIdx + 1}</span>
    </div>
    <div class="word-paper-sheet paper-${paperSize} ${orientation === 'landscape' ? 'landscape' : ''} ${hasPilcrow ? 'show-pilcrow-marks' : ''}"
      data-word-page="${pageIdx}"
      style="
        --margin-top: ${margins.top}cm;
        --margin-bottom: ${margins.bottom}cm;
        --margin-left: ${margins.left}cm;
        --margin-right: ${margins.right}cm;
        --printable-h-mm: ${dims.printableHeightMm}mm;
        font-family: ${fontFamily};
        font-size: ${fontSize}pt;
      ">
      <div class="sheet-content-editable" style="line-height: ${cssLine};">
        <div class="sheet-dynamic-blocks word-continuation-blocks" data-page-blocks="${pageIdx}"></div>
      </div>
      <div class="floating-qr-stage" data-qr-page="${pageIdx}"></div>
    </div>
  `;
  return pw;
}

function captureCanvasCaretState(rootEl) {
  const sel = window.getSelection();
  const activeEl = document.activeElement?.closest('[contenteditable="true"]');
  if (!activeEl || !rootEl.contains(activeEl)) return null;
  if (!sel || sel.rangeCount === 0 || !activeEl.contains(sel.anchorNode)) {
    return { el: activeEl, offset: (activeEl.textContent || '').length };
  }
  const range = sel.getRangeAt(0);
  const pre = document.createRange();
  pre.selectNodeContents(activeEl);
  pre.setEnd(range.startContainer, range.startOffset);
  return { el: activeEl, offset: pre.toString().length };
}

function restoreCanvasCaretState(rootEl, saved) {
  if (!saved || !saved.el || !saved.el.isConnected || !rootEl.contains(saved.el)) return;
  if (document.activeElement === saved.el) {
    const curSel = window.getSelection();
    if (curSel && curSel.rangeCount > 0 && saved.el.contains(curSel.anchorNode)) return;
  }
  const target = saved.el;
  let remaining = saved.offset;
  const walker = document.createTreeWalker(target, NodeFilter.SHOW_TEXT, null);
  let node;
  while ((node = walker.nextNode())) {
    const len = node.nodeValue.length;
    if (remaining <= len) {
      const sel = window.getSelection();
      const r = document.createRange();
      r.setStart(node, Math.max(0, remaining));
      r.collapse(true);
      sel.removeAllRanges();
      sel.addRange(r);
      target.focus();
      target.closest('tr, .paragraph-box-wrapper')?.scrollIntoView({ block: 'nearest' });
      return;
    }
    remaining -= len;
  }
  const sel = window.getSelection();
  const r = document.createRange();
  if (!(target.textContent || '').trim()) {
    r.setStart(target, 0);
    r.collapse(true);
  } else {
    r.selectNodeContents(target);
    r.collapse(false);
  }
  sel.removeAllRanges();
  sel.addRange(r);
  target.focus();
  target.closest('tr, .paragraph-box-wrapper')?.scrollIntoView({ block: 'nearest' });
}

function collapseToSingleWordPage(container) {
  const sheetWrapper = container?.querySelector('#word-sheet-wrapper');
  if (!sheetWrapper) return;
  const firstBody = sheetWrapper.querySelector('.word-page-wrapper[data-word-page="0"] .sheet-content-editable');
  const firstBlocks = sheetWrapper.querySelector('#sheet-dynamic-blocks');
  if (!firstBody || !firstBlocks) return;

  const extraPages = [...sheetWrapper.querySelectorAll('.word-page-wrapper:not([data-word-page="0"])')];
  extraPages.forEach(pw => {
    const blocksEl = pw.querySelector('.sheet-dynamic-blocks');
    if (blocksEl) {
      while (blocksEl.firstElementChild) {
        const node = blocksEl.firstElementChild;
        const prev = firstBlocks.lastElementChild;
        if (node.classList.contains('sheet-identity-wrapper') && node.dataset.splitGroup &&
            prev?.classList.contains('sheet-identity-wrapper') && prev.dataset.splitGroup === node.dataset.splitGroup) {
          const prevTbody = prev.querySelector('tbody');
          const curTbody = node.querySelector('tbody');
          if (prevTbody && curTbody) {
            while (curTbody.firstElementChild) prevTbody.appendChild(curTbody.firstElementChild);
            node.remove();
            continue;
          }
        }
        if (node.classList.contains('sheet-custom-table-wrapper') && node.dataset.splitGroup &&
            prev?.classList.contains('sheet-custom-table-wrapper') && prev.dataset.splitGroup === node.dataset.splitGroup) {
          const prevTbody = prev.querySelector('tbody');
          const curTbody = node.querySelector('tbody');
          if (prevTbody && curTbody) {
            while (curTbody.firstElementChild) prevTbody.appendChild(curTbody.firstElementChild);
            node.remove();
            continue;
          }
        }
        if (node.classList.contains('paragraph-box-wrapper') && node.dataset.splitParaGroup &&
            prev?.classList.contains('paragraph-box-wrapper') && prev.dataset.splitParaGroup === node.dataset.splitParaGroup) {
          const prevBox = prev.querySelector('.paragraph-box');
          const curBox = node.querySelector('.paragraph-box');
          if (prevBox && curBox) {
            prevBox.innerText = `${prevBox.innerText.trimEnd()} ${curBox.innerText.trimStart()}`.trim();
            node.remove();
            continue;
          }
        }
        firstBlocks.appendChild(node);
      }
    }
    const addBar = pw.querySelector('.canvas-add-block-bar');
    const sigWrap = pw.querySelector('.sheet-signature-wrapper');
    if (addBar) firstBody.appendChild(addBar);
    if (sigWrap) firstBody.appendChild(sigWrap);
    pw.remove();
  });
  updateWordPageIndicators(container);
}

function scheduleWordPagination(container) {
  if (!container) return;
  clearTimeout(wordPaginateTimer);
  wordPaginateTimer = setTimeout(() => {
    reflowWordCanvasPagination(container);
  }, 40);
}

/**
 * Mesin Paginasi Multi-Halaman ala Microsoft Word untuk Lembar Kerja:
 * - Menjamin setiap lembar kertas menghormati margin atas/bawah kertas (A4, F4, Letter, Legal).
 * - Ketika teks/tabel melebihi batas margin bawah Halaman i, otomatis membuat Halaman i+1
 *   dan memindahkan baris/blok yang meluap dengan posisi kursor tetap terjaga.
 */
function reflowWordCanvasPagination(container) {
  if (isWordPaginating || !container) return;
  const sheetWrapper = container.querySelector('#word-sheet-wrapper');
  if (!sheetWrapper || sheetWrapper.querySelector('#docx-native-preview-card')) return;
  const basePage = sheetWrapper.querySelector('.word-page-wrapper[data-word-page="0"]');
  if (!basePage) return;

  isWordPaginating = true;
  const savedCaret = captureCanvasCaretState(sheetWrapper);

  try {
    const addBar = sheetWrapper.querySelector('.canvas-add-block-bar');
    const sigWrapper = sheetWrapper.querySelector('.sheet-signature-wrapper');

    const getBottomInSheet = (el, sheetRect, scale) => {
      if (!el || !scale) return 0;
      return (el.getBoundingClientRect().bottom - sheetRect.top) / scale;
    };

    const getTopInSheet = (el, sheetRect, scale) => {
      if (!el || !scale) return 0;
      return (el.getBoundingClientRect().top - sheetRect.top) / scale;
    };

    const ensureNextPage = (curPageIdx, pagesArr) => {
      if (pagesArr[curPageIdx + 1]) return pagesArr[curPageIdx + 1];
      const nextPw = createContinuationWordPage(container, curPageIdx + 1);
      pagesArr[curPageIdx].insertAdjacentElement('afterend', nextPw);
      return nextPw;
    };

    let guard = 0;
    while (guard < 80) {
      guard++;
      let moved = false;
      const pages = [...sheetWrapper.querySelectorAll('.word-page-wrapper')];

      for (let i = 0; i < pages.length; i++) {
        const curPw = pages[i];
        const sheet = curPw.querySelector('.word-paper-sheet');
        const curBody = sheet?.querySelector('.sheet-content-editable');
        const curBlocksEl = curBody?.querySelector('.sheet-dynamic-blocks');
        if (!sheet || !curBody || !curBlocksEl) continue;

        const cs = getComputedStyle(sheet);
        const minH = parseFloat(cs.minHeight) || 0;
        const padBottom = parseFloat(cs.paddingBottom) || 0;
        if (!minH || minH < 100) continue;

        const sheetRect = sheet.getBoundingClientRect();
        const scale = sheet.offsetWidth > 0 ? sheetRect.width / sheet.offsetWidth : 1;
        if (scale <= 0) continue;
        const maxBottomPx = minH - padBottom - 4;

        // 1A-0. Cek apakah ada pemisah halaman manual (.sheet-pagebreak-wrapper) sebelum blok paling akhir
        const manualBreakNode = curBlocksEl.querySelector('.sheet-pagebreak-wrapper');
        if (manualBreakNode && manualBreakNode.nextElementSibling) {
          const nextPw = ensureNextPage(i, pages);
          const nextBody = nextPw.querySelector('.sheet-content-editable');
          const nextBlocksEl = nextBody.querySelector('.sheet-dynamic-blocks');
          if (addBar && addBar.parentNode === curBody) nextBody.appendChild(addBar);
          if (sigWrapper && sigWrapper.parentNode === curBody) nextBody.appendChild(sigWrapper);

          const trailingNodes = [];
          let cursor = manualBreakNode.nextElementSibling;
          while (cursor) {
            trailingNodes.push(cursor);
            cursor = cursor.nextElementSibling;
          }
          for (let tIdx = trailingNodes.length - 1; tIdx >= 0; tIdx--) {
            nextBlocksEl.insertBefore(trailingNodes[tIdx], nextBlocksEl.firstChild);
          }
          moved = true;
          break;
        }

        // 1A. Jika blok TTD berada di halaman ini dan melewati batas margin bawah -> pindahkan ke halaman berikutnya
        if (sigWrapper && sigWrapper.parentNode === curBody && !sigWrapper.classList.contains('hidden')) {
          if (getBottomInSheet(sigWrapper, sheetRect, scale) > maxBottomPx || manualBreakNode) {
            const nextPw = ensureNextPage(i, pages);
            const nextBody = nextPw.querySelector('.sheet-content-editable');
            if (addBar && addBar.parentNode === curBody) nextBody.appendChild(addBar);
            nextBody.appendChild(sigWrapper);
            moved = true;
            break;
          }
        }

        // 1B. Cek blok konten paling bawah di .sheet-dynamic-blocks
        const lastBlock = curBlocksEl.lastElementChild;
        if (lastBlock && getBottomInSheet(lastBlock, sheetRect, scale) > maxBottomPx) {
          const nextPw = ensureNextPage(i, pages);
          const nextBody = nextPw.querySelector('.sheet-content-editable');
          const nextBlocksEl = nextBody.querySelector('.sheet-dynamic-blocks');

          // Jaga urutan dokumen: pastikan addBar & sigWrapper berada setelah seluruh blok
          if (addBar && addBar.parentNode === curBody) nextBody.appendChild(addBar);
          if (sigWrapper && sigWrapper.parentNode === curBody) nextBody.appendChild(sigWrapper);

          if (lastBlock.classList.contains('sheet-identity-wrapper')) {
            const tbody = lastBlock.querySelector('tbody');
            const rows = tbody ? [...tbody.querySelectorAll('tr')] : [];
            const firstRowFits = rows.length > 0 && getBottomInSheet(rows[0], sheetRect, scale) <= maxBottomPx;

            if (rows.length > 1 && (firstRowFits || curBlocksEl.children.length === 1)) {
              const groupId = lastBlock.dataset.splitGroup || `wtbl-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
              lastBlock.dataset.splitGroup = groupId;
              let destWrapper = nextBlocksEl.firstElementChild;
              if (!destWrapper || !destWrapper.classList.contains('sheet-identity-wrapper') || destWrapper.dataset.splitGroup !== groupId) {
                const shell = document.createElement('div');
                shell.innerHTML = buildIdentityTableBlockHtml([{ label: '', value: '' }]).trim();
                destWrapper = shell.firstElementChild;
                destWrapper.removeAttribute('id');
                destWrapper.dataset.splitGroup = groupId;
                const destBody = destWrapper.querySelector('tbody');
                if (destBody) destBody.innerHTML = '';
                nextBlocksEl.insertBefore(destWrapper, nextBlocksEl.firstChild);
              }
              const destTbody = destWrapper.querySelector('tbody');
              destTbody?.insertBefore(rows[rows.length - 1], destTbody.firstChild);
              moved = true;
              break;
            }
          }

          if (lastBlock.classList.contains('paragraph-box-wrapper')) {
            const pBox = lastBlock.querySelector('.paragraph-box');
            const fullText = pBox ? pBox.innerText : '';
            const words = fullText.split(/\s+/).filter(Boolean);
            const blockTopPx = getTopInSheet(lastBlock, sheetRect, scale);
            const blockBottomPx = getBottomInSheet(lastBlock, sheetRect, scale);
            const blockHeightPx = blockBottomPx - blockTopPx;

            // Jika 1 paragraf sangat panjang menyilang jauh dari bagian tengah kertas -> pecah rapi per kata
            if (words.length >= 18 && blockHeightPx > 110 && blockTopPx < maxBottomPx - 80) {
              const availH = Math.max(40, maxBottomPx - blockTopPx - 8);
              const splitIdx = Math.max(4, Math.min(words.length - 4, Math.floor(words.length * (availH / blockHeightPx) * 0.92)));
              const keepText = words.slice(0, splitIdx).join(' ');
              const tailText = words.slice(splitIdx).join(' ');
              const groupId = lastBlock.dataset.splitParaGroup || `wpara-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;
              lastBlock.dataset.splitParaGroup = groupId;
              pBox.innerText = keepText;

              let destParaWrapper = nextBlocksEl.firstElementChild;
              if (destParaWrapper && destParaWrapper.classList.contains('paragraph-box-wrapper') && destParaWrapper.dataset.splitParaGroup === groupId) {
                const destBox = destParaWrapper.querySelector('.paragraph-box');
                if (destBox) destBox.innerText = `${tailText} ${destBox.innerText}`.trim();
              } else {
                const tmp = document.createElement('div');
                tmp.innerHTML = buildParagraphBlockHtml({ text: tailText }).trim();
                destParaWrapper = tmp.firstElementChild;
                destParaWrapper.dataset.splitParaGroup = groupId;
                nextBlocksEl.insertBefore(destParaWrapper, nextBlocksEl.firstChild);
              }
              moved = true;
              break;
            }
          }

          // Pindahkan seluruh blok ke halaman berikutnya selama bukan satu-satunya blok yang buntu
          if (curBlocksEl.children.length > 1 || i === 0) {
            nextBlocksEl.insertBefore(lastBlock, nextBlocksEl.firstChild);
            moved = true;
            break;
          }
        }

        // 2. UNDERFLOW CHECK: Tarik kembali elemen dari halaman i+1 bila kini cukup ruang pada halaman i
        const nextPw = pages[i + 1];
        if (!nextPw) continue;
        if (curBlocksEl.querySelector('.sheet-pagebreak-wrapper')) continue;
        const nextBody = nextPw.querySelector('.sheet-content-editable');
        const nextBlocksEl = nextBody?.querySelector('.sheet-dynamic-blocks');
        if (!nextBlocksEl) continue;

        const firstNextBlock = nextBlocksEl.firstElementChild;
        if (firstNextBlock) {
          const curLastBlock = curBlocksEl.lastElementChild;

          // 2A. Gabung kembali baris tabel terbelah
          if (firstNextBlock.classList.contains('sheet-identity-wrapper') &&
              firstNextBlock.dataset.splitGroup &&
              curLastBlock?.classList.contains('sheet-identity-wrapper') &&
              curLastBlock.dataset.splitGroup === firstNextBlock.dataset.splitGroup) {
            const nextTbody = firstNextBlock.querySelector('tbody');
            const curTbody = curLastBlock.querySelector('tbody');
            const candidateRow = nextTbody?.firstElementChild;
            if (curTbody && candidateRow) {
              curTbody.appendChild(candidateRow);
              if (getBottomInSheet(candidateRow, sheetRect, scale) <= maxBottomPx) {
                if (nextTbody.children.length === 0) firstNextBlock.remove();
                moved = true;
                break;
              } else {
                nextTbody.insertBefore(candidateRow, nextTbody.firstChild);
              }
            }
          }

          // 2B. Gabung kembali paragraf terbelah
          if (firstNextBlock.classList.contains('paragraph-box-wrapper') &&
              firstNextBlock.dataset.splitParaGroup &&
              curLastBlock?.classList.contains('paragraph-box-wrapper') &&
              curLastBlock.dataset.splitParaGroup === firstNextBlock.dataset.splitParaGroup) {
            const curBox = curLastBlock.querySelector('.paragraph-box');
            const nextBox = firstNextBlock.querySelector('.paragraph-box');
            if (curBox && nextBox) {
              const prevText = curBox.innerText;
              curBox.innerText = `${prevText.trimEnd()} ${nextBox.innerText.trimStart()}`.trim();
              if (getBottomInSheet(curLastBlock, sheetRect, scale) <= maxBottomPx) {
                firstNextBlock.remove();
                moved = true;
                break;
              } else {
                curBox.innerText = prevText;
              }
            }
          }

          // 2C. Tarik kembali satu blok utuh dari halaman i+1
          const hasTrailingSigOnCur = Boolean(sigWrapper && sigWrapper.parentNode === curBody && !sigWrapper.classList.contains('hidden'));
          if (!hasTrailingSigOnCur) {
            curBlocksEl.appendChild(firstNextBlock);
            if (getBottomInSheet(firstNextBlock, sheetRect, scale) <= maxBottomPx) {
              moved = true;
              break;
            } else {
              nextBlocksEl.insertBefore(firstNextBlock, nextBlocksEl.firstChild);
            }
          }
        } else if (i + 1 === pages.length - 1) {
          // Halaman i+1 tidak punya blok lagi -> cek apakah addBar / TTD muat kembali ke halaman i
          const hasVisibleSig = Boolean(sigWrapper && sigWrapper.parentNode === nextBody && !sigWrapper.classList.contains('hidden'));
          const hasAddBar = Boolean(addBar && addBar.parentNode === nextBody);
          if (hasVisibleSig) {
            if (hasAddBar) curBody.appendChild(addBar);
            curBody.appendChild(sigWrapper);
            if (getBottomInSheet(sigWrapper, sheetRect, scale) <= maxBottomPx) {
              nextPw.remove();
              moved = true;
              break;
            } else {
              if (hasAddBar) nextBody.appendChild(addBar);
              nextBody.appendChild(sigWrapper);
            }
          } else if (hasAddBar) {
            curBody.appendChild(addBar);
            if (sigWrapper && sigWrapper.parentNode === nextBody) curBody.appendChild(sigWrapper);
            nextPw.remove();
            moved = true;
            break;
          }
        }
      }

      if (!moved) break;
    }

    // Bersihkan halaman lanjutan kosong yang tidak memiliki isi maupun TTD aktif
    const remainingPages = [...sheetWrapper.querySelectorAll('.word-page-wrapper')];
    for (let i = remainingPages.length - 1; i >= 1; i--) {
      const pw = remainingPages[i];
      const blocksEl = pw.querySelector('.sheet-dynamic-blocks');
      const hasBlocks = Boolean(blocksEl && blocksEl.children.length > 0);
      const hasSig = Boolean(pw.querySelector('.sheet-signature-wrapper:not(.hidden)'));
      const prevHasBreak = Boolean(i > 0 && remainingPages[i - 1]?.querySelector('.sheet-pagebreak-wrapper'));
      if (!hasBlocks && !hasSig && !prevHasBreak) {
        const firstBody = sheetWrapper.querySelector('.word-page-wrapper[data-word-page="0"] .sheet-content-editable');
        const addBarNode = pw.querySelector('.canvas-add-block-bar');
        const sigNode = pw.querySelector('.sheet-signature-wrapper');
        if (firstBody && addBarNode) firstBody.appendChild(addBarNode);
        if (firstBody && sigNode) firstBody.appendChild(sigNode);
        pw.remove();
      }
    }

    updateWordPageIndicators(container);
    refreshWordFloatingQrOverlay(container);
  } finally {
    isWordPaginating = false;
    restoreCanvasCaretState(sheetWrapper, savedCaret);
  }
}

function refreshWordFloatingQrOverlay(container) {
  if (!container || !document.body.contains(container)) return;
  const st = store.getState();
  const doc = st.document || {};
  const verification = doc.verification || null;
  const opts = store.getLetterOptions();
  const dims = computePrintableDimsMm(opts);

  const sheets = [...container.querySelectorAll('.word-paper-sheet')];
  sheets.forEach((sh, idx) => {
    sh.style.setProperty('--printable-h-mm', `${dims.printableHeightMm}mm`);
    let stage = sh.querySelector(':scope > .floating-qr-stage');
    if (!stage) {
      stage = document.createElement('div');
      stage.className = 'floating-qr-stage';
      stage.dataset.qrPage = String(idx);
      sh.appendChild(stage);
    }
    stage.dataset.qrPage = String(idx);
  });

  const stages = sheets.map(sh => sh.querySelector(':scope > .floating-qr-stage')).filter(Boolean);
  const toggleBtn = container.querySelector('#btn-toggle-word-qr-canvas');

  const mounted = mountInteractiveFloatingQr({
    stages,
    verification,
    rawPlacement: doc.qrPlacement || null,
    docModel: doc,
    school: st.schoolProfile || {},
    getZoomScale: () => {
      const zPct = Number(container.querySelector('#input-zoom-slider')?.value || store.getState().zoomLevel || 100);
      return Math.max(0.4, zPct / 100);
    },
    onPlacementChange: (nextPlacement) => {
      store.setActiveQrPlacement(nextPlacement, 'standard');
      updateCanvasQrToggleButton(toggleBtn, {
        verification: store.getState().document?.verification,
        placement: nextPlacement
      });
    }
  });

  if (mounted && !doc.qrPlacement && verification?.isVerifiedSah) {
    store.setActiveQrPlacement(mounted, 'standard');
  }
  updateCanvasQrToggleButton(toggleBtn, {
    verification,
    placement: store.getState().document?.qrPlacement || mounted
  });
}

/**
 * Mengatur interaksi keyboard (Backspace/Delete untuk hapus kotak kosong, Enter untuk tambah kotak baru)
 * serta tombol aksi cepat (+ Teks Atas, + Teks Bawah, + Tabel, ✕ Hapus) lintas seluruh halaman kanvas
 */
function bindDynamicBlockControls(container) {
  const sheetWrapper = container.querySelector('#word-sheet-wrapper');
  const firstBlocksContainer = container.querySelector('#sheet-dynamic-blocks');
  if (!sheetWrapper || !firstBlocksContainer) return;

  const getAllBlockNodes = () => collectAllDynamicBlockNodes(container);
  const getLastBlocksContainer = () => {
    const all = [...sheetWrapper.querySelectorAll('.sheet-dynamic-blocks')];
    return all[all.length - 1] || firstBlocksContainer;
  };

  const isCaretAtStart = (el) => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount || !sel.isCollapsed) return false;
    const range = sel.getRangeAt(0);
    const preRange = range.cloneRange();
    preRange.selectNodeContents(el);
    preRange.setEnd(range.startContainer, range.startOffset);
    return preRange.toString().replace(/[\u200b]/g, '').length === 0;
  };

  const isCaretAtEnd = (el) => {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount || !sel.isCollapsed) return false;
    const range = sel.getRangeAt(0);
    const postRange = range.cloneRange();
    postRange.selectNodeContents(el);
    postRange.setStart(range.endContainer, range.endOffset);
    return postRange.toString().replace(/[\u200b]/g, '').length === 0;
  };

  const focusAtEdge = (el, toEnd = true) => {
    if (!el) return;
    el.focus();
    const sel = window.getSelection();
    if (!sel) return;
    const range = document.createRange();
    range.selectNodeContents(el);
    range.collapse(!toEnd);
    sel.removeAllRanges();
    sel.addRange(range);
  };

  const syncParaEmptyAttr = (paraEl) => {
    if (!paraEl) return;
    const plain = paraEl.innerText.replace(/[\s\u00a0\u200b]/g, '');
    const isEmpty = plain.length === 0;
    paraEl.setAttribute('data-empty', isEmpty ? 'true' : 'false');
    const wrapper = paraEl.closest('.paragraph-box-wrapper');
    if (wrapper) {
      wrapper.classList.toggle('is-empty-para', isEmpty);
    }
    if (isEmpty && paraEl.innerHTML !== '') {
      paraEl.innerHTML = '';
    }
  };

  const clearImportedAnchorIfRemoved = (elOrWrapper) => {
    const active = store.getState().activeUploadedDocx;
    if (!active || !Array.isArray(active.paragraphs)) return;
    const anchorEls = [];
    if (elOrWrapper.hasAttribute && elOrWrapper.hasAttribute('data-anchor-index')) {
      anchorEls.push(elOrWrapper);
    }
    if (elOrWrapper.querySelectorAll) {
      elOrWrapper.querySelectorAll('[data-anchor-index]').forEach(n => anchorEls.push(n));
    }
    let changed = false;
    anchorEls.forEach(node => {
      const idx = parseInt(node.dataset.anchorIndex, 10);
      if (!isNaN(idx)) {
        const found = active.paragraphs.find(p => p.index === idx);
        if (found) {
          found.currentText = '';
          changed = true;
        }
      }
    });
    if (changed) store.setActiveUploadedDocx(active);
  };

  const removeBlockWrapper = (wrapper, focusDirection = 'prev') => {
    if (!wrapper) return;
    const allNodesBefore = getAllBlockNodes();
    const idx = allNodesBefore.indexOf(wrapper);
    const prevBlock = idx > 0 ? allNodesBefore[idx - 1] : null;
    const nextBlock = (idx >= 0 && idx < allNodesBefore.length - 1) ? allNodesBefore[idx + 1] : null;

    clearImportedAnchorIfRemoved(wrapper);
    wrapper.remove();

    // Jika semua blok di seluruh halaman habis dihapus, sisakan 1 kotak kosong di Halaman 1
    if (getAllBlockNodes().length === 0) {
      const temp = document.createElement('div');
      temp.innerHTML = buildParagraphBlockHtml({ text: '' }).trim();
      const freshWrapper = temp.firstElementChild;
      firstBlocksContainer.appendChild(freshWrapper);
      const freshPara = freshWrapper.querySelector('.paragraph-box');
      focusAtEdge(freshPara, false);
    } else if (focusDirection === 'prev' && prevBlock) {
      const target = prevBlock.querySelector('.paragraph-box') || prevBlock.querySelector('.id-val, [contenteditable="true"]');
      focusAtEdge(target, true);
    } else if (nextBlock) {
      const target = nextBlock.querySelector('.paragraph-box') || nextBlock.querySelector('.id-label, [contenteditable="true"]');
      focusAtEdge(target, false);
    } else if (prevBlock) {
      const target = prevBlock.querySelector('.paragraph-box') || prevBlock.querySelector('.id-val, [contenteditable="true"]');
      focusAtEdge(target, true);
    }

    scheduleSync(container);
    updateWordCount(container);
    scheduleWordPagination(container);
  };

  const insertParagraphBlock = (refWrapper, position = 'after', initialText = '') => {
    const temp = document.createElement('div');
    temp.innerHTML = buildParagraphBlockHtml({ text: initialText }).trim();
    const newWrapper = temp.firstElementChild;
    const hostContainer = refWrapper?.closest('.sheet-dynamic-blocks') || getLastBlocksContainer();

    if (refWrapper && refWrapper.parentNode === hostContainer) {
      if (position === 'before') {
        hostContainer.insertBefore(newWrapper, refWrapper);
      } else {
        refWrapper.insertAdjacentElement('afterend', newWrapper);
      }
    } else {
      hostContainer.appendChild(newWrapper);
    }

    const newPara = newWrapper.querySelector('.paragraph-box');
    syncParaEmptyAttr(newPara);
    focusAtEdge(newPara, false);
    scheduleSync(container);
    updateWordCount(container);
    reflowWordCanvasPagination(container);
    return newWrapper;
  };

  const insertTableBlock = (refWrapper, position = 'after') => {
    const temp = document.createElement('div');
    temp.innerHTML = buildIdentityTableBlockHtml([
      { label: 'Hari / Tanggal', value: ': ...........................' },
      { label: 'Waktu', value: ': 08.00 WIB s.d Selesai' },
      { label: 'Tempat', value: ': SMK IT Al-Hawari' }
    ]).trim();
    const newWrapper = temp.firstElementChild;
    const hostContainer = refWrapper?.closest('.sheet-dynamic-blocks') || getLastBlocksContainer();

    if (refWrapper && refWrapper.parentNode === hostContainer) {
      if (position === 'before') {
        hostContainer.insertBefore(newWrapper, refWrapper);
      } else {
        refWrapper.insertAdjacentElement('afterend', newWrapper);
      }
    } else {
      hostContainer.appendChild(newWrapper);
    }

    const firstCell = newWrapper.querySelector('.id-label');
    focusAtEdge(firstCell, true);
    scheduleSync(container);
    updateWordCount(container);
    reflowWordCanvasPagination(container);
    return newWrapper;
  };

  const insertGridTableBlock = (refWrapper, position = 'after') => {
    const temp = document.createElement('div');
    temp.innerHTML = buildCustomGridTableBlockHtml({
      borderStyle: 'grid',
      headerRow: true,
      headerShading: false,
      rows: [
        [{ text: 'No', align: 'center' }, { text: 'Nama', align: 'center' }, { text: 'Jabatan / Keterangan', align: 'center' }],
        [{ text: '1', align: 'left' }, { text: 'Nama Pegawai / Peserta', align: 'left' }, { text: 'Guru', align: 'center' }],
        [{ text: '2', align: 'left' }, { text: 'Nama Pegawai / Peserta', align: 'left' }, { text: 'Guru', align: 'center' }]
      ]
    }).trim();
    const newWrapper = temp.firstElementChild;
    const hostContainer = refWrapper?.closest('.sheet-dynamic-blocks') || getLastBlocksContainer();

    if (refWrapper && refWrapper.parentNode === hostContainer) {
      if (position === 'before') hostContainer.insertBefore(newWrapper, refWrapper);
      else refWrapper.insertAdjacentElement('afterend', newWrapper);
    } else {
      hostContainer.appendChild(newWrapper);
    }

    const firstCell = newWrapper.querySelector('th, td');
    if (firstCell) focusAtEdge(firstCell, true);
    scheduleSync(container);
    updateWordCount(container);
    reflowWordCanvasPagination(container);
    return newWrapper;
  };

  const insertPageBreakBlock = (refWrapper, position = 'after') => {
    const temp = document.createElement('div');
    temp.innerHTML = buildPageBreakBlockHtml().trim();
    const newWrapper = temp.firstElementChild;
    const hostContainer = refWrapper?.closest('.sheet-dynamic-blocks') || getLastBlocksContainer();

    if (refWrapper && refWrapper.parentNode === hostContainer) {
      if (position === 'before') hostContainer.insertBefore(newWrapper, refWrapper);
      else refWrapper.insertAdjacentElement('afterend', newWrapper);
    } else {
      hostContainer.appendChild(newWrapper);
    }

    insertParagraphBlock(newWrapper, 'after', '');
    scheduleSync(container);
    updateWordCount(container);
    reflowWordCanvasPagination(container);
    return newWrapper;
  };

  // Sinkronisasi status kosong awal pada semua paragraf
  sheetWrapper.querySelectorAll('.sheet-dynamic-blocks .paragraph-box').forEach(p => syncParaEmptyAttr(p));

  sheetWrapper.addEventListener('focusin', () => {
    updateWordPageIndicators(container);
  });

  // Event Delegation: Input di dalam seluruh .sheet-dynamic-blocks lintas halaman
  sheetWrapper.addEventListener('input', (e) => {
    if (!e.target.closest('.sheet-dynamic-blocks')) return;
    const paraEl = e.target.closest('.paragraph-box');
    if (paraEl) {
      syncParaEmptyAttr(paraEl);
      if (paraEl.classList.contains('imported-field')) {
        const anchorIdx = parseInt(paraEl.dataset.anchorIndex, 10);
        const active = store.getState().activeUploadedDocx;
        if (active && active.paragraphs && !isNaN(anchorIdx)) {
          const found = active.paragraphs.find(p => p.index === anchorIdx);
          if (found) {
            found.currentText = paraEl.innerText;
            store.setActiveUploadedDocx(active);
          }
        }
      }
    }
    scheduleSync(container);
    updateWordCount(container);
    scheduleWordPagination(container);
  });

  // Event Delegation: Keydown (Backspace, Delete, Enter) pada Paragraf & Baris Tabel lintas halaman
  sheetWrapper.addEventListener('keydown', (e) => {
    if (!e.target.closest('.sheet-dynamic-blocks')) return;
    const paraEl = e.target.closest('.paragraph-box');
    if (paraEl) {
      const wrapper = paraEl.closest('.paragraph-box-wrapper');
      const plain = paraEl.innerText.replace(/[\s\u00a0\u200b]/g, '');

      // 1. Jika area teks sudah kosong dan pengguna menekan Backspace atau Delete -> Hapus kotak teks ini!
      if ((e.key === 'Backspace' || e.key === 'Delete') && plain.length === 0) {
        e.preventDefault();
        removeBlockWrapper(wrapper, e.key === 'Backspace' ? 'prev' : 'next');
        return;
      }

      // 2. Jika menekan Backspace tepat di ujung awal paragraf (termasuk lintas batas halaman)
      if (e.key === 'Backspace' && isCaretAtStart(paraEl)) {
        const allNodes = getAllBlockNodes();
        const curIdx = allNodes.indexOf(wrapper);
        const prevWrapper = curIdx > 0 ? allNodes[curIdx - 1] : null;
        if (prevWrapper && prevWrapper.classList.contains('paragraph-box-wrapper')) {
          const prevPara = prevWrapper.querySelector('.paragraph-box');
          const prevPlain = prevPara ? prevPara.innerText.replace(/[\s\u00a0\u200b]/g, '') : '';
          e.preventDefault();
          if (prevPlain.length === 0) {
            removeBlockWrapper(prevWrapper, 'next');
          } else if (prevPara) {
            const prevText = prevPara.innerText.trimEnd();
            const currText = paraEl.innerText.trimStart();
            prevPara.innerText = currText ? `${prevText} ${currText}` : prevText;
            syncParaEmptyAttr(prevPara);
            removeBlockWrapper(wrapper, 'prev');
          }
          return;
        }
      }

      // 3. Jika menekan Delete tepat di ujung akhir paragraf (termasuk lintas batas halaman)
      if (e.key === 'Delete' && isCaretAtEnd(paraEl)) {
        const allNodes = getAllBlockNodes();
        const curIdx = allNodes.indexOf(wrapper);
        const nextWrapper = (curIdx >= 0 && curIdx < allNodes.length - 1) ? allNodes[curIdx + 1] : null;
        if (nextWrapper && nextWrapper.classList.contains('paragraph-box-wrapper')) {
          const nextPara = nextWrapper.querySelector('.paragraph-box');
          const nextPlain = nextPara ? nextPara.innerText.replace(/[\s\u00a0\u200b]/g, '') : '';
          e.preventDefault();
          if (nextPlain.length === 0) {
            removeBlockWrapper(nextWrapper, 'prev');
          } else if (nextPara) {
            const currText = paraEl.innerText.trimEnd();
            const nextText = nextPara.innerText.trimStart();
            paraEl.innerText = nextText ? `${currText} ${nextText}` : currText;
            syncParaEmptyAttr(paraEl);
            removeBlockWrapper(nextWrapper, 'prev');
          }
          return;
        }
      }

      // 4. Tekan Enter (tanpa Shift) -> Buat area teks / paragraf baru tepat di bawahnya
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        let textAfterCursor = '';
        const sel = window.getSelection();
        if (sel && sel.rangeCount) {
          const range = sel.getRangeAt(0);
          const postRange = range.cloneRange();
          postRange.selectNodeContents(paraEl);
          postRange.setStart(range.endContainer, range.endOffset);
          textAfterCursor = postRange.toString().trim();
          if (textAfterCursor) {
            postRange.deleteContents();
            paraEl.innerText = paraEl.innerText.trimEnd();
            syncParaEmptyAttr(paraEl);
          }
        }
        insertParagraphBlock(wrapper, 'after', textAfterCursor);
        return;
      }
    }

    // Interaksi Keyboard di dalam Baris Tabel Identitas / Jadwal
    const tableCell = e.target.closest('.identity-table-word td[contenteditable="true"]');
    if (tableCell) {
      const tr = tableCell.closest('tr');
      const tbody = tr?.closest('tbody');
      if (tr && tbody) {
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          const temp = document.createElement('tbody');
          temp.innerHTML = buildIdentityRowHtml({ label: 'Label Baru', value: ': ...' }).trim();
          const newRow = temp.firstElementChild;
          tr.insertAdjacentElement('afterend', newRow);
          focusAtEdge(newRow.querySelector('.id-label'), true);
          scheduleSync(container);
          reflowWordCanvasPagination(container);
          return;
        }
        if ((e.key === 'Backspace' || e.key === 'Delete')) {
          const lblText = tr.querySelector('.id-label')?.innerText.replace(/[\s\u00a0]/g, '') || '';
          const valText = tr.querySelector('.id-val')?.innerText.replace(/[:\s\u00a0]/g, '') || '';
          if (!lblText && !valText) {
            e.preventDefault();
            if (tbody.querySelectorAll('tr').length > 1) {
              const prevRow = tr.previousElementSibling || tr.nextElementSibling;
              tr.remove();
              if (prevRow) focusAtEdge(prevRow.querySelector('.id-val'), true);
              scheduleSync(container);
              scheduleWordPagination(container);
            } else {
              removeBlockWrapper(tr.closest('.sheet-identity-wrapper'), 'prev');
            }
            return;
          }
        }
      }
    }
  });

  // Event Delegation: Klik Tombol Floating pada Blok & Tabel lintas halaman
  sheetWrapper.addEventListener('click', (e) => {
    if (!e.target.closest('.sheet-dynamic-blocks')) return;
    const blockBtn = e.target.closest('[data-block-action]');
    if (blockBtn) {
      e.preventDefault();
      const action = blockBtn.dataset.blockAction;
      const wrapper = blockBtn.closest('.canvas-block-item');
      if (!wrapper) return;

      if (action === 'delete-block') {
        removeBlockWrapper(wrapper, 'prev');
        showToast('Area dihapus dari lembar kerja.', 'info');
      } else if (action === 'add-para-above') {
        insertParagraphBlock(wrapper, 'before', '');
      } else if (action === 'add-para-below') {
        insertParagraphBlock(wrapper, 'after', '');
      } else if (action === 'add-table-below') {
        insertTableBlock(wrapper, 'after');
      } else if (action === 'add-table-row') {
        const tbody = wrapper.querySelector('.identity-table-word tbody');
        if (tbody) {
          const temp = document.createElement('tbody');
          temp.innerHTML = buildIdentityRowHtml({ label: 'Rincian Baru', value: ': ...........................' }).trim();
          const newRow = temp.firstElementChild;
          tbody.appendChild(newRow);
          focusAtEdge(newRow.querySelector('.id-label'), true);
          scheduleSync(container);
          reflowWordCanvasPagination(container);
        }
      }
      return;
    }

    const gridBtn = e.target.closest('[data-grid-act]');
    if (gridBtn) {
      e.preventDefault();
      const gAct = gridBtn.dataset.gridAct;
      const wrap = gridBtn.closest('.sheet-custom-table-wrapper');
      const tbl = wrap?.querySelector('table.word-custom-grid-table');
      const tbody = tbl?.querySelector('tbody') || tbl;
      if (!wrap || !tbl || !tbody) return;

      const allTrs = [...tbody.querySelectorAll('tr')];
      const maxCols = Math.max(1, ...(allTrs.map(tr =>
        [...tr.children].reduce((sum, el) => sum + (parseInt(el.getAttribute('colspan') || '1', 10) || 1), 0)
      )));

      if (gAct === 'add-row') {
        const tr = document.createElement('tr');
        for (let c = 0; c < maxCols; c++) {
          const td = document.createElement('td');
          td.contentEditable = 'true';
          td.spellcheck = false;
          td.innerText = c === 0 ? String(allTrs.length) : '...';
          tr.appendChild(td);
        }
        tbody.appendChild(tr);
        focusAtEdge(tr.querySelector('td'), true);
      } else if (gAct === 'del-row') {
        if (allTrs.length > 1) {
          allTrs[allTrs.length - 1].remove();
        } else {
          removeBlockWrapper(wrap, 'prev');
          return;
        }
      } else if (gAct === 'add-col') {
        allTrs.forEach((tr, rIdx) => {
          const isHead = tr.dataset.header === 'true' || Boolean(tr.querySelector('th'));
          const cell = document.createElement(isHead ? 'th' : 'td');
          cell.contentEditable = 'true';
          cell.spellcheck = false;
          cell.innerText = isHead ? `Kolom ${maxCols + 1}` : '...';
          tr.appendChild(cell);
        });
      } else if (gAct === 'del-col') {
        if (maxCols <= 1) {
          removeBlockWrapper(wrap, 'prev');
          return;
        }
        allTrs.forEach(tr => {
          const last = tr.lastElementChild;
          if (last) last.remove();
        });
      } else if (gAct === 'cycle-border') {
        const cur = tbl.dataset.borderStyle || 'grid';
        const next = cur === 'grid' ? 'sppd' : (cur === 'sppd' ? 'none' : 'grid');
        tbl.dataset.borderStyle = next;
        tbl.classList.remove('grid-border-sppd', 'grid-border-none');
        if (next === 'sppd') tbl.classList.add('grid-border-sppd');
        if (next === 'none') tbl.classList.add('grid-border-none');
        showToast(`Gaya garis tabel: ${next === 'grid' ? 'Garis Penuh' : next === 'sppd' ? 'SPPD (Tanpa Garis Samping)' : 'Tanpa Garis'}`, 'info');
      } else if (gAct === 'toggle-header-bg') {
        const noBg = tbl.classList.toggle('grid-no-header-bg');
        tbl.dataset.headerShading = noBg ? 'false' : 'true';
      }
      scheduleSync(container);
      updateWordCount(container);
      reflowWordCanvasPagination(container);
      return;
    }

    const rowBtn = e.target.closest('[data-row-action]');
    if (rowBtn) {
      e.preventDefault();
      const action = rowBtn.dataset.rowAction;
      const tr = rowBtn.closest('tr');
      const tbody = tr?.closest('tbody');
      if (!tr || !tbody) return;

      if (action === 'add-row-below') {
        const temp = document.createElement('tbody');
        temp.innerHTML = buildIdentityRowHtml({ label: 'Rincian', value: ': ...........................' }).trim();
        const newRow = temp.firstElementChild;
        tr.insertAdjacentElement('afterend', newRow);
        focusAtEdge(newRow.querySelector('.id-label'), true);
        scheduleSync(container);
        reflowWordCanvasPagination(container);
      } else if (action === 'delete-row') {
        if (tbody.querySelectorAll('tr').length > 1) {
          tr.remove();
          scheduleSync(container);
          scheduleWordPagination(container);
        } else {
          removeBlockWrapper(tr.closest('.sheet-identity-wrapper'), 'prev');
        }
      }
    }
  });

  // Tombol Bar Tambah Area Teks / Tabel / Judul / Sifat / TTD di Bawah Konten
  container.querySelectorAll('[data-canvas-insert]').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      const type = btn.dataset.canvasInsert;
      if (type === 'table') {
        insertTableBlock(null, 'after');
        showToast('Tabel rincian baru ditambahkan.', 'info');
      } else if (type === 'grid-table') {
        insertGridTableBlock(null, 'after');
        showToast('Tabel grid multi-kolom berhasil ditambahkan.', 'info');
      } else if (type === 'pagebreak') {
        insertPageBreakBlock(null, 'after');
        showToast('Pemisah halaman baru (Page Break) ditambahkan.', 'info');
      } else if (type === 'title') {
        const titleSec = container.querySelector('.sheet-section-title');
        const titleEl = container.querySelector('#editable-doc-title');
        const nomorEl = container.querySelector('#editable-doc-nomor');
        if (titleSec) titleSec.classList.remove('hidden');
        if (titleEl && !titleEl.innerText.trim()) titleEl.innerText = 'SURAT DINAS';
        if (nomorEl && !nomorEl.innerText.trim()) {
          const nextNomor = store.getNextFormattedOutgoingNomor({ title: titleEl?.innerText.trim() || 'SURAT DINAS' }).nomorSurat;
          nomorEl.innerText = `Nomor: ${nextNomor}`;
        }
        container.querySelector('[data-canvas-insert="title"]')?.classList.add('hidden');
        if (titleEl) focusAtEdge(titleEl, true);
        scheduleSync(container);
        updateWordCount(container);
        scheduleWordPagination(container);
      } else if (type === 'meta') {
        const metaWrap = container.querySelector('#sheet-meta-recipient-wrapper');
        if (metaWrap) metaWrap.classList.remove('hidden');
        container.querySelector('#btn-restore-meta-recipient')?.click();
        container.querySelector('[data-canvas-insert="meta"]')?.classList.add('hidden');
      } else if (type === 'sig') {
        const sigWrap = container.querySelector('.sheet-signature-wrapper');
        const school = store.getSchoolProfile();
        const sigDate = container.querySelector('#editable-sig-date');
        const sigTitle = container.querySelector('#editable-sig-title');
        const sigName = container.querySelector('#editable-sig-name');
        const sigNip = container.querySelector('#editable-sig-nip');
        if (sigWrap) sigWrap.classList.remove('hidden');
        if (sigDate && !sigDate.innerText.trim()) {
          const todayStr = new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
          sigDate.innerText = `${school.city || 'Garut'}, ${todayStr}`;
        }
        if (sigTitle && !sigTitle.innerText.trim()) sigTitle.innerText = 'Kepala Sekolah,';
        if (sigName && !sigName.innerText.trim()) sigName.innerText = school.principalName || 'Dian Nurhadian, S.Pd.';
        if (sigNip && !sigNip.innerText.trim()) {
          sigNip.innerText = formatSignerNipLine(school.principalNip);
        }
        container.querySelector('[data-canvas-insert="sig"]')?.classList.add('hidden');
        scheduleSync(container);
        updateWordCount(container);
        scheduleWordPagination(container);
      } else {
        insertParagraphBlock(null, 'after', '');
      }
    });
  });

  // Hubungkan Move Handle (panah 4-arah pojok kiri atas), Resize Cursor (col/row), & Drag Boundary MS Word
  const wordDeskEl = container.querySelector('#word-desk-container');
  attachInteractiveTables({
    deskEl: wordDeskEl,
    tableSelector: '.sheet-dynamic-blocks table',
    excludeTableSelector: '.meta-perihal-table',
    getMoveUnitEl: (tbl) => tbl.closest('.canvas-block-item') || tbl,
    getDropTargets: (draggedUnit) =>
      [...sheetWrapper.querySelectorAll('.sheet-dynamic-blocks > .canvas-block-item')].filter(
        el => el !== draggedUnit
      ),
    performMove: (unitEl, targetEl, position) => {
      unitEl.removeAttribute('data-split-group');
      targetEl.insertAdjacentElement(position === 'before' ? 'beforebegin' : 'afterend', unitEl);
    },
    onSelectWholeTable: () => {
      showToast('Seluruh isi tabel dipilih — klik & seret ikon panah 4-arah untuk memindahkan posisi tabel, atau Delete untuk menghapus.', 'info');
    },
    onMutate: (kind) => {
      if (kind === 'delete-table' && getAllBlockNodes().length === 0) {
        insertParagraphBlock(null, 'after', '');
      }
      scheduleSync(container);
      updateWordCount(container);
      reflowWordCanvasPagination(container);
    }
  });
}

function updateWordCount(container) {
  const allBodies = [...container.querySelectorAll('.sheet-content-editable')];
  const bodyText = allBodies.map(el => el.innerText || '').join(' ');
  const words = bodyText.trim().split(/\s+/).filter(w => w.length > 0).length;
  const countEl = container.querySelector('#status-word-count');
  if (countEl) countEl.textContent = `${words} kata`;
}

function scheduleSync(container) {
  scheduleWordPagination(container);
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
 * Menyusun Document Model dari isi lembar kerja saat ini (membaca urutan blok secara dinamis & mengabaikan kotak kosong)
 */
function buildModelFromCanvas(container) {
  const state = store.getState();
  const paperSize = state.letterOptions?.paperSize === 'f4' ? 'suratDinasF4' : 'suratDinas';
  const margins = state.letterOptions?.margins || { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 };
  const showKop = state.letterOptions?.showKop !== false;

  const titleSectionEl = container.querySelector('.sheet-section-title');
  const sigWrapperEl = container.querySelector('.sheet-signature-wrapper');
  const rawTitleText = container.querySelector('#editable-doc-title')?.innerText.trim() || '';
  const nomorText = container.querySelector('#editable-doc-nomor')?.innerText.trim() || '';
  const isTitleVisible = Boolean(titleSectionEl && !titleSectionEl.classList.contains('hidden') && (rawTitleText || nomorText));
  const titleText = rawTitleText || 'Dokumen Baru';

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

  // Parse signature (hanya dianggap ada bila blok tanda tangan aktif & memiliki isi)
  const sigDateStr = container.querySelector('#editable-sig-date')?.innerText.trim() || '';
  const sigTitle = container.querySelector('#editable-sig-title')?.innerText.trim().replace(/,$/, '') || '';
  const sigName = container.querySelector('#editable-sig-name')?.innerText.trim() || '';
  const sigNip = container.querySelector('#editable-sig-nip')?.innerText.trim().replace(/^NIP\.?\s*/i, '') || '';
  const isSigVisible = Boolean(sigWrapperEl && !sigWrapperEl.classList.contains('hidden') && (sigDateStr || sigTitle || sigName || sigNip));

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
  if (isTitleVisible) {
    blocks.push({
      type: 'letterTitle',
      title: rawTitleText,
      nomor: nomorText,
      metaRecipient: metaRecipientData.isDeleted ? null : metaRecipientData
    });
  } else if (!metaRecipientData.isDeleted) {
    blocks.push({
      type: 'letterMetaRecipient',
      ...metaRecipientData
    });
  }

  // 3. Blok Isi Dinamis lintas seluruh halaman (tabel/paragraf terbelah digabung utuh)
  const dynamicItems = collectMergedDynamicItems(container);
  let bodyBlocksAdded = 0;
  for (const item of dynamicItems) {
    if (item.type === 'paragraph') {
      bodyBlocksAdded += 1;
      blocks.push({
        type: 'paragraph',
        style: 'letterBody',
        ...(item.alignment ? { alignment: item.alignment } : {}),
        text: item.text,
        runs: Array.isArray(item.runs) && item.runs.length > 0 ? item.runs : [{ text: item.text }]
      });
    } else if (item.type === 'identityTable') {
      bodyBlocksAdded += 1;
      blocks.push({
        type: 'identityTable',
        ...(Array.isArray(item.colWidthsPct) && item.colWidthsPct.length >= 2 ? { colWidthsPct: item.colWidthsPct } : {}),
        rows: item.rows
      });
    } else if (item.type === 'table') {
      bodyBlocksAdded += 1;
      blocks.push({
        type: 'table',
        borderStyle: item.borderStyle || 'grid',
        headerRow: item.headerRow !== false,
        headerShading: item.headerShading !== false,
        ...(Array.isArray(item.colWidthsPct) ? { colWidthsPct: item.colWidthsPct } : {}),
        ...(item.tableWidthPct ? { tableWidthPct: item.tableWidthPct } : {}),
        rows: item.rows
      });
    } else if (item.type === 'pageBreak') {
      bodyBlocksAdded += 1;
      blocks.push({ type: 'pageBreak' });
    }
  }

  // 4. Signature Block (bila aktif)
  if (isSigVisible) {
    blocks.push({
      type: 'letterSignature',
      city: sigDateStr.split(',')[0]?.trim() || state.schoolProfile?.city || 'Garut',
      date: sigDateStr.split(',')[1]?.trim() || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
      title: sigTitle || 'Kepala Sekolah',
      name: sigName || state.schoolProfile?.principalName || 'Dian Nurhadian, S.Pd.',
      nip: sigNip
    });
  }

  const hasAnyContent = Boolean(isTitleVisible || !metaRecipientData.isDeleted || isSigVisible || bodyBlocksAdded > 0);
  if (bodyBlocksAdded === 0 && !isTitleVisible && !isSigVisible) {
    blocks.push({
      type: 'paragraph',
      style: 'letterBody',
      text: '',
      runs: [{ text: '' }]
    });
  }

  return {
    meta: {
      title: rawTitleText || titleText,
      author: sigName || state.schoolProfile?.principalName || 'SMK IT Al-Hawari',
      year: new Date().getFullYear().toString()
    },
    profile: paperSize,
    customMargins: twipMargins,
    isBlankDraft: !hasAnyContent,
    ...(state.document?.verification ? { verification: state.document.verification } : {}),
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
 * Memasukkan hasil AI Copilot langsung ke lembar kerja Word secara dinamis (tanpa meninggalkan kotak kosong)
 */
function applyAiResultToCanvas(container, aiResult) {
  if (aiResult?.templatePreset === 'surat_tugas_sppd') {
    const stNow = store.getState();
    const optsNow = store.getLetterOptions();
    const sppdDoc = buildSuratTugasSppdDocModel(
      aiResult.formData || aiResult,
      stNow.schoolProfile,
      { ...optsNow, paperSize: aiResult.paperSize || 'f4' }
    );
    store.setState({
      smartDocument: {
        pages: [],
        pendingDocModel: sppdDoc,
        meta: sppdDoc.meta || { title: 'SURAT TUGAS & SPPD' }
      }
    }, false);
    showToast('Paket Surat Tugas & SPPD (3 Halaman 1:1 Word) dibuka otomatis di Editor Word Pro!', 'success');
    document.querySelector('.nav-tab[data-tab="smart-canvas"]')?.click();
    return;
  }

  const titleSectionEl = container.querySelector('.sheet-section-title');
  const sigWrapperEl = container.querySelector('.sheet-signature-wrapper');
  const titleEl = container.querySelector('#editable-doc-title');
  const nomorEl = container.querySelector('#editable-doc-nomor');
  const letterTitle = aiResult.letterTitle || 'SURAT DINAS';
  const smartNomor = formatSmartNomorSurat({
    rawNomor: aiResult.nomorSurat || '',
    contextText: `${letterTitle} ${aiResult.hal || ''} ${aiResult.perihal || ''} ${aiResult.title || ''}`,
    forcedTypeCode: aiResult.kodeJenisSurat || '',
    sequenceNumber: store.getNextOutgoingSequence()
  });

  if (titleSectionEl) titleSectionEl.classList.remove('hidden');
  if (titleEl) titleEl.innerText = letterTitle;
  if (nomorEl) nomorEl.innerText = `Nomor: ${smartNomor}`;

  // Sifat, Perihal & Kepada Yth
  const metaEl = container.querySelector('#editable-meta-recipient');
  const metaWrapper = container.querySelector('#sheet-meta-recipient-wrapper');
  if (metaEl) {
    const btnClear = container.querySelector('#btn-clear-meta-recipient');
    const btnRestore = container.querySelector('#btn-restore-meta-recipient');
    if (aiResult.hideMetaRecipient === true) {
      metaEl.innerHTML = '';
      if (metaWrapper) {
        metaWrapper.classList.add('is-empty', 'hidden');
      }
      if (btnClear) btnClear.classList.add('hidden');
      if (btnRestore) btnRestore.classList.remove('hidden');
    } else {
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
      if (metaWrapper) {
        metaWrapper.classList.remove('is-empty');
        metaWrapper.classList.remove('hidden');
      }
      if (btnClear) btnClear.classList.remove('hidden');
      if (btnRestore) btnRestore.classList.add('hidden');
    }
  }

  // Render blok isi (paragraf, tabel rincian, tabel grid, & page break) sesuai urutan asli dari AI
  collapseToSingleWordPage(container);
  const dynamicContainer = container.querySelector('#sheet-dynamic-blocks');
  if (dynamicContainer && Array.isArray(aiResult.blocks)) {
    const htmlParts = [];
    for (const b of aiResult.blocks) {
      if (!b || !b.type) continue;
      if (b.type === 'pageBreak') {
        htmlParts.push(buildPageBreakBlockHtml());
      } else if (b.type === 'heading' && b.text) {
        htmlParts.push(buildParagraphBlockHtml({
          text: String(b.text).trim(),
          alignment: b.alignment || 'left',
          isBold: true
        }));
      } else if (b.type === 'list' && Array.isArray(b.items)) {
        b.items.forEach((it, idx) => {
          htmlParts.push(buildParagraphBlockHtml({
            text: `${b.ordered !== false ? `${idx + 1}. ` : '• '}${String(it)}`,
            alignment: 'left'
          }));
        });
      } else if (b.type === 'paragraph') {
        const text = (b.text || '').trim();
        if (text) {
          htmlParts.push(buildParagraphBlockHtml({
            text,
            alignment: b.alignment || 'justify',
            isBold: Boolean(b.bold)
          }));
        }
      } else if (b.type === 'identityTable' && Array.isArray(b.rows) && b.rows.length > 0) {
        htmlParts.push(buildIdentityTableBlockHtml(b.rows));
      } else if (b.type === 'table' && Array.isArray(b.rows) && b.rows.length > 0) {
        htmlParts.push(buildCustomGridTableBlockHtml(b));
      }
    }
    if (htmlParts.length > 0) {
      dynamicContainer.innerHTML = htmlParts.join('');
    }
  }

  // Update date & pastikan blok TTD terlihat serta terisi dari profil sekolah bila kanvas sebelumnya kosong
  const schoolSig = store.getState().schoolProfile || {};
  const sigDateEl = container.querySelector('#editable-sig-date');
  const sigTitleEl = container.querySelector('#editable-sig-title');
  const sigNameEl = container.querySelector('#editable-sig-name');
  const sigNipEl = container.querySelector('#editable-sig-nip');
  if (sigWrapperEl) sigWrapperEl.classList.remove('hidden');
  const dateText = aiResult.tanggalSurat || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
  const cityText = schoolSig.city || 'Garut';
  if (sigDateEl && (aiResult.tanggalSurat || !sigDateEl.innerText.trim())) {
    sigDateEl.innerText = `${cityText}, ${dateText}`;
  }
  if (sigTitleEl && !sigTitleEl.innerText.trim()) sigTitleEl.innerText = 'Kepala Sekolah,';
  if (sigNameEl && !sigNameEl.innerText.trim()) sigNameEl.innerText = schoolSig.principalName || 'Dian Nurhadian, S.Pd.';
  if (sigNipEl && !sigNipEl.innerText.trim()) {
    sigNipEl.innerText = formatSignerNipLine(schoolSig.principalNip);
  }

  container.querySelector('[data-canvas-insert="title"]')?.classList.add('hidden');
  container.querySelector('[data-canvas-insert="meta"]')?.classList.add('hidden');
  container.querySelector('[data-canvas-insert="sig"]')?.classList.add('hidden');

  scheduleSync(container);
  updateWordCount(container);
  reflowWordCanvasPagination(container);
}

/**
 * Mengekstrak snapshot isi lembar kerja saat ini sesuai urutan blok dinamis di kanvas
 */
function extractCanvasSnapshot(container) {
  const titleText = container.querySelector('#editable-doc-title')?.innerText.trim() || 'SURAT DINAS';
  const nomorText = container.querySelector('#editable-doc-nomor')?.innerText.trim() || '';

  const metaEl = container.querySelector('#editable-meta-recipient');
  const metaWrapper = container.querySelector('#sheet-meta-recipient-wrapper');
  const isMetaDeleted = !metaEl || metaWrapper?.classList.contains('is-empty') || !metaEl.innerText.trim();

  let sifat = 'Penting';
  let hal = 'Undangan Rapat';
  let tujuan = 'Yth. ...............';

  if (!isMetaDeleted && metaEl) {
    metaEl.querySelectorAll('.meta-perihal-table tr').forEach(tr => {
      const lbl = tr.querySelector('.meta-label')?.innerText.trim().toLowerCase() || '';
      const val = tr.querySelector('.meta-val')?.innerText.trim() || '';
      if (lbl.includes('sifat')) sifat = val;
      if (lbl.includes('perihal') || lbl.includes('hal')) hal = val;
    });
    const addrEl = metaEl.querySelector('.recipient-address-block');
    if (addrEl) {
      const lines = addrEl.innerText.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
      const ythLine = lines.find(l => /^yth/i.test(l));
      if (ythLine) tujuan = ythLine;
    }
  }

  const blocks = collectMergedDynamicItems(container);

  const school = store.getSchoolProfile();
  const todayId = new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
  const sigDateStr = container.querySelector('#editable-sig-date')?.innerText.trim() || `${school.city || 'Garut'}, ${todayId}`;
  const sigTitle = container.querySelector('#editable-sig-title')?.innerText.trim().replace(/,$/, '') || 'Kepala Sekolah';
  const sigName = container.querySelector('#editable-sig-name')?.innerText.trim() || school.principalName || 'Dian Nurhadian, S.Pd.';
  const sigNip = container.querySelector('#editable-sig-nip')?.innerText.trim() || (school.principalNip || '-');

  return {
    letterTitle: titleText,
    nomorSurat: nomorText.replace(/^Nomor:\s*/i, ''),
    hideMetaRecipient: isMetaDeleted,
    sifat,
    hal,
    tujuan,
    blocks,
    signer: {
      city: sigDateStr.split(',')[0]?.trim() || 'Garut',
      date: sigDateStr.split(',')[1]?.trim() || sigDateStr,
      title: sigTitle,
      name: sigName,
      nip: sigNip
    }
  };
}

/**
 * Menerapkan hasil pengeditan AI langsung ke lembar kerja dengan mempertahankan urutan blok dinamis
 */
function applyEditedResultToCanvas(container, edited, prevSnapshot) {
  const state = store.getState();
  const school = store.getSchoolProfile();
  const opts = store.getLetterOptions();

  const finalTitle = (edited.letterTitle || prevSnapshot.letterTitle || 'SURAT DINAS').trim();
  const rawNomor = (edited.nomorSurat || prevSnapshot.nomorSurat || '').trim();
  const nextSeq = store.getNextOutgoingSequence();
  const smartNomor = edited.kodeJenisSurat
    ? formatSmartNomorSurat({ rawNomor, forcedTypeCode: edited.kodeJenisSurat, sequenceNumber: nextSeq })
    : (rawNomor.replace(/^Nomor:\s*/i, '') || formatSmartNomorSurat({ rawNomor, contextText: finalTitle, sequenceNumber: nextSeq }));

  const hideMeta = edited.hideMetaRecipient !== undefined ? Boolean(edited.hideMetaRecipient) : Boolean(prevSnapshot.hideMetaRecipient);
  const sifatVal = edited.sifat || prevSnapshot.sifat || 'Penting';
  const halVal = edited.hal || edited.perihal || prevSnapshot.hal || 'Undangan Rapat';
  const tujuanRaw = edited.tujuan || prevSnapshot.tujuan || 'Yth. ...............';
  const tujuanFormatted = tujuanRaw.toLowerCase().startsWith('yth') ? tujuanRaw : `Yth. ${tujuanRaw}`;

  const metaRecipientObj = hideMeta ? { isDeleted: true, metaRows: [], recipientLines: [] } : {
    isDeleted: false,
    metaRows: [
      { label: 'Sifat', value: sifatVal },
      { label: 'Perihal', value: halVal }
    ],
    recipientLines: [
      'Kepada',
      tujuanFormatted,
      '__________________',
      'Di',
      'Tempat'
    ]
  };

  const newBlocks = [];
  if (opts.showKop !== false) {
    newBlocks.push({
      type: 'letterKop',
      school,
      image: state.schoolProfile?.kopImage || DEFAULT_KOP_DATA_URL
    });
  }

  newBlocks.push({
    type: 'letterTitle',
    title: finalTitle.toUpperCase(),
    nomor: `Nomor: ${smartNomor}`,
    metaRecipient: hideMeta ? null : metaRecipientObj
  });

  // Pertahankan urutan asli setiap blok (paragraf, tabel rincian, tabel grid/SPPD, list, & pageBreak)
  const sourceBlocks = Array.isArray(edited.blocks) && edited.blocks.length > 0 ? edited.blocks : prevSnapshot.blocks;
  newBlocks.push(...convertAiBlocksToDocumentBlocks(sourceBlocks));

  const sig = edited.signer || prevSnapshot.signer || {};
  newBlocks.push({
    type: 'letterSignature',
    city: sig.city || 'Garut',
    date: sig.date || '14 November 2026',
    title: sig.title || 'Kepala Sekolah',
    name: sig.name || school.principalName || 'Dian Nurhadian, S.Pd.',
    nip: sig.nip || '-'
  });

  const updatedDocModel = {
    meta: {
      title: finalTitle,
      author: sig.name || school.principalName || 'SMK IT Al-Hawari',
      year: new Date().getFullYear().toString()
    },
    profile: opts.paperSize === 'f4' ? 'suratDinasF4' : 'suratDinas',
    paperSize: opts.paperSize || 'a4',
    orientation: opts.orientation || 'portrait',
    customMargins: opts.margins,
    ...(state.document?.verification ? { verification: state.document.verification } : {}),
    ...(state.document?.qrPlacement ? { qrPlacement: state.document.qrPlacement } : {}),
    letterData: {
      kopImage: state.schoolProfile?.kopImage || DEFAULT_KOP_DATA_URL,
      metaRecipientDeleted: hideMeta
    },
    blocks: newBlocks
  };

  store.clearActiveUploadedDocx();
  store.setState({ document: updatedDocModel, docxViewMode: 'canvas' }, true);
  renderWordCanvas(container);
}

function bindRecordOutgoingModal(container) {
  const modal = container.querySelector('#modal-record-outgoing');
  if (!modal) return;

  const triggerButtons = [
    container.querySelector('#btn-record-outgoing'),
    container.querySelector('#btn-quick-record-outgoing'),
    container.querySelector('#btn-banner-record-outgoing'),
    container.querySelector('#btn-post-print-record')
  ].filter(Boolean);

  const postPrintBar = container.querySelector('#post-print-outgoing-bar');
  const btnClosePostPrint = container.querySelector('#btn-close-post-print');
  if (btnClosePostPrint && postPrintBar) {
    btnClosePostPrint.addEventListener('click', () => postPrintBar.classList.add('hidden'));
  }

  const btnClose = container.querySelector('#btn-close-outgoing-modal');
  const btnCancel = container.querySelector('#btn-cancel-record-outgoing');
  const btnConfirm = container.querySelector('#btn-confirm-record-outgoing');
  const seqBadge = container.querySelector('#out-seq-badge');
  const inputNomor = container.querySelector('#input-out-nomor');
  const selectType = container.querySelector('#select-out-type-code');
  const inputTitle = container.querySelector('#input-out-title');
  const inputRecipient = container.querySelector('#input-out-recipient');
  const inputPerihal = container.querySelector('#input-out-perihal');

  const closeModal = (resetClearFlag = true) => {
    modal.classList.add('hidden');
    if (resetClearFlag) {
      container.__clearCanvasAfterOutgoing = false;
    }
  };
  if (btnClose) btnClose.addEventListener('click', () => closeModal(true));
  if (btnCancel) btnCancel.addEventListener('click', () => closeModal(true));

  // Klik di luar area kartu modal untuk menutup
  modal.addEventListener('click', (e) => {
    if (e.target === modal) closeModal(true);
  });

  const getActiveSnapshotOrDomModel = () => {
    const state = store.getState();
    const hasDomCanvas = Boolean(container.querySelector('#editable-doc-title') || container.querySelector('#sheet-dynamic-blocks'));
    if (hasDomCanvas) {
      return buildModelFromCanvas(container);
    }
    // Jika sedang di mode "Asli Word" (native_preview), ambil dari model state / activeUploadedDocx
    if (state.document && Array.isArray(state.document.blocks)) {
      return JSON.parse(JSON.stringify(state.document));
    }
    return buildModelFromCanvas(container);
  };

  const extractBestTitle = () => {
    const domTitle = container.querySelector('#editable-doc-title')?.innerText.trim();
    if (domTitle) return domTitle;
    const state = store.getState();
    const titleBlock = state.document?.blocks?.find(b => b.type === 'letterTitle');
    return titleBlock?.title || state.activeUploadedDocx?.parsedMeta?.title || 'SURAT DINAS';
  };

  const extractBestNomorRaw = () => {
    const domNomor = container.querySelector('#editable-doc-nomor')?.innerText.trim();
    if (domNomor) return domNomor;
    const state = store.getState();
    const titleBlock = state.document?.blocks?.find(b => b.type === 'letterTitle');
    return titleBlock?.nomor || state.activeUploadedDocx?.parsedMeta?.nomor || '';
  };

  const extractBestRecipient = () => {
    // 1. Cek baris tabel identitas di DOM yang mengandung "Nama"
    let foundName = '';
    container.querySelectorAll('.identity-table-word tr').forEach(tr => {
      if (foundName) return;
      const lbl = tr.querySelector('.id-label')?.innerText.trim().toLowerCase() || '';
      const val = tr.querySelector('.id-val')?.innerText.replace(/^:\s*/, '').trim() || '';
      if (lbl.includes('nama') && val && !/^\[.*\]$/.test(val) && !/^\.+$/.test(val)) {
        foundName = val;
      }
    });
    if (foundName) return foundName;

    // 2. Cek tujuan di Kepada Yth pada DOM
    const addrEl = container.querySelector('#editable-meta-recipient .recipient-address-block');
    if (addrEl) {
      const yth = addrEl.innerText
        .split(/\r?\n/)
        .map(s => s.trim())
        .find(l => /^yth/i.test(l) && !/\.{3,}/.test(l));
      if (yth) return yth.replace(/^yth\.?\s*/i, '').trim();
    }

    // 3. Fallback ke state model (untuk mode Asli Word)
    const state = store.getState();
    const idTable = state.document?.blocks?.find(b => b.type === 'identityTable');
    if (idTable && Array.isArray(idTable.rows)) {
      const nameRow = idTable.rows.find(r => (r.label || '').toLowerCase().includes('nama'));
      if (nameRow) {
        const cleanVal = String(nameRow.value || '').replace(/^:\s*/, '').trim();
        if (cleanVal && !/^\[.*\]$/.test(cleanVal)) return cleanVal;
      }
    }
    return '';
  };

  const extractBestPerihal = (fallbackTitle) => {
    let foundPerihal = '';
    container.querySelectorAll('#editable-meta-recipient .meta-perihal-table tr').forEach(tr => {
      const lbl = tr.querySelector('.meta-label')?.innerText.trim().toLowerCase() || '';
      const val = tr.querySelector('.meta-val')?.innerText.trim() || '';
      if ((lbl.includes('perihal') || lbl.includes('hal')) && val) {
        foundPerihal = val;
      }
    });
    if (foundPerihal) return foundPerihal;

    const state = store.getState();
    const titleBlock = state.document?.blocks?.find(b => b.type === 'letterTitle');
    const metaPerihal = titleBlock?.metaRecipient?.metaRows?.find(r => /perihal|hal/i.test(r.label || ''));
    return metaPerihal?.value || fallbackTitle || 'Surat Dinas';
  };

  let currentSeqInfo = null;

  const refreshNomorPreview = (forcedCode = '') => {
    const titleVal = inputTitle?.value || '';
    const perihalVal = inputPerihal?.value || '';
    const rawNomor = extractBestNomorRaw();
    currentSeqInfo = store.getNextFormattedOutgoingNomor({
      title: titleVal,
      perihal: perihalVal,
      currentNomor: rawNomor,
      forcedTypeCode: forcedCode
    });
    if (seqBadge) seqBadge.textContent = `#${String(currentSeqInfo.sequenceNumber).padStart(3, '0')}`;
    if (inputNomor) inputNomor.value = currentSeqInfo.nomorSurat;
    if (selectType && !forcedCode) selectType.value = currentSeqInfo.typeCode;
  };

  const openOutgoingModal = () => {
    const titleText = extractBestTitle();
    if (inputTitle) inputTitle.value = titleText;
    if (inputPerihal) inputPerihal.value = extractBestPerihal(titleText);
    if (inputRecipient) inputRecipient.value = extractBestRecipient();

    refreshNomorPreview('');
    modal.classList.remove('hidden');
  };

  container.__openOutgoingModal = openOutgoingModal;

  triggerButtons.forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      container.__clearCanvasAfterOutgoing = false;
      openOutgoingModal();
    });
  });

  if (selectType) {
    selectType.addEventListener('change', () => {
      refreshNomorPreview(selectType.value);
    });
  }

  const buildEntryBase = (forcedCode = '', overrideSignerRole = '') => {
    const title = (inputTitle?.value || '').trim() || extractBestTitle() || 'SURAT DINAS';
    const perihal = (inputPerihal?.value || '').trim() || extractBestPerihal(title) || title;
    const recipientName = (inputRecipient?.value || '').trim() || extractBestRecipient() || 'Umum / Tanpa Nama Penerima';
    const rawNomor = extractBestNomorRaw();
    const seqInfo = currentSeqInfo || store.getNextFormattedOutgoingNomor({
      title,
      perihal,
      currentNomor: rawNomor,
      forcedTypeCode: forcedCode
    });

    const officialNomor = (inputNomor?.value || '').trim() || seqInfo?.nomorSurat || '';
    const nomorEl = container.querySelector('#editable-doc-nomor');
    if (nomorEl && officialNomor) {
      nomorEl.innerText = `Nomor: ${officialNomor}`;
    }

    const activeDocx = store.getState().activeUploadedDocx;
    const school = store.getState().schoolProfile || {};
    const docSnapshot = getActiveSnapshotOrDomModel();
    const curQrPlacement = store.getState().document?.qrPlacement;
    if (curQrPlacement && docSnapshot && !docSnapshot.qrPlacement) {
      docSnapshot.qrPlacement = JSON.parse(JSON.stringify(curQrPlacement));
    }
    const titleBlockInSnap = docSnapshot.blocks?.find(b => b.type === 'letterTitle');
    if (titleBlockInSnap && officialNomor) {
      titleBlockInSnap.nomor = `Nomor: ${officialNomor}`;
    }

    const sigDateEl = container.querySelector('#editable-sig-date');
    const sigTitleEl = container.querySelector('#editable-sig-title');
    const sigNameEl = container.querySelector('#editable-sig-name');
    const sigNipEl = container.querySelector('#editable-sig-nip');
    const sigBlock = docSnapshot.blocks?.find(b => b.type === 'letterSignature');

    if (overrideSignerRole) {
      const signerInfo = buildSignerSelectionInfo(school, overrideSignerRole);
      if (sigBlock) {
        sigBlock.title = signerInfo.signerTitle;
        sigBlock.name = signerInfo.signerName;
        sigBlock.nip = signerInfo.signerNip;
        sigBlock.rank = signerInfo.signerRank;
      }
      if (sigTitleEl) sigTitleEl.innerText = `${signerInfo.signerTitle},`;
      if (sigNameEl) sigNameEl.innerText = signerInfo.signerName;
      if (sigNipEl) sigNipEl.innerText = formatSignerNipLine(signerInfo.signerNip);
    }

    const firstPara = container.querySelector('#sheet-dynamic-blocks .paragraph-box')?.innerText.trim()
      || docSnapshot.blocks?.find(b => b.type === 'paragraph')?.text
      || '';

    return {
      seqInfo,
      entryBase: {
        sequenceNumber: seqInfo?.sequenceNumber,
        year: seqInfo?.year,
        typeCode: forcedCode || selectType?.value || seqInfo?.typeCode,
        nomorSurat: officialNomor,
        title,
        recipientName,
        perihal,
        tanggalSurat: sigDateEl?.innerText.trim()
          || (sigBlock ? `${sigBlock.city || 'Garut'}, ${sigBlock.date || ''}` : '')
          || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
        signerName: sigNameEl?.innerText.trim() || sigBlock?.name || 'Kepala Sekolah',
        summary: firstPara.slice(0, 160),
        sourceMode: 'standard',
        document: docSnapshot,
        uploadedDocx: activeDocx || null
      }
    };
  };

  container.__authorizeStandardLetter = async (signerRole = 'kepala_sekolah', clearAfterSave = false) => {
    currentSeqInfo = null;
    const { entryBase } = buildEntryBase('', signerRole);
    return authorizeAndRecordOutgoing({
      entryBase,
      school: store.getState().schoolProfile,
      signerRole,
      authorizeSah: true,
      syncActiveDocMode: 'standard',
      clearAfterSave
    });
  };

  container.__requestStandardAcc = async (signerRole = 'kepala_sekolah', clearAfterSave = false) => {
    currentSeqInfo = null;
    const { entryBase } = buildEntryBase('', signerRole);
    return requestAccForOutgoing({
      entryBase,
      school: store.getState().schoolProfile,
      requestedSignerRole: signerRole,
      syncActiveDocMode: 'standard',
      clearAfterSave
    });
  };

  if (btnConfirm) {
    const syncConfirmBtnLabel = () => {
      const checked = modal.querySelector('input[name="out-signer-role"]:checked');
      const authorizeSah = modal.querySelector('#chk-out-authorize-sah')?.checked !== false;
      const reqAcc = authorizeSah && checked?.dataset?.requiresAcc === '1';
      btnConfirm.textContent = reqAcc
        ? '📨 Ajukan Minta ACC Kepsek & Catat'
        : authorizeSah
        ? '🛡️ Kunci, Sahkan & Catat'
        : '💾 Catat Surat Keluar';
    };
    modal.querySelectorAll('input[name="out-signer-role"], #chk-out-authorize-sah').forEach(el => {
      el.addEventListener('change', syncConfirmBtnLabel);
    });
    syncConfirmBtnLabel();

    btnConfirm.addEventListener('click', async () => {
      const stNow = store.getState();
      const selectedRadio = modal.querySelector('input[name="out-signer-role"]:checked');
      const selectedRole = selectedRadio?.value
        || resolveDefaultAllowedSignerRole(stNow.authUser, stNow.authPermissions, 'kepala_sekolah');
      const authorizeSah = modal.querySelector('#chk-out-authorize-sah')?.checked !== false;
      const roleAllowed = canSignWithRole(stNow.authUser, stNow.authPermissions, selectedRole).allowed;
      const requiresAcc = authorizeSah && (selectedRadio?.dataset?.requiresAcc === '1' || !roleAllowed);
      const shouldClearAfterSave = Boolean(container.__clearCanvasAfterOutgoing);

      btnConfirm.disabled = true;
      btnConfirm.textContent = requiresAcc
        ? 'Mengirim Antrean ACC Kepala Sekolah...'
        : authorizeSah
        ? 'Menyimpan & Mengesahkan QR...'
        : 'Menyimpan ke Surat Keluar...';

      try {
        const { entryBase } = buildEntryBase('', authorizeSah ? selectedRole : '');
        if (requiresAcc) {
          const { recorded } = await requestAccForOutgoing({
            entryBase,
            school: stNow.schoolProfile,
            requestedSignerRole: selectedRole,
            syncActiveDocMode: 'standard',
            clearAfterSave: shouldClearAfterSave
          });
          closeModal(false);
          container.__clearCanvasAfterOutgoing = false;
          if (postPrintBar) postPrintBar.classList.add('hidden');
          if (shouldClearAfterSave) {
            resetWordCanvasToBlank(container);
          } else {
            renderWordCanvas(container);
          }
          showToast(
            `⏳ Surat #${String(recorded.sequenceNumber).padStart(3, '0')} (${recorded.nomorSurat}) masuk Antrean ACC Kepala Sekolah! Anda bebas melanjutkan membuat surat lain.`,
            'info'
          );
          return;
        }

        const { recorded, cloudResult } = await authorizeAndRecordOutgoing({
          entryBase,
          school: stNow.schoolProfile,
          signerRole: selectedRole,
          authorizeSah,
          syncActiveDocMode: 'standard',
          clearAfterSave: shouldClearAfterSave
        });

        closeModal(false);
        container.__clearCanvasAfterOutgoing = false;
        if (postPrintBar) postPrintBar.classList.add('hidden');

        if (shouldClearAfterSave) {
          resetWordCanvasToBlank(container);
          showToast(
            `✓ Surat #${String(recorded.sequenceNumber).padStart(3, '0')} (${recorded.nomorSurat}) atas nama "${recorded.recipientName}" tersimpan${authorizeSah ? ' & DISAHKAN' : ''} lalu kanvas dikosongkan!`,
            'success'
          );
        } else {
          showToast(
            `✓ Surat #${String(recorded.sequenceNumber).padStart(3, '0')} (${recorded.nomorSurat})${authorizeSah ? (cloudResult.ok ? ' DISAHKAN (QR Cloud Aktif)' : ' DISAHKAN (Lokal)') : ' tersimpan'} di Surat Keluar!`,
            'success'
          );
          renderWordCanvas(container);
        }
      } catch (err) {
        showToast(err?.message || 'Gagal mencatat Surat Keluar.', 'error');
      } finally {
        btnConfirm.disabled = false;
        syncConfirmBtnLabel();
      }
    });
  }
}

/**
 * Helper aktif untuk Cetak / PDF pada Lembar Kerja dengan Modal Pilihan Barcode QR vs TTD Manual Pulpen.
 */
export function triggerWordCanvasPrintWithChoice(container) {
  if (!container) return;
  const doc = store.getState().document;
  const verification = doc?.verification || null;
  const isRevoked = verification?.verificationStatus === 'revoked';
  const isPendingAcc = verification?.verificationStatus === 'pending_acc' || verification?.accStatus === 'pending';
  const isRejectedAcc = verification?.verificationStatus === 'rejected';
  const isVerifiedSah = Boolean(verification?.isVerifiedSah && !isRevoked && verification?.qrDataUrl);
  const titleText = container.querySelector('#editable-doc-title')?.innerText.trim() || doc?.meta?.title || 'SURAT DINAS';
  const nomorText = (container.querySelector('#editable-doc-nomor')?.innerText.trim() || verification?.nomorSurat || '-').replace(/^Nomor:\s*/i, '');

  const runPrint = (mode, qrDataUrl = '', overrideRole = null, overrideVerif = null) => {
    const holder = container.querySelector('.sheet-signature-wrapper .sig-space-holder');
    const sheetEls = [...container.querySelectorAll('.word-paper-sheet')];
    const postPrintBar = container.querySelector('#post-print-outgoing-bar');
    const school = store.getState().schoolProfile || {};
    const v = overrideVerif || verification;
    const effectiveRole = overrideRole || v?.signerRole || 'kepala_sekolah';
    const latestDoc = store.getState().document;
    const printOpts = store.getLetterOptions();
    printWithTemporaryQrOverlay({
      mode,
      pageSheetEls: sheetEls,
      qrPlacement: latestDoc?.qrPlacement || null,
      docModel: latestDoc,
      standardHolderEl: holder,
      qrDataUrl,
      signerRole: effectiveRole,
      principalName: school.principalName || '',
      signerName: v?.signerName || '',
      school,
      paperSize: printOpts.paperSize,
      orientation: printOpts.orientation,
      margins: printOpts.margins,
      onAfterPrint: () => {
        refreshWordFloatingQrOverlay(container);
        if (!isVerifiedSah && mode === 'manual' && postPrintBar) {
          postPrintBar.classList.remove('hidden');
        }
      }
    });
  };

  openPrintSignatureChoiceModal({
    actionType: 'print',
    isVerifiedSah,
    isRevoked,
    isPendingAcc,
    isRejectedAcc,
    rejectedReason: verification?.rejectedReason || '',
    nomorSurat: nomorText,
    title: titleText,
    qrDataUrl: verification?.qrDataUrl || '',
    defaultSignerRole: verification?.signerRole || 'kepala_sekolah',
    signerLabel: verification?.roleLabel || verification?.signerTitle || 'Kepala Sekolah',
    signerName: verification?.signerName || '',
    onChooseQr: () => runPrint('qr', verification.qrDataUrl, verification?.signerRole || 'kepala_sekolah'),
    onChooseManual: () => runPrint('manual'),
    onAuthorizeAndQr: async (selectedRole) => {
      if (typeof container.__authorizeStandardLetter !== 'function') return;
      const result = await container.__authorizeStandardLetter(selectedRole, false);
      runPrint('qr', result.qrDataUrl, selectedRole, result.payload);
    },
    onRequestAcc: async (selectedRole) => {
      if (typeof container.__requestStandardAcc !== 'function') return;
      const res = await container.__requestStandardAcc(selectedRole, false);
      showToast(`⏳ Surat #${String(res.recorded.sequenceNumber).padStart(3, '0')} masuk Antrean ACC Kepala Sekolah!`, 'info');
      renderWordCanvas(container);
    }
  });
}

/**
 * Helper aktif untuk Unduh .docx pada Lembar Kerja dengan Modal Pilihan Barcode QR vs TTD Manual Pulpen.
 */
export function triggerWordCanvasDownloadWithChoice(container) {
  if (!container) return;
  const activeDocx = store.getState().activeUploadedDocx;
  if (activeDocx && activeDocx.rawBytes) {
    (async () => {
      showToast('Menerapkan editan ke berkas Word asli (preservasi format 100% utuh)...', 'info');
      try {
        const patchedBlob = await applyCanvasEditsToDocx(activeDocx.rawBytes, activeDocx.paragraphs || [], { outputType: 'blob' });
        const safeName = (activeDocx.fileName || 'dokumen_teredit.docx').replace(/\.docx$/i, '') + '_edited.docx';
        downloadFile(patchedBlob, safeName);
        showToast('Berkas .docx hasil edit (format asli 100% terjaga) berhasil diunduh.', 'success');
      } catch (err) {
        showToast(`Gagal mem-patch berkas: ${err.message}`, 'error');
      }
    })();
    return;
  }

  const doc = store.getState().document;
  const verification = doc?.verification || null;
  const isRevoked = verification?.verificationStatus === 'revoked';
  const isPendingAcc = verification?.verificationStatus === 'pending_acc' || verification?.accStatus === 'pending';
  const isRejectedAcc = verification?.verificationStatus === 'rejected';
  const isVerifiedSah = Boolean(verification?.isVerifiedSah && !isRevoked && verification?.qrDataUrl);
  const titleText = container.querySelector('#editable-doc-title')?.innerText.trim() || doc?.meta?.title || 'SURAT DINAS';
  const nomorText = (container.querySelector('#editable-doc-nomor')?.innerText.trim() || verification?.nomorSurat || '-').replace(/^Nomor:\s*/i, '');

  const runDownload = async (mode, overrideVerif = null) => {
    const v = overrideVerif || verification;
    const rawModel = buildModelFromCanvas(container);
    const school = store.getState().schoolProfile || {};
    const signerInfo = v ? buildSignerSelectionInfo(school, v.signerRole) : null;
    const finalModel = applySignatureModeToDocModel(rawModel, {
      mode,
      qrDataUrl: mode === 'qr' ? (v?.qrDataUrl || '') : '',
      verificationUrl: mode === 'qr' ? (v?.verificationUrl || '') : '',
      signerRole: v?.signerRole || 'kepala_sekolah',
      signerInfo,
      school,
      syncSignerFields: false
    });
    showToast(`Menyusun berkas Word .docx (${mode === 'qr' ? 'dengan Barcode QR' : 'kosongan TTD manual'})...`, 'info');
    try {
      const blob = await renderDocx(finalModel, 'blob');
      const safeTitle = (finalModel.meta?.title || 'surat_dinas')
        .replace(/[^a-zA-Z0-9_\-]/g, '_')
        .slice(0, 45);
      downloadFile(blob, `${safeTitle}.docx`);
      showToast('Berkas .docx berhasil diunduh.', 'success');
    } catch (err) {
      showToast(`Gagal mengunduh: ${err.message}`, 'error');
    }
  };

  openPrintSignatureChoiceModal({
    actionType: 'docx',
    isVerifiedSah,
    isRevoked,
    isPendingAcc,
    isRejectedAcc,
    rejectedReason: verification?.rejectedReason || '',
    nomorSurat: nomorText,
    title: titleText,
    qrDataUrl: verification?.qrDataUrl || '',
    defaultSignerRole: verification?.signerRole || 'kepala_sekolah',
    signerLabel: verification?.roleLabel || verification?.signerTitle || 'Kepala Sekolah',
    signerName: verification?.signerName || '',
    onChooseQr: () => runDownload('qr'),
    onChooseManual: () => runDownload('manual'),
    onAuthorizeAndQr: async (selectedRole) => {
      if (typeof container.__authorizeStandardLetter !== 'function') return;
      const result = await container.__authorizeStandardLetter(selectedRole, false);
      const roleInfo = buildSignerSelectionInfo(store.getState().schoolProfile, selectedRole);
      await runDownload('qr', {
        qrDataUrl: result.qrDataUrl,
        verificationUrl: result.payload?.verificationUrl,
        signerRole: selectedRole,
        roleLabel: roleInfo.roleLabel
      });
    },
    onRequestAcc: async (selectedRole) => {
      if (typeof container.__requestStandardAcc !== 'function') return;
      const res = await container.__requestStandardAcc(selectedRole, false);
      showToast(`⏳ Surat #${String(res.recorded.sequenceNumber).padStart(3, '0')} masuk Antrean ACC Kepala Sekolah!`, 'info');
      renderWordCanvas(container);
    }
  });
}

function updateWordVaultBadge(container) {
  renderCanvasVaultBadge(
    container?.querySelector('#word-status-vault-badge'),
    store.getActiveVaultEntry('standard')
  );
}

async function buildWordCanvasVaultPayload(container) {
  const activeDocx = store.getState().activeUploadedDocx;
  if (activeDocx) {
    let nextUploaded = activeDocx;
    if (activeDocx.rawBytes) {
      try {
        const patchedBytes = await applyCanvasEditsToDocx(
          activeDocx.rawBytes,
          activeDocx.paragraphs || [],
          { outputType: 'uint8array' }
        );
        nextUploaded = {
          ...activeDocx,
          rawBytes: patchedBytes instanceof Uint8Array ? patchedBytes : new Uint8Array(patchedBytes)
        };
        store.setActiveUploadedDocx(nextUploaded);
      } catch {}
    }
    return {
      isImported: true,
      title: nextUploaded.parsedMeta?.title || nextUploaded.fileName?.replace(/\.docx$/i, '') || 'Dokumen Word Impor',
      nomorSurat: nextUploaded.parsedMeta?.nomor || '-',
      tanggalSurat: nextUploaded.parsedMeta?.date || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
      category: 'Dokumen Impor',
      targetPerson: nextUploaded.parsedMeta?.signer || '-',
      perihal: nextUploaded.parsedMeta?.title || 'Dokumen Word Impor',
      summary: `Dokumen .docx resmi diunggah: "${nextUploaded.fileName}". Preservasi format 100%.`,
      isAiReference: true,
      source: 'upload',
      uploadedDocx: nextUploaded
    };
  }

  const docModel = buildModelFromCanvas(container);
  const titleEl = container.querySelector('#editable-doc-title');
  const nomorEl = container.querySelector('#editable-doc-nomor');
  const sigDateEl = container.querySelector('#editable-sig-date');
  const allParas = collectMergedDynamicItems(container)
    .filter(item => item.type === 'paragraph')
    .map(item => item.text)
    .filter(Boolean);
  const firstTableVal = container.querySelector('.identity-table-word .id-val')?.innerText.replace(/^:\s*/, '').trim() || '-';

  const title = (titleEl?.innerText.trim())
    || (allParas[0]?.slice(0, 60))
    || 'Dokumen Persuratan';
  const nomor = nomorEl?.innerText.trim() || '-';
  const dateText = sigDateEl?.innerText.trim() || new Date().toLocaleDateString('id-ID');
  const summary = (allParas[1] || allParas[0] || 'Dokumen persuratan sekolah.').slice(0, 160);

  return {
    isImported: false,
    title,
    nomorSurat: nomor,
    tanggalSurat: dateText,
    category: 'Surat Dinas',
    targetPerson: firstTableVal,
    perihal: title,
    summary,
    isAiReference: true,
    source: 'created',
    document: docModel
  };
}

/**
 * Menyimpan perubahan (Ctrl+S): jika sedang mengedit file dari Gudang Surat, langsung perbarui file aslinya;
 * jika masih Draf Baru, tampilkan dialog pemberian nama untuk disimpan pertama kali ke Gudang Surat.
 */
async function performWordSave(container, { clearAfterSave = false } = {}) {
  const activeEntry = store.getActiveVaultEntry('standard');
  if (!activeEntry) {
    await openWordSaveVaultModal(container, { mode: 'save', clearAfterSave });
    return;
  }

  const writeUpdateToEntry = async ({ unlock = false } = {}) => {
    const payload = await buildWordCanvasVaultPayload(container);
    const updates = payload.isImported
      ? {
          nomorSurat: payload.nomorSurat || activeEntry.nomorSurat,
          tanggalSurat: payload.tanggalSurat || activeEntry.tanggalSurat,
          targetPerson: payload.targetPerson || activeEntry.targetPerson,
          perihal: payload.perihal || activeEntry.perihal,
          summary: payload.summary || activeEntry.summary,
          uploadedDocx: payload.uploadedDocx,
          ...(unlock ? { isLockedTemplate: false, isMasterTemplate: false } : {})
        }
      : {
          nomorSurat: payload.nomorSurat || activeEntry.nomorSurat,
          tanggalSurat: payload.tanggalSurat || activeEntry.tanggalSurat,
          targetPerson: payload.targetPerson || activeEntry.targetPerson,
          perihal: payload.perihal || activeEntry.perihal,
          summary: payload.summary || activeEntry.summary,
          document: payload.document,
          ...(unlock ? { isLockedTemplate: false, isMasterTemplate: false } : {})
        };

    store.updateLetterVaultEntry(activeEntry.id, updates);
    updateWordVaultBadge(container);
    if (clearAfterSave) {
      resetWordCanvasToBlank(container);
      showToast(`✓ Perubahan disimpan ke "${activeEntry.title}" & kanvas telah dikosongkan!`, 'success');
    } else if (unlock) {
      showToast(`🔓 Kunci dibuka & perubahan disimpan ke "${activeEntry.title}". Anda dapat mengunci Master ⭐ kembali di Gudang Surat.`, 'success');
    } else {
      showToast(`✓ Tersimpan ke "${activeEntry.title}" (Ctrl+S) — file Gudang diperbarui tanpa duplikat.`, 'success');
    }
  };

  if (activeEntry.isLockedTemplate || activeEntry.isMasterTemplate) {
    openLockedMasterSaveChoiceModal(container.querySelector('#word-save-modal-root'), activeEntry, {
      clearAfterSave,
      onSaveAsCopy: () => openWordSaveVaultModal(container, { mode: 'save-as', clearAfterSave }),
      onUnlockAndSave: () => writeUpdateToEntry({ unlock: true })
    });
    return;
  }

  await writeUpdateToEntry();
}

/**
 * Dialog Save (draf baru) atau Save As (duplikat ke file Gudang Surat baru - Ctrl+Shift+S / F12)
 */
async function openWordSaveVaultModal(container, { mode = 'save', clearAfterSave = false } = {}) {
  const modalRoot = container.querySelector('#word-save-modal-root');
  if (!modalRoot) return;
  const isSaveAs = mode === 'save-as';
  const activeEntry = store.getActiveVaultEntry('standard');
  const payload = await buildWordCanvasVaultPayload(container);
  const defaultTitle = isSaveAs
    ? `Salinan - ${activeEntry?.title || payload.title}`
    : payload.title;

  modalRoot.innerHTML = `
    <div class="modal-overlay">
      <div class="modal-box" style="max-width: 480px; width: 92%;">
        <div class="modal-header">
          <h3>${isSaveAs ? ICON_COPY_DOC : ICON_SAVE} ${clearAfterSave ? 'Simpan ke Gudang Sebelum Mengosongkan' : (isSaveAs ? 'Save As / Duplikat ke Gudang Surat' : 'Simpan Dokumen Baru ke Gudang Surat')}</h3>
          <button type="button" class="btn-icon" id="w-save-modal-close">&times;</button>
        </div>
        <div class="modal-body" style="display: grid; gap: 10px;">
          ${isSaveAs && activeEntry ? `<div class="auth-alert" style="background:#eff6ff; border-color:#bfdbfe; color:#1e3a8a; font-size:12px;">Membuat salinan independen baru dari <strong>${escapeHtml(activeEntry.title)}</strong> agar file acuan asli tetap terjaga.</div>` : ''}
          <div>
            <label class="form-label">Judul Arsip</label>
            <input id="w-v-title" class="form-input" value="${escapeHtml(defaultTitle)}" />
            <div id="w-v-dup-warn" class="hidden" style="margin-top:6px; padding:7px 10px; background:#fffbeb; border:1px solid #fde68a; border-radius:6px; font-size:11.5px; color:#92400e;"></div>
          </div>
          <div><label class="form-label">Kategori</label><input id="w-v-cat" class="form-input" value="${escapeHtml(activeEntry?.category || payload.category)}" /></div>
          <div><label class="form-label">Nomor Surat</label><input id="w-v-nomor" class="form-input" value="${escapeHtml(payload.nomorSurat)}" /></div>
          <div><label class="form-label">Perihal</label><input id="w-v-hal" class="form-input" value="${escapeHtml(payload.perihal)}" /></div>
        </div>
        <div class="modal-footer">
          <button type="button" class="btn btn-secondary" id="w-save-modal-cancel">Batal</button>
          <button type="button" class="btn btn-secondary hidden" id="w-save-overwrite-btn">Timpa File Sama</button>
          <button type="button" class="btn btn-primary" id="w-save-modal-ok">${clearAfterSave ? 'Simpan &amp; Kosongkan Kertas' : (isSaveAs ? 'Save As (Simpan Duplikat)' : 'Simpan ke Gudang')}</button>
        </div>
      </div>
    </div>`;

  const closeModal = () => { modalRoot.innerHTML = ''; };
  const titleInput = modalRoot.querySelector('#w-v-title');
  const dupWarnEl = modalRoot.querySelector('#w-v-dup-warn');
  const overwriteBtn = modalRoot.querySelector('#w-save-overwrite-btn');
  let matchedDup = null;

  const checkDup = () => {
    const val = titleInput?.value.trim() || '';
    matchedDup = store.findLetterVaultByTitle(val);
    if (matchedDup) {
      if (matchedDup.isLockedTemplate || matchedDup.isMasterTemplate) {
        dupWarnEl.textContent = `🔒 "${matchedDup.title}" berstatus ⭐ Master Terkunci dan dilindungi agar tidak tertimpa. Ubah judul untuk menyimpan Salinan Baru.`;
        dupWarnEl.classList.remove('hidden');
        overwriteBtn?.classList.add('hidden');
      } else {
        dupWarnEl.textContent = `⚠️ Judul "${matchedDup.title}" sudah ada di Gudang Surat. Ubah judul untuk duplikat terpisah, atau klik "Timpa File Sama".`;
        dupWarnEl.classList.remove('hidden');
        overwriteBtn?.classList.remove('hidden');
      }
    } else {
      dupWarnEl.classList.add('hidden');
      overwriteBtn?.classList.add('hidden');
    }
  };
  titleInput?.addEventListener('input', checkDup);
  checkDup();

  modalRoot.querySelector('#w-save-modal-close').onclick = closeModal;
  modalRoot.querySelector('#w-save-modal-cancel').onclick = closeModal;

  const finalPayload = (titleVal) => ({
    ...payload,
    title: titleVal,
    category: modalRoot.querySelector('#w-v-cat').value.trim() || payload.category,
    nomorSurat: modalRoot.querySelector('#w-v-nomor').value.trim() || '-',
    perihal: modalRoot.querySelector('#w-v-hal').value.trim() || titleVal
  });

  if (overwriteBtn) {
    overwriteBtn.onclick = () => {
      if (!matchedDup) return;
      const titleVal = titleInput.value.trim() || matchedDup.title;
      store.updateLetterVaultEntry(matchedDup.id, finalPayload(titleVal));
      store.setActiveVaultId(matchedDup.id, 'standard');
      closeModal();
      if (clearAfterSave) {
        resetWordCanvasToBlank(container);
        showToast(`✓ "${titleVal}" diperbarui di Gudang Surat & kanvas dikosongkan!`, 'success');
      } else {
        updateWordVaultBadge(container);
        showToast(`✓ File "${titleVal}" di Gudang Surat berhasil diperbarui!`, 'success');
      }
    };
  }

  modalRoot.querySelector('#w-save-modal-ok').onclick = () => {
    const titleVal = titleInput.value.trim() || defaultTitle;
    const created = store.addLetterVaultEntry(finalPayload(titleVal));
    store.setActiveVaultId(created.id, 'standard');
    closeModal();
    if (clearAfterSave) {
      resetWordCanvasToBlank(container);
      showToast(`✓ "${titleVal}" tersimpan di Gudang Surat & kanvas dikosongkan!`, 'success');
    } else {
      updateWordVaultBadge(container);
      showToast(
        isSaveAs
          ? `✓ Save As berhasil! Salinan "${titleVal}" dibuat & kini aktif diedit (Ctrl+S).`
          : `✓ "${titleVal}" disimpan ke Gudang Surat! Edit berikutnya cukup tekan Ctrl+S.`,
        'success'
      );
    }
  };
}

/**
 * Mengecek apakah kanvas Lembar Kerja memiliki isi (judul, sifat/tujuan, paragraf, tabel, TTD, atau berkas .docx aktif)
 */
function hasWordCanvasContent(container) {
  if (store.getState().activeUploadedDocx) return true;

  const titleEl = container.querySelector('#editable-doc-title');
  const nomorEl = container.querySelector('#editable-doc-nomor');
  const titleSection = container.querySelector('.sheet-section-title');
  if (titleSection && !titleSection.classList.contains('hidden')) {
    if ((titleEl?.innerText || '').trim() || (nomorEl?.innerText || '').trim()) return true;
  }

  const metaEl = container.querySelector('#editable-meta-recipient');
  const metaWrap = container.querySelector('#sheet-meta-recipient-wrapper');
  if (metaWrap && !metaWrap.classList.contains('hidden') && !metaWrap.classList.contains('is-empty')) {
    if ((metaEl?.innerText || '').replace(/\u00a0/g, ' ').trim()) return true;
  }

  if (container.querySelector('.sheet-dynamic-blocks .identity-table-word, .sheet-dynamic-blocks .word-custom-grid-table, .sheet-dynamic-blocks .sheet-pagebreak-wrapper')) return true;

  const hasParaText = Array.from(container.querySelectorAll('.sheet-dynamic-blocks .paragraph-box'))
    .some(el => el.innerText.replace(/[\s\u00a0\u200b]/g, '').length > 0);
  if (hasParaText) return true;

  const sigWrap = container.querySelector('.sheet-signature-wrapper');
  if (sigWrap && !sigWrap.classList.contains('hidden')) {
    const sigText = (sigWrap.innerText || '').replace(/\s+/g, '').trim();
    if (sigText.length > 0) return true;
  }

  return false;
}

/**
 * Mengosongkan Lembar Kerja kembali menjadi kertas kosong baru
 */
function resetWordCanvasToBlank(container) {
  clearTimeout(syncTimer);
  container.__clearCanvasAfterOutgoing = false;
  store.clearActiveUploadedDocx();
  store.setActiveVaultId(null, 'standard');
  const emptyDoc = createEmptyDocumentModel(store.getSchoolProfile(), store.getLetterOptions());
  store.setState({ document: emptyDoc, docxViewMode: 'canvas' }, true);
  renderWordCanvas(container);
}

/**
 * Handler tombol Refresh & modal konfirmasi simpan sebelum mengosongkan kanvas
 */
function bindRefreshCanvasModal(container) {
  const btnRefresh = container.querySelector('#btn-refresh-canvas');
  const modal = container.querySelector('#modal-refresh-canvas');
  const btnClose = container.querySelector('#btn-close-refresh-modal');
  const btnCancel = container.querySelector('#btn-cancel-refresh-modal');
  const btnSaveVault = container.querySelector('#btn-refresh-save-vault');
  const btnSaveOutgoing = container.querySelector('#btn-refresh-save-outgoing');
  const btnDiscard = container.querySelector('#btn-refresh-discard');

  if (!btnRefresh) return;

  const closeRefreshModal = () => modal?.classList.add('hidden');
  if (btnClose) btnClose.addEventListener('click', closeRefreshModal);
  if (btnCancel) btnCancel.addEventListener('click', closeRefreshModal);
  if (modal) {
    modal.addEventListener('click', (e) => {
      if (e.target === modal) closeRefreshModal();
    });
  }

  btnRefresh.addEventListener('click', (e) => {
    e.preventDefault();
    if (!hasWordCanvasContent(container)) {
      resetWordCanvasToBlank(container);
      showToast('Lembar kerja sudah dalam keadaan bersih/kosong.', 'info');
      return;
    }
    if (modal) {
      modal.classList.remove('hidden');
    } else {
      resetWordCanvasToBlank(container);
    }
  });

  if (btnSaveVault) {
    btnSaveVault.addEventListener('click', () => {
      closeRefreshModal();
      performWordSave(container, { clearAfterSave: true });
    });
  }

  if (btnSaveOutgoing) {
    btnSaveOutgoing.addEventListener('click', () => {
      closeRefreshModal();
      container.__clearCanvasAfterOutgoing = true;
      if (typeof container.__openOutgoingModal === 'function') {
        container.__openOutgoingModal();
      } else {
        container.querySelector('#btn-record-outgoing')?.click();
      }
    });
  }

  if (btnDiscard) {
    btnDiscard.addEventListener('click', () => {
      closeRefreshModal();
      resetWordCanvasToBlank(container);
      showToast('Kanvas telah dikosongkan ke lembaran baru.', 'info');
    });
  }
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str || '';
  return div.innerHTML;
}
