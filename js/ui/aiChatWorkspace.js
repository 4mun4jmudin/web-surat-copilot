/**
 * aiChatWorkspace.js - Antarmuka Percakapan AI Asisten Profesional (Gaya ChatGPT + Eksekusi Dokumen)
 * Menyediakan dua mode UI yang berbagi riwayat & memori percakapan secara real-time:
 * 1. renderAiChatWorkspace() -> Tab layar penuh "AI Asisten" (Sidebar Riwayat + Thread Chat + Saran Cepat)
 * 2. mountCompactChatDrawer() -> Drawer sidebar di dalam Word Pro & Lembar Kerja
 */

import { store } from '../state.js';
import { showToast } from './notification.js';
import { callAiChatTurn, convertAiBlocksToDocumentBlocks } from '../ai/client.js';
import {
  SCHOOL_UNITS,
  adaptDocumentModelForUnit,
  detectTargetUnitFromText,
  resolveSchoolProfileForUnit,
  resolveDocumentSchoolProfile
} from '../letters/schoolProfile.js';
import {
  CHAT_FOLDERS,
  CHAT_QUICK_TEMPLATES,
  CHAT_HINT_EXAMPLES,
  computeChatMemoryStats,
  filterChatSessions,
  formatChatTimestamp,
  formatTimeOnly
} from '../ai/chatEngine.js';
import {
  shouldAllowLayoutRestructure,
  buildSurgicalIndexFromSmartPages,
  applySurgicalPatchesToSmartPages,
  buildSurgicalIndexFromDocModel,
  applySurgicalPatchesToDocModel
} from '../ai/surgicalEdit.js';
import { aiJsonToSmartDocModel, docModelToSmartPages } from './smartWordCanvas.js';

function esc(str = '') {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getUserInitials() {
  const u = store.getState().authUser;
  const src = String(u?.displayName || u?.username || 'AD').trim();
  const words = src.split(/\s+/).filter(Boolean);
  if (words.length >= 2) {
    return `${words[0][0]}${words[1][0]}`.toUpperCase();
  }
  return src.slice(0, 2).toUpperCase() || 'AD';
}

function formatRichInlineText(raw = '') {
  const escaped = esc(raw);
  return escaped
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\r?\n/g, '<br>');
}

function readFileAsAttachment(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = String(reader.result || '');
      const idx = dataUrl.indexOf(';base64,');
      const base64Data = idx >= 0 ? dataUrl.slice(idx + 8) : '';
      const mimeType = file.type || (idx >= 0 ? dataUrl.slice(5, idx) : 'application/octet-stream');
      resolve({
        id: `att-${Date.now()}-${Math.random().toString(36).slice(2, 5)}`,
        name: file.name || 'lampiran',
        mimeType,
        sizeKb: Math.max(1, Math.round((file.size || 0) / 1024)),
        base64Data
      });
    };
    reader.onerror = () => reject(new Error(`Gagal membaca berkas ${file.name}`));
    reader.readAsDataURL(file);
  });
}

function buildActiveSnapshotFromStore(targetWorkspace = 'smart-canvas') {
  const state = store.getState();
  const school = state.schoolProfile;
  const letterOptions = store.getLetterOptions();

  if (targetWorkspace === 'none') {
    return { snapshot: null, hasContent: false };
  }

  if (targetWorkspace === 'canvas') {
    const doc = state.document;
    const hasContent = Boolean(
      doc && !doc.isBlankDraft && (doc.blocks || []).some(
        b => b && (b.type === 'table' || b.type === 'identityTable' || b.type === 'list' || (b.type === 'paragraph' && String(b.text || b.runs?.map(r => r.text).join('') || '').trim()))
      )
    );
    if (!hasContent) return { snapshot: null, hasContent: false };
    const idx = buildSurgicalIndexFromDocModel(doc, { letterOptions });
    return { snapshot: idx.snapshot, hasContent: true };
  }

  const smart = state.smartDocument;
  const pages = Array.isArray(smart?.pages) ? smart.pages : [];
  const hasSmartContent = pages.some((p) => {
    const plain = String(p?.html || '').replace(/<[^>]+>/g, '').trim();
    return plain.length > 10 || /<table/i.test(p?.html || '');
  });
  if (!hasSmartContent) {
    // Fallback cek apakah dokumen Lembar Kerja punya isi jika Word Pro belum diisi
    const doc = state.document;
    const hasDocContent = Boolean(
      doc && !doc.isBlankDraft && (doc.blocks || []).some(
        b => b && (b.type === 'table' || b.type === 'identityTable' || (b.type === 'paragraph' && String(b.text || b.runs?.map(r => r.text).join('') || '').trim()))
      )
    );
    if (!hasDocContent) return { snapshot: null, hasContent: false };
    const converted = docModelToSmartPages(doc, school);
    const idx = buildSurgicalIndexFromSmartPages(converted.pages, school, { letterOptions });
    return { snapshot: idx.snapshot, hasContent: true, fallbackPages: converted.pages };
  }
  const idx = buildSurgicalIndexFromSmartPages(pages, school, { letterOptions });
  return { snapshot: idx.snapshot, hasContent: true };
}

function computeDryRunChangesForMessage(action, preferredWorkspace = 'smart-canvas') {
  if (!action) return [];
  const state = store.getState();
  const school = state.schoolProfile;
  const opts = store.getLetterOptions();
  const editPlan = action.editPlan;

  if (action.type === 'create_letter' && action.letterDraft) {
    const d = action.letterDraft;
    const items = [];
    if (d.letterTitle || d.title) items.push({ targetLabel: 'Judul Surat', before: '(dokumen baru)', after: d.letterTitle || d.title });
    if (d.kodeJenisSurat || d.nomorSurat) items.push({ targetLabel: 'Kode / Nomor', before: '-', after: `${d.kodeJenisSurat || 'DINAS'} · ${d.nomorSurat || 'Otomatis'}` });
    if (d.hal || d.perihal) items.push({ targetLabel: 'Perihal', before: '-', after: d.hal || d.perihal });
    if (d.tujuan) items.push({ targetLabel: 'Tujuan', before: '-', after: d.tujuan });
    if (Array.isArray(d.blocks)) {
      items.push({ targetLabel: 'Struktur Blok', before: '-', after: `${d.blocks.length} blok isi (paragraf / tabel / daftar)` });
    }
    if (d.pageSetup) {
      const ps = d.pageSetup;
      items.push({
        targetLabel: 'Format Halaman',
        before: '-',
        after: `${String(ps.paperSize || 'a4').toUpperCase()} · ${ps.fontFamily || 'Times New Roman'} ${ps.fontSize || 12}pt`
      });
    }
    return items;
  }

  if (action.type === 'edit_letter' && editPlan) {
    if (Array.isArray(action.changes) && action.changes.length > 0) {
      return action.changes;
    }
    const target = action.targetWorkspace || preferredWorkspace || 'smart-canvas';
    if (target === 'canvas' && state.document) {
      const dry = applySurgicalPatchesToDocModel(state.document, editPlan, {
        allowLayoutRestructure: shouldAllowLayoutRestructure(action.userInstruction || ''),
        letterOptions: opts
      });
      return dry.changes || [];
    }
    const pages = Array.isArray(state.smartDocument?.pages) && state.smartDocument.pages.length > 0
      ? state.smartDocument.pages
      : docModelToSmartPages(state.document || { blocks: [] }, school).pages;
    const dry = applySurgicalPatchesToSmartPages(pages, editPlan, {
      schoolProfile: school,
      allowLayoutRestructure: shouldAllowLayoutRestructure(action.userInstruction || ''),
      letterOptions: opts
    });
    return dry.changes || [];
  }

  return [];
}

function applyActionToStoreWorkspace(action, explicitWorkspace = '') {
  const state = store.getState();
  const baseSchool = state.schoolProfile;
  const detectedUnit = action.targetUnitId
    || detectTargetUnitFromText(action.userInstruction || '', null)
    || detectTargetUnitFromText(`${action.letterDraft?.letterTitle || ''} ${action.letterDraft?.nomorSurat || ''} ${action.letterDraft?.signer?.name || ''}`, null)
    || baseSchool?.unitId
    || 'smk';
  const school = action.schoolProfile || resolveSchoolProfileForUnit(detectedUnit, baseSchool, state.kopGallery || []);
  const prevOpts = store.getLetterOptions();
  const targetWs = explicitWorkspace || action.targetWorkspace || 'smart-canvas';

  if (action.type === 'create_letter' && action.letterDraft) {
    const docModel = aiJsonToSmartDocModel(action.letterDraft, school, prevOpts);
    docModel.targetUnitId = school.unitId;
    const prevDocBackup = state.document ? JSON.parse(JSON.stringify(state.document)) : null;
    const prevSmartBackup = state.smartDocument ? JSON.parse(JSON.stringify(state.smartDocument)) : null;

    if (docModel.extractedPageSetup || docModel.pageSetup) {
      store.updateLetterOptions(docModel.extractedPageSetup || docModel.pageSetup);
    }

    if (targetWs === 'canvas') {
      store.clearActiveUploadedDocx();
      store.setState({ document: docModel, docxViewMode: 'canvas' }, true);
    } else {
      const converted = docModelToSmartPages(docModel, school);
      const nextSmart = {
        pages: converted.pages,
        meta: converted.meta,
        targetUnitId: school.unitId,
        schoolProfile: school
      };
      store.setState({ smartDocument: nextSmart }, false);
      store.setActiveVaultId(null, 'smart');
    }

    return {
      targetWorkspace: targetWs,
      appliedCount: 1,
      changes: computeDryRunChangesForMessage(action, targetWs),
      undoBackup: {
        workspace: targetWs,
        prevDocument: prevDocBackup,
        prevSmartDocument: prevSmartBackup,
        prevOpts
      }
    };
  }

  if (action.type === 'edit_letter' && action.editPlan) {
    const allowLayoutRestructure = shouldAllowLayoutRestructure(action.userInstruction || '');
    const explicitSwitchUnit = detectTargetUnitFromText(action.userInstruction || '', null);
    if (targetWs === 'canvas') {
      const baseDoc = state.document ? JSON.parse(JSON.stringify(state.document)) : null;
      const docSchool = explicitSwitchUnit
        ? school
        : resolveDocumentSchoolProfile(state.document, school, state.kopGallery || []);
      const patchResult = applySurgicalPatchesToDocModel(state.document, action.editPlan, {
        allowLayoutRestructure,
        letterOptions: prevOpts
      });
      if (patchResult.mode === 'restructure_needed' && Array.isArray(action.editPlan.blocks)) {
        const rebuilt = aiJsonToSmartDocModel(action.editPlan, docSchool, prevOpts);
        rebuilt.targetUnitId = docSchool.unitId;
        store.clearActiveUploadedDocx();
        store.setState({ document: rebuilt, docxViewMode: 'canvas' }, true);
        return {
          targetWorkspace: 'canvas',
          appliedCount: 1,
          changes: [{ targetLabel: 'Dokumen Lengkap', before: 'Struktur lama', after: action.editPlan.summaryOfChanges || 'Diperbarui' }],
          undoBackup: { workspace: 'canvas', prevDocument: baseDoc, prevOpts }
        };
      }
      if (patchResult.pageSetupUpdates) {
        store.updateLetterOptions(patchResult.pageSetupUpdates);
      }
      let patchedDoc = {
        ...patchResult.docModel,
        targetUnitId: docSchool.unitId,
        letterData: {
          ...(patchResult.docModel?.letterData || {}),
          schoolProfile: docSchool
        },
        ...(state.document?.verification ? { verification: state.document.verification } : {}),
        ...(state.document?.qrPlacement ? { qrPlacement: state.document.qrPlacement } : {})
      };
      if (explicitSwitchUnit) {
        patchedDoc = adaptDocumentModelForUnit(patchedDoc, explicitSwitchUnit, {
          activeSchoolProfile: baseSchool,
          kopGallery: state.kopGallery || []
        });
      }
      store.clearActiveUploadedDocx();
      store.setState({ document: patchedDoc, docxViewMode: 'canvas' }, true);
      return {
        targetWorkspace: 'canvas',
        appliedCount: patchResult.appliedCount || 1,
        changes: patchResult.changes || [],
        undoBackup: { workspace: 'canvas', prevDocument: baseDoc, prevOpts }
      };
    }

    // Target Word Pro (smart-canvas)
    const prevSmartBackup = state.smartDocument ? JSON.parse(JSON.stringify(state.smartDocument)) : null;
    const effectiveSmartSchool = explicitSwitchUnit
      ? school
      : resolveDocumentSchoolProfile(state.smartDocument, school, state.kopGallery || []);
    let livePages = Array.isArray(state.smartDocument?.pages) ? state.smartDocument.pages : [];
    if (livePages.length === 0 && state.document) {
      livePages = docModelToSmartPages(state.document, effectiveSmartSchool).pages;
    }
    const patchResult = applySurgicalPatchesToSmartPages(livePages, action.editPlan, {
      schoolProfile: effectiveSmartSchool,
      allowLayoutRestructure,
      letterOptions: prevOpts,
      renderBlocksToPageHtml: (aiBlocks = []) => {
        const miniDoc = { blocks: convertAiBlocksToDocumentBlocks(aiBlocks) };
        return docModelToSmartPages(miniDoc, effectiveSmartSchool).pages.map(mp => mp.html).filter(Boolean).join('\n');
      }
    });

    let nextSmart;
    if (patchResult.mode === 'restructure_needed' && Array.isArray(action.editPlan.blocks)) {
      const rebuilt = aiJsonToSmartDocModel(action.editPlan, effectiveSmartSchool, prevOpts);
      rebuilt.targetUnitId = effectiveSmartSchool.unitId;
      const converted = docModelToSmartPages(rebuilt, effectiveSmartSchool);
      nextSmart = {
        pages: converted.pages,
        meta: converted.meta,
        targetUnitId: effectiveSmartSchool.unitId,
        schoolProfile: effectiveSmartSchool,
        ...(state.smartDocument?.verification ? { verification: state.smartDocument.verification } : {}),
        ...(state.smartDocument?.qrPlacement ? { qrPlacement: state.smartDocument.qrPlacement } : {})
      };
    } else {
      if (patchResult.pageSetupUpdates) {
        store.updateLetterOptions(patchResult.pageSetupUpdates);
      }
      nextSmart = {
        ...(state.smartDocument || {}),
        pages: patchResult.pages,
        meta: patchResult.meta || state.smartDocument?.meta,
        targetUnitId: effectiveSmartSchool.unitId,
        schoolProfile: effectiveSmartSchool,
        ...(state.smartDocument?.verification ? { verification: state.smartDocument.verification } : {}),
        ...(state.smartDocument?.qrPlacement ? { qrPlacement: state.smartDocument.qrPlacement } : {})
      };
    }

    store.setState({ smartDocument: nextSmart }, false);
    return {
      targetWorkspace: 'smart-canvas',
      appliedCount: patchResult.appliedCount || 1,
      changes: patchResult.changes || [],
      undoBackup: { workspace: 'smart-canvas', prevSmartDocument: prevSmartBackup, prevOpts }
    };
  }

  return null;
}

const SVG_CHAT_BOT = `<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="8" width="18" height="12" rx="4"/><path d="M12 2v6"/><circle cx="12" cy="2" r="1.3" fill="currentColor"/><circle cx="8.5" cy="14" r="1.4" fill="currentColor"/><circle cx="15.5" cy="14" r="1.4" fill="currentColor"/></svg>`;
const SVG_PLUS = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>`;
const SVG_PIN = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 17v5"/><path d="M9 10.76a2 2 0 0 1-1.11 1.79l-1.78.9A2 2 0 0 0 5 15.24V16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-.76a2 2 0 0 0-1.11-1.79l-1.78-.9A2 2 0 0 1 15 10.76V7a1 1 0 0 1 1-1 2 2 0 0 0 0-4H8a2 2 0 0 0 0 4 1 1 0 0 1 1 1z"/></svg>`;
const SVG_CLIP = `<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m21.44 11.05-9.19 9.19a6 6 0 0 1-8.49-8.49l8.57-8.57A4 4 0 1 1 18 8.84l-8.59 8.57a2 2 0 0 1-2.83-2.83l8.49-8.48"/></svg>`;
const SVG_SEND = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>`;
const SVG_CHECK = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>`;
const SVG_EYE = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>`;

function renderQuickTplIcon(iconKey) {
  if (iconKey === 'calendar') {
    return `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>`;
  }
  if (iconKey === 'briefcase') {
    return `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2"/></svg>`;
  }
  if (iconKey === 'users') {
    return `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>`;
  }
  return `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>`;
}

const SVG_HISTORY = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><polyline points="12 7 12 12 15.5 14"/></svg>`;
const SVG_SPLIT = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="16" rx="2"/><line x1="9" y1="4" x2="9" y2="20"/></svg>`;
const SVG_EXPAND = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg>`;
const SVG_COMPRESS = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="4 14 10 14 10 20"/><polyline points="20 10 14 10 14 4"/><line x1="14" y1="10" x2="21" y2="3"/><line x1="3" y1="21" x2="10" y2="14"/></svg>`;
const SVG_ZAP = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg>`;
const SVG_ARCHIVE_KB = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/></svg>`;
const SVG_SPARKLE = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m12 3 1.9 5.1L19 10l-5.1 1.9L12 17l-1.9-5.1L5 10l5.1-1.9L12 3z"/></svg>`;
const SVG_BOOK = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/></svg>`;
const SVG_CALENDAR = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>`;

const sharedUndoMap = new Map(); // msgId -> undoBackup (berbagi lintas remount drawer Word)
const drawerWideStateMap = new Map(); // editorKey -> boolean (pertahankan Mode Fokus saat re-render kanvas)
const connectedControllers = new Set();

function refreshConnectedChatControllers() {
  for (const ctrl of [...connectedControllers]) {
    if (!ctrl.isConnected()) {
      connectedControllers.delete(ctrl);
      continue;
    }
    ctrl.refresh();
  }
}

/**
 * Controller UI Percakapan bersama untuk Full Tab & Compact Drawer
 */
function initChatController(rootEl, config = {}) {
  const isFull = config.mode === 'full';
  const isDirectCanvasDrawer = !isFull || config.directApply === true;
  const drawerKey = config.editorKind || 'drawer';
  let currentFolder = CHAT_FOLDERS.ALL;
  let searchQuery = '';
  let mobileSidebarOpen = false;
  let drawerHistoryOpen = false;
  let drawerWideOpen = !isFull && Boolean(drawerWideStateMap.get(drawerKey));
  let templatesCollapsed = false;
  let templatesUserToggled = false;
  let isGenerating = false;
  let activeAbortController = null;
  let pendingAttachments = [];
  const expandedDiffIds = new Set();

  const controllerHandle = {
    isConnected: () => Boolean(rootEl && rootEl.isConnected),
    refresh: () => renderAll()
  };
  connectedControllers.add(controllerHandle);

  const getPreferredWorkspace = () => {
    if (config.editorKind) return config.editorKind;
    return store.getActiveChatSession()?.targetWorkspace || 'smart-canvas';
  };

  function renderMessagesHtml(session) {
    const msgs = Array.isArray(session?.messages) ? session.messages : [];
    const initials = getUserInitials();

    if (msgs.length === 0 && !isGenerating) {
      return `
        <div class="aichat-welcome ${isFull ? '' : 'is-drawer'}">
          <div class="aichat-welcome-badge">${SVG_CHAT_BOT}</div>
          <h3>${isFull ? 'AI Asisten Persuratan Sekolah' : 'AI Copilot Siap Menyusun & Mengedit'}</h3>
          <p>
            ${isDirectCanvasDrawer
              ? 'Ketik perintah buat surat baru, ubah nama/tanggal, perbaikan PUEBI, atau atur format Word. Perubahan <strong>langsung diterapkan ke kanvas</strong> secara otomatis.'
              : 'Diskusikan redaksi surat, cek kaidah PUEBI, atur format Word/margin/nomor halaman, atau buat surat dinas resmi multi-halaman langsung dari percakapan ini.'}
          </p>
          <div class="aichat-welcome-tips">
            <span>✓ Mengingat konteks obrolan (multi-turn)</span>
            <span>✓ Bisa baca foto/PDF lampiran</span>
            <span>${isDirectCanvasDrawer ? '⚡ Langsung update kanvas + tombol Urungkan' : '✓ Pratinjau &amp; konfirmasi sebelum ubah editor'}</span>
          </div>
        </div>
      `;
    }

    const parts = msgs.map((m) => {
      const timeTxt = formatTimeOnly(m.timestamp);
      if (m.role === 'user') {
        const attHtml = Array.isArray(m.attachments) && m.attachments.length > 0
          ? `<div class="aichat-msg-atts">${m.attachments.map(a => `<span class="aichat-att-pill">📎 ${esc(a.name)}</span>`).join('')}</div>`
          : '';
        return `
          <div class="aichat-msg-row is-user">
            <div class="aichat-bubble is-user">
              ${attHtml}
              <div class="aichat-bubble-text">${formatRichInlineText(m.text)}</div>
              <div class="aichat-bubble-meta">
                <span>${esc(timeTxt)}</span>
                <span class="aichat-checks">✓✓</span>
              </div>
            </div>
            <div class="aichat-avatar is-user">${esc(initials)}</div>
          </div>
        `;
      }

      const checklistHtml = Array.isArray(m.checklist) && m.checklist.length > 0
        ? `
          <ul class="aichat-checklist">
            ${m.checklist.map(item => `
              <li>
                <span class="aichat-check-badge">${SVG_CHECK}</span>
                <span>${formatRichInlineText(item)}</span>
              </li>
            `).join('')}
          </ul>
        `
        : '';

      const followHtml = (m.followUpQuestion && !(isDirectCanvasDrawer && m.action?.status === 'applied'))
        ? `<div class="aichat-followup">${formatRichInlineText(m.followUpQuestion)}</div>`
        : '';

      let actionBarHtml = '';
      if (m.action && (m.action.type === 'create_letter' || m.action.type === 'edit_letter')) {
        const status = m.action.status || 'pending';
        const isCreate = m.action.type === 'create_letter';
        const actionUnitId = m.action.targetUnitId || m.action.schoolProfile?.unitId || '';
        const actionSchoolObj = actionUnitId ? (m.action.schoolProfile || SCHOOL_UNITS[actionUnitId] || null) : null;
        const unitSignerChip = actionSchoolObj
          ? `<div style="display:inline-flex;align-items:center;gap:6px;margin-top:6px;margin-bottom:2px;padding:3px 9px;border-radius:999px;font-size:11px;font-weight:700;background:${actionUnitId === 'smp' ? '#eff6ff' : '#f0fdf4'};color:${actionUnitId === 'smp' ? '#1d4ed8' : '#15803d'};border:1px solid ${actionUnitId === 'smp' ? '#bfdbfe' : '#bbf7d0'};">
               <span>🏫 ${esc(actionSchoolObj.unitLabel || actionSchoolObj.schoolName)}</span>
               <span>•</span>
               <span>Kepsek: ${esc(actionSchoolObj.principalName)}</span>
             </div>`
          : '';
        const diffOpen = expandedDiffIds.has(m.id);
        let diffHtml = '';
        if (diffOpen) {
          const items = computeDryRunChangesForMessage(m.action, session.targetWorkspace);
          diffHtml = `
            <div class="aichat-diff-box">
              <div class="aichat-diff-title">${isCreate ? 'Ringkasan Draf Surat Siap Diterapkan' : 'Rincian Perubahan (Sebelum → Sesudah)'}</div>
              ${items.length === 0
                ? '<div class="aichat-diff-empty">Semua struktur blok &amp; format siap diperbarui.</div>'
                : items.slice(0, 10).map(ch => `
                    <div class="aichat-diff-row">
                      <strong>${esc(ch.targetLabel || 'Bagian')}</strong>
                      <div class="aichat-diff-vals">
                        ${ch.before && ch.before !== '-' ? `<span class="diff-old">${esc(String(ch.before).slice(0, 75))}</span> → ` : ''}
                        <span class="diff-new">${esc(String(ch.after || '').slice(0, 95))}</span>
                      </div>
                    </div>
                  `).join('')}
            </div>
          `;
        }

        if (status === 'applied') {
          const hasUndo = sharedUndoMap.has(m.id);
          const appliedWsLabel = m.action.appliedWorkspace === 'canvas' ? 'di Lembar Kerja' : 'di Word Pro';
          actionBarHtml = `
            ${unitSignerChip}
            <div class="aichat-action-bar">
              <span class="aichat-applied-pill">${SVG_CHECK} ${isDirectCanvasDrawer ? `Langsung Diterapkan ${appliedWsLabel}` : `Sudah Diterapkan ${appliedWsLabel}`}</span>
              <button type="button" class="aichat-btn-sec" data-msg-act="toggle-diff" data-msg-id="${esc(m.id)}">${SVG_EYE} ${diffOpen ? 'Tutup Rincian' : 'Lihat Perubahan'}</button>
              ${hasUndo ? `<button type="button" class="aichat-btn-sec" data-msg-act="undo" data-msg-id="${esc(m.id)}">↶ Urungkan</button>` : ''}
              ${isFull ? `<button type="button" class="aichat-btn-sec" data-msg-act="open-editor" data-ws="${esc(m.action.appliedWorkspace || 'smart-canvas')}">↗ Buka Editor</button>` : ''}
            </div>
            ${diffHtml}
          `;
        } else {
          actionBarHtml = `
            ${unitSignerChip}
            <div class="aichat-action-bar">
              ${isCreate ? `
                <button type="button" class="aichat-btn-primary" data-msg-act="apply" data-msg-id="${esc(m.id)}" data-ws="${esc(config.editorKind || (session.targetWorkspace === 'canvas' ? 'canvas' : 'smart-canvas'))}">
                  ${SVG_CHECK} <span>${isFull ? 'Ya, Terapkan di Word Pro' : 'Ya, Terapkan ke Editor'}</span>
                </button>
                ${isFull ? `
                  <button type="button" class="aichat-btn-sec" data-msg-act="apply" data-msg-id="${esc(m.id)}" data-ws="canvas">
                    📄 Terapkan di Lembar Kerja
                  </button>
                ` : ''}
              ` : `
                <button type="button" class="aichat-btn-primary" data-msg-act="apply" data-msg-id="${esc(m.id)}" data-ws="${esc(config.editorKind || (session.targetWorkspace === 'canvas' ? 'canvas' : 'smart-canvas'))}">
                  ${SVG_CHECK} <span>Ya, Terapkan Perbaikan</span>
                </button>
              `}
              <button type="button" class="aichat-btn-sec" data-msg-act="toggle-diff" data-msg-id="${esc(m.id)}">
                ${SVG_EYE} <span>${diffOpen ? 'Tutup Perubahan' : 'Lihat Perubahan'}</span>
              </button>
            </div>
            ${diffHtml}
          `;
        }
      }

      return `
        <div class="aichat-msg-row is-assistant">
          <div class="aichat-avatar is-bot" title="${esc(m.modelLabel || 'AI Asisten')}">${SVG_CHAT_BOT}</div>
          <div class="aichat-bubble is-assistant">
            <div class="aichat-bubble-text">${formatRichInlineText(m.text)}</div>
            ${checklistHtml}
            ${followHtml}
            <div class="aichat-bubble-meta is-bot">
              <span>${esc(timeTxt)}${m.modelLabel ? ` • ${esc(m.modelLabel)}` : ''}</span>
            </div>
            ${actionBarHtml}
          </div>
        </div>
      `;
    });

    if (isGenerating) {
      parts.push(`
        <div class="aichat-msg-row is-assistant">
          <div class="aichat-avatar is-bot">${SVG_CHAT_BOT}</div>
          <div class="aichat-bubble is-assistant is-typing">
            <div class="aichat-typing-dots"><span></span><span></span><span></span></div>
            <span class="aichat-typing-label">AI sedang menganalisis konteks surat &amp; menyusun respons...</span>
            <button type="button" class="aichat-btn-stop" data-chat-act="abort">⏹ Hentikan</button>
          </div>
        </div>
      `);
    }

    return parts.join('');
  }

  function renderSidebarSessionsHtml(sessions, activeSession) {
    const stats = computeChatMemoryStats(sessions);
    const filtered = filterChatSessions(sessions, { folder: currentFolder, query: searchQuery });
    const pinnedList = filtered.filter(s => s.pinned && currentFolder !== CHAT_FOLDERS.TRASH && currentFolder !== CHAT_FOLDERS.ARCHIVED);
    const otherList = filtered.filter(s => !pinnedList.includes(s));

    const renderSessionRow = (s) => {
      const isActive = activeSession && s.id === activeSession.id;
      const timeLabel = formatChatTimestamp(s.updatedAt || s.createdAt);
      return `
        <div class="aichat-session-item ${isActive ? 'is-active' : ''}" data-session-id="${esc(s.id)}">
          <button type="button" class="aichat-session-main" data-select-session="${esc(s.id)}" title="${esc(s.title)}">
            <span class="aichat-session-title">${s.pinned ? '📌 ' : ''}${esc(s.title)}</span>
            <span class="aichat-session-time">${esc(timeLabel)}${isActive ? '<i class="aichat-active-dot"></i>' : ''}</span>
          </button>
          <div class="aichat-session-actions">
            ${s.trashed ? `
              <button type="button" data-sess-act="restore" data-session-id="${esc(s.id)}" title="Pulihkan">↩</button>
              <button type="button" data-sess-act="purge" data-session-id="${esc(s.id)}" title="Hapus Permanen">✕</button>
            ` : `
              <button type="button" data-sess-act="pin" data-session-id="${esc(s.id)}" title="${s.pinned ? 'Lepas Pin' : 'Pin Chat'}">${s.pinned ? '★' : '☆'}</button>
              <button type="button" data-sess-act="rename" data-session-id="${esc(s.id)}" title="Ganti Judul">✎</button>
              <button type="button" data-sess-act="trash" data-session-id="${esc(s.id)}" title="Pindah ke Sampah">🗑</button>
            `}
          </div>
        </div>
      `;
    };

    const foldersHtml = [
      { id: CHAT_FOLDERS.ALL, label: 'Semua Chat', icon: '💬', count: stats.allCount },
      { id: CHAT_FOLDERS.PINNED, label: 'Dipin', icon: '📌', count: stats.pinnedCount },
      { id: CHAT_FOLDERS.RECENT, label: 'Terbaru', icon: '🕒', count: stats.recentCount },
      { id: CHAT_FOLDERS.ARCHIVED, label: 'Arsip', icon: '📥', count: stats.archivedCount },
      { id: CHAT_FOLDERS.TRASH, label: 'Sampah', icon: '🗑', count: stats.trashCount }
    ].map(f => `
      <button type="button" class="aichat-folder-btn ${currentFolder === f.id ? 'is-active' : ''}" data-folder="${f.id}">
        <span class="aichat-folder-lead"><span>${f.icon}</span><span>${f.label}</span></span>
        ${f.count > 0 ? `<span class="aichat-folder-badge">${f.count}</span>` : ''}
      </button>
    `).join('');

    return `
      <aside class="aichat-sidebar ${mobileSidebarOpen ? 'is-open' : ''}">
        <div class="aichat-sidebar-top">
          <button type="button" class="aichat-new-btn" data-chat-act="new-chat">
            ${SVG_PLUS} <span>Chat Baru</span>
          </button>
          <div class="aichat-search-wrap">
            <input type="search" class="aichat-search-input" data-chat-search="1" value="${esc(searchQuery)}" placeholder="Cari percakapan..." aria-label="Cari percakapan">
          </div>
          <nav class="aichat-folders-nav">
            ${foldersHtml}
          </nav>
        </div>

        <div class="aichat-sessions-scroll">
          ${pinnedList.length > 0 ? `
            <div class="aichat-section-group">
              <div class="aichat-section-head">▾ Chat Dipin</div>
              ${pinnedList.map(renderSessionRow).join('')}
            </div>
          ` : ''}
          <div class="aichat-section-group">
            <div class="aichat-section-head">▾ ${currentFolder === CHAT_FOLDERS.TRASH ? 'Kotak Sampah' : currentFolder === CHAT_FOLDERS.ARCHIVED ? 'Diarsipkan' : 'Chat Terbaru'}</div>
            ${otherList.length > 0
              ? otherList.map(renderSessionRow).join('')
              : '<div class="aichat-empty-list">Belum ada percakapan di kategori ini.</div>'}
          </div>
        </div>

        <div class="aichat-memory-card">
          <div class="aichat-memory-row">
            <span class="aichat-memory-icon">🗄️</span>
            <div>
              <strong>Memori &amp; Riwayat</strong>
              <small>${stats.totalSaved} chat tersimpan &bull; ${stats.pinnedCount} dipin</small>
            </div>
          </div>
          <div class="aichat-memory-bar"><span style="width:${stats.quotaPct}%"></span></div>
        </div>
      </aside>
      ${mobileSidebarOpen ? '<div class="aichat-sidebar-backdrop" data-chat-act="close-sidebar"></div>' : ''}
    `;
  }

  function renderAll() {
    const activeSession = store.getActiveChatSession();
    const sessions = store.getChatSessions();
    const apiVault = store.getApiVault();
    const activeApi = store.getActiveApi();
    const selectedApiId = activeSession.activeModelId || activeApi?.id || apiVault[0]?.id || '';
    const autoApply = store.getChatAutoApply();
    const kbStats = store.getSmartAiReferencesForPrompt('', 1);
    const activeWs = getPreferredWorkspace();
    const docSnapInfo = buildActiveSnapshotFromStore(activeWs);

    const selectedUnitPref = activeSession.targetUnitPreference || 'auto';
    const unitOptionsHtml = `
      <option value="auto" ${selectedUnitPref === 'auto' ? 'selected' : ''}>🎯 Sekolah: Auto (Deteksi SMP/SMK)</option>
      <option value="smp" ${selectedUnitPref === 'smp' ? 'selected' : ''}>🏫 SMP IT (Yangyang Budiman, S.Pd.)</option>
      <option value="smk" ${selectedUnitPref === 'smk' ? 'selected' : ''}>🏫 SMK IT (Dian Nurhamdani, S.Pd.)</option>
    `;

    const modelOptionsHtml = apiVault.map(api =>
      `<option value="${esc(api.id)}" ${api.id === selectedApiId ? 'selected' : ''}>${esc(api.label || api.model)} [${esc(String(api.provider || 'AI').toUpperCase())}]</option>`
    ).join('');

    const sessionOptionsHtml = sessions
      .filter(s => !s.trashed && !s.archived)
      .slice(0, 15)
      .map(s => `<option value="${esc(s.id)}" ${s.id === activeSession.id ? 'selected' : ''}>${s.pinned ? '📌 ' : ''}${esc(s.title)}</option>`)
      .join('');

    const pendingAttHtml = pendingAttachments.length > 0
      ? `<div class="aichat-pending-atts">${pendingAttachments.map(a => `
          <span class="aichat-pending-chip">
            📎 ${esc(a.name)} (${a.sizeKb} KB)
            <button type="button" data-remove-att="${esc(a.id)}" title="Hapus lampiran">×</button>
          </span>`).join('')}</div>`
      : '';

    const quickCardsHtml = CHAT_QUICK_TEMPLATES.map(t => `
      <button type="button" class="aichat-tpl-card" data-tpl-prompt="${esc(t.prompt)}">
        <span class="aichat-tpl-icon">${renderQuickTplIcon(t.icon)}</span>
        <span class="aichat-tpl-title">${esc(t.title)}</span>
      </button>
    `).join('');

    if (isFull) {
      rootEl.innerHTML = `
        <div class="aichat-page-shell">
          <div class="aichat-hero-strip">
            <div class="aichat-hero-lead">
              <span class="aichat-hero-badge">${SVG_CHAT_BOT}</span>
              <div>
                <h2>AI Asisten Persuratan</h2>
                <p>Asisten cerdas untuk menyusun, memperbaiki, dan mengelola surat resmi sekolah secara percakapan.</p>
              </div>
            </div>
            <div class="aichat-hero-ctx">
              <span class="aichat-kb-pill">📚 ${kbStats.totalKnowledgeCount} arsip &bull; ⭐ ${kbStats.masterCount} master</span>
              ${docSnapInfo.hasContent ? `<span class="aichat-doc-pill">📄 Konteks Aktif: ${esc(docSnapInfo.snapshot?.letterTitle || 'Dokumen Terbuka')}</span>` : ''}
            </div>
          </div>

          <div class="aichat-frame">
            ${renderSidebarSessionsHtml(sessions, activeSession)}

            <section class="aichat-main">
              <header class="aichat-topbar">
                <div class="aichat-topbar-left">
                  <button type="button" class="aichat-mobile-menu-btn" data-chat-act="open-sidebar" title="Buka Riwayat Chat">☰ Riwayat</button>
                  <label class="aichat-model-label">Model AI</label>
                  <select class="aichat-model-select" data-chat-select="model">${modelOptionsHtml}</select>
                  <span class="aichat-status-badge"><i class="aichat-status-dot"></i> Aktif</span>
                </div>
                <div class="aichat-topbar-right">
                  <select class="aichat-ctx-select" data-chat-select="unit" title="Deteksi sekolah &amp; Kepala Sekolah otomatis dari perintah atau pilih manual SMP/SMK">
                    ${unitOptionsHtml}
                  </select>
                  <select class="aichat-ctx-select" data-chat-select="workspace" title="Editor tujuan saat menerapkan surat/perbaikan">
                    <option value="smart-canvas" ${activeSession.targetWorkspace === 'smart-canvas' ? 'selected' : ''}>Target: Word Pro</option>
                    <option value="canvas" ${activeSession.targetWorkspace === 'canvas' ? 'selected' : ''}>Target: Lembar Kerja</option>
                    <option value="none" ${activeSession.targetWorkspace === 'none' ? 'selected' : ''}>Konsultasi Umum</option>
                  </select>
                  <label class="aichat-autoapply-toggle" title="Jika aktif, AI langsung menerapkan perubahan ke editor sekaligus mencatat tombol Urungkan">
                    <input type="checkbox" data-chat-toggle="auto-apply" ${autoApply ? 'checked' : ''}>
                    <span>Auto Terapkan</span>
                  </label>
                </div>
              </header>

              <div class="aichat-messages-scroll" data-chat-scroll="1">
                ${renderMessagesHtml(activeSession)}
              </div>

              <div class="aichat-bottom-dock">
                <div class="aichat-quick-strip">
                  <div class="aichat-quick-head">
                    <span>⚡ <strong>Saran Cepat &amp; Template</strong></span>
                    <div class="aichat-quick-links">
                      <button type="button" class="aichat-link-btn" data-chat-act="toggle-tpls">${templatesCollapsed ? 'Tampilkan ▾' : 'Sembunyikan ▴'}</button>
                      <button type="button" class="aichat-link-btn" data-chat-act="open-vault">Lihat Semua Template →</button>
                    </div>
                  </div>
                  ${templatesCollapsed ? '' : `<div class="aichat-tpl-grid">${quickCardsHtml}</div>`}
                </div>

                ${pendingAttHtml}

                <div class="aichat-composer-card">
                  <input type="file" accept="image/*,application/pdf" class="hidden" data-chat-file="1" multiple>
                  <button type="button" class="aichat-attach-btn" data-chat-act="attach" title="Lampirkan Foto Surat / PDF untuk dibaca AI Vision">
                    ${SVG_CLIP}
                  </button>
                  <textarea
                    rows="1"
                    class="aichat-input"
                    data-chat-input="1"
                    spellcheck="false"
                    placeholder="Ketik perintah atau pertanyaan tentang surat sekolah..."
                  ></textarea>
                  <button type="button" class="aichat-send-btn" data-chat-act="send" ${isGenerating ? 'disabled' : ''} title="Kirim Pesan (Enter)">
                    ${SVG_SEND}
                  </button>
                </div>
                <div class="aichat-hints-row">
                  <span>✨ Contoh: ${CHAT_HINT_EXAMPLES.join(', ')}</span>
                </div>
              </div>
            </section>
          </div>
        </div>
      `;
    } else {
      // Mode Drawer Copilot di dalam Word Pro / Lembar Kerja (bisa diperlebar ke Mode Fokus 2-kolom tanpa pindah tab)
      const shouldRestoreDrawerFocus = document.activeElement?.getAttribute?.('data-chat-input') === '1'
        && rootEl.contains(document.activeElement);
      const msgCount = Array.isArray(activeSession?.messages) ? activeSession.messages.length : 0;
      const showQuickTpls = templatesUserToggled ? !templatesCollapsed : (msgCount <= 2);
      rootEl.classList.toggle('is-drawer-wide', drawerWideOpen);

      rootEl.innerHTML = `
        <div class="aichat-drawer-hero">
          <div class="aichat-drawer-hero-top">
            <div class="aichat-drawer-hero-lead">
              <span class="aichat-hero-badge is-drawer">${SVG_CHAT_BOT}</span>
              <div class="aichat-drawer-hero-titles">
                <div class="aichat-drawer-title-row">
                  <h4>${esc(config.title || 'AI Copilot Word Pro')}</h4>
                  <span class="aichat-beta-pill">${drawerWideOpen ? 'Mode Fokus' : 'Langsung'}</span>
                </div>
                <p class="aichat-drawer-sub">${config.editorKind === 'canvas' ? 'Eksekusi langsung ke kanvas Lembar Kerja aktif.' : 'Eksekusi langsung ke kanvas Word Pro aktif.'}</p>
              </div>
            </div>
            <div class="aichat-drawer-head-actions">
              <button type="button" class="aichat-head-pill-btn ${drawerHistoryOpen || drawerWideOpen ? 'is-active' : ''}" data-chat-act="toggle-drawer-history" title="Buka/tutup Riwayat Chat tersimpan">${SVG_HISTORY} <span>Riwayat</span></button>
              <button type="button" class="aichat-head-pill-btn" data-chat-act="new-chat" title="Mulai percakapan baru">+ Baru</button>
              <button type="button" class="aichat-head-pill-btn aichat-wide-toggle-btn ${drawerWideOpen ? 'is-active' : ''}" data-chat-act="toggle-drawer-wide" title="Perlebar panel menjadi Mode Fokus Chat 2-kolom langsung di dalam editor">${drawerWideOpen ? `${SVG_COMPRESS} <span>Perkecil</span>` : `${SVG_EXPAND} <span>Mode Fokus</span>`}</button>
              <button type="button" class="copilot-close-btn" data-chat-act="close-drawer" title="Tutup Panel">&times;</button>
            </div>
          </div>
        </div>

        <div class="aichat-drawer-stage ${drawerWideOpen ? 'is-wide' : ''}">
          ${(drawerHistoryOpen || drawerWideOpen) ? `
            <div class="aichat-drawer-sidebar-slot ${drawerWideOpen ? 'is-docked' : 'is-overlay'}">
              ${renderSidebarSessionsHtml(sessions, activeSession)}
            </div>
            ${drawerHistoryOpen && !drawerWideOpen ? '<div class="aichat-sidebar-backdrop is-drawer" data-chat-act="close-drawer-history"></div>' : ''}
          ` : ''}

          <div class="aichat-drawer-chatcol">
            <div class="aichat-drawer-topbar">
              <div class="aichat-drawer-model-row">
                <span class="aichat-model-label">Model AI</span>
                <select class="aichat-model-select is-drawer" data-chat-select="model" title="Pilih Model AI">${modelOptionsHtml}</select>
                <span class="aichat-status-badge"><i class="aichat-status-dot"></i> Aktif</span>
              </div>
              <div class="aichat-drawer-ctxbar">
                <span class="aichat-direct-pill" title="Perintah di panel editor otomatis membuat / memperbarui surat di kanvas aktif">${SVG_ZAP}<span>Langsung ke ${config.editorKind === 'canvas' ? 'Lembar Kerja' : 'Word Pro'}</span></span>
                <div class="aichat-drawer-ctx-right">
                  <select class="aichat-model-select is-drawer" data-chat-select="unit" style="max-width:136px;font-size:11px;padding:2px 6px;" title="Target Sekolah &amp; Kepala Sekolah">${unitOptionsHtml}</select>
                  <span class="aichat-drawer-kb" title="Basis Pengetahuan dari Gudang Surat">${SVG_ARCHIVE_KB}<span>${kbStats.totalKnowledgeCount} arsip</span></span>
                  <button type="button" class="aichat-link-btn" data-chat-act="open-vault" style="font-size:11px;">Gudang</button>
                </div>
              </div>
            </div>

            <div class="aichat-messages-scroll is-drawer" data-chat-scroll="1">
              ${renderMessagesHtml(activeSession)}
            </div>

            <div class="aichat-drawer-dock">
              <div class="aichat-drawer-quick-row">
                <button type="button" class="aichat-mini-pill" data-quick-editor="polish" title="Rapikan naskah berantakan menjadi surat resmi">${SVG_SPARKLE}<span>Rapihkan</span></button>
                <button type="button" class="aichat-mini-pill" data-quick-editor="puebi" title="Perbaiki ejaan ke standar PUEBI">${SVG_BOOK}<span>PUEBI</span></button>
                <button type="button" class="aichat-mini-pill" data-quick-editor="today" title="Perbarui tanggal surat ke hari ini">${SVG_CALENDAR}<span>Tgl Hari Ini</span></button>
              </div>

              <div class="aichat-quick-strip is-drawer">
                <div class="aichat-quick-head">
                  <span class="aichat-quick-head-lead">${SVG_ZAP}<strong>Saran Cepat &amp; Template</strong></span>
                  <div class="aichat-quick-links">
                    <button type="button" class="aichat-link-btn" data-chat-act="toggle-tpls">${showQuickTpls ? 'Sembunyikan ▴' : 'Tampilkan ▾'}</button>
                    <button type="button" class="aichat-link-btn" data-chat-act="open-vault">Lihat Semua Template →</button>
                  </div>
                </div>
                ${showQuickTpls ? `<div class="aichat-tpl-grid is-compact">${quickCardsHtml}</div>` : ''}
              </div>

              ${pendingAttHtml}

              <div class="aichat-composer-card is-compact">
                <input type="file" accept="image/*,application/pdf" class="hidden" data-chat-file="1" multiple>
                <button type="button" class="aichat-attach-btn" data-chat-act="attach" title="Lampirkan Foto Surat / PDF untuk dibaca AI Vision">${SVG_CLIP}</button>
                <textarea rows="1" class="aichat-input" data-chat-input="1" spellcheck="false" placeholder="Ketik perintah buat/edit surat & tekan Enter (langsung masuk ke kanvas)..."></textarea>
                <button type="button" class="aichat-send-btn" data-chat-act="send" ${isGenerating ? 'disabled' : ''} title="Kirim &amp; Eksekusi Langsung ke Kanvas (Enter)">${SVG_SEND}</button>
              </div>
              <div class="aichat-hints-row is-drawer">
                <span>Contoh: ${CHAT_HINT_EXAMPLES.slice(0, 2).join(', ')}</span>
              </div>
            </div>
          </div>
        </div>
      `;

      bindEvents();
      const scrollEl = rootEl.querySelector('[data-chat-scroll="1"]');
      if (scrollEl) {
        scrollEl.scrollTop = scrollEl.scrollHeight;
      }
      if (shouldRestoreDrawerFocus) {
        rootEl.querySelector('[data-chat-input="1"]')?.focus({ preventScroll: true });
      }
      return;
    }

    bindEvents();
    const scrollEl = rootEl.querySelector('[data-chat-scroll="1"]');
    if (scrollEl) {
      scrollEl.scrollTop = scrollEl.scrollHeight;
    }
  }

  async function applyMessageAction(msgId, explicitWs = '', { silent = false } = {}) {
    const session = store.getActiveChatSession();
    const msg = (session.messages || []).find(m => m.id === msgId);
    if (!msg?.action) return;

    const targetWs = (isDirectCanvasDrawer && config.editorKind)
      ? config.editorKind
      : (explicitWs || config.editorKind || session.targetWorkspace || 'smart-canvas');
    let applyMeta = null;

    if (config.editorKind && (targetWs === config.editorKind)) {
      const unitCtx = {
        targetUnitId: msg.action.targetUnitId || '',
        schoolProfile: msg.action.schoolProfile || null
      };
      if (msg.action.type === 'create_letter' && typeof config.onApplyCreateDraft === 'function') {
        const draftPayload = {
          ...msg.action.letterDraft,
          ...(unitCtx.targetUnitId ? { _targetUnitId: unitCtx.targetUnitId } : {}),
          ...(unitCtx.schoolProfile ? { _targetSchoolProfile: unitCtx.schoolProfile } : {})
        };
        applyMeta = await config.onApplyCreateDraft(draftPayload, unitCtx);
      } else if (msg.action.type === 'edit_letter' && typeof config.onApplySurgicalEdit === 'function') {
        const editPayload = {
          ...msg.action.editPlan,
          ...(unitCtx.targetUnitId ? { _targetUnitId: unitCtx.targetUnitId } : {}),
          ...(unitCtx.schoolProfile ? { _targetSchoolProfile: unitCtx.schoolProfile } : {})
        };
        applyMeta = await config.onApplySurgicalEdit(editPayload, msg.action.userInstruction || '', unitCtx);
      }
    }

    if (!applyMeta) {
      applyMeta = applyActionToStoreWorkspace(msg.action, targetWs);
    }

    if (applyMeta?.undoBackup) {
      sharedUndoMap.set(msg.id, applyMeta.undoBackup);
    }

    store.updateChatMessage(session.id, msg.id, {
      action: {
        ...msg.action,
        status: 'applied',
        appliedWorkspace: targetWs,
        ...(Array.isArray(applyMeta?.changes) ? { changes: applyMeta.changes } : {})
      }
    });

    if (!silent) {
      showToast(`✓ Langsung diterapkan ke ${targetWs === 'canvas' ? 'Lembar Kerja' : 'Word Pro'}!`, 'success');
    }
    refreshConnectedChatControllers();
  }

  async function undoMessageAction(msgId) {
    const session = store.getActiveChatSession();
    const msg = (session.messages || []).find(m => m.id === msgId);
    const backup = sharedUndoMap.get(msgId);
    if (!msg || !backup) return;

    if (config.editorKind && typeof config.onUndoAction === 'function') {
      config.onUndoAction(backup);
    } else {
      if (backup.prevOpts) {
        store.updateLetterOptions(backup.prevOpts);
      }
      if (backup.workspace === 'canvas' && backup.prevDocument) {
        store.setState({ document: backup.prevDocument, docxViewMode: 'canvas' }, true);
      } else if (backup.prevSmartDocument) {
        store.setState({ smartDocument: backup.prevSmartDocument }, false);
      }
    }

    store.updateChatMessage(session.id, msg.id, {
      action: { ...msg.action, status: 'pending' }
    });
    sharedUndoMap.delete(msgId);
    showToast('↶ Perubahan AI telah diurungkan.', 'info');
    refreshConnectedChatControllers();
  }

  async function sendUserPrompt(overrideText = '') {
    if (isGenerating) return;
    const inputEl = rootEl.querySelector('[data-chat-input="1"]');
    const text = String(overrideText || inputEl?.value || '').trim();
    if (!text && pendingAttachments.length === 0) {
      showToast('Ketikkan pertanyaan atau instruksi surat terlebih dahulu.', 'warning');
      return;
    }

    const session = store.getActiveChatSession();
    const attachmentsToSend = [...pendingAttachments];
    const attMeta = attachmentsToSend.map(({ id, name, mimeType, sizeKb }) => ({ id, name, mimeType, sizeKb }));

    if (inputEl && !overrideText) inputEl.value = '';
    pendingAttachments = [];

    const userMsg = store.appendChatMessage(session.id, {
      role: 'user',
      text: text || `Tolong baca lampiran ${attMeta.map(a => a.name).join(', ')} dan buatkan ringkasan atau tindak lanjut suratnya.`,
      attachments: attMeta
    });

    isGenerating = true;
    activeAbortController = new AbortController();
    renderAll();

    try {
      const latestSession = store.getActiveChatSession();
      const targetWs = getPreferredWorkspace();
      let docSnap = null;
      if (targetWs !== 'none') {
        docSnap = typeof config.getSnapshot === 'function'
          ? config.getSnapshot()
          : buildActiveSnapshotFromStore(targetWs).snapshot;
      }

      const selectedApiId = rootEl.querySelector('[data-chat-select="model"]')?.value || latestSession.activeModelId || store.getActiveApi()?.id;
      const activeEntry = store.getApiVault().find(v => v.id === selectedApiId) || store.getActiveApi();
      const priorMsgs = (latestSession.messages || []).filter(m => m.id !== userMsg.id);

      const preferredUnitPref = rootEl.querySelector('[data-chat-select="unit"]')?.value || latestSession.targetUnitPreference || 'auto';
      const activeDocEffectiveSchool = store.getEffectiveSchoolForDocument();

      const res = await callAiChatTurn({
        conversationHistory: priorMsgs,
        userMessage: userMsg.text,
        activeDocumentSnapshot: docSnap,
        targetWorkspace: targetWs,
        schoolProfile: activeDocEffectiveSchool,
        preferredUnitId: preferredUnitPref,
        preferredApiId: selectedApiId,
        attachments: attachmentsToSend.length > 0 ? attachmentsToSend : null,
        attachmentsMeta: attMeta,
        executionMode: isDirectCanvasDrawer ? 'direct_apply' : 'confirm',
        signal: activeAbortController.signal
      });

      let actionObj = null;
      if (res.intent === 'create_letter' && res.letterDraft) {
        actionObj = {
          type: 'create_letter',
          status: 'pending',
          targetWorkspace: targetWs === 'none' ? 'smart-canvas' : targetWs,
          targetUnitId: res.targetUnitId || '',
          schoolProfile: res.schoolProfile || null,
          userInstruction: userMsg.text,
          letterDraft: res.letterDraft
        };
      } else if (res.intent === 'edit_letter' && res.editPlan) {
        actionObj = {
          type: 'edit_letter',
          status: 'pending',
          targetWorkspace: targetWs === 'none' ? 'smart-canvas' : targetWs,
          targetUnitId: res.targetUnitId || '',
          schoolProfile: res.schoolProfile || null,
          userInstruction: userMsg.text,
          editPlan: res.editPlan
        };
      }

      const savedBotMsg = store.appendChatMessage(latestSession.id, {
        role: 'assistant',
        text: res.reply,
        checklist: res.checklist,
        followUpQuestion: res.followUpQuestion,
        action: actionObj,
        modelLabel: activeEntry?.label?.split('(')[0]?.trim() || activeEntry?.model || 'AI'
      });

      isGenerating = false;
      activeAbortController = null;
      renderAll();

      // Mode drawer nempel di kanvas selalu langsung menerapkan perubahan; Tab AI Asisten mengikuti toggle Auto Terapkan
      if (actionObj && (isDirectCanvasDrawer || store.getChatAutoApply())) {
        await applyMessageAction(savedBotMsg.id, config.editorKind || actionObj.targetWorkspace);
      }
    } catch (err) {
      isGenerating = false;
      const wasAborted = activeAbortController?.signal?.aborted || String(err?.message || '').includes('dibatalkan');
      activeAbortController = null;
      if (!wasAborted) {
        store.appendChatMessage(session.id, {
          role: 'assistant',
          text: `Maaf, permintaan belum dapat diproses: ${err.message}`,
          followUpQuestion: 'Silakan periksa koneksi atau pilih model lain di dropdown Model AI lalu coba kirim ulang.'
        });
        showToast(`AI Asisten: ${err.message}`, 'error');
      }
      renderAll();
    }
  }

  function bindEvents() {
    rootEl.querySelectorAll('[data-chat-act="new-chat"]').forEach(btn => {
      btn.onclick = () => {
        if (isGenerating) activeAbortController?.abort();
        store.createChatSession({ targetWorkspace: getPreferredWorkspace() });
        currentFolder = CHAT_FOLDERS.ALL;
        mobileSidebarOpen = false;
        drawerHistoryOpen = false;
        templatesUserToggled = false;
        renderAll();
      };
    });

    rootEl.querySelector('[data-chat-act="open-sidebar"]')?.addEventListener('click', () => {
      mobileSidebarOpen = true;
      renderAll();
    });
    rootEl.querySelector('[data-chat-act="close-sidebar"]')?.addEventListener('click', () => {
      mobileSidebarOpen = false;
      renderAll();
    });
    rootEl.querySelector('[data-chat-act="toggle-drawer-history"]')?.addEventListener('click', () => {
      if (drawerWideOpen) {
        drawerWideOpen = false;
        drawerWideStateMap.set(drawerKey, false);
        drawerHistoryOpen = false;
      } else {
        drawerHistoryOpen = !drawerHistoryOpen;
      }
      renderAll();
    });
    rootEl.querySelector('[data-chat-act="close-drawer-history"]')?.addEventListener('click', () => {
      drawerHistoryOpen = false;
      renderAll();
    });
    rootEl.querySelector('[data-chat-act="toggle-drawer-wide"]')?.addEventListener('click', () => {
      drawerWideOpen = !drawerWideOpen;
      drawerWideStateMap.set(drawerKey, drawerWideOpen);
      drawerHistoryOpen = false;
      renderAll();
    });

    rootEl.querySelector('[data-chat-act="close-drawer"]')?.addEventListener('click', () => {
      config.onClose?.();
    });
    rootEl.querySelectorAll('[data-chat-act="open-vault"]').forEach(btn => {
      btn.onclick = () => {
        if (typeof config.onOpenVault === 'function') config.onOpenVault();
        else document.querySelector('.nav-tab[data-tab="vault"]')?.click();
      };
    });

    rootEl.querySelector('[data-chat-act="toggle-tpls"]')?.addEventListener('click', () => {
      if (!isFull && !templatesUserToggled) {
        const msgCount = (store.getActiveChatSession()?.messages || []).length;
        templatesCollapsed = msgCount <= 2;
      } else {
        templatesCollapsed = !templatesCollapsed;
      }
      templatesUserToggled = true;
      renderAll();
    });

    rootEl.querySelector('[data-chat-act="abort"]')?.addEventListener('click', () => {
      activeAbortController?.abort();
      isGenerating = false;
      renderAll();
    });

    rootEl.querySelectorAll('[data-folder]').forEach(btn => {
      btn.onclick = () => {
        currentFolder = btn.dataset.folder;
        renderAll();
      };
    });

    const searchInp = rootEl.querySelector('[data-chat-search="1"]');
    if (searchInp) {
      searchInp.oninput = (e) => {
        searchQuery = e.target.value;
        const pos = searchInp.selectionStart || searchQuery.length;
        renderAll();
        const reInp = rootEl.querySelector('[data-chat-search="1"]');
        if (reInp) {
          reInp.focus();
          try { reInp.setSelectionRange(pos, pos); } catch {}
        }
      };
    }

    rootEl.querySelectorAll('[data-select-session]').forEach(btn => {
      btn.onclick = () => {
        store.setActiveChatSession(btn.dataset.selectSession);
        mobileSidebarOpen = false;
        drawerHistoryOpen = false;
        renderAll();
      };
    });

    rootEl.querySelectorAll('[data-sess-act]').forEach(btn => {
      btn.onclick = (e) => {
        e.stopPropagation();
        const sid = btn.dataset.sessionId;
        const act = btn.dataset.sessAct;
        if (act === 'pin') store.togglePinChatSession(sid);
        else if (act === 'trash') store.trashChatSession(sid, true);
        else if (act === 'restore') store.trashChatSession(sid, false);
        else if (act === 'purge') store.deleteChatSessionPermanent(sid);
        else if (act === 'rename') {
          const cur = store.getChatSessions().find(s => s.id === sid);
          const nextTitle = prompt('Judul percakapan:', cur?.title || 'Percakapan');
          if (nextTitle && nextTitle.trim()) {
            store.updateChatSession(sid, { title: nextTitle.trim(), isAutoTitle: false });
          }
        }
        renderAll();
      };
    });

    rootEl.querySelector('[data-chat-select="session"]')?.addEventListener('change', (e) => {
      store.setActiveChatSession(e.target.value);
      renderAll();
    });

    rootEl.querySelector('[data-chat-select="model"]')?.addEventListener('change', (e) => {
      const apiId = e.target.value;
      store.setActiveApi(apiId);
      const sess = store.getActiveChatSession();
      store.updateChatSession(sess.id, { activeModelId: apiId });
    });

    rootEl.querySelector('[data-chat-select="workspace"]')?.addEventListener('change', (e) => {
      const sess = store.getActiveChatSession();
      store.updateChatSession(sess.id, { targetWorkspace: e.target.value });
      renderAll();
    });

    rootEl.querySelector('[data-chat-select="unit"]')?.addEventListener('change', (e) => {
      const sess = store.getActiveChatSession();
      store.updateChatSession(sess.id, { targetUnitPreference: e.target.value });
      renderAll();
    });

    rootEl.querySelector('[data-chat-toggle="auto-apply"]')?.addEventListener('change', (e) => {
      store.setChatAutoApply(e.target.checked);
    });

    rootEl.querySelectorAll('[data-tpl-prompt]').forEach(btn => {
      btn.onclick = () => {
        const inp = rootEl.querySelector('[data-chat-input="1"]');
        if (inp) {
          inp.value = btn.dataset.tplPrompt || '';
          inp.focus();
        }
      };
    });

    rootEl.querySelectorAll('[data-quick-editor]').forEach(btn => {
      btn.onclick = () => {
        if (typeof config.onQuickAction === 'function') {
          config.onQuickAction(btn.dataset.quickEditor);
        }
      };
    });

    const fileInput = rootEl.querySelector('[data-chat-file="1"]');
    rootEl.querySelector('[data-chat-act="attach"]')?.addEventListener('click', () => {
      fileInput?.click();
    });
    fileInput?.addEventListener('change', async (e) => {
      const files = [...(e.target.files || [])].slice(0, 3);
      for (const f of files) {
        try {
          const att = await readFileAsAttachment(f);
          pendingAttachments.push(att);
        } catch (err) {
          showToast(err.message, 'error');
        }
      }
      fileInput.value = '';
      renderAll();
    });

    rootEl.querySelectorAll('[data-remove-att]').forEach(btn => {
      btn.onclick = () => {
        pendingAttachments = pendingAttachments.filter(a => a.id !== btn.dataset.removeAtt);
        renderAll();
      };
    });

    rootEl.querySelectorAll('[data-msg-act="apply"]').forEach(btn => {
      btn.onclick = () => applyMessageAction(btn.dataset.msgId, btn.dataset.ws || '');
    });
    rootEl.querySelectorAll('[data-msg-act="toggle-diff"]').forEach(btn => {
      btn.onclick = () => {
        const id = btn.dataset.msgId;
        if (expandedDiffIds.has(id)) expandedDiffIds.delete(id);
        else expandedDiffIds.add(id);
        renderAll();
      };
    });
    rootEl.querySelectorAll('[data-msg-act="undo"]').forEach(btn => {
      btn.onclick = () => undoMessageAction(btn.dataset.msgId);
    });
    rootEl.querySelectorAll('[data-msg-act="open-editor"]').forEach(btn => {
      btn.onclick = () => {
        const ws = btn.dataset.ws || 'smart-canvas';
        if (typeof config.onOpenEditorTab === 'function') config.onOpenEditorTab(ws);
        else document.querySelector(`.nav-tab[data-tab="${ws}"]`)?.click();
      };
    });

    const composerInput = rootEl.querySelector('[data-chat-input="1"]');
    if (composerInput) {
      composerInput.addEventListener('keydown', (e) => {
        if (e.isComposing) return;
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          e.stopPropagation();
          sendUserPrompt();
        }
      });
      composerInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          e.stopPropagation();
        }
      });
      composerInput.addEventListener('keyup', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          e.stopPropagation();
        }
      });
    }
    rootEl.querySelector('[data-chat-act="send"]')?.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      sendUserPrompt();
    });
  }

  renderAll();
  return { refresh: renderAll };
}

/**
 * Merender halaman utama tab AI Asisten (Full Workspace 2-Kolom gaya ChatGPT)
 */
export function renderAiChatWorkspace(container, options = {}) {
  if (!container) return null;
  return initChatController(container, {
    mode: 'full',
    onOpenEditorTab: options.onNavigateTab,
    onOpenVault: () => options.onNavigateTab?.('vault')
  });
}

/**
 * Memasang Panel Percakapan AI Terpadu di dalam Drawer Word Pro / Lembar Kerja
 */
export function mountCompactChatDrawer(drawerEl, options = {}) {
  if (!drawerEl) return null;
  return initChatController(drawerEl, {
    mode: 'drawer',
    ...options
  });
}
