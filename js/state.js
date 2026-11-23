/**
 * state.js - Reactive State Store & IndexedDB Persistence
 * Mendukung penyimpanan Profil Sekolah, Form Surat Aktif, Gudang API Multi-Provider,
 * Gudang Galeri Kop Surat, Skala Zoom Kanvas, dan Riwayat Undo/Redo.
 */

import {
  DEFAULT_SCHOOL_PROFILE,
  adaptDocumentModelForUnit,
  adaptTextForSchoolUnit,
  getDefaultSchoolProfileForUnit,
  getSchoolCodeFromProfile,
  normalizeSchoolProfileNaming,
  normalizeSmkText,
  normalizeUnitId,
  resolveDocumentSchoolProfile,
  resolveSchoolProfileForUnit
} from './letters/schoolProfile.js';
import { suratKeteranganTemplate } from './letters/templates/suratKeterangan.js';
import { buildSuratTugasSppdDocModel } from './letters/templates/suratTugasSppd.js';
import { selectRelevantAiReferences } from './ai/knowledgeBase.js';
import { DEFAULT_KOP_GALLERY, DEFAULT_KOP_DATA_URL, getDefaultKopGalleryForUnit } from './letters/defaultKop.js';
import { detectLetterTypeCode, formatSequenceNumber, toRomanMonth, buildOrderedNomorSurat, LETTER_TYPE_CODES } from './letters/numbering.js';
import { formatIncomingCode, getNextIncomingSequence, sortIncomingLetters } from './letters/incoming.js';
import { PAGE_NUMBERING_DEFAULT, normalizePageSetupOptions } from './letters/wordFormat.js';
import {
  createChatSession,
  createChatMessage,
  deriveChatTitleFromMessage
} from './ai/chatEngine.js';
import { DEFAULT_ACCESS_PERMISSIONS } from './auth/permissions.js';
import {
  clearStoredAuthToken,
  getActiveWorkspaceUnitId,
  getStoredAuthToken,
  savePrivateWorkspaceCloud,
  saveWorkspaceStatePatch,
  setActiveWorkspaceUnitId
} from './sync/workspaceClient.js';

const DB_NAME = 'ai_school_doc_builder_db';
const DB_VERSION = 7;
const STORE_PROJECTS = 'projects';
const STORE_SCHOOL = 'school_profile';
const STORE_AGENDA = 'agenda_book';
const STORE_API_VAULT = 'api_vault';
const STORE_KOP_GALLERY = 'kop_gallery';
const STORE_LETTER_VAULT = 'letter_vault';
const STORE_OUTGOING_LETTERS = 'outgoing_letters';
const STORE_INCOMING_LETTERS = 'incoming_letters';
const STORE_INCOMING_FILES = 'incoming_files';
const STORE_CHAT_SESSIONS = 'chat_sessions';
const CHAT_SESSIONS_LS_KEY = 'ai_chat_sessions_v1';
const ACTIVE_CHAT_LS_KEY = 'ai_active_chat_id_v1';
const CHAT_AUTO_APPLY_LS_KEY = 'ai_chat_auto_apply_v1';

// Nilai awal form surat keterangan (SMK IT Al-Hawari)
const INITIAL_FORM_DATA = {
  nomorSurat: buildOrderedNomorSurat({ typeCode: 'KET', sequenceNumber: 1 }),
  tanggalSurat: '14 November 2026',
  namaSiswa: 'Ahmad Fauzi',
  nisNisn: '12345 / 0054321987',
  kelasJurusan: 'XII MM 1 / Multimedia',
  ttl: 'Garut, 12 Mei 2008',
  jenisKelamin: 'Laki-laki',
  namaOrtu: 'Bambang Sudarsono',
  pekerjaanOrtu: 'Wiraswasta',
  alamatSiswa: 'Kp. Curug RT 02 / RW 06, Desa Cipareuan, Kec. Cibiuk, Garut',
  tahunAjaran: '2026/2027',
  keperluan: 'Kelengkapan berkas pendaftaran Program Beasiswa Prestasi dan Tunjangan Pendidikan.'
};

/**
 * Membuat Document Model kanvas kosong bersih (tanpa teks contoh bawaan)
 */
export function createEmptyDocumentModel(schoolProfile = DEFAULT_SCHOOL_PROFILE, options = {}) {
  const showKop = options?.showKop !== false;
  const blocks = [];
  if (showKop) {
    blocks.push({ type: 'letterKop', school: schoolProfile });
  }
  blocks.push({ type: 'paragraph', style: 'letterBody', runs: [{ text: '' }] });

  return {
    meta: {
      title: 'Dokumen Baru',
      author: schoolProfile?.schoolName || 'SMK IT Al-Hawari',
      year: new Date().getFullYear().toString()
    },
    profile: options?.paperSize === 'f4' ? 'suratDinasF4' : 'suratDinas',
    paperSize: options?.paperSize || 'a4',
    orientation: options?.orientation || 'portrait',
    customMargins: options?.margins || { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 },
    isBlankDraft: true,
    blocks
  };
}

// Preset Bawaan Gudang Surat (Letter Vault & Acuan AI Referensi Sekolah)
export const DEFAULT_LETTER_VAULT = [
  {
    id: 'vault-ref-001',
    title: 'SURAT KETERANGAN SISWA AKTIF',
    nomorSurat: '421.3 / 085 / KET.SMKITA / XI / 2026',
    tanggalSurat: '14 November 2026',
    category: 'Surat Keterangan',
    targetPerson: 'Ahmad Fauzi (XII MM 1)',
    perihal: 'Keterangan Siswa Aktif Beasiswa Prestasi',
    summary: 'Menerangkan bahwa Ahmad Fauzi adalah siswa aktif kelas XII MM 1 Multimedia SMK IT Al-Hawari Tahun Ajaran 2026/2027 untuk keperluan pendaftaran beasiswa.',
    isAiReference: true,
    isMasterTemplate: true,
    isLockedTemplate: true,
    source: 'builder',
    document: suratKeteranganTemplate.generateDocumentModel(INITIAL_FORM_DATA, DEFAULT_SCHOOL_PROFILE, { paperSize: 'a4' }),
    createdAt: '2026-11-14T08:00:00Z'
  },
  {
    id: 'vault-ref-002',
    title: 'SURAT TUGAS BIMTEK KURIKULUM MERDEKA',
    nomorSurat: '800 / 042 / ST.SMKITA / XI / 2026',
    tanggalSurat: '10 November 2026',
    category: 'Surat Tugas',
    targetPerson: 'Dian Nurhamdani, S.Pd. & Tim Guru Kejuruan',
    perihal: 'Penugasan Pelatihan Implementasi Kurikulum Merdeka',
    summary: 'Menugaskan Kepala Sekolah dan guru kejuruan mengikuti Bimbingan Teknis Implementasi Kurikulum Merdeka di Balai Penjaminan Mutu Pendidikan (BPMP).',
    isAiReference: true,
    isMasterTemplate: false,
    isLockedTemplate: false,
    source: 'builder',
    createdAt: '2026-11-10T09:00:00Z'
  },
  {
    id: 'vault-ref-003',
    title: 'SURAT UNDANGAN RAPAT PLENO EVALUASI SEMESTER',
    nomorSurat: '005 / 028 / UND.SMKITA / XI / 2026',
    tanggalSurat: '8 November 2026',
    category: 'Surat Undangan',
    targetPerson: 'Seluruh Dewan Guru & Staf Tata Usaha',
    perihal: 'Rapat Persiapan Penilaian Akhir Semester (PAS) Ganjil',
    summary: 'Undangan rapat dinas pleno dewan guru dan staf tata usaha dalam rangka koordinasi jadwal, kisi-kisi soal, dan teknis administrasi ujian semester ganjil.',
    isAiReference: true,
    isMasterTemplate: true,
    isLockedTemplate: true,
    source: 'builder',
    createdAt: '2026-11-08T10:00:00Z'
  },
  {
    id: 'vault-ref-004',
    title: 'SURAT TUGAS & SPPD PENJEMPUTAN PKL (3 HALAMAN WORD)',
    nomorSurat: '065 / ST / SMKITA / 12.01 / XI / 2026',
    tanggalSurat: '18 November 2026',
    category: 'Surat Tugas & SPPD',
    targetPerson: 'Guru Pembimbing PKL SMK IT Al-Hawari',
    perihal: 'Surat Tugas & SPPD Penjemputan Peserta Didik PKL',
    summary: 'Paket lengkap 3 halaman 1:1 Microsoft Word: Halaman 1 Surat Tugas dengan tabel pegawai No/Nama/Jabatan, Halaman 2 SPPD 8 poin + Catatan PP & TTD 2 kolom, dan Halaman 3 Keterangan Pejabat & Tempat Dikunjungi dengan Header Merge.',
    isAiReference: true,
    isMasterTemplate: true,
    isLockedTemplate: true,
    source: 'builder',
    preferSmartCanvas: true,
    document: buildSuratTugasSppdDocModel({
      nomorSurat: '065/ST/SMKITA/12.01/XI/2026',
      assignees: [
        { nama: 'Taufik Dwi Setyo Ramadhan, S.Pd.', jabatan: 'Guru' },
        { nama: 'Siti Laras Maemunah, S.Pd.', jabatan: 'Guru' }
      ],
      maksudTugas: 'Penjemputan Peserta Didik PKL',
      tujuanLokasi: 'DRAGON MOTION PICTURE ASIA',
      tempatTiba: 'Dragon Motion Picture',
      maksudSppd: 'Penjemputan siswa kelas 12 Multimedia ke tempat PKL',
      Kendaraan: 'Mobil',
      kendaraanSingkat: 'Roda 4',
      lamaHari: '1'
    }, DEFAULT_SCHOOL_PROFILE, { paperSize: 'f4' }),
    createdAt: '2026-11-18T08:00:00Z'
  }
];

function buildDefaultVaultForUnit(unitId = 'smk') {
  const unit = normalizeUnitId(unitId);
  const rawList = JSON.parse(JSON.stringify(DEFAULT_LETTER_VAULT));
  if (unit === 'smk') return rawList;
  return rawList.map(item => ({
    ...item,
    nomorSurat: adaptTextForSchoolUnit(item.nomorSurat || '', unit).replace(/\b421\.3\b/g, '421.2'),
    targetPerson: adaptTextForSchoolUnit(item.targetPerson || '', unit).replace(/Tim Guru Kejuruan/gi, 'Tim Guru SMP'),
    summary: adaptTextForSchoolUnit(item.summary || '', unit),
    ...(item.document ? { document: adaptDocumentModelForUnit(item.document, unit) } : {})
  }));
}

// Kunci AI Bawaan Sistem (Fallback otomatis bila pengguna belum mengisi API Key sendiri)
const SYSTEM_API_SEEDS = {
  'gemini-auto': '=E1MNd3bIJWamNndLV2U6ZGWohVS4RXRtF3T481TFBVanVDbmJGUOZHcmNkS24kU4IWQuEVQ',
  'groq-fast': '=YmQHFnVUllcVRTZ5gFesh3SZZ0QFhjN2llRzIWekd0VQh2dRdkbqh1UuRTW1EHTLxkNRd2XrN3Z',
  'deepseek-v3': '=UGN1M2NkVTZmJzYlVjNwkjNilDNjFTO4E2YmFDZhFDOts2c'
};

function decodeSystemKey(seed) {
  if (!seed) return '';
  try {
    return globalThis.atob(seed.split('').reverse().join(''));
  } catch (e) {
    return '';
  }
}

export function getBuiltinSystemKey(vaultId) {
  return decodeSystemKey(SYSTEM_API_SEEDS[vaultId] || '');
}

export function isUsingBuiltinSystemKey(entry) {
  if (!entry || !entry.id || !entry.apiKey) return false;
  const builtinKey = getBuiltinSystemKey(entry.id);
  return Boolean(builtinKey && entry.apiKey === builtinKey);
}

// Preset Bawaan Gudang API (Dapat Ditambah/Diedit User)
export const DEFAULT_API_VAULT = [
  {
    id: 'gemini-auto',
    label: 'Google Gemini 3.5 Flash-Lite (Instan & 15 RPM)',
    provider: 'gemini',
    model: 'gemini-3.5-flash-lite',
    apiKey: getBuiltinSystemKey('gemini-auto'),
    baseUrl: 'https://generativelanguage.googleapis.com/v1beta',
    isActive: true,
    isBuiltin: true
  },
  {
    id: 'groq-fast',
    label: 'Groq - GPT OSS 120B (Ultra Cepat ~0.2s)',
    provider: 'groq',
    model: 'openai/gpt-oss-120b',
    apiKey: getBuiltinSystemKey('groq-fast'),
    baseUrl: 'https://api.groq.com/openai/v1',
    isActive: false,
    isBuiltin: true
  },
  {
    id: 'deepseek-v3',
    label: 'DeepSeek Chat (V3 / R1 Reasoner)',
    provider: 'deepseek',
    model: 'deepseek-chat',
    apiKey: getBuiltinSystemKey('deepseek-v3'),
    baseUrl: 'https://api.deepseek.com/v1',
    isActive: false,
    isBuiltin: true
  },
  {
    id: 'qwen-openrouter',
    label: 'Qwen 2.5 72B (OpenRouter / DashScope)',
    provider: 'qwen',
    model: 'qwen/qwen-2.5-72b-instruct',
    apiKey: '',
    baseUrl: 'https://openrouter.ai/api/v1',
    isActive: false,
    isBuiltin: true
  }
];

export function getBuiltinFallbackConfigs() {
  return DEFAULT_API_VAULT
    .filter(item => Boolean(item.apiKey))
    .map(item => ({ ...item }));
}

/**
 * Menghitung skala zoom awal kanvas agar proporsional dan lega di zoom browser 100%
 */
export function computeDefaultDesktopZoom(viewportWidth) {
  const w = typeof viewportWidth === 'number' && viewportWidth > 0
    ? viewportWidth
    : (typeof window !== 'undefined' ? window.innerWidth : 1920);
  if (w <= 860) return 100;
  if (w <= 1180) return 75;
  if (w <= 1480) return 80;
  if (w <= 1720) return 90;
  return 100;
}

class StateStore {
  constructor() {
    this.listeners = new Set();
    this.undoStack = [];
    this.redoStack = [];
    this.maxHistory = 50;

    const initialSchool = JSON.parse(JSON.stringify(DEFAULT_SCHOOL_PROFILE));
    const initialForm = JSON.parse(JSON.stringify(INITIAL_FORM_DATA));
    let savedOptions = null;
    if (typeof localStorage !== 'undefined') {
      try {
        const rawOpt = localStorage.getItem('letter_options');
        if (rawOpt) savedOptions = JSON.parse(rawOpt);
      } catch (e) {}
    }

    const initialOptions = savedOptions || {
      paperSize: 'a4',
      orientation: 'portrait',
      margins: { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 },
      fontFamily: 'Times New Roman',
      fontSize: '12',
      lineSpacing: '1.15',
      paragraphSpacing: '10',
      showKop: true
    };

    const initialDoc = createEmptyDocumentModel(initialSchool, initialOptions);

    // Muat API key lama dari localStorage bila ada
    let legacyKey = '';
    if (typeof localStorage !== 'undefined') {
      legacyKey = localStorage.getItem('ai_api_key') || '';
    }

    const initialVault = JSON.parse(JSON.stringify(DEFAULT_API_VAULT));
    if (legacyKey) {
      initialVault[0].apiKey = legacyKey;
    }

    let savedChatSessions = [];
    let savedActiveChatId = '';
    let savedChatAutoApply = false;
    if (typeof localStorage !== 'undefined') {
      try {
        const rawChats = localStorage.getItem(CHAT_SESSIONS_LS_KEY);
        if (rawChats) {
          const parsedChats = JSON.parse(rawChats);
          if (Array.isArray(parsedChats)) savedChatSessions = parsedChats;
        }
        savedActiveChatId = localStorage.getItem(ACTIVE_CHAT_LS_KEY) || '';
        savedChatAutoApply = localStorage.getItem(CHAT_AUTO_APPLY_LS_KEY) === 'true';
      } catch (e) {}
    }

    this.state = {
      activeTab: 'letter',
      schoolProfile: initialSchool,
      currentTemplateId: 'surat_keterangan',
      currentLetterForm: initialForm,
      letterOptions: initialOptions,
      document: initialDoc,
      isGenerating: false,
      streamingText: '',
      agendaBook: [],
      apiVault: initialVault,
      activeApiId: initialVault.find(v => v.isActive)?.id || 'gemini-auto',
      kopGallery: JSON.parse(JSON.stringify(DEFAULT_KOP_GALLERY)),
      zoomLevel: computeDefaultDesktopZoom(),
      zoomUserAdjusted: false,
      letterVault: JSON.parse(JSON.stringify(DEFAULT_LETTER_VAULT)),
      activeVaultId: null,
      activeSmartVaultId: null,
      outgoingLetters: [],
      incomingLetters: [],
      chatSessions: savedChatSessions,
      activeChatId: savedActiveChatId,
      chatAutoApply: savedChatAutoApply,
      smartDocument: null,
      activeUploadedDocx: null,
      pendingKepsekRevision: null,
      authUser: null,
      activeUnitId: typeof getActiveWorkspaceUnitId === 'function' ? getActiveWorkspaceUnitId() : 'smk',
      authPermissions: { ...DEFAULT_ACCESS_PERMISSIONS },
      driveStatus: {
        configured: false,
        accountEmail: 'tuancraf1@gmail.com',
        rootFolderId: '19nFdoJ4VOEgFmUe_heAFxoFEt-9U9Crw',
        rootFolderUrl: 'https://drive.google.com/drive/folders/19nFdoJ4VOEgFmUe_heAFxoFEt-9U9Crw',
        sheetUrl: '',
        lastSyncAt: '',
        lastError: ''
      },
      managedUsers: [],
      cloudSyncState: 'local',
      privateSyncState: 'local',
      cloudLastSyncAt: '',
      syncRev: 1
    };

    this.cloudSyncQueue = {};
    this.cloudSyncTimer = null;
    this.privateSyncTimer = null;
    this.suppressCloudSync = false;

    this.db = null;
    this.initIndexedDb();
  }

  getActiveUnitId() {
    return normalizeUnitId(this.state.activeUnitId || 'smk');
  }

  getPrivateStorageScope() {
    const unit = this.getActiveUnitId();
    const uid = this.state.authUser?.id || '';
    return uid ? `u_${unit}__usr_${uid}` : `u_${unit}__guest`;
  }

  _unitKey(baseKey) {
    return `${baseKey}__u_${this.getActiveUnitId()}`;
  }

  _projectKey(baseKey) {
    return `${baseKey}__${this.getPrivateStorageScope()}`;
  }

  _getChatStorageKeys() {
    const uid = this.state.authUser?.id || '';
    if (!uid) {
      return {
        sessions: CHAT_SESSIONS_LS_KEY,
        activeId: ACTIVE_CHAT_LS_KEY,
        autoApply: CHAT_AUTO_APPLY_LS_KEY
      };
    }
    const scope = this.getPrivateStorageScope();
    return {
      sessions: `ai_chat_sessions_v2__${scope}`,
      activeId: `ai_active_chat_id_v2__${scope}`,
      autoApply: `ai_chat_auto_apply_v2__${scope}`
    };
  }

  getState() {
    return this.state;
  }

  subscribe(listener) {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  notify() {
    for (const listener of this.listeners) {
      try {
        listener(this.state);
      } catch (err) {
        console.error('State listener error:', err);
      }
    }
  }

  setState(updater, recordHistory = false) {
    const nextState = typeof updater === 'function' ? updater(this.state) : updater;

    if (recordHistory && nextState.document && nextState.document !== this.state.document) {
      this.undoStack.push(JSON.parse(JSON.stringify(this.state.document)));
      if (this.undoStack.length > this.maxHistory) {
        this.undoStack.shift();
      }
      this.redoStack = [];
    }

    this.state = { ...this.state, ...nextState };
    this.notify();

    if (nextState.document) {
      this.saveToIndexedDb(STORE_PROJECTS, this._projectKey('active_document'), this.state.document);
    }
    if (nextState.smartDocument !== undefined) {
      this.saveToIndexedDb(STORE_PROJECTS, this._projectKey('active_smart_document'), this.state.smartDocument);
    }
    if (nextState.schoolProfile) {
      this.saveToIndexedDb(STORE_SCHOOL, this._unitKey('profile'), this.state.schoolProfile);
    }
    if (nextState.agendaBook) {
      this.saveToIndexedDb(STORE_AGENDA, this._unitKey('agenda_list'), this.state.agendaBook);
    }
    if (nextState.apiVault) {
      this.saveToIndexedDb(STORE_API_VAULT, 'vault_list', this.state.apiVault);
    }
    if (nextState.kopGallery) {
      this.saveToIndexedDb(STORE_KOP_GALLERY, this._unitKey('gallery_list'), this.state.kopGallery);
    }

    const touchedPrivate = (
      nextState.document !== undefined ||
      nextState.smartDocument !== undefined ||
      nextState.chatSessions !== undefined ||
      nextState.activeChatId !== undefined ||
      nextState.chatAutoApply !== undefined ||
      nextState.letterOptions !== undefined ||
      nextState.activeVaultId !== undefined ||
      nextState.activeSmartVaultId !== undefined
    );
    if (touchedPrivate && !this.suppressCloudSync && this.state.authUser?.id && getStoredAuthToken()) {
      this.queuePrivateWorkspaceSync(false);
    }
  }

  // --- PEMBARUAN PROFIL SEKOLAH ---
  getSchoolProfile() {
    return this.state.schoolProfile || DEFAULT_SCHOOL_PROFILE;
  }

  getEffectiveSchoolForDocument(docOrItem = null) {
    const targetDoc = docOrItem || this.state.document || this.state.smartDocument?.pendingDocModel || this.state.smartDocument;
    return resolveDocumentSchoolProfile(targetDoc, this.getSchoolProfile(), this.state.kopGallery || []);
  }

  updateSchoolProfile(newProfile) {
    const updated = { ...this.state.schoolProfile, ...newProfile };
    const curDoc = this.state.document;
    let newDoc;
    if (!curDoc || curDoc.isBlankDraft) {
      newDoc = createEmptyDocumentModel(updated, this.getLetterOptions());
    } else {
      newDoc = {
        ...curDoc,
        blocks: (curDoc.blocks || []).map(b => (b.type === 'letterKop' ? { ...b, school: updated, image: updated.kopImage || b.image } : b))
      };
    }

    this.setState({
      schoolProfile: updated,
      document: newDoc
    }, true);
    this.queueCloudSyncPatch({ schoolProfile: updated });
  }

  saveSchoolProfile(profile) {
    this.saveToIndexedDb(STORE_SCHOOL, this._unitKey('profile'), profile);
  }

  saveKopGallery(gallery) {
    this.saveToIndexedDb(STORE_KOP_GALLERY, this._unitKey('gallery_list'), gallery);
  }

  setAuthSessionData({ user = null, permissions = null, driveStatus = null, users = null, unitId = '' } = {}) {
    const patch = {};
    if (user !== undefined) patch.authUser = user;
    const resolvedUnit = unitId
      || (user?.unitId === 'smp' ? 'smp' : (user?.unitId === 'smk' ? 'smk' : (user?.defaultUnit || '')));
    if (resolvedUnit) {
      const cleanUnit = normalizeUnitId(resolvedUnit);
      patch.activeUnitId = cleanUnit;
      setActiveWorkspaceUnitId(cleanUnit);
    }
    if (permissions) patch.authPermissions = { ...DEFAULT_ACCESS_PERMISSIONS, ...permissions };
    if (driveStatus) patch.driveStatus = { ...this.state.driveStatus, ...driveStatus };
    if (Array.isArray(users)) patch.managedUsers = users;
    this.setState(patch, false);
  }

  clearAuthState() {
    clearStoredAuthToken();
    this.cloudSyncQueue = {};
    if (this.cloudSyncTimer) {
      clearTimeout(this.cloudSyncTimer);
      this.cloudSyncTimer = null;
    }
    if (this.privateSyncTimer) {
      clearTimeout(this.privateSyncTimer);
      this.privateSyncTimer = null;
    }
    this.undoStack = [];
    this.redoStack = [];
    const defaultSchool = getDefaultSchoolProfileForUnit(this.getActiveUnitId());
    this.suppressCloudSync = true;
    try {
      this.setState({
        authUser: null,
        authPermissions: { ...DEFAULT_ACCESS_PERMISSIONS },
        managedUsers: [],
        cloudSyncState: 'local',
        privateSyncState: 'local',
        chatSessions: [],
        activeChatId: '',
        chatAutoApply: false,
        smartDocument: null,
        activeUploadedDocx: null,
        activeVaultId: null,
        activeSmartVaultId: null,
        outgoingLetters: [],
        incomingLetters: [],
        pendingKepsekRevision: null,
        document: createEmptyDocumentModel(defaultSchool, this.getLetterOptions())
      }, false);
    } finally {
      this.suppressCloudSync = false;
    }
  }

  mergeApiVaultWithRemote(localVault = [], remoteVaultObj = null) {
    const remoteItems = Array.isArray(remoteVaultObj?.items) ? remoteVaultObj.items : [];
    const byId = new Map();
    for (const item of localVault) {
      if (item?.id) byId.set(item.id, { ...item });
    }
    for (const item of remoteItems) {
      if (!item?.id) continue;
      const prev = byId.get(item.id);
      if (!prev) {
        byId.set(item.id, { ...item });
      } else {
        const prevTs = Number(prev.updatedAt || prev.createdAt || 0);
        const remoteTs = Number(item.updatedAt || item.createdAt || 0);
        if (remoteTs >= prevTs || !prev.apiKey) {
          byId.set(item.id, { ...prev, ...item });
        }
      }
    }
    const hydrated = Array.from(byId.values()).map(entry => {
      if (!(entry.apiKey || '').trim()) {
        const builtinKey = getBuiltinSystemKey(entry.id);
        if (builtinKey) return { ...entry, apiKey: builtinKey };
      }
      return entry;
    });
    const activeId = remoteVaultObj?.activeApiId || hydrated.find(v => v.isActive)?.id || hydrated[0]?.id || 'gemini-auto';
    return {
      activeId,
      items: hydrated.map(v => ({ ...v, isActive: v.id === activeId }))
    };
  }

  applyRemoteBootstrap(payload = {}) {
    if (!payload || !payload.ok) return;
    this.suppressCloudSync = true;
    this.undoStack = [];
    this.redoStack = [];
    try {
      const ws = payload.workspace || {};
      const resolvedUser = payload.user || this.state.authUser;
      const unitId = normalizeUnitId(
        payload.unitId || ws.unitId || (resolvedUser?.unitId === 'smp' ? 'smp' : (resolvedUser?.defaultUnit || this.state.activeUnitId || 'smk'))
      );
      setActiveWorkspaceUnitId(unitId);

      const defaultSchool = getDefaultSchoolProfileForUnit(unitId);
      const normalizedSchool = normalizeSchoolProfileNaming({
        ...defaultSchool,
        ...(ws.schoolProfile || {}),
        unitId
      });

      const resolvedGallery = (Array.isArray(ws.kopGallery) && ws.kopGallery.length > 0)
        ? ws.kopGallery.map(k => (unitId === 'smk' ? {
            ...k,
            name: normalizeSmkText(k.name),
            description: normalizeSmkText(k.description),
            dataUrl: typeof k.dataUrl === 'string' && k.dataUrl.startsWith('data:image/svg+xml')
              ? k.dataUrl.replace(/SMKS(?=(%20|\s))/g, 'SMK')
              : k.dataUrl
          } : { ...k }))
        : getDefaultKopGalleryForUnit(unitId);

      const resolvedVault = (Array.isArray(ws.letterVault) && ws.letterVault.length > 0)
        ? ws.letterVault
        : buildDefaultVaultForUnit(unitId);

      // Pulihkan Ruang Kerja Pribadi & Chat AI eksklusif untuk akun + unit ini saja
      const priv = payload.privateWorkspace && typeof payload.privateWorkspace === 'object'
        ? payload.privateWorkspace
        : null;

      let scopedOpts = priv?.letterOptions
        ? normalizePageSetupOptions(priv.letterOptions)
        : null;
      let scopedChatSessions = Array.isArray(priv?.chatSessions) ? priv.chatSessions : null;
      let scopedActiveChatId = priv?.activeChatId !== undefined ? String(priv.activeChatId || '') : null;
      let scopedAutoApply = priv?.chatAutoApply !== undefined ? Boolean(priv.chatAutoApply) : null;

      if (typeof localStorage !== 'undefined' && resolvedUser?.id) {
        const scope = `u_${unitId}__usr_${resolvedUser.id}`;
        try {
          if (!scopedOpts) {
            const rawOpt = localStorage.getItem(`letter_options__${scope}`);
            if (rawOpt) scopedOpts = normalizePageSetupOptions(JSON.parse(rawOpt));
          }
          if (!scopedChatSessions) {
            const rawChats = localStorage.getItem(`ai_chat_sessions_v2__${scope}`);
            if (rawChats) {
              const parsedChats = JSON.parse(rawChats);
              if (Array.isArray(parsedChats)) scopedChatSessions = parsedChats;
            }
          }
          if (scopedActiveChatId === null) {
            scopedActiveChatId = localStorage.getItem(`ai_active_chat_id_v2__${scope}`) || '';
          }
          if (scopedAutoApply === null) {
            scopedAutoApply = localStorage.getItem(`ai_chat_auto_apply_v2__${scope}`) === 'true';
          }
        } catch {}
      }

      const finalOpts = scopedOpts || {
        paperSize: 'a4',
        orientation: 'portrait',
        margins: { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 },
        fontFamily: 'Times New Roman',
        fontSize: '12',
        lineSpacing: '1.15',
        paragraphSpacing: '6',
        showKop: true,
        pageNumbering: { ...PAGE_NUMBERING_DEFAULT }
      };

      const nextDoc = priv?.document && typeof priv.document === 'object'
        ? (!priv.document.isBlankDraft
            ? {
                ...priv.document,
                blocks: (priv.document.blocks || []).map(b => (
                  b?.type === 'letterKop' ? { ...b, school: normalizedSchool, image: normalizedSchool.kopImage || b.image } : b
                ))
              }
            : createEmptyDocumentModel(normalizedSchool, finalOpts))
        : createEmptyDocumentModel(normalizedSchool, finalOpts);

      const nextPatch = {
        authUser: resolvedUser,
        activeUnitId: unitId,
        authPermissions: payload.permissions ? { ...DEFAULT_ACCESS_PERMISSIONS, ...payload.permissions } : this.state.authPermissions,
        driveStatus: payload.driveStatus ? { ...this.state.driveStatus, ...payload.driveStatus } : this.state.driveStatus,
        cloudSyncState: 'synced',
        privateSyncState: priv ? 'synced' : 'local',
        cloudLastSyncAt: payload.syncedAt || new Date().toISOString(),
        syncRev: Number(payload.syncRev || payload.rev) || this.state.syncRev || 1,
        schoolProfile: normalizedSchool,
        kopGallery: resolvedGallery,
        letterVault: resolvedVault,
        agendaBook: Array.isArray(ws.agendaBook) ? ws.agendaBook : [],
        outgoingLetters: Array.isArray(ws.outgoingLetters) ? ws.outgoingLetters : [],
        incomingLetters: Array.isArray(ws.incomingLetters) ? ws.incomingLetters : [],
        letterOptions: finalOpts,
        document: nextDoc,
        smartDocument: priv?.smartDocument || null,
        activeUploadedDocx: null,
        activeVaultId: priv?.activeVaultId || null,
        activeSmartVaultId: priv?.activeSmartVaultId || null,
        pendingKepsekRevision: null,
        chatSessions: scopedChatSessions || [],
        activeChatId: scopedActiveChatId || '',
        chatAutoApply: Boolean(scopedAutoApply)
      };
      if (Array.isArray(payload.users)) {
        nextPatch.managedUsers = payload.users;
      }
      if (ws.apiVault?.items) {
        const mergedApi = this.mergeApiVaultWithRemote(this.state.apiVault, ws.apiVault);
        nextPatch.apiVault = mergedApi.items;
        nextPatch.activeApiId = mergedApi.activeId;
      }

      this.setState(nextPatch, false);
      this.saveOutgoingLetters(nextPatch.outgoingLetters);
      this.saveIncomingLetters(nextPatch.incomingLetters);
      this.saveLetterVault(nextPatch.letterVault);
      this._persistChatSessions(nextPatch.chatSessions, nextPatch.activeChatId);
      if (typeof localStorage !== 'undefined' && resolvedUser?.id) {
        try {
          const scope = `u_${unitId}__usr_${resolvedUser.id}`;
          localStorage.setItem(`letter_options__${scope}`, JSON.stringify(finalOpts));
        } catch {}
      }
      this.syncActiveDocumentNumberToNextOutgoing(false);

      // Bila cloud pribadi masih kosong, cek cadangan lokal akun+unit ini di IndexedDB
      if (!priv && resolvedUser?.id) {
        this._restorePrivateWorkspaceFromIdbIfEmpty(unitId, resolvedUser.id, normalizedSchool, finalOpts);
      }
    } finally {
      this.suppressCloudSync = false;
    }
  }

  /**
   * Menerapkan perubahan ringan ACC Kepala Sekolah / TU saja tanpa menyentuh ruang kerja, kanvas, profil, maupun chat AI user lain.
   */
  applyLiveWorkspaceDelta(payload = {}) {
    if (!payload || !payload.ok || !payload.changed) return false;
    if (payload.unitId && normalizeUnitId(payload.unitId) !== this.getActiveUnitId()) {
      return false;
    }
    this.suppressCloudSync = true;
    try {
      const ws = payload.workspace || {};
      const nextPatch = {
        syncRev: Number(payload.rev || payload.syncRev) || this.state.syncRev || 1,
        cloudSyncState: 'synced',
        cloudLastSyncAt: payload.syncedAt || new Date().toISOString()
      };

      if (payload.permissions) {
        nextPatch.authPermissions = { ...DEFAULT_ACCESS_PERMISSIONS, ...payload.permissions };
      }
      if (Array.isArray(ws.outgoingLetters)) {
        nextPatch.outgoingLetters = ws.outgoingLetters;
      }
      if (Array.isArray(ws.incomingLetters)) {
        nextPatch.incomingLetters = ws.incomingLetters;
      }

      this.setState(nextPatch, false);

      if (nextPatch.incomingLetters) {
        this.saveIncomingLetters(nextPatch.incomingLetters);
      }

      if (nextPatch.outgoingLetters) {
        this.saveOutgoingLetters(nextPatch.outgoingLetters);
        let verificationUpdated = false;
        const stdOutId = this.state.document?.verification?.outgoingId;
        if (stdOutId) {
          const matchedStd = nextPatch.outgoingLetters.find(x => x.id === stdOutId);
          if (matchedStd && (
            matchedStd.verificationStatus !== this.state.document.verification.verificationStatus ||
            matchedStd.qrDataUrl !== this.state.document.verification.qrDataUrl
          )) {
            this.state.document.verification = {
              ...this.state.document.verification,
              isVerifiedSah: Boolean(matchedStd.isVerifiedSah),
              verificationStatus: matchedStd.verificationStatus,
              accStatus: matchedStd.accStatus || '',
              verificationId: matchedStd.verificationId || '',
              verificationUrl: matchedStd.verificationUrl || '',
              qrDataUrl: matchedStd.qrDataUrl || '',
              rejectedReason: matchedStd.rejectedReason || ''
            };
            verificationUpdated = true;
          }
        }
        const smartOutId = this.state.smartDocument?.verification?.outgoingId;
        if (smartOutId) {
          const matchedSmart = nextPatch.outgoingLetters.find(x => x.id === smartOutId);
          if (matchedSmart && (
            matchedSmart.verificationStatus !== this.state.smartDocument.verification.verificationStatus ||
            matchedSmart.qrDataUrl !== this.state.smartDocument.verification.qrDataUrl
          )) {
            this.state.smartDocument.verification = {
              ...this.state.smartDocument.verification,
              isVerifiedSah: Boolean(matchedSmart.isVerifiedSah),
              verificationStatus: matchedSmart.verificationStatus,
              accStatus: matchedSmart.accStatus || '',
              verificationId: matchedSmart.verificationId || '',
              verificationUrl: matchedSmart.verificationUrl || '',
              qrDataUrl: matchedSmart.qrDataUrl || '',
              rejectedReason: matchedSmart.rejectedReason || ''
            };
            verificationUpdated = true;
          }
        }
        if (verificationUpdated && typeof window !== 'undefined' && typeof CustomEvent !== 'undefined') {
          window.dispatchEvent(new CustomEvent('docbuilder:verification-updated'));
        }
      }
      return true;
    } finally {
      this.suppressCloudSync = false;
    }
  }

  upsertOutgoingLetterFromCloud(item) {
    if (!item?.id) return item;
    if (item.unitId && normalizeUnitId(item.unitId) !== this.getActiveUnitId()) {
      return item;
    }
    const list = this.getOutgoingLetters();
    const exists = list.some(x => x.id === item.id);
    const updated = (exists ? list.map(x => (x.id === item.id ? item : x)) : [item, ...list]).sort((a, b) => {
      const yDiff = Number(b.year || 0) - Number(a.year || 0);
      if (yDiff !== 0) return yDiff;
      return (Number(b.sequenceNumber) || 0) - (Number(a.sequenceNumber) || 0);
    });
    this.setState({ outgoingLetters: updated }, false);
    this.saveOutgoingLetters(updated);
    return item;
  }

  buildPrivateWorkspaceSnapshot() {
    const cleanChats = (Array.isArray(this.state.chatSessions) ? this.state.chatSessions : []).slice(0, 60).map((s) => ({
      ...s,
      messages: (s.messages || []).map((m) => ({
        ...m,
        attachments: (m.attachments || []).map(({ base64Data, dataUrl, ...meta }) => meta)
      }))
    }));
    return {
      unitId: this.getActiveUnitId(),
      document: this.state.document || null,
      smartDocument: this.state.smartDocument || null,
      currentLetterForm: this.state.currentLetterForm || null,
      letterOptions: this.getLetterOptions(),
      activeVaultId: this.state.activeVaultId || null,
      activeSmartVaultId: this.state.activeSmartVaultId || null,
      chatSessions: cleanChats,
      activeChatId: this.state.activeChatId || '',
      chatAutoApply: Boolean(this.state.chatAutoApply),
      updatedAt: new Date().toISOString()
    };
  }

  queuePrivateWorkspaceSync(immediate = false) {
    if (this.suppressCloudSync || typeof window === 'undefined') return;
    if (!this.state.authUser?.id || !getStoredAuthToken()) return;
    if (this.privateSyncTimer) clearTimeout(this.privateSyncTimer);
    if (immediate) {
      this.flushPrivateWorkspaceNow();
      return;
    }
    this.privateSyncTimer = setTimeout(() => {
      this.flushPrivateWorkspaceNow();
    }, 1600);
  }

  async flushPrivateWorkspaceNow() {
    if (this.privateSyncTimer) {
      clearTimeout(this.privateSyncTimer);
      this.privateSyncTimer = null;
    }
    if (!this.state.authUser?.id || !getStoredAuthToken()) return false;
    const unitId = this.getActiveUnitId();
    const snapshot = this.buildPrivateWorkspaceSnapshot();
    this.setState({ privateSyncState: 'saving' }, false);
    const res = await savePrivateWorkspaceCloud(snapshot, unitId);
    if (res.ok) {
      this.suppressCloudSync = true;
      try {
        this.setState({ privateSyncState: 'saved' }, false);
      } finally {
        this.suppressCloudSync = false;
      }
      return true;
    }
    this.suppressCloudSync = true;
    try {
      this.setState({ privateSyncState: 'error' }, false);
    } finally {
      this.suppressCloudSync = false;
    }
    return false;
  }

  queueCloudSyncPatch(patch = {}, immediate = false) {
    if (this.suppressCloudSync || typeof window === 'undefined') return;
    if (!this.state.authUser || !getStoredAuthToken()) return;

    this.cloudSyncQueue = { ...this.cloudSyncQueue, ...patch };
    if (this.cloudSyncTimer) clearTimeout(this.cloudSyncTimer);
    if (immediate) {
      this.flushCloudSyncQueue();
      return;
    }
    this.cloudSyncTimer = setTimeout(() => {
      this.flushCloudSyncQueue();
    }, 350);
  }

  async flushCloudSyncQueue() {
    this.cloudSyncTimer = null;
    const payload = { ...this.cloudSyncQueue };
    this.cloudSyncQueue = {};
    if (Object.keys(payload).length === 0 || !getStoredAuthToken()) return;

    this.setState({ cloudSyncState: 'syncing' }, false);
    const res = await saveWorkspaceStatePatch({ ...payload, activeUnitId: this.getActiveUnitId() });
    if (res.ok) {
      this.setState({
        cloudSyncState: 'synced',
        cloudLastSyncAt: res.syncedAt || new Date().toISOString()
      }, false);
    } else {
      this.setState({ cloudSyncState: 'error' }, false);
    }
  }

  // --- PEMBARUAN FORM SURAT ---
  updateLetterForm(newFormData) {
    const updatedForm = { ...this.state.currentLetterForm, ...newFormData };
    const newDoc = this.regenerateActiveDocument(updatedForm, this.state.schoolProfile, this.state.letterOptions);

    this.setState({
      currentLetterForm: updatedForm,
      document: newDoc
    }, false);
  }

  getLetterOptions() {
    const cur = this.state.letterOptions || {};
    return {
      paperSize: cur.paperSize || 'a4',
      orientation: cur.orientation || 'portrait',
      margins: cur.margins || { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 },
      fontFamily: cur.fontFamily || 'Times New Roman',
      fontSize: String(cur.fontSize || '12'),
      lineSpacing: String(cur.lineSpacing || '1.15'),
      paragraphSpacing: String(cur.paragraphSpacing || '6'),
      showKop: cur.showKop !== undefined ? Boolean(cur.showKop) : true,
      pageNumbering: cur.pageNumbering && typeof cur.pageNumbering === 'object'
        ? { ...PAGE_NUMBERING_DEFAULT, ...cur.pageNumbering }
        : { ...PAGE_NUMBERING_DEFAULT }
    };
  }

  updateLetterOptions(opts) {
    const updatedOpts = normalizePageSetupOptions(opts, this.getLetterOptions());
    let currentDoc = this.state.document;
    if (currentDoc) {
      currentDoc = {
        ...currentDoc,
        profile: updatedOpts.paperSize === 'f4' ? 'suratDinasF4' : 'suratDinas',
        customMargins: updatedOpts.margins,
        paperSize: updatedOpts.paperSize,
        orientation: updatedOpts.orientation || 'portrait',
        pageSetup: updatedOpts
      };
    } else {
      currentDoc = this.regenerateActiveDocument(this.state.currentLetterForm, this.state.schoolProfile, updatedOpts);
    }

    this.setState({
      letterOptions: updatedOpts,
      document: currentDoc
    }, true);

    if (typeof localStorage !== 'undefined') {
      try {
        const uid = this.state.authUser?.id || '';
        const optKey = uid ? `letter_options__${this.getPrivateStorageScope()}` : 'letter_options';
        localStorage.setItem(optKey, JSON.stringify(updatedOpts));
      } catch (e) {}
    }
  }

  regenerateActiveDocument(formData, schoolProfile, options) {
    return suratKeteranganTemplate.generateDocumentModel(formData, schoolProfile, options);
  }

  // --- ZOOM LEMBAR KERTAS ---
  setZoomLevel(zoom, { fromUser = true } = {}) {
    const clamped = Math.max(50, Math.min(200, zoom));
    this.setState({
      zoomLevel: clamped,
      ...(fromUser ? { zoomUserAdjusted: true } : {})
    }, false);
  }

  // --- GUDANG API (API VAULT) METHODS ---
  getApiVault() {
    return this.state.apiVault || [];
  }

  getActiveApi() {
    const vault = this.getApiVault();
    const active = vault.find(a => a.id === this.state.activeApiId) || vault.find(a => a.isActive) || vault[0];
    return active;
  }

  setActiveApi(id) {
    const now = Date.now();
    const vault = this.getApiVault().map(item => ({
      ...item,
      isActive: item.id === id,
      updatedAt: item.id === id ? now : item.updatedAt
    }));
    this.setState({
      apiVault: vault,
      activeApiId: id
    });
    this.queueCloudSyncPatch({ apiVault: vault, activeApiId: id, replaceApiVault: true });
  }

  addApiVaultEntry(entry) {
    const now = Date.now();
    const newEntry = {
      id: entry.id || `api-${now}-${Math.random().toString(36).slice(2, 6)}`,
      label: entry.label || 'Koneksi AI Baru',
      provider: entry.provider || 'gemini',
      model: entry.model || 'gemini-2.0-flash',
      apiKey: entry.apiKey || '',
      baseUrl: entry.baseUrl || '',
      isActive: !!entry.isActive,
      createdAt: now,
      updatedAt: now
    };

    let vault = [...this.getApiVault()];
    if (newEntry.isActive) {
      vault = vault.map(v => ({ ...v, isActive: false }));
    }
    vault.push(newEntry);

    const nextActiveId = newEntry.isActive ? newEntry.id : this.state.activeApiId;
    this.setState({
      apiVault: vault,
      activeApiId: nextActiveId
    });
    this.queueCloudSyncPatch({ apiVault: vault, activeApiId: nextActiveId, replaceApiVault: true });
  }

  updateApiVaultEntry(id, patch) {
    const now = Date.now();
    const vault = this.getApiVault().map(item => {
      if (item.id === id) {
        return { ...item, ...patch, updatedAt: now };
      }
      if (patch.isActive && item.id !== id) {
        return { ...item, isActive: false };
      }
      return item;
    });

    const activeId = patch.isActive ? id : this.state.activeApiId;
    this.setState({ apiVault: vault, activeApiId: activeId });
    this.queueCloudSyncPatch({ apiVault: vault, activeApiId: activeId, replaceApiVault: true });
  }

  removeApiVaultEntry(id) {
    const vault = this.getApiVault().filter(item => item.id !== id);
    let activeId = this.state.activeApiId;
    if (activeId === id) {
      activeId = vault[0]?.id || '';
      if (vault[0]) vault[0].isActive = true;
    }
    this.setState({ apiVault: vault, activeApiId: activeId });
    this.queueCloudSyncPatch({ apiVault: vault, activeApiId: activeId, replaceApiVault: true });
  }

  // --- GUDANG KOP SURAT (KOP GALLERY) METHODS ---
  getKopGallery() {
    return this.state.kopGallery || [];
  }

  addKopGalleryEntry({ name, dataUrl, category = 'Kustom' }) {
    const newKop = {
      id: `kop-${Date.now()}`,
      name: name || `Kop Surat ${new Date().toLocaleDateString('id-ID')}`,
      category,
      isDefault: false,
      dataUrl,
      description: 'Diunggah oleh pengguna dan disimpan di galeri lokal.'
    };

    const gallery = [newKop, ...this.getKopGallery()];
    this.setState({ kopGallery: gallery });
    this.queueCloudSyncPatch({
      kopGallery: gallery,
      syncDriveItem: { type: 'kop', item: newKop }
    });
    return newKop;
  }

  removeKopGalleryEntry(id) {
    const gallery = this.getKopGallery().filter(k => k.id !== id);
    this.setState({ kopGallery: gallery });
    this.queueCloudSyncPatch({ kopGallery: gallery });
  }

  applyKopFromGallery(id) {
    const kop = this.getKopGallery().find(k => k.id === id);
    if (!kop) return;

    this.updateSchoolProfile({ kopImage: kop.dataUrl });
    this.updateLetterOptions({ showKop: true });
  }

  // --- GUDANG SURAT (LETTER VAULT) METHODS ---
  getLetterVault() {
    return this.state.letterVault || [];
  }

  buildVaultVersionSnapshot(entryLike = {}, note = 'Simpan Perubahan') {
    const actor = this.state.authUser?.displayName || this.state.authUser?.username || 'Pengguna';
    let docClone = null;
    if (entryLike.document && typeof entryLike.document === 'object') {
      try {
        docClone = JSON.parse(JSON.stringify(entryLike.document));
        if (docClone?.letterData?.kopImage && String(docClone.letterData.kopImage).length > 4000) {
          docClone.letterData.kopImage = '';
        }
      } catch {
        docClone = null;
      }
    }
    return {
      versionId: `ver-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      savedAt: new Date().toISOString(),
      savedBy: actor,
      note: String(note || 'Simpan Perubahan').trim().slice(0, 140),
      title: entryLike.title || 'Surat Dinas',
      nomorSurat: entryLike.nomorSurat || '-',
      perihal: entryLike.perihal || '-',
      document: docClone
    };
  }

  addLetterVaultEntry(entry) {
    const vault = this.getLetterVault();
    const initialHistory = Array.isArray(entry.versionHistory) && entry.versionHistory.length > 0
      ? entry.versionHistory.slice(0, 10)
      : (entry.document ? [this.buildVaultVersionSnapshot(entry, entry.versionNote || 'Versi Awal')] : []);
    const newEntry = {
      id: entry.id || `letter-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
      title: entry.title || 'Surat Dinas',
      nomorSurat: entry.nomorSurat || '-',
      tanggalSurat: entry.tanggalSurat || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
      category: entry.category || 'Surat Dinas',
      targetPerson: entry.targetPerson || '-',
      perihal: entry.perihal || '-',
      summary: entry.summary || '',
      isAiReference: entry.isAiReference !== undefined ? entry.isAiReference : true,
      isMasterTemplate: Boolean(entry.isMasterTemplate),
      isLockedTemplate: Boolean(entry.isLockedTemplate || entry.isMasterTemplate),
      source: entry.source || 'builder',
      document: entry.document || null,
      uploadedDocx: entry.uploadedDocx || null,
      versionHistory: initialHistory,
      createdAt: entry.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    const updated = [newEntry, ...vault];
    this.setState({ letterVault: updated }, false);
    this.saveLetterVault(updated);
    this.queueCloudSyncPatch({
      letterVault: updated,
      replaceLetterVault: true,
      syncDriveItem: { type: 'vault', item: newEntry }
    });
    return newEntry;
  }

  getActiveVaultId(mode = 'standard') {
    return mode === 'smart' ? (this.state.activeSmartVaultId || null) : (this.state.activeVaultId || null);
  }

  getActiveVaultEntry(mode = 'standard') {
    const id = this.getActiveVaultId(mode);
    if (!id) return null;
    return this.getLetterVault().find(item => item.id === id) || null;
  }

  findLetterVaultByTitle(title = '', excludeId = null) {
    const norm = String(title || '').trim().toLowerCase();
    if (!norm) return null;
    return this.getLetterVault().find(item =>
      item.id !== excludeId && String(item.title || '').trim().toLowerCase() === norm
    ) || null;
  }

  setActiveVaultId(id, mode = 'standard') {
    const patch = mode === 'smart'
      ? { activeSmartVaultId: id || null }
      : { activeVaultId: id || null };
    this.setState(patch, false);
    this.saveToIndexedDb(STORE_PROJECTS, this._projectKey('active_vault_refs'), {
      activeVaultId: this.state.activeVaultId || null,
      activeSmartVaultId: this.state.activeSmartVaultId || null
    });
  }

  updateLetterVaultEntry(id, updates, options = {}) {
    const vault = this.getLetterVault();
    const idx = vault.findIndex(item => item.id === id);
    if (idx === -1) return null;

    const currentEntry = vault[idx];
    const nextMerged = {
      ...currentEntry,
      ...updates,
      updatedAt: new Date().toISOString()
    };

    if (updates && updates.document) {
      const existingHistory = Array.isArray(currentEntry.versionHistory) ? [...currentEntry.versionHistory] : [];
      if (existingHistory.length === 0 && currentEntry.document) {
        existingHistory.push(this.buildVaultVersionSnapshot(currentEntry, 'Versi Sebelumnya'));
      }
      const note = options.versionNote || updates.versionNote || 'Pembaruan Dokumen (Ctrl+S)';
      const newSnap = this.buildVaultVersionSnapshot(nextMerged, note);
      const latestDocStr = existingHistory[0]?.document ? JSON.stringify(existingHistory[0].document) : '';
      const nextDocStr = newSnap.document ? JSON.stringify(newSnap.document) : '';
      if (options.forceVersionSnapshot || !latestDocStr || latestDocStr !== nextDocStr) {
        nextMerged.versionHistory = [newSnap, ...existingHistory].slice(0, 10);
      } else {
        nextMerged.versionHistory = existingHistory.slice(0, 10);
      }
      delete nextMerged.versionNote;
    }

    const updated = vault.map((item, i) => (i === idx ? nextMerged : item));
    this.setState({ letterVault: updated }, false);
    this.saveLetterVault(updated);
    this.queueCloudSyncPatch({
      letterVault: updated,
      replaceLetterVault: true,
      syncDriveItem: { type: 'vault', item: nextMerged }
    });
    return nextMerged;
  }

  restoreLetterVaultVersion(vaultId, versionId) {
    const vault = this.getLetterVault();
    const entry = vault.find(item => item.id === vaultId);
    if (!entry) return null;
    const history = Array.isArray(entry.versionHistory) ? entry.versionHistory : [];
    const targetVer = history.find(v => v.versionId === versionId);
    if (!targetVer || !targetVer.document) return null;

    const restoredDoc = JSON.parse(JSON.stringify(targetVer.document));
    const timeLabel = targetVer.savedAt
      ? new Date(targetVer.savedAt).toLocaleString('id-ID', { day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit' })
      : '';
    return this.updateLetterVaultEntry(
      vaultId,
      {
        title: targetVer.title || entry.title,
        nomorSurat: targetVer.nomorSurat || entry.nomorSurat,
        perihal: targetVer.perihal || entry.perihal,
        document: restoredDoc
      },
      {
        versionNote: `Pulihkan versi (${timeLabel || targetVer.note})`,
        forceVersionSnapshot: true
      }
    );
  }

  setPendingKepsekRevision(revision) {
    this.setState({ pendingKepsekRevision: revision ? { ...revision } : null }, false);
  }

  getPendingKepsekRevision() {
    return this.state.pendingKepsekRevision || null;
  }

  clearPendingKepsekRevision() {
    this.setState({ pendingKepsekRevision: null }, false);
  }

  removeLetterVaultEntry(id, { bypassLock = false } = {}) {
    const current = this.getLetterVault();
    const target = current.find(item => item.id === id);
    if (!target) return false;
    if ((target.isLockedTemplate || target.isMasterTemplate) && !bypassLock) {
      return false;
    }

    const vault = current.filter(item => item.id !== id);
    const patch = { letterVault: vault };
    let refsChanged = false;
    if (this.state.activeVaultId === id) {
      patch.activeVaultId = null;
      refsChanged = true;
    }
    if (this.state.activeSmartVaultId === id) {
      patch.activeSmartVaultId = null;
      refsChanged = true;
    }
    this.setState(patch, false);
    this.saveLetterVault(vault);
    if (refsChanged) {
      this.saveToIndexedDb(STORE_PROJECTS, 'active_vault_refs', {
        activeVaultId: this.state.activeVaultId || null,
        activeSmartVaultId: this.state.activeSmartVaultId || null
      });
    }
    this.queueCloudSyncPatch({ letterVault: vault, replaceLetterVault: true, deletedIds: [id] });
    return true;
  }

  toggleMasterTemplate(id) {
    const vault = this.getLetterVault();
    let newStatus = false;
    let updatedItem = null;
    const updated = vault.map(item => {
      if (item.id === id) {
        newStatus = !Boolean(item.isMasterTemplate || item.isLockedTemplate);
        updatedItem = {
          ...item,
          isMasterTemplate: newStatus,
          isLockedTemplate: newStatus,
          isAiReference: newStatus ? true : item.isAiReference,
          updatedAt: new Date().toISOString()
        };
        return updatedItem;
      }
      return item;
    });
    this.setState({ letterVault: updated }, false);
    this.saveLetterVault(updated);
    this.queueCloudSyncPatch({
      letterVault: updated,
      replaceLetterVault: true,
      ...(updatedItem ? { syncDriveItem: { type: 'vault', item: updatedItem } } : {})
    });
    return updatedItem;
  }

  toggleLetterAiReference(id) {
    const vault = this.getLetterVault();
    let newStatus = false;
    const updated = vault.map(item => {
      if (item.id === id) {
        newStatus = !item.isAiReference;
        return {
          ...item,
          isAiReference: newStatus,
          ...(newStatus ? {} : { isMasterTemplate: false, isLockedTemplate: false }),
          updatedAt: new Date().toISOString()
        };
      }
      return item;
    });
    this.setState({ letterVault: updated }, false);
    this.saveLetterVault(updated);
    this.queueCloudSyncPatch({ letterVault: updated, replaceLetterVault: true });
    return newStatus;
  }

  getAiReferences(limit = 4, queryPrompt = '') {
    return selectRelevantAiReferences(queryPrompt, this.getLetterVault(), { limit }).references;
  }

  getSmartAiReferencesForPrompt(userPrompt = '', limit = 3) {
    return selectRelevantAiReferences(userPrompt, this.getLetterVault(), { limit });
  }

  isSmartTemplateDocument(doc) {
    if (!doc || typeof doc !== 'object') return false;
    const tplId = doc.letterData?.templateId || '';
    if (['surat_tugas_sppd', 'smart_word_pro', 'smart_word_ai', 'sppd_only'].includes(tplId)) {
      return true;
    }
    const blocks = doc.blocks || [];
    return blocks.some(b =>
      b?.type === 'pageBreak' ||
      (b?.type === 'table' && (b.borderStyle === 'sppd' || b.borderStyle === 'none' || (b.rows || []).some(r =>
        (Array.isArray(r) ? r : r?.cells || []).some(c => c && typeof c === 'object' && (c.colspan > 1 || c.rowspan > 1 || Array.isArray(c.paragraphs)))
      )))
    );
  }

  loadLetterToCanvas(id, { targetUnitId = '' } = {}) {
    const vault = this.getLetterVault();
    const target = vault.find(v => v.id === id);
    if (!target) return false;

    const effectiveUnit = targetUnitId ? normalizeUnitId(targetUnitId) : this.getActiveUnitId();
    const targetSchool = resolveSchoolProfileForUnit(effectiveUnit, this.state.schoolProfile, this.state.kopGallery || []);

    if (target.uploadedDocx) {
      const clonedUpload = JSON.parse(JSON.stringify(target.uploadedDocx));
      if (targetUnitId) {
        if (clonedUpload.parsedMeta) {
          clonedUpload.parsedMeta = {
            ...clonedUpload.parsedMeta,
            title: adaptTextForSchoolUnit(clonedUpload.parsedMeta.title || '', effectiveUnit),
            nomor: adaptTextForSchoolUnit(clonedUpload.parsedMeta.nomor || '', effectiveUnit),
            signer: targetSchool.principalName
          };
        }
        if (Array.isArray(clonedUpload.paragraphs)) {
          clonedUpload.paragraphs = clonedUpload.paragraphs.map(p => ({
            ...p,
            currentText: adaptTextForSchoolUnit(p.currentText ?? p.originalText ?? '', effectiveUnit)
          }));
        }
        if (clonedUpload.structured?.signature) {
          clonedUpload.structured.signature = {
            ...clonedUpload.structured.signature,
            title: targetSchool.principalTitle || 'Kepala Sekolah',
            signer: targetSchool.principalName,
            nip: targetSchool.principalNip || '-'
          };
        }
        clonedUpload.targetUnitId = effectiveUnit;
        clonedUpload.targetSchoolProfile = targetSchool;
      }

      const nextDoc = target.document
        ? adaptDocumentModelForUnit(target.document, effectiveUnit, {
            activeSchoolProfile: this.state.schoolProfile,
            kopGallery: this.state.kopGallery || []
          })
        : (this.state.document && targetUnitId
            ? adaptDocumentModelForUnit(this.state.document, effectiveUnit, {
                activeSchoolProfile: this.state.schoolProfile,
                kopGallery: this.state.kopGallery || []
              })
            : this.state.document);

      this.setState({
        activeUploadedDocx: clonedUpload,
        activeVaultId: id,
        document: nextDoc
      }, true);
      this.saveToIndexedDb(STORE_PROJECTS, this._projectKey('active_vault_refs'), {
        activeVaultId: id,
        activeSmartVaultId: this.state.activeSmartVaultId || null
      });
      return true;
    }

    if (target.document) {
      const titleRaw = (target.document.blocks || []).find(b => b.type === 'letterTitle');
      const nextInfo = titleRaw
        ? this.getNextFormattedOutgoingNomor({
            title: titleRaw.title || target.title,
            perihal: target.perihal,
            currentNomor: titleRaw.nomor || target.nomorSurat,
            schoolProfile: targetSchool
          })
        : null;
      const cloned = adaptDocumentModelForUnit(target.document, effectiveUnit, {
        activeSchoolProfile: this.state.schoolProfile,
        kopGallery: this.state.kopGallery || [],
        nextNomorSurat: nextInfo?.nomorSurat || ''
      });
      const titleBlock = (cloned.blocks || []).find(b => b.type === 'letterTitle');
      if (target.preferSmartCanvas || this.isSmartTemplateDocument(cloned)) {
        this.setState({
          activeUploadedDocx: null,
          activeSmartVaultId: id,
          document: cloned,
          smartDocument: {
            pages: [],
            pendingDocModel: cloned,
            targetUnitId: effectiveUnit,
            schoolProfile: targetSchool,
            meta: {
              title: cloned.meta?.title || target.title || 'Dokumen Word Pro',
              nomorSurat: titleBlock ? (titleBlock.nomor || '').replace(/^Nomor\s*:\s*/i, '').trim() : (target.nomorSurat || '')
            }
          }
        }, true);
        this.saveToIndexedDb(STORE_PROJECTS, this._projectKey('active_vault_refs'), {
          activeVaultId: this.state.activeVaultId || null,
          activeSmartVaultId: id
        });
        return true;
      }
      this.setState({
        activeUploadedDocx: null,
        activeVaultId: id,
        document: cloned
      }, true);
      this.saveToIndexedDb(STORE_PROJECTS, this._projectKey('active_vault_refs'), {
        activeVaultId: id,
        activeSmartVaultId: this.state.activeSmartVaultId || null
      });
      return true;
    }

    // Jika surat tidak memiliki model docx terpisah, adaptasi ke dokumen kanvas saat ini dengan identitas & Kepsek unit terpilih
    const adaptedDoc = adaptDocumentModelForUnit(
      JSON.parse(JSON.stringify(this.state.document || createEmptyDocumentModel(targetSchool, this.getLetterOptions()))),
      effectiveUnit,
      {
        activeSchoolProfile: this.state.schoolProfile,
        kopGallery: this.state.kopGallery || [],
        nextNomorSurat: this.getNextFormattedOutgoingNomor({
          title: target.title,
          perihal: target.perihal,
          currentNomor: target.nomorSurat,
          schoolProfile: targetSchool
        }).nomorSurat
      }
    );
    if (!adaptedDoc.metadata) adaptedDoc.metadata = {};
    adaptedDoc.metadata.title = adaptTextForSchoolUnit(target.title || 'SURAT DINAS', effectiveUnit);
    adaptedDoc.metadata.nomorSurat = adaptTextForSchoolUnit(target.nomorSurat || '', effectiveUnit);
    this.setState({
      activeUploadedDocx: null,
      activeVaultId: id,
      document: adaptedDoc
    }, true);
    this.saveToIndexedDb(STORE_PROJECTS, this._projectKey('active_vault_refs'), {
      activeVaultId: id,
      activeSmartVaultId: this.state.activeSmartVaultId || null
    });
    return true;
  }

  setActiveUploadedDocx(docxData) {
    this.setState({ activeUploadedDocx: docxData }, false);
  }

  clearActiveUploadedDocx() {
    this.setState({ activeUploadedDocx: null }, false);
  }

  saveLetterVault(vaultList) {
    this.saveToIndexedDb(STORE_LETTER_VAULT, this._unitKey('vault_list'), vaultList);
  }

  // --- SURAT KELUAR (ARSIP TETAP / TERKUNCI DENGAN PENOMORAN BERURUTAN) ---
  getOutgoingLetters() {
    return this.state.outgoingLetters || [];
  }

  getNextOutgoingSequence(year = new Date().getFullYear()) {
    const list = this.getOutgoingLetters();
    const targetYear = String(year);
    const sameYear = list.filter(item => String(item.year || '') === targetYear);
    if (sameYear.length === 0) return 1;
    const maxSeq = sameYear.reduce((max, item) => Math.max(max, Number(item.sequenceNumber) || 0), 0);
    return maxSeq + 1;
  }

  getNextFormattedOutgoingNomor({ title = '', perihal = '', currentNomor = '', forcedTypeCode = '', date = new Date(), schoolProfile = null, unitId = '' } = {}) {
    const d = date instanceof Date && !isNaN(date.getTime()) ? date : new Date();
    const year = d.getFullYear().toString();
    const seqNum = this.getNextOutgoingSequence(year);
    const typeCode = (forcedTypeCode || detectLetterTypeCode(`${title} ${perihal} ${currentNomor}`)).toUpperCase();
    const effectiveProfile = schoolProfile || this.state.schoolProfile;
    const schoolCode = getSchoolCodeFromProfile(effectiveProfile, unitId || effectiveProfile?.unitId || this.getActiveUnitId());
    return {
      sequenceNumber: seqNum,
      typeCode,
      year,
      schoolCode,
      nomorSurat: buildOrderedNomorSurat({ typeCode, sequenceNumber: seqNum, schoolCode, date: d })
    };
  }

  /**
   * Menyinkronkan nomor surat pada dokumen kanvas aktif dengan nomor urut Surat Keluar berikutnya
   */
  syncActiveDocumentNumberToNextOutgoing(forceAdvance = false) {
    const doc = this.state.document;
    if (!doc || !Array.isArray(doc.blocks)) return null;
    const titleBlock = doc.blocks.find(b => b.type === 'letterTitle');
    if (!titleBlock) return null;

    const usedNomors = new Set(
      this.getOutgoingLetters().map(item => String(item.nomorSurat || '').replace(/^Nomor:\s*/i, '').trim())
    );
    const curClean = String(titleBlock.nomor || '').replace(/^Nomor:\s*/i, '').trim();
    const isAlreadyUsed = curClean && usedNomors.has(curClean);
    const isLegacyDefault = /\/\s*(085|050)\s*\//i.test(curClean);
    const isSeq001WhenHasOutgoing = this.getOutgoingLetters().length > 0 && /\/\s*001\s*\//i.test(curClean);

    if (forceAdvance || isAlreadyUsed || isLegacyDefault || isSeq001WhenHasOutgoing || !curClean) {
      const nextInfo = this.getNextFormattedOutgoingNomor({
        title: titleBlock.title,
        currentNomor: curClean
      });
      const updatedDoc = JSON.parse(JSON.stringify(doc));
      const tb = updatedDoc.blocks.find(b => b.type === 'letterTitle');
      if (tb) tb.nomor = `Nomor: ${nextInfo.nomorSurat}`;

      this.setState({
        currentLetterForm: { ...this.state.currentLetterForm, nomorSurat: nextInfo.nomorSurat },
        document: updatedDoc
      }, false);
      return nextInfo;
    }
    return null;
  }

  addOutgoingLetter(entry) {
    const list = this.getOutgoingLetters();
    const now = new Date();
    const year = String(entry.year || now.getFullYear());
    const seqNum = Number(entry.sequenceNumber) || this.getNextOutgoingSequence(year);
    const cleanNomor = String(entry.nomorSurat || '').replace(/^Nomor:\s*/i, '').trim()
      || this.getNextFormattedOutgoingNomor({
        title: entry.title,
        perihal: entry.perihal,
        forcedTypeCode: entry.typeCode,
        date: now
      }).nomorSurat;

    const lockedSnapshot = entry.document ? JSON.parse(JSON.stringify(entry.document)) : null;
    const lockedSmartPages = entry.smartPagesSnapshot ? JSON.parse(JSON.stringify(entry.smartPagesSnapshot)) : null;
    const lockedUploadedDocx = entry.uploadedDocx ? { ...entry.uploadedDocx } : null;

    const newEntry = {
      id: entry.id || `out-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      sequenceNumber: seqNum,
      year,
      typeCode: (entry.typeCode || detectLetterTypeCode(`${entry.title || ''} ${entry.perihal || ''} ${cleanNomor}`)).toUpperCase(),
      nomorSurat: cleanNomor,
      title: (entry.title || 'SURAT DINAS').trim(),
      recipientName: (entry.recipientName || '-').trim(),
      perihal: (entry.perihal || entry.title || '-').trim(),
      tanggalSurat: entry.tanggalSurat || now.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' }),
      summary: entry.summary || '',
      signerName: entry.signerName || this.state.schoolProfile?.principalName || 'Kepala Sekolah',
      signerRole: entry.signerRole || 'kepala_sekolah',
      signerTitle: entry.signerTitle || this.state.schoolProfile?.principalTitle || 'Kepala Sekolah',
      signerNip: entry.signerNip ?? this.state.schoolProfile?.principalNip ?? '-',
      signerRank: entry.signerRank || this.state.schoolProfile?.principalRank || '',
      verificationStatus: entry.verificationStatus || (entry.isVerifiedSah ? 'sah' : 'unverified'),
      isVerifiedSah: Boolean(entry.isVerifiedSah ?? (entry.verificationStatus === 'sah')),
      verificationId: entry.verificationId || '',
      verificationUrl: entry.verificationUrl || '',
      qrDataUrl: entry.qrDataUrl || '',
      manageToken: entry.manageToken || '',
      verificationStatement: entry.verificationStatement || '',
      verificationDetails: Array.isArray(entry.verificationDetails) ? entry.verificationDetails : [],
      cloudSynced: Boolean(entry.cloudSynced),
      authorizedAt: entry.authorizedAt || (entry.verificationStatus === 'sah' || entry.isVerifiedSah ? now.toISOString() : null),
      revokedAt: entry.revokedAt || null,
      revokedReason: entry.revokedReason || '',
      sourceMode: entry.sourceMode || 'standard',
      isLocked: true, // Surat keluar bersifat tetap & tidak dapat diedit kembali
      documentSnapshot: lockedSnapshot,
      ...(lockedSmartPages ? { smartPagesSnapshot: lockedSmartPages } : {}),
      uploadedDocx: lockedUploadedDocx,
      recordedAt: now.toISOString()
    };

    // Urutkan berdasarkan tahun terbaru lalu nomor urut terbesar di atas
    const updated = [newEntry, ...list].sort((a, b) => {
      const yDiff = Number(b.year || 0) - Number(a.year || 0);
      if (yDiff !== 0) return yDiff;
      return (Number(b.sequenceNumber) || 0) - (Number(a.sequenceNumber) || 0);
    });

    this.setState({ outgoingLetters: updated }, false);
    this.saveOutgoingLetters(updated);
    // Otomatis majukan nomor surat di Lembar Kerja ke nomor urut berikutnya (kecuali diminta tahan saat langsung cetak)
    if (!entry.skipAutoAdvanceNumber) {
      this.syncActiveDocumentNumberToNextOutgoing(true);
    }
    return newEntry;
  }

  /**
   * Memperbarui status verifikasi QR / sinkronisasi cloud pada catatan Surat Keluar
   * tanpa pernah membuka kunci isi surat keluar (isLocked tetap true).
   */
  updateOutgoingLetterVerification(id, updates = {}) {
    const list = this.getOutgoingLetters();
    let targetEntry = null;
    const updated = list.map(item => {
      if (item.id !== id) return item;
      const merged = {
        ...item,
        verificationStatus: updates.verificationStatus ?? item.verificationStatus,
        isVerifiedSah: updates.isVerifiedSah ?? ((updates.verificationStatus ?? item.verificationStatus) === 'sah'),
        verificationId: updates.verificationId ?? item.verificationId,
        verificationUrl: updates.verificationUrl ?? item.verificationUrl,
        qrDataUrl: updates.qrDataUrl ?? item.qrDataUrl,
        manageToken: updates.manageToken ?? item.manageToken,
        signerRole: updates.signerRole ?? item.signerRole,
        signerTitle: updates.signerTitle ?? item.signerTitle,
        signerName: updates.signerName ?? item.signerName,
        signerNip: updates.signerNip ?? item.signerNip,
        signerRank: updates.signerRank ?? item.signerRank,
        verificationStatement: updates.verificationStatement ?? item.verificationStatement,
        verificationDetails: updates.verificationDetails ?? item.verificationDetails,
        cloudSynced: updates.cloudSynced ?? item.cloudSynced,
        authorizedAt: updates.authorizedAt ?? item.authorizedAt,
        revokedAt: updates.revokedAt !== undefined ? updates.revokedAt : item.revokedAt,
        revokedReason: updates.revokedReason !== undefined ? updates.revokedReason : item.revokedReason
      };
      targetEntry = merged;
      return merged;
    });
    if (!targetEntry) return null;
    this.setState({ outgoingLetters: updated }, false);
    this.saveOutgoingLetters(updated);

    // Sinkronkan badge pengesahan pada kanvas aktif bila merujuk verificationId / nomor yang sama
    if (this.state.document?.verification?.outgoingId === id) {
      this.setActiveDocumentVerification({
        ...this.state.document.verification,
        ...updates,
        isVerifiedSah: targetEntry.isVerifiedSah,
        verificationStatus: targetEntry.verificationStatus
      }, 'standard');
    }
    if (this.state.smartDocument?.verification?.outgoingId === id) {
      this.setActiveDocumentVerification({
        ...this.state.smartDocument.verification,
        ...updates,
        isVerifiedSah: targetEntry.isVerifiedSah,
        verificationStatus: targetEntry.verificationStatus
      }, 'smart');
    }
    return targetEntry;
  }

  getOutgoingLetterByVerificationId(verificationId) {
    const id = String(verificationId || '').trim();
    if (!id) return null;
    return this.getOutgoingLetters().find(item => item.verificationId === id) || null;
  }

  /**
   * Menyematkan / menghapus metadata status SAH & QR Code pada dokumen kanvas aktif.
   */
  setActiveDocumentVerification(verificationMeta, mode = 'standard') {
    if (mode === 'smart') {
      const smartDoc = this.state.smartDocument;
      if (!smartDoc) return;
      const nextSmart = { ...smartDoc, verification: verificationMeta || null };
      this.setState({ smartDocument: nextSmart }, false);
    } else {
      const doc = this.state.document;
      if (!doc) return;
      const nextDoc = { ...doc, verification: verificationMeta || null };
      this.setState({ document: nextDoc }, false);
    }
    if (typeof window !== 'undefined' && typeof CustomEvent !== 'undefined') {
      window.dispatchEvent(new CustomEvent('docbuilder:verification-updated'));
    }
  }

  /**
   * Menyimpan konfigurasi posisi/lapisan/visibilitas Barcode QR mengambang pada kanvas aktif
   * serta menyinkronkannya ke arsip Surat Keluar terkait bila sudah tercatat.
   */
  setActiveQrPlacement(qrPlacement, mode = 'standard') {
    const cleanPlacement = qrPlacement ? { ...qrPlacement } : null;
    const isSmart = mode === 'smart';
    const targetDoc = isSmart ? this.state.smartDocument : this.state.document;
    if (!targetDoc) return;

    this.setState(
      isSmart
        ? { smartDocument: { ...targetDoc, qrPlacement: cleanPlacement } }
        : { document: { ...targetDoc, qrPlacement: cleanPlacement } },
      false
    );

    const outId = targetDoc.verification?.outgoingId;
    if (outId && cleanPlacement) {
      const list = this.getOutgoingLetters();
      if (list.some(item => item.id === outId)) {
        const updated = list.map(item => (item.id === outId ? {
          ...item,
          qrPlacement: cleanPlacement,
          ...(item.documentSnapshot ? { documentSnapshot: { ...item.documentSnapshot, qrPlacement: cleanPlacement } } : {})
        } : item));
        this.setState({ outgoingLetters: updated }, false);
        this.saveOutgoingLetters(updated);
      }
    }
  }

  removeOutgoingLetter(id) {
    const updated = this.getOutgoingLetters().filter(item => item.id !== id);
    this.setState({ outgoingLetters: updated }, false);
    this.saveOutgoingLetters(updated);
    return updated;
  }

  saveOutgoingLetters(list) {
    this.saveToIndexedDb(STORE_OUTGOING_LETTERS, this._unitKey('outgoing_list'), list);
  }

  /**
   * Menyalin isi surat keluar menjadi draf baru di kanvas tanpa pernah mengubah catatan arsip surat keluar aslinya
   */
  cloneOutgoingToNewDraft(id) {
    const target = this.getOutgoingLetters().find(item => item.id === id);
    if (!target || !target.documentSnapshot) return false;
    const clonedDoc = JSON.parse(JSON.stringify(target.documentSnapshot));
    delete clonedDoc.verification;
    delete clonedDoc.qrPlacement;
    for (const blk of (clonedDoc.blocks || [])) {
      if (blk?.type === 'letterSignature') {
        delete blk.signatureMode;
        delete blk.qrImage;
        delete blk.qrCaption;
        delete blk.verificationUrl;
      }
    }
    const nextInfo = this.getNextFormattedOutgoingNomor({
      title: target.title,
      perihal: target.perihal,
      forcedTypeCode: target.typeCode
    });
    const titleBlock = (clonedDoc.blocks || []).find(b => b.type === 'letterTitle');
    if (titleBlock) {
      titleBlock.nomor = `Nomor: ${nextInfo.nomorSurat}`;
    }
    this.setState({
      activeUploadedDocx: null,
      document: clonedDoc
    }, true);
    return true;
  }

  // --- SURAT MASUK (UPLOAD GAMBAR / PDF / FOTO KAMERA & SINKRON DRIVE) ---
  getIncomingLetters() {
    return Array.isArray(this.state.incomingLetters) ? this.state.incomingLetters : [];
  }

  getNextIncomingCodeInfo(date = new Date()) {
    const list = this.getIncomingLetters();
    const d = date instanceof Date && !Number.isNaN(date.getTime()) ? date : new Date();
    const year = String(d.getFullYear());
    const sequenceNumber = getNextIncomingSequence(list, year);
    return {
      sequenceNumber,
      year,
      kodeSurat: formatIncomingCode({ sequenceNumber, date: d })
    };
  }

  saveIncomingLetters(list) {
    this.saveToIndexedDb(STORE_INCOMING_LETTERS, this._unitKey('incoming_list'), Array.isArray(list) ? list : []);
  }

  upsertIncomingLetterLocal(entry = {}) {
    if (!entry?.id) return null;
    const list = this.getIncomingLetters();
    const exists = list.some(i => i.id === entry.id);
    const next = sortIncomingLetters(
      exists ? list.map(i => (i.id === entry.id ? { ...i, ...entry } : i)) : [entry, ...list],
      'newest'
    );
    this.setState({ incomingLetters: next }, false);
    this.saveIncomingLetters(next);
    return entry;
  }

  removeIncomingLetterLocal(id) {
    const next = this.getIncomingLetters().filter(i => i.id !== id);
    this.setState({ incomingLetters: next }, false);
    this.saveIncomingLetters(next);
    if (this.db && id) {
      try {
        const tx = this.db.transaction([STORE_INCOMING_FILES], 'readwrite');
        tx.objectStore(STORE_INCOMING_FILES).delete(id);
      } catch {}
    }
    return next;
  }

  saveIncomingFileCache(id, filePayload = {}) {
    if (!id || !filePayload?.fileBase64) return;
    this.saveToIndexedDb(STORE_INCOMING_FILES, id, {
      id,
      fileName: filePayload.fileName || '',
      mimeType: filePayload.mimeType || 'application/octet-stream',
      fileKind: filePayload.fileKind || 'image',
      fileBase64: filePayload.fileBase64,
      cachedAt: new Date().toISOString()
    });
  }

  getIncomingFileCache(id) {
    return new Promise((resolve) => {
      if (!this.db || !id) {
        resolve(null);
        return;
      }
      try {
        const tx = this.db.transaction([STORE_INCOMING_FILES], 'readonly');
        const req = tx.objectStore(STORE_INCOMING_FILES).get(id);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => resolve(null);
      } catch {
        resolve(null);
      }
    });
  }

  // --- PERCAKAPAN AI COPILOT MULTI-TURN (TERISOLASI PER AKUN + UNIT) ---
  _persistChatSessions(sessions = this.state.chatSessions, activeId = this.state.activeChatId) {
    const cleanList = Array.isArray(sessions) ? sessions.slice(0, 60) : [];
    const keys = this._getChatStorageKeys();
    this.saveToIndexedDb(STORE_CHAT_SESSIONS, this._projectKey('sessions_list'), cleanList);
    this.saveToIndexedDb(STORE_CHAT_SESSIONS, this._projectKey('active_chat_id'), activeId || '');
    if (typeof localStorage !== 'undefined') {
      try {
        // Simpan tanpa data base64 lampiran besar di localStorage agar kuota 5MB aman
        const compactForLs = cleanList.map((s) => ({
          ...s,
          messages: (s.messages || []).map((m) => ({
            ...m,
            attachments: (m.attachments || []).map(({ base64Data, dataUrl, ...meta }) => meta)
          }))
        }));
        localStorage.setItem(keys.sessions, JSON.stringify(compactForLs));
        localStorage.setItem(keys.activeId, activeId || '');
      } catch (e) {}
    }
  }

  getChatSessions() {
    return Array.isArray(this.state.chatSessions) ? this.state.chatSessions : [];
  }

  getActiveChatId() {
    return this.state.activeChatId || '';
  }

  getActiveChatSession() {
    const sessions = this.getChatSessions();
    let active = sessions.find(s => s.id === this.state.activeChatId && !s.trashed);
    if (!active) {
      active = sessions.find(s => !s.trashed && !s.archived);
    }
    if (!active) {
      return this.createChatSession();
    }
    if (active.id !== this.state.activeChatId) {
      this.setState({ activeChatId: active.id }, false);
      this._persistChatSessions(sessions, active.id);
    }
    return active;
  }

  createChatSession({ title = 'Percakapan Baru', targetWorkspace = 'smart-canvas', activeModelId = '' } = {}) {
    const sessions = this.getChatSessions();
    const newSession = createChatSession({
      title,
      targetWorkspace,
      activeModelId: activeModelId || this.state.activeApiId || ''
    });
    const nextList = [newSession, ...sessions].slice(0, 60);
    this.setState({ chatSessions: nextList, activeChatId: newSession.id }, false);
    this._persistChatSessions(nextList, newSession.id);
    return newSession;
  }

  setActiveChatSession(sessionId) {
    const exists = this.getChatSessions().find(s => s.id === sessionId);
    if (!exists) return null;
    this.setState({ activeChatId: sessionId }, false);
    this._persistChatSessions(this.getChatSessions(), sessionId);
    return exists;
  }

  updateChatSession(sessionId, patch = {}) {
    const sessions = this.getChatSessions();
    let updatedSession = null;
    const nextList = sessions.map((s) => {
      if (s.id !== sessionId) return s;
      updatedSession = {
        ...s,
        ...patch,
        updatedAt: new Date().toISOString()
      };
      return updatedSession;
    });
    if (!updatedSession) return null;
    this.setState({ chatSessions: nextList }, false);
    this._persistChatSessions(nextList, this.state.activeChatId);
    return updatedSession;
  }

  appendChatMessage(sessionId, messageInput = {}) {
    let sessions = this.getChatSessions();
    let target = sessions.find(s => s.id === sessionId);
    if (!target) {
      target = this.createChatSession({ targetWorkspace: 'smart-canvas' });
      sessions = this.getChatSessions();
    }
    const msg = createChatMessage(messageInput);
    const shouldRetitle = msg.role === 'user' && (target.isAutoTitle !== false || target.title === 'Percakapan Baru') && (target.messages || []).length === 0;
    const nextTitle = shouldRetitle ? deriveChatTitleFromMessage(msg.text) : target.title;
    const nowIso = new Date().toISOString();
    const nextList = sessions.map((s) => {
      if (s.id !== target.id) return s;
      return {
        ...s,
        title: nextTitle,
        updatedAt: nowIso,
        messages: [...(s.messages || []), msg]
      };
    });
    this.setState({ chatSessions: nextList, activeChatId: target.id }, false);
    this._persistChatSessions(nextList, target.id);
    return msg;
  }

  updateChatMessage(sessionId, messageId, patch = {}) {
    const sessions = this.getChatSessions();
    const nextList = sessions.map((s) => {
      if (s.id !== sessionId) return s;
      return {
        ...s,
        messages: (s.messages || []).map(m => (m.id === messageId ? { ...m, ...patch } : m))
      };
    });
    this.setState({ chatSessions: nextList }, false);
    this._persistChatSessions(nextList, this.state.activeChatId);
  }

  togglePinChatSession(sessionId) {
    const cur = this.getChatSessions().find(s => s.id === sessionId);
    if (!cur) return null;
    return this.updateChatSession(sessionId, { pinned: !cur.pinned, archived: false });
  }

  archiveChatSession(sessionId, archived = true) {
    return this.updateChatSession(sessionId, { archived: Boolean(archived), pinned: false, trashed: false });
  }

  trashChatSession(sessionId, trashed = true) {
    const updated = this.updateChatSession(sessionId, {
      trashed: Boolean(trashed),
      ...(trashed ? { pinned: false } : {})
    });
    if (trashed && this.state.activeChatId === sessionId) {
      const nextActive = this.getChatSessions().find(s => !s.trashed && !s.archived && s.id !== sessionId);
      if (nextActive) {
        this.setActiveChatSession(nextActive.id);
      }
    }
    return updated;
  }

  deleteChatSessionPermanent(sessionId) {
    const nextList = this.getChatSessions().filter(s => s.id !== sessionId);
    let nextActiveId = this.state.activeChatId;
    if (nextActiveId === sessionId) {
      nextActiveId = nextList.find(s => !s.trashed && !s.archived)?.id || '';
    }
    this.setState({ chatSessions: nextList, activeChatId: nextActiveId }, false);
    this._persistChatSessions(nextList, nextActiveId);
    return nextList;
  }

  getChatAutoApply() {
    return Boolean(this.state.chatAutoApply);
  }

  setChatAutoApply(enabled) {
    const val = Boolean(enabled);
    this.setState({ chatAutoApply: val }, false);
    if (typeof localStorage !== 'undefined') {
      try {
        const keys = this._getChatStorageKeys();
        localStorage.setItem(keys.autoApply, String(val));
      } catch (e) {}
    }
  }

  // --- UNDO / REDO ---
  canUndo() { return this.undoStack.length > 0; }
  canRedo() { return this.redoStack.length > 0; }

  undo() {
    if (!this.canUndo()) return;
    const prevDoc = this.undoStack.pop();
    this.redoStack.push(JSON.parse(JSON.stringify(this.state.document)));
    this.state = { ...this.state, document: prevDoc };
    this.notify();
    this.saveToIndexedDb(STORE_PROJECTS, this._projectKey('active_document'), prevDoc);
    this.queuePrivateWorkspaceSync(false);
  }

  redo() {
    if (!this.canRedo()) return;
    const nextDoc = this.redoStack.pop();
    this.undoStack.push(JSON.parse(JSON.stringify(this.state.document)));
    this.state = { ...this.state, document: nextDoc };
    this.notify();
    this.saveToIndexedDb(STORE_PROJECTS, this._projectKey('active_document'), nextDoc);
    this.queuePrivateWorkspaceSync(false);
  }

  // --- INDEXEDDB STORAGE ---
  initIndexedDb() {
    if (typeof indexedDB === 'undefined') return;

    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(STORE_PROJECTS)) {
        db.createObjectStore(STORE_PROJECTS);
      }
      if (!db.objectStoreNames.contains(STORE_SCHOOL)) {
        db.createObjectStore(STORE_SCHOOL);
      }
      if (!db.objectStoreNames.contains(STORE_AGENDA)) {
        db.createObjectStore(STORE_AGENDA);
      }
      if (!db.objectStoreNames.contains(STORE_API_VAULT)) {
        db.createObjectStore(STORE_API_VAULT);
      }
      if (!db.objectStoreNames.contains(STORE_KOP_GALLERY)) {
        db.createObjectStore(STORE_KOP_GALLERY);
      }
      if (!db.objectStoreNames.contains(STORE_LETTER_VAULT)) {
        db.createObjectStore(STORE_LETTER_VAULT);
      }
      if (!db.objectStoreNames.contains(STORE_OUTGOING_LETTERS)) {
        db.createObjectStore(STORE_OUTGOING_LETTERS);
      }
      if (!db.objectStoreNames.contains(STORE_INCOMING_LETTERS)) {
        db.createObjectStore(STORE_INCOMING_LETTERS);
      }
      if (!db.objectStoreNames.contains(STORE_INCOMING_FILES)) {
        db.createObjectStore(STORE_INCOMING_FILES);
      }
      if (!db.objectStoreNames.contains(STORE_CHAT_SESSIONS)) {
        db.createObjectStore(STORE_CHAT_SESSIONS);
      }
    };
    req.onsuccess = (e) => {
      this.db = e.target.result;
      this.restoreFromIndexedDb();
    };
    req.onerror = (e) => {
      console.warn('IndexedDB gagal dibuka:', e);
    };
  }

  saveToIndexedDb(storeName, key, value) {
    if (!this.db) return;
    try {
      const tx = this.db.transaction([storeName], 'readwrite');
      const store = tx.objectStore(storeName);
      store.put(value, key);
    } catch (e) {
      console.warn(`Gagal menyimpan ke IndexedDB (${storeName}):`, e);
    }
  }

  readFromIndexedDb(storeName, key) {
    return new Promise((resolve) => {
      if (!this.db || !key) {
        resolve(null);
        return;
      }
      try {
        const tx = this.db.transaction([storeName], 'readonly');
        const req = tx.objectStore(storeName).get(key);
        req.onsuccess = () => resolve(req.result ?? null);
        req.onerror = () => resolve(null);
      } catch {
        resolve(null);
      }
    });
  }

  async _restorePrivateWorkspaceFromIdbIfEmpty(unitId, userId, schoolProfile, fallbackOpts) {
    if (!this.db || !unitId || !userId) return;
    const scope = `u_${normalizeUnitId(unitId)}__usr_${userId}`;
    if (this.getPrivateStorageScope() !== scope) return;

    const [docRec, smartRec, refsRec, chatsRec, activeChatRec] = await Promise.all([
      this.readFromIndexedDb(STORE_PROJECTS, `active_document__${scope}`),
      this.readFromIndexedDb(STORE_PROJECTS, `active_smart_document__${scope}`),
      this.readFromIndexedDb(STORE_PROJECTS, `active_vault_refs__${scope}`),
      this.readFromIndexedDb(STORE_CHAT_SESSIONS, `sessions_list__${scope}`),
      this.readFromIndexedDb(STORE_CHAT_SESSIONS, `active_chat_id__${scope}`)
    ]);

    if (this.getPrivateStorageScope() !== scope) return;

    const patch = {};
    if (docRec && typeof docRec === 'object' && !docRec.isBlankDraft) {
      patch.document = {
        ...docRec,
        blocks: (docRec.blocks || []).map(b => (
          b?.type === 'letterKop' ? { ...b, school: schoolProfile, image: schoolProfile.kopImage || b.image } : b
        ))
      };
    }
    if (smartRec && Array.isArray(smartRec.pages)) {
      patch.smartDocument = smartRec;
    }
    if (refsRec && typeof refsRec === 'object') {
      patch.activeVaultId = refsRec.activeVaultId || null;
      patch.activeSmartVaultId = refsRec.activeSmartVaultId || null;
    }
    if (Array.isArray(chatsRec) && chatsRec.length > 0 && (!this.state.chatSessions || this.state.chatSessions.length === 0)) {
      patch.chatSessions = chatsRec;
      if (typeof activeChatRec === 'string' && activeChatRec) {
        patch.activeChatId = activeChatRec;
      }
    }
    if (Object.keys(patch).length > 0) {
      this.suppressCloudSync = true;
      try {
        this.setState(patch, false);
      } finally {
        this.suppressCloudSync = false;
      }
    }
  }

  restoreFromIndexedDb() {
    if (!this.db) return;
    // Jangan menimpa sesi akun cloud yang sudah ter-bootstrap
    if (this.state.authUser?.id) return;

    const unitId = this.getActiveUnitId();
    const uKey = (base) => `${base}__u_${unitId}`;
    const pKey = (base) => `${base}__${this.getPrivateStorageScope()}`;

    // 1. Pulihkan Profil Sekolah (offline fallback)
    try {
      const txSchool = this.db.transaction([STORE_SCHOOL], 'readonly');
      const storeSchool = txSchool.objectStore(STORE_SCHOOL);
      const reqSchool = storeSchool.get(uKey('profile'));
      reqSchool.onsuccess = () => {
        if (this.state.authUser?.id) return;
        const found = reqSchool.result;
        if (found) {
          const loadedSchool = normalizeSchoolProfileNaming(found);
          this.setState({
            schoolProfile: loadedSchool,
            document: this.regenerateActiveDocument(this.state.currentLetterForm, loadedSchool, this.state.letterOptions)
          }, false);
        }
      };
    } catch (e) {}

    // 2. Pulihkan Agenda Book
    try {
      const txAgenda = this.db.transaction([STORE_AGENDA], 'readonly');
      const reqAgenda = txAgenda.objectStore(STORE_AGENDA).get(uKey('agenda_list'));
      reqAgenda.onsuccess = () => {
        if (!this.state.authUser?.id && Array.isArray(reqAgenda.result)) {
          this.setState({ agendaBook: reqAgenda.result }, false);
        }
      };
    } catch (e) {}

    // 3. Pulihkan Gudang API (API Vault)
    try {
      const txVault = this.db.transaction([STORE_API_VAULT], 'readonly');
      const reqVault = txVault.objectStore(STORE_API_VAULT).get('vault_list');
      reqVault.onsuccess = () => {
        if (this.state.authUser?.id) return;
        if (Array.isArray(reqVault.result) && reqVault.result.length > 0) {
          const migratedVault = reqVault.result.map(entry => {
            if (entry.provider === 'gemini') {
              const m = (entry.model || '').toLowerCase();
              if (!m || m === 'gemini-2.5-flash' || m.includes('omni') || m.includes('2.0-flash') || m.includes('2-flash') || m.includes('1.5-pro') || m === 'gemini-pro') {
                return {
                  ...entry,
                  model: 'gemini-3.5-flash-lite',
                  label: 'Google Gemini 3.5 Flash-Lite (Instan & 15 RPM)'
                };
              }
            } else if (entry.provider === 'groq') {
              const m = (entry.model || '').toLowerCase();
              if (!m || m.includes('llama-3.3') || m.includes('llama-3.1') || m.includes('mixtral')) {
                return {
                  ...entry,
                  model: 'openai/gpt-oss-120b',
                  label: 'Groq - GPT OSS 120B (Ultra Cepat ~0.2s)'
                };
              }
            }
            return entry;
          });

          let backfilled = false;
          const hydratedVault = migratedVault.map(entry => {
            if (!(entry.apiKey || '').trim()) {
              const builtinKey = getBuiltinSystemKey(entry.id);
              if (builtinKey) {
                backfilled = true;
                return { ...entry, apiKey: builtinKey };
              }
            }
            return entry;
          });

          const activeEntry = hydratedVault.find(v => v.isActive) || hydratedVault[0];
          this.setState({
            apiVault: hydratedVault,
            activeApiId: activeEntry?.id || this.state.activeApiId
          }, backfilled);
        }
      };
    } catch (e) {}

    // 4. Pulihkan Gudang Kop Surat (Kop Gallery)
    try {
      const txKop = this.db.transaction([STORE_KOP_GALLERY], 'readonly');
      const reqKop = txKop.objectStore(STORE_KOP_GALLERY).get(uKey('gallery_list'));
      reqKop.onsuccess = () => {
        if (!this.state.authUser?.id && Array.isArray(reqKop.result) && reqKop.result.length > 0) {
          this.setState({ kopGallery: reqKop.result }, false);
        }
      };
    } catch (e) {}

    // 5. Pulihkan Gudang Surat (Letter Vault)
    try {
      const txLetter = this.db.transaction([STORE_LETTER_VAULT], 'readonly');
      const reqLetter = txLetter.objectStore(STORE_LETTER_VAULT).get(uKey('vault_list'));
      reqLetter.onsuccess = () => {
        if (!this.state.authUser?.id && Array.isArray(reqLetter.result)) {
          this.setState({ letterVault: reqLetter.result }, false);
        }
      };
    } catch (e) {}

    // 6. Pulihkan Buku Catatan Surat Keluar
    try {
      const txOutgoing = this.db.transaction([STORE_OUTGOING_LETTERS], 'readonly');
      const reqOutgoing = txOutgoing.objectStore(STORE_OUTGOING_LETTERS).get(uKey('outgoing_list'));
      reqOutgoing.onsuccess = () => {
        if (!this.state.authUser?.id && Array.isArray(reqOutgoing.result)) {
          this.setState({ outgoingLetters: reqOutgoing.result }, false);
          this.syncActiveDocumentNumberToNextOutgoing(false);
        }
      };
    } catch (e) {}

    // 6b. Pulihkan Daftar Surat Masuk
    try {
      const txIncoming = this.db.transaction([STORE_INCOMING_LETTERS], 'readonly');
      const reqIncoming = txIncoming.objectStore(STORE_INCOMING_LETTERS).get(uKey('incoming_list'));
      reqIncoming.onsuccess = () => {
        if (!this.state.authUser?.id && Array.isArray(reqIncoming.result)) {
          this.setState({ incomingLetters: sortIncomingLetters(reqIncoming.result, 'newest') }, false);
        }
      };
    } catch (e) {}

    try {
      const txProjects = this.db.transaction([STORE_PROJECTS], 'readonly');
      const storeProjects = txProjects.objectStore(STORE_PROJECTS);
      const reqSmart = storeProjects.get(pKey('active_smart_document'));
      reqSmart.onsuccess = () => {
        if (!this.state.authUser?.id && reqSmart.result && Array.isArray(reqSmart.result.pages)) {
          this.setState({ smartDocument: reqSmart.result }, false);
        }
      };
      const reqRefs = storeProjects.get(pKey('active_vault_refs'));
      reqRefs.onsuccess = () => {
        if (!this.state.authUser?.id && reqRefs.result && typeof reqRefs.result === 'object') {
          this.setState({
            activeVaultId: reqRefs.result.activeVaultId || null,
            activeSmartVaultId: reqRefs.result.activeSmartVaultId || null
          }, false);
        }
      };
    } catch (e) {}

    // 7. Pulihkan Riwayat Percakapan AI
    try {
      const txChat = this.db.transaction([STORE_CHAT_SESSIONS], 'readonly');
      const storeChat = txChat.objectStore(STORE_CHAT_SESSIONS);
      const reqList = storeChat.get(pKey('sessions_list'));
      const reqActive = storeChat.get(pKey('active_chat_id'));
      reqList.onsuccess = () => {
        if (!this.state.authUser?.id && Array.isArray(reqList.result) && reqList.result.length > 0) {
          this.setState({ chatSessions: reqList.result }, false);
        }
      };
      reqActive.onsuccess = () => {
        if (!this.state.authUser?.id && typeof reqActive.result === 'string' && reqActive.result) {
          this.setState({ activeChatId: reqActive.result }, false);
        }
      };
    } catch (e) {}
  }
}

export const store = new StateStore();
