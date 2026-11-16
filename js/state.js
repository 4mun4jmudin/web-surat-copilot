/**
 * state.js - Reactive State Store & IndexedDB Persistence
 * Mendukung penyimpanan Profil Sekolah, Form Surat Aktif, Gudang API Multi-Provider,
 * Gudang Galeri Kop Surat, Skala Zoom Kanvas, dan Riwayat Undo/Redo.
 */

import { DEFAULT_SCHOOL_PROFILE } from './letters/schoolProfile.js';
import { suratKeteranganTemplate } from './letters/templates/suratKeterangan.js';
import { DEFAULT_KOP_GALLERY, DEFAULT_KOP_DATA_URL } from './letters/defaultKop.js';
import { detectLetterTypeCode, formatSequenceNumber, toRomanMonth, buildOrderedNomorSurat, LETTER_TYPE_CODES } from './letters/numbering.js';

const DB_NAME = 'ai_school_doc_builder_db';
const DB_VERSION = 5;
const STORE_PROJECTS = 'projects';
const STORE_SCHOOL = 'school_profile';
const STORE_AGENDA = 'agenda_book';
const STORE_API_VAULT = 'api_vault';
const STORE_KOP_GALLERY = 'kop_gallery';
const STORE_LETTER_VAULT = 'letter_vault';
const STORE_OUTGOING_LETTERS = 'outgoing_letters';

// Nilai awal form surat keterangan (SMKS IT Al-Hawari)
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
      author: schoolProfile?.schoolName || 'SMKS IT Al-Hawari',
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
    source: 'builder',
    createdAt: '2026-11-14T08:00:00Z'
  },
  {
    id: 'vault-ref-002',
    title: 'SURAT TUGAS BIMTEK KURIKULUM MERDEKA',
    nomorSurat: '800 / 042 / ST.SMKITA / XI / 2026',
    tanggalSurat: '10 November 2026',
    category: 'Surat Tugas',
    targetPerson: 'Dian Nurhadian, S.Pd. & Tim Guru Kejuruan',
    perihal: 'Penugasan Pelatihan Implementasi Kurikulum Merdeka',
    summary: 'Menugaskan Kepala Sekolah dan guru kejuruan Multimedia dan ATPH mengikuti Bimbingan Teknis Implementasi Kurikulum Merdeka di Balai Penjaminan Mutu Pendidikan (BPMP).',
    isAiReference: true,
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
    source: 'builder',
    createdAt: '2026-11-08T10:00:00Z'
  }
];

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
      zoomLevel: 100, // 100%
      letterVault: JSON.parse(JSON.stringify(DEFAULT_LETTER_VAULT)),
      outgoingLetters: [],
      smartDocument: null,
      activeUploadedDocx: null
    };

    this.db = null;
    this.initIndexedDb();
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
      this.saveToIndexedDb(STORE_PROJECTS, 'active_document', this.state.document);
    }
    if (nextState.smartDocument) {
      this.saveToIndexedDb(STORE_PROJECTS, 'active_smart_document', this.state.smartDocument);
    }
    if (nextState.schoolProfile) {
      this.saveToIndexedDb(STORE_SCHOOL, 'profile', this.state.schoolProfile);
    }
    if (nextState.agendaBook) {
      this.saveToIndexedDb(STORE_AGENDA, 'agenda_list', this.state.agendaBook);
    }
    if (nextState.apiVault) {
      this.saveToIndexedDb(STORE_API_VAULT, 'vault_list', this.state.apiVault);
    }
    if (nextState.kopGallery) {
      this.saveToIndexedDb(STORE_KOP_GALLERY, 'gallery_list', this.state.kopGallery);
    }
  }

  // --- PEMBARUAN PROFIL SEKOLAH ---
  getSchoolProfile() {
    return this.state.schoolProfile || DEFAULT_SCHOOL_PROFILE;
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
    return this.state.letterOptions || {
      paperSize: 'a4',
      orientation: 'portrait',
      margins: { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 },
      fontFamily: 'Times New Roman',
      fontSize: '12',
      lineSpacing: '1.15',
      paragraphSpacing: '10',
      showKop: true
    };
  }

  updateLetterOptions(opts) {
    const updatedOpts = { ...this.getLetterOptions(), ...opts };
    let currentDoc = this.state.document;
    if (currentDoc) {
      currentDoc = {
        ...currentDoc,
        profile: updatedOpts.paperSize === 'f4' ? 'suratDinasF4' : 'suratDinas',
        customMargins: updatedOpts.margins,
        paperSize: updatedOpts.paperSize,
        orientation: updatedOpts.orientation || 'portrait'
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
        localStorage.setItem('letter_options', JSON.stringify(updatedOpts));
      } catch (e) {}
    }
  }

  regenerateActiveDocument(formData, schoolProfile, options) {
    return suratKeteranganTemplate.generateDocumentModel(formData, schoolProfile, options);
  }

  // --- ZOOM LEMBAR KERTAS ---
  setZoomLevel(zoom) {
    const clamped = Math.max(50, Math.min(200, zoom));
    this.setState({ zoomLevel: clamped }, false);
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
    const vault = this.getApiVault().map(item => ({
      ...item,
      isActive: item.id === id
    }));
    this.setState({
      apiVault: vault,
      activeApiId: id
    });
  }

  addApiVaultEntry(entry) {
    const newEntry = {
      id: entry.id || `api-${Date.now()}`,
      label: entry.label || 'Koneksi AI Baru',
      provider: entry.provider || 'gemini',
      model: entry.model || 'gemini-2.0-flash',
      apiKey: entry.apiKey || '',
      baseUrl: entry.baseUrl || '',
      isActive: !!entry.isActive,
      createdAt: Date.now()
    };

    let vault = [...this.getApiVault()];
    if (newEntry.isActive) {
      vault = vault.map(v => ({ ...v, isActive: false }));
    }
    vault.push(newEntry);

    this.setState({
      apiVault: vault,
      activeApiId: newEntry.isActive ? newEntry.id : this.state.activeApiId
    });
  }

  updateApiVaultEntry(id, patch) {
    const vault = this.getApiVault().map(item => {
      if (item.id === id) {
        return { ...item, ...patch };
      }
      if (patch.isActive && item.id !== id) {
        return { ...item, isActive: false };
      }
      return item;
    });

    const activeId = patch.isActive ? id : this.state.activeApiId;
    this.setState({ apiVault: vault, activeApiId: activeId });
  }

  removeApiVaultEntry(id) {
    const vault = this.getApiVault().filter(item => item.id !== id);
    let activeId = this.state.activeApiId;
    if (activeId === id) {
      activeId = vault[0]?.id || '';
      if (vault[0]) vault[0].isActive = true;
    }
    this.setState({ apiVault: vault, activeApiId: activeId });
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
    return newKop;
  }

  removeKopGalleryEntry(id) {
    const gallery = this.getKopGallery().filter(k => k.id !== id);
    this.setState({ kopGallery: gallery });
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

  addLetterVaultEntry(entry) {
    const vault = this.getLetterVault();
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
      source: entry.source || 'builder',
      document: entry.document || null,
      uploadedDocx: entry.uploadedDocx || null,
      createdAt: entry.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    const updated = [newEntry, ...vault];
    this.setState({ letterVault: updated }, false);
    this.saveLetterVault(updated);
    return newEntry;
  }

  updateLetterVaultEntry(id, updates) {
    const vault = this.getLetterVault();
    const updated = vault.map(item => {
      if (item.id === id) {
        return { ...item, ...updates, updatedAt: new Date().toISOString() };
      }
      return item;
    });
    this.setState({ letterVault: updated }, false);
    this.saveLetterVault(updated);
  }

  removeLetterVaultEntry(id) {
    const vault = this.getLetterVault().filter(item => item.id !== id);
    this.setState({ letterVault: vault }, false);
    this.saveLetterVault(vault);
  }

  toggleLetterAiReference(id) {
    const vault = this.getLetterVault();
    let newStatus = false;
    const updated = vault.map(item => {
      if (item.id === id) {
        newStatus = !item.isAiReference;
        return { ...item, isAiReference: newStatus, updatedAt: new Date().toISOString() };
      }
      return item;
    });
    this.setState({ letterVault: updated }, false);
    this.saveLetterVault(updated);
    return newStatus;
  }

  getAiReferences(limit = 4) {
    const vault = this.getLetterVault();
    const activeRefs = vault.filter(v => v.isAiReference);
    return activeRefs.slice(0, limit).map(v => ({
      title: v.title,
      nomorSurat: v.nomorSurat,
      category: v.category,
      targetPerson: v.targetPerson,
      perihal: v.perihal,
      summary: v.summary,
      isAiReference: v.isAiReference
    }));
  }

  loadLetterToCanvas(id) {
    const vault = this.getLetterVault();
    const target = vault.find(v => v.id === id);
    if (!target) return false;

    if (target.uploadedDocx) {
      this.setState({
        activeUploadedDocx: target.uploadedDocx,
        document: target.document || this.state.document
      }, true);
      return true;
    }

    if (target.document) {
      const cloned = JSON.parse(JSON.stringify(target.document));
      delete cloned.verification;
      const titleBlock = (cloned.blocks || []).find(b => b.type === 'letterTitle');
      if (titleBlock) {
        const nextInfo = this.getNextFormattedOutgoingNomor({
          title: titleBlock.title || target.title,
          perihal: target.perihal,
          currentNomor: titleBlock.nomor || target.nomorSurat
        });
        titleBlock.nomor = `Nomor: ${nextInfo.nomorSurat}`;
      }
      this.setState({
        activeUploadedDocx: null,
        document: cloned
      }, true);
      return true;
    }

    // Jika surat tidak memiliki model docx terpisah, adaptasi ke dokumen kanvas saat ini
    const doc = JSON.parse(JSON.stringify(this.state.document || {}));
    if (!doc.metadata) doc.metadata = {};
    doc.metadata.title = target.title || 'SURAT DINAS';
    doc.metadata.nomorSurat = target.nomorSurat || '';
    this.setState({
      activeUploadedDocx: null,
      document: doc
    }, true);
    return true;
  }

  setActiveUploadedDocx(docxData) {
    this.setState({ activeUploadedDocx: docxData }, false);
  }

  clearActiveUploadedDocx() {
    this.setState({ activeUploadedDocx: null }, false);
  }

  saveLetterVault(vaultList) {
    this.saveToIndexedDb(STORE_LETTER_VAULT, 'vault_list', vaultList);
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

  getNextFormattedOutgoingNomor({ title = '', perihal = '', currentNomor = '', forcedTypeCode = '', date = new Date() } = {}) {
    const d = date instanceof Date && !isNaN(date.getTime()) ? date : new Date();
    const year = d.getFullYear().toString();
    const seqNum = this.getNextOutgoingSequence(year);
    const typeCode = (forcedTypeCode || detectLetterTypeCode(`${title} ${perihal} ${currentNomor}`)).toUpperCase();
    return {
      sequenceNumber: seqNum,
      typeCode,
      year,
      nomorSurat: buildOrderedNomorSurat({ typeCode, sequenceNumber: seqNum, date: d })
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
      return;
    }
    const doc = this.state.document;
    if (!doc) return;
    const nextDoc = { ...doc, verification: verificationMeta || null };
    this.setState({ document: nextDoc }, false);
  }

  removeOutgoingLetter(id) {
    const updated = this.getOutgoingLetters().filter(item => item.id !== id);
    this.setState({ outgoingLetters: updated }, false);
    this.saveOutgoingLetters(updated);
    return updated;
  }

  saveOutgoingLetters(list) {
    this.saveToIndexedDb(STORE_OUTGOING_LETTERS, 'outgoing_list', list);
  }

  /**
   * Menyalin isi surat keluar menjadi draf baru di kanvas tanpa pernah mengubah catatan arsip surat keluar aslinya
   */
  cloneOutgoingToNewDraft(id) {
    const target = this.getOutgoingLetters().find(item => item.id === id);
    if (!target || !target.documentSnapshot) return false;
    const clonedDoc = JSON.parse(JSON.stringify(target.documentSnapshot));
    delete clonedDoc.verification;
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

  // --- UNDO / REDO ---
  canUndo() { return this.undoStack.length > 0; }
  canRedo() { return this.redoStack.length > 0; }

  undo() {
    if (!this.canUndo()) return;
    const prevDoc = this.undoStack.pop();
    this.redoStack.push(JSON.parse(JSON.stringify(this.state.document)));
    this.state = { ...this.state, document: prevDoc };
    this.notify();
    this.saveToIndexedDb(STORE_PROJECTS, 'active_document', prevDoc);
  }

  redo() {
    if (!this.canRedo()) return;
    const nextDoc = this.redoStack.pop();
    this.undoStack.push(JSON.parse(JSON.stringify(this.state.document)));
    this.state = { ...this.state, document: nextDoc };
    this.notify();
    this.saveToIndexedDb(STORE_PROJECTS, 'active_document', nextDoc);
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

  restoreFromIndexedDb() {
    if (!this.db) return;

    // 1. Pulihkan Profil Sekolah
    try {
      const txSchool = this.db.transaction([STORE_SCHOOL], 'readonly');
      const storeSchool = txSchool.objectStore(STORE_SCHOOL);
      const reqSchool = storeSchool.get('profile');
      reqSchool.onsuccess = () => {
        if (reqSchool.result) {
          let loadedSchool = reqSchool.result;
          // Migrasi otomatis jika profil sekolah di browser masih sekolah bawaan lama (Surabaya) atau kode surat lama
          if (!loadedSchool.schoolName || loadedSchool.schoolName.includes('SURABAYA') || loadedSchool.principalName?.includes('Eko Wahyudi') || !loadedSchool.numberingPattern?.includes('SMKITA')) {
            loadedSchool = {
              ...JSON.parse(JSON.stringify(DEFAULT_SCHOOL_PROFILE)),
              ...loadedSchool,
              numberingPattern: '421.3/{urut}/SMKITA/{bulanRomawi}/{tahun}'
            };
            this.saveSchoolProfile(loadedSchool);
          }
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
      const storeAgenda = txAgenda.objectStore(STORE_AGENDA);
      const reqAgenda = storeAgenda.get('agenda_list');
      reqAgenda.onsuccess = () => {
        if (Array.isArray(reqAgenda.result)) {
          this.setState({ agendaBook: reqAgenda.result }, false);
        }
      };
    } catch (e) {}

    // 3. Pulihkan Gudang API (API Vault)
    try {
      const txVault = this.db.transaction([STORE_API_VAULT], 'readonly');
      const storeVault = txVault.objectStore(STORE_API_VAULT);
      const reqVault = storeVault.get('vault_list');
      reqVault.onsuccess = () => {
        if (Array.isArray(reqVault.result) && reqVault.result.length > 0) {
          // Migrasi otomatis jika masih ada model Gemini/Groq versi usang yang tersimpan di IndexedDB browser
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
      const storeKop = txKop.objectStore(STORE_KOP_GALLERY);
      const reqKop = storeKop.get('gallery_list');
      reqKop.onsuccess = () => {
        if (Array.isArray(reqKop.result) && reqKop.result.length > 0) {
          let loadedGallery = reqKop.result;
          // Migrasi otomatis ke galeri SMKS IT Al-Hawari jika masih galeri bawaan lama
          if (loadedGallery.some(k => k.name?.includes('SURABAYA') || k.dataUrl?.includes('SURABAYA') || k.description?.includes('pemerintah provinsi dan dinas pendidikan dengan garis ganda'))) {
            loadedGallery = JSON.parse(JSON.stringify(DEFAULT_KOP_GALLERY));
            this.saveKopGallery(loadedGallery);
          }
          this.setState({ kopGallery: loadedGallery }, false);
        }
      };
    } catch (e) {}

    // 5. Pulihkan Gudang Surat (Letter Vault)
    try {
      const txLetter = this.db.transaction([STORE_LETTER_VAULT], 'readonly');
      const storeLetter = txLetter.objectStore(STORE_LETTER_VAULT);
      const reqLetter = storeLetter.get('vault_list');
      reqLetter.onsuccess = () => {
        if (Array.isArray(reqLetter.result) && reqLetter.result.length > 0) {
          this.setState({ letterVault: reqLetter.result }, false);
        }
      };
    } catch (e) {}

    // 6. Pulihkan Buku Catatan Surat Keluar (Outgoing Letters - Terkunci)
    try {
      const txOutgoing = this.db.transaction([STORE_OUTGOING_LETTERS], 'readonly');
      const storeOutgoing = txOutgoing.objectStore(STORE_OUTGOING_LETTERS);
      const reqOutgoing = storeOutgoing.get('outgoing_list');
      reqOutgoing.onsuccess = () => {
        if (Array.isArray(reqOutgoing.result)) {
          this.setState({ outgoingLetters: reqOutgoing.result }, false);
          this.syncActiveDocumentNumberToNextOutgoing(false);
        }
      };
    } catch (e) {}

    try {
      const txProjects = this.db.transaction([STORE_PROJECTS], 'readonly');
      const storeProjects = txProjects.objectStore(STORE_PROJECTS);
      const reqSmart = storeProjects.get('active_smart_document');
      reqSmart.onsuccess = () => {
        if (reqSmart.result && Array.isArray(reqSmart.result.pages)) {
          this.setState({ smartDocument: reqSmart.result }, false);
        }
      };
    } catch (e) {}
  }
}

export const store = new StateStore();
