/**
 * state.js - Reactive State Store & IndexedDB Persistence
 * Mendukung penyimpanan Profil Sekolah, Form Surat Aktif, Gudang API Multi-Provider,
 * Gudang Galeri Kop Surat, Skala Zoom Kanvas, dan Riwayat Undo/Redo.
 */

import { DEFAULT_SCHOOL_PROFILE } from './letters/schoolProfile.js';
import { suratKeteranganTemplate } from './letters/templates/suratKeterangan.js';
import { DEFAULT_KOP_GALLERY, DEFAULT_KOP_DATA_URL } from './letters/defaultKop.js';

const DB_NAME = 'ai_school_doc_builder_db';
const DB_VERSION = 4;
const STORE_PROJECTS = 'projects';
const STORE_SCHOOL = 'school_profile';
const STORE_AGENDA = 'agenda_book';
const STORE_API_VAULT = 'api_vault';
const STORE_KOP_GALLERY = 'kop_gallery';
const STORE_LETTER_VAULT = 'letter_vault';

// Nilai awal form surat keterangan (SMKS IT Al-Hawari)
const INITIAL_FORM_DATA = {
  nomorSurat: '421.3 / 085 / KET.SMKITA / XI / 2026',
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

// Preset Bawaan Gudang API (Dapat Ditambah/Diedit User)
export const DEFAULT_API_VAULT = [
  {
    id: 'gemini-auto',
    label: 'Google Gemini 3.5 Flash-Lite (Instan & 15 RPM)',
    provider: 'gemini',
    model: 'gemini-3.5-flash-lite',
    apiKey: '',
    baseUrl: 'https://generativelanguage.googleapis.com/v1beta',
    isActive: true,
    isBuiltin: true
  },
  {
    id: 'groq-fast',
    label: 'Groq - GPT OSS 120B (Ultra Cepat ~0.2s)',
    provider: 'groq',
    model: 'openai/gpt-oss-120b',
    apiKey: '',
    baseUrl: 'https://api.groq.com/openai/v1',
    isActive: false,
    isBuiltin: true
  },
  {
    id: 'deepseek-v3',
    label: 'DeepSeek Chat (V3 / R1 Reasoner)',
    provider: 'deepseek',
    model: 'deepseek-chat',
    apiKey: '',
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

    const initialDoc = suratKeteranganTemplate.generateDocumentModel(
      initialForm,
      initialSchool,
      initialOptions
    );

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
    const newDoc = this.regenerateActiveDocument(this.state.currentLetterForm, updated, this.state.letterOptions);

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
      this.setState({
        activeUploadedDocx: null,
        document: JSON.parse(JSON.stringify(target.document))
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

          const activeEntry = migratedVault.find(v => v.isActive) || migratedVault[0];
          this.setState({
            apiVault: migratedVault,
            activeApiId: activeEntry?.id || this.state.activeApiId
          }, false);
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
  }
}

export const store = new StateStore();
