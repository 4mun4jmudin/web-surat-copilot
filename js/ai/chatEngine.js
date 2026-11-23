/**
 * chatEngine.js - Mesin Percakapan AI Copilot Multi-Turn & Manajemen Riwayat Chat
 * Menyatukan percakapan konsultatif gaya ChatGPT dengan eksekusi pembuat surat
 * dan AI Smart Surgical Editor (Word Pro & Lembar Kerja).
 */

import { formatKnowledgeReferencesSection } from './prompts.js';
import { shouldAllowLayoutRestructure } from './surgicalEdit.js';

export const CHAT_FOLDERS = Object.freeze({
  ALL: 'all',
  PINNED: 'pinned',
  RECENT: 'recent',
  ARCHIVED: 'archived',
  TRASH: 'trash'
});

export const CHAT_QUICK_TEMPLATES = Object.freeze([
  {
    id: 'tpl-skb',
    code: 'KET',
    title: 'Buat Surat Keterangan (SKB)',
    icon: 'doc',
    prompt: 'Buatkan Surat Keterangan Siswa Aktif / Berkelakuan Baik untuk keperluan administrasi sekolah.'
  },
  {
    id: 'tpl-und',
    code: 'UND',
    title: 'Undangan Rapat (UND)',
    icon: 'calendar',
    prompt: 'Buatkan Surat Undangan Rapat Dinas Dewan Guru beserta jadwal acara.'
  },
  {
    id: 'tpl-st',
    code: 'ST',
    title: 'Surat Tugas (ST)',
    icon: 'briefcase',
    prompt: 'Buatkan Surat Tugas Guru beserta lampiran SPPD 8 poin dan halaman Keterangan.'
  },
  {
    id: 'tpl-pgl',
    code: 'PGL',
    title: 'Panggilan (PGL)',
    icon: 'users',
    prompt: 'Buatkan Surat Panggilan Orang Tua/Wali Siswa terkait konsultasi pembinaan siswa.'
  }
]);

export const CHAT_HINT_EXAMPLES = Object.freeze([
  '"Buat surat tugas kegiatan lomba"',
  '"Ubah penerima ke Wali Kelas X"',
  '"Perbaiki PUEBI & atur margin kiri 3cm"'
]);

export function createChatSession({
  id = '',
  title = 'Percakapan Baru',
  targetWorkspace = 'smart-canvas',
  activeModelId = '',
  pinned = false,
  archived = false,
  trashed = false,
  messages = []
} = {}) {
  const now = new Date().toISOString();
  return {
    id: id || `chat-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
    title: String(title || 'Percakapan Baru').trim() || 'Percakapan Baru',
    pinned: Boolean(pinned),
    archived: Boolean(archived),
    trashed: Boolean(trashed),
    isAutoTitle: true,
    targetWorkspace: targetWorkspace === 'canvas' || targetWorkspace === 'none' ? targetWorkspace : 'smart-canvas',
    activeModelId: activeModelId || '',
    messages: Array.isArray(messages) ? messages : [],
    createdAt: now,
    updatedAt: now
  };
}

export function createChatMessage({
  id = '',
  role = 'user',
  text = '',
  replyIntro = '',
  checklist = [],
  followUpQuestion = '',
  action = null,
  attachments = [],
  modelLabel = '',
  timestamp = ''
} = {}) {
  return {
    id: id || `msg-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
    role: role === 'assistant' ? 'assistant' : 'user',
    text: String(text || ''),
    ...(replyIntro ? { replyIntro: String(replyIntro) } : {}),
    checklist: Array.isArray(checklist) ? checklist.filter(Boolean).map(String) : [],
    ...(followUpQuestion ? { followUpQuestion: String(followUpQuestion) } : {}),
    ...(action && typeof action === 'object' ? { action } : {}),
    attachments: Array.isArray(attachments) ? attachments : [],
    ...(modelLabel ? { modelLabel: String(modelLabel) } : {}),
    timestamp: timestamp || new Date().toISOString()
  };
}

export function deriveChatTitleFromMessage(rawText = '') {
  const cleaned = String(rawText || '')
    .replace(/\s+/g, ' ')
    .replace(/^(tolong|mohon|bisa|coba|silakan|hai ai|halo|bro)\s+/i, '')
    .trim();
  if (!cleaned) return 'Percakapan Baru';
  const capitalized = cleaned.charAt(0).toUpperCase() + cleaned.slice(1);
  return capitalized.length > 42 ? `${capitalized.slice(0, 39).trimEnd()}...` : capitalized;
}

export function formatChatTimestamp(isoString, nowMs = Date.now()) {
  if (!isoString) return '';
  const d = new Date(isoString);
  if (Number.isNaN(d.getTime())) return '';
  const now = new Date(nowMs);
  const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const startOfTarget = new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
  const diffDays = Math.round((startOfToday - startOfTarget) / 86400000);

  if (diffDays === 0) {
    return d.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }).replace('.', ':');
  }
  if (diffDays === 1) return 'Kemarin';
  if (diffDays > 1 && diffDays < 7) {
    return d.toLocaleDateString('id-ID', { weekday: 'short', day: 'numeric', month: 'short' });
  }
  return d.toLocaleDateString('id-ID', { day: 'numeric', month: 'short' });
}

export function formatTimeOnly(isoString) {
  if (!isoString) return '';
  const d = new Date(isoString);
  if (Number.isNaN(d.getTime())) return '';
  return d.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }).replace('.', ':');
}

export function computeChatMemoryStats(sessions = []) {
  const list = Array.isArray(sessions) ? sessions : [];
  const active = list.filter(s => !s.trashed);
  const allNonArchived = active.filter(s => !s.archived);
  const pinned = active.filter(s => s.pinned && !s.archived);
  const archived = active.filter(s => s.archived);
  const trashed = list.filter(s => s.trashed);
  const weekAgo = Date.now() - 7 * 86400000;
  const recent = allNonArchived.filter(s => new Date(s.updatedAt || s.createdAt || 0).getTime() >= weekAgo);
  const totalSaved = active.length;
  const quotaPct = Math.max(totalSaved > 0 ? 8 : 0, Math.min(100, Math.round((totalSaved / 40) * 100)));

  return {
    totalSaved,
    allCount: allNonArchived.length,
    pinnedCount: pinned.length,
    recentCount: recent.length,
    archivedCount: archived.length,
    trashCount: trashed.length,
    quotaPct
  };
}

export function filterChatSessions(sessions = [], { folder = CHAT_FOLDERS.ALL, query = '' } = {}) {
  const list = Array.isArray(sessions) ? sessions : [];
  const q = String(query || '').trim().toLowerCase();

  let filtered = list.filter((s) => {
    if (!s) return false;
    if (folder === CHAT_FOLDERS.TRASH) return Boolean(s.trashed);
    if (s.trashed) return false;
    if (folder === CHAT_FOLDERS.ARCHIVED) return Boolean(s.archived);
    if (s.archived) return false;
    if (folder === CHAT_FOLDERS.PINNED) return Boolean(s.pinned);
    if (folder === CHAT_FOLDERS.RECENT) {
      const weekAgo = Date.now() - 14 * 86400000;
      return new Date(s.updatedAt || s.createdAt || 0).getTime() >= weekAgo;
    }
    return true;
  });

  if (q) {
    filtered = filtered.filter((s) => {
      if ((s.title || '').toLowerCase().includes(q)) return true;
      return (s.messages || []).some(m => (m.text || '').toLowerCase().includes(q));
    });
  }

  return filtered.sort((a, b) => {
    if (folder !== CHAT_FOLDERS.TRASH && folder !== CHAT_FOLDERS.ARCHIVED) {
      if (Boolean(a.pinned) !== Boolean(b.pinned)) return a.pinned ? -1 : 1;
    }
    return new Date(b.updatedAt || b.createdAt || 0).getTime() - new Date(a.updatedAt || a.createdAt || 0).getTime();
  });
}

export function buildConversationHistoryForAi(messages = [], limit = 8) {
  const list = Array.isArray(messages) ? messages : [];
  return list
    .filter(m => m && (m.role === 'user' || m.role === 'assistant'))
    .slice(-limit)
    .map((m) => {
      if (m.role === 'user') {
        return { role: 'user', text: String(m.text || '').trim() };
      }
      const parts = [m.text];
      if (Array.isArray(m.checklist) && m.checklist.length > 0) {
        parts.push(`Poin tindakan: ${m.checklist.join('; ')}`);
      }
      if (m.followUpQuestion) {
        parts.push(m.followUpQuestion);
      }
      return { role: 'assistant', text: parts.filter(Boolean).join('\n').trim() };
    })
    .filter(m => m.text);
}

export const CHAT_COPILOT_SYSTEM_PROMPT = `
Kamu adalah "AI Asisten Persuratan Sekolah" yang profesional, ramah, cerdas, dan terstruktur (seperti ChatGPT/Claude khusus administrasi surat dinas 2 unit sekolah Yayasan Al-Hawari: SMK IT Al-Hawari & SMP IT Al-Hawari - Cibiuk).
Kembalikan HANYA objek JSON valid (tanpa bungkus markdown \`\`\`json).

DATA RESMI 2 UNIT & 2 PEJABAT KEPALA SEKOLAH:
1. Unit SMK IT AL-HAWARI (kode unit "smk", kode nomor "SMKITA"):
   - Kepala Sekolah: "Dian Nurhamdani, S.Pd."
   - Format nomor contoh: "050/UND.SMKITA/XI/2026" atau "421.3/085/KET.SMKITA/XI/2026"
2. Unit SMP IT AL-HAWARI - CIBIUK (kode unit "smp", kode nomor "SMPITA"):
   - Kepala Sekolah: "Yangyang Budiman, S.Pd."
   - Format nomor contoh: "050/UND.SMPITA/XI/2026" atau "421.2/085/KET.SMPITA/XI/2026"
Jika pengguna meminta surat untuk SMP (atau menyebut SMP / SMPITA / Yangyang Budiman / kelas VII-IX), WAJIB pakai profil, kode SMPITA, dan Kepala Sekolah "Yangyang Budiman, S.Pd.". Jika untuk SMK (atau menyebut SMK / SMKITA / Dian Nurhamdani / PKL / jurusan), WAJIB pakai profil, kode SMKITA, dan Kepala Sekolah "Dian Nurhamdani, S.Pd.". Ikuti KONTEKS SEKOLAH (UNIT TARGET) yang dikirim pada pesan.

TUGAS UTAMA: Pahami maksud pengguna dari riwayat percakapan & pesan terbaru, lalu pilih SATU dari 3 nilai "intent":
1. "chat" -> Jika pengguna bertanya, berkonsultasi, berdiskusi, meminta saran redaksi/kode nomor surat, menanyakan isi dokumen atau lampiran, atau belum menyuruh membuat/mengubah surat.
2. "create_letter" -> Jika pengguna meminta membuatkan/menulis draf surat baru (termasuk membalas lampiran gambar/PDF menjadi surat dinas/surat tugas/undangan/keterangan baru).
3. "edit_letter" -> Jika pengguna meminta memperbaiki ejaan PUEBI, mengoreksi kalimat, mengganti data (nama/tanggal/nomor/tujuan/baris tabel SPPD), atau mengatur format Word (kertas A4/F4/Letter/Legal, potret/lanskap, margin cm, font, ukuran huruf, spasi, indentasi, perataan, gaya numbering 1/a/A/i/I/bullet, nomor halaman) pada DOKUMEN TERBUKA saat ini.

SKEMA JSON WAJIB:
{
  "intent": "chat" | "create_letter" | "edit_letter",
  "reply": "Penjelasan ramah, jelas, & profesional dalam Bahasa Indonesia PUEBI (2-4 kalimat).",
  "checklist": [
    "Poin ringkas tindakan / temuan / saran (opsional 0-5 poin singkat)"
  ],
  "followUpQuestion": "Pertanyaan tindak lanjut yang membantu pengguna (mis. 'Apakah Anda ingin saya langsung menerapkan perbaikannya pada editor?' atau saran lanjutan bila intent = chat).",
    "letterDraft": {
    "// Hanya diisi jika intent === 'create_letter'": true,
    "targetUnit": "smk | smp",
    "title": "Judul Singkat",
    "letterTitle": "JUDUL SURAT HURUF KAPITAL",
    "kodeJenisSurat": "UND | TU | KEP | KET | ST | PGL | PBR | PRM | DISP | BA | REK | PKL | KUR | KES",
    "nomorSurat": "050/UND.SMKITA/XI/2026",
    "tanggalSurat": "14 November 2026",
    "hideMetaRecipient": false,
    "sifat": "Penting",
    "hal": "Perihal Surat",
    "lampiran": "-",
    "tujuan": "Yth. ...",
    "signer": {
      "city": "Garut",
      "date": "14 November 2026",
      "title": "Kepala Sekolah",
      "name": "Dian Nurhamdani, S.Pd. atau Yangyang Budiman, S.Pd. sesuai unit",
      "nip": "-"
    },
    "pageSetup": {
      "paperSize": "a4",
      "orientation": "portrait",
      "margins": { "top": 2.0, "bottom": 2.0, "left": 2.5, "right": 2.0 },
      "fontFamily": "Times New Roman",
      "fontSize": "12",
      "lineSpacing": "1.15",
      "paragraphSpacing": "6",
      "pageNumbering": { "enabled": false, "position": "bottom-center", "format": "decimal", "startAt": 1 }
    },
    "blocks": [
      { "type": "paragraph", "text": "...", "alignment": "both" }
    ],
    "tembusan": []
  },
  "editPlan": {
    "// Hanya diisi jika intent === 'edit_letter'": true,
    "editScope": "surgical",
    "summaryOfChanges": "Ringkasan singkat perubahan",
    "patches": [
      { "op": "setField", "target": "title" | "nomor", "text": "..." },
      { "op": "setMetaRecipient", "blockId": "p0_meta0", "sifat": "...", "hal": "...", "tujuanLines": ["Kepada", "Yth. ...", "Di Tempat"] },
      { "op": "setParagraph", "blockId": "p0_b0", "text": "..." },
      { "op": "setParagraphFormat", "blockId": "p0_b0" | "*", "alignment": "left"|"center"|"right"|"both", "firstLineIndentCm": 1.25, "leftIndentCm": 0, "spacingBeforePt": 0, "spacingAfterPt": 6, "lineSpacing": "1.5", "fontSizePt": 12, "fontFamily": "Arial" },
      { "op": "setPageFormat", "paperSize": "a4"|"f4"|"letter"|"legal", "orientation": "portrait"|"landscape", "margins": { "top": 2.0, "bottom": 2.0, "left": 2.5, "right": 2.0 }, "fontFamily": "Times New Roman", "fontSize": "12", "lineSpacing": "1.15", "paragraphSpacing": "6", "showKop": true, "pageNumbering": { "enabled": true, "position": "bottom-center"|"bottom-right"|"top-right", "format": "decimal"|"lowerRoman"|"upperRoman", "startAt": 1 } },
      { "op": "setIdentityRow", "blockId": "p0_id0", "row": 0, "label": "...", "value": ": ..." },
      { "op": "setTableCell", "blockId": "p0_t0", "row": 1, "col": 1, "text": "..." },
      { "op": "addTableRow", "blockId": "p0_t0", "cells": ["..."] },
      { "op": "setTableFormat", "blockId": "p0_t0", "borderStyle": "grid"|"sppd"|"none", "headerShading": true|false },
      { "op": "setList", "blockId": "p0_list0", "listStyle": "decimal"|"lower-alpha"|"upper-alpha"|"lower-roman"|"upper-roman"|"bullet", "startAt": 1 },
      { "op": "convertParagraphsToList", "blockIds": ["p0_b0", "p0_b1"], "listStyle": "lower-alpha" },
      { "op": "setSignature", "blockId": "p0_sig0", "city": "Garut", "date": "...", "title": "...", "name": "...", "nip": "..." }
    ]
  }
}

ATURAN KETAT:
- JIKA intent === "edit_letter" & mode kunci aktif: WAJIB pakai "editScope": "surgical" dengan "patches" tepat di "blockId" / "(row, col)" dokumen terbuka. DILARANG membuat ulang tabel SPPD atau menghapus "[QR_SLOT_TERKUNCI]".
- Selalu gunakan Nama Sekolah, Nama Kepala Sekolah, dan Kode Unit pada KONTEKS SEKOLAH yang diberikan (mis. SMK IT Al-Hawari / SMKITA atau SMP IT Al-Hawari - Cibiuk / SMPITA). Jangan mengarang data sensitif jika belum disebut pengguna (gunakan placeholder [Nama Siswa], dll.).
- Jika dokumen terbuka masih kosong tetapi pengguna menyuruh membuat surat, gunakan intent "create_letter" dan isi "letterDraft".
- Untuk "Surat Tugas + SPPD", boleh mengisi "letterDraft" dengan { "templatePreset": "surat_tugas_sppd", "letterTitle": "SURAT TUGAS & SPPD", "kodeJenisSurat": "ST", "formData": { "assignees": [...], "maksudTugas": "...", "tujuanLokasi": "...", "maksudSppd": "...", "kendaraan": "Mobil", "lamaHari": "1", "tglBerangkat": "...", "tglKembali": "..." }, "blocks": [{ "type": "paragraph", "text": "Template Surat Tugas & SPPD" }] }.
`.trim();

export const DIRECT_CANVAS_SYSTEM_SUFFIX = `
ATURAN KHUSUS MODE KANVAS LANGSUNG (WAJIB):
- Saat pengguna memberi instruksi di dalam editor kanvas untuk membuat surat baru, menambah jadwal/isi, memperbaiki ejaan/kalimat, mengganti data, atau mengubah format Word, DILARANG membalas dengan intent "chat" hanya untuk meminta konfirmasi atau menanyakan izin penerapan!
- LANGSUNG eksekusi: kembalikan "intent": "create_letter" (dengan "letterDraft" lengkap) atau "intent": "edit_letter" (dengan "editPlan" lengkap).
- Jika ada data detail yang belum disebut (misal nomor, jam, nama program), tetap buat/edit suratnya sekarang juga memakai placeholder rapi seperti [Nama Program] atau isian standar sekolah yang wajar agar dokumen di kanvas langsung terbarui seketika.
- Kosongkan "followUpQuestion" (atau isi saran opsional sangat singkat); JANGAN pernah bertanya "Apakah Anda ingin saya terapkan?" karena perubahan langsung diterapkan otomatis ke kanvas.
`.trim();

/**
 * Mendeteksi apakah pesan pengguna merupakan perintah aksi buat/ubah dokumen (bukan sekadar tanya jawab)
 */
export function isDirectExecutionCommand(message = '') {
  const txt = String(message || '').trim().toLowerCase();
  if (!txt) return false;
  const actionRe = /\b(buat(?:kan)?|bikin(?:kan)?|susun(?:kan)?|tulis(?:kan)?|cetak|rancang|ubah|ganti|revisi|perbaiki|rapikan|rapihkan|koreksi|sunting|edit|tambah(?:kan)?|sisipkan|hapus|hilangkan|atur|sesuaikan|jadikan|terapkan|perbesar|perkecil|spasi|margin|numbering|penomoran|kertas|f4|a4|lanskap|landscape|potret|portrait|kop)\b/i;
  if (!actionRe.test(txt)) return false;
  const pureQuestionRe = /^(apa|apakah|bagaimana|gimana|kenapa|mengapa|kapan|dimana|siapa|bolehkah|bisakah|tolong jelaskan|jelaskan)\b/i;
  if (pureQuestionRe.test(txt) && !/\b(buat(?:kan)?|bikin|tulis|ubah|ganti|perbaiki|rapikan|rapihkan|atur|tambah(?:kan)?|hapus|revisi)\b/i.test(txt)) {
    return false;
  }
  return true;
}

export function buildChatCopilotUserPrompt({
  userMessage = '',
  activeDocumentSnapshot = null,
  targetWorkspace = 'smart-canvas',
  schoolName = 'SMK IT Al-Hawari',
  principalName = 'Dian Nurhamdani, S.Pd.',
  schoolCode = '',
  unitId = '',
  npsn = '',
  address = '',
  tuName = '',
  dateStr = '',
  references = [],
  matchMode = 'fresh',
  totalKnowledgeCount = 0,
  allowLayoutRestructure = false,
  attachmentsMeta = [],
  executionMode = 'confirm'
} = {}) {
  const isSmp = unitId === 'smp' || /\bsmp\b/i.test(String(schoolName || schoolCode || ''));
  const resolvedSchoolCode = schoolCode || (isSmp ? 'SMPITA' : 'SMKITA');
  const resolvedPrincipal = principalName || (isSmp ? 'Yangyang Budiman, S.Pd.' : 'Dian Nurhamdani, S.Pd.');
  const refSection = formatKnowledgeReferencesSection({
    references,
    matchMode,
    totalKnowledgeCount,
    modePurpose: 'percakapan & penyusunan/revisi surat'
  });

  const hasOpenDoc = Boolean(activeDocumentSnapshot && targetWorkspace !== 'none');
  const openDocSection = hasOpenDoc
    ? `DOKUMEN TERBUKA SAAT INI (Editor: ${targetWorkspace === 'canvas' ? 'Lembar Kerja Standar' : 'Word Pro Multi-Halaman'}):\n${JSON.stringify(activeDocumentSnapshot, null, 2)}`
    : 'DOKUMEN TERBUKA SAAT INI: Kosong / Belum ada draf aktif.';

  const modeDirective = allowLayoutRestructure
    ? 'Status Kunci Tabel: RESTRUKTURISASI DIIZINKAN bila memang diminta eksplisit.'
    : 'Status Kunci Tabel: MODE BEDAH (SURGICAL) AKTIF — jaga struktur tabel & slot QR tetap utuh jika melakukan edit_letter.';

  const directApplyDirective = executionMode === 'direct_apply'
    ? '\nMODE EKSEKUSI LANGSUNG KE KANVAS AKTIF: Pengguna sedang bekerja langsung di dalam editor. Apabila pesan pengguna menginstruksikan membuat surat baru, menambah isi/jadwal/bagian, mengubah teks/identitas/data, atau mengatur format (meski detailnya singkat — lengkapi bagian standar secara wajar atau gunakan placeholder [..] yang rapi), WAJIB langsung kembalikan intent "create_letter" atau "edit_letter" beserta payload lengkap agar langsung diperbarui di kanvas tanpa bertanya konfirmasi penerapan! Gunakan intent "chat" HANYA jika pengguna murni bertanya/berkonsultasi tanpa meminta pembuatan atau perubahan dokumen. Untuk followUpQuestion pada create_letter/edit_letter, kosongkan saja atau beri saran penyempurnaan lanjutan yang singkat.'
    : '';

  const attSection = Array.isArray(attachmentsMeta) && attachmentsMeta.length > 0
    ? `Lampiran Pengguna Saat Ini: ${attachmentsMeta.map(a => `${a.name || 'berkas'} (${a.mimeType || 'file'})`).join(', ')}. Baca isi lampiran ini dengan teliti bila relevan dengan pertanyaan/perintah.`
    : '';

  return `
PESAN PENGGUNA SAAT INI:
"${userMessage}"
${attSection ? `\n${attSection}` : ''}

KONTEKS SEKOLAH (UNIT TARGET TERDETEKSI: ${isSmp ? 'SMP' : 'SMK'}):
- targetUnit: "${isSmp ? 'smp' : 'smk'}"
- Nama Sekolah Resmi: ${schoolName || (isSmp ? 'SMP IT AL-HAWARI - CIBIUK' : 'SMK IT AL-HAWARI')}
- Kepala Sekolah (Penanda Tangan Wajib): ${resolvedPrincipal}
- Kode Penomoran Unit: ${resolvedSchoolCode} (gunakan format {prefix}/{urut}/{KODE_JENIS}.${resolvedSchoolCode}/{bulanRomawi}/{tahun})${npsn ? `\n- NPSN: ${npsn}` : ''}${address ? `\n- Alamat: ${address}` : ''}${tuName ? `\n- Kepala Tata Usaha (TU): ${tuName}` : ''}
- Tanggal Hari Ini: ${dateStr || new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' })}
- ${modeDirective}${directApplyDirective}

${openDocSection}

${refSection}

Kembalikan JSON terstruktur sesuai skema CHAT_COPILOT_SYSTEM_PROMPT.
`.trim();
}

export function normalizeAiChatResponse(parsedJson, {
  userMessage = '',
  hasActiveDoc = false,
  allowLayoutRestructure = false,
  executionMode = 'confirm',
  targetUnitId = '',
  targetSchoolProfile = null
} = {}) {
  const raw = (parsedJson && typeof parsedJson === 'object') ? parsedJson : {};
  let intent = String(raw.intent || '').trim().toLowerCase();

  let letterDraft = raw.letterDraft && typeof raw.letterDraft === 'object' ? raw.letterDraft : null;
  let editPlan = raw.editPlan && typeof raw.editPlan === 'object' ? raw.editPlan : null;

  // Backward-compatible fallback jika LLM mengembalikan payload flat (tanpa wrapper intent) atau salah melabeli 'chat' padahal menyertakan draf/patch
  if (!['chat', 'create_letter', 'edit_letter'].includes(intent) || (executionMode === 'direct_apply' && intent === 'chat' && (letterDraft || editPlan || Array.isArray(raw.patches) || Array.isArray(raw.blocks)))) {
    if (editPlan && (Array.isArray(editPlan.patches) || Array.isArray(editPlan.blocks))) {
      intent = 'edit_letter';
    } else if (letterDraft && (Array.isArray(letterDraft.blocks) || letterDraft.templatePreset || letterDraft.letterTitle)) {
      intent = 'create_letter';
    } else if (Array.isArray(raw.patches) && raw.patches.length > 0) {
      intent = 'edit_letter';
      editPlan = {
        editScope: raw.editScope || 'surgical',
        summaryOfChanges: raw.summaryOfChanges || raw.reply || 'Perbaikan surat siap diterapkan.',
        patches: raw.patches
      };
    } else if (Array.isArray(raw.blocks) && raw.blocks.length > 0) {
      if (hasActiveDoc && raw.editScope) {
        intent = 'edit_letter';
        editPlan = { ...raw };
      } else {
        intent = 'create_letter';
        letterDraft = { ...raw };
      }
    } else if (raw.templatePreset) {
      intent = 'create_letter';
      letterDraft = { ...raw };
    } else {
      intent = 'chat';
    }
  }

  if (intent === 'create_letter' && !letterDraft && (Array.isArray(raw.blocks) || raw.templatePreset)) {
    letterDraft = { ...raw };
  }
  if (intent === 'edit_letter' && !editPlan && (Array.isArray(raw.patches) || Array.isArray(raw.blocks))) {
    editPlan = { ...raw };
  }

  // Pada kanvas kosong di mode eksekusi langsung, editPlan penuh otomatis dipromosikan menjadi create_letter
  if (executionMode === 'direct_apply' && !hasActiveDoc && intent === 'edit_letter' && editPlan && !letterDraft && Array.isArray(editPlan.blocks) && editPlan.blocks.length > 0) {
    intent = 'create_letter';
    letterDraft = { ...editPlan };
    editPlan = null;
  }

  // Paksa mode surgical pada editPlan bila user tidak meminta rombak tata letak
  const allowRestructure = allowLayoutRestructure || shouldAllowLayoutRestructure(userMessage);
  if (intent === 'edit_letter' && editPlan && Array.isArray(editPlan.patches) && editPlan.patches.length > 0 && !allowRestructure) {
    editPlan = { ...editPlan, editScope: 'surgical' };
  }

  if (intent === 'create_letter' && !letterDraft) {
    intent = 'chat';
  }
  if (intent === 'edit_letter' && !editPlan) {
    intent = 'chat';
  }

  // Selaraskan targetUnit & penandatangan Kepala Sekolah otomatis bila ada konteks target unit SMP/SMK
  const resolvedUnit = targetUnitId || targetSchoolProfile?.unitId
    || (letterDraft?.targetUnit && /^(smp|smk)$/i.test(String(letterDraft.targetUnit)) ? String(letterDraft.targetUnit).toLowerCase() : '');
  if (intent === 'create_letter' && letterDraft && resolvedUnit) {
    const defaultPrincipal = resolvedUnit === 'smp' ? 'Yangyang Budiman, S.Pd.' : 'Dian Nurhamdani, S.Pd.';
    const expectedPrincipal = targetSchoolProfile?.principalName || defaultPrincipal;
    const expectedCode = targetSchoolProfile?.schoolCode || (resolvedUnit === 'smp' ? 'SMPITA' : 'SMKITA');
    const prevSig = (letterDraft.signer && typeof letterDraft.signer === 'object') ? letterDraft.signer : {};
    const sigTitle = String(prevSig.title || '').trim();
    const isExplicitTu = /\btata\s*usaha\b|\bktu\b|\btu\b/i.test(sigTitle) && !/kepala\s*sekolah/i.test(sigTitle);
    const sigNameStr = String(prevSig.name || '').trim();
    const mentionsOtherKepsek = resolvedUnit === 'smp'
      ? /dian\s+nurh/i.test(sigNameStr)
      : /yangyang\s+budiman/i.test(sigNameStr);
    const fixedNomor = letterDraft.nomorSurat
      ? String(letterDraft.nomorSurat).replace(/\b(SMKITA|SMPITA)\b/gi, expectedCode)
      : letterDraft.nomorSurat;

    letterDraft = {
      ...letterDraft,
      targetUnit: resolvedUnit,
      ...(fixedNomor ? { nomorSurat: fixedNomor } : {}),
      signer: {
        ...prevSig,
        city: prevSig.city || targetSchoolProfile?.city || 'Garut',
        date: prevSig.date || letterDraft.tanggalSurat || '',
        title: isExplicitTu ? (targetSchoolProfile?.tuTitle || sigTitle) : (prevSig.title || targetSchoolProfile?.principalTitle || 'Kepala Sekolah'),
        name: isExplicitTu
          ? (prevSig.name || targetSchoolProfile?.tuName || expectedPrincipal)
          : ((!sigNameStr || mentionsOtherKepsek || /kepala\s*sekolah/i.test(sigTitle || 'Kepala Sekolah')) ? expectedPrincipal : sigNameStr),
        nip: prevSig.nip || targetSchoolProfile?.principalNip || '-'
      }
    };
  }

  const checklist = Array.isArray(raw.checklist)
    ? raw.checklist.map(x => String(x || '').trim()).filter(Boolean).slice(0, 6)
    : [];

  let reply = String(raw.reply || raw.message || raw.answer || raw.summaryOfChanges || '').trim();
  if (!reply) {
    if (intent === 'create_letter') {
      reply = `Baik, saya telah menyusun draf "${letterDraft.letterTitle || letterDraft.title || 'Surat Resmi Sekolah'}" sesuai kebutuhan Anda.`;
    } else if (intent === 'edit_letter') {
      reply = editPlan.summaryOfChanges || 'Baik, saya sudah menyiapkan rancangan perbaikan dokumen dengan tetap menjaga struktur dan tabel yang rapi.';
    } else {
      reply = 'Baik, ada hal lain pada surat ini yang ingin Anda diskusikan atau sesuaikan?';
    }
  }

  let followUpQuestion = String(raw.followUpQuestion || '').trim();
  if (executionMode === 'direct_apply' && (intent === 'create_letter' || intent === 'edit_letter')) {
    if (/apakah (anda|bapak\/ibu) ingin.*(menerapkan|terapkan|perbaikannya pada editor|ke editor)/i.test(followUpQuestion)) {
      followUpQuestion = '';
    }
  } else if (!followUpQuestion) {
    if (intent === 'create_letter') {
      followUpQuestion = 'Apakah Anda ingin langsung menerapkan draf surat ini ke editor sekarang?';
    } else if (intent === 'edit_letter') {
      followUpQuestion = 'Apakah Anda ingin saya langsung menerapkan perbaikannya pada editor?';
    }
  }

  return {
    intent,
    reply,
    checklist,
    followUpQuestion,
    letterDraft,
    editPlan
  };
}
