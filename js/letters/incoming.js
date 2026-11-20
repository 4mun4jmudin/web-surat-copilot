/**
 * js/letters/incoming.js - Helper Penomoran Kode, Pencarian, Kompresi Gambar/Foto,
 * dan Spesifikasi Bridge Google Drive untuk Surat Masuk.
 */

const ROMAN_MONTHS = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'];

export const MAX_INCOMING_RAW_BYTES = 3_800_000; // ~3.8 MB biner (~5 MB JSON base64 agar aman < 6 MB limit Netlify Functions)

export const INCOMING_CATEGORIES = [
  'Undangan',
  'Surat Edaran',
  'Surat Dinas / Pemberitahuan',
  'Permohonan / Pengajuan',
  'Surat Tugas / Rekomendasi',
  'Laporan / Keuangan',
  'Lainnya'
];

export function formatIncomingCode({ sequenceNumber = 1, date = new Date(), prefix = 'SM' } = {}) {
  const d = date instanceof Date && !Number.isNaN(date.getTime()) ? date : new Date();
  const seq = String(Math.max(1, Number(sequenceNumber) || 1)).padStart(3, '0');
  const roman = ROMAN_MONTHS[d.getMonth()] || 'I';
  const year = d.getFullYear();
  return `${prefix}-${seq}/${roman}/${year}`;
}

export function getNextIncomingSequence(list = [], year = new Date().getFullYear()) {
  const targetYear = String(year);
  const maxSeq = (Array.isArray(list) ? list : [])
    .filter(item => String(item?.year || '') === targetYear || !item?.year)
    .reduce((max, item) => Math.max(max, Number(item?.sequenceNumber) || 0), 0);
  return maxSeq + 1;
}

export function detectIncomingFileKind(fileName = '', mimeType = '', sourceHint = '') {
  const mime = String(mimeType || '').toLowerCase();
  const name = String(fileName || '').toLowerCase();
  if (mime.includes('pdf') || name.endsWith('.pdf')) return 'pdf';
  if (sourceHint === 'camera_photo') return 'photo';
  if (mime.startsWith('image/') || /\.(jpg|jpeg|png|webp|gif|bmp|heic)$/i.test(name)) return 'image';
  return 'image';
}

export function formatFileSizeLabel(bytes = 0) {
  const n = Number(bytes) || 0;
  if (n <= 0) return '-';
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(2)} MB`;
}

export function deriveDefaultTitleFromFilename(fileName = '') {
  const base = String(fileName || '')
    .replace(/\.[^/.]+$/, '')
    .replace(/[_\-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return base || '';
}

function damerauLevenshteinLocal(a, b) {
  const m = a.length;
  const n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + cost);
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) {
        dp[i][j] = Math.min(dp[i][j], dp[i - 2][j - 2] + cost);
      }
    }
  }
  return dp[m][n];
}

export function scoreIncomingMatch(item = {}, query = '') {
  const q = String(query || '').trim().toLowerCase();
  if (!q) return 1;

  const kode = String(item.kodeSurat || '').toLowerCase();
  const nama = String(item.namaSurat || '').toLowerCase();
  const nomorAsli = String(item.nomorSuratAsli || '').toLowerCase();
  const pengirim = String(item.pengirim || '').toLowerCase();
  const kategori = String(item.kategori || '').toLowerCase();
  const catatan = String(item.catatan || '').toLowerCase();
  const fileName = String(item.fileName || '').toLowerCase();
  const combined = `${kode} ${nama} ${nomorAsli} ${pengirim} ${kategori} ${catatan} ${fileName}`;

  if (combined.includes(q)) {
    if (kode.includes(q) || nama.includes(q)) return 100;
    return 82;
  }

  const tokens = q.split(/\s+/).filter(Boolean);
  const words = combined.split(/[\s\/.,\-_()]+/).filter(Boolean);
  let matched = 0;
  let scoreSum = 0;

  for (const tok of tokens) {
    let best = 0;
    for (const w of words) {
      if (w.includes(tok)) {
        best = 1;
        break;
      }
      const maxLen = Math.max(tok.length, w.length);
      if (maxLen >= 3) {
        const dist = damerauLevenshteinLocal(tok, w);
        const tol = tok.length <= 4 ? 1 : tok.length <= 8 ? 2 : 3;
        if (dist <= tol) {
          best = Math.max(best, 1 - dist / maxLen);
        }
      }
    }
    if (best >= 0.45) {
      matched += 1;
      scoreSum += best;
    }
  }

  if (matched === 0) return 0;
  return (matched / tokens.length) * 55 + scoreSum * 10;
}

export function sortIncomingLetters(list = [], mode = 'newest') {
  return [...(Array.isArray(list) ? list : [])].sort((a, b) => {
    if (mode === 'oldest') {
      return new Date(a?.receivedAt || a?.createdAt || 0).getTime() - new Date(b?.receivedAt || b?.createdAt || 0).getTime();
    }
    if (mode === 'code_asc') {
      return String(a?.kodeSurat || '').localeCompare(String(b?.kodeSurat || ''), 'id');
    }
    if (mode === 'name_asc') {
      return String(a?.namaSurat || '').localeCompare(String(b?.namaSurat || ''), 'id');
    }
    // newest
    const tDiff = new Date(b?.receivedAt || b?.createdAt || 0).getTime() - new Date(a?.receivedAt || a?.createdAt || 0).getTime();
    if (tDiff !== 0) return tDiff;
    return (Number(b?.sequenceNumber) || 0) - (Number(a?.sequenceNumber) || 0);
  });
}

/**
 * Membaca File/Blob menjadi base64 murni + membuat thumbnail kecil dan mengompres gambar bila berukuran besar.
 */
export async function prepareIncomingFilePayload(fileOrBlob, {
  sourceHint = '',
  customFileName = ''
} = {}) {
  if (!fileOrBlob) throw new Error('Berkas belum dipilih.');
  const fileName = customFileName || fileOrBlob.name || `surat_masuk_${Date.now()}.jpg`;
  const origMime = fileOrBlob.type || (fileName.toLowerCase().endsWith('.pdf') ? 'application/pdf' : 'image/jpeg');
  const fileKind = detectIncomingFileKind(fileName, origMime, sourceHint);

  if (fileKind === 'pdf') {
    if (fileOrBlob.size > MAX_INCOMING_RAW_BYTES) {
      throw new Error(`Ukuran PDF terlalu besar (${formatFileSizeLabel(fileOrBlob.size)}). Batas maksimal upload PDF adalah 3,8 MB.`);
    }
    const dataUrl = await readBlobAsDataUrl(fileOrBlob);
    const fileBase64 = dataUrl.includes(',') ? dataUrl.split(',')[1] : '';
    return {
      fileName: fileName.toLowerCase().endsWith('.pdf') ? fileName : `${fileName}.pdf`,
      mimeType: 'application/pdf',
      fileKind: 'pdf',
      fileSizeBytes: fileOrBlob.size || Math.round((fileBase64.length * 3) / 4),
      fileBase64,
      previewThumbDataUrl: '',
      localDataUrl: dataUrl
    };
  }

  // Untuk Gambar / Foto Kamera: kompres otomatis agar tajam dibaca tapi ringan disimpan & cepat masuk Drive
  const compressed = await compressImageInBrowser(fileOrBlob, fileName, origMime);
  return {
    fileName: compressed.fileName,
    mimeType: compressed.mimeType,
    fileKind,
    fileSizeBytes: compressed.fileSizeBytes,
    fileBase64: compressed.fileBase64,
    previewThumbDataUrl: compressed.previewThumbDataUrl,
    localDataUrl: compressed.localDataUrl
  };
}

function readBlobAsDataUrl(blob) {
  return new Promise((resolve, reject) => {
    if (typeof FileReader === 'undefined') {
      reject(new Error('FileReader tidak tersedia di lingkungan ini.'));
      return;
    }
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ''));
    reader.onerror = () => reject(new Error('Gagal membaca berkas.'));
    reader.readAsDataURL(blob);
  });
}

async function compressImageInBrowser(blob, origName, origMime) {
  const rawDataUrl = await readBlobAsDataUrl(blob);
  if (typeof Image === 'undefined' || typeof document === 'undefined') {
    const rawB64 = rawDataUrl.includes(',') ? rawDataUrl.split(',')[1] : '';
    return {
      fileName: origName,
      mimeType: origMime || 'image/jpeg',
      fileSizeBytes: blob.size || Math.round((rawB64.length * 3) / 4),
      fileBase64: rawB64,
      previewThumbDataUrl: rawB64.length < 40000 ? rawDataUrl : '',
      localDataUrl: rawDataUrl
    };
  }

  const img = await new Promise((resolve, reject) => {
    const im = new Image();
    im.onload = () => resolve(im);
    im.onerror = () => reject(new Error('Berkas gambar tidak dapat dimuat.'));
    im.src = rawDataUrl;
  });

  const maxSide = 1600;
  let targetW = img.naturalWidth || img.width || 1200;
  let targetH = img.naturalHeight || img.height || 1600;
  if (targetW > maxSide || targetH > maxSide) {
    const ratio = Math.min(maxSide / targetW, maxSide / targetH);
    targetW = Math.max(1, Math.round(targetW * ratio));
    targetH = Math.max(1, Math.round(targetH * ratio));
  }

  const canvas = document.createElement('canvas');
  canvas.width = targetW;
  canvas.height = targetH;
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, targetW, targetH);
  ctx.drawImage(img, 0, 0, targetW, targetH);

  // Gunakan PNG hanya bila gambar kecil & bukan foto, selainnya JPEG kualitas 0.84 agar hemat & tajam
  const keepPng = origMime === 'image/png' && blob.size <= 320 * 1024 && targetW <= 1200 && targetH <= 1200;
  const outMime = keepPng ? 'image/png' : 'image/jpeg';
  let outDataUrl = outMime === 'image/png'
    ? canvas.toDataURL('image/png')
    : canvas.toDataURL('image/jpeg', 0.84);

  // Jika masih mendekati batas, turunkan kualitas sedikit
  if (outDataUrl.length > MAX_INCOMING_RAW_BYTES * 1.25 && outMime === 'image/jpeg') {
    outDataUrl = canvas.toDataURL('image/jpeg', 0.68);
  }

  const outBase64 = outDataUrl.includes(',') ? outDataUrl.split(',')[1] : '';
  const estimatedBytes = Math.round((outBase64.length * 3) / 4);
  if (estimatedBytes > MAX_INCOMING_RAW_BYTES) {
    throw new Error(`Ukuran gambar masih terlalu besar (${formatFileSizeLabel(estimatedBytes)}). Kecilkan resolusi terlebih dahulu.`);
  }

  // Thumbnail mungil untuk tampil instan di tabel daftar Surat Masuk
  const thumbMax = 160;
  const tRatio = Math.min(1, thumbMax / targetW, thumbMax / targetH);
  const tCanvas = document.createElement('canvas');
  tCanvas.width = Math.max(1, Math.round(targetW * tRatio));
  tCanvas.height = Math.max(1, Math.round(targetH * tRatio));
  const tCtx = tCanvas.getContext('2d');
  tCtx.fillStyle = '#ffffff';
  tCtx.fillRect(0, 0, tCanvas.width, tCanvas.height);
  tCtx.drawImage(img, 0, 0, tCanvas.width, tCanvas.height);
  const thumbUrl = tCanvas.toDataURL('image/jpeg', 0.65);

  const safeExt = outMime === 'image/png' ? '.png' : '.jpg';
  const cleanName = origName.replace(/\.[^/.]+$/, '') + safeExt;

  return {
    fileName: cleanName,
    mimeType: outMime,
    fileSizeBytes: estimatedBytes,
    fileBase64: outBase64,
    previewThumbDataUrl: thumbUrl,
    localDataUrl: outDataUrl
  };
}

/**
 * Snippet Google Apps Script (Code.gs) untuk disisipkan ke doPost(e) Bridge Drive yang sudah ada
 * (tidak mengubah secret maupun fungsi Surat Keluar yang sudah berjalan).
 */
export const APPS_SCRIPT_INCOMING_SNIPPET = `// 1. Tempel blok IF ini DI DALAM blok try { ... } fungsi doPost(e),
//    TEPAT SEBELUM baris: return jsonOut({ ok: false, error: 'Aksi tidak dikenali: ' + action });
//    (JANGAN ditempel di luar try/finally atau di paling bawah setelah tutup kurung '}' doPost!)
if (action === 'archive_incoming') {
  return handleIncomingArchive_(e);
}
if (action === 'delete_incoming') {
  return handleIncomingDelete_(e);
}

// 2. Tempel kedua fungsi pembantu ini di bagian paling bawah file Code.gs Anda (di luar doPost):
function handleIncomingArchive_(e) {
  var payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  var root = DriveApp.getFolderById(payload.rootFolderId || '19nFdoJ4VOEgFmUe_heAFxoFEt-9U9Crw');
  var subFolders = root.getFoldersByName('04_Surat_Masuk');
  var inFolder = subFolders.hasNext() ? subFolders.next() : root.createFolder('04_Surat_Masuk');

  var item = payload.item || {};
  var mime = payload.mimeType || item.mimeType || 'application/octet-stream';
  var safeKode = String(item.kodeSurat || 'SM').replace(/[\\\\/:*?"<>|]/g, '-');
  var safeNama = String(item.namaSurat || 'Surat_Masuk').replace(/[\\\\/:*?"<>|]/g, '-').slice(0, 60);
  var origName = String(payload.fileName || item.fileName || '');
  var extMatch = origName.match(/(\\.[a-zA-Z0-9]{2,5})$/);
  var ext = extMatch ? extMatch[1].toLowerCase() : (mime.indexOf('pdf') !== -1 ? '.pdf' : (mime.indexOf('png') !== -1 ? '.png' : '.jpg'));
  var fileName = safeKode + '_' + safeNama + ext;
  var bytes = Utilities.base64Decode(payload.fileBase64 || '');
  var blob = Utilities.newBlob(bytes, mime, fileName);
  var file = inFolder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (shareErr) {}

  var sheetUrl = '';
  try {
    var dbIt = root.getFoldersByName('00_Database_Persuratan');
    var dbFolder = dbIt.hasNext() ? dbIt.next() : root;
    var ssFiles = dbFolder.getFilesByName('DB_Persuratan_SMKS_IT_AlHawari');
    if (!ssFiles.hasNext()) {
      ssFiles = root.getFilesByName('DB_Persuratan_SMKS_IT_AlHawari');
    }
    if (ssFiles.hasNext()) {
      var ss = SpreadsheetApp.open(ssFiles.next());
      var sh = ss.getSheetByName('Surat_Masuk') || ss.insertSheet('Surat_Masuk');
      if (sh.getLastRow() === 0) {
        sh.appendRow([
          'ID', 'Tahun', 'No Urut', 'Kode Surat', 'Nama Surat',
          'Nomor Surat Asli', 'Pengirim', 'Tanggal Surat', 'Kategori', 'Jenis File',
          'Nama File', 'Ukuran (Bytes)', 'Catatan', 'Diunggah Oleh', 'Waktu Diterima',
          'Drive File ID', 'Drive Link'
        ]);
        sh.setFrozenRows(1);
        sh.getRange(1, 1, 1, 17).setFontWeight('bold').setBackground('#156c3b').setFontColor('#ffffff');
      }
      sh.appendRow([
        String(item.id || ''), String(item.year || ''), Number(item.sequenceNumber || 0),
        String(item.kodeSurat || '-'), String(item.namaSurat || '-'),
        String(item.nomorSuratAsli || '-'), String(item.pengirim || '-'),
        String(item.tanggalSurat || '-'), String(item.kategori || 'Lainnya'),
        String(item.fileKind || 'image'), file.getName(), Number(item.fileSizeBytes || 0),
        String(item.catatan || ''), String(item.createdBy || payload.actorUsername || '-'),
        String(item.receivedAt || new Date().toISOString()), file.getId(), file.getUrl()
      ]);
      sheetUrl = ss.getUrl();
    }
  } catch (sheetErr) {}

  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    driveFileId: file.getId(),
    driveUrl: file.getUrl(),
    driveDownloadUrl: 'https://drive.google.com/uc?export=download&id=' + file.getId(),
    driveFolderUrl: inFolder.getUrl(),
    rootFolderUrl: root.getUrl(),
    sheetUrl: sheetUrl,
    syncedAt: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

function handleIncomingDelete_(e) {
  var payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  if (payload.driveFileId) {
    try { DriveApp.getFileById(payload.driveFileId).setTrashed(true); } catch (delErr) {}
  }
  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    syncedAt: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}`;

/**
 * Kode Google Apps Script lengkap All-in-One (Surat Keluar, Gudang, Kop, Backup, + Surat Masuk).
 */
export const APPS_SCRIPT_FULL_BRIDGE_CODE = `/**
 * Google Apps Script Backend Bridge â€” AI DocBuilder SMKS IT Al-Hawari
 * Mengelola penyimpanan file .docx, arsip kop, dan Google Sheet master otomatis
 * di Google Drive (tuancraf1@gmail.com) atas permintaan terautentikasi dari Netlify Functions.
 */

var DEFAULT_FOLDER_IDS = {
  ROOT: '19nFdoJ4VOEgFmUe_heAFxoFEt-9U9Crw',
  DATABASE: '1ifQdDSke3-iVxkp3OT5He8Zji_zlJRg_',
  SURAT_KELUAR: '1rcaOdng-xgbJlbJ12gqawSLsWKFRrlz3',
  GUDANG_SURAT: '1bb7vlTvl31e3YDvoE0MlrNiX0NRkFVQ8',
  KOP_LOGO: '1odtooBDSXUnIn7BvR3Qjygg-7MkyveNA'
};

var MONTH_NAMES_ID = [
  '01_Januari', '02_Februari', '03_Maret', '04_April',
  '05_Mei', '06_Juni', '07_Juli', '08_Agustus',
  '09_September', '10_Oktober', '11_November', '12_Desember'
];

function jsonOut(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getScriptProp(key, fallback) {
  var val = PropertiesService.getScriptProperties().getProperty(key);
  return val ? String(val).trim() : (fallback || '');
}

/**
 * Fungsi setup satu kali (otomatis dipanggil via clasp run / saat deploy)
 * untuk menyimpan secret & memastikan Google Sheet master siap.
 */
function configureBridgeProperties(secret, rootFolderId) {
  var props = PropertiesService.getScriptProperties();
  if (secret) props.setProperty('GAS_WEBHOOK_SECRET', String(secret).trim());
  if (rootFolderId) props.setProperty('DRIVE_ROOT_FOLDER_ID', String(rootFolderId).trim());
  var info = ensureDriveStructureAndSheet();
  return { ok: true, info: info };
}

function getOrCreateFolderByName(parentFolder, folderName) {
  var it = parentFolder.getFoldersByName(folderName);
  if (it.hasNext()) return it.next();
  return parentFolder.createFolder(folderName);
}

function resolveFolderByIdOrName(propId, defaultId, parentFolder, fallbackName) {
  var candidates = [propId, defaultId];
  for (var i = 0; i < candidates.length; i++) {
    if (candidates[i]) {
      try {
        var f = DriveApp.getFolderById(candidates[i]);
        if (f && !f.isTrashed()) return f;
      } catch (e) {}
    }
  }
  if (parentFolder && fallbackName) {
    return getOrCreateFolderByName(parentFolder, fallbackName);
  }
  return DriveApp.getRootFolder();
}

function ensureDriveStructureAndSheet() {
  var props = PropertiesService.getScriptProperties();
  var rootFolder = resolveFolderByIdOrName(
    props.getProperty('DRIVE_ROOT_FOLDER_ID'),
    DEFAULT_FOLDER_IDS.ROOT,
    DriveApp.getRootFolder(),
    'Persuratan SMKS IT Al-Hawari'
  );
  props.setProperty('DRIVE_ROOT_FOLDER_ID', rootFolder.getId());

  var dbFolder = resolveFolderByIdOrName(
    props.getProperty('DRIVE_DB_FOLDER_ID'),
    DEFAULT_FOLDER_IDS.DATABASE,
    rootFolder,
    '00_Database_Persuratan'
  );
  var outFolder = resolveFolderByIdOrName(
    props.getProperty('DRIVE_OUTGOING_FOLDER_ID'),
    DEFAULT_FOLDER_IDS.SURAT_KELUAR,
    rootFolder,
    '01_Surat_Keluar'
  );
  var vaultFolder = resolveFolderByIdOrName(
    props.getProperty('DRIVE_VAULT_FOLDER_ID'),
    DEFAULT_FOLDER_IDS.GUDANG_SURAT,
    rootFolder,
    '02_Gudang_Surat_Template'
  );
  var kopFolder = resolveFolderByIdOrName(
    props.getProperty('DRIVE_KOP_FOLDER_ID'),
    DEFAULT_FOLDER_IDS.KOP_LOGO,
    rootFolder,
    '03_Kop_dan_Logo'
  );
  var incomingFolder = resolveFolderByIdOrName(
    props.getProperty('DRIVE_INCOMING_FOLDER_ID'),
    '',
    rootFolder,
    '04_Surat_Masuk'
  );

  var sheetId = props.getProperty('MASTER_SHEET_ID');
  var ss = null;
  if (sheetId) {
    try {
      ss = SpreadsheetApp.openById(sheetId);
    } catch (e) {
      ss = null;
    }
  }
  if (!ss) {
    var existingSheets = dbFolder.getFilesByName('DB_Persuratan_SMKS_IT_AlHawari');
    if (existingSheets.hasNext()) {
      ss = SpreadsheetApp.openById(existingSheets.next().getId());
    } else {
      ss = SpreadsheetApp.create('DB_Persuratan_SMKS_IT_AlHawari');
      var ssFile = DriveApp.getFileById(ss.getId());
      ssFile.moveTo(dbFolder);
    }
    props.setProperty('MASTER_SHEET_ID', ss.getId());
  }

  ensureSheetTabs(ss);

  return {
    rootFolderId: rootFolder.getId(),
    rootFolderUrl: rootFolder.getUrl(),
    dbFolderId: dbFolder.getId(),
    outgoingFolderId: outFolder.getId(),
    vaultFolderId: vaultFolder.getId(),
    kopFolderId: kopFolder.getId(),
    incomingFolderId: incomingFolder.getId(),
    sheetId: ss.getId(),
    sheetUrl: ss.getUrl(),
    rootFolder: rootFolder,
    dbFolder: dbFolder,
    outFolder: outFolder,
    vaultFolder: vaultFolder,
    kopFolder: kopFolder,
    incomingFolder: incomingFolder,
    spreadsheet: ss
  };
}

function ensureSheetTabs(ss) {
  var specs = [
    {
      name: 'Surat_Keluar',
      headers: [
        'ID', 'Tahun', 'No Urut', 'Kode Jenis', 'Nomor Surat',
        'Judul Surat', 'Perihal', 'Penerima / Atas Nama', 'Tanggal Surat',
        'Peran Penandatangan', 'Nama Penandatangan', 'Status QR', 'ID Verifikasi',
        'Dicatat Oleh', 'Waktu Catat', 'Drive File ID', 'Drive Link'
      ]
    },
    {
      name: 'Gudang_Surat',
      headers: [
        'ID', 'Judul', 'Nomor Contoh', 'Kategori', 'Perihal',
        'Acuan AI', 'Diperbarui', 'Drive File ID', 'Drive Link'
      ]
    },
    {
      name: 'Log_Verifikasi',
      headers: [
        'Waktu', 'Aksi', 'ID Verifikasi', 'Nomor Surat', 'Judul',
        'Status', 'Peran Sah', 'Oleh Akun', 'Keterangan'
      ]
    },
    {
      name: 'Kop_Galeri',
      headers: [
        'ID', 'Nama Kop', 'Kategori', 'Drive Image ID', 'Drive Image Link', 'Diperbarui'
      ]
    },
    {
      name: 'Surat_Masuk',
      headers: [
        'ID', 'Tahun', 'No Urut', 'Kode Surat', 'Nama / Judul Surat',
        'Nomor Asli', 'Asal / Pengirim', 'Tanggal Surat', 'Kategori',
        'Jenis Berkas', 'Nama File', 'Ukuran (Bytes)', 'Catatan',
        'Dicatat Oleh', 'Waktu Terima', 'Drive File ID', 'Drive Link'
      ]
    }
  ];

  for (var i = 0; i < specs.length; i++) {
    var spec = specs[i];
    var sheet = ss.getSheetByName(spec.name);
    if (!sheet) {
      sheet = ss.insertSheet(spec.name);
    }
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, spec.headers.length)
        .setFontWeight('bold')
        .setBackground('#156c3b')
        .setFontColor('#ffffff');
    }
  }

  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1 && defaultSheet.getLastRow() <= 1) {
    try { ss.deleteSheet(defaultSheet); } catch (e) {}
  }
}

function upsertSheetRow(sheet, idValue, rowArray) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var r = 0; r < ids.length; r++) {
      if (String(ids[r][0]) === String(idValue)) {
        sheet.getRange(r + 2, 1, 1, rowArray.length).setValues([rowArray]);
        return;
      }
    }
  }
  sheet.appendRow(rowArray);
}

function upsertDriveFile(folder, fileName, bytesOrText, mimeType, existingFileId) {
  var blob = typeof bytesOrText === 'string'
    ? Utilities.newBlob(bytesOrText, mimeType, fileName)
    : Utilities.newBlob(bytesOrText, mimeType, fileName);

  if (existingFileId) {
    try {
      var oldFile = DriveApp.getFileById(existingFileId);
      if (oldFile && !oldFile.isTrashed()) {
        oldFile.setName(fileName);
        oldFile.setTrashed(true);
      }
    } catch (e) {}
  }

  var existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) {
    var f = existing.next();
    try { f.setTrashed(true); } catch (e) {}
  }

  var created = folder.createFile(blob);
  return {
    id: created.getId(),
    name: created.getName(),
    url: created.getUrl(),
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + created.getId()
  };
}

function sanitizeFileSlug(value, fallback) {
  var cleaned = String(value || fallback || 'dokumen')
    .replace(/[^a-zA-Z0-9_\\-\\s]/g, '')
    .trim()
    .replace(/\\s+/g, '_')
    .slice(0, 48);
  return cleaned || 'dokumen';
}

function resolveOutgoingMonthFolder(outRootFolder, yearVal, dateIsoOrText) {
  var y = String(yearVal || new Date().getFullYear()).replace(/[^0-9]/g, '').slice(0, 4) || String(new Date().getFullYear());
  var yearFolder = getOrCreateFolderByName(outRootFolder, y);

  var mIdx = new Date().getMonth();
  if (dateIsoOrText) {
    var d = new Date(dateIsoOrText);
    if (!isNaN(d.getTime())) mIdx = d.getMonth();
  }
  var monthName = MONTH_NAMES_ID[mIdx] || MONTH_NAMES_ID[0];
  return getOrCreateFolderByName(yearFolder, monthName);
}

function doGet(e) {
  var secretConfigured = Boolean(getScriptProp('GAS_WEBHOOK_SECRET', ''));
  return jsonOut({
    ok: true,
    service: 'AI DocBuilder SMKS IT Al-Hawari â€” Drive Backend',
    secretConfigured: secretConfigured,
    timestamp: new Date().toISOString()
  });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    var body = JSON.parse(raw);
    var action = String(body.action || 'ping').trim();
    var props = PropertiesService.getScriptProperties();

    // Setup satu kali jika GAS_WEBHOOK_SECRET belum pernah diinisialisasi
    if (action === 'initialize_bridge') {
      var existingSecret = props.getProperty('GAS_WEBHOOK_SECRET');
      if (existingSecret) {
        return jsonOut({ ok: false, status: 403, error: 'Bridge sudah diinisialisasi sebelumnya.' });
      }
      var incomingSecret = String(body.secret || '').trim();
      var incomingRoot = String(body.rootFolderId || '').trim();
      if (incomingSecret.length < 32 || incomingRoot !== DEFAULT_FOLDER_IDS.ROOT) {
        return jsonOut({ ok: false, status: 400, error: 'Parameter inisialisasi bridge tidak valid.' });
      }
      props.setProperty('GAS_WEBHOOK_SECRET', incomingSecret);
      props.setProperty('DRIVE_ROOT_FOLDER_ID', incomingRoot);
      var initCtx = ensureDriveStructureAndSheet();
      return jsonOut({
        ok: true,
        action: 'initialize_bridge',
        rootFolderId: initCtx.rootFolderId,
        rootFolderUrl: initCtx.rootFolderUrl,
        sheetId: initCtx.sheetId,
        sheetUrl: initCtx.sheetUrl,
        syncedAt: new Date().toISOString()
      });
    }

    var expectedSecret = getScriptProp('GAS_WEBHOOK_SECRET', '');
    if (!expectedSecret || String(body.secret || '') !== expectedSecret) {
      return jsonOut({ ok: false, status: 403, error: 'Unauthorized: secret bridge Google Drive tidak cocok.' });
    }

    var ctx = ensureDriveStructureAndSheet();

    if (action === 'ping' || action === 'setup') {
      return jsonOut({
        ok: true,
        action: action,
        rootFolderId: ctx.rootFolderId,
        rootFolderUrl: ctx.rootFolderUrl,
        sheetId: ctx.sheetId,
        sheetUrl: ctx.sheetUrl,
        timestamp: new Date().toISOString()
      });
    }

    if (action === 'archive_outgoing') {
      var item = body.item || {};
      var targetFolder = resolveOutgoingMonthFolder(ctx.outFolder, item.year, item.recordedAt || item.authorizedAt);
      var seqPad = String(item.sequenceNumber || 0);
      while (seqPad.length < 3) seqPad = '0' + seqPad;
      var baseSlug = seqPad + '_' + sanitizeFileSlug(item.typeCode || 'SURAT', 'SURAT') + '_' + sanitizeFileSlug(item.recipientName || item.title, 'Dokumen');

      var docxResult = null;
      if (body.docxBase64) {
        var docxBytes = Utilities.base64Decode(String(body.docxBase64));
        docxResult = upsertDriveFile(
          targetFolder,
          baseSlug + '.docx',
          docxBytes,
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          item.driveDocxFileId || ''
        );
      }

      var snapshotCopy = JSON.parse(JSON.stringify(item));
      delete snapshotCopy.manageToken;
      var jsonResult = upsertDriveFile(
        targetFolder,
        baseSlug + '.json',
        JSON.stringify(snapshotCopy, null, 2),
        'application/json',
        item.driveJsonFileId || ''
      );

      var sheetOut = ctx.spreadsheet.getSheetByName('Surat_Keluar');
      upsertSheetRow(sheetOut, item.id || baseSlug, [
        String(item.id || ''),
        String(item.year || ''),
        Number(item.sequenceNumber || 0),
        String(item.typeCode || 'DINAS'),
        String(item.nomorSurat || '-'),
        String(item.title || '-'),
        String(item.perihal || '-'),
        String(item.recipientName || '-'),
        String(item.tanggalSurat || '-'),
        String(item.signerRole || 'kepala_sekolah'),
        String(item.signerName || '-'),
        String(item.verificationStatus || 'unverified').toUpperCase(),
        String(item.verificationId || '-'),
        String(item.recordedBy || body.actorUsername || '-'),
        String(item.recordedAt || new Date().toISOString()),
        docxResult ? docxResult.id : (item.driveDocxFileId || jsonResult.id),
        docxResult ? docxResult.url : (item.driveDocxUrl || jsonResult.url)
      ]);

      if (item.verificationId) {
        var logSheet = ctx.spreadsheet.getSheetByName('Log_Verifikasi');
        logSheet.appendRow([
          new Date().toISOString(),
          'CATAT / ARSIP',
          String(item.verificationId || ''),
          String(item.nomorSurat || '-'),
          String(item.title || '-'),
          String(item.verificationStatus || 'sah').toUpperCase(),
          String(item.signerRole || 'kepala_sekolah'),
          String(item.authorizedBy || body.actorUsername || '-'),
          String(item.delegatedNote || '')
        ]);
      }

      return jsonOut({
        ok: true,
        driveDocxFileId: docxResult ? docxResult.id : (item.driveDocxFileId || ''),
        driveDocxUrl: docxResult ? docxResult.url : (item.driveDocxUrl || ''),
        driveJsonFileId: jsonResult.id,
        driveJsonUrl: jsonResult.url,
        driveFolderUrl: targetFolder.getUrl(),
        sheetUrl: ctx.sheetUrl,
        syncedAt: new Date().toISOString()
      });
    }

    if (action === 'update_outgoing_status') {
      var upd = body.item || {};
      var sheetSt = ctx.spreadsheet.getSheetByName('Surat_Keluar');
      upsertSheetRow(sheetSt, upd.id || upd.verificationId, [
        String(upd.id || ''),
        String(upd.year || ''),
        Number(upd.sequenceNumber || 0),
        String(upd.typeCode || 'DINAS'),
        String(upd.nomorSurat || '-'),
        String(upd.title || '-'),
        String(upd.perihal || '-'),
        String(upd.recipientName || '-'),
        String(upd.tanggalSurat || '-'),
        String(upd.signerRole || 'kepala_sekolah'),
        String(upd.signerName || '-'),
        String(upd.verificationStatus || 'unverified').toUpperCase(),
        String(upd.verificationId || '-'),
        String(upd.recordedBy || '-'),
        String(upd.recordedAt || ''),
        String(upd.driveDocxFileId || upd.driveJsonFileId || ''),
        String(upd.driveDocxUrl || upd.driveJsonUrl || '')
      ]);
      var logSt = ctx.spreadsheet.getSheetByName('Log_Verifikasi');
      logSt.appendRow([
        new Date().toISOString(),
        String(body.statusAction || 'UBAH_STATUS'),
        String(upd.verificationId || ''),
        String(upd.nomorSurat || '-'),
        String(upd.title || '-'),
        String(upd.verificationStatus || '').toUpperCase(),
        String(upd.signerRole || ''),
        String(body.actorUsername || '-'),
        String(upd.revokedReason || body.note || '')
      ]);
      return jsonOut({ ok: true, sheetUrl: ctx.sheetUrl, syncedAt: new Date().toISOString() });
    }

    if (action === 'archive_vault') {
      var vItem = body.item || {};
      var vSlug = sanitizeFileSlug(vItem.title || vItem.id, 'Template_Surat') + '_' + String(vItem.id || '').slice(-6);
      var vDocxRes = null;
      if (body.docxBase64) {
        var vBytes = Utilities.base64Decode(String(body.docxBase64));
        vDocxRes = upsertDriveFile(
          ctx.vaultFolder,
          vSlug + '.docx',
          vBytes,
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          vItem.driveFileId || ''
        );
      }
      var vJsonRes = upsertDriveFile(
        ctx.vaultFolder,
        vSlug + '.json',
        JSON.stringify(vItem, null, 2),
        'application/json',
        vItem.driveJsonFileId || ''
      );
      var sheetVault = ctx.spreadsheet.getSheetByName('Gudang_Surat');
      upsertSheetRow(sheetVault, vItem.id || vSlug, [
        String(vItem.id || ''),
        String(vItem.title || '-'),
        String(vItem.nomorSurat || '-'),
        String(vItem.category || 'Surat Dinas'),
        String(vItem.perihal || '-'),
        vItem.isAiReference ? 'YA' : 'TIDAK',
        String(vItem.updatedAt || new Date().toISOString()),
        vDocxRes ? vDocxRes.id : vJsonRes.id,
        vDocxRes ? vDocxRes.url : vJsonRes.url
      ]);
      return jsonOut({
        ok: true,
        driveFileId: vDocxRes ? vDocxRes.id : vJsonRes.id,
        driveUrl: vDocxRes ? vDocxRes.url : vJsonRes.url,
        driveJsonFileId: vJsonRes.id,
        sheetUrl: ctx.sheetUrl,
        syncedAt: new Date().toISOString()
      });
    }

    if (action === 'archive_kop') {
      var kop = body.item || {};
      var kopRes = null;
      if (kop.dataUrl && String(kop.dataUrl).indexOf('base64,') !== -1) {
        var parts = String(kop.dataUrl).split('base64,');
        var mimeMatch = parts[0].match(/data:([^;]+);/);
        var mime = mimeMatch ? mimeMatch[1] : 'image/png';
        var ext = mime.indexOf('jpeg') !== -1 || mime.indexOf('jpg') !== -1 ? '.jpg' : '.png';
        var kBytes = Utilities.base64Decode(parts[1]);
        kopRes = upsertDriveFile(
          ctx.kopFolder,
          sanitizeFileSlug(kop.name || kop.id, 'kop_surat') + ext,
          kBytes,
          mime,
          kop.driveFileId || ''
        );
        var sheetKop = ctx.spreadsheet.getSheetByName('Kop_Galeri');
        upsertSheetRow(sheetKop, kop.id || kop.name, [
          String(kop.id || ''),
          String(kop.name || 'Kop Surat'),
          String(kop.category || 'Kustom'),
          kopRes.id,
          kopRes.url,
          new Date().toISOString()
        ]);
      }
      return jsonOut({
        ok: true,
        driveFileId: kopRes ? kopRes.id : '',
        driveUrl: kopRes ? kopRes.url : '',
        syncedAt: new Date().toISOString()
      });
    }

    if (action === 'backup_snapshot') {
      var snapshotPayload = body.snapshot || {};
      var backupRes = upsertDriveFile(
        ctx.dbFolder,
        'workspace_backup_latest.json',
        JSON.stringify(snapshotPayload, null, 2),
        'application/json',
        ''
      );
      return jsonOut({
        ok: true,
        backupFileId: backupRes.id,
        backupFileUrl: backupRes.url,
        rootFolderUrl: ctx.rootFolderUrl,
        sheetUrl: ctx.sheetUrl,
        syncedAt: new Date().toISOString()
      });
    }

    if (action === 'archive_incoming') {
      var inItem = body.item || {};
      var inFolder = resolveOutgoingMonthFolder(ctx.incomingFolder, inItem.year, inItem.receivedAt || inItem.createdAt);
      var inMime = String(body.mimeType || inItem.mimeType || 'application/octet-stream');
      var kodeSlug = sanitizeFileSlug(inItem.kodeSurat || 'SM', 'SM');
      var namaSlug = sanitizeFileSlug(inItem.namaSurat || 'Surat_Masuk', 'Surat_Masuk');
      var origName = String(body.fileName || inItem.fileName || '');
      var extMatch = origName.match(/(\\.[a-zA-Z0-9]{2,5})$/);
      var ext = extMatch ? extMatch[1].toLowerCase() : (inMime.indexOf('pdf') !== -1 ? '.pdf' : (inMime.indexOf('png') !== -1 ? '.png' : '.jpg'));
      var inFileName = kodeSlug + '_' + namaSlug + ext;

      var inBytes = Utilities.base64Decode(String(body.fileBase64 || ''));
      var inRes = upsertDriveFile(inFolder, inFileName, inBytes, inMime, inItem.driveFileId || '');
      try {
        DriveApp.getFileById(inRes.id).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (shareErr) {}

      var sheetIn = ctx.spreadsheet.getSheetByName('Surat_Masuk');
      if (sheetIn) {
        upsertSheetRow(sheetIn, inItem.id || inFileName, [
          String(inItem.id || ''),
          String(inItem.year || ''),
          Number(inItem.sequenceNumber || 0),
          String(inItem.kodeSurat || '-'),
          String(inItem.namaSurat || '-'),
          String(inItem.nomorSuratAsli || '-'),
          String(inItem.pengirim || '-'),
          String(inItem.tanggalSurat || '-'),
          String(inItem.kategori || 'Lainnya'),
          String(inItem.fileKind || 'image'),
          inRes.name,
          Number(inItem.fileSizeBytes || 0),
          String(inItem.catatan || ''),
          String(inItem.createdBy || body.actorUsername || '-'),
          String(inItem.receivedAt || new Date().toISOString()),
          inRes.id,
          inRes.url
        ]);
      }

      return jsonOut({
        ok: true,
        driveFileId: inRes.id,
        driveUrl: inRes.url,
        driveDownloadUrl: inRes.downloadUrl,
        driveFolderUrl: inFolder.getUrl(),
        rootFolderUrl: ctx.rootFolderUrl,
        sheetUrl: ctx.sheetUrl,
        syncedAt: new Date().toISOString()
      });
    }

    if (action === 'delete_incoming') {
      if (body.driveFileId) {
        try { DriveApp.getFileById(String(body.driveFileId)).setTrashed(true); } catch (delErr) {}
      }
      return jsonOut({ ok: true, syncedAt: new Date().toISOString() });
    }

    return jsonOut({ ok: false, error: 'Aksi tidak dikenali: ' + action });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}`;
