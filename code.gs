/**
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
    .replace(/[^a-zA-Z0-9_\-\s]/g, '')
    .trim()
    .replace(/\s+/g, '_')
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
      var extMatch = origName.match(/(\.[a-zA-Z0-9]{2,5})$/);
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
}
