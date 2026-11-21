/**
 * agenda.js - Pengelola Buku Agenda Surat Keluar & Riwayat Dokumen
 * Mendukung pencatatan surat, filter/pencarian, buka kembali, duplikasi, dan ekspor CSV / Excel.
 */

/**
 * @typedef {Object} AgendaItem
 * @property {string} id ID unik
 * @property {number} noUrut Nomor urut agenda
 * @property {string} nomorSurat Nomor surat lengkap
 * @property {string} tanggalSurat Tanggal surat
 * @property {string} jenisSurat Jenis / kategori surat
 * @property {string} perihal Perihal atau keperluan surat
 * @property {string} tujuan Pihak yang dituju atau nama siswa/guru
 * @property {string} penandatangan Nama pejabat penandatangan
 * @property {string} createdAt Waktu pembuatan ISO
 * @property {object} docModel Snapshot Document Model untuk dibuka kembali
 */

/**
 * Menambahkan entri baru ke Buku Agenda
 * @param {AgendaItem[]} agendaList 
 * @param {Omit<AgendaItem, 'id'|'noUrut'|'createdAt'>} entryData 
 * @returns {{ updatedList: AgendaItem[], newItem: AgendaItem }}
 */
export function addAgendaEntry(agendaList = [], entryData) {
  const currentYear = new Date().getFullYear();
  // Hitung nomor urut agenda berikutnya untuk tahun ini
  const thisYearEntries = agendaList.filter(item => {
    const itemYear = new Date(item.createdAt || item.tanggalSurat).getFullYear();
    return itemYear === currentYear;
  });

  const nextNoUrut = thisYearEntries.length + 1;
  const newItem = {
    id: `agenda_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
    noUrut: nextNoUrut,
    nomorSurat: entryData.nomorSurat || '-',
    tanggalSurat: entryData.tanggalSurat || new Date().toLocaleDateString('id-ID'),
    jenisSurat: entryData.jenisSurat || 'Surat Dinas',
    perihal: entryData.perihal || 'Administrasi Sekolah',
    tujuan: entryData.tujuan || '-',
    penandatangan: entryData.penandatangan || 'Kepala Sekolah',
    createdAt: new Date().toISOString(),
    docModel: entryData.docModel || null
  };

  const updatedList = [newItem, ...agendaList];
  return { updatedList, newItem };
}

/**
 * Menghapus entri dari Buku Agenda
 * @param {AgendaItem[]} agendaList 
 * @param {string} id 
 * @returns {AgendaItem[]}
 */
export function removeAgendaEntry(agendaList = [], id) {
  return agendaList.filter(item => item.id !== id);
}

/**
 * Filter / Cari catatan Buku Agenda berdasarkan kata kunci
 * @param {AgendaItem[]} agendaList 
 * @param {string} query 
 * @returns {AgendaItem[]}
 */
export function filterAgendaList(agendaList = [], query = '') {
  if (!query || query.trim() === '') return agendaList;
  const q = query.toLowerCase().trim();

  return agendaList.filter(item => {
    return (
      item.nomorSurat.toLowerCase().includes(q) ||
      item.perihal.toLowerCase().includes(q) ||
      item.tujuan.toLowerCase().includes(q) ||
      item.jenisSurat.toLowerCase().includes(q) ||
      item.penandatangan.toLowerCase().includes(q) ||
      item.tanggalSurat.toLowerCase().includes(q)
    );
  });
}

/**
 * Mengonversi daftar Buku Agenda menjadi string CSV (UTF-8 dengan BOM untuk Excel)
 * @param {AgendaItem[]} agendaList 
 * @param {string} [schoolName]
 * @returns {string}
 */
export function exportAgendaToCsv(agendaList = [], schoolName = 'Sekolah') {
  const headers = [
    'No. Urut',
    'Nomor Surat',
    'Tanggal Surat',
    'Jenis Surat',
    'Perihal / Keperluan',
    'Tujuan / Nama',
    'Penandatangan',
    'Waktu Catat'
  ];

  const rows = agendaList.map((item, idx) => {
    return [
      item.noUrut || (idx + 1),
      cleanCsvField(item.nomorSurat),
      cleanCsvField(item.tanggalSurat),
      cleanCsvField(item.jenisSurat),
      cleanCsvField(item.perihal),
      cleanCsvField(item.tujuan),
      cleanCsvField(item.penandatangan),
      cleanCsvField(item.createdAt ? new Date(item.createdAt).toLocaleString('id-ID') : '')
    ].join(',');
  });

  // Tambahkan BOM UTF-8 (\uFEFF) agar Microsoft Excel membuka aksen dan koma dengan tepat
  const csvContent = '\uFEFF' + [
    `"BUKU AGENDA SURAT KELUAR - ${schoolName.toUpperCase()}"`,
    `"Tanggal Ekspor: ${new Date().toLocaleDateString('id-ID')}"`,
    '',
    headers.map(h => `"${h}"`).join(','),
    ...rows
  ].join('\r\n');

  return csvContent;
}

function cleanCsvField(str) {
  if (str === null || str === undefined) return '""';
  const clean = String(str).replace(/"/g, '""');
  return `"${clean}"`;
}

export const AGENDA_MONTH_LABELS = [
  'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
];

const ID_MONTH_MAP = {
  januari: '01', jan: '01',
  februari: '02', feb: '02', peb: '02',
  maret: '03', mar: '03',
  april: '04', apr: '04',
  mei: '05', may: '05',
  juni: '06', jun: '06',
  juli: '07', jul: '07',
  agustus: '08', agu: '08', aug: '08',
  september: '09', sep: '09', sept: '09',
  oktober: '10', okt: '10', oct: '10',
  november: '11', nov: '11', nop: '11',
  desember: '12', des: '12', dec: '12'
};

/**
 * Mengekstrak { year, month ('01'..'12') } dari format tanggal ISO, YYYY-MM-DD,
 * ataupun string tanggal Indonesia (misal "Garut, 18 November 2026").
 */
export function extractDateYearMonth(rawDate) {
  if (!rawDate) return null;
  const s = String(rawDate).trim();
  if (!s) return null;

  const isoMatch = s.match(/^(\d{4})-(\d{1,2})(?:-(\d{1,2}))?/);
  if (isoMatch) {
    const m = String(Number(isoMatch[2])).padStart(2, '0');
    if (Number(m) >= 1 && Number(m) <= 12) {
      return { year: isoMatch[1], month: m };
    }
  }

  const lower = s.toLowerCase();
  const yearMatch = lower.match(/\b(20\d{2}|19\d{2})\b/);
  if (yearMatch) {
    for (const [mName, mNum] of Object.entries(ID_MONTH_MAP)) {
      if (new RegExp(`\\b${mName}\\b`, 'i').test(lower)) {
        return { year: yearMatch[1], month: mNum };
      }
    }
  }

  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) {
    return {
      year: String(d.getFullYear()),
      month: String(d.getMonth() + 1).padStart(2, '0')
    };
  }
  return null;
}

export function filterLettersByPeriod(items = [], { month = 'all', year = 'all', getDateValue } = {}) {
  const targetYear = String(year || 'all').trim();
  const targetMonth = month && month !== 'all' ? String(month).padStart(2, '0') : 'all';
  if (targetYear === 'all' && targetMonth === 'all') return [...items];

  return items.filter((item) => {
    const raw = typeof getDateValue === 'function'
      ? getDateValue(item)
      : (item.tanggalSurat || item.receivedAt || item.recordedAt || item.createdAt);
    const ym = extractDateYearMonth(raw) || extractDateYearMonth(item.recordedAt || item.receivedAt || item.createdAt);
    if (!ym) return false;
    if (targetYear !== 'all' && ym.year !== targetYear) return false;
    if (targetMonth !== 'all' && ym.month !== targetMonth) return false;
    return true;
  });
}

export function collectAvailableAgendaYears(items = [], getDateValue) {
  const set = new Set([String(new Date().getFullYear())]);
  for (const item of items) {
    const raw = typeof getDateValue === 'function'
      ? getDateValue(item)
      : (item.tanggalSurat || item.receivedAt || item.recordedAt || item.createdAt);
    const ym = extractDateYearMonth(raw) || extractDateYearMonth(item.recordedAt || item.receivedAt || item.createdAt);
    if (ym?.year) set.add(ym.year);
  }
  return [...set].sort((a, b) => Number(b) - Number(a));
}

export function formatPeriodLabel(month = 'all', year = 'all') {
  const mIdx = month && month !== 'all' ? Number(month) - 1 : -1;
  const mLabel = mIdx >= 0 && mIdx < 12 ? AGENDA_MONTH_LABELS[mIdx] : 'Semua Bulan';
  const yLabel = year && year !== 'all' ? `Tahun ${year}` : 'Semua Tahun';
  return `${mLabel} · ${yLabel}`;
}

/**
 * Mengekspor daftar Surat Keluar resmi ke format CSV (UTF-8 BOM, siap buka di Microsoft Excel).
 */
export function exportOutgoingLettersCsv(items = [], { schoolName = 'SMK IT Al-Hawari', periodLabel = 'Semua Periode' } = {}) {
  const headers = [
    'No. Urut',
    'Nomor Surat',
    'Kode Jenis',
    'Tanggal Surat',
    'Judul Surat',
    'Perihal',
    'Atas Nama / Penerima',
    'Penandatangan',
    'Status Verifikasi',
    'Link Drive'
  ];

  const rows = items.map((item, idx) => [
    item.sequenceNumber || (idx + 1),
    cleanCsvField(item.nomorSurat || '-'),
    cleanCsvField(item.typeCode || '-'),
    cleanCsvField(item.tanggalSurat || '-'),
    cleanCsvField(item.title || '-'),
    cleanCsvField(item.perihal || '-'),
    cleanCsvField(item.recipientName || '-'),
    cleanCsvField(item.signerName || '-'),
    cleanCsvField(item.isVerifiedSah || item.verificationStatus === 'sah' ? 'SAH (QR Terverifikasi)' : (item.verificationStatus || 'Manual')),
    cleanCsvField(item.driveDocxUrl || item.driveJsonUrl || '-')
  ].join(','));

  return '\uFEFF' + [
    `"BUKU AGENDA SURAT KELUAR - ${String(schoolName).toUpperCase()}"`,
    `"Periode: ${periodLabel} | Diekspor: ${new Date().toLocaleDateString('id-ID')}"`,
    '',
    headers.map(h => `"${h}"`).join(','),
    ...rows
  ].join('\r\n');
}

/**
 * Mengekspor daftar Surat Masuk resmi ke format CSV (UTF-8 BOM, siap buka di Microsoft Excel).
 */
export function exportIncomingLettersCsv(items = [], { schoolName = 'SMK IT Al-Hawari', periodLabel = 'Semua Periode' } = {}) {
  const headers = [
    'No',
    'Kode Surat Masuk',
    'Nomor Surat Asli',
    'Tanggal Surat',
    'Nama / Judul Surat',
    'Asal / Pengirim',
    'Kategori',
    'Jenis Berkas',
    'Status Drive',
    'Catatan / Disposisi',
    'Link Drive'
  ];

  const rows = items.map((item, idx) => [
    idx + 1,
    cleanCsvField(item.kodeSurat || '-'),
    cleanCsvField(item.nomorSuratAsli || '-'),
    cleanCsvField(item.tanggalSurat || '-'),
    cleanCsvField(item.namaSurat || '-'),
    cleanCsvField(item.pengirim || '-'),
    cleanCsvField(item.kategori || '-'),
    cleanCsvField(item.fileKind === 'pdf' ? 'PDF' : (item.fileKind === 'photo' ? 'Foto Kamera' : 'Gambar')),
    cleanCsvField(item.driveSynced ? 'Tersimpan di Drive' : 'Lokal / Belum Sync'),
    cleanCsvField(item.catatan || '-'),
    cleanCsvField(item.driveUrl || '-')
  ].join(','));

  return '\uFEFF' + [
    `"BUKU AGENDA SURAT MASUK - ${String(schoolName).toUpperCase()}"`,
    `"Periode: ${periodLabel} | Diekspor: ${new Date().toLocaleDateString('id-ID')}"`,
    '',
    headers.map(h => `"${h}"`).join(','),
    ...rows
  ].join('\r\n');
}

function escapeAgendaHtml(str = '') {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Menyusun dokumen HTML siap cetak resmi (A4 Landscape) untuk Buku Agenda Surat Masuk / Keluar.
 */
export function buildPrintableAgendaHtml({
  kind = 'outgoing',
  school = {},
  periodLabel = 'Semua Periode',
  items = []
} = {}) {
  const isIncoming = kind === 'incoming';
  const titleText = isIncoming ? 'BUKU AGENDA SURAT MASUK' : 'BUKU AGENDA SURAT KELUAR';
  const schoolName = school.schoolName || 'SMK IT Al-Hawari';
  const address = school.address || 'Kp. Curug RT/RW 002/006 Desa Cipareuan Kec. Cibiuk Kab. Garut';
  const principalTitle = school.principalTitle || 'Kepala Sekolah';
  const principalName = school.principalName || 'Dian Nurhadian, S.Pd.';
  const todayStr = new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });

  const tableHead = isIncoming
    ? '<tr><th style="width:4%">No</th><th style="width:13%">Kode Agenda</th><th style="width:15%">Nomor Asli</th><th style="width:11%">Tanggal</th><th style="width:22%">Nama / Perihal Surat</th><th style="width:16%">Pengirim</th><th style="width:10%">Kategori</th><th style="width:9%">Catatan</th></tr>'
    : '<tr><th style="width:4%">No</th><th style="width:18%">Nomor Surat</th><th style="width:11%">Tanggal</th><th style="width:24%">Judul / Perihal Surat</th><th style="width:18%">Penerima / Tujuan</th><th style="width:14%">Penandatangan</th><th style="width:11%">Status</th></tr>';

  const tableRows = items.length === 0
    ? `<tr><td colspan="${isIncoming ? 8 : 7}" style="text-align:center; padding:18px; color:#64748b;">Tidak ada catatan surat pada periode ${escapeAgendaHtml(periodLabel)}.</td></tr>`
    : items.map((item, idx) => {
        if (isIncoming) {
          return `<tr>
            <td style="text-align:center">${idx + 1}</td>
            <td><strong>${escapeAgendaHtml(item.kodeSurat || '-')}</strong></td>
            <td>${escapeAgendaHtml(item.nomorSuratAsli || '-')}</td>
            <td>${escapeAgendaHtml(item.tanggalSurat || '-')}</td>
            <td><strong>${escapeAgendaHtml(item.namaSurat || '-')}</strong></td>
            <td>${escapeAgendaHtml(item.pengirim || '-')}</td>
            <td>${escapeAgendaHtml(item.kategori || '-')}</td>
            <td>${escapeAgendaHtml(item.catatan || '-')}</td>
          </tr>`;
        }
        const statusStr = (item.isVerifiedSah || item.verificationStatus === 'sah')
          ? 'SAH (QR)'
          : item.verificationStatus === 'pending_acc'
          ? 'Menunggu ACC'
          : item.verificationStatus === 'rejected'
          ? 'Ditolak'
          : 'Manual';
        return `<tr>
          <td style="text-align:center">${item.sequenceNumber || (idx + 1)}</td>
          <td><strong>${escapeAgendaHtml(item.nomorSurat || '-')}</strong></td>
          <td>${escapeAgendaHtml(item.tanggalSurat || '-')}</td>
          <td><strong>${escapeAgendaHtml(item.title || '-')}</strong><br><small>${escapeAgendaHtml(item.perihal || '-')}</small></td>
          <td>${escapeAgendaHtml(item.recipientName || '-')}</td>
          <td>${escapeAgendaHtml(item.signerName || '-')}</td>
          <td style="text-align:center">${escapeAgendaHtml(statusStr)}</td>
        </tr>`;
      }).join('');

  return `<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <title>${escapeAgendaHtml(titleText)} - ${escapeAgendaHtml(schoolName)}</title>
  <style>
    @page { size: A4 landscape; margin: 1.4cm; }
    body { font-family: "Times New Roman", serif; color: #0f172a; margin: 0; font-size: 11pt; }
    .kop { text-align: center; border-bottom: 3px double #0f172a; padding-bottom: 8px; margin-bottom: 14px; }
    .kop h1 { margin: 0; font-size: 15pt; letter-spacing: 0.5px; text-transform: uppercase; }
    .kop p { margin: 2px 0 0 0; font-size: 10pt; color: #334155; }
    .doc-title { text-align: center; margin-bottom: 12px; }
    .doc-title h2 { margin: 0; font-size: 13pt; text-decoration: underline; }
    .doc-title div { font-size: 10.5pt; margin-top: 3px; color: #334155; }
    table { width: 100%; border-collapse: collapse; margin-top: 8px; font-size: 10pt; }
    th, td { border: 1px solid #334155; padding: 6px 7px; vertical-align: top; text-align: left; }
    th { background: #e2e8f0; text-align: center; font-weight: 700; }
    .sig-row { margin-top: 22px; display: flex; justify-content: flex-end; }
    .sig-box { width: 260px; text-align: left; font-size: 10.5pt; }
  </style>
</head>
<body>
  <div class="kop">
    <h1>${escapeAgendaHtml(schoolName)}</h1>
    <p>${escapeAgendaHtml(address)}</p>
  </div>
  <div class="doc-title">
    <h2>${escapeAgendaHtml(titleText)}</h2>
    <div>Periode: ${escapeAgendaHtml(periodLabel)} · Total: ${items.length} surat</div>
  </div>
  <table>
    <thead>${tableHead}</thead>
    <tbody>${tableRows}</tbody>
  </table>
  <div class="sig-row">
    <div class="sig-box">
      ${escapeAgendaHtml(school.city || 'Garut')}, ${escapeAgendaHtml(todayStr)}<br>
      ${escapeAgendaHtml(principalTitle)},<br><br><br><br>
      <strong><u>${escapeAgendaHtml(principalName)}</u></strong>
    </div>
  </div>
</body>
</html>`;
}
