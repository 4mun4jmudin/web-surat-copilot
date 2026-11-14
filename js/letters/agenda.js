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
