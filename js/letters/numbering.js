/**
 * numbering.js - Sistem Penomoran Surat Otomatis Administrasi Sekolah
 * Mendukung pola penomoran dinamis, konversi bulan Romawi, dan counter per tahun.
 */

// Konversi angka bulan (1-12) ke angka Romawi baku surat dinas
const ROMAN_MONTHS = ['', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'];

/**
 * Mengonversi bulan menjadi angka Romawi
 * @param {number|string} month 1 - 12
 * @returns {string} Contoh: 'XI'
 */
export function toRomanMonth(month) {
  const m = parseInt(month, 10);
  return ROMAN_MONTHS[m] || 'I';
}

/**
 * Format angka urut dengan padding nol di depan
 * @param {number} num 
 * @param {number} padLength Default 3 (contoh: 1 -> "001", 85 -> "085")
 * @returns {string}
 */
export function formatSequenceNumber(num, padLength = 3) {
  return String(num).padStart(padLength, '0');
}

/**
 * Kode Klasifikasi Baku Arsip Surat Dinas Pendidikan (Permendagri / Kemendikbudristek)
 */
export const LETTER_CLASSIFICATION_CODES = {
  surat_keterangan: '421.3', // Kesiswaan
  surat_undangan: '005',     // Undangan Dinas
  surat_tugas: '800',        // Kepegawaian / Penugasan
  surat_dispensasi: '421.3', // Kesiswaan
  surat_pemberitahuan: '421',// Umum Pendidikan
  surat_permohonan: '028',   // Permohonan
  berita_acara: '000.5',     // Berita Acara
  notulen_rapat: '005.1',    // Notulen
  daftar_hadir: '005.2',     // Daftar Hadir
  sertifikat: '421.4'        // Piagam / Sertifikat
};

/**
 * Kode Jenis & Unit Surat Sekolah (Format: {urut}/{jenis}.SMKITA/{bulanRomawi}/{tahun})
 */
export const LETTER_TYPE_CODES = {
  UND: { code: 'UND', prefix: '005', label: 'Surat Undangan (UND)', defaultPerihal: 'Undangan Rapat' },
  TU:  { code: 'TU',  prefix: '421', label: 'Surat Tata Usaha / Administrasi (TU)', defaultPerihal: 'Administrasi Tata Usaha' },
  KEP: { code: 'KEP', prefix: '421', label: 'Surat Kepala Sekolah / SK (KEP)', defaultPerihal: 'Keputusan Kepala Sekolah' },
  KET: { code: 'KET', prefix: '421.3', label: 'Surat Keterangan Siswa/Guru (KET)', defaultPerihal: 'Keterangan Siswa Aktif' },
  ST:  { code: 'ST',  prefix: '800', label: 'Surat Tugas / Penugasan (ST)', defaultPerihal: 'Penugasan Dinas' },
  PGL: { code: 'PGL', prefix: '421.3', label: 'Surat Panggilan Ortu / Siswa (PGL)', defaultPerihal: 'Panggilan Orang Tua / Wali Siswa' },
  PBR: { code: 'PBR', prefix: '421', label: 'Surat Pemberitahuan / Edaran (PBR)', defaultPerihal: 'Pemberitahuan Kegiatan Sekolah' },
  PRM: { code: 'PRM', prefix: '028', label: 'Surat Permohonan (PRM)', defaultPerihal: 'Permohonan Kerjasama / Izin' },
  DISP:{ code: 'DISP',prefix: '421.3', label: 'Surat Dispensasi (DISP)', defaultPerihal: 'Dispensasi Kegiatan Siswa' },
  BA:  { code: 'BA',  prefix: '000.5', label: 'Berita Acara (BA)', defaultPerihal: 'Berita Acara' },
  REK: { code: 'REK', prefix: '421.3', label: 'Surat Rekomendasi / Pengantar (REK)', defaultPerihal: 'Rekomendasi' },
  PKL: { code: 'PKL', prefix: '421.5', label: 'Surat Prakerin / PKL / Hubin (PKL)', defaultPerihal: 'Praktik Kerja Lapangan (PKL)' },
  KUR: { code: 'KUR', prefix: '421.1', label: 'Surat Bidang Kurikulum (KUR)', defaultPerihal: 'Kegiatan Akademik & Kurikulum' },
  KES: { code: 'KES', prefix: '421.3', label: 'Surat Bidang Kesiswaan / OSIS (KES)', defaultPerihal: 'Kegiatan Kesiswaan' }
};

/**
 * Mendeteksi kode jenis surat (UND, TU, KEP, KET, ST, PGL, PBR, PRM, DISP, BA, REK, PKL, KUR, KES)
 * secara otomatis dari teks perintah, judul surat, atau perihal.
 * @param {string} text 
 * @returns {string} Kode jenis surat (contoh: 'UND', 'TU', 'KEP', dll.)
 */
export function detectLetterTypeCode(text = '') {
  const s = String(text || '').toLowerCase();

  if (/\bundangan\b|\bMengundang\b|\brapat\s+pleno\b|\brapat\s+dinas\b|\brapat\s+orang\s+tua\b|\brapat\s+komite\b/i.test(s)) {
    return 'UND';
  }
  if (/\btata\s+usaha\b|\bdari\s+tu\b|\bstaf\s+tu\b|\bka\.?\s*tu\b|\badministrasi\s+tu\b|\bkode\s+tu\b/i.test(s)) {
    return 'TU';
  }
  if (/\bkeputusan\b|\bsk\s+kepala\b|\bkepala\s+sekolah\b|\bkebijakan\s+kepala\b|\bpengangkatan\b|\bpenetapan\b|\bkode\s+kep\b/i.test(s) && !/\bmenerangkan\s+bahwa\b/i.test(s)) {
    return 'KEP';
  }
  if (/\bpanggilan\b|\bmemanggil\b|\bkonseling\b|\bbimbingan\s+konseling\b|\bpelanggaran\b|\bsp\s*1\b|\bsp\s*2\b/i.test(s)) {
    return 'PGL';
  }
  if (/\bsurat\s+tugas\b|\bmenugaskan\b|\bpenugasan\b|\bbimtek\b|\bdelegasi\b|\bsppd\b/i.test(s)) {
    return 'ST';
  }
  if (/\bdispensasi\b|\bizin\s+tidak\s+mengikuti\b|\bizin\s+lomba\b/i.test(s)) {
    return 'DISP';
  }
  if (/\bberita\s+acara\b|\bserah\s+terima\b/i.test(s)) {
    return 'BA';
  }
  if (/\bpemberitahuan\b|\bedaran\b|\bpengumuman\b|\blibur\b|\binfo\s+kegiatan\b/i.test(s)) {
    return 'PBR';
  }
  if (/\bpermohonan\b|\bpengajuan\b|\bmemohon\b|\bbantuan\b|\baudiensi\b|\bnarasumber\b/i.test(s)) {
    return 'PRM';
  }
  if (/\brekomendasi\b|\bpengantar\b/i.test(s)) {
    return 'REK';
  }
  if (/\bprakerin\b|\bpkl\b|\bpraktik\s+kerja\b|\bmagang\b|\bhubin\b|\bdudi\b/i.test(s)) {
    return 'PKL';
  }
  if (/\bkurikulum\b|\bujian\b|\bpenilaian\s+akhir\b|\bpas\b|\bpat\b|\bakademik\b/i.test(s)) {
    return 'KUR';
  }
  if (/\bkesiswaan\b|\bosis\b|\bekstrakurikuler\b|\bpramuka\b|\bldks\b/i.test(s)) {
    return 'KES';
  }
  if (/\bketerangan\b|\bsiswa\s+aktif\b|\bkelakuan\s+baik\b|\bskkb\b|\bbebas\s+pustaka\b/i.test(s)) {
    return 'KET';
  }

  return 'KEP';
}

/**
 * Memformat atau menormalisasi nomor surat agar selalu menyertakan kode jenis surat yang tepat
 * Contoh hasil: "050/UND.SMKITA/XI/2026", "050/TU.SMKITA/XI/2026", "050/KEP.SMKITA/XI/2026"
 * @param {object} params
 * @param {string} [params.rawNomor] Nomor surat dari input / AI
 * @param {string} [params.contextText] Teks judul / prompt / perihal untuk deteksi otomatis kode
 * @param {string} [params.forcedTypeCode] Kode jenis surat spesifik bila dipilih langsung (mis. 'UND', 'TU', 'KEP')
 * @param {string} [params.schoolCode='SMKITA'] Kode sekolah baku
 * @param {Date} [params.date]
 * @returns {string} Nomor surat lengkap (tanpa awalan "Nomor: ")
 */
export function formatSmartNomorSurat({
  rawNomor = '',
  contextText = '',
  forcedTypeCode = '',
  schoolCode = 'SMKITA',
  date = new Date()
} = {}) {
  const d = date instanceof Date && !isNaN(date.getTime()) ? date : new Date();
  const romanMonth = toRomanMonth(d.getMonth() + 1);
  const year = d.getFullYear().toString();
  const typeCode = (forcedTypeCode || detectLetterTypeCode(`${contextText} ${rawNomor}`)).toUpperCase();

  let clean = String(rawNomor || '').trim().replace(/^(nomor|no\.?)\s*:\s*/i, '').trim();

  // Jika kosong atau masih placeholder, rakit nomor surat baru sesuai kode jenis
  if (!clean || clean === '-' || /\[.*\]/.test(clean)) {
    return `050/${typeCode}.${schoolCode}/${romanMonth}/${year}`;
  }

  // Normalisasi variasi kode sekolah lama (SMKIT, SMKS.IT-AH, SMK-IT, dll) menjadi SMKITA
  clean = clean.replace(/\b(SMKS\.?IT-?AH|SMK-?IT|SMKIT)\b/gi, schoolCode);

  // Cek apakah sudah memiliki format {KODE}.{SCHOOL} atau {KODE}/{SCHOOL}
  const knownCodesPattern = Object.keys(LETTER_TYPE_CODES).join('|');
  const hasTypeDotSchool = new RegExp(`\\b(${knownCodesPattern})[\\.\\/-]${schoolCode}\\b`, 'i');

  if (hasTypeDotSchool.test(clean)) {
    if (forcedTypeCode) {
      // Ganti kode jenis lama dengan forcedTypeCode
      return clean.replace(hasTypeDotSchool, `${typeCode}.${schoolCode}`);
    }
    // Pastikan kode jenis & sekolah huruf besar rapi (misal und.SMKITA -> UND.SMKITA)
    return clean.replace(hasTypeDotSchool, (m, c) => `${c.toUpperCase()}.${schoolCode}`);
  }

  // Jika ada SMKITA tapi belum ada kode jenis di depannya (misal: "421.3 / 085 / SMKITA / XI / 2026" atau "050/SMKITA/XI/2026")
  if (new RegExp(`\\b${schoolCode}\\b`, 'i').test(clean)) {
    return clean.replace(new RegExp(`\\b${schoolCode}\\b`, 'i'), `${typeCode}.${schoolCode}`);
  }

  return `050/${typeCode}.${schoolCode}/${romanMonth}/${year}`;
}

/**
 * Menghasilkan nomor surat berdasarkan pola profil sekolah dan counter
 * @param {object} params
 * @param {string} [params.pattern] Pola nomor surat, contoh: '421.3/{urut}/SMKITA/{bulanRomawi}/{tahun}'
 * @param {number} params.counter Nomor urut (integer)
 * @param {string} [params.templateId] ID template surat untuk mengambil kode klasifikasi
 * @param {Date|string} [params.date] Tanggal surat
 * @param {number} [params.padLength=3] Jumlah digit padding
 * @returns {string} Nomor surat siap pakai
 */
export function generateLetterNumber({
  pattern = '421.3/{urut}/SMKITA/{bulanRomawi}/{tahun}',
  counter = 1,
  templateId = 'surat_keterangan',
  date = new Date(),
  padLength = 3
}) {
  const d = date instanceof Date ? date : new Date(date);
  const validDate = isNaN(d.getTime()) ? new Date() : d;

  const year = validDate.getFullYear().toString();
  const monthNum = validDate.getMonth() + 1;
  const romanMonth = toRomanMonth(monthNum);
  const paddedMonth = String(monthNum).padStart(2, '0');
  const formattedUrut = formatSequenceNumber(counter, padLength);
  const letterCode = LETTER_CLASSIFICATION_CODES[templateId] || '421';
  const typeCode = detectLetterTypeCode(templateId);

  let result = pattern;
  result = result.replace(/\{urut\}/g, formattedUrut);
  result = result.replace(/\{kode\}/g, letterCode);
  result = result.replace(/\{jenis\}/g, typeCode);
  result = result.replace(/\{bulanRomawi\}/g, romanMonth);
  result = result.replace(/\{bulan\}/g, paddedMonth);
  result = result.replace(/\{tahun\}/g, year);

  return result;
}

