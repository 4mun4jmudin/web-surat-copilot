/**
 * schoolProfile.js - Pengelola Data Profil Sekolah & Kop Surat
 * Tersimpan permanen di IndexedDB. Digunakan sebagai identitas baku seluruh surat.
 */

export const DEFAULT_SCHOOL_PROFILE = {
  government: 'PEMERINTAH DAERAH PROVINSI JAWA BARAT',
  department: 'DINAS PENDIDIKAN',
  subOffice: 'CABANG DINAS PENDIDIKAN WILAYAH XI',
  foundation: 'Yayasan Nunuk Kartini Al-Irsyad',
  schoolName: 'SMKS IT AL-HAWARI',
  npsn: '69758993',
  skOperasional: '425.11/1136-DISDIK (30 April 2013)',
  accreditation: 'Swasta / Akreditasi B',
  address: 'Kp. Curug RT 02 / RW 06, Desa Cipareuan, Kec. Cibiuk',
  city: 'Garut',
  province: 'Jawa Barat',
  postalCode: '44193',
  phone: '+62 895-3266-12155',
  email: 'smkitalhawari@gmail.com',
  website: 'al-hawari.sch.id',
  principalName: 'Dian Nurhadian, S.Pd.',
  principalNip: '-',
  principalRank: 'Kepala SMKS IT Al-Hawari',
  principalTitle: 'Kepala Sekolah',
  tuName: 'Staf Tata Usaha (TU)',
  tuNip: '-',
  tuRank: 'Tata Usaha SMKS IT Al-Hawari',
  tuTitle: 'Kepala Tata Usaha',
  numberingPattern: '421.3/{urut}/SMKITA/{bulanRomawi}/{tahun}',
  majors: [
    'Multimedia (MM)',
    'Agribisnis Tanaman Pangan dan Hortikultura (ATPH)',
    'Desain Pemodelan dan Informasi Bangunan (DPIB)'
  ],
  // Logo placeholder standar Pendidikan / Al-Hawari (Data URL PNG)
  logo: 'data:image/png;base64,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'
};

