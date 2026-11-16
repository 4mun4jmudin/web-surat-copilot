/**
 * defaultKop.js - Default Banner Kop Surat Berformat PNG & SVG dan Galeri Kop Bawaan
 */

export function createKopSvg({
  provinsi = 'PEMERINTAH DAERAH PROVINSI JAWA BARAT',
  dinas = 'DINAS PENDIDIKAN',
  cabang = 'CABANG DINAS PENDIDIKAN WILAYAH XI',
  sekolah = 'SMK IT AL-HAWARI',
  alamat = 'Kp. Curug RT 02/RW 06, Desa Cipareuan, Kec. Cibiuk, Kab. Garut 44193',
  kontak = 'Website: al-hawari.sch.id | Email: smkitalhawari@gmail.com | WA: +62 895-3266-12155',
  primaryColor = '#15803d',
  secondaryColor = '#ca8a04',
  badgeText = 'IT'
} = {}) {
  return `
<svg xmlns="http://www.w3.org/2000/svg" width="900" height="180" viewBox="0 0 900 180">
  <rect width="900" height="180" fill="#ffffff"/>
  <g transform="translate(35, 15)">
    <circle cx="50" cy="50" r="46" fill="${primaryColor}" stroke="${secondaryColor}" stroke-width="4"/>
    <path d="M50,18 L62,44 L50,38 L38,44 Z" fill="#fef3c7"/>
    <path d="M50,38 L68,68 L50,62 L32,68 Z" fill="#ffffff"/>
    <path d="M50,62 L74,84 L50,76 L26,84 Z" fill="#86efac"/>
    <circle cx="50" cy="18" r="5" fill="#ca8a04"/>
    <text x="50" y="65" font-family="'Times New Roman', serif" font-size="11" font-weight="bold" fill="#ffffff" text-anchor="middle">${badgeText}</text>
  </g>
  <g text-anchor="middle" font-family="'Times New Roman', serif" fill="#000000">
    <text x="500" y="32" font-size="16" font-weight="bold" letter-spacing="0.5">${provinsi}</text>
    <text x="500" y="54" font-size="16" font-weight="bold" letter-spacing="0.5">${dinas}</text>
    ${cabang ? `<text x="500" y="74" font-size="13.5" font-weight="bold">${cabang}</text>` : ''}
    <text x="500" y="104" font-size="23" font-weight="bold" letter-spacing="1">${sekolah}</text>
    <text x="500" y="126" font-size="12">${alamat}</text>
    <text x="500" y="142" font-size="11.5">${kontak}</text>
  </g>
  <line x1="20" y1="160" x2="880" y2="160" stroke="#000000" stroke-width="3"/>
  <line x1="20" y1="165" x2="880" y2="165" stroke="#000000" stroke-width="1"/>
</svg>
  `.trim();
}

export const DEFAULT_KOP_SVG = createKopSvg();
export const DEFAULT_KOP_DATA_URL = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(DEFAULT_KOP_SVG)}`;

export const DEFAULT_KOP_PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAlgAAADIAQMAAADv0mK5AAAABlBMVEX///8AAABVwtN+AAAAAXRSTlMAQObYZgAAAFtJREFUeNrtwTEBAAAAwqD1T20ND6AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAOBvAiUAAAEfQy1PAAAAAElFTkSuQmCC';

// Daftar Template Galeri Kop Bawaan Sistem
export const DEFAULT_KOP_GALLERY = [
  {
    id: 'kop-smk-dinas',
    name: 'Kop Resmi SMK IT Al-Hawari (Dinas Pendidikan)',
    category: 'Resmi Dinas',
    isDefault: true,
    dataUrl: DEFAULT_KOP_DATA_URL,
    description: 'Format baku Pemerintah Provinsi Jawa Barat, Dinas Pendidikan Wilayah XI Garut, dan SMK IT Al-Hawari.'
  },
  {
    id: 'kop-yayasan-al-irsyad',
    name: 'Kop Yayasan Nunuk Kartini Al-Irsyad',
    category: 'Yayasan',
    isDefault: false,
    dataUrl: `data:image/svg+xml;charset=utf-8,${encodeURIComponent(createKopSvg({
      provinsi: 'YAYASAN NUNUK KARTINI AL-IRSYAD',
      dinas: 'AKTA NOTARIS & KEMENKUMHAM RI',
      cabang: 'SK OPERASIONAL: 425.11/1136-DISDIK',
      sekolah: 'SMK ISLAM TERPADU AL-HAWARI',
      badgeText: 'YYS',
      primaryColor: '#0f766e',
      secondaryColor: '#f59e0b',
      alamat: 'Kp. Curug RT 02 / RW 06, Desa Cipareuan, Kec. Cibiuk, Garut 44193',
      kontak: 'Website: al-hawari.sch.id | Email: smkitalhawari@gmail.com | WA: +62 895-3266-12155'
    }))}`,
    description: 'Format kop naungan Yayasan Nunuk Kartini Al-Irsyad untuk surat keputusan yayasan dan kelembagaan.'
  },
  {
    id: 'kop-komite-sekolah',
    name: 'Kop Komite Sekolah SMK IT Al-Hawari',
    category: 'Komite',
    isDefault: false,
    dataUrl: `data:image/svg+xml;charset=utf-8,${encodeURIComponent(createKopSvg({
      provinsi: 'KOMITE SEKOLAH',
      dinas: 'SMK IT AL-HAWARI CIBIUK GARUT',
      cabang: 'FORUM KERJASAMA MASYARAKAT & WALI MURID',
      sekolah: 'PENGURUS KOMITE SEKOLAH',
      badgeText: 'KOMITE',
      primaryColor: '#7c3aed',
      secondaryColor: '#ec4899',
      alamat: 'Kp. Curug RT 02 / RW 06, Desa Cipareuan, Kec. Cibiuk, Garut',
      kontak: 'Email: komite@al-hawari.sch.id | Kontak: +62 895-3266-12155'
    }))}`,
    description: 'Khusus untuk surat undangan rapat komite, edaran iuran partisipasi, dan komunikasi wali murid.'
  },
  {
    id: 'kop-osis-mpk',
    name: 'Kop OSIS & MPK SMK IT Al-Hawari',
    category: 'Kesiswaan',
    isDefault: false,
    dataUrl: `data:image/svg+xml;charset=utf-8,${encodeURIComponent(createKopSvg({
      provinsi: 'ORGANISASI SISWA INTRA SEKOLAH (OSIS)',
      dinas: 'MAJELIS PERMUSYAWARATAN KELAS (MPK)',
      cabang: '',
      sekolah: 'SMK IT AL-HAWARI',
      badgeText: 'OSIS',
      primaryColor: '#15803d',
      secondaryColor: '#1d4ed8',
      alamat: 'Sekretariat OSIS/MPK, Kp. Curug RT 02 / RW 06, Cibiuk, Garut',
      kontak: 'Instagram: @smkitalhawari | Email: osis@al-hawari.sch.id'
    }))}`,
    description: 'Kop resmi kegiatan siswa, perlombaan sekolah, class meeting, dan surat mandat delegasi OSIS.'
  }
];
