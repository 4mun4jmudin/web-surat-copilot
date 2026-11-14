/**
 * sample.js - Dokumen Contoh Laporan PKL SMK Lengkap
 * Sesuai skema Document Model dan siap dirender menjadi .docx
 */

export const samplePklDocument = {
  meta: {
    title: 'LAPORAN PRAKTIK KERJA LAPANGAN DI PT TEKNOLOGI NUSANTARA',
    author: 'Ahmad Fauzi',
    nisn: '0054321987',
    school: 'SMK Negeri 1 Surabaya',
    program: 'Rekayasa Perangkat Lunak',
    year: '2026',
    city: 'Surabaya'
  },
  profile: 'pkl',
  blocks: [
    // --- 1. COVER PAGE ---
    {
      type: 'paragraph',
      runs: [
        { text: 'LAPORAN PRAKTIK KERJA LAPANGAN\nDI PT TEKNOLOGI INOVASI NUSANTARA\n\n', bold: true }
      ]
    },
    {
      type: 'paragraph',
      runs: [
        { text: 'Diajukan untuk Memenuhi Salah Satu Syarat Kelulusan\nProgram Keahlian Rekayasa Perangkat Lunak\n\n\n\n', italic: true }
      ]
    },
    {
      type: 'paragraph',
      runs: [
        { text: 'Disusun Oleh:\n', bold: true },
        { text: 'Nama: Ahmad Fauzi\nNISN: 0054321987\nKelas: XII RPL 2\n\n\n\n' }
      ]
    },
    {
      type: 'paragraph',
      runs: [
        { text: 'SMK NEGERI 1 SURABAYA\nDINAS PENDIDIKAN JAWA TIMUR\n2026', bold: true }
      ]
    },

    // --- 2. SECTION BREAK MENUJU BAGIAN AWAL (ROMAN i, ii) ---
    {
      type: 'sectionBreak',
      numberFormat: 'roman',
      restart: true
    },

    // KATA PENGANTAR
    {
      type: 'heading',
      level: 1,
      text: 'KATA PENGANTAR'
    },
    {
      type: 'paragraph',
      runs: [
        { text: 'Puji dan syukur penulis panjatkan ke hadirat Allah SWT atas segala rahmat dan hidayah-Nya, sehingga penulis dapat menyelesaikan Praktik Kerja Lapangan (PKL) di PT Teknologi Inovasi Nusantara serta menyusun laporan ini tepat pada waktunya.' }
      ]
    },
    {
      type: 'paragraph',
      runs: [
        { text: 'Laporan ini disusun sebagai pertanggungjawaban atas pelaksanaan PKL yang telah dilaksanakan mulai tanggal 1 Januari hingga 31 Maret 2026. Penulis menyadari bahwa keberhasilan pelaksanaan PKL dan penulisan laporan ini tidak lepas dari bimbingan dan dukungan berbagai pihak.' }
      ]
    },

    { type: 'pageBreak' },

    // DAFTAR ISI
    {
      type: 'toc',
      title: 'DAFTAR ISI'
    },

    { type: 'pageBreak' },

    // DAFTAR TABEL & GAMBAR
    {
      type: 'listOfTables',
      title: 'DAFTAR TABEL'
    },
    {
      type: 'listOfFigures',
      title: 'DAFTAR GAMBAR'
    },

    // --- 3. SECTION BREAK MENUJU ISI UTAMA (ARAB 1, 2, 3...) ---
    {
      type: 'sectionBreak',
      numberFormat: 'decimal',
      restart: true
    },

    // BAB I: PENDAHULUAN
    {
      type: 'heading',
      level: 1,
      text: 'PENDAHULUAN'
    },
    {
      type: 'heading',
      level: 2,
      text: 'Latar Belakang Praktik Kerja Lapangan'
    },
    {
      type: 'paragraph',
      runs: [
        { text: 'Praktik Kerja Lapangan (PKL) merupakan program pembelajaran yang dirancang untuk membekali siswa Sekolah Menengah Kejuruan (SMK) dengan ' },
        { text: 'kompetensi praktis di dunia usaha dan dunia industri (DUDI)', bold: true },
        { text: '. Melalui program ini, siswa tidak hanya mengaplikasikan ilmu pemrograman web dan basis data yang dipelajari di kelas, namun juga memahami budaya kerja, etika profesional, dan metode pengembangan perangkat lunak modern seperti Agile Scrum.' }
      ]
    },
    {
      type: 'heading',
      level: 2,
      text: 'Tujuan dan Manfaat PKL'
    },
    {
      type: 'paragraph',
      runs: [
        { text: 'Tujuan diselenggarakannya Praktik Kerja Lapangan antara lain:' }
      ]
    },
    {
      type: 'list',
      ordered: true,
      items: [
        'Meningkatkan keterampilan teknis siswa dalam pengembangan aplikasi berbasis web.',
        'Menumbuhkan etos kerja profesional, tanggung jawab, dan kedisiplinan tinggi.',
        'Mengenal alur kerja kolaboratif menggunakan version control system (Git) di tim engineer.',
        'Menjembatani keterkaitan antara kurikulum sekolah dengan kebutuhan industri teknologi.'
      ]
    },

    // BAB II: PROFIL INDUSTRI DAN PEMBAHASAN
    {
      type: 'heading',
      level: 1,
      text: 'PEMBAHASAN KEGIATAN PRAKTIK'
    },
    {
      type: 'heading',
      level: 2,
      text: 'Struktur Organisasi Perusahaan'
    },
    {
      type: 'paragraph',
      runs: [
        { text: 'Selama melaksanakan kegiatan di PT Teknologi Inovasi Nusantara, penulis ditempatkan di bawah supervisi langsung Senior Web Developer. Struktur koordinasi tim dapat diamati pada bagan berikut ini.' }
      ]
    },
    {
      type: 'figure',
      id: 'fig_org',
      caption: 'Struktur Organisasi Divisi Rekayasa Perangkat Lunak',
      src: 'placeholder_diagram',
      widthCm: 14
    },
    {
      type: 'paragraph',
      runs: [
        { text: 'Sebagaimana ditunjukkan pada ' },
        { text: 'Gambar 1' }, // Cross-ref
        { text: ', koordinasi tugas teknis dilaporkan setiap hari melalui sesi daily stand-up meeting.' }
      ]
    },
    {
      type: 'heading',
      level: 2,
      text: 'Rekapitulasi Kegiatan Teknis'
    },
    {
      type: 'paragraph',
      runs: [
        { text: 'Berikut rekapitulasi pekerjaan utama yang diselesaikan selama durasi pelaksanaan PKL:' }
      ]
    },
    {
      type: 'table',
      id: 'tab_rekap',
      caption: 'Rekapitulasi Proyek dan Hasil Capaian PKL',
      headerRow: true,
      rows: [
        ['No', 'Bulan', 'Modul Pekerjaan', 'Hasil Capaian', 'Status'],
        ['1', 'Januari', 'Setup Database & Auth', 'Implementasi otentikasi JWT dan schema DB', 'Selesai'],
        ['2', 'Februari', 'Pembuatan API Persuratan', 'Endpoint CRUD surat keluar dan verifikasi', 'Selesai'],
        ['3', 'Maret', 'Integrasi Client & QA', 'UI form responsif dan pengujian sistem', 'Selesai']
      ]
    },
    {
      type: 'paragraph',
      runs: [
        { text: 'Rincian jadwal mingguan dari kegiatan di ' },
        { text: 'Tabel 1' },
        { text: ' dipantau menggunakan papan kerja Jira secara transparan.' }
      ]
    },

    // BAB III: PENUTUP
    {
      type: 'heading',
      level: 1,
      text: 'PENUTUP'
    },
    {
      type: 'heading',
      level: 2,
      text: 'Kesimpulan'
    },
    {
      type: 'paragraph',
      runs: [
        { text: 'Berdasarkan pelaksanaan PKL selama tiga bulan, disimpulkan bahwa pengalaman langsung di industri secara signifikan mempercepat pemahaman arsitektur software bersih, pentingnya clean code, serta kemampuan pemecahan masalah secara terstruktur.' }
      ]
    },
    {
      type: 'heading',
      level: 2,
      text: 'Saran'
    },
    {
      type: 'list',
      ordered: false,
      items: [
        'Bagi sekolah: Perlu peningkatan materi framework modern dan integrasi cloud sejak semester awal.',
        'Bagi adik kelas: Manfaatkan waktu luang untuk mempelajari dokumentasi resmi dan git collaboration.',
        'Bagi industri: Program mentoring intensif sangat bermanfaat dan diharapkan terus berlanjut.'
      ]
    },

    { type: 'pageBreak' },

    // DAFTAR PUSTAKA
    {
      type: 'heading',
      level: 1,
      text: 'DAFTAR PUSTAKA'
    },
    {
      type: 'reference',
      id: 'ref1',
      text: 'Dinas Pendidikan Provinsi Jawa Timur. (2025). Panduan Teknis Pelaksanaan Praktik Kerja Lapangan (PKL) Sekolah Menengah Kejuruan. Surabaya: Kemendikbudristek.'
    },
    {
      type: 'reference',
      id: 'ref2',
      text: 'Pressman, R. S., & Maxim, B. R. (2020). Software Engineering: A Practitioner\'s Approach (9th ed.). New York: McGraw-Hill Education.'
    },
    {
      type: 'reference',
      id: 'ref3',
      text: 'Mozilla Developer Network. (2026). Open Packaging Conventions and Web Standards in Word Processing. Diperoleh dari https://developer.mozilla.org/'
    }
  ]
};

