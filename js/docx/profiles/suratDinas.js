/**
 * suratDinas.js - Profil Kertas & Tipografi Standar Surat Dinas Sekolah
 * Mendukung A4 dan F4/Folio dengan margin dinas resmi.
 */

export const suratDinasProfile = {
  id: 'suratDinas',
  name: 'Surat Dinas Sekolah',
  description: 'Standar format surat resmi instansi pendidikan (A4 / F4 Folio, Times New Roman 12pt)',

  // Pilihan ukuran kertas
  pageSizes: {
    a4: {
      name: 'A4 (21,0 x 29,7 cm)',
      width: 11906,  // 210 mm
      height: 16838, // 297 mm
      orientation: 'portrait'
    },
    f4: {
      name: 'F4 / Folio (21,5 x 33,0 cm)',
      width: 12189,  // 215 mm
      height: 18709, // 330 mm
      orientation: 'portrait'
    },
    letter: {
      name: 'Letter / Kuarto (21,59 x 27,94 cm)',
      width: 12240,  // 8.5 in
      height: 15840, // 11 in
      orientation: 'portrait'
    },
    legal: {
      name: 'Legal (21,59 x 35,56 cm)',
      width: 12240,  // 8.5 in
      height: 20160, // 14 in
      orientation: 'portrait'
    }
  },

  // Default A4
  page: {
    width: 11906,
    height: 16838,
    orientation: 'portrait'
  },

  // Margin dinas resmi: Atas 2.0cm, Kiri 2.5cm, Kanan 2.0cm, Bawah 2.0cm (dalam twips: 1cm ≈ 567 twips)
  margins: {
    top: 1134,    // 2.0 cm
    bottom: 1134, // 2.0 cm
    left: 1418,   // 2.5 cm
    right: 1134,  // 2.0 cm
    header: 567,  // 1.0 cm
    footer: 567
  },

  // Tipografi
  fonts: {
    primary: 'Times New Roman',
    fallback: 'Arial',
    ascii: 'Times New Roman',
    hAnsi: 'Times New Roman'
  },

  typography: {
    bodySize: 24,       // 12pt (dalam half-points)
    kopGovSize: 24,     // 12pt (Pemerintah Provinsi / Dinas)
    kopSchoolSize: 28,  // 14pt (Nama Sekolah Tebal)
    kopContactSize: 18, // 9pt (Alamat, telp, email)
    titleSize: 28,      // 14pt (Judul Surat)
    lineSpacing: 276,   // 1.15 baris (240 * 1.15 = 276)
    lineSpacing15: 360, // 1.5 baris
    paragraphSpacingAfter: 72, // ~3.6pt
    alignment: 'both'
  },

  headings: {
    heading1: { size: 28, spacingBefore: 240, spacingAfter: 240, alignment: 'center', allCaps: true },
    heading2: { size: 24, spacingBefore: 180, spacingAfter: 120, alignment: 'left' },
    heading3: { size: 24, spacingBefore: 120, spacingAfter: 60, alignment: 'left' }
  },

  caption: {
    size: 20,
    spacingBefore: 60,
    spacingAfter: 120,
    alignment: 'center'
  }
};
