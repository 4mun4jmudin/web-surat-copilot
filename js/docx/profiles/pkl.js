/**
 * pkl.js - Profil Laporan Praktik Kerja Lapangan (PKL) SMK
 * Mengikuti standar penulisan laporan akademik kejuruan di Indonesia.
 */

export const pklProfile = {
  id: 'pkl',
  name: 'Laporan PKL SMK',
  description: 'Format baku laporan Praktik Kerja Lapangan SMK (A4, margin 4-4-3-3 cm, Times New Roman 12pt, spasi 1.5)',
  
  // Ukuran kertas (A4 dalam twips: 1cm ≈ 567 twips)
  page: {
    width: 11906,   // 21.0 cm
    height: 16838,  // 29.7 cm
    orientation: 'portrait'
  },

  // Margin halaman dalam twips (Left 4cm, Top 4cm, Right 3cm, Bottom 3cm)
  margins: {
    top: 2268,      // 4.0 cm
    bottom: 1701,   // 3.0 cm
    left: 2268,     // 4.0 cm
    right: 1701,    // 3.0 cm
    header: 720,    // 1.27 cm
    footer: 720
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
    lineSpacing: 360,   // Spasi 1.5 baris (360/240 = 1.5)
    paragraphIndent: 709, // Indent baris pertama 1.25 cm
    paragraphSpacingAfter: 120, // 6pt
    alignment: 'both'   // Rata kiri-kanan (justify)
  },

  headings: {
    heading1: {
      name: 'Heading 1',
      size: 28,         // 14pt
      bold: true,
      allCaps: true,
      alignment: 'center',
      spacingBefore: 240, // 12pt
      spacingAfter: 240,
      keepNext: true,
      prefix: 'BAB '
    },
    heading2: {
      name: 'Heading 2',
      size: 24,         // 12pt
      bold: true,
      alignment: 'left',
      spacingBefore: 180, // 9pt
      spacingAfter: 120,  // 6pt
      keepNext: true
    },
    heading3: {
      name: 'Heading 3',
      size: 24,         // 12pt
      bold: true,
      italic: false,
      alignment: 'left',
      spacingBefore: 120, // 6pt
      spacingAfter: 60,
      keepNext: true
    }
  },

  caption: {
    size: 20,           // 10pt
    italic: true,
    alignment: 'center',
    spacingBefore: 60,
    spacingAfter: 120
  },

  table: {
    borderSize: 4,      // 1/2 pt
    borderColor: '000000',
    headerBackground: 'F2F2F2',
    fontSize: 22,       // 11pt
    cellPadding: {
      top: 100,
      bottom: 100,
      left: 150,
      right: 150
    }
  },

  numberingRules: {
    cover: { numberFormat: 'none' },
    frontMatter: { numberFormat: 'roman', start: 1, position: 'bottom-center' },
    body: { numberFormat: 'decimal', start: 1, position: 'bottom-right' }
  }
};

