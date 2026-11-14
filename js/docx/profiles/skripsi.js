/**
 * skripsi.js - Profil Skripsi / Tugas Akhir Perguruan Tinggi
 */
export const skripsiProfile = {
  id: 'skripsi',
  name: 'Skripsi / Tugas Akhir',
  description: 'Format baku skripsi universitas (A4, margin 4-4-3-3 cm, Times New Roman 12pt, spasi 2.0/double, bab berangka romawi)',

  page: {
    width: 11906,
    height: 16838,
    orientation: 'portrait'
  },

  margins: {
    top: 2268,      // 4 cm
    bottom: 1701,   // 3 cm
    left: 2268,     // 4 cm
    right: 1701,    // 3 cm
    header: 720,
    footer: 720
  },

  fonts: {
    primary: 'Times New Roman',
    fallback: 'Calibri'
  },

  typography: {
    bodySize: 24,       // 12pt
    lineSpacing: 480,   // Double space (2.0)
    paragraphIndent: 709, // 1.25 cm
    paragraphSpacingAfter: 0,
    alignment: 'both'
  },

  headings: {
    heading1: {
      name: 'Heading 1',
      size: 28,         // 14pt
      bold: true,
      allCaps: true,
      alignment: 'center',
      spacingBefore: 240,
      spacingAfter: 240,
      keepNext: true,
      prefix: 'BAB '
    },
    heading2: {
      name: 'Heading 2',
      size: 24,
      bold: true,
      alignment: 'left',
      spacingBefore: 180,
      spacingAfter: 120,
      keepNext: true
    },
    heading3: {
      name: 'Heading 3',
      size: 24,
      bold: true,
      alignment: 'left',
      spacingBefore: 120,
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
    borderSize: 4,
    borderColor: '000000',
    headerBackground: 'F0F0F0',
    fontSize: 22,
    cellPadding: { top: 100, bottom: 100, left: 150, right: 150 }
  },

  numberingRules: {
    cover: { numberFormat: 'none' },
    frontMatter: { numberFormat: 'roman', start: 1, position: 'bottom-center' },
    body: { numberFormat: 'decimal', start: 1, position: 'bottom-right' }
  }
};

