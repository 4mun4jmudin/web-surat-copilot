/**
 * jurnal.js - Profil Artikel / Jurnal Ilmiah (IEEE / Standard Template)
 */
export const jurnalProfile = {
  id: 'jurnal',
  name: 'Jurnal Ilmiah (Standard)',
  description: 'Format jurnal ilmiah (A4, margin 2.5 cm, font Arial/Times 10pt, spasi 1.0, single/two column format)',

  page: {
    width: 11906,
    height: 16838,
    orientation: 'portrait'
  },

  margins: {
    top: 1418,      // 2.5 cm
    bottom: 1418,   // 2.5 cm
    left: 1418,     // 2.5 cm
    right: 1418,    // 2.5 cm
    header: 720,
    footer: 720
  },

  fonts: {
    primary: 'Arial',
    fallback: 'Times New Roman'
  },

  typography: {
    bodySize: 20,       // 10pt
    lineSpacing: 240,   // Single space (1.0)
    paragraphIndent: 360, // ~0.63 cm
    paragraphSpacingAfter: 72, // ~3.6pt
    alignment: 'both'
  },

  headings: {
    heading1: {
      name: 'Heading 1',
      size: 22,         // 11pt
      bold: true,
      allCaps: false,
      alignment: 'left',
      spacingBefore: 180,
      spacingAfter: 90,
      keepNext: true
    },
    heading2: {
      name: 'Heading 2',
      size: 20,         // 10pt
      bold: true,
      alignment: 'left',
      spacingBefore: 120,
      spacingAfter: 60,
      keepNext: true
    },
    heading3: {
      name: 'Heading 3',
      size: 20,         // 10pt
      bold: false,
      italic: true,
      alignment: 'left',
      spacingBefore: 90,
      spacingAfter: 45,
      keepNext: true
    }
  },

  caption: {
    size: 18,           // 9pt
    italic: true,
    alignment: 'center',
    spacingBefore: 60,
    spacingAfter: 90
  },

  table: {
    borderSize: 4,
    borderColor: '000000',
    headerBackground: 'F5F5F5',
    fontSize: 18,       // 9pt
    cellPadding: { top: 80, bottom: 80, left: 100, right: 100 }
  },

  numberingRules: {
    cover: { numberFormat: 'none' },
    frontMatter: { numberFormat: 'decimal', start: 1, position: 'bottom-center' },
    body: { numberFormat: 'decimal', start: 1, position: 'bottom-center' }
  }
};

