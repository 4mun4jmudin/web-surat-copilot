import test from 'node:test';
import assert from 'node:assert/strict';

import {
  normalizeListStyle,
  getListStyleSpec,
  parseListNodeMeta,
  toRomanNumeral,
  formatPageNumberLabel,
  normalizePageNumbering,
  normalizePageSetupOptions,
  extractPageSetupFromAiPayload
} from '../js/letters/wordFormat.js';
import { parseHtmlNodes } from '../js/utils/htmlAst.js';
import {
  shouldAllowLayoutRestructure,
  buildSurgicalIndexFromSmartPages,
  applySurgicalPatchesToSmartPages,
  buildSurgicalIndexFromDocModel,
  applySurgicalPatchesToDocModel
} from '../js/ai/surgicalEdit.js';
import {
  parseSmartPageHtmlToBlocks,
  docModelToSmartPages,
  buildDocModelFromSmartPagesData,
  aiJsonToSmartDocModel
} from '../js/ui/smartWordCanvas.js';
import { buildSuratTugasSppdDocModel } from '../js/letters/templates/suratTugasSppd.js';
import { DEFAULT_SCHOOL_PROFILE } from '../js/letters/schoolProfile.js';
import { renderDocx } from '../js/docx/render.js';
import { readZipPackage } from '../js/utils/zip.js';

test('wordFormat: normalisasi numbering bertingkat, romawi, pageNumbering & pageSetup AI', () => {
  assert.equal(normalizeListStyle('a'), 'lower-alpha');
  assert.equal(normalizeListStyle('A'), 'upper-alpha');
  assert.equal(normalizeListStyle('i'), 'lower-roman');
  assert.equal(normalizeListStyle('I'), 'upper-roman');
  assert.equal(normalizeListStyle('bullet'), 'bullet');
  assert.equal(normalizeListStyle('', true), 'decimal');

  const specRoman = getListStyleSpec('lower-roman', true);
  assert.equal(specRoman.ooxmlFmt, 'lowerRoman');
  assert.equal(specRoman.ooxmlNumId, '6');

  const parsedUl = parseListNodeMeta(parseHtmlNodes('<ol type="A" start="3" data-list-style="upper-alpha"><li>Satu</li></ol>')[0]);
  assert.deepEqual(parsedUl, { ordered: true, listStyle: 'upper-alpha', startAt: 3 });

  assert.equal(toRomanNumeral(4, false), 'iv');
  assert.equal(toRomanNumeral(9, true), 'IX');
  assert.equal(formatPageNumberLabel(3, 'lowerRoman'), 'iii');
  assert.equal(formatPageNumberLabel(2, 'decimal'), '2');

  const pn = normalizePageNumbering({ enabled: true, position: 'top-right', format: 'lower-roman', startAt: 5 });
  assert.deepEqual(pn, { enabled: true, position: 'top-right', format: 'lowerRoman', startAt: 5 });

  const extracted = extractPageSetupFromAiPayload({
    patches: [
      { op: 'setPageFormat', paperSize: 'f4', fontFamily: 'Calibri', fontSize: 11, margins: { left: 3.0 }, pageNumbering: { enabled: true, position: 'bottom-right' } }
    ]
  }, { paperSize: 'a4', orientation: 'portrait', margins: { top: 2, bottom: 2, left: 2.5, right: 2 } });

  assert.ok(extracted);
  assert.equal(extracted.partial.paperSize, 'f4');
  assert.equal(extracted.partial.fontFamily, 'Calibri');
  assert.equal(extracted.partial.fontSize, '11');
  assert.equal(extracted.full.margins.left, 3.0);
  assert.equal(extracted.full.margins.top, 2.0);
  assert.equal(extracted.full.pageNumbering.position, 'bottom-right');
});

test('Surgical Format Ops (Word Pro & DocModel): setPageFormat, setParagraphFormat, setList, convertParagraphsToList, & tabel SPPD tetap utuh', () => {
  const sppdModel = buildSuratTugasSppdDocModel({
    nomorSurat: '800/010/ST.SMKITA/XI/2026',
    assignees: [{ nama: 'Rina Marlina, S.Pd.', jabatan: 'Guru' }],
    tujuanLokasi: 'SMKN 1 Garut',
    maksudTugas: 'Koordinasi Ujian'
  }, DEFAULT_SCHOOL_PROFILE, { paperSize: 'f4' });

  // Pastikan instruksi format tidak memicu restrukturisasi paksa yang merusak tabel SPPD
  assert.equal(shouldAllowLayoutRestructure('Ubah font jadi Arial 11pt, margin kiri 3cm, spasi 1.5, dan nomor halaman di kanan bawah'), false);

  const converted = docModelToSmartPages(sppdModel, DEFAULT_SCHOOL_PROFILE);
  const index = buildSurgicalIndexFromSmartPages(converted.pages, DEFAULT_SCHOOL_PROFILE, {
    letterOptions: { paperSize: 'f4', fontFamily: 'Times New Roman' }
  });

  const firstParaId = index.orderedBlocks.find((b) => b.kind === 'paragraph')?.blockId;
  const tableBlock = index.snapshot.pages.flatMap((p) => p.blocks).find((b) => b.type === 'table');
  assert.ok(firstParaId);
  assert.ok(tableBlock);

  const patchRes = applySurgicalPatchesToSmartPages(converted.pages, {
    summaryOfChanges: 'Format font Arial 11pt, indent paragraf, & border sppd',
    patches: [
      { op: 'setPageFormat', fontFamily: 'Arial', fontSize: 11, lineSpacing: 1.5, margins: { left: 3.0 }, pageNumbering: { enabled: true, position: 'bottom-center', format: 'decimal', startAt: 1 } },
      { op: 'setParagraphFormat', blockId: firstParaId, firstLineIndentCm: 1.25, spacingAfterPt: 8, alignment: 'both' },
      { op: 'setTableFormat', blockId: tableBlock.blockId, headerShading: false }
    ]
  }, DEFAULT_SCHOOL_PROFILE, {
    allowLayoutRestructure: false,
    letterOptions: { paperSize: 'f4' }
  });

  assert.equal(patchRes.mode, 'surgical');
  assert.equal(patchRes.appliedCount, 3);
  assert.equal(patchRes.pageSetupUpdates.fontFamily, 'Arial');
  assert.equal(patchRes.pageSetupUpdates.lineSpacing, '1.5');
  assert.equal(patchRes.pageSetupUpdates.margins.left, 3.0);
  assert.equal(patchRes.pageSetupUpdates.pageNumbering.enabled, true);

  // Verifikasi tabel SPPD tidak rusak & atribut data-col-widths / rowspan / no-header-bg tetap terjaga
  assert.ok(patchRes.pages.some((p) => p.html.includes('text-indent: 1.25cm')));
  assert.ok(patchRes.pages.some((p) => p.html.includes('no-header-bg')));
  assert.ok(patchRes.pages.some((p) => p.html.includes('data-col-widths=')));

  // Uji DocModel Surgical ops: convertParagraphsToList & setList tanpa menghapus items
  const sampleDoc = {
    profile: 'suratDinas',
    paperSize: 'a4',
    blocks: [
      { type: 'letterTitle', title: 'SURAT PEMBERITAHUAN', nomor: 'Nomor: 001/SMKITA/2026' },
      { type: 'paragraph', text: 'Menyiapkan laptop dan charger.', runs: [{ text: 'Menyiapkan laptop dan charger.' }] },
      { type: 'paragraph', text: 'Membawa kartu peserta ujian.', runs: [{ text: 'Membawa kartu peserta ujian.' }] },
      { type: 'list', ordered: true, listStyle: 'decimal', items: ['Item satu', 'Item dua'] }
    ]
  };
  const dmIdx = buildSurgicalIndexFromDocModel(sampleDoc);
  const dmPatch = applySurgicalPatchesToDocModel(sampleDoc, {
    patches: [
      { op: 'setPageFormat', fontFamily: 'Calibri', pageNumbering: { enabled: true, position: 'top-right', format: 'lowerRoman', startAt: 2 } },
      { op: 'setList', blockId: 'p0_list0', listStyle: 'lower-alpha', startAt: 3 },
      { op: 'convertParagraphsToList', blockIds: ['p0_b0', 'p0_b1'], listStyle: 'upper-roman' }
    ]
  });

  assert.equal(dmPatch.appliedCount, 3);
  assert.equal(dmPatch.docModel.fontFamily, 'Calibri');
  assert.equal(dmPatch.docModel.pageNumbering.position, 'top-right');
  const convertedList = dmPatch.docModel.blocks.find((b) => b.type === 'list' && b.listStyle === 'upper-roman');
  assert.ok(convertedList);
  assert.equal(convertedList.items.length, 2);
  const styledList = dmPatch.docModel.blocks.find((b) => b.type === 'list' && b.listStyle === 'lower-alpha');
  assert.equal(styledList.startAt, 3);
  assert.deepEqual(styledList.items, ['Item satu', 'Item dua']);
});

test('renderDocx & Word Pro HTML round-trip menerapkan seluruh format Word (.docx OOXML)', async () => {
  const html = `
    <p style="text-align: justify; text-indent: 1.25cm; padding-left: 0.5cm; line-height: 1.5; margin-bottom: 10pt;">Paragraf dengan inden dan spasi penuh.</p>
    <ol data-list-style="lower-alpha" type="a" start="4" style="list-style-type: lower-alpha;">
      <li>Poin Keempat</li>
      <li>Poin Kelima</li>
    </ol>
  `;
  const parsedBlocks = parseSmartPageHtmlToBlocks(html, DEFAULT_SCHOOL_PROFILE);
  assert.equal(parsedBlocks[0].firstLineIndentCm, 1.25);
  assert.equal(parsedBlocks[0].leftIndentCm, 0.5);
  assert.equal(parsedBlocks[0].lineSpacing, 1.5);
  assert.equal(parsedBlocks[1].type, 'list');
  assert.equal(parsedBlocks[1].listStyle, 'lower-alpha');
  assert.equal(parsedBlocks[1].startAt, 4);

  const docModel = buildDocModelFromSmartPagesData({
    pages: [{ id: 'p1', html, isManualBreak: false }],
    schoolProfile: DEFAULT_SCHOOL_PROFILE,
    letterOptions: {
      paperSize: 'legal',
      orientation: 'landscape',
      margins: { top: 2.5, bottom: 2.5, left: 3.0, right: 2.5 },
      fontFamily: 'Cambria',
      fontSize: '11',
      lineSpacing: '1.5',
      paragraphSpacing: '10',
      showKop: false,
      pageNumbering: {
        enabled: true,
        position: 'top-right',
        format: 'upperRoman',
        startAt: 2
      }
    },
    meta: { title: 'DOKUMEN FORMAT PENUH', nomorSurat: '800/020/SMKITA/XI/2026' }
  });

  // Uji round-trip ke Smart Pages
  const backToPages = docModelToSmartPages(docModel, DEFAULT_SCHOOL_PROFILE);
  assert.ok(backToPages.pages[0].html.includes('text-indent: 1.25cm'));
  assert.ok(backToPages.pages[0].html.includes('data-list-style="lower-alpha"'));
  assert.ok(backToPages.pages[0].html.includes('start="4"'));

  // Uji ekspor OOXML .docx nyata
  const zipBytes = await renderDocx(docModel, 'uint8array');
  const pkg = await readZipPackage(zipBytes);

  const stylesXml = await pkg.file('word/styles.xml').async('string');
  const numberingXml = await pkg.file('word/numbering.xml').async('string');
  const documentXml = await pkg.file('word/document.xml').async('string');
  const relsXml = await pkg.file('word/_rels/document.xml.rels').async('string');
  const contentTypesXml = await pkg.file('[Content_Types].xml').async('string');

  assert.ok(stylesXml.includes('w:ascii="Cambria"'), 'Global fontFamily Cambria harus masuk styles.xml');
  assert.ok(stylesXml.includes('<w:sz w:val="22"/>'), 'Global fontSize 11pt (22 half-pt) harus masuk styles.xml');
  assert.ok(stylesXml.includes('w:line="360"'), 'Global lineSpacing 1.5 (360) harus masuk styles.xml');
  assert.ok(stylesXml.includes('w:after="200"'), 'Global paragraphSpacing 10pt (200 twips) harus masuk styles.xml');

  assert.ok(numberingXml.includes('lowerLetter'), 'lowerLetter numbering definition wajib ada');
  assert.ok(numberingXml.includes('<w:startOverride w:val="4"/>'), 'Custom list startAt=4 wajib menghasilkan startOverride');
  assert.ok(documentXml.includes('w:numId w:val="20"'), 'List block harus merujuk custom numId override startAt=4');

  assert.ok(documentXml.includes('w:orient="landscape"'), 'Orientasi landscape harus masuk document.xml');
  assert.ok(documentXml.includes('w:firstLine="709"'), 'firstLineIndent 1.25cm (709 twips) harus masuk document.xml');
  assert.ok(documentXml.includes('w:left="284"'), 'leftIndent 0.5cm (284 twips) harus masuk document.xml');
  assert.ok(documentXml.includes('<w:headerReference w:type="default" r:id="rIdHeader1"/>'), 'Top-right page number header wajib terpasang');
  assert.ok(documentXml.includes('<w:pgNumType w:fmt="upperRoman" w:start="2"/>'), 'Format upperRoman & start=2 wajib masuk pgNumType');
  assert.ok(relsXml.includes('Target="header1.xml"'));
  assert.ok(contentTypesXml.includes('/word/header1.xml'));
  assert.ok(pkg.file('word/header1.xml'));

  // Uji aiJsonToSmartDocModel membawa pageSetup lengkap
  const aiModel = aiJsonToSmartDocModel({
    letterTitle: 'SURAT KEPUTUSAN',
    pageSetup: {
      paperSize: 'f4',
      fontFamily: 'Georgia',
      fontSize: 11,
      pageNumbering: { enabled: true, position: 'bottom-center', format: 'decimal', startAt: 1 }
    },
    blocks: [
      { type: 'paragraph', text: 'Menimbang dst.', firstLineIndentCm: 1.25 },
      { type: 'list', ordered: true, listStyle: 'upper-alpha', items: ['Kesatu', 'Kedua'] }
    ]
  }, DEFAULT_SCHOOL_PROFILE);
  assert.equal(aiModel.fontFamily, 'Georgia');
  assert.equal(aiModel.fontSize, '11');
  assert.equal(aiModel.paperSize, 'f4');
  assert.equal(aiModel.pageNumbering.enabled, true);
  assert.equal(aiModel.blocks.find((b) => b.type === 'list')?.listStyle, 'upper-alpha');
});
