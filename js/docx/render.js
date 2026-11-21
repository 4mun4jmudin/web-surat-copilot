/**
 * render.js - Renderer OOXML Microsoft Word (.docx)
 * Mendukung Kop Surat Gambar Banner, Pengaturan Margin Kustom (A4/F4), dan Komponen Surat Resmi.
 */

import { escapeXml, buildTextElement } from '../utils/xml.js';
import { createZipPackage } from '../utils/zip.js';
import { pklProfile } from './profiles/pkl.js';
import { skripsiProfile } from './profiles/skripsi.js';
import { jurnalProfile } from './profiles/jurnal.js';
import { suratDinasProfile } from './profiles/suratDinas.js';
import { buildStylesXml } from './styles.js';
import { buildNumberingXml } from './numbering.js';
import {
  buildKopSuratXml,
  buildKopBannerDrawingXml,
  buildLetterTitleXml,
  buildLetterMetaRecipientXml,
  buildIdentityTableXml,
  buildLetterSignatureXml,
  buildQrDrawingXml,
  parseBase64Image
} from '../letters/builder.js';
import {
  buildPageField,
  buildSeqField,
  buildRefField,
  buildTocBlock,
  buildListOfFiguresBlock,
  buildListOfTablesBlock,
  buildBookmarkStart,
  buildBookmarkEnd
} from './fields.js';

const suratDinasF4Profile = {
  ...suratDinasProfile,
  id: 'suratDinasF4',
  name: 'Surat Dinas Sekolah (F4 / Folio)',
  page: suratDinasProfile.pageSizes.f4
};

const PROFILES = {
  suratdinas: suratDinasProfile,
  suratdinasf4: suratDinasF4Profile,
  pkl: pklProfile,
  skripsi: skripsiProfile,
  jurnal: jurnalProfile
};

export function getProfile(profileName) {
  const key = (profileName || 'suratDinas').toLowerCase();
  return PROFILES[key] || PROFILES.suratdinas;
}

function renderRuns(runs = []) {
  return runs.map((run, idx) => {
    const rPrParts = [];
    if (run.bold) rPrParts.push('<w:b/>');
    if (run.italic) rPrParts.push('<w:i/>');
    if (run.underline) rPrParts.push('<w:u w:val="single"/>');
    if (run.letterSpacing) rPrParts.push(`<w:spacing w:val="${Math.round(Number(run.letterSpacing) * 20)}"/>`);
    if (run.fontSizeHalfPt) {
      rPrParts.push(`<w:sz w:val="${Math.round(Number(run.fontSizeHalfPt))}"/>`);
      rPrParts.push(`<w:szCs w:val="${Math.round(Number(run.fontSizeHalfPt))}"/>`);
    }
    if (run.sup) rPrParts.push('<w:vertAlign w:val="superscript"/>');
    if (run.sub) rPrParts.push('<w:vertAlign w:val="subscript"/>');
    if (run.color) rPrParts.push(`<w:color w:val="${escapeXml(String(run.color).replace(/^#/, ''))}"/>`);

    const rPrXml = rPrParts.length > 0 ? `<w:rPr>${rPrParts.join('')}</w:rPr>` : '';
    const rawText = String(run.text || '');
    const tabParts = rawText.split('\t');
    const textXml = tabParts.length > 1
      ? tabParts.map(seg => (seg ? buildTextElement(seg) : '')).join('<w:tab/>')
      : buildTextElement(rawText);

    let runXml = `<w:r>${rPrXml}${textXml}</w:r>`;

    if (run.bookmark) {
      const bId = String(idx + 100);
      runXml = `${buildBookmarkStart(bId, run.bookmark)}${runXml}${buildBookmarkEnd(bId)}`;
    }

    return runXml;
  }).join('');
}

function buildTableBordersXml(borderStyle = 'grid') {
  const b = (val = 'single', sz = '4') => val === 'none'
    ? 'w:val="none"'
    : `w:val="${val}" w:sz="${sz}" w:space="0" w:color="000000"`;

  if (borderStyle === 'none') {
    return `
      <w:tblBorders>
        <w:top ${b('none')}/><w:bottom ${b('none')}/>
        <w:left ${b('none')}/><w:right ${b('none')}/>
        <w:insideH ${b('none')}/><w:insideV ${b('none')}/>
      </w:tblBorders>
    `.trim();
  }
  if (borderStyle === 'sppd') {
    return `
      <w:tblBorders>
        <w:top ${b('single', '6')}/><w:bottom ${b('single', '6')}/>
        <w:left ${b('none')}/><w:right ${b('none')}/>
        <w:insideH ${b('single', '4')}/><w:insideV ${b('single', '4')}/>
      </w:tblBorders>
    `.trim();
  }
  return `
    <w:tblBorders>
      <w:top ${b('single', '4')}/><w:bottom ${b('single', '4')}/>
      <w:left ${b('single', '4')}/><w:right ${b('single', '4')}/>
      <w:insideH ${b('single', '4')}/><w:insideV ${b('single', '4')}/>
    </w:tblBorders>
  `.trim();
}

function normalizeTableRowEntry(rowEntry) {
  if (Array.isArray(rowEntry)) {
    return { cells: rowEntry, isHeader: null, minHeightTwips: 0, cantSplit: true };
  }
  if (rowEntry && typeof rowEntry === 'object' && Array.isArray(rowEntry.cells)) {
    const minH = rowEntry.minHeightTwips
      ? Math.round(Number(rowEntry.minHeightTwips))
      : (rowEntry.minHeightCm ? Math.round(Number(rowEntry.minHeightCm) * 567) : 0);
    return {
      cells: rowEntry.cells,
      isHeader: rowEntry.isHeader ?? null,
      minHeightTwips: minH,
      cantSplit: rowEntry.cantSplit !== false
    };
  }
  return { cells: [], isHeader: false, minHeightTwips: 0, cantSplit: true };
}

function normalizeTableCellEntry(cell, fallbackAlign = 'left', fallbackBold = false) {
  if (cell && typeof cell === 'object' && !Array.isArray(cell)) {
    const colspan = Math.max(1, Number(cell.colspan) || 1);
    const rowspan = Math.max(1, Number(cell.rowspan) || 1);
    const align = cell.align || fallbackAlign;
    const bold = cell.bold !== undefined ? Boolean(cell.bold) : fallbackBold;
    const italic = Boolean(cell.italic);
    const underline = Boolean(cell.underline);
    let paragraphs = null;
    if (Array.isArray(cell.paragraphs) && cell.paragraphs.length > 0) {
      paragraphs = cell.paragraphs.map(p => (typeof p === 'string' ? { text: p, align, bold, italic, underline } : { align, bold, italic, underline, ...p }));
    } else if (Array.isArray(cell.lines) && cell.lines.length > 0) {
      paragraphs = cell.lines.map(l => (typeof l === 'string' ? { text: l, align, bold, italic, underline } : { align, bold, italic, underline, ...l }));
    } else {
      const rawText = cell.text !== undefined ? String(cell.text) : '';
      const splitLines = rawText.split(/\r?\n/);
      paragraphs = splitLines.map(lineText => ({ text: lineText, align, bold, italic, underline }));
    }
    return {
      paragraphs,
      colspan,
      rowspan,
      vAlign: cell.vAlign || '',
      shade: cell.shade || '',
      qrSlot: Boolean(cell.qrSlot),
      qrCaption: cell.qrCaption || ''
    };
  }
  const splitLines = String(cell ?? '').split(/\r?\n/);
  return {
    paragraphs: splitLines.map(lineText => ({ text: lineText, align: fallbackAlign, bold: fallbackBold })),
    colspan: 1,
    rowspan: 1,
    vAlign: '',
    shade: '',
    qrSlot: false,
    qrCaption: ''
  };
}

function renderRichTableCellParagraphs(cellObj, { hasQrImage = false, borderStyle = 'grid' } = {}) {
  const marTop = borderStyle === 'none' ? 30 : 75;
  const marSide = borderStyle === 'none' ? 60 : 110;
  const pXml = (cellObj.paragraphs || []).map((p) => {
    if (p.qrSlot && hasQrImage) {
      return `
        ${buildQrDrawingXml('rIdQr', 2.0, 2.0)}
        ${p.qrCaption ? `<w:p><w:pPr><w:spacing w:before="0" w:after="20" w:line="200" w:lineRule="auto"/><w:jc w:val="${p.align || 'center'}"/></w:pPr><w:r><w:rPr><w:i/><w:sz w:val="15"/><w:color w:val="166534"/></w:rPr><w:t>${escapeXml(p.qrCaption)}</w:t></w:r></w:p>` : ''}
      `.trim();
    }
    const jc = p.align === 'both' || p.align === 'justify' ? 'both' : (p.align || 'left');
    const runs = Array.isArray(p.runs) && p.runs.length > 0
      ? p.runs.map(r => ({
          ...r,
          bold: r.bold !== undefined ? r.bold : p.bold,
          italic: r.italic !== undefined ? r.italic : p.italic,
          underline: r.underline !== undefined ? r.underline : p.underline,
          fontSizeHalfPt: r.fontSizeHalfPt || p.fontSizeHalfPt
        }))
      : [{
          text: String(p.text ?? ''),
          bold: Boolean(p.bold),
          italic: Boolean(p.italic),
          underline: Boolean(p.underline),
          fontSizeHalfPt: p.fontSizeHalfPt
        }];
    const after = p.spacingAfter !== undefined ? p.spacingAfter : 20;
    const before = p.spacingBefore !== undefined ? p.spacingBefore : 0;
    return `
      <w:p>
        <w:pPr>
          <w:pStyle w:val="Normal"/>
          <w:spacing w:before="${before}" w:after="${after}" w:line="240" w:lineRule="auto"/>
          <w:jc w:val="${jc}"/>
        </w:pPr>
        ${renderRuns(runs)}
      </w:p>
    `.trim();
  }).join('');

  const qrInCellFallback = cellObj.qrSlot && hasQrImage && !(cellObj.paragraphs || []).some(p => p.qrSlot)
    ? buildQrDrawingXml('rIdQr', 2.0, 2.0)
    : '';

  return {
    contentXml: (pXml + qrInCellFallback) || '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>',
    marTop,
    marSide
  };
}

export function renderRichTableXml(block, profile, { hasQrImage = false } = {}) {
  const rawRows = (block.rows || []).map(normalizeTableRowEntry);
  if (rawRows.length === 0) return '';

  const borderStyle = block.borderStyle || 'grid';
  const useHeaderShading = block.headerShading !== undefined ? Boolean(block.headerShading) : (borderStyle === 'grid' && block.headerRow !== false);
  const baseTwips = Math.max(4000, (profile.page?.width || 11906) - (profile.margins?.left || 1418) - (profile.margins?.right || 1134));
  const widthPctFactor = (block.tableWidthPct && Number(block.tableWidthPct) >= 30 && Number(block.tableWidthPct) < 100)
    ? (Number(block.tableWidthPct) / 100)
    : 1;
  const printableTwips = Math.max(3000, Math.round(baseTwips * widthPctFactor));

  // Hitung jumlah kolom maksimal dengan memperhitungkan colspan & rowspan
  let numCols = Array.isArray(block.colWidthsPct) && block.colWidthsPct.length > 0 ? block.colWidthsPct.length : 0;
  {
    const simTracker = [];
    for (const r of rawRows) {
      let c = 0;
      for (const cell of r.cells) {
        while ((simTracker[c] || 0) > 0) c++;
        const cs = Math.max(1, Number(cell?.colspan) || 1);
        const rs = Math.max(1, Number(cell?.rowspan) || 1);
        if (rs > 1) {
          for (let k = 0; k < cs; k++) simTracker[c + k] = Math.max(simTracker[c + k] || 0, rs - 1);
        }
        c += cs;
        numCols = Math.max(numCols, c);
      }
      for (let k = 0; k < simTracker.length; k++) {
        if (simTracker[k] > 0) simTracker[k]--;
      }
    }
  }
  numCols = Math.max(1, numCols);

  const rawPcts = Array.isArray(block.colWidthsPct) && block.colWidthsPct.length === numCols
    ? block.colWidthsPct.map(n => Math.max(1, Number(n) || 1))
    : Array.from({ length: numCols }, () => 100 / numCols);
  const pctSum = rawPcts.reduce((a, b) => a + b, 0) || 100;
  const colWidthsTwips = rawPcts.map(p => Math.round((p / pctSum) * printableTwips));

  const tblGridXml = `<w:tblGrid>${colWidthsTwips.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`;

  const tracker = Array.from({ length: numCols }, () => 0);
  const rowsXml = rawRows.map((rowEntry, rIdx) => {
    const isHeader = rowEntry.isHeader !== null ? rowEntry.isHeader : (rIdx === 0 && block.headerRow !== false);
    const defaultAlign = isHeader ? 'center' : 'left';
    const cellsXmlParts = [];
    let colPtr = 0;
    let cellIdx = 0;

    while (colPtr < numCols) {
      if (tracker[colPtr] > 0) {
        const w = colWidthsTwips[colPtr] || Math.round(printableTwips / numCols);
        cellsXmlParts.push(`
          <w:tc>
            <w:tcPr>
              <w:tcW w:w="${w}" w:type="dxa"/>
              <w:vMerge/>
            </w:tcPr>
            <w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>
          </w:tc>
        `.trim());
        tracker[colPtr]--;
        colPtr++;
        continue;
      }

      if (cellIdx >= rowEntry.cells.length) {
        const w = colWidthsTwips[colPtr] || Math.round(printableTwips / numCols);
        cellsXmlParts.push(`
          <w:tc>
            <w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>
            <w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>
          </w:tc>
        `.trim());
        colPtr++;
        continue;
      }

      const rawCell = rowEntry.cells[cellIdx++];
      const cell = normalizeTableCellEntry(rawCell, defaultAlign, isHeader && (block.headerBold !== false));
      const spanCols = Math.min(cell.colspan, numCols - colPtr);
      let spanWidth = 0;
      for (let k = 0; k < spanCols; k++) {
        spanWidth += colWidthsTwips[colPtr + k] || Math.round(printableTwips / numCols);
        if (cell.rowspan > 1) {
          tracker[colPtr + k] = cell.rowspan - 1;
        }
      }

      const { contentXml, marTop, marSide } = renderRichTableCellParagraphs(cell, { hasQrImage, borderStyle });
      const shadeColor = cell.shade || (isHeader && useHeaderShading ? 'F2F2F2' : '');
      const tcPrParts = [
        `<w:tcW w:w="${spanWidth}" w:type="dxa"/>`,
        spanCols > 1 ? `<w:gridSpan w:val="${spanCols}"/>` : '',
        cell.rowspan > 1 ? '<w:vMerge w:val="restart"/>' : '',
        shadeColor ? `<w:shd w:val="clear" w:color="auto" w:fill="${escapeXml(shadeColor)}"/>` : '',
        `<w:tcMar><w:top w:w="${marTop}" w:type="dxa"/><w:bottom w:w="${marTop}" w:type="dxa"/><w:left w:w="${marSide}" w:type="dxa"/><w:right w:w="${marSide}" w:type="dxa"/></w:tcMar>`,
        cell.vAlign ? `<w:vAlign w:val="${cell.vAlign === 'middle' ? 'center' : cell.vAlign}"/>` : ''
      ].filter(Boolean).join('');

      cellsXmlParts.push(`<w:tc><w:tcPr>${tcPrParts}</w:tcPr>${contentXml}</w:tc>`);
      colPtr += spanCols;
    }

    const trPrParts = [
      rowEntry.cantSplit ? '<w:cantSplit/>' : '',
      isHeader ? '<w:tblHeader/>' : '',
      rowEntry.minHeightTwips > 0 ? `<w:trHeight w:val="${rowEntry.minHeightTwips}" w:hRule="atLeast"/>` : ''
    ].filter(Boolean).join('');

    return `<w:tr>${trPrParts ? `<w:trPr>${trPrParts}</w:trPr>` : ''}${cellsXmlParts.join('')}</w:tr>`;
  }).join('');

  return `
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="TableNormal"/>
        <w:tblW w:w="${printableTwips}" w:type="dxa"/>
        ${Array.isArray(block.colWidthsPct) && block.colWidthsPct.length > 0 ? '<w:tblLayout w:type="fixed"/>' : ''}
        <w:jc w:val="${widthPctFactor < 1 ? 'left' : 'center'}"/>
        ${buildTableBordersXml(borderStyle)}
      </w:tblPr>
      ${tblGridXml}
      ${rowsXml}
    </w:tbl>
    <w:p><w:pPr><w:spacing w:before="0" w:after="60" w:line="220" w:lineRule="auto"/></w:pPr></w:p>
  `.trim();
}

function buildSectPr(profile, numberFormat = 'none', restartNumbering = false, footerRelId = null) {
  const p = profile.page;
  const m = profile.margins;

  let footerXml = '';
  if (footerRelId) {
    footerXml = `<w:footerReference w:type="default" r:id="${footerRelId}"/>`;
  }

  let pgNumTypeXml = '';
  if (numberFormat === 'roman') {
    pgNumTypeXml = `<w:pgNumType w:fmt="lowerRoman"${restartNumbering ? ' w:start="1"' : ''}/>`;
  } else if (numberFormat === 'decimal') {
    pgNumTypeXml = `<w:pgNumType w:fmt="decimal"${restartNumbering ? ' w:start="1"' : ''}/>`;
  }

  return `
    <w:sectPr>
      ${footerXml}
      <w:pgSz w:w="${p.width}" w:h="${p.height}" w:orient="${p.orientation}"/>
      <w:pgMar w:top="${m.top}" w:right="${m.right}" w:bottom="${m.bottom}" w:left="${m.left}" w:header="${m.header}" w:footer="${m.footer}"/>
      ${pgNumTypeXml}
    </w:sectPr>
  `.trim();
}

export function renderDocumentBody(doc, profile, hasKopImage = false, hasQrImage = false) {
  const bodyXmlParts = [];
  let figureCounter = 0;
  let tableCounter = 0;
  let bookmarkCounter = 200;

  const isLetter = doc.profile && doc.profile.toLowerCase().startsWith('suratdinas');
  let currentSectionFormat = isLetter ? 'none' : 'none';

  for (let i = 0; i < doc.blocks.length; i++) {
    const block = doc.blocks[i];

    switch (block.type) {
      // 1. KOP SURAT (Dukungan Banner Gambar & Tabel Profil)
      case 'letterKop': {
        if (block.school) {
          bodyXmlParts.push(buildKopSuratXml(block.school, hasKopImage));
        } else if (hasKopImage) {
          bodyXmlParts.push(buildKopBannerDrawingXml('rIdKop', block.widthCm || 16.5, block.heightCm || 3.3));
        }
        break;
      }

      // 2. JUDUL DAN NOMOR SURAT
      case 'letterTitle': {
        bodyXmlParts.push(buildLetterTitleXml(block.title, block.nomor, block.metaRecipient));
        break;
      }

      // 2B. BLOK SIFAT, PERIHAL & KEPADA YTH
      case 'letterMetaRecipient': {
        bodyXmlParts.push(buildLetterMetaRecipientXml(block));
        break;
      }

      // 3. TABEL IDENTITAS SISWA / PEGAWAI
      case 'identityTable': {
        bodyXmlParts.push(buildIdentityTableXml(block.rows || [], block.colWidthsPct || null));
        break;
      }

      // 4. BLOK TANDA TANGAN RESMI
      case 'letterSignature': {
        bodyXmlParts.push(buildLetterSignatureXml(block, { hasQrImage }));
        break;
      }

      case 'heading': {
        const lvl = block.level || 1;
        const styleName = `Heading${lvl}`;
        const ilvl = lvl - 1;
        const bId = block.id || `head_${i}`;
        bookmarkCounter++;

        bodyXmlParts.push(`
          <w:p>
            <w:pPr>
              <w:pStyle w:val="${styleName}"/>
              <w:numPr>
                <w:ilvl w:val="${ilvl}"/>
                <w:numId w:val="1"/>
              </w:numPr>
            </w:pPr>
            ${buildBookmarkStart(bookmarkCounter, bId)}
            <w:r>
              <w:t>${escapeXml(block.text)}</w:t>
            </w:r>
            ${buildBookmarkEnd(bookmarkCounter)}
          </w:p>
        `);
        break;
      }

      case 'paragraph': {
        const runs = block.runs || [{ text: block.text || '' }];
        const isLetterBody = block.style === 'letterBody' || isLetter;
        const indentVal = isLetterBody ? 0 : (profile.typography?.paragraphIndent || 709);
        const lineSpacing = profile.typography?.lineSpacing || 276;
        const spacingAfter = block.spacingAfter !== undefined ? Number(block.spacingAfter) : (isLetterBody ? 100 : (profile.typography?.paragraphSpacingAfter || 120));
        const spacingBefore = block.spacingBefore !== undefined ? Number(block.spacingBefore) : 0;

        bodyXmlParts.push(`
          <w:p>
            <w:pPr>
              <w:pStyle w:val="Normal"/>
              ${indentVal ? `<w:ind w:firstLine="${indentVal}"/>` : ''}
              <w:spacing w:before="${spacingBefore}" w:line="${lineSpacing}" w:lineRule="auto" w:after="${spacingAfter}"/>
              <w:jc w:val="${block.alignment || profile.typography?.alignment || 'both'}"/>
            </w:pPr>
            ${renderRuns(runs)}
          </w:p>
        `);
        break;
      }

      case 'list': {
        const numId = block.ordered ? '2' : '3';
        (block.items || []).forEach(item => {
          let itemRuns = typeof item === 'string' ? [{ text: item }] : (Array.isArray(item) ? item : []);
          bodyXmlParts.push(`
            <w:p>
              <w:pPr>
                <w:pStyle w:val="Normal"/>
                <w:numPr>
                  <w:ilvl w:val="0"/>
                  <w:numId w:val="${numId}"/>
                </w:numPr>
                <w:ind w:left="720" w:hanging="360"/>
              </w:pPr>
              ${renderRuns(itemRuns)}
            </w:p>
          `);
        });
        break;
      }

      case 'table': {
        tableCounter++;
        const tblCaption = block.caption || `Tabel ${tableCounter}`;
        const bId = block.id || `table_${tableCounter}`;
        bookmarkCounter++;

        if (block.caption) {
          bodyXmlParts.push(`
            <w:p>
              <w:pPr>
                <w:pStyle w:val="Caption"/>
                <w:jc w:val="left"/>
              </w:pPr>
              ${buildBookmarkStart(bookmarkCounter, bId)}
              <w:r><w:t>Tabel </w:t></w:r>
              ${buildSeqField('Tabel', tableCounter)}
              <w:r><w:t>. ${escapeXml(tblCaption)}</w:t></w:r>
              ${buildBookmarkEnd(bookmarkCounter)}
            </w:p>
          `);
        }

        bodyXmlParts.push(renderRichTableXml(block, profile, { hasQrImage }));
        break;
      }

      case 'figure': {
        figureCounter++;
        const figCaption = block.caption || `Gambar ${figureCounter}`;
        const bId = block.id || `fig_${figureCounter}`;
        bookmarkCounter++;

        bodyXmlParts.push(`
          <w:p>
            <w:pPr>
              <w:pStyle w:val="Normal"/>
              <w:jc w:val="center"/>
              <w:spacing w:before="120" w:after="120"/>
            </w:pPr>
            <w:r>
              <w:rPr><w:b/><w:color w:val="666666"/></w:rPr>
              <w:t>[ Ilustrasi: ${escapeXml(figCaption)} ]</w:t>
            </w:r>
          </w:p>
          <w:p>
            <w:pPr>
              <w:pStyle w:val="Caption"/>
              <w:jc w:val="center"/>
            </w:pPr>
            ${buildBookmarkStart(bookmarkCounter, bId)}
            <w:r><w:t>Gambar </w:t></w:r>
            ${buildSeqField('Gambar', figureCounter)}
            <w:r><w:t>. ${escapeXml(figCaption)}</w:t></w:r>
            ${buildBookmarkEnd(bookmarkCounter)}
          </w:p>
        `);
        break;
      }

      case 'pageBreak': {
        bodyXmlParts.push(`
          <w:p>
            <w:r><w:br w:type="page"/></w:r>
          </w:p>
        `);
        break;
      }

      case 'sectionBreak': {
        const nextFmt = block.numberFormat || 'decimal';
        const restart = block.restart !== false;
        let footerId = currentSectionFormat === 'none' ? null : (currentSectionFormat === 'roman' ? 'rIdFooter1' : 'rIdFooter2');

        bodyXmlParts.push(`
          <w:p>
            <w:pPr>
              ${buildSectPr(profile, currentSectionFormat, restart, footerId)}
            </w:pPr>
          </w:p>
        `);
        currentSectionFormat = nextFmt;
        break;
      }

      case 'toc':
        bodyXmlParts.push(buildTocBlock(block.title || 'DAFTAR ISI'));
        break;
      case 'listOfFigures':
        bodyXmlParts.push(buildListOfFiguresBlock(block.title || 'DAFTAR GAMBAR'));
        break;
      case 'listOfTables':
        bodyXmlParts.push(buildListOfTablesBlock(block.title || 'DAFTAR TABEL'));
        break;

      case 'reference': {
        const bId = block.id || `ref_${i}`;
        bookmarkCounter++;
        bodyXmlParts.push(`
          <w:p>
            <w:pPr>
              <w:pStyle w:val="Normal"/>
              <w:ind w:left="720" w:hanging="720"/>
              <w:spacing w:line="240" w:lineRule="auto" w:after="120"/>
            </w:pPr>
            ${buildBookmarkStart(bookmarkCounter, bId)}
            <w:r><w:t>${escapeXml(block.text)}</w:t></w:r>
            ${buildBookmarkEnd(bookmarkCounter)}
          </w:p>
        `);
        break;
      }

      case 'citation': {
        bodyXmlParts.push(`<w:r><w:t>[${escapeXml(block.refId)}]</w:t></w:r>`);
        break;
      }

      case 'crossRef': {
        bodyXmlParts.push(`
          <w:r><w:t>${escapeXml(block.prefix || '')} </w:t></w:r>
          ${buildRefField(block.targetId, '1')}
        `);
        break;
      }

      default:
        break;
    }
  }

  const finalSectPr = buildSectPr(profile, currentSectionFormat, false, null);

  return `
    <w:body>
      ${bodyXmlParts.join('\n')}
      ${finalSectPr}
    </w:body>
  `.trim();
}

export function buildDocumentXml(doc, profile, hasKopImage = false, hasQrImage = false) {
  const bodyXml = renderDocumentBody(doc, profile, hasKopImage, hasQrImage);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  ${bodyXml}
</w:document>`.trim();
}

export function buildContentTypesXml(hasKopImage = false, hasQrImage = false) {
  const hasAnyImage = hasKopImage || hasQrImage;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${hasAnyImage ? '<Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="svg" ContentType="image/svg+xml"/>' : ''}
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>`.trim();
}

export function buildRootRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`.trim();
}

export function buildDocRelsXml(hasKopImage = false, hasQrImage = false) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rIdFooter2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>
  ${hasKopImage ? '<Relationship Id="rIdKop" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/kop.png"/><Relationship Id="rIdLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/>' : ''}
  ${hasQrImage ? '<Relationship Id="rIdQr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/qr.png"/>' : ''}
</Relationships>`.trim();
}

export function buildSettingsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:updateFields w:val="true"/>
  <w:defaultTabStop w:val="720"/>
</w:settings>`.trim();
}

export function buildFooter1Xml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>
    ${buildPageField()}
  </w:p>
</w:ftr>`.trim();
}

export function buildFooter2Xml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="right"/></w:pPr>
    ${buildPageField()}
  </w:p>
</w:ftr>`.trim();
}

/**
 * Fungsi Utama: Merender Document Model JSON menjadi paket berkas .docx
 */
export async function renderDocx(docModel, outputType = 'blob') {
  let profile = { ...getProfile(docModel.profile) };

  // 1. Terapkan ukuran kertas kustom (A4, F4, Letter, Legal)
  const paperKey = (docModel.paperSize || docModel.letterData?.options?.paperSize || 'a4').toLowerCase();
  if (profile.pageSizes && profile.pageSizes[paperKey]) {
    profile.page = { ...profile.pageSizes[paperKey] };
  }

  // 2. Terapkan orientasi kertas (portrait / landscape)
  const orientation = (docModel.orientation || docModel.letterData?.options?.orientation || 'portrait').toLowerCase();
  if (orientation === 'landscape') {
    const origW = profile.page.width;
    const origH = profile.page.height;
    profile.page.width = Math.max(origW, origH);
    profile.page.height = Math.min(origW, origH);
    profile.page.orientation = 'landscape';
  }

  // 3. Terapkan margin kustom
  const customM = docModel.customMargins || docModel.letterData?.options?.margins;
  if (customM) {
    const toTwips = (val) => (typeof val === 'number' && val <= 25 ? Math.round(val * 567) : val);
    profile.margins = {
      ...profile.margins,
      top: customM.top !== undefined ? toTwips(customM.top) : profile.margins.top,
      bottom: customM.bottom !== undefined ? toTwips(customM.bottom) : profile.margins.bottom,
      left: customM.left !== undefined ? toTwips(customM.left) : profile.margins.left,
      right: customM.right !== undefined ? toTwips(customM.right) : profile.margins.right
    };
  }

  // Cari apakah ada gambar banner kop surat atau logo sekolah
  let kopParsed = null;
  const kopBlock = (docModel.blocks || []).find(b => b.type === 'letterKop');
  const kopDataUrl = kopBlock?.image || docModel.letterData?.kopImage || kopBlock?.school?.logo || docModel.letterData?.schoolProfile?.logo;
  if (kopDataUrl) {
    kopParsed = parseBase64Image(kopDataUrl);
  }

  // Cari gambar Barcode QR pengesahan jika signatureMode === 'qr'
  let qrParsed = null;
  const sigBlock = (docModel.blocks || []).find(b => b.type === 'letterSignature' && b.signatureMode === 'qr' && b.qrImage);
  const qrDataUrl = sigBlock?.qrImage || docModel.qrSignatureOverlay?.qrImage || '';
  if (qrDataUrl) {
    qrParsed = parseBase64Image(qrDataUrl);
  }

  const hasKopImage = !!kopParsed;
  const hasQrImage = !!qrParsed;

  const files = {
    '[Content_Types].xml': buildContentTypesXml(hasKopImage, hasQrImage),
    '_rels/.rels': buildRootRelsXml(),
    'word/_rels/document.xml.rels': buildDocRelsXml(hasKopImage, hasQrImage),
    'word/settings.xml': buildSettingsXml(),
    'word/styles.xml': buildStylesXml(profile),
    'word/numbering.xml': buildNumberingXml(),
    'word/footer1.xml': buildFooter1Xml(),
    'word/footer2.xml': buildFooter2Xml(),
    'word/document.xml': buildDocumentXml(docModel, profile, hasKopImage, hasQrImage)
  };

  if (hasKopImage) {
    files['word/media/kop.png'] = kopParsed.bytes;
    files['word/media/logo.png'] = kopParsed.bytes;
  }

  if (hasQrImage) {
    files['word/media/qr.png'] = qrParsed.bytes;
  }

  return createZipPackage(files, outputType);
}
