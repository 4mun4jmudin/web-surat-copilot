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
    if (run.sup) rPrParts.push('<w:vertAlign w:val="superscript"/>');
    if (run.sub) rPrParts.push('<w:vertAlign w:val="subscript"/>');

    const rPrXml = rPrParts.length > 0 ? `<w:rPr>${rPrParts.join('')}</w:rPr>` : '';
    const textXml = buildTextElement(run.text || '');

    let runXml = `<w:r>${rPrXml}${textXml}</w:r>`;

    if (run.bookmark) {
      const bId = String(idx + 100);
      runXml = `${buildBookmarkStart(bId, run.bookmark)}${runXml}${buildBookmarkEnd(bId)}`;
    }

    return runXml;
  }).join('');
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

export function renderDocumentBody(doc, profile, hasKopImage = false) {
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
        bodyXmlParts.push(buildIdentityTableXml(block.rows || []));
        break;
      }

      // 4. BLOK TANDA TANGAN RESMI
      case 'letterSignature': {
        bodyXmlParts.push(buildLetterSignatureXml(block));
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
        const spacingAfter = isLetterBody ? 120 : (profile.typography?.paragraphSpacingAfter || 120);

        bodyXmlParts.push(`
          <w:p>
            <w:pPr>
              <w:pStyle w:val="Normal"/>
              ${indentVal ? `<w:ind w:firstLine="${indentVal}"/>` : ''}
              <w:spacing w:line="${lineSpacing}" w:lineRule="auto" w:after="${spacingAfter}"/>
              <w:jc w:val="${profile.typography?.alignment || 'both'}"/>
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

        const rowsXml = (block.rows || []).map((row, rIdx) => {
          const isHeader = rIdx === 0 && block.headerRow !== false;
          const cellsXml = (row || []).map(cell => {
            const cellText = typeof cell === 'object' && cell.text !== undefined ? cell.text : String(cell);
            return `
              <w:tc>
                <w:tcPr>
                  <w:tcMar>
                    <w:top w:w="100" w:type="dxa"/>
                    <w:bottom w:w="100" w:type="dxa"/>
                    <w:left w:w="120" w:type="dxa"/>
                    <w:right w:w="120" w:type="dxa"/>
                  </w:tcMar>
                  ${isHeader ? '<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>' : ''}
                </w:tcPr>
                <w:p>
                  <w:pPr>
                    <w:pStyle w:val="Normal"/>
                    <w:spacing w:line="240" w:lineRule="auto" w:after="0"/>
                    <w:jc w:val="${isHeader ? 'center' : 'left'}"/>
                  </w:pPr>
                  <w:r>
                    ${isHeader ? '<w:rPr><w:b/></w:rPr>' : ''}
                    <w:t>${escapeXml(cellText)}</w:t>
                  </w:r>
                </w:p>
              </w:tc>
            `;
          }).join('');

          return `
            <w:tr>
              <w:trPr>${isHeader ? '<w:tblHeader/>' : ''}</w:trPr>
              ${cellsXml}
            </w:tr>
          `;
        }).join('');

        bodyXmlParts.push(`
          <w:tbl>
            <w:tblPr>
              <w:tblStyle w:val="TableNormal"/>
              <w:tblW w:w="0" w:type="auto"/>
              <w:tblBorders>
                <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
                <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
              </w:tblBorders>
              <w:jc w:val="center"/>
            </w:tblPr>
            ${rowsXml}
          </w:tbl>
        `);
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

export function buildDocumentXml(doc, profile, hasKopImage = false) {
  const bodyXml = renderDocumentBody(doc, profile, hasKopImage);
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

export function buildContentTypesXml(hasKopImage = false) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${hasKopImage ? '<Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="svg" ContentType="image/svg+xml"/>' : ''}
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

export function buildDocRelsXml(hasKopImage = false) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rIdFooter2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>
  ${hasKopImage ? '<Relationship Id="rIdKop" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/kop.png"/><Relationship Id="rIdLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.png"/>' : ''}
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

  const hasKopImage = !!kopParsed;

  const files = {
    '[Content_Types].xml': buildContentTypesXml(hasKopImage),
    '_rels/.rels': buildRootRelsXml(),
    'word/_rels/document.xml.rels': buildDocRelsXml(hasKopImage),
    'word/settings.xml': buildSettingsXml(),
    'word/styles.xml': buildStylesXml(profile),
    'word/numbering.xml': buildNumberingXml(),
    'word/footer1.xml': buildFooter1Xml(),
    'word/footer2.xml': buildFooter2Xml(),
    'word/document.xml': buildDocumentXml(docModel, profile, hasKopImage)
  };

  if (hasKopImage) {
    files['word/media/kop.png'] = kopParsed.bytes;
    files['word/media/logo.png'] = kopParsed.bytes;
  }

  return createZipPackage(files, outputType);
}
