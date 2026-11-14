/**
 * styles.js - Generator word/styles.xml berbasis Profile Dokumen
 * Mendefinisikan style terpusat (Heading 1-3, Normal, Caption, Header, Footer)
 */

import { escapeXml } from '../utils/xml.js';

/**
 * Menghasilkan konten lengkap berkas word/styles.xml
 * @param {object} profile 
 * @returns {string} XML string untuk word/styles.xml
 */
export function buildStylesXml(profile) {
  const fontPrimary = profile.fonts?.primary || 'Times New Roman';
  const bodySize = profile.typography?.bodySize || 24;
  const lineSpacing = profile.typography?.lineSpacing || 276;
  const spacingAfter = profile.typography?.paragraphSpacingAfter || 120;
  const align = profile.typography?.alignment || 'both';

  const h1 = profile.headings?.heading1 || { size: 28, spacingBefore: 240, spacingAfter: 240, alignment: 'center', allCaps: true };
  const h2 = profile.headings?.heading2 || { size: 24, spacingBefore: 180, spacingAfter: 120, alignment: 'left' };
  const h3 = profile.headings?.heading3 || { size: 24, spacingBefore: 120, spacingAfter: 60, alignment: 'left' };
  const cap = profile.caption || { size: 20, spacingBefore: 60, spacingAfter: 120, alignment: 'center' };

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="${escapeXml(fontPrimary)}" w:hAnsi="${escapeXml(fontPrimary)}" w:cs="${escapeXml(fontPrimary)}"/>
        <w:sz w:val="${bodySize}"/>
        <w:szCs w:val="${bodySize}"/>
        <w:lang w:val="id-ID"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:line="${lineSpacing}" w:lineRule="auto" w:after="${spacingAfter}"/>
        <w:jc w:val="${align}"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>

  <!-- Style Normal -->
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:line="${lineSpacing}" w:lineRule="auto" w:after="${spacingAfter}"/>
      <w:jc w:val="${align}"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="${escapeXml(fontPrimary)}" w:hAnsi="${escapeXml(fontPrimary)}"/>
      <w:sz w:val="${bodySize}"/>
      <w:szCs w:val="${bodySize}"/>
    </w:rPr>
  </w:style>

  <!-- Style Heading 1 -->
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="${h1.spacingBefore || 240}" w:after="${h1.spacingAfter || 240}"/>
      <w:jc w:val="${h1.alignment || 'center'}"/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="${escapeXml(fontPrimary)}" w:hAnsi="${escapeXml(fontPrimary)}"/>
      <w:b/>
      ${h1.allCaps ? '<w:caps/>' : ''}
      <w:sz w:val="${h1.size || 28}"/>
      <w:szCs w:val="${h1.size || 28}"/>
    </w:rPr>
  </w:style>

  <!-- Style Heading 2 -->
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="${h2.spacingBefore || 180}" w:after="${h2.spacingAfter || 120}"/>
      <w:jc w:val="${h2.alignment || 'left'}"/>
      <w:outlineLvl w:val="1"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="${escapeXml(fontPrimary)}" w:hAnsi="${escapeXml(fontPrimary)}"/>
      <w:b/>
      <w:sz w:val="${h2.size || 24}"/>
      <w:szCs w:val="${h2.size || 24}"/>
    </w:rPr>
  </w:style>

  <!-- Style Heading 3 -->
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="${h3.spacingBefore || 120}" w:after="${h3.spacingAfter || 60}"/>
      <w:jc w:val="${h3.alignment || 'left'}"/>
      <w:outlineLvl w:val="2"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="${escapeXml(fontPrimary)}" w:hAnsi="${escapeXml(fontPrimary)}"/>
      <w:b/>
      ${h3.italic ? '<w:i/>' : ''}
      <w:sz w:val="${h3.size || 24}"/>
      <w:szCs w:val="${h3.size || 24}"/>
    </w:rPr>
  </w:style>

  <!-- Style Caption -->
  <w:style w:type="paragraph" w:styleId="Caption">
    <w:name w:val="caption"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="${cap.spacingBefore || 60}" w:after="${cap.spacingAfter || 120}"/>
      <w:jc w:val="${cap.alignment || 'center'}"/>
    </w:pPr>
    <w:rPr>
      <w:i/>
      <w:sz w:val="${cap.size || 20}"/>
      <w:szCs w:val="${cap.size || 20}"/>
    </w:rPr>
  </w:style>

  <!-- Style Header -->
  <w:style w:type="paragraph" w:styleId="Header">
    <w:name w:val="header"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:sz w:val="18"/>
      <w:szCs w:val="18"/>
    </w:rPr>
  </w:style>

  <!-- Style Footer -->
  <w:style w:type="paragraph" w:styleId="Footer">
    <w:name w:val="footer"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:sz w:val="20"/>
      <w:szCs w:val="20"/>
    </w:rPr>
  </w:style>

  <!-- Table Normal -->
  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
    <w:name w:val="Normal Table"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:tblPr>
      <w:tblInd w:w="0" w:type="dxa"/>
      <w:tblCellMar>
        <w:top w:w="120" w:type="dxa"/>
        <w:left w:w="160" w:type="dxa"/>
        <w:bottom w:w="120" w:type="dxa"/>
        <w:right w:w="160" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>

</w:styles>
`.trim();
}
