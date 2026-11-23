/**
 * numbering.js - Generator word/numbering.xml
 * Menyediakan definisi penomoran bertingkat asli Word:
 * - Multilevel Heading (BAB I -> 1.1 -> 1.1.1) terhubung langsung ke Heading styles (numId=1)
 * - Ordered Decimal (1., 2., 3.) (numId=2)
 * - Unordered / Bulleted (•) (numId=3)
 * - Ordered Huruf Kecil (a., b., c.) (numId=4)
 * - Ordered Huruf Besar (A., B., C.) (numId=5)
 * - Ordered Romawi Kecil (i., ii., iii.) (numId=6)
 * - Ordered Romawi Besar (I., II., III.) (numId=7)
 * - Custom startAt overrides (numId=10+) bila daftar dimulai dari angka > 1
 */

function buildSimpleListAbstractNum(abstractNumId, numFmt, lvlText) {
  return `
  <w:abstractNum w:abstractNumId="${abstractNumId}">
    <w:multiLevelType w:val="multilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="${numFmt}"/>
      <w:lvlText w:val="${lvlText}"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2) "/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="1080" w:hanging="360"/>
      </w:pPr>
    </w:lvl>
  </w:abstractNum>`.trim();
}

/**
 * Menghasilkan konten XML untuk word/numbering.xml
 * @param {Array<{numId: string|number, abstractNumId: string|number, startAt: number}>} [customOverrides]
 * @returns {string} XML string untuk numbering.xml
 */
export function buildNumberingXml(customOverrides = []) {
  const customNumsXml = (Array.isArray(customOverrides) ? customOverrides : [])
    .filter(o => o && o.numId && o.abstractNumId)
    .map(o => `
  <w:num w:numId="${o.numId}">
    <w:abstractNumId w:val="${o.abstractNumId}"/>
    <w:lvlOverride w:ilvl="0">
      <w:startOverride w:val="${Math.max(1, parseInt(o.startAt, 10) || 1)}"/>
    </w:lvlOverride>
  </w:num>`.trim())
    .join('\n\n  ');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

  <!-- 1. MULTILEVEL HEADINGS (BAB I, 1.1, 1.1.1) -->
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="multilevel"/>
    
    <!-- Level 0: BAB I, BAB II (Heading 1) -->
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="upperRoman"/>
      <w:pStyle w:val="Heading1"/>
      <w:lvlText w:val="BAB %1&#x000A;"/>
      <w:lvlJc w:val="center"/>
      <w:pPr>
        <w:jc w:val="center"/>
        <w:ind w:left="0" w:firstLine="0"/>
      </w:pPr>
      <w:rPr>
        <w:b/>
      </w:rPr>
    </w:lvl>

    <!-- Level 1: 1.1, 1.2 (Heading 2) -->
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:pStyle w:val="Heading2"/>
      <w:lvlText w:val="%1.%2 "/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="0" w:firstLine="0"/>
      </w:pPr>
      <w:rPr>
        <w:b/>
      </w:rPr>
    </w:lvl>

    <!-- Level 2: 1.1.1, 1.1.2 (Heading 3) -->
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:pStyle w:val="Heading3"/>
      <w:lvlText w:val="%1.%2.%3 "/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="0" w:firstLine="0"/>
      </w:pPr>
      <w:rPr>
        <w:b/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>

  <!-- 2. ORDERED LIST (1., 2., 3.) -->
  ${buildSimpleListAbstractNum(2, 'decimal', '%1. ')}

  <!-- 3. BULLETED LIST (•) -->
  <w:abstractNum w:abstractNumId="3">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#xF0B7;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="o"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="1080" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:hint="default"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>

  <!-- 4-7. VARIASI NUMBERING WORD (a., A., i., I.) -->
  ${buildSimpleListAbstractNum(4, 'lowerLetter', '%1. ')}
  ${buildSimpleListAbstractNum(5, 'upperLetter', '%1. ')}
  ${buildSimpleListAbstractNum(6, 'lowerRoman', '%1. ')}
  ${buildSimpleListAbstractNum(7, 'upperRoman', '%1. ')}

  <!-- Instance Penomoran Nyata (numId) -->
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>
  <w:num w:numId="3"><w:abstractNumId w:val="3"/></w:num>
  <w:num w:numId="4"><w:abstractNumId w:val="4"/></w:num>
  <w:num w:numId="5"><w:abstractNumId w:val="5"/></w:num>
  <w:num w:numId="6"><w:abstractNumId w:val="6"/></w:num>
  <w:num w:numId="7"><w:abstractNumId w:val="7"/></w:num>
  ${customNumsXml}

</w:numbering>`.trim();
}

