/**
 * numbering.js - Generator word/numbering.xml
 * Menyediakan definisi penomoran bertingkat asli Word:
 * - Multilevel Heading (BAB I -> 1.1 -> 1.1.1) terhubung langsung ke Heading styles
 * - Ordered list (1, 2, 3 / a, b, c)
 * - Unordered list (bullets)
 */

/**
 * Menghasilkan konten XML untuk word/numbering.xml
 * @returns {string} XML string untuk numbering.xml
 */
export function buildNumberingXml() {
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
  <w:abstractNum w:abstractNumId="2">
    <w:multiLevelType w:val="multilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1. "/>
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
  </w:abstractNum>

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

  <!-- Instance Penomoran Nyata (numId) -->
  <w:num w:numId="1">
    <w:abstractNumId w:val="1"/>
  </w:num>
  
  <w:num w:numId="2">
    <w:abstractNumId w:val="2"/>
  </w:num>

  <w:num w:numId="3">
    <w:abstractNumId w:val="3"/>
  </w:num>

</w:numbering>`.trim();
}

