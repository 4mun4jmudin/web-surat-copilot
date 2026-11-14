/**
 * xml.js - Utilitas manipulasi dan sanitasi XML/OOXML
 * Berjalan universal di Browser dan Node.js
 */

/**
 * Escape karakter spesial XML agar aman dari kerusakan struktur / injeksi XML.
 * @param {string} unsafe 
 * @returns {string}
 */
export function escapeXml(unsafe) {
  if (unsafe === null || unsafe === undefined) return '';
  return String(unsafe)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Menghasilkan elemen `<w:t>` dengan preservasi spasi jika ada spasi di awal/akhir
 * @param {string} text 
 * @returns {string}
 */
export function buildTextElement(text) {
  if (!text) return '<w:t></w:t>';
  const escaped = escapeXml(text);
  const needsPreserve = /^\s|\s$/.test(text) || text.length === 0;
  return needsPreserve
    ? `<w:t xml:space="preserve">${escaped}</w:t>`
    : `<w:t>${escaped}</w:t>`;
}

/**
 * Dapatkan implementasi DOMParser (baik bawaan browser maupun modul di Node)
 */
export async function getDOMParser() {
  if (typeof DOMParser !== 'undefined') {
    return DOMParser;
  }
  if (typeof globalThis !== 'undefined' && globalThis.DOMParser) {
    return globalThis.DOMParser;
  }
  try {
    const xmldom = await import('@xmldom/xmldom');
    globalThis.DOMParser = xmldom.DOMParser;
    globalThis.XMLSerializer = xmldom.XMLSerializer;
    return xmldom.DOMParser;
  } catch (err) {
    throw new Error('DOMParser tidak tersedia di lingkungan ini.');
  }
}

/**
 * Dapatkan implementasi XMLSerializer
 */
export async function getXMLSerializer() {
  if (typeof XMLSerializer !== 'undefined') {
    return XMLSerializer;
  }
  if (typeof globalThis !== 'undefined' && globalThis.XMLSerializer) {
    return globalThis.XMLSerializer;
  }
  try {
    const xmldom = await import('@xmldom/xmldom');
    globalThis.XMLSerializer = xmldom.XMLSerializer;
    return xmldom.XMLSerializer;
  } catch (err) {
    throw new Error('XMLSerializer tidak tersedia di lingkungan ini.');
  }
}

/**
 * Mengubah string XML menjadi DOM Document.
 * @param {string} xmlString 
 * @returns {Promise<Document>}
 */
export async function parseXmlAsync(xmlString) {
  const Parser = await getDOMParser();
  const parser = new Parser();
  const doc = parser.parseFromString(xmlString, 'application/xml');
  const parserError = doc.getElementsByTagName('parsererror')[0];
  if (parserError) {
    throw new Error(`XML Parse Error: ${parserError.textContent}`);
  }
  return doc;
}

/**
 * Mengubah XML DOM Document kembali menjadi string XML.
 * @param {Document|Node} domNode 
 * @returns {Promise<string>}
 */
export async function serializeXmlAsync(domNode) {
  const Serializer = await getXMLSerializer();
  const serializer = new Serializer();
  return serializer.serializeToString(domNode);
}
