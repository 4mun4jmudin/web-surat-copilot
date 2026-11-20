/**
 * htmlAst.js - Parser & Serializer AST HTML ringan (zero-dependency, berjalan di Browser & Node.js)
 * Digunakan oleh Smart Word Pro Canvas & AI Surgical Patcher agar struktur tabel, <colgroup>,
 * colspan/rowspan, tinggi baris, atribut data-*, dan slot QR tetap utuh 100% saat diedit.
 */

export const VOID_TAGS = new Set([
  'br', 'hr', 'img', 'input', 'meta', 'link', 'col', 'wbr', 'source'
]);

export function escapeHtmlText(str = '') {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export function escapeHtmlTextContent(str = '') {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

export function decodeHtmlEntities(str = '') {
  return String(str)
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&#(\d+);/g, (_, dec) => String.fromCharCode(Number(dec)))
    .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
}

function parseAttributes(attrStr = '') {
  const attrs = {};
  const re = /([a-zA-Z0-9_:-]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+)))?/g;
  let m;
  while ((m = re.exec(attrStr)) !== null) {
    const key = m[1].toLowerCase();
    const val = m[2] ?? m[3] ?? m[4] ?? '';
    attrs[key] = decodeHtmlEntities(val);
  }
  return attrs;
}

/**
 * Mengurai string HTML menjadi daftar node AST:
 * `{ type: 'element' | 'text', tag, attrs, children, text }`.
 */
export function parseHtmlNodes(html = '') {
  const root = { type: 'element', tag: 'root', attrs: {}, children: [] };
  const stack = [root];
  const tokenRe = /<!--[\s\S]*?-->|<\/?([a-zA-Z][a-zA-Z0-9]*)([^>]*)>|([^<]+)/g;
  let match;

  while ((match = tokenRe.exec(String(html))) !== null) {
    if (match[3] !== undefined) {
      const txt = decodeHtmlEntities(match[3]);
      if (txt) {
        stack[stack.length - 1].children.push({ type: 'text', text: txt });
      }
      continue;
    }
    if (!match[1]) continue;

    const fullTag = match[0];
    const tag = match[1].toLowerCase();
    const rawAttrs = match[2] || '';
    const isClosing = fullTag.startsWith('</');
    const isSelfClosing = fullTag.endsWith('/>') || VOID_TAGS.has(tag);

    if (isClosing) {
      for (let i = stack.length - 1; i >= 1; i--) {
        if (stack[i].tag === tag) {
          stack.length = i;
          break;
        }
      }
    } else {
      const node = {
        type: 'element',
        tag,
        attrs: parseAttributes(rawAttrs),
        children: []
      };
      stack[stack.length - 1].children.push(node);
      if (!isSelfClosing) {
        stack.push(node);
      }
    }
  }

  return root.children;
}

function serializeAttrEntry(key, val) {
  if (val === undefined || val === null) return '';
  return `${key}="${escapeHtmlText(String(val))}"`;
}

/**
 * Menyerialisasi kembali daftar node AST menjadi HTML valid dengan mempertahankan semua tag & atribut.
 */
export function serializeHtmlNodes(nodes = []) {
  if (!Array.isArray(nodes) || nodes.length === 0) return '';
  return nodes.map((node) => {
    if (!node) return '';
    if (node.type === 'text') {
      return escapeHtmlTextContent(node.text || '');
    }
    if (node.type !== 'element') return '';
    const tag = node.tag || 'div';
    const attrParts = Object.entries(node.attrs || {})
      .map(([k, v]) => serializeAttrEntry(k, v))
      .filter(Boolean);
    const attrStr = attrParts.length > 0 ? ` ${attrParts.join(' ')}` : '';
    if (VOID_TAGS.has(tag)) {
      return `<${tag}${attrStr}>`;
    }
    const inner = serializeHtmlNodes(node.children || []);
    return `<${tag}${attrStr}>${inner}</${tag}>`;
  }).join('');
}

export function getNodePlainText(node) {
  if (!node) return '';
  if (node.type === 'text') return node.text;
  if (node.tag === 'br') return '\n';
  return (node.children || []).map(getNodePlainText).join('');
}

export function cloneAstNode(node) {
  if (!node) return null;
  if (node.type === 'text') {
    return { type: 'text', text: node.text || '' };
  }
  return {
    type: 'element',
    tag: node.tag,
    attrs: { ...(node.attrs || {}) },
    children: (node.children || []).map(cloneAstNode).filter(Boolean)
  };
}
