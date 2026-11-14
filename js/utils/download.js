/**
 * download.js - Utilitas pengunduhan file di browser
 */

/**
 * Memicu pengunduhan berkas Blob di browser
 * @param {Blob|Uint8Array} data 
 * @param {string} filename 
 * @param {string} mimeType 
 */
export function downloadFile(data, filename, mimeType = 'application/octet-stream') {
  if (typeof window === 'undefined' || typeof document === 'undefined') {
    throw new Error('downloadFile hanya dapat dipanggil di lingkungan browser.');
  }

  const blob = data instanceof Blob ? data : new Blob([data], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 100);
}

/**
 * Menyimpan teks sebagai file (misalnya JSON, Markdown)
 * @param {string} content 
 * @param {string} filename 
 * @param {string} mimeType 
 */
export function downloadText(content, filename, mimeType = 'text/plain;charset=utf-8') {
  const blob = new Blob([content], { type: mimeType });
  downloadFile(blob, filename, mimeType);
}

