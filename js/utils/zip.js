/**
 * zip.js - Wrapper JSZip universal (Browser & Node.js)
 */

/**
 * Mendapatkan constructor JSZip baik di Browser maupun Node.js
 * @returns {Promise<typeof import('jszip')>}
 */
export async function getJSZip() {
  if (typeof window !== 'undefined' && window.JSZip) {
    return window.JSZip;
  }
  if (typeof globalThis !== 'undefined' && globalThis.JSZip) {
    return globalThis.JSZip;
  }
  // Fallback untuk lingkungan Node.js
  try {
    const mod = await import('jszip');
    return mod.default || mod;
  } catch (err) {
    throw new Error('JSZip tidak ditemukan. Pastikan vendor/jszip.min.js dimuat atau paket jszip terpasang.');
  }
}

/**
 * Membuat paket ZIP dari daftar file { path: content }
 * @param {Record<string, string|Uint8Array|Blob>} files 
 * @param {'blob'|'uint8array'|'nodebuffer'} outputType 
 * @returns {Promise<Blob|Uint8Array|Buffer>}
 */
export async function createZipPackage(files, outputType = 'blob') {
  const JSZip = await getJSZip();
  const zip = new JSZip();

  for (const [path, content] of Object.entries(files)) {
    zip.file(path, content);
  }

  return zip.generateAsync({
    type: outputType,
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
  });
}

/**
 * Membaca paket ZIP menjadi kumpulan file
 * @param {Blob|Uint8Array|ArrayBuffer} data 
 * @returns {Promise<any>} Objek JSZip instance
 */
export async function readZipPackage(data) {
  const JSZip = await getJSZip();
  const zip = new JSZip();
  return zip.loadAsync(data);
}

