/**
 * preview.js - Pengelola Pratinjau Dokumen .docx di Browser (docx-preview)
 */

import { renderDocx } from '../docx/render.js';
import { showToast } from './notification.js';

let currentDocxBlob = null;

/**
 * Merender Document Model ke kontainer pratinjau
 * @param {object} docModel 
 * @param {HTMLElement} containerElement 
 */
export async function updateDocumentPreview(docModel, containerElement) {
  if (!containerElement) return;

  containerElement.innerHTML = `
    <div class="preview-loading">
      <div class="spinner"></div>
      <p>Menyusun dan merender pratinjau dokumen...</p>
    </div>
  `;

  try {
    const docxBlob = await renderDocx(docModel, 'blob');
    currentDocxBlob = docxBlob;

    containerElement.innerHTML = '';
    
    // Gunakan docx-preview dari window.docx
    if (typeof window.docx !== 'undefined' && window.docx.renderAsync) {
      await window.docx.renderAsync(docxBlob, containerElement, null, {
        className: 'docx-preview-wrapper',
        inWrapper: true,
        ignoreWidth: false,
        ignoreHeight: false,
        breakPages: true,
        useBase64URL: true
      });
    } else {
      // Fallback jika vendor library belum selesai load
      containerElement.innerHTML = `
        <div class="preview-card-fallback">
          <h3>${docModel.meta?.title || 'Dokumen Word'}</h3>
          <p><em>Profil: ${docModel.profile?.toUpperCase()} | ${docModel.blocks?.length || 0} Blok Konten</em></p>
          <div class="alert alert-info">
            Pratinjau visual memerlukan <code>vendor/docx-preview.min.js</code>. Anda tetap dapat mengunduh berkas .docx utuh secara langsung.
          </div>
        </div>
      `;
    }
  } catch (err) {
    console.error('Error saat merender preview:', err);
    containerElement.innerHTML = `
      <div class="preview-error">
        <h4>Gagal menampilkan pratinjau</h4>
        <p>${err.message}</p>
      </div>
    `;
    showToast(`Gagal merender pratinjau: ${err.message}`, 'error');
  }
}

/**
 * Mendapatkan Blob berkas docx yang terakhir dirender
 */
export function getCurrentDocxBlob() {
  return currentDocxBlob;
}

