/**
 * scripts/build-netlify.mjs - Menyiapkan folder publish dist/ yang bersih untuk deploy Netlify
 * Hanya menyalin aset runtime aplikasi dan mengecualikan folder sensitif/pengembangan (API/, tests/, node_modules/).
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const rootDir = path.resolve(__dirname, '..');
const distDir = path.join(rootDir, 'dist');

const ROOT_FILES = [
  'index.html',
  'manifest.webmanifest',
  'sw.js',
  'robots.txt',
  '_headers',
  '_redirects',
  '404.html'
];

const ROOT_DIRS = [
  'css',
  'img',
  'js',
  'vendor'
];

/**
 * Menghapus dan menyusun ulang folder dist/ dengan aset produksi saja.
 */
function buildDist() {
  fs.rmSync(distDir, { recursive: true, force: true });
  fs.mkdirSync(distDir, { recursive: true });

  for (const fileName of ROOT_FILES) {
    const srcPath = path.join(rootDir, fileName);
    if (!fs.existsSync(srcPath)) {
      throw new Error(`Berkas wajib tidak ditemukan: ${fileName}`);
    }
    fs.copyFileSync(srcPath, path.join(distDir, fileName));
  }

  for (const dirName of ROOT_DIRS) {
    const srcDir = path.join(rootDir, dirName);
    if (!fs.existsSync(srcDir)) {
      throw new Error(`Direktori wajib tidak ditemukan: ${dirName}`);
    }
    fs.cpSync(srcDir, path.join(distDir, dirName), { recursive: true });
  }

  // Validasi bahwa seluruh aset precache di sw.js tersedia di dist/
  const swText = fs.readFileSync(path.join(distDir, 'sw.js'), 'utf8');
  const precacheMatch = swText.match(/const PRECACHE_ASSETS = \[([\s\S]*?)\];/);
  if (precacheMatch) {
    const assetMatches = [...precacheMatch[1].matchAll(/'([^']+)'/g)].map(m => m[1]);
    for (const relAsset of assetMatches) {
      if (relAsset === './') continue;
      const normalized = relAsset.replace(/^\.\//, '');
      const fullTarget = path.join(distDir, normalized);
      if (!fs.existsSync(fullTarget)) {
        throw new Error(`Aset PRECACHE sw.js tidak ditemukan di dist/: ${normalized}`);
      }
    }
  }

  console.log('[build-netlify] Folder dist/ siap deploy (bersih tanpa API/, tests/, node_modules/).');
}

buildDist();
