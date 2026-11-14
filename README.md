# AI Document Builder (.docx)

Sistem pembuat dan pengedit dokumen Microsoft Word (`.docx`) standar formal dan akademik (Laporan PKL SMK, Skripsi, Jurnal Ilmiah) dengan tata letak baku, penomoran bertingkat otomatis, serta dibantu AI untuk menyusun kerangka dan menulis konten secara terstruktur.

---

## 🌟 Fitur Utama

1. **Mode A: Pembuat Dokumen Baru (Render JSON -> .docx)**
   - Menghasilkan dokumen Word asli yang rapi dan terstandarisasi.
   - Menggunakan fitur bawaan Word: *Multilevel Headings*, `numbering.xml`, field `PAGE`, field `TOC` (Daftar Isi otomatis ter-update), field `SEQ` (nomor gambar/tabel otomatis), dan bookmark + field `REF` (rujukan silang).
   - Pengaturan halaman otomatis per section (*Cover* tanpa nomor, *Bagian Awal* angka romawi `i, ii`, *Isi Bab* angka arab `1, 2, 3...`).
   - Profil baku: **Laporan PKL SMK** (A4, margin 4-4-3-3 cm, Times New Roman 12pt, spasi 1.5), **Skripsi** (spasi 2.0), dan **Jurnal Ilmiah**.

2. **Mode B: Edit Dokumen Word yang Diunggah (Safe Patcher)**
   - Mengedit isi dokumen tanpa merusak tata letak, margin, atau batas halaman (*sectPr*).
   - Pencarian berbasis anchor teks yang menggabungkan seluruh teks run (`<w:r>`) dalam satu paragraf.
   - Menyalin properti gaya paragraf tetangga (`pPr` dan `rPr`) agar format selaras sempurna.
   - Opsi penandaan **Track Changes** Word (`<w:ins>` dan `<w:del>`) agar revisi AI dapat disetujui/ditolak di Microsoft Word.
   - Laporan perubahan visual dan audit keselamatan dokumen otomatis.

3. **Alur Kerja AI Terstruktur (Anti-Halusinasi & Anti-Kerusakan)**
   - AI **TIDAK PERNAH** menghasilkan XML atau biner `.docx` langsung.
   - Alur: Permintaan User ➔ AI ➔ JSON (*Document Model*) ➔ Validasi Skema ➔ Renderer Deterministik ➔ `.docx`.
   - Tahap terarah: User input topik ➔ AI menyusun *Outline* ➔ User edit/setujui ➔ AI menulis tiap subbab berurutan dengan membawa ringkasan konteks bab sebelumnya ➔ Gabung menjadi *Document Model*.
   - Fitur *AI Rewrite* per blok: Perbaiki, Perpanjang, Persingkat, atau Formalkan gaya bahasa.
   - Otomatis melakukan *retry* maksimal 2 kali jika JSON dari AI tidak sesuai skema.

4. **Editor Blok Visual di Web**
   - Tambah, hapus, dan pindahkan posisi blok (*reorder*).
   - Riwayat *Undo* dan *Redo* (*history stack*).
   - Penyimpanan otomatis proyek ke **IndexedDB** browser.
   - Ekspor dan impor proyek dalam format `.json`.

5. **Pemeriksa Kualitas & Konsistensi Dokumen**
   - Deteksi nomor gambar/tabel dobel atau loncat (*SEQ*).
   - Deteksi rujukan silang yang hilang (*broken bookmarks*).
   - Deteksi heading level yang melompat (misal dari level 1 langsung ke level 3).
   - Deteksi paragraf kosong berlebih dan peringatan ukuran kertas non-A4.
   - Tombol *Auto-Fix* sekali klik untuk membersihkan masalah umum.

6. **Format Ekspor Tambahan**
   - Unduh berkas `.docx` asli.
   - Cetak langsung atau simpan sebagai PDF via `window.print()` dengan CSS `@media print`.
   - Ekspor ke format Markdown (`.md`).

---

## 🛠 Batasan Teknologi

- **Full JavaScript (ES Modules murni)**: Tanpa framework (tanpa React, Vue, Svelte, Angular).
- **Tanpa Build Step**: Tanpa Webpack, Vite, Babel, atau TypeScript. File dijalankan apa adanya.
- **Library Pihak Ketiga Lokal**: Disimpan di `/vendor/` (`jszip.min.js`, `docx-preview.min.js`). Tidak bergantung pada CDN runtime.
- **Serverless Backend Ringan**: Cloudflare Pages Functions (`/functions/api/ai.js`) sebagai proxy streaming dan pelindung API key.
- **BYOK (Bring Your Own Key)**: Tersedia mode input API key Google Gemini langsung di browser dan disimpan di `localStorage` untuk development lokal.

---

## 🚀 Cara Menjalankan Secara Lokal

Karena aplikasi ini murni file statis ES Modules tanpa proses build, Anda dapat menjalankannya dengan server statis lokal apa pun.

### Opsi 1: Menggunakan Node `serve`
```bash
npx serve . -l 3000
```
Buka browser di: `http://localhost:3000`

### Opsi 2: Menggunakan Python
```bash
# Python 3
python -m http.server 3000
```
Buka browser di: `http://localhost:3000`

### Opsi 3: Menggunakan VS Code Live Server
Cukup klik kanan pada `index.html` dan pilih **Open with Live Server**.

---

## 🧪 Menjalankan Pengujian Otomatis

Aplikasi menggunakan runner tes bawaan Node.js (`node --test`) tanpa dependensi Jest/Vitest.

Jalankan perintah berikut di terminal:
```bash
npm test
```
atau:
```bash
node --test tests/*.test.js
```

Tes mencakup:
- `tests/schema.test.js`: Validasi skema Document Model dan normalisasi input.
- `tests/render.test.js`: Validitas pembuatan paket berkas `.docx` (Mode A) dan kepatuhan elemen OOXML.
- `tests/patch.test.js`: Validitas manipulasi berkas `.docx` eksisting (Mode B) dan penandaan Track Changes.
- `tests/consistency.test.js`: Algoritma pendeteksi inkonsistensi penomoran, rujukan, dan auto-fix.

---

## ☁ Cara Deploy ke Cloudflare Pages

1. **Push Proyek ke GitHub / GitLab**.
2. Masuk ke dashboard **Cloudflare** > **Workers & Pages** > **Create application** > **Pages** > **Connect to Git**.
3. Pilih repositori proyek ini.
4. Pada bagian **Build settings**:
   - **Framework preset**: *None*
   - **Build command**: *(Kosongkan, tidak ada build step)*
   - **Build output directory**: `.` *(titik atau root)*
5. Klik **Save and Deploy**. Cloudflare Pages secara otomatis akan mendeteksi `/functions/api/ai.js` sebagai serverless function proxy API AI.

### Mengatur Environment Variable API Key di Cloudflare:
1. Di dashboard proyek Cloudflare Pages Anda, buka tab **Settings** > **Environment variables**.
2. Tambahkan variabel baru:
   - **Variable name**: `AI_API_KEY`
   - **Value**: Masukkan API Key Google Gemini Anda (dapat diperoleh gratis dari [Google AI Studio](https://aistudio.google.com/)).
3. Klik **Save**.
4. Lakukan *Redeploy* agar fungsi serverless dapat membaca variabel lingkungan tersebut.

*(Catatan: Pengguna juga dapat menggunakan mode BYOK langsung melalui menu **Pengaturan (BYOK)** di pojok kanan antarmuka web tanpa perlu menyetel server backend).*

---

## 📂 Struktur Direktori

```
/
├── index.html                 # UI Aplikasi Utama (Layout responsif, pratinjau, editor)
├── css/
│   └── app.css                # Styling antarmuka & print CSS (@media print)
├── js/
│   ├── main.js                # Bootstrap aplikasi & router navigasi
│   ├── state.js               # Reactive store, history undo/redo, auto-save IndexedDB
│   ├── ui/                    # Komponen antarmuka modular
│   │   ├── preview.js         # Renderer docx-preview di browser
│   │   ├── editor.js          # Visual block editor & per-block AI rewrite
│   │   ├── patchView.js       # UI Mode B (upload & patch docx)
│   │   ├── aiView.js          # UI alur kerja AI (outline -> subbab -> doc)
│   │   ├── consistencyView.js # UI pemeriksa inkonsistensi & auto-fix
│   │   ├── settingsView.js    # UI pengaturan API Key BYOK & model
│   │   └── notification.js    # Toast notifikasi bahasa Indonesia
│   ├── ai/
│   │   ├── client.js          # Klien API AI, streaming SSE, retry max 2x
│   │   └── prompts.js         # Prompt terstruktur (outline, section, rewrite)
│   ├── model/
│   │   ├── schema.js          # Validator mandiri Document Model
│   │   └── sample.js          # Contoh dokumen PKL lengkap
│   ├── docx/
│   │   ├── render.js          # Mode A: Document Model -> OOXML (.docx)
│   │   ├── patch.js           # Mode B: Patcher docx dengan kloning pPr/rPr
│   │   ├── styles.js          # word/styles.xml dinamis berbasis profil
│   │   ├── numbering.js       # word/numbering.xml penomoran bertingkat
│   │   ├── fields.js          # Field codes (PAGE, TOC, SEQ, REF, Bookmarks)
│   │   └── profiles/          # Profil dokumen baku (pkl.js, skripsi.js, jurnal.js)
│   ├── checks/
│   │   └── consistency.js     # Pemeriksa inkonsistensi struktur dokumen
│   └── utils/
│       ├── xml.js             # Sanitasi XML & helper DOM parser
│       ├── zip.js             # Wrapper JSZip universal (Browser & Node)
│       └── download.js        # Utilitas download berkas
├── functions/
│   └── api/
│       └── ai.js              # Serverless Function Cloudflare Pages (Proxy AI & SSE)
├── vendor/
│   ├── jszip.min.js           # Pustaka kompresi ZIP lokal
│   └── docx-preview.min.js    # Pustaka pratinjau docx lokal
├── tests/
│   ├── schema.test.js         # Pengujian skema
│   ├── render.test.js         # Pengujian renderer docx
│   ├── patch.test.js          # Pengujian mode B patcher
│   └── consistency.test.js    # Pengujian pemeriksa kualitas
├── package.json
└── README.md
```

# web-surat-copilot
