
Berikut konsep lengkapnya, dari fondasi sampai alur kerja AI.

1. Konsep inti: AI tidak menulis file Word, AI menulis instruksi

Ini yang paling penting. Jangan minta AI menghasilkan XML atau file .docx langsung, karena hasilnya sering rusak dan tidak konsisten. Polanya:

Permintaan user → AI → JSON terstruktur → Renderer (kode biasa) → file .docx

AI hanya menentukan isi dan struktur. Kode yang deterministik yang mengurus format, penomoran, dan margin. Karena itu hasilnya rapi dan bisa diulang.

Contoh JSON yang dihasilkan AI:

json
{
  "type": "heading", "level": 1, "text": "PENDAHULUAN"
},
{
  "type": "paragraph", "text": "Praktik Kerja Lapangan merupakan..."
},
{
  "type": "figure", "src": "img_01", "caption": "Tampilan aplikasi"
}

Kamu menjalankan validasi skema (misalnya Zod atau JSON Schema) sebelum diberikan ke renderer.

2. Pahami isi file .docx (OOXML)

File .docx adalah ZIP berisi XML. Bagian pentingnya:

File	Fungsi
word/document.xml	Isi dokumen (paragraf, tabel, gambar)
word/styles.xml	Definisi style (Heading, Normal, Caption)
word/numbering.xml	Definisi penomoran dan list
word/header*.xml, footer*.xml	Header dan footer
word/media/	Gambar
word/_rels/	Relasi antar bagian
[Content_Types].xml	Daftar tipe konten

Konsep kunci yang bikin dokumen rapi: section (sectPr) untuk ukuran kertas, margin, dan nomor halaman yang berbeda per bagian (cover, romawi, angka), styles sebagai pusat format, dan field untuk nomor halaman, daftar isi, dan nomor gambar otomatis.

3. Dua mode kerja

Mode A: Buat dokumen baru. Ambil template, isi dengan konten dari AI, hasilkan .docx.

Mode B: Edit dokumen yang sudah ada. Ini yang lebih sulit tapi paling bernilai. Langkahnya:

Unzip .docx
Parse document.xml
Cari lokasi target (berdasarkan teks anchor, nomor paragraf, atau heading)
Sisipkan atau ganti elemen, dengan menyalin properti paragraf tetangga (pPr, rPr)
Zip ulang dan validasi

Jebakan klasik di Mode B: satu kalimat di Word sering terpecah jadi banyak <w:r> (run), jadi pencarian teks harus menggabungkan run dulu. Selain itu, jangan sampai merusak rsid, bookmark, atau relasi gambar.

4. Sistem penomoran dan struktur
   Heading bertingkat memakai style Heading 1/2/3 yang terhubung ke numbering.xml (1, 1.1, 1.1.1).
   Daftar isi memakai field TOC yang membaca Heading.
   Nomor gambar dan tabel memakai field SEQ (Gambar 1, Gambar 2, otomatis).
   Rujukan silang memakai bookmark + field REF, jadi kalau nomor berubah, rujukan ikut berubah.
   Nomor halaman memakai field PAGE, dengan section berbeda untuk format romawi dan angka.
   Daftar pustaka/sitasi: simpan data referensi terstruktur (BibTeX/CSL JSON), lalu render sesuai gaya (APA, IEEE).
5. Sistem template dan style

Buat template profil per jenis dokumen: laporan PKL, skripsi, jurnal, surat. Tiap profil menyimpan ukuran kertas, margin, font, spasi, gaya heading, gaya caption, dan aturan halaman (cover tanpa nomor, dst). User tinggal memilih profil, dan semua dokumen konsisten. Ini juga cara menangani pedoman kampus yang berbeda-beda.

6. Pipeline AI
   Outline generation: AI membuat kerangka bab dan subbab, user menyetujui.
   Section generation: tiap subbab ditulis terpisah (lebih cepat, bisa paralel, lebih terkontrol), dengan streaming ke UI.
   Konteks: kirim ringkasan dokumen, outline, dan gaya bahasa ke tiap panggilan supaya antar bagian nyambung.
   Rewrite/edit: user menyorot teks lalu minta “perbaiki”, “perpanjang”, “ubah jadi formal”. AI mengembalikan patch, bukan seluruh dokumen.
   RAG (opsional): user unggah PDF atau data, AI menulis berdasarkan sumber itu dan memberi sitasi.
   Pemeriksa konsistensi: contoh dari chat kamu tadi, deteksi nomor gambar dobel, rujukan yang salah, istilah yang tidak konsisten, atau judul yang huruf besar-kecilnya tidak seragam.

Gunakan structured output / tool calling supaya AI selalu mengembalikan JSON valid, dan jalankan validasi serta retry kalau gagal.

7. Validasi dan kontrol kualitas

Tahap ini yang membedakan sistem biasa dengan yang andal:

Validasi XML terhadap skema OOXML
Buka file dengan LibreOffice headless, lalu konversi ke PDF
Render halaman jadi gambar dan bandingkan sebelum/sesudah edit (visual diff), untuk membuktikan halaman lain tidak berubah
Cek otomatis: ukuran kertas A4, font tersedia, tidak ada paragraf kosong berlebih, nomor gambar berurutan
Tes buka di Word asli sebelum rilis
8. Arsitektur sistem
[Frontend]  ── chat panel + editor + preview
     │
[API Backend] ── auth, project, versioning
     │
 ├── [AI Service]       → LLM API, prompt, validasi JSON
 ├── [Document Engine]  → render & patch docx
 ├── [Converter]        → LibreOffice (preview PDF)
 └── [Storage]          → file, gambar, versi dokumen
     │
[Queue/Worker] ── tugas berat berjalan di background

Teknologi yang umum dipakai:

Frontend: React/Next.js, editor TipTap atau ProseMirror, preview dengan PDF.js atau docx-preview
Backend: Node.js (library docx) atau Python (python-docx, docxtpl, lxml)
Konversi dan preview: LibreOffice headless di container
Database: PostgreSQL; penyimpanan file: S3 atau sejenisnya
Antrian: Redis + BullMQ atau Celery
Sandbox: jalankan pemrosesan file di container terisolasi
9. Fitur pendukung
Versioning: setiap perubahan disimpan, bisa undo atau kembali ke versi lama
Track changes: AI menyisipkan perubahan sebagai revisi Word sehingga user bisa setuju atau tolak
Ekspor: DOCX, PDF, dan Markdown
Kolaborasi dan komentar (tahap lanjut)
Kontrol biaya: batasi token, cache hasil, pakai model kecil untuk tugas sederhana
10. Keamanan
Validasi dan batasi ukuran file unggahan (docx bisa berisi makro atau XML berbahaya)
Jangan eksekusi kode dari AI di server utama, selalu di sandbox
Pisahkan data antar user
Hati-hati dengan data sensitif yang dikirim ke API AI
11. Urutan membangun (roadmap)
MVP: pilih template → AI buat outline → AI isi bab → renderer menghasilkan .docx dengan heading, penomoran, daftar isi, nomor halaman
Preview PDF dan unduh
Editor di web + regenerate per bagian
Edit dokumen yang diunggah (Mode B) dengan visual diff
Daftar gambar/tabel, sitasi, pemeriksa konsistensi
Track changes, versioning, kolaborasi
