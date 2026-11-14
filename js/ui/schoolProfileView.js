/**
 * schoolProfileView.js - Formulir Pengaturan Profil Sekolah & Gudang Kop Surat
 * Tersimpan permanen di IndexedDB.
 */

import { store } from '../state.js';
import { showToast } from './notification.js';

export function renderSchoolProfileView(container) {
  if (!container) return;

  const school = store.getState().schoolProfile;
  const kopGallery = store.getKopGallery();

  container.innerHTML = `
    <div class="school-profile-container" style="max-width: 900px; margin: 24px auto; padding: 0 16px; overflow-y: auto; height: 100%;">
      <div class="card" style="background: white; border-radius: 8px; padding: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.06);">
        <h2 style="font-size: 20px; font-weight: 700; color: var(--dark); display: flex; align-items: center; gap: 8px;">
          <span>🏫</span> Profil Sekolah & Pengaturan Kop Surat
        </h2>
        <p class="text-muted" style="margin-top: 4px; font-size: 13px;">
          Data ini menjadi kop resmi dan identitas penandatangan pada seluruh dokumen dan surat dinas.
          Tersimpan permanen di perangkat browser Anda (IndexedDB).
        </p>

        <!-- KOLEKSI GUDANG KOP SURAT TERSIMPAN -->
        <div style="margin: 20px 0; background: #f8fafc; border: 1px solid var(--border); border-radius: 6px; padding: 14px;">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
            <strong style="font-size: 14px; color: var(--dark);">🖼 Gudang Banner Kop Tersimpan (${kopGallery.length})</strong>
            <span style="font-size: 12px; color: var(--text-muted);">Pilih salah satu untuk menjadi kop aktif:</span>
          </div>

          <div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(240px, 1fr)); gap: 10px;">
            ${kopGallery.map(k => {
              const isActive = (school.kopImage || '') === k.dataUrl;
              return `
                <div style="border: 2px solid ${isActive ? 'var(--primary)' : 'var(--border)'}; border-radius: 6px; padding: 8px; background: white; display: flex; flex-direction: column; justify-content: space-between;">
                  <div>
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px;">
                      <span style="font-size: 10px; font-weight: 700; background: #e2e8f0; padding: 1px 5px; border-radius: 3px;">${escapeHtml(k.category || 'Umum')}</span>
                      ${isActive ? '<span style="color: #166534; font-size: 10.5px; font-weight: bold;">✓ Aktif</span>' : ''}
                    </div>
                    <strong style="font-size: 12px; display: block; margin-bottom: 6px;">${escapeHtml(k.name)}</strong>
                    <div style="border: 1px solid #e2e8f0; border-radius: 4px; padding: 4px; background: #fafafa;">
                      <img src="${k.dataUrl}" alt="${escapeHtml(k.name)}" style="width: 100%; height: 50px; object-fit: contain; display: block;">
                    </div>
                  </div>
                  <button type="button" class="btn btn-sm ${isActive ? 'btn-success' : 'btn-secondary'} btn-apply-school-kop" data-id="${k.id}" style="margin-top: 8px; width: 100%;">
                    ${isActive ? '✓ Kop Aktif Saat Ini' : 'Gunakan Kop Ini'}
                  </button>
                </div>
              `;
            }).join('')}
          </div>
        </div>

        <form id="form-school-profile">
          <h4 style="margin: 20px 0 10px 0; border-bottom: 1px solid #e2e8f0; padding-bottom: 6px;">1. Instansi Pembina, Yayasan & Legalitas Sekolah</h4>
          <div class="form-row">
            <div class="form-group col">
              <label>Pemerintah Daerah / Yayasan Naungan:</label>
              <input type="text" name="government" class="form-control" value="${escapeHtml(school.government || 'PEMERINTAH DAERAH PROVINSI JAWA BARAT')}">
            </div>
            <div class="form-group col">
              <label>Yayasan Penyelenggara:</label>
              <input type="text" name="foundation" class="form-control" value="${escapeHtml(school.foundation || 'Yayasan Nunuk Kartini Al-Irsyad')}">
            </div>
          </div>
          <div class="form-row">
            <div class="form-group col">
              <label>Dinas Pendidikan:</label>
              <input type="text" name="department" class="form-control" value="${escapeHtml(school.department || 'DINAS PENDIDIKAN')}">
            </div>
            <div class="form-group col">
              <label>Cabang Dinas (Opsional):</label>
              <input type="text" name="subOffice" class="form-control" value="${escapeHtml(school.subOffice || 'CABANG DINAS PENDIDIKAN WILAYAH XI')}">
            </div>
          </div>
          <div class="form-row">
            <div class="form-group col">
              <label>Nama Resmi Sekolah:</label>
              <input type="text" name="schoolName" class="form-control" value="${escapeHtml(school.schoolName || 'SMKS IT AL-HAWARI')}">
            </div>
            <div class="form-group col">
              <label>NPSN:</label>
              <input type="text" name="npsn" class="form-control" value="${escapeHtml(school.npsn || '69758993')}">
            </div>
          </div>
          <div class="form-row">
            <div class="form-group col">
              <label>SK Operasional & Tanggal:</label>
              <input type="text" name="skOperasional" class="form-control" value="${escapeHtml(school.skOperasional || '425.11/1136-DISDIK (30 April 2013)')}">
            </div>
            <div class="form-group col">
              <label>Status & Akreditasi:</label>
              <input type="text" name="accreditation" class="form-control" value="${escapeHtml(school.accreditation || 'Swasta / Akreditasi B')}">
            </div>
          </div>

          <div style="background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 6px; padding: 10px 14px; margin: 10px 0;">
            <strong style="font-size: 12px; color: #166534;">🎓 Kompetensi Keahlian (Jurusan Resmi SMKS IT Al-Hawari):</strong>
            <ul style="margin: 4px 0 0 18px; font-size: 12px; color: #15803d; line-height: 1.5;">
              <li><strong>Multimedia (MM)</strong> — Desain Grafis, Animasi, Fotografi, Videografi, Media Digital.</li>
              <li><strong>Agribisnis Tanaman Pangan dan Hortikultura (ATPH)</strong> — Budi Daya Pangan, Sayuran, Hortikultura Modern.</li>
              <li><strong>Desain Pemodelan dan Informasi Bangunan (DPIB)</strong> — Gambar Konstruksi, Arsitektur, RAB Bangunan.</li>
            </ul>
          </div>

          <h4 style="margin: 20px 0 10px 0; border-bottom: 1px solid #e2e8f0; padding-bottom: 6px;">2. Alamat & Kontak</h4>
          <div class="form-group">
            <label>Alamat Lengkap:</label>
            <input type="text" name="address" class="form-control" value="${escapeHtml(school.address || 'Kp. Curug RT 02 / RW 06, Desa Cipareuan, Kec. Cibiuk')}">
          </div>
          <div class="form-row">
            <div class="form-group col">
              <label>Kota / Kabupaten:</label>
              <input type="text" name="city" class="form-control" value="${escapeHtml(school.city || 'Garut')}">
            </div>
            <div class="form-group col">
              <label>Kode Pos:</label>
              <input type="text" name="postalCode" class="form-control" value="${escapeHtml(school.postalCode || '44193')}">
            </div>
          </div>
          <div class="form-row">
            <div class="form-group col">
              <label>Nomor Telepon / WhatsApp:</label>
              <input type="text" name="phone" class="form-control" value="${escapeHtml(school.phone || '+62 895-3266-12155')}">
            </div>
            <div class="form-group col">
              <label>Email Resmi:</label>
              <input type="email" name="email" class="form-control" value="${escapeHtml(school.email || 'smkitalhawari@gmail.com')}">
            </div>
            <div class="form-group col">
              <label>Website:</label>
              <input type="text" name="website" class="form-control" value="${escapeHtml(school.website || 'al-hawari.sch.id')}">
            </div>
          </div>

          <h4 style="margin: 20px 0 10px 0; border-bottom: 1px solid #e2e8f0; padding-bottom: 6px;">3. Pejabat Penandatangan (Kepala Sekolah)</h4>
          <div class="form-row">
            <div class="form-group col">
              <label>Nama Lengkap & Gelar:</label>
              <input type="text" name="principalName" class="form-control" value="${escapeHtml(school.principalName || 'Dian Nurhadian, S.Pd.')}">
            </div>
            <div class="form-group col">
              <label>NIP (Isi '-' jika yayasan/swasta):</label>
              <input type="text" name="principalNip" class="form-control" value="${escapeHtml(school.principalNip || '-')}">
            </div>
          </div>
          <div class="form-row">
            <div class="form-group col">
              <label>Pangkat / Golongan / Keterangan:</label>
              <input type="text" name="principalRank" class="form-control" value="${escapeHtml(school.principalRank || 'Kepala SMKS IT Al-Hawari')}">
            </div>
            <div class="form-group col">
              <label>Jabatan:</label>
              <input type="text" name="principalTitle" class="form-control" value="${escapeHtml(school.principalTitle || 'Kepala Sekolah')}">
            </div>
          </div>

          <h4 style="margin: 20px 0 10px 0; border-bottom: 1px solid #e2e8f0; padding-bottom: 6px;">4. Pola Format Penomoran & Kode Jenis Surat</h4>
          <div class="form-group">
            <label>Pola Format Nomor Surat:</label>
            <input type="text" name="numberingPattern" class="form-control" value="${escapeHtml(school.numberingPattern || '421.3/{urut}/SMKITA/{bulanRomawi}/{tahun}')}">
            <small class="text-muted">Gunakan placeholder: <code>{urut}</code>, <code>{kode}</code>, <code>{jenis}</code>, <code>{bulanRomawi}</code>, <code>{tahun}</code></small>
            <div style="margin-top: 8px; background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 6px; padding: 10px 12px; font-size: 11.5px; color: #1e3a8a; line-height: 1.5;">
              <strong>📌 Kode Jenis Surat Otomatis (AI & Kanvas):</strong><br>
              • <code>UND</code> (Surat Undangan: <i>050/UND.SMKITA/...</i>) &nbsp;|&nbsp;
              • <code>TU</code> (Tata Usaha: <i>050/TU.SMKITA/...</i>) &nbsp;|&nbsp;
              • <code>KEP</code> (Kepala Sekolah / SK: <i>050/KEP.SMKITA/...</i>)<br>
              • <code>KET</code> (Surat Keterangan) &nbsp;|&nbsp;
              • <code>ST</code> (Surat Tugas) &nbsp;|&nbsp;
              • <code>PGL</code> (Panggilan Ortu/Siswa) &nbsp;|&nbsp;
              • <code>PBR</code> (Pemberitahuan/Edaran) &nbsp;|&nbsp;
              • <code>PRM</code> (Permohonan) &nbsp;|&nbsp;
              • <code>DISP</code> (Dispensasi) &nbsp;|&nbsp;
              • <code>BA</code> (Berita Acara) &nbsp;|&nbsp;
              • <code>PKL</code> (Prakerin/Hubin) &nbsp;|&nbsp;
              • <code>KUR</code> (Kurikulum) &nbsp;|&nbsp;
              • <code>KES</code> (Kesiswaan)
            </div>
          </div>

          <div class="btn-group" style="margin-top: 20px;">
            <button type="submit" class="btn btn-primary">💾 Simpan Profil Sekolah</button>
          </div>
        </form>
      </div>
    </div>
  `;

  // Bind pilih kop dari galeri
  container.querySelectorAll('.btn-apply-school-kop').forEach(btn => {
    btn.addEventListener('click', () => {
      const id = btn.dataset.id;
      const kop = store.getKopGallery().find(k => k.id === id);
      if (kop) {
        store.updateSchoolProfile({ kopImage: kop.dataUrl });
        showToast(`Kop aktif diatur ke: ${kop.name}`, 'success');
        renderSchoolProfileView(container);
      }
    });
  });

  // Form submit handling
  const form = container.querySelector('#form-school-profile');
  form.onsubmit = (e) => {
    e.preventDefault();
    const formData = new FormData(form);
    const updated = {};
    for (const [key, val] of formData.entries()) {
      updated[key] = val.trim();
    }
    store.updateSchoolProfile(updated);
    renderSchoolProfileView(container);
    showToast('Profil sekolah dan kop surat berhasil disimpan permanen.', 'success');
  };
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str || '';
  return div.innerHTML;
}
