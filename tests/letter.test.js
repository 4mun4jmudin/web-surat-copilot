import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DEFAULT_SCHOOL_PROFILE } from '../js/letters/schoolProfile.js';
import { suratKeteranganTemplate } from '../js/letters/templates/suratKeterangan.js';
import { validateDocument } from '../js/model/schema.js';
import { renderDocx, getProfile } from '../js/docx/render.js';
import { readZipPackage } from '../js/utils/zip.js';

test('Surat Dinas: Profil suratDinas (A4) dan suratDinasF4 (Folio) memiliki dimensi baku', () => {
  const a4 = getProfile('suratDinas');
  const f4 = getProfile('suratDinasF4');

  assert.equal(a4.page.width, 11906, 'A4 lebar harus 210mm');
  assert.equal(a4.page.height, 16838, 'A4 tinggi harus 297mm');
  assert.equal(f4.page.width, 12189, 'F4 lebar harus 215mm');
  assert.equal(f4.page.height, 18709, 'F4 tinggi harus 330mm');
});

test('Template Surat Keterangan: Menghasilkan Document Model yang valid', () => {
  const sampleFormData = {
    nomorSurat: '421.3 / 100 / SMKITA / XI / 2026',
    tanggalSurat: '14 November 2026',
    namaSiswa: 'Budi Santoso',
    nisNisn: '11223 / 0071234567',
    kelasJurusan: 'XII RPL 1 / Rekayasa Perangkat Lunak',
    ttl: 'Surabaya, 10 Januari 2008',
    jenisKelamin: 'Laki-laki',
    namaOrtu: 'Sutrisno',
    pekerjaanOrtu: 'Wiraswasta',
    alamatSiswa: 'Jl. Rungkut Asri No. 5, Surabaya',
    tahunAjaran: '2026/2027',
    keperluan: 'Pendaftaran beasiswa perguruan tinggi'
  };

  const docModel = suratKeteranganTemplate.generateDocumentModel(sampleFormData, DEFAULT_SCHOOL_PROFILE);
  const validation = validateDocument(docModel);

  assert.equal(validation.valid, true, `Seharusnya valid, error: ${validation.errors.join('; ')}`);
  assert.equal(docModel.blocks[0].type, 'letterKop');
  assert.equal(docModel.blocks[1].type, 'letterTitle');
  assert.equal(docModel.blocks[3].type, 'identityTable');
  assert.equal(docModel.blocks[7].type, 'letterSignature');
});

test('Renderer DOCX: Merender Surat Dinas lengkap dengan Kop, Garis Ganda, dan Logo ter-embed', async () => {
  const sampleFormData = {
    nomorSurat: '421.3 / 100 / SMKITA / XI / 2026',
    tanggalSurat: '14 November 2026',
    namaSiswa: 'Budi Santoso',
    nisNisn: '11223 / 0071234567',
    kelasJurusan: 'XII RPL 1',
    ttl: 'Surabaya, 10 Januari 2008',
    jenisKelamin: 'Laki-laki',
    tahunAjaran: '2026/2027',
    keperluan: 'Kelengkapan administrasi'
  };

  const docModel = suratKeteranganTemplate.generateDocumentModel(sampleFormData, DEFAULT_SCHOOL_PROFILE, { paperSize: 'a4' });
  const docxBytes = await renderDocx(docModel, 'uint8array');

  assert.ok(docxBytes instanceof Uint8Array);
  assert.ok(docxBytes.length > 500);

  const zip = await readZipPackage(docxBytes);
  
  // Pastikan logo/kop di-embed ke word/media/
  const logoFile = zip.file('word/media/logo.png') || zip.file('word/media/kop.png');
  assert.ok(logoFile, 'Berkas word/media/logo.png atau kop.png harus ada di dalam docx');

  // Periksa konten XML
  const documentXml = await zip.file('word/document.xml').async('text');
  assert.match(documentXml, /PEMERINTAH DAERAH PROVINSI JAWA BARAT/, 'Harus ada teks kop pemerintah');
  assert.match(documentXml, /SMKS IT AL-HAWARI/, 'Harus ada nama sekolah');
  assert.match(documentXml, /w:val="double"/, 'Harus ada garis ganda kop surat');
  assert.match(documentXml, /SURAT KETERANGAN SISWA AKTIF/, 'Harus ada judul surat');
  assert.match(documentXml, /Budi Santoso/, 'Harus ada nama siswa');
  assert.match(documentXml, /Dian Nurhadian, S\.Pd\./, 'Harus ada nama kepala sekolah');
});

test('Profil Sekolah Bawaan: SMKS IT Al-Hawari memiliki identitas lengkap dan akurat', () => {
  assert.equal(DEFAULT_SCHOOL_PROFILE.schoolName, 'SMKS IT AL-HAWARI');
  assert.equal(DEFAULT_SCHOOL_PROFILE.npsn, '69758993');
  assert.equal(DEFAULT_SCHOOL_PROFILE.principalName, 'Dian Nurhadian, S.Pd.');
  assert.equal(DEFAULT_SCHOOL_PROFILE.foundation, 'Yayasan Nunuk Kartini Al-Irsyad');
  assert.equal(DEFAULT_SCHOOL_PROFILE.city, 'Garut');
  assert.equal(DEFAULT_SCHOOL_PROFILE.postalCode, '44193');
  assert.equal(DEFAULT_SCHOOL_PROFILE.email, 'smkitalhawari@gmail.com');
  assert.equal(DEFAULT_SCHOOL_PROFILE.website, 'al-hawari.sch.id');
  assert.ok(DEFAULT_SCHOOL_PROFILE.majors.includes('Multimedia (MM)'));
  assert.ok(DEFAULT_SCHOOL_PROFILE.majors.includes('Agribisnis Tanaman Pangan dan Hortikultura (ATPH)'));
  assert.ok(DEFAULT_SCHOOL_PROFILE.majors.includes('Desain Pemodelan dan Informasi Bangunan (DPIB)'));
  assert.equal(DEFAULT_SCHOOL_PROFILE.numberingPattern, '421.3/{urut}/SMKITA/{bulanRomawi}/{tahun}');
});
