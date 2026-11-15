/**
 * suratKeterangan.js - Template Surat Keterangan Siswa Aktif
 * Format resmi administrasi sekolah untuk keperluan beasiswa, tunjangan orang tua, lomba, dll.
 */

export const suratKeteranganTemplate = {
  id: 'surat_keterangan',
  name: 'Surat Keterangan Siswa Aktif',
  category: 'Kesiswaan',
  description: 'Menerangkan bahwa siswa yang bersangkutan adalah peserta didik aktif dan terdaftar di sekolah.',
  
  // Skema Form Field
  fields: [
    {
      name: 'nomorSurat',
      label: 'Nomor Surat',
      type: 'text',
      placeholder: 'Contoh: 421.3/001/KET.SMKITA/XI/2026',
      defaultValue: '421.3/001/KET.SMKITA/XI/2026',
      required: true
    },
    {
      name: 'tanggalSurat',
      label: 'Tanggal Surat',
      type: 'text',
      placeholder: 'Contoh: 14 November 2026',
      defaultValue: '14 November 2026',
      required: true
    },
    {
      name: 'namaSiswa',
      label: 'Nama Lengkap Siswa',
      type: 'text',
      placeholder: 'Nama siswa (huruf kapital atau title case)',
      defaultValue: 'Ahmad Fauzi',
      required: true
    },
    {
      name: 'nisNisn',
      label: 'NIS / NISN',
      type: 'text',
      placeholder: 'Contoh: 12345 / 0054321987',
      defaultValue: '12345 / 0054321987',
      required: true
    },
    {
      name: 'kelasJurusan',
      label: 'Kelas / Program Keahlian',
      type: 'text',
      placeholder: 'Contoh: XII MM 1 / Multimedia',
      defaultValue: 'XII MM 1 / Multimedia',
      required: true
    },
    {
      name: 'ttl',
      label: 'Tempat, Tanggal Lahir',
      type: 'text',
      placeholder: 'Contoh: Garut, 12 Mei 2008',
      defaultValue: 'Garut, 12 Mei 2008',
      required: true
    },
    {
      name: 'jenisKelamin',
      label: 'Jenis Kelamin',
      type: 'select',
      options: ['Laki-laki', 'Perempuan'],
      defaultValue: 'Laki-laki',
      required: true
    },
    {
      name: 'namaOrtu',
      label: 'Nama Orang Tua / Wali',
      type: 'text',
      placeholder: 'Contoh: Bambang Sudarsono',
      defaultValue: 'Bambang Sudarsono',
      required: false
    },
    {
      name: 'pekerjaanOrtu',
      label: 'Pekerjaan Orang Tua',
      type: 'text',
      placeholder: 'Contoh: Wiraswasta / Karyawan / PNS',
      defaultValue: 'Wiraswasta',
      required: false
    },
    {
      name: 'alamatSiswa',
      label: 'Alamat Tinggal Siswa',
      type: 'text',
      placeholder: 'Alamat domisili lengkap',
      defaultValue: 'Kp. Curug RT 02 / RW 06, Desa Cipareuan, Kec. Cibiuk, Garut',
      required: false
    },
    {
      name: 'tahunAjaran',
      label: 'Tahun Ajaran Aktif',
      type: 'text',
      placeholder: 'Contoh: 2026/2027',
      defaultValue: '2026/2027',
      required: true
    },
    {
      name: 'keperluan',
      label: 'Keperluan / Maksud Surat',
      type: 'textarea',
      placeholder: 'Maksud penggunaan surat keterangan...',
      defaultValue: 'Kelengkapan berkas pendaftaran Program Beasiswa Prestasi dan Tunjangan Pendidikan.',
      required: true
    }
  ],

  /**
   * Menghasilkan Document Model JSON lengkap dari data form dan profil sekolah
   * @param {object} formData 
   * @param {object} schoolProfile 
   * @param {object} [options]
   * @returns {object} Document Model JSON
   */
  generateDocumentModel(formData, schoolProfile, options = {}) {
    const profileKey = options.paperSize === 'f4' ? 'suratDinasF4' : 'suratDinas';

    return {
      meta: {
        title: `Surat Keterangan Siswa Aktif - ${formData.namaSiswa || 'Siswa'}`,
        author: schoolProfile.schoolName || 'Tata Usaha Sekolah',
        year: new Date().getFullYear().toString()
      },
      profile: profileKey,
      letterData: {
        templateId: 'surat_keterangan',
        schoolProfile,
        formData,
        options
      },
      blocks: [
        // 1. KOP SURAT BERGARIS GANDA
        {
          type: 'letterKop',
          school: schoolProfile
        },

        // 2. JUDUL DAN NOMOR SURAT (Beserta Sifat, Perihal & Kepada Yth Default)
        {
          type: 'letterTitle',
          title: 'SURAT KETERANGAN SISWA AKTIF',
          nomor: `Nomor: ${formData.nomorSurat || '-'}`,
          metaRecipient: {
            metaRows: [
              { label: 'Sifat', value: 'Penting' },
              { label: 'Perihal', value: formData.perihal || 'Undangan Rapat' }
            ],
            recipientLines: [
              'Kepada',
              'Yth. ...............',
              '__________________',
              'Di',
              'Tempat'
            ]
          }
        },

        // 3. PARAGRAF PEMBUKA
        {
          type: 'paragraph',
          style: 'letterBody',
          runs: [
            { text: `Yang bertanda tangan di bawah ini, ${schoolProfile.principalTitle || 'Kepala'} ${schoolProfile.schoolName || 'Sekolah'}, dengan ini menerangkan bahwa:` }
          ]
        },

        // 4. IDENTITAS SISWA (Tabel tanpa border)
        {
          type: 'identityTable',
          rows: [
            { label: 'Nama Lengkap', value: `: ${formData.namaSiswa || '-'}` },
            { label: 'NIS / NISN', value: `: ${formData.nisNisn || '-'}` },
            { label: 'Kelas / Program Keahlian', value: `: ${formData.kelasJurusan || '-'}` },
            { label: 'Tempat, Tanggal Lahir', value: `: ${formData.ttl || '-'}` },
            { label: 'Jenis Kelamin', value: `: ${formData.jenisKelamin || '-'}` },
            ...(formData.namaOrtu ? [
              { label: 'Nama Orang Tua / Wali', value: `: ${formData.namaOrtu}` },
              { label: 'Pekerjaan Orang Tua', value: `: ${formData.pekerjaanOrtu || '-'}` }
            ] : []),
            ...(formData.alamatSiswa ? [
              { label: 'Alamat Tinggal', value: `: ${formData.alamatSiswa}` }
            ] : [])
          ]
        },

        // 5. PERNYATAAN STATUS SISWA
        {
          type: 'paragraph',
          style: 'letterBody',
          runs: [
            { text: `Adalah benar-benar peserta didik yang tercatat aktif terdaftar pada ${schoolProfile.schoolName || 'sekolah ini'} pada Tahun Ajaran ${formData.tahunAjaran || '2026/2027'}, serta senantiasa menaati seluruh tata tertib dan berkelakuan baik.` }
          ]
        },

        // 6. MAKSUD / KEPERLUAN SURAT
        {
          type: 'paragraph',
          style: 'letterBody',
          runs: [
            { text: `Surat keterangan ini diterbitkan dan diberikan kepada yang bersangkutan untuk keperluan: ` },
            { text: formData.keperluan || 'Administrasi sekolah.', bold: true }
          ]
        },

        // 7. PENUTUP
        {
          type: 'paragraph',
          style: 'letterBody',
          runs: [
            { text: 'Demikian surat keterangan ini kami buat dengan sebenarnya agar dapat dipergunakan sebagaimana mestinya.' }
          ]
        },

        // 8. BLOK TANDA TANGAN KEPALA SEKOLAH
        {
          type: 'letterSignature',
          city: schoolProfile.city || 'Garut',
          date: formData.tanggalSurat || '14 November 2026',
          title: schoolProfile.principalTitle || 'Kepala Sekolah',
          name: schoolProfile.principalName || 'Dian Nurhadian, S.Pd.',
          nip: schoolProfile.principalNip || '-',
          rank: schoolProfile.principalRank || ''
        }
      ]
    };
  }
};

