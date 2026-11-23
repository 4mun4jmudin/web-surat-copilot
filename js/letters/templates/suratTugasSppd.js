/**
 * suratTugasSppd.js - Template Resmi Surat Tugas & SPPD (Surat Perintah Perjalanan Dinas) 3 Halaman
 * Mengikuti tata letak persis dokumen Microsoft Word sekolah (Halaman 1 Surat Tugas + Tabel Pegawai,
 * Halaman 2 SPPD 8 Poin + Catatan Hukum & TTD 2 Kolom, Halaman 3 Keterangan I & II dengan Header Merge).
 */

import { DEFAULT_SCHOOL_PROFILE, getSchoolCodeFromProfile, normalizeSchoolProfileNaming } from '../schoolProfile.js';
import { formatSmartNomorSurat } from '../numbering.js';

function getTodayFormattedId() {
  try {
    return new Date().toLocaleDateString('id-ID', {
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    });
  } catch {
    return '18 November 2026';
  }
}

export function normalizeAssigneesList(raw = null) {
  if (Array.isArray(raw) && raw.length > 0) {
    return raw
      .map(item => ({
        nama: String(item?.nama || item?.name || '').trim(),
        jabatan: String(item?.jabatan || item?.role || 'Guru').trim()
      }))
      .filter(item => item.nama);
  }
  if (typeof raw === 'string' && raw.trim()) {
    return raw
      .split(/\r?\n/)
      .map(line => line.trim())
      .filter(Boolean)
      .map(line => {
        const parts = line.split(/\||-|—/);
        return {
          nama: (parts[0] || '').trim(),
          jabatan: (parts[1] || 'Guru').trim()
        };
      })
      .filter(item => item.nama);
  }
  return [
    { nama: 'Taufik Dwi Setyo Ramadhan, S.Pd.', jabatan: 'Guru' },
    { nama: 'Siti Laras Maemunah, S.Pd.', jabatan: 'Guru' }
  ];
}

export function buildSuratTugasBlocks(data = {}, school = {}) {
  const s = normalizeSchoolProfileNaming({ ...DEFAULT_SCHOOL_PROFILE, ...school });
  const schoolCode = getSchoolCodeFromProfile(s);
  const schoolDisplay = s.unitLabel || (s.unitId === 'smp' ? 'SMP IT Al-Hawari - Cibiuk' : 'SMK IT Al-Hawari');
  const unitKerja = s.schoolName || 'SMK IT AL-HAWARI';
  const pemberiNama = data.pemberiNama || s.principalName || 'Dian Nurhamdani, S.Pd.';
  const pemberiJabatan = data.pemberiJabatan || 'Kepala Sekolah';
  const alamatSekolah = data.alamatSekolah || `${s.address || 'Kp. Curug RT/RW 002/006 Desa Cipareuan Kec. Cibiuk'} Kab. ${s.city || 'Garut'} ${s.postalCode || '44193'}`;
  const nomorSurat = data.nomorSurat
    ? formatSmartNomorSurat({ rawNomor: data.nomorSurat, forcedTypeCode: 'ST', contextText: 'Surat Tugas', schoolCode })
    : formatSmartNomorSurat({ forcedTypeCode: 'ST', contextText: 'Surat Tugas', sequenceNumber: data.sequenceNumber || 1, schoolCode });
  const assignees = normalizeAssigneesList(data.assignees || data.pegawai);
  const maksudTugas = data.maksudTugas || 'Penjemputan Peserta Didik PKL';
  const tujuanLokasi = data.tujuanLokasi || 'DRAGON MOTION PICTURE ASIA';
  const tempatTiba = data.tempatTiba || data.tujuanLokasi || 'Dragon Motion Picture';
  const tanggalSurat = data.tanggalSurat || getTodayFormattedId();
  const city = data.city || s.city || 'Garut';

  const assigneeTableRows = [
    [
      { text: 'No', align: 'center' },
      { text: 'Nama', align: 'center' },
      { text: 'Jabatan', align: 'center' }
    ],
    ...assignees.map((p, idx) => [
      { text: String(idx + 1), align: 'left' },
      { text: p.nama, align: 'center' },
      { text: p.jabatan || 'Guru', align: 'center' }
    ])
  ];

  return [
    {
      type: 'letterTitle',
      title: 'SURAT TUGAS',
      nomor: `Nomor: ${nomorSurat}`
    },
    {
      type: 'paragraph',
      style: 'letterBody',
      alignment: 'left',
      runs: [{ text: 'Yang bertanda tangan dibawah ini:' }]
    },
    {
      type: 'identityTable',
      rows: [
        { label: 'Nama', value: `: ${pemberiNama}` },
        { label: 'Jabatan', value: `: ${pemberiJabatan}` },
        { label: 'Unit Kerja', value: `: ${unitKerja}` },
        { label: 'Alamat', value: `: ${alamatSekolah}` }
      ]
    },
    {
      type: 'paragraph',
      style: 'letterBody',
      alignment: 'left',
      runs: [{ text: `Memberikan tugas kepada Guru ${schoolDisplay}:` }]
    },
    {
      type: 'table',
      borderStyle: 'grid',
      headerRow: true,
      headerShading: false,
      headerBold: false,
      colWidthsPct: [9, 51, 40],
      rows: assigneeTableRows
    },
    {
      type: 'paragraph',
      style: 'letterBody',
      alignment: 'left',
      runs: [{ text: `Untuk ${maksudTugas} Bertempat di ${tujuanLokasi}` }]
    },
    {
      type: 'paragraph',
      style: 'letterBody',
      alignment: 'left',
      runs: [{ text: 'Demikian surat ini kami buat agar digunakan sebagai mestinya.' }]
    },
    {
      type: 'letterSignature',
      city,
      date: tanggalSurat,
      title: `Kepala ${schoolDisplay}`,
      name: pemberiNama,
      nip: s.principalNip || '-'
    },
    {
      type: 'paragraph',
      style: 'letterBody',
      alignment: 'left',
      spacingBefore: 120,
      spacingAfter: 360,
      runs: [
        { text: 'Tiba di ' },
        { text: tempatTiba, underline: true }
      ]
    },
    {
      type: 'paragraph',
      style: 'letterBody',
      alignment: 'left',
      runs: [{ text: '_______________________' }]
    }
  ];
}

export function buildSppdBlocks(data = {}, school = {}) {
  const s = normalizeSchoolProfileNaming({ ...DEFAULT_SCHOOL_PROFILE, ...school });
  const schoolDisplay = s.unitLabel || (s.unitId === 'smp' ? 'SMP IT Al-Hawari - Cibiuk' : 'SMK IT Al-Hawari');
  const unitKerja = s.schoolName || 'SMK IT AL-HAWARI';
  const pemberiNama = data.pemberiNama || s.principalName || 'Dian Nurhamdani, S.Pd.';
  const assignees = normalizeAssigneesList(data.assignees || data.pegawai);
  const tujuanLokasi = data.tujuanLokasi || 'DRAGON MOTION PICTURE';
  const maksudSppd = data.maksudSppd || data.maksudTugas || 'Penjemputan siswa kelas 12 Multimedia ke tempat PKL';
  const kendaraan = data.kendaraan || 'Mobil';
  const kendaraanSingkat = data.kendaraanSingkat || (kendaraan.toLowerCase().includes('motor') ? 'Roda 2' : 'Roda 4');
  const lamaHari = String(data.lamaHari || '1').trim();
  const tanggalSurat = data.tanggalSurat || getTodayFormattedId();
  const tglBerangkat = data.tglBerangkat || tanggalSurat;
  const tglKembali = data.tglKembali || tglBerangkat;
  const biaya = data.biaya !== undefined ? String(data.biaya) : '';
  const atasBeban = data.atasBeban || '.........................................................';
  const pasalAnggaran = data.pasalAnggaran || '.........................................................';

  const pegawaiListText = assignees.map((p, idx) => `${idx + 1}.  ${p.nama}`).join('\n');
  const jabatanListText = assignees.map((p, idx) => `${idx + 1}.  ${p.jabatan || 'Guru'}`).join('\n');

  const sppdFrontTable = {
    type: 'table',
    borderStyle: 'sppd',
    headerRow: false,
    headerShading: false,
    headerBold: false,
    colWidthsPct: [46, 54],
    rows: [
      [
        { text: '1.   Pejabat yang memberi perintah' },
        { text: `Kepala ${schoolDisplay}\nKec. Cibiuk Kab. Garut` }
      ],
      [
        { text: '2.   Nama Pegawai yang diperintahkan\n      Mengadakan perjalanan.' },
        { text: pegawaiListText }
      ],
      [
        { text: '3.   Jabatan, Pangkat dari pegawai yang\n      diperintahkan' },
        { text: jabatanListText }
      ],
      [
        { text: '4.   Perjalanan dinas yang diperintahkan' },
        { text: `Dari   : ${unitKerja}\nKe     : ${tujuanLokasi}\nDengan menggunakan kendaraan: ${kendaraan}` }
      ],
      [
        { text: '5.   Perjalanan yang direncanakan' },
        {
          text: `A.  Selama   ${lamaHari}   hari\n     Dari tanggal   ${tglBerangkat}\n     s.d. tanggal    ${tglKembali}.\nB.  Dengan lumpsum biaya Rp.${biaya ? ` ${biaya}` : ''}`
        }
      ],
      [
        { text: '6.   Maksud mengadakan perjalanan' },
        { text: maksudSppd }
      ],
      [
        { text: '7.   Perhitungan biaya perjalanan' },
        { text: `Atas beban     : ${atasBeban}\nPasal anggaran : ${pasalAnggaran}` }
      ],
      [
        { text: '8.   Keterangan' },
        { text: 'Lihat sebelah' }
      ]
    ]
  };

  const sppdFooterLayoutTable = {
    type: 'table',
    borderStyle: 'none',
    headerRow: false,
    headerShading: false,
    headerBold: false,
    colWidthsPct: [55, 45],
    rows: [
      [
        {
          paragraphs: [
            { text: 'CATATAN:', italic: true, bold: true, spacingAfter: 80 },
            {
              text: 'Pemegang kuasa yang berhak memberi perintah jalan diperingatkan dengan tegas bahwa perjalanan dinas hanya boleh dilakukan atas biaya-biaya Negara. Jika ini diperlukan untuk kepentingan Negara (lihat pasal 2 dan pasal 3 peraturan Perjalanan Dinas dalam Negeri; buat tahun 1995). Jika Negara menderita kerugian yang disebabkan karena pemberian perintah jalan yang tidak beralasan maka kerugian itu akan dibebankan pada pemberi perintah jalan itu (pasal 11 Peraturan Pemerintah No.331/1995 jo. Pasal 50 Peraturan Pemerintah No.5/1975).',
              fontSizeHalfPt: 20,
              align: 'both'
            }
          ]
        },
        {
          paragraphs: [
            { text: tanggalSurat, align: 'center' },
            { text: `Kepala ${schoolDisplay}`, align: 'center' },
            { text: 'Kec. Cibiuk Kab. Garut', align: 'center', spacingAfter: 40 },
            { text: '', align: 'center', qrSlot: true, qrCaption: 'Disahkan secara elektronik', spacingAfter: 20 },
            { text: '', spacingAfter: 360 },
            { text: pemberiNama, align: 'center', bold: true, underline: true }
          ]
        }
      ]
    ]
  };

  const keteranganTabel1 = {
    type: 'table',
    borderStyle: 'grid',
    headerRow: true,
    headerShading: false,
    headerBold: false,
    colWidthsPct: [34, 14, 15, 11, 13, 13],
    rows: [
      {
        isHeader: true,
        cells: [
          { text: 'Tempat kedudukan pegawai yang\ndiberi perintah jalan', rowspan: 2, align: 'center', vAlign: 'center' },
          { text: 'Berangkat', colspan: 2, align: 'center', vAlign: 'center' },
          { text: 'Tanda\nTangan', rowspan: 2, align: 'center', vAlign: 'center' },
          { text: 'Tiba Kembali', colspan: 2, align: 'center', vAlign: 'center' }
        ]
      },
      {
        isHeader: true,
        cells: [
          { text: 'Tanggal', align: 'center', vAlign: 'center' },
          { text: 'Dengan\nKendaraan', align: 'center', vAlign: 'center' },
          { text: 'Tanggal', align: 'center', vAlign: 'center' },
          { text: 'Tanda\nTangan', align: 'center', vAlign: 'center' }
        ]
      },
      {
        isHeader: false,
        minHeightCm: 4.5,
        cells: [
          { text: schoolDisplay, align: 'left', vAlign: 'center' },
          { text: `${tglBerangkat}.`, align: 'left', vAlign: 'center' },
          { text: kendaraanSingkat, align: 'center', vAlign: 'center' },
          { text: '', align: 'center' },
          { text: `${tglKembali}.`, align: 'left', vAlign: 'center' },
          { text: '', align: 'center' }
        ]
      }
    ]
  };

  const keteranganTabel2 = {
    type: 'table',
    borderStyle: 'grid',
    headerRow: true,
    headerShading: false,
    headerBold: false,
    colWidthsPct: [32, 14, 12, 14, 15, 13],
    rows: [
      {
        isHeader: true,
        cells: [
          { text: 'Kantor yang dikunjungi', rowspan: 2, align: 'center', vAlign: 'center' },
          { text: 'Datang', colspan: 2, align: 'center', vAlign: 'center' },
          { text: 'Berangkat lagi', colspan: 3, align: 'center', vAlign: 'center' }
        ]
      },
      {
        isHeader: true,
        cells: [
          { text: 'Tanggal', align: 'center', vAlign: 'center' },
          { text: 'Tanda\nTangan', align: 'center', vAlign: 'center' },
          { text: 'Tanggal', align: 'center', vAlign: 'center' },
          { text: 'Dengan\nKendaraan', align: 'center', vAlign: 'center' },
          { text: 'Tanda\nTangan', align: 'center', vAlign: 'center' }
        ]
      },
      {
        isHeader: false,
        minHeightCm: 4.8,
        cells: [
          { text: tujuanLokasi, bold: true, align: 'left', vAlign: 'center' },
          { text: `${tglBerangkat}.`, align: 'left', vAlign: 'center' },
          { text: '', align: 'center' },
          { text: `${tglKembali}.`, align: 'left', vAlign: 'center' },
          { text: kendaraanSingkat, align: 'center', vAlign: 'center' },
          { text: '', align: 'center' }
        ]
      }
    ]
  };

  return [
    {
      type: 'paragraph',
      style: 'letterBody',
      alignment: 'center',
      spacingAfter: 20,
      runs: [{ text: 'Surat Perintah Perjalanan Dinas', bold: true, fontSizeHalfPt: 26 }]
    },
    {
      type: 'paragraph',
      style: 'letterBody',
      alignment: 'center',
      spacingAfter: 160,
      runs: [{ text: '( S P P D )', bold: true, fontSizeHalfPt: 26 }]
    },
    sppdFrontTable,
    sppdFooterLayoutTable,
    { type: 'pageBreak' },
    {
      type: 'paragraph',
      style: 'letterBody',
      alignment: 'center',
      spacingAfter: 140,
      runs: [{ text: 'K e t e r a n g a n :', bold: true }]
    },
    {
      type: 'paragraph',
      style: 'letterBody',
      alignment: 'left',
      spacingAfter: 80,
      runs: [{ text: 'I.   DARI PEJABAT PEMBERI PERINTAH JALAN' }]
    },
    keteranganTabel1,
    {
      type: 'paragraph',
      style: 'letterBody',
      alignment: 'left',
      spacingBefore: 240,
      spacingAfter: 80,
      runs: [{ text: 'II.  DARI TEMPAT YANG DIKUNJUNGI' }]
    },
    keteranganTabel2
  ];
}

export function buildSuratTugasSppdDocModel(formData = {}, schoolProfile = {}, options = {}) {
  const s = normalizeSchoolProfileNaming({ ...DEFAULT_SCHOOL_PROFILE, ...schoolProfile });
  const includeSuratTugas = options.includeSuratTugas !== false;
  const includeSppd = options.includeSppd !== false;
  const showKop = options.showKop !== false;
  const paperSize = options.paperSize || 'f4';
  const orientation = options.orientation || 'portrait';
  const margins = options.margins || { top: 2.0, bottom: 2.0, left: 2.5, right: 2.0 };

  const blocks = [];
  if (showKop && includeSuratTugas) {
    blocks.push({
      type: 'letterKop',
      school: s,
      image: s.kopImage || ''
    });
  }

  if (includeSuratTugas) {
    blocks.push(...buildSuratTugasBlocks(formData, s));
  }

  if (includeSuratTugas && includeSppd) {
    blocks.push({ type: 'pageBreak' });
  }

  if (includeSppd) {
    blocks.push(...buildSppdBlocks(formData, s));
  }

  const titleLabel = includeSuratTugas && includeSppd
    ? 'SURAT TUGAS & SPPD'
    : (includeSppd ? 'SURAT PERINTAH PERJALANAN DINAS (SPPD)' : 'SURAT TUGAS');

  return {
    meta: {
      title: `${titleLabel}${formData.tujuanLokasi ? ` - ${formData.tujuanLokasi}` : ''}`,
      author: s.schoolName || 'SMK IT Al-Hawari',
      year: new Date().getFullYear().toString()
    },
    profile: paperSize === 'f4' ? 'suratDinasF4' : 'suratDinas',
    paperSize,
    orientation,
    customMargins: margins,
    letterData: {
      templateId: includeSuratTugas && includeSppd ? 'surat_tugas_sppd' : (includeSppd ? 'sppd_only' : 'surat_tugas'),
      schoolProfile: s,
      formData,
      options
    },
    blocks
  };
}
