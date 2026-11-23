/**
 * schoolProfile.js - Pengelola Data Profil Sekolah & Kop Surat Multi-Unit (Yayasan Al-Irsyad / Al-Hawari)
 * Menyimpan definisi resmi unit SMP IT Al-Hawari - Cibiuk & SMK IT Al-Hawari serta helper isolasi & adaptasi unit.
 */

import { DEFAULT_KOP_DATA_URL, DEFAULT_SMP_KOP_DATA_URL } from './defaultKop.js';
import { formatSmartNomorSurat } from './numbering.js';

export const VALID_UNIT_IDS = Object.freeze(['smk', 'smp']);

export const SCHOOL_UNITS = Object.freeze({
  smk: Object.freeze({
    unitId: 'smk',
    unitLabel: 'SMK IT Al-Hawari',
    unitShortBadge: 'SMK',
    schoolCode: 'SMKITA',
    government: 'PEMERINTAH DAERAH PROVINSI JAWA BARAT',
    department: 'DINAS PENDIDIKAN',
    subOffice: 'CABANG DINAS PENDIDIKAN WILAYAH XI',
    foundation: 'Yayasan Nunuk Kartini Al-Irsyad',
    npyp: 'AA2743',
    schoolName: 'SMK IT AL-HAWARI',
    npsn: '69758993',
    skOperasional: '425.11/1136-DISDIK (30 April 2013)',
    accreditation: 'Swasta / Akreditasi B',
    educationLevel: 'SMK (Sekolah Menengah Kejuruan)',
    address: 'Kp. Curug RT 02 / RW 06, Desa Cipareuan, Kec. Cibiuk',
    city: 'Garut',
    province: 'Jawa Barat',
    postalCode: '44193',
    phone: '+62 895-3266-12155',
    email: 'smkitalhawari@gmail.com',
    website: 'al-hawari.sch.id',
    principalName: 'Dian Nurhamdani, S.Pd.',
    principalNip: '-',
    principalRank: 'Kepala SMK IT Al-Hawari',
    principalTitle: 'Kepala Sekolah',
    tuName: 'Tanti Nurkomalasari',
    tuNip: '-',
    tuRank: 'Tata Usaha SMK IT Al-Hawari',
    tuTitle: 'Kepala Tata Usaha',
    numberingPattern: '421.3/{urut}/SMKITA/{bulanRomawi}/{tahun}',
    majors: [
      'Desain Komunikasi Visual (DKV)',
      'Multimedia (MM)',
      'Agribisnis Tanaman Pangan dan Hortikultura (ATPH)',
      'Desain Pemodelan dan Informasi Bangunan (DPIB)'
    ]
  }),
  smp: Object.freeze({
    unitId: 'smp',
    unitLabel: 'SMP IT Al-Hawari - Cibiuk',
    unitShortBadge: 'SMP',
    schoolCode: 'SMPITA',
    government: 'PEMERINTAH KABUPATEN GARUT',
    department: 'DINAS PENDIDIKAN',
    subOffice: 'KECAMATAN CIBIUK',
    foundation: 'Yayasan Nunuk Kartini Al-Irsyad',
    npyp: 'AA2743',
    schoolName: 'SMP IT AL-HAWARI - CIBIUK',
    npsn: '20279898',
    skOperasional: '425.11/794-DISDIK (16 Mei 2019)',
    accreditation: 'Swasta / Akreditasi B',
    educationLevel: 'SMP (Sekolah Menengah Pertama)',
    address: 'Jl. Raya Cibiuk, Desa Cipareuan, Kec. Cibiuk',
    city: 'Garut',
    province: 'Jawa Barat',
    postalCode: '44193',
    phone: '+62 895-3266-12155',
    email: 'smpit_alhawari09@yahoo.co.id',
    website: 'al-hawari.sch.id',
    principalName: 'Yangyang Budiman, S.Pd.',
    principalNip: '-',
    principalRank: 'Kepala SMP IT Al-Hawari - Cibiuk',
    principalTitle: 'Kepala Sekolah',
    tuName: 'Yoga Yudistira',
    tuNip: '-',
    tuRank: 'Tata Usaha SMP IT Al-Hawari - Cibiuk',
    tuTitle: 'Kepala Tata Usaha',
    numberingPattern: '421.2/{urut}/SMPITA/{bulanRomawi}/{tahun}',
    majors: [
      'Kelas VII',
      'Kelas VIII',
      'Kelas IX'
    ]
  })
});

export function normalizeUnitId(unitId = 'smk') {
  const raw = String(unitId || '').trim().toLowerCase();
  if (raw === 'smp' || raw.includes('smp')) return 'smp';
  return 'smk';
}

export function getDefaultSchoolProfileForUnit(unitId = 'smk') {
  const id = normalizeUnitId(unitId);
  const template = SCHOOL_UNITS[id] || SCHOOL_UNITS.smk;
  return {
    ...JSON.parse(JSON.stringify(template)),
    logo: DEFAULT_SCHOOL_PROFILE.logo
  };
}

export function getSchoolCodeFromProfile(profile = {}, fallbackUnitId = '') {
  const explicit = String(profile?.schoolCode || '').trim().toUpperCase();
  if (explicit) return explicit;
  const pattern = String(profile?.numberingPattern || '');
  const match = pattern.match(/\/([A-Z0-9_-]+)\/\{bulanRomawi\}/i);
  if (match?.[1]) return match[1].toUpperCase();
  const unit = profile?.unitId || fallbackUnitId || (
    /\bsmp\b/i.test(String(profile?.schoolName || '')) ? 'smp' : 'smk'
  );
  return normalizeUnitId(unit) === 'smp' ? 'SMPITA' : 'SMKITA';
}

export const DEFAULT_SCHOOL_PROFILE = {
  ...JSON.parse(JSON.stringify(SCHOOL_UNITS.smk)),
  // Logo placeholder standar Pendidikan / Al-Hawari (Data URL PNG)
  logo: 'data:image/png;base64,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'
};

export function normalizeSmkText(value) {
  if (typeof value !== 'string' || !value) return value;
  return value.replace(/\bSMKS\b/g, 'SMK').replace(/\bsmks\b/g, 'smk');
}

export function normalizeSchoolProfileNaming(profile) {
  if (!profile || typeof profile !== 'object') return profile;
  const isSmp = profile.unitId === 'smp' || /\bsmp\b/i.test(String(profile.schoolName || ''));
  if (isSmp) {
    return {
      ...profile,
      unitId: 'smp',
      schoolCode: profile.schoolCode || 'SMPITA'
    };
  }
  let kopImg = profile.kopImage;
  if (typeof kopImg === 'string' && kopImg.startsWith('data:image/svg+xml')) {
    kopImg = kopImg.replace(/SMKS(?=(%20|\s))/g, 'SMK');
  }
  return {
    ...profile,
    unitId: profile.unitId || 'smk',
    schoolCode: profile.schoolCode || 'SMKITA',
    schoolName: normalizeSmkText(profile.schoolName),
    principalRank: normalizeSmkText(profile.principalRank),
    tuRank: normalizeSmkText(profile.tuRank),
    ...(kopImg !== undefined ? { kopImage: kopImg } : {})
  };
}

/**
 * Mendeteksi target unit sekolah ('smp' atau 'smk') dari teks prompt AI / isi dokumen.
 * Mengutamakan frasa target eksplisit ("untuk SMP", "kop SMK", nama Kepsek "Yangyang Budiman" / "Dian Nurhamdani").
 */
export function detectTargetUnitFromText(text = '', fallbackUnitId = null) {
  const s = String(text || '').trim();
  if (!s) return fallbackUnitId ? normalizeUnitId(fallbackUnitId) : null;

  // 1. Cek frasa perintah/tujuan eksplisit terlebih dahulu (mis. "ubah surat SMK ini untuk SMP")
  const explicitTargetRe = /(?:untuk|buat(?:kan)?|ke|versi|kop|kepsek|kepala\s*sekolah|atas\s*nama|unit|milik|tujuan|ubah\s*ke|ganti\s*ke|sesuaikan(?:\s+dengan)?)\s+(?:unit\s+|sekolah\s+|surat\s+)?(smp(?:\s*it)?|smk(?:\s*it)?)\b/gi;
  let lastExplicit = null;
  let m;
  while ((m = explicitTargetRe.exec(s)) !== null) {
    lastExplicit = /^smp/i.test(m[1]) ? 'smp' : 'smk';
  }
  if (lastExplicit) return lastExplicit;

  // 2. Nama kepala sekolah / TU / kode khas masing-masing unit
  const smpDirect = /\b(smp\s*it|smpit|smpita|smp|yangyang(?:\s+budiman)?|yoga\s+yudistira|20279898|421\.2)\b/i.test(s)
    || /\bkelas\s*(vii|viii|ix|7|8|9)\b/i.test(s);
  const smkDirect = /\b(smk\s*it|smkit|smkita|smks?|dian\s+nurhamdani|dian\s+nurhadian|tanti(?:\s+nurkomalasari)?|69758993|421\.3|kejuruan|prakerin|\bpkl\b|dkv|dpib|atph)\b/i.test(s)
    || /\bkelas\s*(x|xi|xii|10|11|12)\b/i.test(s);

  if (smpDirect && !smkDirect) return 'smp';
  if (smkDirect && !smpDirect) return 'smk';

  return fallbackUnitId ? normalizeUnitId(fallbackUnitId) : null;
}

/**
 * Mengambil profil sekolah lengkap + Kop SVG/custom yang cocok untuk unit target ('smp' | 'smk').
 */
export function resolveSchoolProfileForUnit(targetUnitId = 'smk', activeSchoolProfile = null, kopGallery = []) {
  const unitId = normalizeUnitId(targetUnitId);
  const base = getDefaultSchoolProfileForUnit(unitId);
  const activeMatchesUnit = Boolean(
    activeSchoolProfile &&
    typeof activeSchoolProfile === 'object' &&
    normalizeUnitId(activeSchoolProfile.unitId || (/\bsmp\b/i.test(String(activeSchoolProfile.schoolName || '')) ? 'smp' : 'smk')) === unitId
  );

  let resolvedKop = unitId === 'smp' ? DEFAULT_SMP_KOP_DATA_URL : DEFAULT_KOP_DATA_URL;
  if (activeMatchesUnit && activeSchoolProfile.kopImage) {
    resolvedKop = activeSchoolProfile.kopImage;
  } else if (Array.isArray(kopGallery) && kopGallery.length > 0) {
    const unitNeedle = unitId === 'smp' ? /\bsmp\b/i : /\bsmk\b/i;
    const matchedKop = kopGallery.find(k => k?.dataUrl && unitNeedle.test(String(k.name || k.id || '')));
    if (matchedKop?.dataUrl) {
      resolvedKop = matchedKop.dataUrl;
    }
  }

  const merged = activeMatchesUnit
    ? normalizeSchoolProfileNaming({ ...base, ...activeSchoolProfile, unitId })
    : normalizeSchoolProfileNaming({ ...base, unitId });

  return {
    ...merged,
    unitId,
    schoolCode: merged.schoolCode || (unitId === 'smp' ? 'SMPITA' : 'SMKITA'),
    kopImage: resolvedKop
  };
}

/**
 * Mengambil profil sekolah efektif yang melekat pada dokumen kanvas / surat keluar.
 */
export function resolveDocumentSchoolProfile(docOrItem = null, fallbackSchoolProfile = null, kopGallery = []) {
  if (!docOrItem || typeof docOrItem !== 'object') {
    const fbUnit = normalizeUnitId(fallbackSchoolProfile?.unitId || 'smk');
    return resolveSchoolProfileForUnit(fbUnit, fallbackSchoolProfile, kopGallery);
  }
  const embeddedProfile = docOrItem.letterData?.schoolProfile || docOrItem.schoolProfile || (docOrItem.blocks || []).find(b => b?.type === 'letterKop')?.school || null;
  const sigBlock = (docOrItem.blocks || []).find(b => b?.type === 'letterSignature') || null;
  const detectedUnit = docOrItem.targetUnitId
    || embeddedProfile?.unitId
    || detectTargetUnitFromText(
      [
        embeddedProfile?.schoolName || '',
        docOrItem.nomorSurat || docOrItem.letterData?.formData?.nomorSurat || '',
        sigBlock?.name || docOrItem.signerName || ''
      ].join(' '),
      null
    )
    || normalizeUnitId(fallbackSchoolProfile?.unitId || 'smk');

  return resolveSchoolProfileForUnit(detectedUnit, embeddedProfile || fallbackSchoolProfile, kopGallery);
}

/**
 * Mengadaptasi teks surat (nama sekolah, kepala sekolah, NPSN, kode unit, alamat, kop teks)
 * agar sesuai target unit ('smp' atau 'smk').
 */
export function adaptTextForSchoolUnit(text = '', targetUnitId = 'smk') {
  const raw = String(text ?? '');
  if (!raw) return raw;
  const unitId = normalizeUnitId(targetUnitId);
  const smk = SCHOOL_UNITS.smk;
  const smp = SCHOOL_UNITS.smp;

  if (unitId === 'smp') {
    return raw
      .replace(/PEMERINTAH\s+DAERAH\s+PROVINSI\s+JAWA\s+BARAT/g, smp.government)
      .replace(/CABANG\s+DINAS\s+PENDIDIKAN\s+WILAYAH\s+XI/g, smp.subOffice)
      .replace(/\bSMKS?\s+IT\s+AL-?HAWARI(?:\s*-\s*CIBIUK)?\b/gi, (m) => (m === m.toUpperCase() ? smp.schoolName : smp.unitLabel))
      .replace(/\bSMKITA\b/g, smp.schoolCode)
      .replace(/\bDian\s+Nurh(?:amdani|adian)(?:,\s*S\.?\s*Pd\.?)?/gi, smp.principalName)
      .replace(/\bTanti\s+Nurkomalasari\b/gi, smp.tuName)
      .replace(/\b69758993\b/g, smp.npsn)
      .replace(/smkitalhawari@gmail\.com/gi, smp.email)
      .replace(/Kp\.?\s*Curug\s*RT\.?\s*0?2\s*\/?\s*RW\.?\s*0?6\s*,?\s*Desa\s*Cipareuan\s*,?\s*Kec\.?\s*Cibiuk/gi, smp.address)
      .replace(/\b421\.3(?=\/[^\s]*KET\.SMPITA)/g, '421.2');
  }

  return raw
    .replace(/PEMERINTAH\s+KABUPATEN\s+GARUT/g, smk.government)
    .replace(/KECAMATAN\s+CIBIUK/g, smk.subOffice)
    .replace(/\bSMP\s+IT\s+AL-?HAWARI(?:\s*-\s*CIBIUK)?\b/gi, (m) => (m === m.toUpperCase() ? smk.schoolName : smk.unitLabel))
    .replace(/\bSMPITA\b/g, smk.schoolCode)
    .replace(/\bYangyang\s+Budiman(?:,\s*S\.?\s*Pd\.?)?/gi, smk.principalName)
    .replace(/\bYoga\s+Yudistira\b/gi, smk.tuName)
    .replace(/\b20279898\b/g, smk.npsn)
    .replace(/smpit_alhawari09@yahoo\.co\.id/gi, smk.email)
    .replace(/Jl\.?\s*Raya\s*Cibiuk\s*,?\s*Desa\s*Cipareuan\s*,?\s*Kec\.?\s*Cibiuk/gi, smk.address)
    .replace(/\b421\.2(?=\/[^\s]*KET\.SMKITA)/g, '421.3');
}

function adaptTextArrayDeep(value, targetUnitId) {
  if (typeof value === 'string') return adaptTextForSchoolUnit(value, targetUnitId);
  if (Array.isArray(value)) return value.map(item => adaptTextArrayDeep(item, targetUnitId));
  if (value && typeof value === 'object') {
    const out = {};
    for (const [k, v] of Object.entries(value)) {
      if (k === 'image' || k === 'qrImage' || k === 'kopImage' || k === 'logo' || k === 'qrDataUrl') {
        out[k] = v;
      } else {
        out[k] = adaptTextArrayDeep(v, targetUnitId);
      }
    }
    return out;
  }
  return value;
}

/**
 * Mengadaptasi DocumentModel lengkap (Kop, Nomor, Isi, Tabel SPPD, Blok TTD Kepsek, metadata unit)
 * agar siap dipakai & diedit oleh unit target ('smp' atau 'smk').
 */
export function adaptDocumentModelForUnit(
  docModel,
  targetUnitId = 'smk',
  { activeSchoolProfile = null, kopGallery = [], nextNomorSurat = '' } = {}
) {
  if (!docModel || typeof docModel !== 'object') return docModel;
  const unitId = normalizeUnitId(targetUnitId);
  const targetSchool = resolveSchoolProfileForUnit(unitId, activeSchoolProfile, kopGallery);
  const cloned = JSON.parse(JSON.stringify(docModel));
  delete cloned.verification;
  delete cloned.qrPlacement;

  let detectedNomor = nextNomorSurat ? String(nextNomorSurat).replace(/^Nomor:\s*/i, '').trim() : '';

  const adaptedBlocks = (cloned.blocks || []).map((block) => {
    if (!block || typeof block !== 'object') return block;

    if (block.type === 'letterKop') {
      return {
        ...block,
        school: targetSchool,
        image: targetSchool.kopImage || block.image
      };
    }

    if (block.type === 'letterTitle') {
      const updatedTitle = adaptTextForSchoolUnit(block.title || '', unitId);
      const rawNom = String(block.nomor || '').replace(/^Nomor:\s*/i, '').trim();
      const updatedNomor = detectedNomor || formatSmartNomorSurat({
        rawNomor: adaptTextForSchoolUnit(rawNom, unitId),
        contextText: `${updatedTitle} ${cloned.meta?.title || ''}`,
        schoolCode: targetSchool.schoolCode
      });
      detectedNomor = updatedNomor;
      return {
        ...block,
        title: updatedTitle,
        nomor: `Nomor: ${updatedNomor}`,
        ...(block.metaRecipient ? { metaRecipient: adaptTextArrayDeep(block.metaRecipient, unitId) } : {})
      };
    }

    if (block.type === 'letterSignature') {
      const prevTitle = String(block.title || '').trim();
      const isTuSigner = /\btata\s*usaha\b|\bktu\b|\btu\b/i.test(prevTitle) && !/kepala\s*sekolah/i.test(prevTitle);
      return {
        ...block,
        city: targetSchool.city || block.city || 'Garut',
        title: isTuSigner ? (targetSchool.tuTitle || prevTitle || 'Kepala Tata Usaha') : (targetSchool.principalTitle || 'Kepala Sekolah'),
        name: isTuSigner ? (targetSchool.tuName || block.name) : targetSchool.principalName,
        nip: isTuSigner ? (targetSchool.tuNip || '-') : (targetSchool.principalNip || '-')
      };
    }

    return adaptTextArrayDeep(block, unitId);
  });

  cloned.blocks = adaptedBlocks;
  cloned.targetUnitId = unitId;
  cloned.meta = {
    ...(cloned.meta || {}),
    title: adaptTextForSchoolUnit(cloned.meta?.title || 'SURAT DINAS', unitId),
    author: targetSchool.schoolName
  };
  cloned.letterData = {
    ...(cloned.letterData || {}),
    schoolProfile: targetSchool,
    formData: {
      ...(cloned.letterData?.formData ? adaptTextArrayDeep(cloned.letterData.formData, unitId) : {}),
      ...(detectedNomor ? { nomorSurat: detectedNomor } : {}),
      signerName: targetSchool.principalName
    }
  };

  return cloned;
}


