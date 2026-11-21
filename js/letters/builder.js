/**
 * builder.js - Mesin Penyusun Komponen OOXML Surat Dinas Sekolah
 * Mendukung Kop Surat Gambar Banner, Kop Surat Tabel Profil, Judul Surat, Tabel Identitas, dan Blok TTD.
 */

import { escapeXml, buildTextElement } from '../utils/xml.js';

/**
 * Mengonversi base64 data URL menjadi Uint8Array
 * @param {string} dataUrl 
 * @returns {{ bytes: Uint8Array, extension: string, mime: string }|null}
 */
export function parseBase64Image(dataUrl) {
  if (!dataUrl || typeof dataUrl !== 'string') return null;

  const match = dataUrl.match(/^data:(image\/([a-zA-Z0-9+]+));base64,(.+)$/);
  if (match) {
    const mime = match[1];
    let extension = match[2].toLowerCase();
    if (extension === 'jpeg') extension = 'jpg';
    const base64Str = match[3];

    let binaryStr;
    if (typeof atob !== 'undefined') {
      binaryStr = atob(base64Str);
    } else {
      binaryStr = Buffer.from(base64Str, 'base64').toString('binary');
    }

    const len = binaryStr.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
      bytes[i] = binaryStr.charCodeAt(i);
    }
    return { bytes, extension, mime };
  }

  if (dataUrl.startsWith('data:image/svg+xml')) {
    const base64Index = dataUrl.indexOf(';base64,');
    if (base64Index !== -1) {
      const b64 = dataUrl.slice(base64Index + 8);
      const binaryStr = typeof atob !== 'undefined' ? atob(b64) : Buffer.from(b64, 'base64').toString('binary');
      const bytes = new Uint8Array(binaryStr.length);
      for (let i = 0; i < binaryStr.length; i++) bytes[i] = binaryStr.charCodeAt(i);
      return { bytes, extension: 'svg', mime: 'image/svg+xml' };
    }
  }

  return null;
}

/**
 * Menghasilkan DrawingML untuk Logo di Kop Surat
 */
export function buildLogoDrawingXml(relId = 'rIdLogo', widthCm = 2.1, heightCm = 2.1) {
  const cx = Math.round(widthCm * 360000);
  const cy = Math.round(heightCm * 360000);

  return `
    <w:drawing>
      <wp:inline distT="0" distB="0" distL="0" distR="0">
        <wp:extent cx="${cx}" cy="${cy}"/>
        <wp:docPr id="1" name="Logo Sekolah"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:nvPicPr>
                <pic:cNvPr id="1" name="logo.png"/>
                <pic:cNvPicPr/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="${relId}"/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="${cx}" cy="${cy}"/>
                </a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>
  `.trim();
}

/**
 * Menghasilkan DrawingML untuk Kop Surat Gambar Banner
 * Lebar standar: 16.5 cm (5,940,000 EMUs), Tinggi proporsional ~3.3 cm (1,188,000 EMUs)
 */
export function buildKopBannerDrawingXml(relId = 'rIdKop', widthCm = 16.5, heightCm = 3.3) {
  const cx = Math.round(widthCm * 360000);
  const cy = Math.round(heightCm * 360000);

  return `
    <w:p>
      <w:pPr>
        <w:jc w:val="center"/>
        <w:spacing w:before="0" w:after="180" w:line="240" w:lineRule="auto"/>
      </w:pPr>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="${cx}" cy="${cy}"/>
            <wp:docPr id="1" name="Kop Surat"/>
            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:nvPicPr>
                    <pic:cNvPr id="1" name="kop.png"/>
                    <pic:cNvPicPr/>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="${relId}"/>
                    <a:stretch><a:fillRect/></a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="${cx}" cy="${cy}"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  `.trim();
}

/**
 * Menghasilkan OOXML untuk Kop Surat Berbasis Profil Sekolah
 */
export function buildKopSuratXml(school, hasLogo = true) {
  const gov = school.government || 'PEMERINTAH DAERAH PROVINSI JAWA BARAT';
  const dept = school.department || 'DINAS PENDIDIKAN';
  const sub = school.subOffice || 'CABANG DINAS PENDIDIKAN WILAYAH XI';
  const name = school.schoolName || 'SMK IT AL-HAWARI';
  const addr = school.address || '';
  const city = school.city ? `${school.city} ${school.postalCode || ''}` : '';
  const contact = `Telp. ${school.phone || '-'} | Email: ${school.email || '-'} | Web: ${school.website || '-'}`;

  const logoColWidth = hasLogo ? 1600 : 0;
  const textColWidth = hasLogo ? 8200 : 9800;

  return `
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="0" w:type="auto"/>
        <w:tblBorders>
          <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/>
        </w:tblBorders>
        <w:jc w:val="center"/>
      </w:tblPr>
      <w:tr>
        ${hasLogo ? `
        <w:tc>
          <w:tcPr><w:tcW w:w="${logoColWidth}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r>${buildLogoDrawingXml('rIdLogo', 2.1, 2.1)}</w:r>
          </w:p>
        </w:tc>` : ''}
        <w:tc>
          <w:tcPr><w:tcW w:w="${textColWidth}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:after="20" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/><w:b/></w:rPr><w:t>${escapeXml(gov.toUpperCase())}</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:after="20" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/><w:b/></w:rPr><w:t>${escapeXml(dept.toUpperCase())}</w:t></w:r>
          </w:p>
          ${sub ? `
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:after="20" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/><w:b/></w:rPr><w:t>${escapeXml(sub.toUpperCase())}</w:t></w:r>
          </w:p>` : ''}
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:after="40" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="28"/><w:b/></w:rPr><w:t>${escapeXml(name.toUpperCase())}</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="18"/></w:rPr><w:t>${escapeXml(addr)}, ${escapeXml(city)}</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:jc w:val="center"/><w:spacing w:after="0" w:line="220" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="18"/></w:rPr><w:t>${escapeXml(contact)}</w:t></w:r>
          </w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:p>
      <w:pPr>
        <w:pBdr><w:bottom w:val="double" w:sz="18" w:space="4" w:color="000000"/></w:pBdr>
        <w:spacing w:before="60" w:after="240" w:line="240" w:lineRule="auto"/>
      </w:pPr>
      <w:r><w:t></w:t></w:r>
    </w:p>
  `.trim();
}

/**
 * Menghasilkan OOXML untuk Judul Surat & Nomor Surat
 */
export function buildLetterTitleXml(title, nomor, metaRecipient = null) {
  return `
    <w:p>
      <w:pPr>
        <w:jc w:val="center"/>
        <w:spacing w:before="120" w:after="20" w:line="240" w:lineRule="auto"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
          <w:b/>
          <w:u w:val="single"/>
          <w:sz w:val="28"/>
        </w:rPr>
        <w:t>${escapeXml(title)}</w:t>
      </w:r>
    </w:p>
    ${nomor ? `
    <w:p>
      <w:pPr>
        <w:jc w:val="center"/>
        <w:spacing w:before="0" w:after="180" w:line="240" w:lineRule="auto"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
          <w:sz w:val="24"/>
        </w:rPr>
        <w:t>${escapeXml(nomor)}</w:t>
      </w:r>
    </w:p>` : ''}
    ${metaRecipient ? buildLetterMetaRecipientXml(metaRecipient) : ''}
  `.trim();
}

/**
 * Menghasilkan OOXML untuk Blok Sifat, Perihal & Kepada Yth (Di Tempat)
 */
export function buildLetterMetaRecipientXml(block = {}) {
  if (!block || block.isDeleted) return '';

  const metaRows = Array.isArray(block.metaRows) ? block.metaRows.filter(r => (r.label && r.label.trim()) || (r.value && r.value.trim())) : [];
  const recipientLines = Array.isArray(block.recipientLines) ? block.recipientLines.filter(l => typeof l === 'string' && l.trim().length > 0) : [];

  if (metaRows.length === 0 && recipientLines.length === 0) return '';

  const parts = [];

  if (metaRows.length > 0) {
    const rowsXml = metaRows.map(r => {
      const lbl = (r.label || '').trim();
      const val = (r.value || '').trim().replace(/^:\s*/, '');
      const isPerihal = /^(perihal|hal)$/i.test(lbl);
      return `
        <w:tr>
          <w:trPr><w:cantSplit/></w:trPr>
          <w:tc>
            <w:tcPr>
              <w:tcW w:w="1300" w:type="dxa"/>
              <w:tcMar><w:top w:w="20" w:type="dxa"/><w:bottom w:w="20" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:right w:w="40" w:type="dxa"/></w:tcMar>
            </w:tcPr>
            <w:p>
              <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
              <w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>${escapeXml(lbl)}</w:t></w:r>
            </w:p>
          </w:tc>
          <w:tc>
            <w:tcPr>
              <w:tcW w:w="220" w:type="dxa"/>
              <w:tcMar><w:top w:w="20" w:type="dxa"/><w:bottom w:w="20" w:type="dxa"/><w:left w:w="20" w:type="dxa"/><w:right w:w="40" w:type="dxa"/></w:tcMar>
            </w:tcPr>
            <w:p>
              <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
              <w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>:</w:t></w:r>
            </w:p>
          </w:tc>
          <w:tc>
            <w:tcPr>
              <w:tcW w:w="7500" w:type="dxa"/>
              <w:tcMar><w:top w:w="20" w:type="dxa"/><w:bottom w:w="20" w:type="dxa"/><w:left w:w="20" w:type="dxa"/><w:right w:w="40" w:type="dxa"/></w:tcMar>
            </w:tcPr>
            <w:p>
              <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
              <w:r>
                <w:rPr>
                  <w:sz w:val="24"/>
                  ${isPerihal ? '<w:b/><w:i/>' : ''}
                </w:rPr>
                <w:t>${escapeXml(val)}</w:t>
              </w:r>
            </w:p>
          </w:tc>
        </w:tr>
      `;
    }).join('');

    parts.push(`
      <w:tbl>
        <w:tblPr>
          <w:tblW w:w="0" w:type="auto"/>
          <w:tblInd w:w="0" w:type="dxa"/>
          <w:tblBorders>
            <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/>
          </w:tblBorders>
        </w:tblPr>
        ${rowsXml}
      </w:tbl>
    `);
  }

  if (recipientLines.length > 0) {
    recipientLines.forEach((line, idx) => {
      const trimmed = line.trim();
      const isDi = /^di$/i.test(trimmed);
      const isLast = idx === recipientLines.length - 1;
      const beforeTwips = (idx === 0 && metaRows.length > 0) ? 140 : (isDi ? 120 : 0);
      const afterTwips = isLast ? 160 : 20;

      parts.push(`
        <w:p>
          <w:pPr>
            <w:jc w:val="left"/>
            <w:spacing w:before="${beforeTwips}" w:after="${afterTwips}" w:line="240" w:lineRule="auto"/>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
              <w:sz w:val="24"/>
            </w:rPr>
            <w:t>${escapeXml(trimmed)}</w:t>
          </w:r>
        </w:p>
      `);
    });
  } else if (metaRows.length > 0) {
    parts.push(`<w:p><w:pPr><w:spacing w:before="60" w:after="100"/></w:pPr></w:p>`);
  }

  return parts.join('\n').trim();
}

/**
 * Menghasilkan OOXML untuk Tabel Identitas Siswa/Pegawai (tanpa border, indent rapi)
 */
export function buildIdentityTableXml(rows = [], colWidthsPct = null) {
  const hasCustomW = Array.isArray(colWidthsPct) && colWidthsPct.length >= 2;
  const p0 = hasCustomW ? Math.max(10, Math.min(80, Number(colWidthsPct[0]) || 28)) : 28;
  const w0 = hasCustomW ? Math.round(9400 * (p0 / 100)) : 2600;
  const w1 = hasCustomW ? Math.max(1200, 9400 - w0) : 6800;

  const rowsXml = rows.map(r => `
    <w:tr>
      <w:trPr><w:cantSplit/></w:trPr>
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="${w0}" w:type="dxa"/>
          <w:tcMar><w:top w:w="40" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:right w:w="40" w:type="dxa"/></w:tcMar>
        </w:tcPr>
        <w:p>
          <w:pPr><w:spacing w:after="0" w:line="260" w:lineRule="auto"/></w:pPr>
          <w:r><w:rPr><w:sz w:val="24"/></w:rPr>${buildTextElement(r.label)}</w:r>
        </w:p>
      </w:tc>
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="${w1}" w:type="dxa"/>
          <w:tcMar><w:top w:w="40" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:left w:w="40" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tcMar>
        </w:tcPr>
        <w:p>
          <w:pPr><w:spacing w:after="0" w:line="260" w:lineRule="auto"/></w:pPr>
          <w:r><w:rPr><w:sz w:val="24"/><w:b/></w:rPr>${buildTextElement(r.value)}</w:r>
        </w:p>
      </w:tc>
    </w:tr>
  `).join('');

  return `
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="0" w:type="auto"/>
        <w:tblInd w:w="567" w:type="dxa"/>
        <w:tblBorders>
          <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="${w0}"/><w:gridCol w:w="${w1}"/></w:tblGrid>
      ${rowsXml}
    </w:tbl>
    <w:p><w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr></w:p>
  `.trim();
}

/**
/**
 * Menghasilkan OOXML drawing untuk Barcode QR Pengesahan Surat Keluar
 */
export function buildQrDrawingXml(relId = 'rIdQr', widthCm = 2.4, heightCm = 2.4) {
  const cx = Math.round(widthCm * 360000);
  const cy = Math.round(heightCm * 360000);
  return `
    <w:p>
      <w:pPr><w:spacing w:before="80" w:after="40" w:line="240" w:lineRule="auto"/></w:pPr>
      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="${cx}" cy="${cy}"/>
            <wp:docPr id="99" name="QR Pengesahan Surat"/>
            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:nvPicPr>
                    <pic:cNvPr id="99" name="qr.png"/>
                    <pic:cNvPicPr/>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="${relId}"/>
                    <a:stretch><a:fillRect/></a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="${cx}" cy="${cy}"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
  `.trim();
}

/**
 * Menghasilkan OOXML untuk Blok Tanda Tangan Resmi (rata kanan bawah),
 * mendukung mode Barcode QR pengesahan atau ruang kosong TTD manual pulpen.
 */
export function buildLetterSignatureXml(sig, options = {}) {
  const city = sig.city || 'Garut';
  const date = sig.date || '';
  const title = sig.title || 'Kepala Sekolah';
  const name = sig.name || 'Dian Nurhadian, S.Pd.';
  const rawNip = String(sig.nip ?? '').trim();
  const nip = (!rawNip || rawNip === '-')
    ? 'NIP. -'
    : (/^NIP\b/i.test(rawNip) ? rawNip : `NIP. ${rawNip}`);
  const useQr = sig.signatureMode === 'qr' && (options?.hasQrImage || Boolean(sig.qrImage));
  const qrCaption = sig.qrCaption || 'Disahkan secara elektronik';
  const qrSubCaption = String(sig.qrSubCaption || '').trim();

  const signatureBodyXml = useQr
    ? `
          ${buildQrDrawingXml('rIdQr', 2.4, 2.4)}
          <w:p>
            <w:pPr><w:spacing w:before="0" w:after="${qrSubCaption ? '10' : '40'}" w:line="200" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:sz w:val="16"/><w:i/><w:color w:val="166534"/></w:rPr><w:t>${escapeXml(qrCaption)}</w:t></w:r>
          </w:p>
          ${qrSubCaption ? `
          <w:p>
            <w:pPr><w:spacing w:before="0" w:after="40" w:line="200" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:sz w:val="17"/><w:b/><w:color w:val="166534"/></w:rPr><w:t>${escapeXml(qrSubCaption)}</w:t></w:r>
          </w:p>` : ''}
          <w:p>
            <w:pPr><w:spacing w:before="40" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r>
              <w:rPr><w:sz w:val="24"/><w:b/><w:u w:val="single"/></w:rPr>
              <w:t>${escapeXml(name)}</w:t>
            </w:r>
          </w:p>
    `
    : `
          <w:p>
            <w:pPr><w:spacing w:before="720" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r>
              <w:rPr><w:sz w:val="24"/><w:b/><w:u w:val="single"/></w:rPr>
              <w:t>${escapeXml(name)}</w:t>
            </w:r>
          </w:p>
    `;

  return `
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="0" w:type="auto"/>
        <w:tblBorders>
          <w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/>
        </w:tblBorders>
      </w:tblPr>
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="5400" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t></w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="4400" w:type="dxa"/></w:tcPr>
          <w:p>
            <w:pPr><w:spacing w:before="180" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>${escapeXml(city)}, ${escapeXml(date)}</w:t></w:r>
          </w:p>
          <w:p>
            <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>${escapeXml(title)},</w:t></w:r>
          </w:p>
          ${signatureBodyXml}
          ${nip ? `
          <w:p>
            <w:pPr><w:spacing w:after="0" w:line="220" w:lineRule="auto"/></w:pPr>
            <w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>${escapeXml(nip)}</w:t></w:r>
          </w:p>` : ''}
        </w:tc>
      </w:tr>
    </w:tbl>
  `.trim();
}
