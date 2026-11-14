import { test } from 'node:test';
import assert from 'node:assert/strict';
import { renderDocx } from '../js/docx/render.js';
import { patchUploadedDocx, extractParagraphs } from '../js/docx/patch.js';
import { parseXmlAsync } from '../js/utils/xml.js';
import { samplePklDocument } from '../js/model/sample.js';

test('Mode B (Patcher): Sisipkan paragraf baru tanpa merusak sectPr', async () => {
  // Buat file docx awal
  const originalDocx = await renderDocx(samplePklDocument, 'uint8array');

  // Jalankan patch: insertAfter
  const ops = [
    {
      action: 'insertAfter',
      anchor: { type: 'text', value: 'Latar Belakang Praktik Kerja Lapangan' },
      text: 'Ini adalah paragraf penjelasan tambahan yang disisipkan secara mulus oleh AI.'
    }
  ];

  const result = await patchUploadedDocx(originalDocx, ops, { outputType: 'uint8array' });

  assert.equal(result.report.safetyPassed, true, 'Safety audit harus lolos');
  assert.equal(result.report.executedChanges.length, 1);
  assert.match(result.patchedXml, /Ini adalah paragraf penjelasan tambahan yang disisipkan/);

  // Verifikasi paragraf di dalam DOM patched
  const patchedDom = await parseXmlAsync(result.patchedXml);
  const paragraphs = extractParagraphs(patchedDom);
  const found = paragraphs.some(p => p.fullText.includes('Ini adalah paragraf penjelasan tambahan'));
  assert.equal(found, true, 'Paragraf baru harus ditemukan dalam daftar paragraf yang diekstrak');
});

test('Mode B (Patcher): Dukungan Track Changes (<w:ins> dan <w:del>)', async () => {
  const originalDocx = await renderDocx(samplePklDocument, 'uint8array');

  const ops = [
    {
      action: 'replace',
      anchor: { type: 'text', value: 'Berdasarkan pelaksanaan PKL selama tiga bulan' },
      text: 'Berdasarkan evaluasi komprehensif selama pelaksanaan PKL di industri'
    }
  ];

  const result = await patchUploadedDocx(originalDocx, ops, {
    trackChanges: true,
    author: 'Penguji AI',
    outputType: 'uint8array'
  });

  assert.equal(result.report.safetyPassed, true);
  // Verifikasi ada elemen w:ins dan w:del
  assert.match(result.patchedXml, /<w:ins w:id="[^"]+" w:author="Penguji AI"/);
  assert.match(result.patchedXml, /<w:del w:id="[^"]+" w:author="Penguji AI"/);
});

