import { test } from 'node:test';
import assert from 'node:assert/strict';
import { validateDocument, normalizeDocument } from '../js/model/schema.js';
import { samplePklDocument } from '../js/model/sample.js';

test('Schema Validator: Dokumen valid samplePklDocument lolos validasi', () => {
  const result = validateDocument(samplePklDocument);
  assert.equal(result.valid, true, `Seharusnya valid, error: ${result.errors.join('; ')}`);
  assert.equal(result.errors.length, 0);
});

test('Schema Validator: Menolak dokumen kosong atau non-objek', () => {
  assert.equal(validateDocument(null).valid, false);
  assert.equal(validateDocument([]).valid, false);
  assert.equal(validateDocument({}).valid, false);
});

test('Schema Validator: Memastikan pengecekan level heading (harus 1, 2, atau 3)', () => {
  const invalidDoc = {
    meta: { title: 'Test' },
    profile: 'pkl',
    blocks: [
      { type: 'heading', level: 4, text: 'Heading 4 tidak diizinkan' }
    ]
  };
  const result = validateDocument(invalidDoc);
  assert.equal(result.valid, false);
  assert.match(result.errors[0], /angka 1, 2, atau 3/);
});

test('Schema Validator: Normalisasi paragraph text menjadi runs', () => {
  const rawDoc = {
    meta: { title: 'Test' },
    profile: 'pkl',
    blocks: [
      { type: 'paragraph', text: 'Paragraf string tunggal' }
    ]
  };
  const normalized = normalizeDocument(rawDoc);
  assert.equal(Array.isArray(normalized.blocks[0].runs), true);
  assert.equal(normalized.blocks[0].runs[0].text, 'Paragraf string tunggal');
});

