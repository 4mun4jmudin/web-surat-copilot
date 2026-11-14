import { test } from 'node:test';
import assert from 'node:assert/strict';
import { checkDocumentConsistency, autoFixDocument } from '../js/checks/consistency.js';
import { samplePklDocument } from '../js/model/sample.js';

test('Consistency Check: samplePklDocument minim masalah', () => {
  const result = checkDocumentConsistency(samplePklDocument);
  assert.equal(result.summary.errors, 0, `Tidak boleh ada error fatal: ${JSON.stringify(result.issues)}`);
});

test('Consistency Check: Mendeteksi rujukan silang rusak dan sitasi hilang', () => {
  const brokenDoc = {
    meta: { title: 'Test Broken' },
    profile: 'pkl',
    blocks: [
      { type: 'crossRef', targetId: 'fig_unknown' },
      { type: 'citation', refId: 'ref_nonexistent' }
    ]
  };

  const result = checkDocumentConsistency(brokenDoc);
  assert.ok(result.summary.errors > 0 || result.summary.warnings > 0);
  const brokenRefIssue = result.issues.find(i => i.code === 'BROKEN_CROSS_REFERENCE');
  const unresolvedCitIssue = result.issues.find(i => i.code === 'UNRESOLVED_CITATION');

  assert.ok(brokenRefIssue, 'Harus mendeteksi broken cross reference');
  assert.ok(unresolvedCitIssue, 'Harus mendeteksi unresolved citation');
});

test('Consistency Check: Mendeteksi heading level yang melompat', () => {
  const docWithJump = {
    meta: { title: 'Heading Jump' },
    profile: 'pkl',
    blocks: [
      { type: 'heading', level: 1, text: 'BAB I' },
      { type: 'heading', level: 3, text: '1.1.1 Sub-sub Bab' } // Melompat dari 1 langsung ke 3!
    ]
  };

  const result = checkDocumentConsistency(docWithJump);
  const jumpIssue = result.issues.find(i => i.code === 'HEADING_LEVEL_SKIPPED');
  assert.ok(jumpIssue, 'Harus mendeteksi heading level skipped');
});

test('Consistency Check: Auto-fix membersihkan paragraf kosong berlebih', () => {
  const dirtyDoc = {
    meta: { title: 'Dirty' },
    profile: 'pkl',
    blocks: [
      { type: 'paragraph', text: 'Paragraf 1' },
      { type: 'paragraph', text: '' },
      { type: 'paragraph', text: '' },
      { type: 'paragraph', text: '' },
      { type: 'paragraph', text: 'Paragraf 2' }
    ]
  };

  const cleaned = autoFixDocument(dirtyDoc);
  // Paragraf kosong berlebih harus dikurangi
  assert.ok(cleaned.blocks.length < dirtyDoc.blocks.length);
});

