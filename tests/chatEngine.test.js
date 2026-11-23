import test from 'node:test';
import assert from 'node:assert/strict';

import {
  CHAT_FOLDERS,
  DIRECT_CANVAS_SYSTEM_SUFFIX,
  isDirectExecutionCommand,
  createChatSession,
  createChatMessage,
  deriveChatTitleFromMessage,
  formatChatTimestamp,
  computeChatMemoryStats,
  filterChatSessions,
  buildConversationHistoryForAi,
  buildChatCopilotUserPrompt,
  normalizeAiChatResponse
} from '../js/ai/chatEngine.js';
import { store } from '../js/state.js';

test('chatEngine: pembuatan session, penamaan otomatis, folder filter (Dipin/Terbaru/Arsip/Sampah), & memory stats', () => {
  const s1 = createChatSession({ id: 'c1', title: 'Percakapan Baru', targetWorkspace: 'smart-canvas' });
  const s2 = createChatSession({ id: 'c2', title: 'Format Undangan Rapat', pinned: true });
  const s3 = createChatSession({ id: 'c3', title: 'Draf Lama', archived: true });
  const s4 = createChatSession({ id: 'c4', title: 'Dihapus', trashed: true });

  assert.equal(deriveChatTitleFromMessage('tolong perbaiki ejaan surat keterangan'), 'Perbaiki ejaan surat keterangan');
  assert.equal(deriveChatTitleFromMessage('tolong perbaiki ejaan surat keterangan berkelakuan baik ini'), 'Perbaiki ejaan surat keterangan berkela...');
  assert.ok(formatChatTimestamp(new Date().toISOString()).includes(':'));

  const stats = computeChatMemoryStats([s1, s2, s3, s4]);
  assert.equal(stats.allCount, 2);
  assert.equal(stats.pinnedCount, 1);
  assert.equal(stats.archivedCount, 1);
  assert.equal(stats.trashCount, 1);

  const pinned = filterChatSessions([s1, s2, s3, s4], { folder: CHAT_FOLDERS.PINNED });
  assert.deepEqual(pinned.map(x => x.id), ['c2']);

  const all = filterChatSessions([s1, s2, s3, s4], { folder: CHAT_FOLDERS.ALL });
  assert.equal(all[0].id, 'c2', 'Chat dipin selalu tampil di atas');
  assert.equal(all.length, 2);

  const trash = filterChatSessions([s1, s2, s3, s4], { folder: CHAT_FOLDERS.TRASH });
  assert.deepEqual(trash.map(x => x.id), ['c4']);

  const bySearch = filterChatSessions([s1, s2, s3], { folder: CHAT_FOLDERS.ALL, query: 'undangan' });
  assert.deepEqual(bySearch.map(x => x.id), ['c2']);
});

test('chatEngine: riwayat multi-turn & normalisasi respons AI (chat vs create_letter vs edit_letter + surgical lock SPPD)', () => {
  const msgs = [
    createChatMessage({ role: 'user', text: 'Kode surat untuk undangan rapat guru apa?' }),
    createChatMessage({
      role: 'assistant',
      text: 'Gunakan kode UND.',
      checklist: ['Format: 050/UND.SMKITA/XI/2026'],
      followUpQuestion: 'Mau saya buatkan drafnya sekalian?'
    })
  ];
  const hist = buildConversationHistoryForAi(msgs, 8);
  assert.equal(hist.length, 2);
  assert.equal(hist[0].role, 'user');
  assert.equal(hist[1].role, 'assistant');
  assert.ok(hist[1].text.includes('UND.SMKITA'));

  // Konsultasi murni (chat)
  const chatRes = normalizeAiChatResponse({
    intent: 'chat',
    reply: 'Untuk rapat guru gunakan kode UND.',
    checklist: ['Sertakan hari, tanggal, waktu, tempat'],
    followUpQuestion: 'Apakah ingin langsung dibuatkan drafnya?'
  }, { userMessage: 'kode surat undangan apa?', hasActiveDoc: true });
  assert.equal(chatRes.intent, 'chat');
  assert.equal(chatRes.letterDraft, null);
  assert.equal(chatRes.editPlan, null);

  // Perintah edit biasa dipaksa tetap surgical agar tabel SPPD tidak dirombak
  const editRes = normalizeAiChatResponse({
    intent: 'edit_letter',
    reply: 'Saya akan memperbaiki ejaan surat keterangan ini tanpa mengubah struktur tabel.',
    checklist: ['Memastikan ejaan sesuai PUEBI', 'Menjaga kolom tetap rapi'],
    editPlan: {
      editScope: 'restructure',
      patches: [{ op: 'setParagraph', blockId: 'p0_b0', text: 'Yang bertanda tangan di bawah ini:' }]
    }
  }, { userMessage: 'Perbaiki ejaan dan format surat keterangan ini tanpa ubah tabel', hasActiveDoc: true });
  assert.equal(editRes.intent, 'edit_letter');
  assert.equal(editRes.editPlan.editScope, 'surgical');
  assert.equal(editRes.checklist.length, 2);

  // Fallback jika LLM mengembalikan letterTitle/blocks langsung
  const draftRes = normalizeAiChatResponse({
    letterTitle: 'SURAT UNDANGAN',
    blocks: [{ type: 'paragraph', text: 'Dengan hormat...' }]
  }, { userMessage: 'Buatkan surat undangan', hasActiveDoc: false });
  assert.equal(draftRes.intent, 'create_letter');
  assert.equal(draftRes.letterDraft.letterTitle, 'SURAT UNDANGAN');
  assert.ok(draftRes.followUpQuestion.includes('menerapkan'), 'Mode confirm menyertakan tawaran terapkan');

  // Mode direct_apply (drawer Word Pro / Lembar Kerja) langsung eksekusi tanpa tanya konfirmasi penerapan
  const directPrompt = buildChatCopilotUserPrompt({
    userMessage: 'Buatkan surat undangan rapat besok',
    executionMode: 'direct_apply'
  });
  assert.ok(directPrompt.includes('MODE EKSEKUSI LANGSUNG KE KANVAS AKTIF'));

  const directEditRes = normalizeAiChatResponse({
    intent: 'edit_letter',
    reply: 'PUEBI diperbarui.',
    followUpQuestion: 'Apakah Anda ingin saya langsung menerapkan perbaikannya pada editor?',
    editPlan: { editScope: 'surgical', patches: [{ op: 'setParagraph', blockId: 'p0_b0', text: 'Dengan hormat,' }] }
  }, { userMessage: 'Rapikan ejaan', hasActiveDoc: true, executionMode: 'direct_apply' });
  assert.equal(directEditRes.intent, 'edit_letter');
  assert.equal(directEditRes.followUpQuestion, '', 'Pertanyaan konfirmasi menerapkan dihilangkan saat direct_apply');

  // Auto-promote jika model melabeli 'chat' tapi tetap menyertakan blocks/patches saat direct_apply
  const promotedRes = normalizeAiChatResponse({
    intent: 'chat',
    reply: 'Berikut drafnya.',
    letterDraft: {
      letterTitle: 'SURAT KETERANGAN',
      blocks: [{ type: 'paragraph', text: 'Dengan hormat,' }]
    }
  }, { userMessage: 'Buatkan surat keterangan', hasActiveDoc: false, executionMode: 'direct_apply' });
  assert.equal(promotedRes.intent, 'create_letter');
  assert.equal(promotedRes.letterDraft.letterTitle, 'SURAT KETERANGAN');
  assert.ok(DIRECT_CANVAS_SYSTEM_SUFFIX.includes('LANGSUNG eksekusi'));

  // Deteksi perintah aksi langsung vs tanya jawab konsultasi biasa
  assert.equal(isDirectExecutionCommand('Buatkan surat undangan rapat guru besok jam 09.00'), true);
  assert.equal(isDirectExecutionCommand('Ubah margin kiri jadi 3 cm dan spasi 1.5'), true);
  assert.equal(isDirectExecutionCommand('Apa kode surat yang tepat untuk undangan rapat?'), false);

  // Pada kanvas kosong di direct_apply, edit_letter dengan blocks penuh dipromosikan menjadi create_letter
  const blankDirectCreate = normalizeAiChatResponse({
    intent: 'edit_letter',
    reply: 'Surat telah dibuat.',
    editPlan: {
      letterTitle: 'SURAT PEMBERITAHUAN',
      blocks: [{ type: 'paragraph', text: 'Dengan hormat,' }]
    }
  }, { userMessage: 'Buatkan surat pemberitahuan', hasActiveDoc: false, executionMode: 'direct_apply' });
  assert.equal(blankDirectCreate.intent, 'create_letter');
  assert.equal(blankDirectCreate.letterDraft.letterTitle, 'SURAT PEMBERITAHUAN');
});

test('StateStore: manajemen sesi chat multi-turn (create, append, pin, archive, trash, auto-apply toggle)', () => {
  const created = store.createChatSession({ title: 'Percakapan Baru', targetWorkspace: 'smart-canvas' });
  assert.ok(created.id);

  store.appendChatMessage(created.id, { role: 'user', text: 'Tolong rapikan ejaan surat keterangan berkelakuan baik' });
  const afterMsg = store.getChatSessions().find(s => s.id === created.id);
  assert.equal(afterMsg.messages.length, 1);
  assert.ok(afterMsg.title.toLowerCase().includes('rapikan'), 'Judul otomatis diambil dari pesan pertama user');

  store.togglePinChatSession(created.id);
  assert.equal(store.getChatSessions().find(s => s.id === created.id).pinned, true);

  store.setChatAutoApply(true);
  assert.equal(store.getChatAutoApply(), true);

  store.trashChatSession(created.id, true);
  assert.equal(store.getChatSessions().find(s => s.id === created.id).trashed, true);

  store.deleteChatSessionPermanent(created.id);
  assert.equal(store.getChatSessions().some(s => s.id === created.id), false);
});
