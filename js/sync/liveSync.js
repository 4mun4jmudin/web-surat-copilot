/**
 * liveSync.js - Mesin Polling Ringan Status Antrean ACC & Pengesahan QR Tanpa Mengganggu Pekerjaan
 * Hanya mengecek status ACC/Surat Keluar secara ringan (lite=1) agar tidak menyinkronkan atau menimpa
 * kanvas, draf, profil, maupun obrolan AI antar perangkat.
 */

import { store } from '../state.js';
import { getActiveWorkspaceUnitId, getStoredAuthToken, pollWorkspaceUpdates } from './workspaceClient.js';

const POLL_INTERVAL_VISIBLE_MS = 15000;
const POLL_INTERVAL_HIDDEN_MS = 60000;
const MIN_WAKE_GAP_MS = 8000;

let pollTimer = null;
let inFlight = false;
let currentSyncRev = 1;
let lastPollTimestamp = 0;
let deltaListener = null;
let previousStatusSnapshot = new Map();
let lifecycleBound = false;

function buildStatusMap(list = []) {
  const map = new Map();
  if (!Array.isArray(list)) return map;
  for (const item of list) {
    if (!item?.id) continue;
    map.set(item.id, {
      status: item.verificationStatus || 'unverified',
      accStatus: item.accStatus || '',
      item
    });
  }
  return map;
}

export function seedOutgoingStatusBaseline(list = []) {
  previousStatusSnapshot = buildStatusMap(list);
}

export function detectOutgoingAccTransitions(prevMap, nextList = []) {
  const approved = [];
  const rejected = [];
  const newRequests = [];

  for (const item of (Array.isArray(nextList) ? nextList : [])) {
    if (!item?.id) continue;
    const prev = prevMap.get(item.id);
    const nextStatus = item.verificationStatus || 'unverified';
    const prevStatus = prev?.status || '';

    if (prev && prevStatus === 'pending_acc' && nextStatus === 'sah') {
      approved.push(item);
    } else if (prev && prevStatus === 'pending_acc' && nextStatus === 'rejected') {
      rejected.push(item);
    } else if (nextStatus === 'pending_acc' && (!prev || prevStatus !== 'pending_acc')) {
      newRequests.push(item);
    }
  }

  return { approved, rejected, newRequests };
}

export function initLiveSync({ onDeltaApplied } = {}) {
  deltaListener = typeof onDeltaApplied === 'function' ? onDeltaApplied : null;
  if (!lifecycleBound && typeof window !== 'undefined' && typeof document !== 'undefined') {
    lifecycleBound = true;
    const wakePoll = () => {
      if (document.visibilityState === 'visible' && getStoredAuthToken()) {
        if (Date.now() - lastPollTimestamp >= MIN_WAKE_GAP_MS) {
          forcePollNow();
        }
      }
    };
    document.addEventListener('visibilitychange', wakePoll);
    window.addEventListener('focus', wakePoll);
    window.addEventListener('online', wakePoll);
  }
}

export function setLastKnownSyncRev(rev) {
  const num = Number(rev);
  if (Number.isFinite(num) && num > 0) {
    currentSyncRev = Math.max(currentSyncRev, num);
  }
}

export function getLastKnownSyncRev() {
  return currentSyncRev;
}

function scheduleNextPoll() {
  clearTimeout(pollTimer);
  pollTimer = null;
  if (typeof window === 'undefined') return;
  if (!getStoredAuthToken()) return;
  const delay = typeof document !== 'undefined' && document.visibilityState === 'hidden'
    ? POLL_INTERVAL_HIDDEN_MS
    : POLL_INTERVAL_VISIBLE_MS;
  pollTimer = setTimeout(() => {
    forcePollNow();
  }, delay);
}

export function startLiveSync(initialRev = null) {
  stopLiveSync();
  const st = store.getState();
  if (initialRev) setLastKnownSyncRev(initialRev);
  else if (st.syncRev) setLastKnownSyncRev(st.syncRev);
  seedOutgoingStatusBaseline(st.outgoingLetters || []);
  lastPollTimestamp = Date.now();
  scheduleNextPoll();
}

export function stopLiveSync() {
  clearTimeout(pollTimer);
  pollTimer = null;
  inFlight = false;
  currentSyncRev = 1;
  lastPollTimestamp = 0;
  previousStatusSnapshot = new Map();
}

export async function forcePollNow({ lite = true } = {}) {
  if (inFlight || typeof window === 'undefined') return false;
  if (!getStoredAuthToken() || !store.getState().authUser) {
    stopLiveSync();
    return false;
  }

  inFlight = true;
  lastPollTimestamp = Date.now();
  try {
    const unitId = store.getActiveUnitId?.() || getActiveWorkspaceUnitId();
    const res = await pollWorkspaceUpdates(currentSyncRev, { lite, unitId });
    if (res?.ok) {
      if (res.rev) setLastKnownSyncRev(res.rev);
      if (res.changed) {
        const prevMap = previousStatusSnapshot.size > 0
          ? previousStatusSnapshot
          : buildStatusMap(store.getState().outgoingLetters || []);
        const nextOutgoing = res.workspace?.outgoingLetters || store.getState().outgoingLetters || [];
        const transitions = detectOutgoingAccTransitions(prevMap, nextOutgoing);
        seedOutgoingStatusBaseline(nextOutgoing);

        store.applyLiveWorkspaceDelta(res);

        if (deltaListener) {
          deltaListener({
            rev: res.rev,
            lastEvent: res.lastEvent || 'update',
            pendingAccCount: Number(res.pendingAccCount) || 0,
            transitions
          });
        }
        return true;
      }
    }
    return false;
  } catch {
    return false;
  } finally {
    inFlight = false;
    scheduleNextPoll();
  }
}
