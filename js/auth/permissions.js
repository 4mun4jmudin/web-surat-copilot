/**
 * js/auth/permissions.js - Pembantu Otorisasi RBAC & Delegasi Pengesahan Legal
 * Mengatur hak akses Administrator, Kepala Sekolah, dan Tata Usaha (TU).
 */

import { SIGNER_ROLE_OPTIONS } from '../letters/verification.js';

export const ROLE_LABELS = Object.freeze({
  admin: 'Administrator',
  kepsek: 'Kepala Sekolah',
  tu: 'Tata Usaha (TU)'
});

export const DEFAULT_ACCESS_PERMISSIONS = Object.freeze({
  allowTuSignPrincipal: false,
  allowTuSignKepsekDanTu: false,
  allowTuRevokePrincipalQr: false,
  allowTuManageApiVault: true,
  allowTuDeleteOutgoing: false,
  allowTuDeleteIncoming: true,
  updatedAt: '',
  updatedBy: 'system',
  updatedByRole: 'admin',
  note: ''
});

export function getRoleLabel(role) {
  return ROLE_LABELS[String(role || '').toLowerCase()] || 'Pengguna';
}

export function isAdminUser(user) {
  return Boolean(user && user.role === 'admin' && user.active !== false);
}

export function isKepsekUser(user) {
  return Boolean(user && user.role === 'kepsek' && user.active !== false);
}

export function canManageUsers(user) {
  return isAdminUser(user);
}

export function canManageDelegation(user) {
  return isAdminUser(user) || isKepsekUser(user);
}

export function canSignWithRole(user, permissions = DEFAULT_ACCESS_PERMISSIONS, requestedRole = 'kepala_sekolah') {
  // Jika belum memakai mode login (unit test / sesi lokal tanpa auth), izinkan agar kompatibel
  if (!user) {
    return { allowed: true, delegated: false, reason: '' };
  }
  if (user.active === false) {
    return { allowed: false, delegated: false, reason: 'Akun dinonaktifkan oleh Administrator.' };
  }

  const normalized = ['kepala_sekolah', 'tata_usaha', 'kepsek_dan_tu'].includes(requestedRole)
    ? requestedRole
    : 'kepala_sekolah';

  if (user.role === 'admin' || user.role === 'kepsek') {
    return { allowed: true, delegated: false, reason: '' };
  }

  if (user.role === 'tu') {
    if (normalized === 'tata_usaha') {
      return { allowed: true, delegated: false, reason: '' };
    }
    if (normalized === 'kepala_sekolah') {
      return permissions?.allowTuSignPrincipal === true
        ? { allowed: true, delegated: true, reason: '' }
        : {
            allowed: false,
            delegated: false,
            reason: 'Terkunci: perlu izin persetujuan dari Kepala Sekolah atau Admin'
          };
    }
    if (normalized === 'kepsek_dan_tu') {
      return (permissions?.allowTuSignKepsekDanTu === true || permissions?.allowTuSignPrincipal === true)
        ? { allowed: true, delegated: true, reason: '' }
        : {
            allowed: false,
            delegated: false,
            reason: 'Terkunci: perlu izin delegasi Kepsek / Admin'
          };
    }
  }

  return { allowed: false, delegated: false, reason: 'Peran tidak memiliki akses pengesahan.' };
}

export function getSelectableSignerRoles(user, permissions = DEFAULT_ACCESS_PERMISSIONS, roleOptions = SIGNER_ROLE_OPTIONS) {
  return roleOptions.map(opt => {
    const check = canSignWithRole(user, permissions, opt.id);
    return {
      ...opt,
      allowed: check.allowed,
      delegated: check.delegated,
      lockReason: check.reason
    };
  });
}

export function resolveDefaultAllowedSignerRole(user, permissions = DEFAULT_ACCESS_PERMISSIONS, preferredRole = 'kepala_sekolah') {
  const options = getSelectableSignerRoles(user, permissions);
  const preferred = options.find(o => o.id === preferredRole && o.allowed);
  if (preferred) return preferred.id;
  const firstAllowed = options.find(o => o.allowed);
  return firstAllowed?.id || 'tata_usaha';
}

export function canRevokeOutgoingItem(user, permissions = DEFAULT_ACCESS_PERMISSIONS, item = null) {
  if (!user) return true;
  if (user.active === false) return false;
  if (user.role === 'admin' || user.role === 'kepsek') return true;
  if (user.role === 'tu') {
    if (item?.signerRole === 'tata_usaha') return true;
    return permissions?.allowTuRevokePrincipalQr === true;
  }
  return false;
}

export function canDeleteOutgoingItem(user, permissions = DEFAULT_ACCESS_PERMISSIONS) {
  if (!user) return true;
  if (user.active === false) return false;
  if (user.role === 'admin' || user.role === 'kepsek') return true;
  if (user.role === 'tu') return permissions?.allowTuDeleteOutgoing === true;
  return false;
}

export function canDeleteIncomingItem(user, permissions = DEFAULT_ACCESS_PERMISSIONS) {
  if (!user) return true;
  if (user.active === false) return false;
  if (user.role === 'admin' || user.role === 'kepsek') return true;
  if (user.role === 'tu') return permissions?.allowTuDeleteIncoming !== false;
  return false;
}

export function canManageApiVault(user, permissions = DEFAULT_ACCESS_PERMISSIONS) {
  if (!user) return true;
  if (user.active === false) return false;
  if (user.role === 'admin' || user.role === 'kepsek') return true;
  if (user.role === 'tu') return permissions?.allowTuManageApiVault !== false;
  return false;
}

export function isMobileOrPhoneEnvironment({ innerWidth, userAgent } = {}) {
  const width = typeof innerWidth === 'number'
    ? innerWidth
    : (typeof window !== 'undefined' ? window.innerWidth : 1280);
  const ua = typeof userAgent === 'string'
    ? userAgent
    : (typeof navigator !== 'undefined' ? navigator.userAgent : '');
  const isPhoneUa = /Android.*Mobile|webOS|iPhone|iPod|BlackBerry|IEMobile|Opera Mini/i.test(ua);
  return width <= 860 || isPhoneUa;
}

export function isCanvasLockedForUser(user, env = undefined) {
  if (!user || user.role !== 'kepsek') return false;
  return isMobileOrPhoneEnvironment(env);
}

export function canApproveOrRejectAcc(user) {
  if (!user || user.active === false) return false;
  return user.role === 'admin' || user.role === 'kepsek';
}

export function getPendingAccLetters(outgoingLetters = []) {
  if (!Array.isArray(outgoingLetters)) return [];
  return outgoingLetters.filter(item => item?.verificationStatus === 'pending_acc' || item?.accStatus === 'pending');
}

export function canSwitchUnits(user) {
  if (!user || user.active === false) return false;
  return String(user.unitId || '').toLowerCase() === 'all';
}

export function getUnitDisplayName(unitId = 'smk') {
  return String(unitId || '').toLowerCase() === 'smp'
    ? 'SMP IT Al-Hawari'
    : 'SMK IT Al-Hawari';
}
