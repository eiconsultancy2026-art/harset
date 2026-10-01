'use strict';

/**
 * Canonical Schedule III data helpers.
 * The UI, merge endpoint and Excel validator all use the same rules here.
 */

const NOTE_KEY_RE = /^\s*Note\s*(\d+)\s*-/i;

function toFiniteNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  if (typeof value !== 'string') return null;
  const cleaned = value.replace(/₹/g, '').replace(/,/g, '').trim();
  if (!cleaned) return null;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

function noteMap(data) {
  const result = {};
  const mapping = data && typeof data.mapping === 'object' ? data.mapping : {};
  for (const [key, value] of Object.entries(mapping)) {
    const match = String(key).match(NOTE_KEY_RE);
    if (!match) continue;
    const n = Number(match[1]);
    const parsed = toFiniteNumber(value);
    if (Number.isInteger(n) && parsed !== null) result[n] = parsed;
  }
  return result;
}

function getNote(data, note) {
  const map = noteMap(data);
  return Object.prototype.hasOwnProperty.call(map, note) ? map[note] : null;
}

function canonicalCompanyName(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/\s*\(\s*from\s+[^)]*\)\s*$/i, '')
    .replace(/\s*-\s*$/, '')
    .trim()
    .toLowerCase();
}

function displayCompanyName(value) {
  return String(value || '').trim().replace(/\s+/g, ' ');
}

function canonicalCompanyType(value) {
  const type = String(value || '').trim().toLowerCase();
  if (type === 'private limited' || type === 'private limited company') return 'Private Limited';
  if (type === 'partnership firm' || type === 'partnership') return 'Partnership Firm';
  if (type === 'proprietorship' || type === 'proprietor') return 'Proprietorship';
  if (type === 'huf') return 'HUF';
  if (type === 'rwa') return 'RWA';
  if (type === 'trust') return 'Trust';
  return String(value || '').trim();
}

function inferCompanyType(data, explicitType) {
  if (explicitType) return canonicalCompanyType(explicitType);
  return canonicalCompanyType(data && data.companyType);
}

function expectedNoteCount(companyType) {
  if (companyType === 'Private Limited') return 18;
  return companyType === 'RWA' || companyType === 'Trust' ? 21 : 25;
}

function snapshot(data, noteCount) {
  const notes = {};
  let nonZero = 0;
  for (let n = 1; n <= noteCount; n++) {
    const value = getNote(data, n);
    notes[n] = value;
    if (value !== null && Math.abs(value) > 0.000001) nonZero++;
  }
  return { notes, nonZero };
}

function validateMappingPayload(data, label, companyType) {
  const errors = [];
  if (!data || typeof data !== 'object') return [`${label} mapping is missing.`];
  const company = displayCompanyName(data.company);
  if (!company) errors.push(`${label} mapping does not contain a company name.`);
  const type = inferCompanyType(data, companyType);
  if (!type) errors.push(`${label} mapping does not contain a company type.`);
  const count = expectedNoteCount(type);
  const notes = snapshot(data, count).notes;
  for (let n = 1; n <= count; n++) {
    if (notes[n] === null) errors.push(`${label}: Note ${n} is missing or non-numeric.`);
  }
  return errors;
}

module.exports = {
  toFiniteNumber,
  noteMap,
  getNote,
  canonicalCompanyName,
  displayCompanyName,
  canonicalCompanyType,
  inferCompanyType,
  expectedNoteCount,
  snapshot,
  validateMappingPayload
};
