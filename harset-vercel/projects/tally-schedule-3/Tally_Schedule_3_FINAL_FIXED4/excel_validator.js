'use strict';

const ExcelJS = require('exceljs');
const {
  getNote,
  expectedNoteCount,
  toFiniteNumber,
  inferCompanyType
} = require('./schedule3_model');

const EPSILON = 0.01;

function numericCell(cell) {
  const value = cell?.value;
  if (value && typeof value === 'object' && Object.prototype.hasOwnProperty.call(value, 'result')) {
    return toFiniteNumber(value.result);
  }
  return toFiniteNumber(value);
}

function findAuditSheet(workbook) {
  return workbook.getWorksheet('Tally Data Audit') ||
    workbook.worksheets.find(ws => /tally\s+data\s+audit/i.test(ws.name));
}

function readAuditNotes(ws, noteCount) {
  const found = new Map();
  for (let r = 1; r <= ws.rowCount; r++) {
    const first = String(ws.getCell(r, 1).value ?? '').trim();
    const m = first.match(/^Note\s+(\d+)$/i);
    if (!m) continue;
    const note = Number(m[1]);
    if (!Number.isInteger(note) || note < 1 || note > noteCount) continue;

    // Generators use either C/D or D/E for current/previous values.
    const candidates = [
      [3, 4], [4, 5], [2, 3]
    ];
    let pair = null;
    let bestScore = -1;
    for (const [a, b] of candidates) {
      const av = numericCell(ws.getCell(r, a));
      const bv = numericCell(ws.getCell(r, b));
      const score = (av !== null ? 1 : 0) + (bv !== null ? 1 : 0);
      if (score > bestScore) {
        pair = [av ?? 0, bv ?? 0];
        bestScore = score;
      }
    }
    if (!pair) continue;
    found.set(note, pair);
  }
  return found;
}

function compare(expected, actual, note, side, errors) {
  if (expected === null || actual === null) {
    errors.push(`Note ${note} ${side}: missing numeric value (expected=${expected}, excel=${actual})`);
    return;
  }
  if (Math.abs(expected - actual) > EPSILON) {
    errors.push(`Note ${note} ${side}: expected ${expected}, Excel contains ${actual}`);
  }
}

async function validateGeneratedWorkbook(filePath, leftData, rightData, leftYear, rightYear, companyType) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const type = inferCompanyType(leftData, companyType);
  const noteCount = expectedNoteCount(type);
  const audit = findAuditSheet(workbook);
  if (!audit) throw new Error('Generated workbook is missing the required "Tally Data Audit" sheet.');

  const actual = readAuditNotes(audit, noteCount);
  const errors = [];
  for (let n = 1; n <= noteCount; n++) {
    const expectedLeft = getNote(leftData, n);
    const expectedRight = getNote(rightData, n);
    const pair = actual.get(n);
    if (!pair) {
      errors.push(`Note ${n}: missing from Excel audit sheet.`);
      continue;
    }
    compare(expectedLeft, pair[0], n, `current year ${leftYear}`, errors);
    compare(expectedRight, pair[1], n, `previous year ${rightYear}`, errors);
  }

  // The audit sheet must cover every note exactly once. Duplicate rows can
  // otherwise make a broken workbook appear valid.
  const seen = new Set();
  for (let r = 1; r <= audit.rowCount; r++) {
    const m = String(audit.getCell(r, 1).value ?? '').trim().match(/^Note\s+(\d+)$/i);
    if (m) {
      const n = Number(m[1]);
      if (n >= 1 && n <= noteCount) seen.add(n);
    }
  }
  if (seen.size !== noteCount) {
    errors.push(`Audit coverage incomplete: found ${seen.size}/${noteCount} notes.`);
  }

  if (errors.length) {
    throw new Error(`Excel validation failed. The generated workbook does not match the portal data:\n${errors.join('\n')}`);
  }

  return {
    valid: true,
    noteCount,
    companyType: type,
    years: { current: leftYear, previous: rightYear }
  };
}

module.exports = { validateGeneratedWorkbook, readAuditNotes };
