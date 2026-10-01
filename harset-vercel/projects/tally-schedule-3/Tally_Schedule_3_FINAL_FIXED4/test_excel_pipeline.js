'use strict';
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { validateGeneratedWorkbook } = require('./excel_validator');

async function main() {
  const dir = path.join(__dirname, 'cache');
  fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, '__pipeline_validation.xlsx');

  const makeData = (company, values) => ({
    company,
    companyType: 'Partnership Firm',
    mapping: Object.fromEntries(values.map((v, i) => [`Note ${i + 1} - Test`, v]))
  });
  const left = makeData('Demo Co', Array.from({ length: 25 }, (_, i) => i + 100));
  const right = makeData('Demo Co', Array.from({ length: 25 }, (_, i) => i + 50));

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Balance Sheet');
  ws.addRow(['placeholder']);
  const audit = wb.addWorksheet('Tally Data Audit');
  audit.addRow(['Note', 'Particular', '31 Mar 2026', '31 Mar 2025']);
  for (let n = 1; n <= 25; n++) audit.addRow([`Note ${n}`, 'Schedule III total', left.mapping[`Note ${n} - Test`], right.mapping[`Note ${n} - Test`]]);
  await wb.xlsx.writeFile(file);

  const result = await validateGeneratedWorkbook(file, left, right, 2026, 2025, 'Partnership Firm');
  if (!result.valid || result.noteCount !== 25) throw new Error('Pipeline validator did not pass the valid workbook.');

  // Regression: change only the previous-year value and ensure validation fails.
  const bad = new ExcelJS.Workbook();
  await bad.xlsx.readFile(file);
  bad.getWorksheet('Tally Data Audit').getCell(2, 4).value = 0;
  await bad.xlsx.writeFile(file);
  let failed = false;
  try { await validateGeneratedWorkbook(file, left, right, 2026, 2025, 'Partnership Firm'); }
  catch (e) { failed = /previous year 2025/.test(e.message); }
  if (!failed) throw new Error('Regression test failed: previous-year mismatch was not detected.');

  fs.unlinkSync(file);
  console.log('PASS: Excel pipeline validation catches current and previous year mismatches.');
}

main().catch(err => { console.error(err); process.exit(1); });
