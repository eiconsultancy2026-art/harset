const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// Maps company type to a template xlsx file (relative to project root)
const TEMPLATE_TYPES = {
    'Partnership Firm': 'Format for Properiotorship and Partnership.xlsx',
    'Proprietorship':   'Format for Properiotorship and Partnership.xlsx',
    'RWA':              'Format for trust and RWAs.xlsx',
    'Trust':            'Format for trust and RWAs.xlsx',
};

function getTemplatePath(companyType) {
    const file = TEMPLATE_TYPES[companyType];
    if (!file) return null;
    return path.join(__dirname, file);
}

async function generateFromTemplate(templatePath, leftName, leftData, rightName, rightData, outputPath, leftYear, rightYear) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(templatePath);

    const y1 = leftYear  || new Date().getFullYear();
    const y2 = rightYear || (parseInt(y1) - 1);
    const NUM_FMT = '[>=10000000]##\\,##\\,##\\,##0.00;[>=100000]##\\,##\\,##0.00;##,##0.00';

    function getNote(data, n) {
        if (!data || !data.mapping) return 0;
        const key = Object.keys(data.mapping).find(k => new RegExp(`^\\s*Note\\s*${n}\\s*-`, 'i').test(String(k)));
        if (key) { const v = Number(data.mapping[key]); return Number.isFinite(v) ? v : 0; }
        return 0;
    }

    // Walk every cell in every sheet: replace placeholder tokens with data
    // Placeholder convention in the template:
    //   {L1}..{L18}  → current-year Note values
    //   {R1}..{R18}  → previous-year Note values
    //   {LY}         → current year label
    //   {RY}         → previous year label
    //   {LNAME}      → left company name
    //   {RNAME}      → right company name
    wb.eachSheet(ws => {
        ws.eachRow(row => {
            row.eachCell({ includeEmpty: false }, cell => {
                const v = cell.value;
                if (typeof v !== 'string') return;

                let replaced = v
                    .replace(/\{LY\}/g, y1)
                    .replace(/\{RY\}/g, y2)
                    .replace(/\{LNAME\}/g, leftName)
                    .replace(/\{RNAME\}/g, rightName);

                // Note substitutions: {L14} etc.
                replaced = replaced.replace(/\{L(\d+)\}/g, (_, n) => {
                    const val = getNote(leftData, parseInt(n));
                    cell.numFmt = NUM_FMT;
                    return val;
                });
                replaced = replaced.replace(/\{R(\d+)\}/g, (_, n) => {
                    const val = getNote(rightData, parseInt(n));
                    cell.numFmt = NUM_FMT;
                    return val;
                });

                // If the whole cell became a plain number string, convert to number
                if (replaced !== v) {
                    const asNum = Number(replaced);
                    cell.value = isNaN(asNum) || replaced.trim() === '' ? replaced : asNum;
                    if (typeof cell.value === 'number') {
                        cell.numFmt = NUM_FMT;
                    }
                }
            });
        });
    });

    // Add an audit sheet so no Tally value disappears silently.
    const audit = wb.addWorksheet('Tally Data Audit');
    audit.columns = [
        { width: 10 }, { width: 42 }, { width: 28 }, { width: 20 }, { width: 20 }, { width: 45 }
    ];
    audit.addRow(['Note', 'Particular', 'Source / Variable', `31 Mar ${y1}`, `31 Mar ${y2}`, 'Status']).font = { bold: true };
    function addAudit(note, label, source, lv, rv) {
        const l = Number(lv) || 0, r = Number(rv) || 0;
        const status = (l !== 0 || r !== 0) ? 'Mapped from Tally' : 'Zero / not returned for this mapping';
        audit.addRow([note, label, source, l, r, status]);
    }
    for (let n = 1; n <= 18; n++) {
        addAudit(n, `Schedule III Note ${n}`, `mapping[Note ${n}]`, getNote(leftData,n), getNote(rightData,n));
    }
    const breakdownGroups = [
        ['Note 3 - Secured Loans', leftData?.loansInfo?.secured, rightData?.loansInfo?.secured],
        ['Note 3 - Unsecured Loans', leftData?.loansInfo?.unsecured, rightData?.loansInfo?.unsecured],
        ['Note 6 - Trade Payables', leftData?.sundryCredInfo?.ledgers, rightData?.sundryCredInfo?.ledgers],
        ['Note 11 - Sundry Debtors', leftData?.sundryDebtorsInfo?.ledgers, rightData?.sundryDebtorsInfo?.ledgers],
        ['Note 10 - Investments', leftData?.investmentsInfo?.ledgers, rightData?.investmentsInfo?.ledgers],
        ['Note 15 - Other Income', leftData?.indirectIncomesInfo?.items, rightData?.indirectIncomesInfo?.items],
        ['Note 16 - Employee Benefits', leftData?.note16Info?.items, rightData?.note16Info?.items]
    ];
    for (const [label, li, ri] of breakdownGroups) {
        const keys = [...new Set([...Object.keys(li || {}), ...Object.keys(ri || {})])].sort();
        for (const k of keys) addAudit(label, k, 'Detailed Tally response', li?.[k] || 0, ri?.[k] || 0);
    }
    audit.eachRow(row => row.eachCell(cell => { if (typeof cell.value === 'number') cell.numFmt = NUM_FMT; }));

    await wb.xlsx.writeFile(outputPath);
    console.log(`[Template] Saved to ${outputPath}`);
}

/**
 * RWA / Trust merged Excel:
 *   Sheet 1 – Balance Sheet  (from template)
 *   Sheet 2 – Statement of P&L  (from template)
 *   Sheet 3 – Notes 1 to 3
 *   Sheet 4 – Note 4 Borrowings
 *   Sheet 5 – Notes 5 to 8
 *   Sheet 6 – Note 9 PPE
 *   Sheet 7 – Notes 10 to 15
 *   Sheet 8 – Notes 16 to 21
 */
async function generateMergedExcelRwaTrust(leftName, leftData, rightName, rightData, outputPath, leftYear, rightYear) {
    const templatePath = path.join(__dirname, 'Format for trust and RWAs.xlsx');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(templatePath);

    const y1 = leftYear  || new Date().getFullYear();
    const y2 = rightYear || (parseInt(y1) - 1);
    const NUM_FMT = '[>=10000000]##\\,##\\,##\\,##0.00;[>=100000]##\\,##\\,##0.00;##,##0.00';
    const THIN   = { style: 'thin', color: { argb: 'FF000000' } };
    const BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };

    // Helper: find Note N total from mapping
    function N(data, n) {
        if (!data || !data.mapping) return 0;
        const key = Object.keys(data.mapping).find(k => k.startsWith(`Note ${n} -`) || k.startsWith(`Note ${n}  -`));
        return key ? (data.mapping[key] || 0) : 0;
    }

    // Helper: union of keys from two objects
    function keys(a, b) {
        return [...new Set([...Object.keys(a || {}), ...Object.keys(b || {})])].sort();
    }

    // Fill {LY},{RY},{LNAME},{RNAME},{L1}..{L21},{R1}..{R21} in template sheets
    wb.eachSheet(ws => {
        ws.eachRow(row => {
            row.eachCell({ includeEmpty: false }, cell => {
                if (typeof cell.value !== 'string') return;
                let v = cell.value
                    .replace(/\{LY\}/g, y1).replace(/\{RY\}/g, y2)
                    .replace(/\{LNAME\}/g, leftName).replace(/\{RNAME\}/g, rightName);
                v = v.replace(/\{L(\d+)\}/g, (_, n) => N(leftData,  parseInt(n)));
                v = v.replace(/\{R(\d+)\}/g, (_, n) => N(rightData, parseInt(n)));
                if (v !== cell.value) {
                    const num = Number(v);
                    cell.value = (!isNaN(num) && v.trim() !== '') ? num : v;
                    if (typeof cell.value === 'number') cell.numFmt = NUM_FMT;
                }
            });
        });
    });

    function addSheet(name) {
        const existing = wb.getWorksheet(name);
        if (existing) wb.removeWorksheet(existing.id);
        return wb.addWorksheet(name);
    }

    function hdr(ws, title, cols) {
        ws.columns = cols;
        const r1 = ws.addRow([`NOTES TO THE FINANCIAL STATEMENTS FOR THE YEAR ENDED 31 MARCH ${y1}`]);
        ws.mergeCells(1, 1, 1, cols.length);
        r1.font = { bold: true, name: 'Times New Roman', size: 11 };
        r1.getCell(1).alignment = { horizontal: 'center' };
        ws.addRow([]);
        const r3 = ws.addRow([title]);
        r3.font = { bold: true, name: 'Times New Roman', size: 12 };
        ws.mergeCells(3, 1, 3, cols.length);
        ws.addRow([]);
    }

    function colHdr(ws, colCount) {
        const labels = colCount === 2
            ? ['Particulars', `As at 31 March ${y1} Rs.`, `As at 31 March ${y2} Rs.`]
            : ['Particulars', `As at 31 March ${y1} Rs.`, `As at 31 March ${y2} Rs.`];
        const r = ws.addRow(labels);
        r.font = { bold: true, name: 'Times New Roman' };
        r.eachCell(c => { c.border = BORDER; c.alignment = { horizontal: 'center', wrapText: true }; });
        return r;
    }

    function dataRow(ws, label, vL, vR, bold) {
        const r = ws.addRow([label, vL || 0, vR || 0]);
        if (bold) r.font = { bold: true, name: 'Times New Roman' };
        r.getCell(2).numFmt = NUM_FMT; r.getCell(3).numFmt = NUM_FMT;
        r.eachCell({ includeEmpty: true }, c => c.border = BORDER);
        return r;
    }

    function sectionTitle(ws, text) {
        const r = ws.addRow([text]);
        r.font = { bold: true, name: 'Times New Roman', size: 12 };
        ws.addRow([]);
    }

    // ── Sheet 3: Notes 1, 2, 3 ─────────────────────────────────────────────
    const ws3 = addSheet('Note 1 to 3');
    hdr(ws3, 'Notes 1, 2 and 3', [{ width: 50 }, { width: 22 }, { width: 22 }]);

    ws3.addRow(['Note 1 – Brief about the Entity']).font = { bold: true, size: 12 };
    ws3.addRow(['(Narrative – see attached disclosure)']).font = { italic: true };
    ws3.addRow([]);

    ws3.addRow(['Note 2 – Significant Accounting Policies']).font = { bold: true, size: 12 };
    ws3.addRow(['(Narrative – see attached disclosure)']).font = { italic: true };
    ws3.addRow([]);

    ws3.addRow(['Note 3 – NPOs Funds']).font = { bold: true, size: 12 };
    colHdr(ws3, 2);

    // NPOs funds breakdown (from summaryGroups / Capital Account)
    const n3L = N(leftData, 3),  n3R = N(rightData, 3);
    const capL = leftData?.summaryGroups?.['Capital Account']  || 0;
    const capR = rightData?.summaryGroups?.['Capital Account'] || 0;
    const surpL = leftData?.summaryGroups?.['Excess of income over expenditure'] || 0;
    const surpR = rightData?.summaryGroups?.['Excess of income over expenditure'] || 0;
    dataRow(ws3, 'Capital Account / Opening Balance', capL, capR);
    dataRow(ws3, 'Add: Excess of Income over Expenditure', surpL, surpR);
    dataRow(ws3, 'Total NPOs Funds', n3L, n3R, true);

    // ── Sheet 4: Note 4 – Borrowings ───────────────────────────────────────
    const ws4 = addSheet('Notes to BS 4');
    hdr(ws4, 'Note 4 – Borrowings', [{ width: 50 }, { width: 22 }, { width: 22 }]);
    colHdr(ws4, 2);

    ws4.addRow(['(A) Secured']).font = { bold: true };
    keys(leftData?.loansInfo?.secured, rightData?.loansInfo?.secured).forEach(k => {
        dataRow(ws4, `  ${k}`, leftData?.loansInfo?.secured?.[k] || 0, rightData?.loansInfo?.secured?.[k] || 0);
    });
    dataRow(ws4, 'Total Secured', leftData?.loansInfo?.secTotal || 0, rightData?.loansInfo?.secTotal || 0, true);
    ws4.addRow([]);
    ws4.addRow(['(B) Unsecured']).font = { bold: true };
    keys(leftData?.loansInfo?.unsecured, rightData?.loansInfo?.unsecured).forEach(k => {
        dataRow(ws4, `  ${k}`, leftData?.loansInfo?.unsecured?.[k] || 0, rightData?.loansInfo?.unsecured?.[k] || 0);
    });
    dataRow(ws4, 'Total Unsecured', leftData?.loansInfo?.unsecTotal || 0, rightData?.loansInfo?.unsecTotal || 0, true);
    ws4.addRow([]);
    dataRow(ws4, 'Total Borrowings (A + B)', N(leftData, 4), N(rightData, 4), true);

    // ── Sheet 5: Notes 5, 6, 7, 8 ─────────────────────────────────────────
    const ws5 = addSheet('Notes to BS 5 to 8');
    hdr(ws5, 'Notes 5 to 8 – Liabilities', [{ width: 50 }, { width: 22 }, { width: 22 }]);

    // Note 5 – Other Long-term Liabilities
    sectionTitle(ws5, 'Note 5 – Other Long-term Liabilities');
    colHdr(ws5, 2);
    dataRow(ws5, 'Total Other Long-term Liabilities', N(leftData, 5), N(rightData, 5), true);
    ws5.addRow([]);

    // Note 6 – Provisions
    sectionTitle(ws5, 'Note 6 – Provisions');
    colHdr(ws5, 2);
    keys(leftData?.provisionsInfo?.shortTermItems, rightData?.provisionsInfo?.shortTermItems).forEach(k => {
        dataRow(ws5, `  ${k}`,
            leftData?.provisionsInfo?.shortTermItems?.[k] || 0,
            rightData?.provisionsInfo?.shortTermItems?.[k] || 0);
    });
    const ptL = leftData?.provisionsInfo?.provisionForTaxation || 0;
    const ptR = rightData?.provisionsInfo?.provisionForTaxation || 0;
    if (ptL || ptR) dataRow(ws5, '  Provision for Taxation', ptL, ptR);
    dataRow(ws5, 'Total Provisions', N(leftData, 6), N(rightData, 6), true);
    ws5.addRow([]);

    // Note 7 – Payables
    sectionTitle(ws5, 'Note 7 – Payables (Trade Payables)');
    colHdr(ws5, 2);
    keys(leftData?.sundryCredInfo?.ledgers, rightData?.sundryCredInfo?.ledgers).forEach(k => {
        dataRow(ws5, `  ${k}`,
            leftData?.sundryCredInfo?.ledgers?.[k] || 0,
            rightData?.sundryCredInfo?.ledgers?.[k] || 0);
    });
    dataRow(ws5, 'Total Payables', N(leftData, 7), N(rightData, 7), true);
    ws5.addRow([]);

    // Note 8 – Other Current Liabilities
    sectionTitle(ws5, 'Note 8 – Other Current Liabilities');
    colHdr(ws5, 2);
    keys(leftData?.otherCurrLiabInfo?.items, rightData?.otherCurrLiabInfo?.items).forEach(k => {
        dataRow(ws5, `  ${k}`,
            leftData?.otherCurrLiabInfo?.items?.[k] || 0,
            rightData?.otherCurrLiabInfo?.items?.[k] || 0);
    });
    dataRow(ws5, 'Total Other Current Liabilities', Math.abs(N(leftData, 8)), Math.abs(N(rightData, 8)), true);

    // ── Sheet 6: Note 9 – PPE ──────────────────────────────────────────────
    const ws6 = addSheet('Notes to BS 9');
    hdr(ws6, 'Note 9 – Property, Plant & Equipment (PPE)', [{ width: 50 }, { width: 22 }, { width: 22 }]);
    colHdr(ws6, 2);
    dataRow(ws6, 'Gross Block (Opening Balance)', 0, 0);
    dataRow(ws6, 'Additions during the year', N(leftData, 9), N(rightData, 9));
    dataRow(ws6, 'Disposals during the year', 0, 0);
    dataRow(ws6, 'Gross Block (Closing Balance)', N(leftData, 9), N(rightData, 9));
    ws6.addRow([]);
    dataRow(ws6, 'Less: Accumulated Depreciation', N(leftData, 20), N(rightData, 20));
    dataRow(ws6, 'Net Block (WDV)', N(leftData, 9), N(rightData, 9), true);

    // ── Sheet 7: Notes 10–15 ──────────────────────────────────────────────
    const ws7 = addSheet('Notes to BS 10 to 15');
    hdr(ws7, 'Notes 10 to 15 – Assets', [{ width: 50 }, { width: 22 }, { width: 22 }]);

    // Note 10 – Investments
    sectionTitle(ws7, 'Note 10 – Investments (Non-current & Current)');
    colHdr(ws7, 2);
    keys(leftData?.investmentsInfo?.ledgers, rightData?.investmentsInfo?.ledgers).forEach(k => {
        dataRow(ws7, `  ${k}`,
            leftData?.investmentsInfo?.ledgers?.[k] || 0,
            rightData?.investmentsInfo?.ledgers?.[k] || 0);
    });
    dataRow(ws7, 'Total Investments', N(leftData, 10), N(rightData, 10), true);
    ws7.addRow([]);

    // Note 11 – Loans and Advances
    sectionTitle(ws7, 'Note 11 – Loans and Advances');
    colHdr(ws7, 2);
    keys(leftData?.advancesInfo?.items, rightData?.advancesInfo?.items).forEach(k => {
        dataRow(ws7, `  ${k}`,
            leftData?.advancesInfo?.items?.[k] || 0,
            rightData?.advancesInfo?.items?.[k] || 0);
    });
    dataRow(ws7, 'Total Loans and Advances', N(leftData, 11), N(rightData, 11), true);
    ws7.addRow([]);

    // Note 12 – Other Non-current Assets
    sectionTitle(ws7, 'Note 12 – Other Non-current Assets');
    colHdr(ws7, 2);
    dataRow(ws7, 'Total Other Non-current Assets', N(leftData, 12), N(rightData, 12), true);
    ws7.addRow([]);

    // Note 13 – Receivables
    sectionTitle(ws7, 'Note 13 – Receivables (Members / Trade Debtors)');
    colHdr(ws7, 2);
    keys(leftData?.sundryDebtorsInfo?.ledgers, rightData?.sundryDebtorsInfo?.ledgers).forEach(k => {
        dataRow(ws7, `  ${k}`,
            leftData?.sundryDebtorsInfo?.ledgers?.[k] || 0,
            rightData?.sundryDebtorsInfo?.ledgers?.[k] || 0);
    });
    dataRow(ws7, 'Total Receivables', N(leftData, 13), N(rightData, 13), true);
    ws7.addRow([]);

    // Note 14 – Cash and Bank Balance
    sectionTitle(ws7, 'Note 14 – Cash and Bank Balance');
    colHdr(ws7, 2);
    dataRow(ws7, 'Balances with Banks', leftData?.cashInfo?.bankAccounts || 0, rightData?.cashInfo?.bankAccounts || 0);
    dataRow(ws7, 'Cash on Hand',         leftData?.cashInfo?.cashInHand  || 0, rightData?.cashInfo?.cashInHand  || 0);
    dataRow(ws7, 'Total Cash and Bank Balance', N(leftData, 14), N(rightData, 14), true);
    ws7.addRow([]);

    // Note 15 – Other Current Assets
    sectionTitle(ws7, 'Note 15 – Other Current Assets');
    colHdr(ws7, 2);
    dataRow(ws7, 'Total Other Current Assets', N(leftData, 15), N(rightData, 15), true);

    // ── Sheet 8: Notes 16–21 ──────────────────────────────────────────────
    const ws8 = addSheet('Notes to P&L 16 to 21');
    hdr(ws8, 'Notes 16 to 21 – Income & Expenditure', [{ width: 50 }, { width: 22 }, { width: 22 }]);

    const fyHdr = (ws) => {
        const r = ws.addRow(['Particulars', `For year ended 31 Mar ${y1} Rs.`, `For year ended 31 Mar ${y2} Rs.`]);
        r.font = { bold: true, name: 'Times New Roman' };
        r.eachCell(c => { c.border = BORDER; c.alignment = { horizontal: 'center', wrapText: true }; });
    };

    // Note 16 – Other Income
    sectionTitle(ws8, 'Note 16 – Other Income');
    fyHdr(ws8);
    keys(leftData?.indirectIncomesInfo?.items, rightData?.indirectIncomesInfo?.items).forEach(k => {
        dataRow(ws8, `  ${k}`,
            leftData?.indirectIncomesInfo?.items?.[k] || 0,
            rightData?.indirectIncomesInfo?.items?.[k] || 0);
    });
    dataRow(ws8, 'Total Other Income', N(leftData, 16), N(rightData, 16), true);
    ws8.addRow([]);

    // Note 17 – Cost of Goods Sold
    sectionTitle(ws8, 'Note 17 – Cost of Goods Sold');
    fyHdr(ws8);
    dataRow(ws8, 'Total Cost of Goods Sold', N(leftData, 17), N(rightData, 17), true);
    ws8.addRow([]);

    // Note 18 – Employee Benefits Expense
    sectionTitle(ws8, 'Note 18 – Employee Benefits Expense');
    fyHdr(ws8);
    keys(leftData?.note16Info?.items, rightData?.note16Info?.items).forEach(k => {
        dataRow(ws8, `  ${k}`,
            leftData?.note16Info?.items?.[k] || 0,
            rightData?.note16Info?.items?.[k] || 0);
    });
    dataRow(ws8, 'Total Employee Benefits Expense', N(leftData, 18), N(rightData, 18), true);
    ws8.addRow([]);

    // Note 19 – Finance Cost
    sectionTitle(ws8, 'Note 19 – Finance Cost');
    fyHdr(ws8);
    dataRow(ws8, 'Total Finance Cost', N(leftData, 19), N(rightData, 19), true);
    ws8.addRow([]);

    // Note 20 – Depreciation and Amortisation
    sectionTitle(ws8, 'Note 20 – Depreciation and Amortisation Expense');
    fyHdr(ws8);
    dataRow(ws8, 'Depreciation on PPE', N(leftData, 20), N(rightData, 20));
    dataRow(ws8, 'Total Depreciation', N(leftData, 20), N(rightData, 20), true);
    ws8.addRow([]);

    // Note 21 – Other Expenses
    sectionTitle(ws8, 'Note 21 – Other Expenses');
    fyHdr(ws8);
    dataRow(ws8, 'Total Other Expenses', N(leftData, 21), N(rightData, 21), true);

    // ── Save ───────────────────────────────────────────────────────────────
    const auditRwa = wb.addWorksheet('Tally Data Audit');
    auditRwa.columns = [{width:20},{width:45},{width:22},{width:22}];
    auditRwa.addRow(['Note','Particular',`31 Mar ${y1}`,`31 Mar ${y2}`]).font={bold:true};
    for(let n=1;n<=21;n++) auditRwa.addRow([`Note ${n}`,'Schedule III total',N(leftData,n),N(rightData,n)]);
    auditRwa.eachRow(r=>r.eachCell(c=>{if(typeof c.value==='number')c.numFmt=NUM_FMT;}));
    await wb.xlsx.writeFile(outputPath);
    console.log(`[RWA/Trust] Saved merged Excel to ${outputPath}`);
}

async function generateMergedExcel(leftName, leftData, rightName, rightData, outputPath, leftYear, rightYear, companyType, leftPvtInfo, rightPvtInfo) {
    // Route to template-based generation for Prop/Partnership/RWA/Trust
    const templatePath = getTemplatePath(companyType);
    if (templatePath && fs.existsSync(templatePath)) {
        console.log(`[Excel] Using template: ${path.basename(templatePath)} for type: ${companyType}`);
        return generateFromTemplate(templatePath, leftName, leftData, rightName, rightData, outputPath, leftYear, rightYear);
    }

    // ── Fallback: full programmatic generation (HUF / Private Limited / unknown) ──
    const workbook = new ExcelJS.Workbook();
    
    const y1 = leftYear || new Date().getFullYear();
    const y2 = rightYear || (parseInt(y1) - 1);

    const NUM_FMT = '[>=10000000]##\\,##\\,##\\,##0.00;[>=100000]##\\,##\\,##0.00;##,##0.00';
    const BORDER_THIN = { style: 'thin', color: { argb: 'FF000000' } };
    
    function applyBorders(row, cols) {
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            if (cols && !cols.includes(colNumber)) return;
            cell.border = { top: BORDER_THIN, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN };
        });
    }

    function getNote(data, n) {
        if (!data || !data.mapping) return 0;
        const key = Object.keys(data.mapping).find(k => new RegExp(`^\\s*Note\\s*${n}\\s*-`, 'i').test(String(k)));
        if (key) { const v = Number(data.mapping[key]); return Number.isFinite(v) ? v : 0; }
        return 0;
    }
    
    function getKeys(obj1, obj2) {
        const keys = new Set([...Object.keys(obj1 || {}), ...Object.keys(obj2 || {})]);
        return Array.from(keys).sort();
    }


    // 2. BS
    const wsBS = workbook.addWorksheet('BS');
    wsBS.columns = [{ width: 8 }, { width: 45 }, { width: 10 }, { width: 25 }, { width: 25 }];
    wsBS.addRow([`BALANCE SHEET AS AT 31st MARCH ${y1}`]).font = { bold: true };
    wsBS.mergeCells(1, 1, 1, wsBS.columns.length);
    wsBS.addRow([]);
    const bsHead = wsBS.addRow(['ID', 'Particulars', 'Note No.', `As at 31st March, ${y1}\nRs.`, `As at 31st March, ${y2}\nRs.`]);
    bsHead.height = 30;
    bsHead.eachCell(c => { c.font = { bold: true, size: 12, name: 'Times New Roman' }; c.alignment = { horizontal: 'center', wrapText: true, vertical: 'middle' }; c.border = { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN }; });
    
    const bsRows = [
        ['I', 'EQUITY AND LIABILITIES', '', '', ''],
        ['1', 'Shareholders’ funds', '', '', ''],
        ['', '(a) Share capital', '1', getNote(leftData, 1), getNote(rightData, 1)],
        ['', '(b) Reserves and surplus', '2', getNote(leftData, 2), getNote(rightData, 2)],
        ['', '(c) Money received against share warrants', '', '', ''],
        ['2', 'Share application money pending allotment', '', '', ''],
        ['3', 'Non-current liabilities', '', '', ''],
        ['', '(a) Long-term borrowings', '3', getNote(leftData, 3), getNote(rightData, 3)],
        ['', '(b) Deferred tax liabilities (net)', '4', getNote(leftData, 4), getNote(rightData, 4)],
        ['', '(c) Other Long-term liabilities', '', '', ''],
        ['', '(d) Long-term provisions', '', '', ''],
        ['4', 'Current liabilities', '', '', ''],
        ['', '(a) Short-term borrowings', '', '', ''],
        ['', '(b) Trade payables', '', '', ''],
        ['', '    (i) Total outstanding dues of micro enterprises and small enterprises', '', '', ''],
        ['', '    (ii) Total outstanding dues of creditors other than micro enterprises and small enterprises', '6', getNote(leftData, 6), getNote(rightData, 6)],
        ['', '(c) Other current liabilities', '7', getNote(leftData, 7), getNote(rightData, 7)],
        ['', '(d) Short-term provisions', '5', getNote(leftData, 5), getNote(rightData, 5)]
    ];
    let lEqTotal = 0, rEqTotal = 0;
    for(let i=1; i<=7; i++){ lEqTotal+=getNote(leftData,i); rEqTotal+=getNote(rightData,i); }
    bsRows.push(['', 'TOTAL', '', lEqTotal, rEqTotal]);
    
    bsRows.push(['', '', '', '', '']);
    bsRows.push(['II', 'ASSETS', '', '', '']);
    bsRows.push(['1', 'Non-current assets', '', '', '']);
    bsRows.push(['', '(a) Property Plant and Equipment and Intangible assets', '', '', '']);
    bsRows.push(['', '    (i) Property, Plant and Equipment', '8', getNote(leftData, 8), getNote(rightData, 8)]);
    bsRows.push(['', '    (ii) Intangible assets', '', '', '']);
    bsRows.push(['', '    (iii) Capital Work In Progress', '', '', '']);
    bsRows.push(['', '    (iv) Intangible Assets Under Developments', '', '', '']);
    bsRows.push(['', '(b) Non-current investments', '10', getNote(leftData, 10), getNote(rightData, 10)]);
    bsRows.push(['', '(c) Deferred tax assets (net)', '', '', '']);
    bsRows.push(['', '(d) Long-term loans and advances', '9', getNote(leftData, 9), getNote(rightData, 9)]);
    bsRows.push(['', '(e) Other non-current assets', '', '', '']);
    bsRows.push(['2', 'Current assets', '', '', '']);
    bsRows.push(['', '(a) Current investments', '', '', '']);
    bsRows.push(['', '(b) Inventories', '', '', '']);
    bsRows.push(['', '(c) Trade receivables', '11', getNote(leftData, 11), getNote(rightData, 11)]);
    bsRows.push(['', '(d) Cash and cash equivalents', '12', getNote(leftData, 12), getNote(rightData, 12)]);
    bsRows.push(['', '(e) Short-term loans and advances', '', '', '']);
    bsRows.push(['', '(f) Other current assets', '13', getNote(leftData, 13), getNote(rightData, 13)]);

    // Assets TOTAL: sum Notes 8 (PPE) + 9 (LT Loans) + 10 (Investments) + 11 (Trade Receivables) + 12 (Cash) + 13 (Other Current Assets)
    // Note 11 appears only once (under Trade receivables); Other non-current assets row is intentionally blank.
    let lAsTotal = 0, rAsTotal = 0;
    for(let i=8; i<=13; i++){ lAsTotal+=getNote(leftData,i); rAsTotal+=getNote(rightData,i); }
    bsRows.push(['', 'TOTAL', '', lAsTotal, rAsTotal]);

    bsRows.forEach(rData => {
        const r = wsBS.addRow(rData);
        if (rData[0] || rData[1] === 'TOTAL') r.font = { bold: true };
        applyBorders(r);
    });

    // 3. PL
    const wsPL = workbook.addWorksheet('PL');
    wsPL.columns = [{ width: 8 }, { width: 50 }, { width: 10 }, { width: 25 }, { width: 25 }];
    wsPL.addRow([`STATEMENT OF PROFIT AND LOSS ACCOUNT FOR THE YEAR ENDED 31st MARCH ${y1}`]).font = { bold: true };
    wsPL.mergeCells(1, 1, 1, wsPL.columns.length);
    wsPL.addRow([]);
    const plHead = wsPL.addRow(['ID', 'Particulars', 'Note No.', `As at 31st March, ${y1}\nRs.`, `As at 31st March, ${y2}\nRs.`]);
    plHead.height = 30;
    plHead.eachCell(c => { c.font = { bold: true, size: 12, name: 'Times New Roman' }; c.alignment = { horizontal: 'center', wrapText: true, vertical: 'middle' }; c.border = { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN }; });
    
    const lInc = getNote(leftData, 14) + getNote(leftData, 15);
    const rInc = getNote(rightData, 14) + getNote(rightData, 15);
    const lExp = getNote(leftData, 16) + getNote(leftData, 17) + getNote(leftData, 18);
    const rExp = getNote(rightData, 16) + getNote(rightData, 17) + getNote(rightData, 18);
    
    const plRows = [
        ['', 'INCOME', '', '', ''],
        ['I', 'Revenue from operations', '14', getNote(leftData, 14), getNote(rightData, 14)],
        ['II', 'Other Income', '15', getNote(leftData, 15), getNote(rightData, 15)],
        ['III', 'TOTAL INCOME ( I + II )', '', lInc, rInc],
        ['', '', '', '', ''],
        ['IV', 'EXPENSES', '', '', ''],
        ['', '(a) Cost of materials consumed', '', '', ''],
        ['', '(b) Purchases of Stock In Trade', '', '', ''],
        ['', '(c) Changes in inventories of finished goods,', '', '', ''],
        ['', '(d) Changes in work-in-progress and stock-in-trade', '', '', ''],
        ['', '(e) Employee benefits expenses', '16', getNote(leftData, 16), getNote(rightData, 16)],
        ['', '(f) Depreciation and amortisation expenses', '17', getNote(leftData, 17), getNote(rightData, 17)],
        ['', '(g) Finance costs', '', '', ''],
        ['', '(h) Other expenses', '18', getNote(leftData, 18), getNote(rightData, 18)],
        ['', 'TOTAL EXPENSES', '', lExp, rExp],
        ['', '', '', '', ''],
        ['V', 'Profit before exceptional and extraordinary items and tax (III-IV)', '', lInc - lExp, rInc - rExp],
        ['VI', 'Exceptional items', '', '', ''],
        ['VII', 'Profit before extraordinary items and tax ( V- VI)', '', lInc - lExp, rInc - rExp],
        ['VIII', 'Extraordinary Items', '', '', ''],
        ['IX', 'Profit before tax (VII-VIII)', '', lInc - lExp, rInc - rExp],
        ['X', 'Tax Expense:', '', '', ''],
        ['', '(a) Current tax expense', '', '', ''],
        ['', '(b) Deferred tax', '', '', ''],
        ['XI', 'Profit / (Loss) from continuing operations (VII-VIII)', '', lInc - lExp, rInc - rExp]
    ];
    
    plRows.forEach(rData => {
        const r = wsPL.addRow(rData);
        if (rData[0] || rData[1].startsWith('TOTAL')) r.font = { bold: true };
        applyBorders(r);
    });

    function addCommonHeader(ws, title) {
        ws.addRow([`NOTES TO THE FINANCIAL STATEMENTS FOR THE YEAR ENDED 31 MARCH ${y1}`]).font = { bold: true };
        ws.mergeCells(1, 1, 1, ws.columns.length);
        ws.addRow([]);
        ws.addRow([title]).font = { bold: true, size: 12 };
        ws.mergeCells(3, 1, 3, ws.columns.length);
        ws.addRow([]);
    }

    // 4. Note 1 Sh Cap
    const ws1 = workbook.addWorksheet('Note 1 Sh Cap');
    ws1.columns = [{ width: 50 }, { width: 25 }, { width: 25 }];
    addCommonHeader(ws1, 'Note 1 - Share Capital');
    ws1.addRow(['Particulars', `As at 31 March, ${y1}\nRs.`, `As at 31 March, ${y2}\nRs.`]).font = { bold: true };
    ws1.addRow(['(a) Authorised']).font = { bold: true };
    const lAuth = leftPvtInfo ? `${leftPvtInfo.noShares || 0} Equity shares of Rs.${leftPvtInfo.faceVal || 0}/- each` : 'Equity shares of Rs.1/- each';
    const rAuth = rightPvtInfo ? `${rightPvtInfo.noShares || 0} Equity shares of Rs.${rightPvtInfo.faceVal || 0}/- each` : 'Equity shares of Rs.1/- each';
    
    // We display the authorized capital amount in the columns
    ws1.addRow([lAuth, leftPvtInfo ? leftPvtInfo.authCap : '', rightPvtInfo ? rightPvtInfo.authCap : '']);
    ws1.addRow(['Preference shares', '', '']);
    ws1.addRow(['(b) Issued']).font = { bold: true };
    ws1.addRow(['Subscribed and fully paid up', getNote(leftData, 1), getNote(rightData, 1)]);
    ws1.addRow(['Subscribed and not fully paid up', '', '']);
    ws1.addRow(['Total', getNote(leftData, 1), getNote(rightData, 1)]).font = { bold: true };
    
    // 5. Note 2 & 3
    const ws23 = workbook.addWorksheet('Note 2& Note 3');
    ws23.columns = [{ width: 50 }, { width: 25 }, { width: 25 }];
    addCommonHeader(ws23, 'Note 2 - Reserves & Surplus');
    ws23.addRow(['Particulars', `As at 31 March, ${y1}\nRs.`, `As at 31 March, ${y2}\nRs.`]).font = { bold: true };
    ws23.addRow(['(B) Surplus / (Deficit) in Statement of Profit and Loss']).font = { bold: true };
    ws23.addRow(['Opening balance', '', '']);
    ws23.addRow(['Profit / (Loss) for the year', getNote(leftData, 2), getNote(rightData, 2)]);
    ws23.addRow(['Closing balance', getNote(leftData, 2), getNote(rightData, 2)]).font = { bold: true };
    ws23.addRow(['Total', getNote(leftData, 2), getNote(rightData, 2)]).font = { bold: true };
    ws23.addRow([]);
    
    ws23.addRow(['Note 3 - Long-term Borrowings']).font = { bold: true, size: 12 };
    ws23.addRow(['Particulars', `As at 31 March, ${y1}\nRs.`, `As at 31 March, ${y2}\nRs.`]).font = { bold: true };
    ws23.addRow(['(A) Secured']).font = { bold: true };
    
    const lSecTotal = leftData?.loansInfo?.secTotal || 0;
    const rSecTotal = rightData?.loansInfo?.secTotal || 0;
    const secKeys = getKeys(leftData?.loansInfo?.secured, rightData?.loansInfo?.secured);
    secKeys.forEach(k => {
        ws23.addRow([k, leftData?.loansInfo?.secured?.[k] || 0, rightData?.loansInfo?.secured?.[k] || 0]);
    });
    ws23.addRow(['Secured Loans', lSecTotal, rSecTotal]).font = { bold: true };
    
    ws23.addRow(['(B) Unsecured']).font = { bold: true };
    const lUnsecTotal = leftData?.loansInfo?.unsecTotal || 0;
    const rUnsecTotal = rightData?.loansInfo?.unsecTotal || 0;
    const unsecKeys = getKeys(leftData?.loansInfo?.unsecured, rightData?.loansInfo?.unsecured);
    unsecKeys.forEach(k => {
        ws23.addRow([k, leftData?.loansInfo?.unsecured?.[k] || 0, rightData?.loansInfo?.unsecured?.[k] || 0]);
    });
    ws23.addRow(['Unsecured Loans', lUnsecTotal, rUnsecTotal]).font = { bold: true };
    ws23.addRow(['Total', getNote(leftData, 3), getNote(rightData, 3)]).font = { bold: true };

    // 6. Note 4 Oth LTL
    const ws4 = workbook.addWorksheet('Note 4 Oth LTL');
    ws4.columns = [{ width: 50 }, { width: 25 }, { width: 25 }];
    addCommonHeader(ws4, 'Note 4 - Other Long-Term Liabilities');
    ws4.addRow(['Particulars', `As at 31 March, ${y1}\nRs.`, `As at 31 March, ${y2}\nRs.`]).font = { bold: true };
    ws4.addRow(['Trade payables', '', '']);
    ws4.addRow(['Total', '', '']).font = { bold: true };

    // 7. Note 5 LTP - STPL
    const ws5 = workbook.addWorksheet('Note 5 LTP - STPL');
    ws5.columns = [{ width: 40 }, { width: 15 }, { width: 15 }, { width: 15 }, { width: 15 }];
    addCommonHeader(ws5, 'Note 5 - Long-Term Provisions and Short-Term Provisions');
    ws5.addRow(['Particulars', `As at 31 March, ${y1}`, '', `As at 31 March, ${y2}`, '']).font = { bold: true };
    ws5.addRow(['', 'Long-term', 'Short-term', 'Long-term', 'Short-term']).font = { bold: true };
    ws5.mergeCells('A5:A6');
    ws5.mergeCells('B5:C5');
    ws5.mergeCells('D5:E5');
    ws5.getCell('A5').alignment = { vertical: 'middle', horizontal: 'center' };
    ws5.getCell('B5').alignment = { vertical: 'middle', horizontal: 'center' };
    ws5.getCell('D5').alignment = { vertical: 'middle', horizontal: 'center' };
    
    const provKeys = getKeys(leftData?.provisionsInfo?.shortTermItems, rightData?.provisionsInfo?.shortTermItems);
    provKeys.forEach(k => {
        ws5.addRow([k, '', leftData?.provisionsInfo?.shortTermItems?.[k] || 0, '', rightData?.provisionsInfo?.shortTermItems?.[k] || 0]);
    });
    ws5.addRow(['Provision for Taxation', '', leftData?.provisionsInfo?.provisionForTaxation || 0, '', rightData?.provisionsInfo?.provisionForTaxation || 0]);
    ws5.addRow(['Total', '', getNote(leftData, 5), '', getNote(rightData, 5)]).font = { bold: true };

    // 8. Note 6 - Trd Pbl
    const ws6 = workbook.addWorksheet('Note 6 - Trd Pbl');
    ws6.columns = [{ width: 50 }, { width: 25 }, { width: 25 }];
    addCommonHeader(ws6, 'Note 6 - Trade Payables');
    ws6.addRow(['Particulars', `As at 31 March, ${y1}\nRs.`, `As at 31 March, ${y2}\nRs.`]).font = { bold: true };
    ws6.addRow(['(A) Total outstanding dues of micro enterprises and small enterprises', '', '']);
    ws6.addRow(['(B) Total outstanding dues of creditors other than micro enterprises and small enterprises', '', '']);
    const credKeys = getKeys(leftData?.sundryCredInfo?.ledgers, rightData?.sundryCredInfo?.ledgers);
    credKeys.forEach(k => {
        ws6.addRow([k, leftData?.sundryCredInfo?.ledgers?.[k] || 0, rightData?.sundryCredInfo?.ledgers?.[k] || 0]);
    });
    ws6.addRow(['Total', getNote(leftData, 6), getNote(rightData, 6)]).font = { bold: true };
    
    ws6.addRow([]);
    ws6.addRow([`Ageing for trade payables as at 31st March, ${y1}`]).font = { bold: true };
    ws6.addRow(['Particulars', 'Not due', 'Less than 1 year', '1 - 2 years', '2 - 3 years', 'More than 3 years', 'Total']).font = { bold: true };
    ws6.addRow(['(i) Undisputed dues - MSME', '', '', '', '', '', '']);
    ws6.addRow(['(ii) Undisputed dues - Others', '', getNote(leftData, 6), '', '', '', getNote(leftData, 6)]);

    // 9. Note 7 - OCL
    const ws7 = workbook.addWorksheet('Note 7 - OCL');
    ws7.columns = [{ width: 50 }, { width: 25 }, { width: 25 }];
    addCommonHeader(ws7, 'Note 7 - Other Current Liabilities');
    ws7.addRow(['Particulars', `As at 31 March, ${y1}\nRs.`, `As at 31 March, ${y2}\nRs.`]).font = { bold: true };
    const oclKeys = getKeys(leftData?.otherCurrLiabInfo?.items, rightData?.otherCurrLiabInfo?.items);
    oclKeys.forEach(k => {
        ws7.addRow([k, leftData?.otherCurrLiabInfo?.items?.[k] || 0, rightData?.otherCurrLiabInfo?.items?.[k] || 0]);
    });
    ws7.addRow(['Total', leftData?.otherCurrLiabInfo?.total ?? getNote(leftData, 7), rightData?.otherCurrLiabInfo?.total ?? getNote(rightData, 7)]).font = { bold: true };

    // 10. Note 8 - PPE
    const ws8 = workbook.addWorksheet('Note 8 - PPE');
    const ppeKeys = getKeys(leftData?.fixedAssetsInfo?.ledgers, rightData?.fixedAssetsInfo?.ledgers);
    ws8.columns = [{ width: 40 }, ...ppeKeys.map(() => ({ width: 20 })), { width: 20 }];
    addCommonHeader(ws8, 'NOTE 08 - PROPERTY PLANT AND EQUIPMENT');
    ws8.addRow(['a. Details of PPE']).font = { bold: true };
    ws8.addRow(['Particulars', ...ppeKeys, 'Total Tangible Assets']).font = { bold: true };
    const ppeRow = (data, year) => [
        `Balance as at March 31, ${year}`,
        ...ppeKeys.map(key => data?.fixedAssetsInfo?.ledgers?.[key] || 0),
        getNote(data, 8)
    ];
    ws8.addRow(ppeRow(leftData, y1));
    if (rightData) ws8.addRow(ppeRow(rightData, y2));
    
    // 11. Note 09 & 10
    const ws9 = workbook.addWorksheet('Note 09  & 10- LT Loans & Adv');
    ws9.columns = [{ width: 50 }, { width: 25 }, { width: 25 }];
    addCommonHeader(ws9, 'Note 9: Long term Loans and Advances');
    ws9.addRow(['Particulars', `As at 31 March, ${y1}\nRs.`, `As at 31 March, ${y2}\nRs.`]).font = { bold: true };
    ws9.addRow(['Unsecured advances : Considered Good']).font = { bold: true };
    const advKeys = getKeys(leftData?.advancesInfo?.items, rightData?.advancesInfo?.items);
    advKeys.forEach(k => {
        ws9.addRow([k, leftData?.advancesInfo?.items?.[k] || 0, rightData?.advancesInfo?.items?.[k] || 0]);
    });
    ws9.addRow(['Total Advances', getNote(leftData, 9), getNote(rightData, 9)]).font = { bold: true };
    
    ws9.addRow([]);
    ws9.addRow(['Note 10: Investments']).font = { bold: true, size: 12 };
    ws9.addRow(['Particulars', `As at 31 March, ${y1}\nRs.`, `As at 31 March, ${y2}\nRs.`]).font = { bold: true };
    const invKeys = getKeys(leftData?.investmentsInfo?.ledgers, rightData?.investmentsInfo?.ledgers);
    invKeys.forEach(k => {
        ws9.addRow([k, leftData?.investmentsInfo?.ledgers?.[k] || 0, rightData?.investmentsInfo?.ledgers?.[k] || 0]);
    });
    ws9.addRow(['Total', getNote(leftData, 10), getNote(rightData, 10)]).font = { bold: true };

    // 12. Note 11 - Other Non Curr Assets
    const ws11 = workbook.addWorksheet('Note 11 - Other Non Curr Assets');
    ws11.columns = [{ width: 50 }, { width: 25 }, { width: 25 }];
    addCommonHeader(ws11, 'Note 11 Other Non Current Assets');
    ws11.addRow(['Particulars', `As at 31 March, ${y1}\nRs.`, `As at 31 March, ${y2}\nRs.`]).font = { bold: true };
    const debtKeys = getKeys(leftData?.sundryDebtorsInfo?.ledgers, rightData?.sundryDebtorsInfo?.ledgers);
    debtKeys.forEach(k => {
        ws11.addRow([k, leftData?.sundryDebtorsInfo?.ledgers?.[k] || 0, rightData?.sundryDebtorsInfo?.ledgers?.[k] || 0]);
    });
    ws11.addRow(['Total', getNote(leftData, 11), getNote(rightData, 11)]).font = { bold: true };

    // 13. Note 12 & Note 13
    const ws12 = workbook.addWorksheet('Note 12 & Note 13');
    ws12.columns = [{ width: 50 }, { width: 25 }, { width: 25 }];
    addCommonHeader(ws12, 'Note 12 Cash And Cash Equivalents');
    ws12.addRow(['Particulars', `As at 31 March, ${y1}\nRs.`, `As at 31 March, ${y2}\nRs.`]).font = { bold: true };
    ws12.addRow(['Balances with Banks', leftData?.cashInfo?.bankAccounts || 0, rightData?.cashInfo?.bankAccounts || 0]);
    ws12.addRow(['Cheques, drafts on hand', '', '']);
    ws12.addRow(['Cash on Hand', leftData?.cashInfo?.cashInHand || 0, rightData?.cashInfo?.cashInHand || 0]);
    ws12.addRow(['Total', getNote(leftData, 12), getNote(rightData, 12)]).font = { bold: true };
    
    ws12.addRow([]);
    ws12.addRow(['Note 13 Other Current Assets']).font = { bold: true, size: 12 };
    ws12.addRow(['Particulars', `As at 31 March, ${y1}\nRs.`, `As at 31 March, ${y2}\nRs.`]).font = { bold: true };
    ws12.addRow(['Total', getNote(leftData, 13), getNote(rightData, 13)]).font = { bold: true };

    // 14. Note 14 - 18 - PL Schedules
    const ws14 = workbook.addWorksheet('Note 14 - 18 - PL Schedules');
    ws14.columns = [{ width: 50 }, { width: 25 }, { width: 25 }];
    addCommonHeader(ws14, 'Note 14 Revenue From Operations');
    ws14.addRow(['Particulars', `For the year ended 31 March, ${y1}\nRs.`, `For the year ended 31 March, ${y2}\nRs.`]).font = { bold: true };
    ws14.addRow(['Sale of Services', getNote(leftData, 14), getNote(rightData, 14)]);
    ws14.addRow(['Total', getNote(leftData, 14), getNote(rightData, 14)]).font = { bold: true };
    
    ws14.addRow([]);
    ws14.addRow(['Note 15 Other Income']).font = { bold: true, size: 12 };
    ws14.addRow(['Particulars', `For the year ended 31 March, ${y1}\nRs.`, `For the year ended 31 March, ${y2}\nRs.`]).font = { bold: true };
    const incKeys = getKeys(leftData?.indirectIncomesInfo?.items, rightData?.indirectIncomesInfo?.items);
    incKeys.forEach(k => {
        ws14.addRow([k, leftData?.indirectIncomesInfo?.items?.[k] || 0, rightData?.indirectIncomesInfo?.items?.[k] || 0]);
    });
    ws14.addRow(['Total', getNote(leftData, 15), getNote(rightData, 15)]).font = { bold: true };
    
    ws14.addRow([]);
    ws14.addRow(['Note 16 Employee Benefit Expenses']).font = { bold: true, size: 12 };
    ws14.addRow(['Particulars', `For the year ended 31 March, ${y1}\nRs.`, `For the year ended 31 March, ${y2}\nRs.`]).font = { bold: true };
    const empKeys = getKeys(leftData?.note16Info?.items, rightData?.note16Info?.items);
    empKeys.forEach(k => {
        ws14.addRow([k, leftData?.note16Info?.items?.[k] || 0, rightData?.note16Info?.items?.[k] || 0]);
    });
    ws14.addRow(['Total', getNote(leftData, 16), getNote(rightData, 16)]).font = { bold: true };
    
    ws14.addRow([]);
    ws14.addRow(['Note 17 Depreciation and Amortisation Expenses']).font = { bold: true, size: 12 };
    ws14.addRow(['Particulars', `For the year ended 31 March, ${y1}\nRs.`, `For the year ended 31 March, ${y2}\nRs.`]).font = { bold: true };
    ws14.addRow(['Depreciation on property, plant and equipment', getNote(leftData, 17), getNote(rightData, 17)]);
    ws14.addRow(['Total', getNote(leftData, 17), getNote(rightData, 17)]).font = { bold: true };

    ws14.addRow([]);
    ws14.addRow(['Note 18 Other expenses']).font = { bold: true, size: 12 };
    ws14.addRow(['Particulars', `For the year ended 31 March, ${y1}\nRs.`, `For the year ended 31 March, ${y2}\nRs.`]).font = { bold: true };
    // Assuming other expenses are not individually broken down in the mapping, we just put total
    ws14.addRow(['Total', getNote(leftData, 18), getNote(rightData, 18)]).font = { bold: true };

    const audit = workbook.addWorksheet('Tally Data Audit');
    audit.columns = [{ width: 12 }, { width: 42 }, { width: 24 }, { width: 22 }, { width: 22 }];
    audit.addRow(['Note', 'Particular', 'Source / Variable', `31 Mar ${y1}`, `31 Mar ${y2}`]).font = { bold: true };
    for (let n = 1; n <= 18; n++) {
        audit.addRow([`Note ${n}`, `Schedule III Note ${n}`, `mapping[Note ${n}]`, getNote(leftData, n), getNote(rightData, n)]);
    }

    // Format all sheets
    workbook.eachSheet((sheet) => {
        sheet.eachRow((row, rowNumber) => {
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                // Default font
                cell.font = cell.font || {};
                cell.font.name = 'Times New Roman';
                if (!cell.font.size) cell.font.size = 11;
                
                // Alignment
                cell.alignment = cell.alignment || {};
                cell.alignment.vertical = 'middle';
                
                // Title rows (Row 1)
                if (rowNumber === 1) {
                    cell.font.size = 12;
                    cell.font.bold = true;
                    if (cell.value) cell.alignment.horizontal = 'center';
                }
                
                // Note headers
                if (cell.value && typeof cell.value === 'string' && cell.value.startsWith('Note ')) {
                    cell.font.size = 12;
                    cell.font.bold = true;
                }
                
                // Format table headers (Row 3 for BS and PL)
                if (rowNumber === 3 && (sheet.name === 'BS' || sheet.name === 'PL')) {
                    cell.font.size = 12;
                    cell.font.bold = true;
                    cell.alignment.horizontal = 'center';
                    cell.alignment.wrapText = true;
                }

                // Any cell containing \n (like '...2024\nRs.') should wrap text and center align
                if (typeof cell.value === 'string' && cell.value.includes('\n')) {
                    cell.alignment.wrapText = true;
                    cell.alignment.horizontal = 'center';
                }
                
                // Wrap and center all table headers (bold cells in header rows)
                if (cell.font.bold && typeof cell.value === 'string' && rowNumber > 1 && rowNumber <= 6) {
                    cell.alignment.wrapText = true;
                    cell.alignment.horizontal = 'center';
                }
                
                // Center align the 'Particulars' header in Notes
                if (typeof cell.value === 'string' && cell.value === 'Particulars') {
                    cell.alignment.horizontal = 'center';
                }

                // Ensure borders for notes tables (rough heuristic: if we are below the header rows and column 1 has particulars)
                if (sheet.name.startsWith('Note') && rowNumber >= 5) {
                    if (row.getCell(1).value || row.getCell(2).value) {
                         cell.border = { top: BORDER_THIN, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN };
                    }
                }

                // Indentation for particulars (assuming column 2 for BS/PL or column 1 for Notes)
                if (typeof cell.value === 'string') {
                    const val = cell.value.trim();
                    if (val.startsWith('(a)') || val.startsWith('(b)') || val.startsWith('(c)') || val.startsWith('(d)') || val.startsWith('(e)') || val.startsWith('(f)') || val.startsWith('(g)') || val.startsWith('(h)')) {
                        cell.alignment.indent = 1;
                    } else if (val.startsWith('(i)') || val.startsWith('(ii)') || val.startsWith('(iii)') || val.startsWith('(iv)')) {
                        cell.alignment.indent = 2;
                    }
                    
                    // Also center align Note No and ID
                    if ((sheet.name === 'BS' || sheet.name === 'PL') && (colNumber === 1 || colNumber === 3)) {
                        cell.alignment.horizontal = 'center';
                    }
                }

                // Numbers formatting
                if (typeof cell.value === 'number') {
                    cell.numFmt = NUM_FMT;
                    cell.alignment.horizontal = 'right';
                } else if (cell.value === '-') {
                    cell.alignment.horizontal = 'center';
                }
            });
        });
    });

    try {
        await workbook.xlsx.writeFile(outputPath);
        console.log(`Saved merged file to ${outputPath}`);
    } catch(err) {
        console.error("Failed to write Excel:", err);
        throw err;
    }
}

const dynamicExcel = require('./dynamic_schedule3_excel');
module.exports = { generateMergedExcel, generateMergedExcelRwaTrust, ...dynamicExcel };
