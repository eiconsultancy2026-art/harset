const ExcelJS = require('exceljs');

async function generateMergedExcelRwaTrust(leftName, rightName, leftData, rightData, outputPath, leftYear, rightYear) {
    const wb = new ExcelJS.Workbook();
    const y1 = leftYear  || new Date().getFullYear();
    const y2 = rightYear || (parseInt(y1) - 1);
    const NUM = '[>=10000000]##\\,##\\,##\\,##0.00;[>=100000]##\\,##\\,##0.00;##,##0.00';
    const TN  = 'Times New Roman';
    const T   = { style: 'thin', color: { argb: 'FF000000' } };
    const B   = { top: T, left: T, bottom: T, right: T };

    // ── Helpers ───────────────────────────────────────────────────────────────
    function noteVal(data, num) {
        if (!data?.mapping) return 0;
        const k = Object.keys(data.mapping).find(k => new RegExp(`^Note\\s+${num}\\s*-`, 'i').test(k));
        return k ? (data.mapping[k] || 0) : 0;
    }

    function unionKeys(a, b) {
        return [...new Set([...Object.keys(a || {}), ...Object.keys(b || {})])].sort();
    }

    function dataRow(ws, label, vL, vR, bold) {
        const r = ws.addRow([label, vL || 0, vR || 0]);
        r.font = bold ? { bold: true, name: TN } : { name: TN };
        r.getCell(2).numFmt = NUM;
        r.getCell(3).numFmt = NUM;
        r.eachCell({ includeEmpty: true }, c => c.border = B);
        return r;
    }

    function secTitle(ws, text) {
        const r = ws.addRow([text]);
        r.font = { bold: true, name: TN, size: 12 };
        ws.mergeCells(ws.lastRow.number, 1, ws.lastRow.number, 3);
        return r;
    }

    function sheetHeader(ws, cols) {
        ws.columns = cols || [{ width: 50 }, { width: 22 }, { width: 22 }];
        const r = ws.addRow([`NOTES TO FINANCIAL STATEMENTS FOR THE YEAR ENDED 31ST MARCH ${y1}`]);
        ws.mergeCells(1, 1, 1, 3);
        r.font = { bold: true, name: TN, size: 11 };
        r.getCell(1).alignment = { horizontal: 'center' };
        ws.addRow([]);
    }

    function colHdr(ws, isIncome) {
        const c1 = isIncome ? `For year ended 31 Mar ${y1} (Rs.)` : `As at 31 Mar ${y1} (Rs.)`;
        const c2 = isIncome ? `For year ended 31 Mar ${y2} (Rs.)` : `As at 31 Mar ${y2} (Rs.)`;
        const r = ws.addRow(['Particulars', c1, c2]);
        r.font = { bold: true, name: TN };
        r.height = 30;
        r.eachCell(c => { c.border = B; c.alignment = { horizontal: 'center', wrapText: true }; });
        return r;
    }

    // ── SHEET 1: Balance Sheet ────────────────────────────────────────────────
    const wsBS = wb.addWorksheet('Balance Sheet');
    wsBS.columns = [{ width: 6 }, { width: 52 }, { width: 8 }, { width: 22 }, { width: 22 }];

    const bsT = wsBS.addRow([`BALANCE SHEET AS AT 31ST MARCH ${y1}`]);
    wsBS.mergeCells(1, 1, 1, 5);
    bsT.font = { bold: true, name: TN, size: 13 };
    bsT.getCell(1).alignment = { horizontal: 'center' };

    wsBS.addRow([`Name of the Entity: ${leftName} vs ${rightName}`]).font = { italic: true, name: TN };
    wsBS.addRow([]);

    const bsHdr = wsBS.addRow(['', 'Particulars', 'Note No.', `As at 31 Mar ${y1} (Rs.)`, `As at 31 Mar ${y2} (Rs.)`]);
    bsHdr.font = { bold: true, name: TN };
    bsHdr.height = 35;
    bsHdr.eachCell(c => { c.border = B; c.alignment = { horizontal: 'center', wrapText: true, vertical: 'middle' }; });

    function bsSection(ws, label) {
        const r = ws.addRow(['', label, '', '', '']);
        r.font = { bold: true, name: TN, size: 11 };
        ws.mergeCells(ws.lastRow.number, 2, ws.lastRow.number, 5);
    }

    function bsRow(ws, label, note) {
        const vL = noteVal(leftData, note), vR = noteVal(rightData, note);
        const r = ws.addRow(['', label, note, vL, vR]);
        r.font = { name: TN };
        r.getCell(4).numFmt = NUM; r.getCell(5).numFmt = NUM;
        r.eachCell({ includeEmpty: true }, c => c.border = B);
        return [vL, vR];
    }

    function bsTotalRow(ws, label, vL, vR) {
        const r = ws.addRow(['', label, '', vL, vR]);
        r.font = { bold: true, name: TN };
        r.getCell(4).numFmt = NUM; r.getCell(5).numFmt = NUM;
        r.eachCell({ includeEmpty: true }, c => c.border = B);
    }

    bsSection(wsBS, 'I. SOURCES OF FUNDS (LIABILITIES)');
    let liabL = 0, liabR = 0;
    [[3,'NPOs Funds'],[4,'Borrowings'],[5,'Other Long-term Liabilities'],
     [6,'Provisions'],[7,'Payables'],[8,'Other Current Liabilities']].forEach(([note, label]) => {
        const [vL, vR] = bsRow(wsBS, label, note);
        liabL += vL; liabR += vR;
    });
    bsTotalRow(wsBS, 'Total Sources of Funds', liabL, liabR);
    wsBS.addRow([]);

    bsSection(wsBS, 'II. APPLICATION OF FUNDS (ASSETS)');
    let assetL = 0, assetR = 0;
    [[9,'PPE (Fixed Assets)'],[10,'Investments'],[11,'Loans and Advances'],
     [12,'Other Non-current Assets'],[13,'Receivables'],[14,'Cash and Bank Balance'],[15,'Other Current Assets']].forEach(([note, label]) => {
        const [vL, vR] = bsRow(wsBS, label, note);
        assetL += vL; assetR += vR;
    });
    bsTotalRow(wsBS, 'Total Application of Funds', assetL, assetR);

    // ── SHEET 2: Statement of P&L ─────────────────────────────────────────────
    const wsPL = wb.addWorksheet('Income & Expenditure');
    wsPL.columns = [{ width: 6 }, { width: 52 }, { width: 8 }, { width: 22 }, { width: 22 }];

    const plT = wsPL.addRow([`STATEMENT OF INCOME AND EXPENDITURE FOR THE YEAR ENDED 31ST MARCH ${y1}`]);
    wsPL.mergeCells(1, 1, 1, 5);
    plT.font = { bold: true, name: TN, size: 13 };
    plT.getCell(1).alignment = { horizontal: 'center' };
    wsPL.addRow([`Name of the Entity: ${leftName} vs ${rightName}`]).font = { italic: true, name: TN };
    wsPL.addRow([]);

    const plHdr = wsPL.addRow(['', 'Particulars', 'Note No.', `For year ended 31 Mar ${y1} (Rs.)`, `For year ended 31 Mar ${y2} (Rs.)`]);
    plHdr.font = { bold: true, name: TN };
    plHdr.height = 35;
    plHdr.eachCell(c => { c.border = B; c.alignment = { horizontal: 'center', wrapText: true, vertical: 'middle' }; });

    function plSection(ws, label) {
        const r = ws.addRow(['', label, '', '', '']);
        r.font = { bold: true, name: TN, size: 11 };
        ws.mergeCells(ws.lastRow.number, 2, ws.lastRow.number, 5);
    }

    function plDataRow(ws, label, note) {
        const vL = noteVal(leftData, note), vR = noteVal(rightData, note);
        const r = ws.addRow(['', label, note, vL, vR]);
        r.font = { name: TN };
        r.getCell(4).numFmt = NUM; r.getCell(5).numFmt = NUM;
        r.eachCell({ includeEmpty: true }, c => c.border = B);
        return [vL, vR];
    }

    function plTotalRow(ws, label, vL, vR) {
        const r = ws.addRow(['', label, '', vL, vR]);
        r.font = { bold: true, name: TN };
        r.getCell(4).numFmt = NUM; r.getCell(5).numFmt = NUM;
        r.eachCell({ includeEmpty: true }, c => c.border = B);
    }

    plSection(wsPL, 'I. INCOME');
    const [incL, incR] = plDataRow(wsPL, 'Other Income (Receipts)', 16);
    plTotalRow(wsPL, 'Total Income', incL, incR);
    wsPL.addRow([]);

    plSection(wsPL, 'II. EXPENDITURE');
    let expL = 0, expR = 0;
    [[17,'Cost of Goods Sold'],[18,'Employee Benefits Expense'],
     [19,'Finance Cost'],[20,'Depreciation and Amortisation'],[21,'Other Expenses']].forEach(([note, label]) => {
        const [vL, vR] = plDataRow(wsPL, label, note);
        expL += vL; expR += vR;
    });
    plTotalRow(wsPL, 'Total Expenditure', expL, expR);
    wsPL.addRow([]);
    plTotalRow(wsPL, 'Excess of Income over Expenditure', incL - expL, incR - expR);

    // ── SHEET 3: Notes 1 to 3 ────────────────────────────────────────────────
    const ws3 = wb.addWorksheet('Note 1 to 3');
    sheetHeader(ws3);
    secTitle(ws3, 'Note 1 – Brief about the Entity');
    ws3.addRow(['(Narrative note – attach disclosure separately)']).font = { italic: true, name: TN };
    ws3.addRow([]);
    secTitle(ws3, 'Note 2 – Significant Accounting Policies');
    ws3.addRow(['(Narrative note – attach disclosure separately)']).font = { italic: true, name: TN };
    ws3.addRow([]);
    secTitle(ws3, 'Note 3 – NPOs Funds');
    colHdr(ws3);
    const capL  = leftData?.summaryGroups?.['Capital Account'] || 0;
    const capR  = rightData?.summaryGroups?.['Capital Account'] || 0;
    const surpL = leftData?.summaryGroups?.['Excess of income over expenditure'] || 0;
    const surpR = rightData?.summaryGroups?.['Excess of income over expenditure'] || 0;
    dataRow(ws3, 'Capital Account (Opening Balance)', capL, capR);
    dataRow(ws3, 'Add: Excess of Income over Expenditure', surpL, surpR);
    dataRow(ws3, 'Total NPOs Funds', noteVal(leftData,3), noteVal(rightData,3), true);

    // ── SHEET 4: Note 4 – Borrowings ─────────────────────────────────────────
    const ws4 = wb.addWorksheet('Notes to BS 4');
    sheetHeader(ws4);
    secTitle(ws4, 'Note 4 – Borrowings');
    colHdr(ws4);
    ws4.addRow(['(A) Secured']).font = { bold: true, name: TN };
    unionKeys(leftData?.loansInfo?.secured, rightData?.loansInfo?.secured).forEach(k =>
        dataRow(ws4, '  ' + k, leftData?.loansInfo?.secured?.[k] || 0, rightData?.loansInfo?.secured?.[k] || 0));
    dataRow(ws4, 'Total Secured', leftData?.loansInfo?.secTotal || 0, rightData?.loansInfo?.secTotal || 0, true);
    ws4.addRow([]);
    ws4.addRow(['(B) Unsecured']).font = { bold: true, name: TN };
    unionKeys(leftData?.loansInfo?.unsecured, rightData?.loansInfo?.unsecured).forEach(k =>
        dataRow(ws4, '  ' + k, leftData?.loansInfo?.unsecured?.[k] || 0, rightData?.loansInfo?.unsecured?.[k] || 0));
    dataRow(ws4, 'Total Unsecured', leftData?.loansInfo?.unsecTotal || 0, rightData?.loansInfo?.unsecTotal || 0, true);
    ws4.addRow([]);
    dataRow(ws4, 'Total Borrowings (A + B)', noteVal(leftData,4), noteVal(rightData,4), true);

    // ── SHEET 5: Notes 5 to 8 ────────────────────────────────────────────────
    const ws5 = wb.addWorksheet('Notes to BS 5 to 8');
    sheetHeader(ws5);

    secTitle(ws5, 'Note 5 – Other Long-term Liabilities');
    colHdr(ws5);
    dataRow(ws5, 'Total Other Long-term Liabilities', noteVal(leftData,5), noteVal(rightData,5), true);
    ws5.addRow([]);

    secTitle(ws5, 'Note 6 – Provisions');
    colHdr(ws5);
    unionKeys(leftData?.provisionsInfo?.shortTermItems, rightData?.provisionsInfo?.shortTermItems).forEach(k =>
        dataRow(ws5, '  ' + k, leftData?.provisionsInfo?.shortTermItems?.[k]||0, rightData?.provisionsInfo?.shortTermItems?.[k]||0));
    const ptL = leftData?.provisionsInfo?.provisionForTaxation || 0;
    const ptR = rightData?.provisionsInfo?.provisionForTaxation || 0;
    if (ptL || ptR) dataRow(ws5, '  Provision for Taxation', ptL, ptR);
    dataRow(ws5, 'Total Provisions', noteVal(leftData,6), noteVal(rightData,6), true);
    ws5.addRow([]);

    secTitle(ws5, 'Note 7 – Payables (Trade Payables)');
    colHdr(ws5);
    unionKeys(leftData?.sundryCredInfo?.ledgers, rightData?.sundryCredInfo?.ledgers).forEach(k =>
        dataRow(ws5, '  ' + k, leftData?.sundryCredInfo?.ledgers?.[k]||0, rightData?.sundryCredInfo?.ledgers?.[k]||0));
    dataRow(ws5, 'Total Payables', noteVal(leftData,7), noteVal(rightData,7), true);
    ws5.addRow([]);

    secTitle(ws5, 'Note 8 – Other Current Liabilities');
    colHdr(ws5);
    unionKeys(leftData?.otherCurrLiabInfo?.items, rightData?.otherCurrLiabInfo?.items).forEach(k =>
        dataRow(ws5, '  ' + k, leftData?.otherCurrLiabInfo?.items?.[k]||0, rightData?.otherCurrLiabInfo?.items?.[k]||0));
    dataRow(ws5, 'Total Other Current Liabilities', Math.abs(noteVal(leftData,8)), Math.abs(noteVal(rightData,8)), true);

    // ── SHEET 6: Note 9 – PPE ────────────────────────────────────────────────
    const ws6 = wb.addWorksheet('Notes to BS 9');
    sheetHeader(ws6);
    secTitle(ws6, 'Note 9 – Property, Plant & Equipment (PPE)');
    colHdr(ws6);
    dataRow(ws6, 'Gross Block (Opening Balance)', 0, 0);
    dataRow(ws6, 'Additions during the year', noteVal(leftData,9), noteVal(rightData,9));
    dataRow(ws6, 'Disposals / Deductions', 0, 0);
    dataRow(ws6, 'Gross Block (Closing Balance)', noteVal(leftData,9), noteVal(rightData,9));
    dataRow(ws6, 'Less: Accumulated Depreciation', noteVal(leftData,20), noteVal(rightData,20));
    dataRow(ws6, 'Net Block (Written Down Value)', noteVal(leftData,9), noteVal(rightData,9), true);

    // ── SHEET 7: Notes 10 to 15 ──────────────────────────────────────────────
    const ws7 = wb.addWorksheet('Notes to BS 10 to 15');
    sheetHeader(ws7);

    secTitle(ws7, 'Note 10 – Investments (Non-current & Current)');
    colHdr(ws7);
    unionKeys(leftData?.investmentsInfo?.ledgers, rightData?.investmentsInfo?.ledgers).forEach(k =>
        dataRow(ws7, '  ' + k, leftData?.investmentsInfo?.ledgers?.[k]||0, rightData?.investmentsInfo?.ledgers?.[k]||0));
    dataRow(ws7, 'Total Investments', noteVal(leftData,10), noteVal(rightData,10), true);
    ws7.addRow([]);

    secTitle(ws7, 'Note 11 – Loans and Advances');
    colHdr(ws7);
    unionKeys(leftData?.advancesInfo?.items, rightData?.advancesInfo?.items).forEach(k =>
        dataRow(ws7, '  ' + k, leftData?.advancesInfo?.items?.[k]||0, rightData?.advancesInfo?.items?.[k]||0));
    dataRow(ws7, 'Total Loans and Advances', noteVal(leftData,11), noteVal(rightData,11), true);
    ws7.addRow([]);

    secTitle(ws7, 'Note 12 – Other Non-current Assets');
    colHdr(ws7);
    dataRow(ws7, 'Total Other Non-current Assets', noteVal(leftData,12), noteVal(rightData,12), true);
    ws7.addRow([]);

    secTitle(ws7, 'Note 13 – Receivables (Members / Debtors)');
    colHdr(ws7);
    unionKeys(leftData?.sundryDebtorsInfo?.ledgers, rightData?.sundryDebtorsInfo?.ledgers).forEach(k =>
        dataRow(ws7, '  ' + k, leftData?.sundryDebtorsInfo?.ledgers?.[k]||0, rightData?.sundryDebtorsInfo?.ledgers?.[k]||0));
    dataRow(ws7, 'Total Receivables', noteVal(leftData,13), noteVal(rightData,13), true);
    ws7.addRow([]);

    secTitle(ws7, 'Note 14 – Cash and Bank Balance');
    colHdr(ws7);
    dataRow(ws7, 'Balances with Banks', leftData?.cashInfo?.bankAccounts||0, rightData?.cashInfo?.bankAccounts||0);
    dataRow(ws7, 'Cash on Hand',        leftData?.cashInfo?.cashInHand||0,   rightData?.cashInfo?.cashInHand||0);
    dataRow(ws7, 'Total Cash and Bank Balance', noteVal(leftData,14), noteVal(rightData,14), true);
    ws7.addRow([]);

    secTitle(ws7, 'Note 15 – Other Current Assets');
    colHdr(ws7);
    dataRow(ws7, 'Total Other Current Assets', noteVal(leftData,15), noteVal(rightData,15), true);

    // ── SHEET 8: Notes 16 to 21 ──────────────────────────────────────────────
    const ws8 = wb.addWorksheet('Notes to Income & Exp 16 to 21');
    sheetHeader(ws8);

    secTitle(ws8, 'Note 16 – Other Income (Receipts)');
    colHdr(ws8, true);
    unionKeys(leftData?.indirectIncomesInfo?.items, rightData?.indirectIncomesInfo?.items).forEach(k =>
        dataRow(ws8, '  ' + k, leftData?.indirectIncomesInfo?.items?.[k]||0, rightData?.indirectIncomesInfo?.items?.[k]||0));
    dataRow(ws8, 'Total Other Income', noteVal(leftData,16), noteVal(rightData,16), true);
    ws8.addRow([]);

    secTitle(ws8, 'Note 17 – Cost of Goods Sold');
    colHdr(ws8, true);
    dataRow(ws8, 'Total Cost of Goods Sold', noteVal(leftData,17), noteVal(rightData,17), true);
    ws8.addRow([]);

    secTitle(ws8, 'Note 18 – Employee Benefits Expense');
    colHdr(ws8, true);
    unionKeys(leftData?.note16Info?.items, rightData?.note16Info?.items).forEach(k =>
        dataRow(ws8, '  ' + k, leftData?.note16Info?.items?.[k]||0, rightData?.note16Info?.items?.[k]||0));
    dataRow(ws8, 'Total Employee Benefits Expense', noteVal(leftData,18), noteVal(rightData,18), true);
    ws8.addRow([]);

    secTitle(ws8, 'Note 19 – Finance Cost');
    colHdr(ws8, true);
    dataRow(ws8, 'Total Finance Cost', noteVal(leftData,19), noteVal(rightData,19), true);
    ws8.addRow([]);

    secTitle(ws8, 'Note 20 – Depreciation and Amortisation Expense');
    colHdr(ws8, true);
    dataRow(ws8, 'Depreciation on PPE', noteVal(leftData,20), noteVal(rightData,20));
    dataRow(ws8, 'Total Depreciation', noteVal(leftData,20), noteVal(rightData,20), true);
    ws8.addRow([]);

    secTitle(ws8, 'Note 21 – Other Expenses');
    colHdr(ws8, true);
    dataRow(ws8, 'Total Other Expenses', noteVal(leftData,21), noteVal(rightData,21), true);

    const audit = wb.addWorksheet('Tally Data Audit');
    audit.columns=[{width:20},{width:45},{width:22},{width:22}];
    audit.addRow(['Note','Particular',`31 Mar ${leftYear}`,`31 Mar ${rightYear}`]).font={bold:true};
    for(let n=1;n<=21;n++){ const l=noteVal(leftData,n)||0, r=noteVal(rightData,n)||0; audit.addRow([`Note ${n}`,'Schedule III total',l,r]); }
    const groups=[['Note 4 Borrowings',leftData?.loansInfo?.secured,rightData?.loansInfo?.secured],['Note 4 Unsecured Loans',leftData?.loansInfo?.unsecured,rightData?.loansInfo?.unsecured],['Note 13 Receivables',leftData?.sundryDebtorsInfo?.ledgers,rightData?.sundryDebtorsInfo?.ledgers],['Note 16 Other Income',leftData?.indirectIncomesInfo?.items,rightData?.indirectIncomesInfo?.items],['Note 18 Employee Benefits',leftData?.note16Info?.items,rightData?.note16Info?.items]];
    for(const [label,a,b] of groups){ for(const k of [...new Set([...Object.keys(a||{}),...Object.keys(b||{})])].sort()) audit.addRow([label,k,a?.[k]||0,b?.[k]||0]); }
    audit.eachRow(r=>r.eachCell(c=>{if(typeof c.value==='number')c.numFmt='#,##0.00;[Red]-#,##0.00';}));
    await wb.xlsx.writeFile(outputPath);
    console.log(`[RWA/Trust] Fresh Excel saved: ${outputPath}`);
}

module.exports = { generateMergedExcelRwaTrust };
