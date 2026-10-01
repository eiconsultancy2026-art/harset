const ExcelJS = require('exceljs');

/**
 * Partnership / Proprietorship / HUF merged Excel – 8 sheets:
 *   Sheet 1 – Balance Sheet (Notes 3–10 liabilities, Notes 11–18 assets)
 *   Sheet 2 – Statement of P&L (Notes 19–25)
 *   Sheet 3 – Notes 1 to 3
 *   Sheet 4 – Notes 4 to 6
 *   Sheet 5 – Notes 7 to 10
 *   Sheet 6 – Note 11 (PPE)
 *   Sheet 7 – Notes 12 to 18
 *   Sheet 8 – Notes 19 to 25
 *
 * Note mapping (Partnership/Prop/HUF):
 *   3  Owner's Capital Account     9  Trade Payables          15 Inventory
 *   4  Reserves and Surplus        10 Other Current Liabilities 16 Trade Receivables
 *   5  Borrowings                  11 PPE                     17 Cash and Bank Balance
 *   6  Deferred Tax Liabilities    12 Investments             18 Other Current Assets
 *   7  Other Long-term Liabilities 13 Loans and Advances      19 Revenue from Operations
 *   8  Provisions                  14 Non-current Assets      20 Other Income
 *                                                             21 Cost of Goods Sold
 *                                                             22 Employee Benefits Expense
 *                                                             23 Finance Cost
 *                                                             24 Depreciation & Amortisation
 *                                                             25 Other Expenses
 */
async function generateMergedExcelPartnership(leftName, leftData, rightName, rightData, outputPath, leftYear, rightYear) {
    const wb  = new ExcelJS.Workbook();
    const y1  = leftYear  || new Date().getFullYear();
    const y2  = rightYear || (parseInt(y1) - 1);
    const NUM = '[>=10000000]##\\,##\\,##\\,##0.00;[>=100000]##\\,##\\,##0.00;##,##0.00';
    const TN  = 'Times New Roman';
    const T   = { style: 'thin', color: { argb: 'FF000000' } };
    const BDR = { top: T, left: T, bottom: T, right: T };

    // ── Helpers ───────────────────────────────────────────────────────────────
    function num(v) {
        if (v === null || v === undefined || v === '') return 0;
        if (typeof v === 'number') return Number.isFinite(v) ? v : 0;
        const raw = String(v).replace(/&amp;/g, '&').trim();
        if (!raw) return 0;
        const paren = /^\(.*\)$/.test(raw);
        const dr = /\bDR\b|\bDEBIT\b/i.test(raw);
        const cr = /\bCR\b|\bCREDIT\b/i.test(raw);
        let x = raw.replace(/[₹,]/g,'').replace(/\bDR\b|\bCR\b|\bDEBIT\b|\bCREDIT\b/ig,'').trim();
        x = x.replace(/[()]/g,'');
        x = Number.parseFloat(x.replace(/[^0-9.+-]/g,''));
        if (!Number.isFinite(x)) return 0;
        if (paren) x = -Math.abs(x);
        else if (cr) x = -Math.abs(x);
        else if (dr) x = Math.abs(x);
        return x;
    }

    function mappingValue(data, numValue) {
        const map = data?.mapping || {};
        const key = Object.keys(map).find(k =>
            new RegExp(`^Note\\s+${numValue}\\s*-`, 'i').test(k)
        );
        return key ? num(map[key]) : 0;
    }

    // The portal values are the source of truth for the merge.
    // Excel must print exactly the same Note value that renderReport() displays.
    // A previous implementation reconstructed zero values from unrelated raw
    // fields; that could silently replace a valid portal value.
    function nv(data, note) {
        return mappingValue(data, note);
    }

    function uk(a, b) {
        return [...new Set([...Object.keys(a||{}), ...Object.keys(b||{})])].sort();
    }

    function dataRow(ws, label, vL, vR, bold) {
        const r = ws.addRow([label, vL||0, vR||0]);
        r.font = bold ? { bold:true, name:TN } : { name:TN };
        r.getCell(2).numFmt = NUM;
        r.getCell(3).numFmt = NUM;
        r.eachCell({ includeEmpty:true }, c => c.border = BDR);
        return r;
    }

    function secTitle(ws, text) {
        const r = ws.addRow([text]);
        r.font = { bold:true, name:TN, size:12 };
        ws.mergeCells(ws.lastRow.number, 1, ws.lastRow.number, 3);
        return r;
    }

    function sheetHeader(ws) {
        ws.columns = [{ width:52 }, { width:22 }, { width:22 }];
        const r = ws.addRow([`NOTES TO FINANCIAL STATEMENTS FOR THE YEAR ENDED 31ST MARCH ${y1}`]);
        ws.mergeCells(1, 1, 1, 3);
        r.font = { bold:true, name:TN, size:11 };
        r.getCell(1).alignment = { horizontal:'center' };
        ws.addRow([]);
    }

    function colHdr(ws, isIncome) {
        const c1 = isIncome ? `For year ended 31 Mar ${y1} (Rs.)` : `As at 31 Mar ${y1} (Rs.)`;
        const c2 = isIncome ? `For year ended 31 Mar ${y2} (Rs.)` : `As at 31 Mar ${y2} (Rs.)`;
        const r  = ws.addRow(['Particulars', c1, c2]);
        r.font   = { bold:true, name:TN };
        r.height = 30;
        r.eachCell(c => { c.border = BDR; c.alignment = { horizontal:'center', wrapText:true }; });
        return r;
    }

    // ── SHEET 1: Balance Sheet ────────────────────────────────────────────────
    const wsBS = wb.addWorksheet('Balance Sheet');
    wsBS.columns = [{ width:6 }, { width:52 }, { width:8 }, { width:22 }, { width:22 }];

    const bsT = wsBS.addRow([`BALANCE SHEET AS AT 31ST MARCH ${y1}`]);
    wsBS.mergeCells(1,1,1,5);
    bsT.font = { bold:true, name:TN, size:13 };
    bsT.getCell(1).alignment = { horizontal:'center' };
    wsBS.addRow([`${leftName} vs ${rightName}`]).font = { italic:true, name:TN };
    wsBS.addRow([]);

    const bsHdr = wsBS.addRow(['', 'Particulars', 'Note No.', `As at 31 Mar ${y1} (Rs.)`, `As at 31 Mar ${y2} (Rs.)`]);
    bsHdr.font = { bold:true, name:TN }; bsHdr.height = 35;
    bsHdr.eachCell(c => { c.border = BDR; c.alignment = { horizontal:'center', wrapText:true, vertical:'middle' }; });

    function bsSec(ws, label) {
        const r = ws.addRow(['', label, '', '', '']);
        r.font = { bold:true, name:TN, size:11 };
        ws.mergeCells(ws.lastRow.number, 2, ws.lastRow.number, 5);
    }
    function bsRow(ws, label, note) {
        const vL = nv(leftData,note), vR = nv(rightData,note);
        const r  = ws.addRow(['', label, note, vL, vR]);
        r.font = { name:TN };
        r.getCell(4).numFmt = NUM; r.getCell(5).numFmt = NUM;
        r.eachCell({ includeEmpty:true }, c => c.border = BDR);
        return [vL, vR];
    }
    function bsTotal(ws, label, vL, vR) {
        const r = ws.addRow(['', label, '', vL, vR]);
        r.font = { bold:true, name:TN };
        r.getCell(4).numFmt = NUM; r.getCell(5).numFmt = NUM;
        r.eachCell({ includeEmpty:true }, c => c.border = BDR);
    }

    bsSec(wsBS, 'I. SOURCES OF FUNDS (LIABILITIES)');
    let liabL=0, liabR=0;
    [[3,"Owner's Capital Account"],[4,'Reserves and Surplus'],[5,'Borrowings'],
     [6,'Deferred Tax Liabilities'],[7,'Other Long-term Liabilities'],
     [8,'Provisions'],[9,'Trade Payables'],[10,'Other Current Liabilities']].forEach(([note,label]) => {
        const [vL,vR] = bsRow(wsBS, label, note);
        liabL += vL; liabR += vR;
    });
    bsTotal(wsBS, 'Total Sources of Funds', liabL, liabR);
    wsBS.addRow([]);

    bsSec(wsBS, 'II. APPLICATION OF FUNDS (ASSETS)');
    let assetL=0, assetR=0;
    [[11,'PPE (Fixed Assets)'],[12,'Investments'],[13,'Loans and Advances'],
     [14,'Non-current Assets'],[15,'Inventory'],[16,'Trade Receivables'],
     [17,'Cash and Bank Balance'],[18,'Other Current Assets']].forEach(([note,label]) => {
        const [vL,vR] = bsRow(wsBS, label, note);
        assetL += vL; assetR += vR;
    });
    bsTotal(wsBS, 'Total Application of Funds', assetL, assetR);

    // ── SHEET 2: Statement of P&L ─────────────────────────────────────────────
    const wsPL = wb.addWorksheet('Statement of P&L');
    wsPL.columns = [{ width:6 }, { width:52 }, { width:8 }, { width:22 }, { width:22 }];

    const plT = wsPL.addRow([`STATEMENT OF PROFIT AND LOSS FOR THE YEAR ENDED 31ST MARCH ${y1}`]);
    wsPL.mergeCells(1,1,1,5);
    plT.font = { bold:true, name:TN, size:13 };
    plT.getCell(1).alignment = { horizontal:'center' };
    wsPL.addRow([`${leftName} vs ${rightName}`]).font = { italic:true, name:TN };
    wsPL.addRow([]);

    const plHdr = wsPL.addRow(['', 'Particulars', 'Note No.', `For year ended 31 Mar ${y1} (Rs.)`, `For year ended 31 Mar ${y2} (Rs.)`]);
    plHdr.font = { bold:true, name:TN }; plHdr.height = 35;
    plHdr.eachCell(c => { c.border = BDR; c.alignment = { horizontal:'center', wrapText:true, vertical:'middle' }; });

    function plSec(ws, label) {
        const r = ws.addRow(['', label, '', '', '']);
        r.font = { bold:true, name:TN, size:11 };
        ws.mergeCells(ws.lastRow.number, 2, ws.lastRow.number, 5);
    }
    function plDataRow(ws, label, note) {
        const vL = nv(leftData,note), vR = nv(rightData,note);
        const r = ws.addRow(['', label, note, vL, vR]);
        r.font = { name:TN };
        r.getCell(4).numFmt = NUM; r.getCell(5).numFmt = NUM;
        r.eachCell({ includeEmpty:true }, c => c.border = BDR);
        return [vL, vR];
    }
    function plTotal(ws, label, vL, vR) {
        const r = ws.addRow(['', label, '', vL, vR]);
        r.font = { bold:true, name:TN };
        r.getCell(4).numFmt = NUM; r.getCell(5).numFmt = NUM;
        r.eachCell({ includeEmpty:true }, c => c.border = BDR);
    }

    plSec(wsPL, 'I. INCOME');
    const [r19L,r19R] = plDataRow(wsPL, 'Revenue from Operations', 19);
    const [r20L,r20R] = plDataRow(wsPL, 'Other Income', 20);
    const totIncL = r19L+r20L, totIncR = r19R+r20R;
    plTotal(wsPL, 'Total Income', totIncL, totIncR);
    wsPL.addRow([]);

    plSec(wsPL, 'II. EXPENDITURE');
    let expL=0, expR=0;
    [[21,'Cost of Goods Sold'],[22,'Employee Benefits Expense'],
     [23,'Finance Cost'],[24,'Depreciation and Amortisation'],[25,'Other Expenses']].forEach(([note,label]) => {
        const [vL,vR] = plDataRow(wsPL, label, note);
        expL += vL; expR += vR;
    });
    plTotal(wsPL, 'Total Expenditure', expL, expR);
    wsPL.addRow([]);
    plTotal(wsPL, 'III. Net Profit / (Loss)', totIncL-expL, totIncR-expR);

    // ── SHEET 3: Notes 1 to 3 ────────────────────────────────────────────────
    const ws3 = wb.addWorksheet('Note 1 to 3');
    sheetHeader(ws3);
    secTitle(ws3, 'Note 1 – Brief about the Entity');
    ws3.addRow(['(Narrative note – attach disclosure separately)']).font = { italic:true, name:TN };
    ws3.addRow([]);
    secTitle(ws3, 'Note 2 – Significant Accounting Policies');
    ws3.addRow(['(Narrative note – attach disclosure separately)']).font = { italic:true, name:TN };
    ws3.addRow([]);
    secTitle(ws3, "Note 3 – Owner's Capital Account");
    colHdr(ws3);
    uk(leftData?.summaryGroups, rightData?.summaryGroups).forEach(k =>
        dataRow(ws3, '  ' + k, leftData?.summaryGroups?.[k]||0, rightData?.summaryGroups?.[k]||0));
    dataRow(ws3, "Total Owner's Capital", nv(leftData,3), nv(rightData,3), true);

    // ── SHEET 4: Notes 4 to 6 ────────────────────────────────────────────────
    const ws4 = wb.addWorksheet('Notes to BS 4 to 6');
    sheetHeader(ws4);

    secTitle(ws4, 'Note 4 – Reserves and Surplus');
    colHdr(ws4);
    dataRow(ws4, 'Total Reserves and Surplus', nv(leftData,4), nv(rightData,4), true);
    ws4.addRow([]);

    secTitle(ws4, 'Note 5 – Borrowings');
    colHdr(ws4);
    ws4.addRow(['(A) Secured']).font = { bold:true, name:TN };
    uk(leftData?.loansInfo?.secured, rightData?.loansInfo?.secured).forEach(k =>
        dataRow(ws4, '  '+k, leftData?.loansInfo?.secured?.[k]||0, rightData?.loansInfo?.secured?.[k]||0));
    dataRow(ws4, 'Total Secured', leftData?.loansInfo?.secTotal||0, rightData?.loansInfo?.secTotal||0, true);
    ws4.addRow([]);
    ws4.addRow(['(B) Unsecured']).font = { bold:true, name:TN };
    uk(leftData?.loansInfo?.unsecured, rightData?.loansInfo?.unsecured).forEach(k =>
        dataRow(ws4, '  '+k, leftData?.loansInfo?.unsecured?.[k]||0, rightData?.loansInfo?.unsecured?.[k]||0));
    dataRow(ws4, 'Total Unsecured', leftData?.loansInfo?.unsecTotal||0, rightData?.loansInfo?.unsecTotal||0, true);
    ws4.addRow([]);
    dataRow(ws4, 'Total Borrowings', nv(leftData,5), nv(rightData,5), true);
    ws4.addRow([]);

    secTitle(ws4, 'Note 6 – Deferred Tax Liabilities');
    colHdr(ws4);
    dataRow(ws4, 'Total Deferred Tax Liabilities', nv(leftData,6), nv(rightData,6), true);

    // ── SHEET 5: Notes 7 to 10 ───────────────────────────────────────────────
    const ws5 = wb.addWorksheet('Notes to BS 7 to 10');
    sheetHeader(ws5);

    secTitle(ws5, 'Note 7 – Other Long-term Liabilities');
    colHdr(ws5);
    dataRow(ws5, 'Total Other Long-term Liabilities', nv(leftData,7), nv(rightData,7), true);
    ws5.addRow([]);

    secTitle(ws5, 'Note 8 – Provisions');
    colHdr(ws5);
    uk(leftData?.provisionsInfo?.shortTermItems, rightData?.provisionsInfo?.shortTermItems).forEach(k =>
        dataRow(ws5, '  '+k, leftData?.provisionsInfo?.shortTermItems?.[k]||0, rightData?.provisionsInfo?.shortTermItems?.[k]||0));
    const ptL=leftData?.provisionsInfo?.provisionForTaxation||0, ptR=rightData?.provisionsInfo?.provisionForTaxation||0;
    if(ptL||ptR) dataRow(ws5, '  Provision for Taxation', ptL, ptR);
    dataRow(ws5, 'Total Provisions', nv(leftData,8), nv(rightData,8), true);
    ws5.addRow([]);

    secTitle(ws5, 'Note 9 – Trade Payables');
    colHdr(ws5);
    uk(leftData?.sundryCredInfo?.ledgers, rightData?.sundryCredInfo?.ledgers).forEach(k =>
        dataRow(ws5, '  '+k, leftData?.sundryCredInfo?.ledgers?.[k]||0, rightData?.sundryCredInfo?.ledgers?.[k]||0));
    dataRow(ws5, 'Total Trade Payables', nv(leftData,9), nv(rightData,9), true);
    ws5.addRow([]);

    secTitle(ws5, 'Note 10 – Other Current Liabilities');
    colHdr(ws5);
    uk(leftData?.otherCurrLiabInfo?.items, rightData?.otherCurrLiabInfo?.items).forEach(k =>
        dataRow(ws5, '  '+k, leftData?.otherCurrLiabInfo?.items?.[k]||0, rightData?.otherCurrLiabInfo?.items?.[k]||0));
    dataRow(ws5, 'Total Other Current Liabilities', nv(leftData,10), nv(rightData,10), true);

    // ── SHEET 6: Note 11 – PPE ───────────────────────────────────────────────
    const ws6 = wb.addWorksheet('Notes to BS 11');
    sheetHeader(ws6);
    secTitle(ws6, 'Note 11 – Property, Plant & Equipment (PPE)');
    colHdr(ws6);
    dataRow(ws6, 'Gross Block (Opening Balance)', 0, 0);
    dataRow(ws6, 'Additions during the year', nv(leftData,11), nv(rightData,11));
    dataRow(ws6, 'Disposals / Deductions', 0, 0);
    dataRow(ws6, 'Gross Block (Closing Balance)', nv(leftData,11), nv(rightData,11));
    dataRow(ws6, 'Less: Accumulated Depreciation', nv(leftData,24), nv(rightData,24));
    if (leftData?.fixedAssetsInfo?.ledgers || rightData?.fixedAssetsInfo?.ledgers) {
        ws6.addRow([]);
        dataRow(ws6, 'PPE Ledger Breakup', '', '', true);
        uk(leftData?.fixedAssetsInfo?.ledgers, rightData?.fixedAssetsInfo?.ledgers).forEach(k =>
            dataRow(ws6, '  ' + k, leftData?.fixedAssetsInfo?.ledgers?.[k] || 0, rightData?.fixedAssetsInfo?.ledgers?.[k] || 0));
        dataRow(ws6, 'Total PPE Ledger Breakup', leftData?.fixedAssetsInfo?.total || 0, rightData?.fixedAssetsInfo?.total || 0, true);
    }
    dataRow(ws6, 'Net Block (Written Down Value)', nv(leftData,11), nv(rightData,11), true);

    // ── SHEET 7: Notes 12 to 18 ──────────────────────────────────────────────
    const ws7 = wb.addWorksheet('Notes to BS 12 to 18');
    sheetHeader(ws7);

    secTitle(ws7, 'Note 12 – Investments (Non-current & Current)');
    colHdr(ws7);
    uk(leftData?.investmentsInfo?.ledgers, rightData?.investmentsInfo?.ledgers).forEach(k =>
        dataRow(ws7, '  '+k, leftData?.investmentsInfo?.ledgers?.[k]||0, rightData?.investmentsInfo?.ledgers?.[k]||0));
    dataRow(ws7, 'Total Investments', nv(leftData,12), nv(rightData,12), true);
    ws7.addRow([]);

    secTitle(ws7, 'Note 13 – Loans and Advances');
    colHdr(ws7);
    uk(leftData?.advancesInfo?.items, rightData?.advancesInfo?.items).forEach(k =>
        dataRow(ws7, '  '+k, leftData?.advancesInfo?.items?.[k]||0, rightData?.advancesInfo?.items?.[k]||0));
    dataRow(ws7, 'Total Loans and Advances', nv(leftData,13), nv(rightData,13), true);
    ws7.addRow([]);

    secTitle(ws7, 'Note 14 – Non-current Assets');
    colHdr(ws7);
    dataRow(ws7, 'Total Non-current Assets', nv(leftData,14), nv(rightData,14), true);
    ws7.addRow([]);

    secTitle(ws7, 'Note 15 – Inventory');
    colHdr(ws7);
    dataRow(ws7, 'Total Inventory', nv(leftData,15), nv(rightData,15), true);
    ws7.addRow([]);

    secTitle(ws7, 'Note 16 – Trade Receivables');
    colHdr(ws7);
    uk(leftData?.sundryDebtorsInfo?.ledgers, rightData?.sundryDebtorsInfo?.ledgers).forEach(k =>
        dataRow(ws7, '  '+k, leftData?.sundryDebtorsInfo?.ledgers?.[k]||0, rightData?.sundryDebtorsInfo?.ledgers?.[k]||0));
    dataRow(ws7, 'Total Trade Receivables', nv(leftData,16), nv(rightData,16), true);
    ws7.addRow([]);

    secTitle(ws7, 'Note 17 – Cash and Bank Balance');
    colHdr(ws7);
    dataRow(ws7, 'Balances with Banks', leftData?.cashInfo?.bankAccounts||0, rightData?.cashInfo?.bankAccounts||0);
    dataRow(ws7, 'Cash on Hand',        leftData?.cashInfo?.cashInHand||0,   rightData?.cashInfo?.cashInHand||0);
    dataRow(ws7, 'Total Cash and Bank Balance', nv(leftData,17), nv(rightData,17), true);
    ws7.addRow([]);

    secTitle(ws7, 'Note 18 – Other Current Assets');
    colHdr(ws7);
    dataRow(ws7, 'Total Other Current Assets', nv(leftData,18), nv(rightData,18), true);

    // ── SHEET 8: Notes 19 to 25 ──────────────────────────────────────────────
    const ws8 = wb.addWorksheet('Notes to P&L 19 to 25');
    sheetHeader(ws8);

    secTitle(ws8, 'Note 19 – Revenue from Operations');
    colHdr(ws8, true);
    dataRow(ws8, 'Total Revenue from Operations', nv(leftData,19), nv(rightData,19), true);
    ws8.addRow([]);

    secTitle(ws8, 'Note 20 – Other Income');
    colHdr(ws8, true);
    uk(leftData?.indirectIncomesInfo?.items, rightData?.indirectIncomesInfo?.items).forEach(k =>
        dataRow(ws8, '  '+k, leftData?.indirectIncomesInfo?.items?.[k]||0, rightData?.indirectIncomesInfo?.items?.[k]||0));
    dataRow(ws8, 'Total Other Income', nv(leftData,20), nv(rightData,20), true);
    ws8.addRow([]);

    secTitle(ws8, 'Note 21 – Cost of Goods Sold');
    colHdr(ws8, true);
    dataRow(ws8, 'Total Cost of Goods Sold', nv(leftData,21), nv(rightData,21), true);
    ws8.addRow([]);

    secTitle(ws8, 'Note 22 – Employee Benefits Expense');
    colHdr(ws8, true);
    uk(leftData?.note16Info?.items, rightData?.note16Info?.items).forEach(k =>
        dataRow(ws8, '  '+k, leftData?.note16Info?.items?.[k]||0, rightData?.note16Info?.items?.[k]||0));
    dataRow(ws8, 'Total Employee Benefits Expense', nv(leftData,22), nv(rightData,22), true);
    ws8.addRow([]);

    secTitle(ws8, 'Note 23 – Finance Cost');
    colHdr(ws8, true);
    dataRow(ws8, 'Total Finance Cost', nv(leftData,23), nv(rightData,23), true);
    ws8.addRow([]);

    secTitle(ws8, 'Note 24 – Depreciation and Amortisation Expense');
    colHdr(ws8, true);
    dataRow(ws8, 'Depreciation on PPE', nv(leftData,24), nv(rightData,24));
    dataRow(ws8, 'Total Depreciation', nv(leftData,24), nv(rightData,24), true);
    ws8.addRow([]);

    secTitle(ws8, 'Note 25 – Other Expenses');
    colHdr(ws8, true);
    dataRow(ws8, 'Total Other Expenses', nv(leftData,25), nv(rightData,25), true);

    const audit = wb.addWorksheet('Tally Data Audit');
    audit.columns=[{width:20},{width:45},{width:22},{width:22}];
    audit.addRow(['Note','Particular',`31 Mar ${leftYear}`,`31 Mar ${rightYear}`]).font={bold:true};
    for(let n=1;n<=25;n++){ const l=nv(leftData,n)||0, r=nv(rightData,n)||0; audit.addRow([`Note ${n}`,'Schedule III total',l,r]); }
    const groups=[['Note 3 Secured Loans',leftData?.loansInfo?.secured,rightData?.loansInfo?.secured],['Note 3 Unsecured Loans',leftData?.loansInfo?.unsecured,rightData?.loansInfo?.unsecured],['Note 6 Trade Payables',leftData?.sundryCredInfo?.ledgers,rightData?.sundryCredInfo?.ledgers],['Note 11 Sundry Debtors',leftData?.sundryDebtorsInfo?.ledgers,rightData?.sundryDebtorsInfo?.ledgers],['Note 20 Other Income',leftData?.indirectIncomesInfo?.items,rightData?.indirectIncomesInfo?.items],['Note 22 Employee Benefits',leftData?.note16Info?.items,rightData?.note16Info?.items]];
    for(const [label,a,b] of groups){ for(const k of [...new Set([...Object.keys(a||{}),...Object.keys(b||{})])].sort()) audit.addRow([label,k,a?.[k]||0,b?.[k]||0]); }
    audit.eachRow(r=>r.eachCell(c=>{if(typeof c.value==='number')c.numFmt='#,##0.00;[Red]-#,##0.00';}));
    await wb.xlsx.writeFile(outputPath);
    console.log(`[Partnership/Prop/HUF] Fresh Excel saved: ${outputPath}`);
}

module.exports = { generateMergedExcelPartnership };
