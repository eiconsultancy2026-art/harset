const { normalizeAmount } = require('./xml_parser_utils');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

async function extractMapping() {
    const sessionId = process.env.SESSION_ID || '';
    const responsesDir = sessionId ? path.join(__dirname, 'responses', sessionId) : path.join(__dirname, 'responses');

    const bsXml = fs.readFileSync(path.join(responsesDir, 'balance_sheet_response.xml'), 'utf8');
    const plXml = fs.readFileSync(path.join(responsesDir, 'profit_loss_response.xml'), 'utf8');

    let bsSummaryXml = "";
    try { bsSummaryXml = fs.readFileSync(path.join(responsesDir, 'balance_sheet_summary.xml'), 'utf8'); } catch (e) { }

    let loansXml = "";
    try { loansXml = fs.readFileSync(path.join(responsesDir, 'loans_liability_response.xml'), 'utf8'); } catch (e) { }

    let currLiabXml = "";
    try { currLiabXml = fs.readFileSync(path.join(responsesDir, 'current_liabilities_response.xml'), 'utf8'); } catch (e) { }

    let currAssetsXml = "";
    try { currAssetsXml = fs.readFileSync(path.join(responsesDir, 'current_assets_response.xml'), 'utf8'); } catch (e) { }

    let indirectExpXml = "";
    try { indirectExpXml = fs.readFileSync(path.join(responsesDir, 'indirect_expenses_response.xml'), 'utf8'); } catch (e) { }

    let sundryDebtorsXml = "";
    try { sundryDebtorsXml = fs.readFileSync(path.join(responsesDir, 'sundry_debtors_response.xml'), 'utf8'); } catch (e) { }

    let sundryCreditorsXml = "";
    try { sundryCreditorsXml = fs.readFileSync(path.join(responsesDir, 'sundry_creditors_response.xml'), 'utf8'); } catch (e) { }

    let output = "";
    const log = (msg = "") => {
        console.log(msg);
        output += msg + "\n";
    };

    const traceLog = [];
    const trace = (note, source, variable, value, logic = "", breakdown = null) => {
        const entry = { 
            Note: note, 
            Source: source, 
            Tally_Variable: variable, 
            Value: value, 
            Logic: logic 
        };
        if (breakdown && Object.keys(breakdown).length > 0) {
            entry.Breakdown = breakdown;
        }
        traceLog.push(entry);
    };
    // ---------------------------------------------------------------
    // Parses balance_sheet_summary.xml to get authoritative Tally totals.
    // ---------------------------------------------------------------
    const SUMMARY_LIABILITY_GROUPS = new Set([
        'Capital Account', 'Loans (Liability)', 'Current Liabilities',
        'Reserves & Surplus', 'Profit & Loss A/c',
        'Excess of income over expenditure', 'Suspense A/c',
        'Branch / Divisions'
    ]);
    const SUMMARY_ASSET_GROUPS = new Set([
        'Fixed Assets', 'Investments', 'Current Assets',
        'Misc. Expenses (Asset)', 'Loans & Advances (Asset)'
    ]);

    function parseSummaryTotals(xmlStr) {
        if (!xmlStr) return { liabilitiesTotal: 0, assetsTotal: 0, groups: {} };
        const groups = {};
        const chunks = xmlStr.split('<BSNAME>');
        for (let i = 1; i < chunks.length; i++) {
            const chunk = chunks[i];
            const nameMatch = chunk.match(/<DSPDISPNAME>([^<]+)<\/DSPDISPNAME>/);
            if (!nameMatch) continue;
            const name = nameMatch[1].trim().replace(/&amp;/g, '&');
            const amtMatch = chunk.match(/<BSMAINAMT>([-0-9.]+)<\/BSMAINAMT>/);
            const val = amtMatch ? normalizeAmount(amtMatch[1]) : 0;
            if (val !== 0) groups[name] = val;
        }
        let liabilitiesTotal = 0;
        let assetsTotal = 0;
        for (const [name, val] of Object.entries(groups)) {
            if (SUMMARY_LIABILITY_GROUPS.has(name)) liabilitiesTotal += Math.abs(val);
            else if (SUMMARY_ASSET_GROUPS.has(name)) assetsTotal += Math.abs(val);
        }
        if (liabilitiesTotal === 0 && assetsTotal > 0) liabilitiesTotal = assetsTotal;
        if (assetsTotal === 0 && liabilitiesTotal > 0) assetsTotal = liabilitiesTotal;
        return { liabilitiesTotal, assetsTotal, groups };
    }

    function parseValuesEasier(xmlStr) {
        const results = {};
        const names = xmlStr.split('<DSPDISPNAME>');
        for (let i = 1; i < names.length; i++) {
            const namePart = names[i];
            let name = namePart.split('</DSPDISPNAME>')[0].replace(/&#13;&#10;/g, '').trim();
            name = name.replace(/&amp;/g, '&');

            let val = 0;
            const subAmtMatch = namePart.match(/<BSSUBAMT>([-0-9.]+)<\/BSSUBAMT>/);
            const mainAmtMatch = namePart.match(/<BSMAINAMT>([-0-9.]+)<\/BSMAINAMT>/);
            const plSubMatch = namePart.match(/<PLSUBAMT>([-0-9.]+)<\/PLSUBAMT>/);

            if (subAmtMatch) val = normalizeAmount(subAmtMatch[1]);
            else if (mainAmtMatch) val = normalizeAmount(mainAmtMatch[1]);
            else if (plSubMatch) val = normalizeAmount(plSubMatch[1]);

            results[name] = val;
        }
        return results;
    }

    function parseLoans(xmlStr) {
        if (!xmlStr) return null;
        const names = xmlStr.split('<DSPDISPNAME>');
        let unsecuredTotal = null;
        let runningSum = 0;

        const secured = {};
        const unsecured = {};
        let inUnsecuredGroup = false;

        for (let i = 1; i < names.length; i++) {
            const namePart = names[i];
            let name = namePart.split('</DSPDISPNAME>')[0].trim();
            name = name.replace(/&amp;/g, '&');

            let val = 0;
            const crcMatch = namePart.match(/<DSPCLCRAMTA>([-0-9.]+)<\/DSPCLCRAMTA>/);
            const drMatch = namePart.match(/<DSPCLDRAMTA>([-0-9.]+)<\/DSPCLDRAMTA>/);
            if (crcMatch) val = normalizeAmount(crcMatch[1]);
            else if (drMatch) val = normalizeAmount(drMatch[1]);

            if (name === "Unsecured Loans") {
                unsecuredTotal = val;
                inUnsecuredGroup = true;
                runningSum = 0;
                continue;
            }

            if (inUnsecuredGroup) {
                unsecured[name] = val;
                runningSum += val;
                if (Math.abs(runningSum - unsecuredTotal) < 0.01) {
                    inUnsecuredGroup = false;
                }
            } else {
                secured[name] = val;
            }
        }

        let secTotal = Object.values(secured).reduce((a, b) => a + b, 0);
        let unsecTotal = Object.values(unsecured).reduce((a, b) => a + b, 0);

        return { secured, unsecured, secTotal, unsecTotal, total: secTotal + unsecTotal };
    }

    function parseProvisions(xmlStr, mainData) {
        if (!xmlStr) return null;
        const names = xmlStr.split('<DSPDISPNAME>');
        let shortTermTotal = 0;
        let runningSum = 0;

        const shortTermItems = {};
        let inShortTerm = false;
        let provisionForTaxation = 0;

        for (let i = 1; i < names.length; i++) {
            const namePart = names[i];
            let name = namePart.split('</DSPDISPNAME>')[0].trim();
            name = name.replace(/&amp;/g, '&');

            let val = 0;
            const crcMatch = namePart.match(/<DSPCLCRAMTA>([-0-9.]+)<\/DSPCLCRAMTA>/);
            const drMatch = namePart.match(/<DSPCLDRAMTA>([-0-9.]+)<\/DSPCLDRAMTA>/);
            if (crcMatch) val = normalizeAmount(crcMatch[1]);
            else if (drMatch) val = normalizeAmount(drMatch[1]);

            if (name === "Short Term Provisions" || name === "Provisions") {
                shortTermTotal = val;
                inShortTerm = true;
                runningSum = 0;
                continue;
            }

            if (inShortTerm) {
                shortTermItems[name] = val;
                runningSum += val;
                if (Math.abs(runningSum - shortTermTotal) < 0.01) {
                    inShortTerm = false;
                }
            } else if (name === "Income Tax Payable" || name.toLowerCase().includes("provision for taxation")) {
                provisionForTaxation += val;
            }
        }

        // if (mainData['Income Tax AY 2025-26']) {
        //     provisionForTaxation += Math.abs(mainData['Income Tax AY 2025-26']);
        // }

        return { shortTermItems, provisionForTaxation, total: shortTermTotal + provisionForTaxation };
    }

    // ---------------------------------------------------------------
    // Parses Note 6 – Trade Payables (Sundry Creditors breakdown)
    // Enters collection mode on the "Sundry Creditors" group header
    // and collects every child ledger until the next sibling group.
    // ---------------------------------------------------------------
    const CURR_LIAB_TOP_GROUPS = new Set([
        'Duties & Taxes', 'Sundry Creditors', 'Other Current Liabilities',
        'Short Term Provisions'
    ]);

    function parseSundryCreditors(xmlStr) {
        if (!xmlStr) return null;

        const blocks = xmlStr.split('<DSPACCNAME>');
        const ledgers = {};
        
        for (let i = 1; i < blocks.length; i++) {
            const block = blocks[i];
            const nameMatch = block.match(/<DSPDISPNAME>([^<]+)<\/DSPDISPNAME>/);
            if (!nameMatch) continue;
            const name = nameMatch[1].trim().replace(/&amp;/g, '&');

            const infoStart = block.indexOf('<DSPACCINFO>');
            if (infoStart === -1) continue;
            const ib = block.slice(infoStart);

            const drMatch = ib.match(/<DSPCLDRAMTA>([-0-9.]+)<\/DSPCLDRAMTA>/);
            const crMatch = ib.match(/<DSPCLCRAMTA>([-0-9.]+)<\/DSPCLCRAMTA>/);
            const dr = drMatch ? Math.abs(normalizeAmount(drMatch[1])) : 0;
            const cr = crMatch ? Math.abs(normalizeAmount(crMatch[1])) : 0;
            const net = cr - dr;

            if (net === 0) continue;
            ledgers[name] = net;
        }

        const total = Object.values(ledgers).reduce((a, b) => a + b, 0);
        return { ledgers, total };
    }

    // ---------------------------------------------------------------
    // Parses Note 7 – Other Current Liabilities (itemised)
    // Ledgers belonging to Sundry Creditors are explicitly excluded
    // so they don't bleed into Note 7.
    // Ledgers/groups to skip entirely — group headers and Short-Term Provisions
    // NOTE: 'Professional Tax Payable' is intentionally NOT skipped here.
    // It can carry a debit balance that must offset the group total.
    const NOTE7_SKIP = new Set([
        'Sundry Creditors', 'Short Term Provisions', 'Other Current Liabilities',
        'Duties & Taxes', 'Provisions'
    ]);
    const PROVISION_NAMES = new Set([
        'Provision for Accounting Charges', 'Provision for Audit Fees',
        'Provision for Secretarial Charges'
    ]);

    function parseOtherCurrentLiabilities(xmlStr) {
        if (!xmlStr) return null;

        const blocks = xmlStr.split('<DSPACCNAME>');
        const items = {};
        let insideExcludedGroup = false;

        for (let i = 1; i < blocks.length; i++) {
            const block = blocks[i];
            const nameMatch = block.match(/<DSPDISPNAME>([^<]+)<\/DSPDISPNAME>/);
            if (!nameMatch) continue;
            const name = nameMatch[1].trim().replace(/&amp;/g, '&');

            // Track entry into Sundry Creditors sub-group
            if (name === 'Sundry Creditors' || name === 'Short Term Provisions' || name === 'Provisions') {
                insideExcludedGroup = true;
                continue;
            }

            // Exit Sundry Creditors mode at the next top-level group
            if (insideExcludedGroup && CURR_LIAB_TOP_GROUPS.has(name)) {
                insideExcludedGroup = false;
                // fall through — handle this new group header below
            }

            // Skip Sundry Creditor child ledgers
            if (insideExcludedGroup) continue;

            // Skip group headers, provisions, and unwanted ledgers
            if (NOTE7_SKIP.has(name) || PROVISION_NAMES.has(name)) continue;
            if (name === 'Income Tax Payable') continue;

            const infoStart = block.indexOf('<DSPACCINFO>');
            if (infoStart === -1) continue;
            const infoBlock = block.slice(infoStart);

            let val = 0;
            const crMatch = infoBlock.match(/<DSPCLCRAMTA>([-0-9.]+)<\/DSPCLCRAMTA>/);
            const drMatch = infoBlock.match(/<DSPCLDRAMTA>([-0-9.]+)<\/DSPCLDRAMTA>/);
            if (crMatch) val = Math.abs(normalizeAmount(crMatch[1]));
            // DSPCLDRAMTA is already stored as a negative value for DR-balance items in a liability
            // group (e.g. Professional Tax Payable = -3240). Using parseFloat directly preserves
            // the negative sign so it correctly nets against the credit items in the total.
            else if (drMatch) val = normalizeAmount(drMatch[1]);

            if (val === 0) continue;
            // Aggregate multiple entries (e.g. netting DR and CR for the same ledger name)
            items[name] = (items[name] || 0) + val;
        }

        const total = Object.values(items).reduce((a, b) => a + b, 0);
        return { items, total };
    }

    // ---------------------------------------------------------------
    // Known top-level group names that mark the END of Current Asset sub-groups
    const CURRENT_ASSET_GROUPS = new Set([
        'Opening Stock', 'Sundry Debtors', 'Cash-in-Hand', 'Cash-in-hand', 'Bank Accounts',
        'Advances & Deposit', 'Loans & Advances (Asset)', 'Deposits (Assets)', 'Deposits (Asset)', 'GST ITC', 'GST',
        'Other Current Assets'
    ]);

    // ---------------------------------------------------------------
    // Parses Note 11 – Sundry Debtors (child ledger breakdown)
    // Walks the current-assets XML, enters collection mode at the
    // "Sundry Debtors" group header and harvests child ledgers until
    // the next sibling top-level group.
    // ---------------------------------------------------------------
    function parseSundryDebtors(xmlStr) {
        if (!xmlStr) return null;

        const blocks = xmlStr.split('<DSPACCNAME>');
        const ledgers = {};

        for (let i = 1; i < blocks.length; i++) {
            const block = blocks[i];
            const nameMatch = block.match(/<DSPDISPNAME>([^<]+)<\/DSPDISPNAME>/);
            if (!nameMatch) continue;
            const name = nameMatch[1].trim().replace(/&amp;/g, '&');

            const infoStart = block.indexOf('<DSPACCINFO>');
            if (infoStart === -1) continue;
            const ib = block.slice(infoStart);

            const drMatch = ib.match(/<DSPCLDRAMTA>([-0-9.]+)<\/DSPCLDRAMTA>/);
            const crMatch = ib.match(/<DSPCLCRAMTA>([-0-9.]+)<\/DSPCLCRAMTA>/);
            const dr = drMatch ? Math.abs(normalizeAmount(drMatch[1])) : 0;
            const cr = crMatch ? Math.abs(normalizeAmount(crMatch[1])) : 0;
            const net = dr - cr;

            if (net === 0) continue;
            ledgers[name] = net;
        }

        const total = Object.values(ledgers).reduce((a, b) => a + b, 0);
        return { ledgers, total };
    }

    // ---------------------------------------------------------------
    // Parses Note 10 – Investments (Deposits (Assets) breakdown)
    // ---------------------------------------------------------------
    function parseInvestments(xmlStr) {
        if (!xmlStr) return null;

        const blocks = xmlStr.split('<DSPACCNAME>');
        const ledgers = {};
        let groupTotal = 0;
        let collecting = false;

        for (let i = 1; i < blocks.length; i++) {
            const block = blocks[i];
            const nameMatch = block.match(/<DSPDISPNAME>([^<]+)<\/DSPDISPNAME>/);
            if (!nameMatch) continue;
            const name = nameMatch[1].trim().replace(/&amp;/g, '&');

            if (name === 'Deposits (Assets)' || name === 'Deposits (Asset)') {
                const infoStart = block.indexOf('<DSPACCINFO>');
                if (infoStart !== -1) {
                    const ib = block.slice(infoStart);
                    const drMatch = ib.match(/<DSPCLDRAMTA>([-0-9.]+)<\/DSPCLDRAMTA>/);
                    const crMatch = ib.match(/<DSPCLCRAMTA>([-0-9.]+)<\/DSPCLCRAMTA>/);
                    const dr = drMatch ? Math.abs(normalizeAmount(drMatch[1])) : 0;
                    const cr = crMatch ? Math.abs(normalizeAmount(crMatch[1])) : 0;
                    groupTotal = dr - cr;
                }
                collecting = true;
                continue;
            }

            if (collecting && CURRENT_ASSET_GROUPS.has(name)) {
                collecting = false;
                break;
            }

            if (!collecting) continue;
            
            // Skip the "Term Deposit" subgroup header itself so we only capture its child ledgers
            if (name === 'Term Deposit') continue;

            const infoStart = block.indexOf('<DSPACCINFO>');
            if (infoStart === -1) continue;
            const ib = block.slice(infoStart);

            const drMatch = ib.match(/<DSPCLDRAMTA>([-0-9.]+)<\/DSPCLDRAMTA>/);
            const crMatch = ib.match(/<DSPCLCRAMTA>([-0-9.]+)<\/DSPCLCRAMTA>/);
            const dr = drMatch ? Math.abs(normalizeAmount(drMatch[1])) : 0;
            const cr = crMatch ? Math.abs(normalizeAmount(crMatch[1])) : 0;
            const net = dr - cr;

            if (net === 0) continue;
            ledgers[name] = net;
        }

        const total = groupTotal || Object.values(ledgers).reduce((a, b) => a + b, 0);
        return { ledgers, total };
    }

    // ---------------------------------------------------------------
    // Parses Note 9 – Advances & Deposit (itemised)
    // Enters collection mode on the "Advances & Deposit" group header
    // and collects every child ledger until the next sibling group.
    // ---------------------------------------------------------------

    function parseAdvancesAndDeposit(xmlStr) {
        if (!xmlStr) return null;

        const blocks = xmlStr.split('<DSPACCNAME>');
        const items = {};
        let collecting = false;
        let groupTotal = 0;

        for (let i = 1; i < blocks.length; i++) {
            const block = blocks[i];

            const nameMatch = block.match(/<DSPDISPNAME>([^<]+)<\/DSPDISPNAME>/);
            if (!nameMatch) continue;
            const name = nameMatch[1].trim().replace(/&amp;/g, '&');

            // Support both Tally group naming conventions for long-term advances
            if (name === 'Advances & Deposit' || name === 'Loans & Advances (Asset)') {
                // This is the group header — grab its DR value as the group total
                const infoBlock = block.slice(block.indexOf('<DSPACCINFO>'));
                const drMatch = infoBlock.match(/<DSPCLDRAMTA>([-0-9.]+)<\/DSPCLDRAMTA>/);
                if (drMatch) groupTotal = Math.abs(normalizeAmount(drMatch[1]));
                collecting = true;
                continue;
            }

            // Stop collecting when we hit the next sibling group
            if (collecting && CURRENT_ASSET_GROUPS.has(name)) {
                collecting = false;
                break;
            }

            if (!collecting) continue;

            // Extract value from the DSPACCINFO block
            const infoStart = block.indexOf('<DSPACCINFO>');
            if (infoStart === -1) continue;
            const infoBlock = block.slice(infoStart);

            let val = 0;
            const crMatch = infoBlock.match(/<DSPCLCRAMTA>([-0-9.]+)<\/DSPCLCRAMTA>/);
            const drMatch = infoBlock.match(/<DSPCLDRAMTA>([-0-9.]+)<\/DSPCLDRAMTA>/);
            if (crMatch) val = normalizeAmount(crMatch[1]);
            else if (drMatch) val = Math.abs(normalizeAmount(drMatch[1]));

            // Include even zero-value items so they show as "-" like in the image
            items[name] = val;
        }

        const total = groupTotal || Object.values(items).reduce((a, b) => a + b, 0);
        return { items, total };
    }

    // ---------------------------------------------------------------
    // Parses Note 12 – Cash & Cash Equivalents
    // Extracts the group totals for Bank Accounts and Cash-in-Hand
    // and presents them as fixed rows matching the Schedule 3 format.
    // ---------------------------------------------------------------
    function parseCashAndEquivalents(xmlStr) {
        if (!xmlStr) return null;

        const blocks = xmlStr.split('<DSPACCNAME>');
        let bankAccounts = 0;
        let cashInHand = 0;

        for (let i = 1; i < blocks.length; i++) {
            const block = blocks[i];

            const nameMatch = block.match(/<DSPDISPNAME>([^<]+)<\/DSPDISPNAME>/);
            if (!nameMatch) continue;
            const name = nameMatch[1].trim().replace(/&amp;/g, '&');

            // We only care about the two group-header rows
            if (name !== 'Bank Accounts' && name !== 'Cash-in-Hand' && name !== 'Cash-in-hand') continue;

            const infoStart = block.indexOf('<DSPACCINFO>');
            if (infoStart === -1) continue;
            const infoBlock = block.slice(infoStart);

            let val = 0;
            const crMatch = infoBlock.match(/<DSPCLCRAMTA>([-0-9.]+)<\/DSPCLCRAMTA>/);
            const drMatch = infoBlock.match(/<DSPCLDRAMTA>([-0-9.]+)<\/DSPCLDRAMTA>/);
            if (crMatch) val = normalizeAmount(crMatch[1]);
            else if (drMatch) val = Math.abs(normalizeAmount(drMatch[1]));

            if (name === 'Bank Accounts') bankAccounts = val;
            if (name === 'Cash-in-Hand' || name === 'Cash-in-hand') cashInHand = val;
        }

        const total = bankAccounts + cashInHand;
        return { bankAccounts, cashInHand, total };
    }

    // ---------------------------------------------------------------
    // Parses GST ITC group total from current assets XML
    // Reads the 'GST ITC' group header's net DR balance directly so the
    // correct asset value is used (DR - CR), as opposed to the BSMAINAMT
    // tag in the BS XML which may differ due to rounding or timing.
    // ---------------------------------------------------------------
    function parseGstItc(xmlStr) {
        if (!xmlStr) return 0;
        const blocks = xmlStr.split('<DSPACCNAME>');
        for (let i = 1; i < blocks.length; i++) {
            const block = blocks[i];
            const nameMatch = block.match(/<DSPDISPNAME>([^<]+)<\/DSPDISPNAME>/);
            if (!nameMatch) continue;
            const name = nameMatch[1].trim().replace(/&amp;/g, '&');
            if (name !== 'GST ITC') continue;
            const infoStart = block.indexOf('<DSPACCINFO>');
            if (infoStart === -1) continue;
            const ib = block.slice(infoStart);
            const drMatch = ib.match(/<DSPCLDRAMTA>([-0-9.]+)<\/DSPCLDRAMTA>/);
            const crMatch = ib.match(/<DSPCLCRAMTA>([-0-9.]+)<\/DSPCLCRAMTA>/);
            const dr = drMatch ? Math.abs(normalizeAmount(drMatch[1])) : 0;
            const cr = crMatch ? Math.abs(normalizeAmount(crMatch[1])) : 0;
            return dr - cr; // net DR = asset value
        }
        return 0;
    }

    // ---------------------------------------------------------------
    // Parses Note 15 – Other Income (Indirect Incomes, itemised)
    // The P&L XML uses a different structure:
    //   Group header: <DSPACCNAME><DSPDISPNAME>...</DSPDISPNAME></DSPACCNAME>
    //                 followed by <PLAMT><BSMAINAMT>total</BSMAINAMT></PLAMT>
    //   Children:     <BSNAME><DSPACCNAME><DSPDISPNAME>...</DSPDISPNAME>...
    //                 followed by <BSAMT><BSSUBAMT>value</BSSUBAMT></BSAMT>
    // We collect BSNAME children between Indirect Incomes and the
    // next top-level DSPACCNAME group.
    // ---------------------------------------------------------------
    function parseIndirectIncomes(xmlStr) {
        if (!xmlStr) return null;

        // Split on the raw text to find the Indirect Incomes section
        const markerTop = '<DSPDISPNAME>Indirect Incomes</DSPDISPNAME>';
        const startIdx = xmlStr.indexOf(markerTop);
        if (startIdx === -1) return null;

        // Find where the next top-level group starts.
        // Top-level groups appear as "\n <DSPACCNAME>" (newline + 1 space).
        // Children are nested inside <BSNAME> as "   <DSPACCNAME>" (3+ spaces),
        // so we must NOT match those. We use a regex to find only the top-level pattern.
        const afterHeader = xmlStr.indexOf('</PLAMT>', startIdx) + 8; // skip past PLAMT
        const topLevelPattern = /\n <DSPACCNAME>/g;
        topLevelPattern.lastIndex = afterHeader;
        const nextTopMatch = topLevelPattern.exec(xmlStr);
        const nextGroupIdx = nextTopMatch ? nextTopMatch.index : -1;
        const section = nextGroupIdx === -1
            ? xmlStr.slice(afterHeader)
            : xmlStr.slice(afterHeader, nextGroupIdx);

        // Extract group total from BSMAINAMT in the PLAMT just above afterHeader
        const plSection = xmlStr.slice(startIdx, afterHeader);
        const mainAmtMatch = plSection.match(/<BSMAINAMT>([-0-9.]+)<\/BSMAINAMT>/);
        const groupTotal = mainAmtMatch ? Math.abs(normalizeAmount(mainAmtMatch[1])) : 0;

        // Extract each BSNAME child
        const items = {};
        const bsChunks = section.split('<BSNAME>');
        for (let i = 1; i < bsChunks.length; i++) {
            const chunk = bsChunks[i];

            const nameMatch = chunk.match(/<DSPDISPNAME>([^<]+)<\/DSPDISPNAME>/);
            if (!nameMatch) continue;
            const name = nameMatch[1].trim().replace(/&amp;/g, '&');

            let val = 0;
            const subAmtMatch = chunk.match(/<BSSUBAMT>([-0-9.]+)<\/BSSUBAMT>/);
            if (subAmtMatch) val = Math.abs(normalizeAmount(subAmtMatch[1]));

            items[name] = val;
        }

        const total = groupTotal || Object.values(items).reduce((a, b) => a + b, 0);
        return { items, total };
    }

    // ---------------------------------------------------------------
    // Parses Note 16 – Employee Benefit Expenses
    // Reads from indirect_expenses_response.xml which uses a flat
    // <DSPACCNAME>/<DSPACCINFO> structure (no <BSNAME> nesting).
    // We collect only the known employee-benefit ledgers.
    // ---------------------------------------------------------------
    // Employee-benefit ledgers are not always named consistently in Tally.
    // Capture salary/wages/PF/leave/bonus/staff-welfare variants instead of
    // relying on a four-ledger allow-list that silently converts new ledgers to 0.
    const EMPLOYEE_BENEFIT_PATTERNS = [
        /salary/i, /wages?/i, /staff\s*welfare/i, /provident\s*fund/i, /\bpf\b/i,
        /leave\s*salary/i, /gratuity/i, /employee\s*benefit/i, /bonus/i
    ];
    const NOTE16_DISPLAY_NAMES = {
        'Salaries and Wages': 'Salaries and wages',
        'Salary': 'Salaries and wages',
        'Staff Welfare Expenses': 'Staff Welfare',
        'Staff  Welfare Expenses': 'Staff Welfare',
        'Management Contribution - PF': 'PF - Management Contribution'
    };
    function isEmployeeBenefitLedger(name) {
        const n = String(name || '').trim();
        return EMPLOYEE_BENEFIT_PATTERNS.some(re => re.test(n));
    }

    function parseNote16(xmlStr) {
        if (!xmlStr) return null;
        const blocks = xmlStr.split('<DSPACCNAME>');
        const items = {};
        for (let i = 1; i < blocks.length; i++) {
            const block = blocks[i];
            const nameMatch = block.match(/<DSPDISPNAME>([^<]*)<\/DSPDISPNAME>/i);
            if (!nameMatch) continue;
            const rawName = nameMatch[1].trim().replace(/&amp;/g, '&');
            if (!isEmployeeBenefitLedger(rawName)) continue;
            const infoStart = block.indexOf('<DSPACCINFO>');
            if (infoStart === -1) continue;
            const infoBlock = block.slice(infoStart);
            const drMatch = infoBlock.match(/<DSPCLDRAMTA>([^<]*)<\/DSPCLDRAMTA>/i);
            const crMatch = infoBlock.match(/<DSPCLCRAMTA>([^<]*)<\/DSPCLCRAMTA>/i);
            const dr = drMatch ? Math.abs(normalizeAmount(drMatch[1])) : 0;
            const cr = crMatch ? Math.abs(normalizeAmount(crMatch[1])) : 0;
            const val = Math.abs(dr - cr);
            if (val === 0) continue;
            const displayName = NOTE16_DISPLAY_NAMES[rawName] || rawName;
            items[displayName] = (items[displayName] || 0) + val;
        }
        const total = Object.values(items).reduce((a,b) => a+b, 0);
        return { items, total };
    }

    const bsData = parseValuesEasier(bsXml);
    const plData = parseValuesEasier(plXml);
    const data = { ...bsData, ...plData };

    function parseFixedAssets(xmlStr) {
        if (!xmlStr) return { ledgers: {}, total: 0 };
        const marker = /<DSPDISPNAME>\s*Fixed Assets\s*<\/DSPDISPNAME>/i.exec(xmlStr);
        if (!marker) return { ledgers: {}, total: 0 };
        const tail = xmlStr.slice(marker.index + marker[0].length);
        const nextGroup = /<DSPDISPNAME>\s*(?:Current Assets|Current Liabilities|Loans \(Liability\)|Capital Account|Reserves & Surplus|Suspense A\/c)\s*<\/DSPDISPNAME>/i.exec(tail);
        const section = nextGroup ? tail.slice(0, nextGroup.index) : tail;
        const ledgers = {};
        const rowRe = /<DSPDISPNAME>([\s\S]*?)<\/DSPDISPNAME>([\s\S]*?)(?=<DSPDISPNAME>|$)/gi;
        let row;
        while ((row = rowRe.exec(section))) {
            const name = row[1].replace(/<!\[CDATA\[|\]\]>/g, '').replace(/&amp;/g, '&').trim();
            if (!name || /^Fixed Assets$/i.test(name)) continue;
            const amount = row[2].match(/<(?:BSSUBAMT|BSMAINAMT)>([\s\S]*?)<\/(?:BSSUBAMT|BSMAINAMT)>/i);
            if (!amount) continue;
            const value = Math.abs(normalizeAmount(amount[1]));
            if (value > 0) ledgers[name] = (ledgers[name] || 0) + value;
        }
        return { ledgers, total: Object.values(ledgers).reduce((sum, value) => sum + value, 0) };
    }

    const loansInfo = parseLoans(loansXml);
    let provisionsInfo = parseProvisions(currLiabXml, data);
    const sundryCredInfo    = parseSundryCreditors(sundryCreditorsXml);
    const sundryDebtorsInfo  = parseSundryDebtors(sundryDebtorsXml);
    const investmentsInfo    = parseInvestments(currAssetsXml);
    const otherCurrLiabInfo  = parseOtherCurrentLiabilities(currLiabXml);
    const advancesInfo       = parseAdvancesAndDeposit(currAssetsXml);
    const cashInfo           = parseCashAndEquivalents(currAssetsXml);
    const gstItcAmount       = parseGstItc(currAssetsXml);
    const indirectIncomesInfo = parseIndirectIncomes(plXml);
    const note16Info         = parseNote16(indirectExpXml);
    const fixedAssetsInfo    = parseFixedAssets(bsXml);

    const mapping = {};

    let shareCapitalVal = 0;
    for (const key of Object.keys(data)) {
        const lowerKey = key.toLowerCase();
        if (lowerKey === 'share capital' || lowerKey === 'paid up capital' || lowerKey === 'capital account') {
            const val = Math.abs(data[key] || 0);
            shareCapitalVal += val;
            trace('Note 1', 'balance_sheet_response.xml', key, val, 'Matched Share Capital/Capital Account keywords');
        }
    }
    mapping['Note 1 - Share Capital'] = shareCapitalVal;

    // Note 2 – Reserves & Surplus (closing balance from BS XML)
    const plNetProfit = (data['Sales Accounts'] || 0)
        + (data['Direct Incomes'] || 0)
        + (data['Indirect Incomes'] || 0)
        + (data['Add: Purchase Accounts'] || 0)
        + (data['Direct Expenses'] || 0)
        + (data['Indirect Expenses'] || 0);
    
    let note2Val = 0;
    let note2Var = '';
    if (data['Reserves & Surplus']) { note2Val = data['Reserves & Surplus']; note2Var = 'Reserves & Surplus'; }
    else if (data['Profit & Loss A/c']) { note2Val = data['Profit & Loss A/c']; note2Var = 'Profit & Loss A/c'; }
    else if (data['Excess of income over expenditure']) { note2Val = data['Excess of income over expenditure']; note2Var = 'Excess of income over expenditure'; }
    else { note2Val = plNetProfit; note2Var = 'Calculated Net Profit (Sales + Incomes - Expenses)'; }
    
    mapping['Note 2 - Profit/Loss'] = note2Val;
    trace('Note 2', 'balance_sheet_response.xml', note2Var, note2Val, 'Primary Profit/Loss balance from Tally');
    
    // Account for Suspense A/c if it exists in BS
    if (data['Suspense A/c']) {
        const susVal = data['Suspense A/c'];
        mapping['Note 2 - Profit/Loss'] = (mapping['Note 2 - Profit/Loss'] || 0) + susVal;
        trace('Note 2', 'balance_sheet_response.xml', 'Suspense A/c', susVal, 'Aggregated Suspense A/c into reserves');
    }
    if (loansInfo) {
        mapping['Note 3 - Long-Term Borrowings'] = loansInfo.total;
        const loanBreakdown = { 'Secured Loans': loansInfo.secTotal, 'Unsecured Loans': loansInfo.unsecTotal, ...loansInfo.secured, ...loansInfo.unsecured };
        trace('Note 3', 'loans_liability_response.xml', 'Loans (Liability)', loansInfo.total, 'Parsed from detailed Loans Group Summary', loanBreakdown);
    }
    mapping['Note 4 - '] = 0;
    
    // Note 5: Long-Term and Short-Term Provisions
    const provisionLedgers = ['Provision for Accounting Charges', 'Provision for Audit Fees', 'Provision for Secretarial Charges', 'Provision for Taxation'];
    let catchAllProv = 0;
    for (let pName of provisionLedgers) {
        if (data[pName]) {
            catchAllProv += Math.abs(data[pName]);
            trace('Note 5', 'balance_sheet_response.xml', pName, data[pName], 'Captured via ledger name fallback');
        }
    }

    if (catchAllProv > 0 && (!provisionsInfo || provisionsInfo.total === 0)) {
        provisionsInfo = { shortTermItems: { 'Other Provisions': catchAllProv }, total: catchAllProv, provisionForTaxation: 0 };
    }
    mapping['Note 5 - Long-Term Provisions and Short-Term Provisions'] = provisionsInfo ? provisionsInfo.total : 0;
    if (provisionsInfo && provisionsInfo.total > 0 && !provisionsInfo.shortTermItems['Other Provisions']) {
         const provBreakdown = { ...provisionsInfo.shortTermItems };
         if (provisionsInfo.provisionForTaxation > 0) provBreakdown['Provision for Taxation'] = provisionsInfo.provisionForTaxation;
         trace('Note 5', 'current_liabilities_response.xml', 'Provisions (Group)', provisionsInfo.total, 'Parsed from detailed Current Liabilities Group Summary', provBreakdown);
    }

    const bsSundryCreditors = Math.abs(data['Sundry Creditors'] || 0);
    mapping['Note 6 - Trade Payables'] = (sundryCredInfo && Math.abs(sundryCredInfo.total - bsSundryCreditors) < 10) ? sundryCredInfo.total : bsSundryCreditors;
    if (sundryCredInfo && mapping['Note 6 - Trade Payables'] === sundryCredInfo.total) trace('Note 6', 'sundry_creditors_response.xml', 'Sundry Creditors (Group)', sundryCredInfo.total, 'Extracted from detailed group summary', sundryCredInfo.ledgers);
    else trace('Note 6', 'balance_sheet_response.xml', 'Sundry Creditors', mapping['Note 6 - Trade Payables'], 'Using authoritative Balance Sheet total');

    let note7Val = otherCurrLiabInfo ? otherCurrLiabInfo.total : (Math.abs(data['Other Current Liabilities'] || 0) + Math.abs(data['TDS Payable - Other Services'] || 0) + Math.abs(data['Duties & Taxes'] || 0));
    if (otherCurrLiabInfo) trace('Note 7', 'current_liabilities_response.xml', 'Other Current Liabilities (Detailed)', otherCurrLiabInfo.total, 'Sum of all non-skipped ledgers in curr liab', otherCurrLiabInfo.items);
    
    const totalCurrLiab = Math.abs(data['Current Liabilities'] || 0);
    const n5Val = mapping['Note 5 - Long-Term Provisions and Short-Term Provisions'] || 0;
    const n6Val = mapping['Note 6 - Trade Payables'] || 0;

    if (totalCurrLiab > 0 && Math.abs((note7Val + n6Val + n5Val) - totalCurrLiab) > 1) {
        const oldVal = note7Val;
        note7Val = totalCurrLiab - n6Val - n5Val;
        trace('Note 7', 'Logic Fallback', 'Current Liabilities - Note 6 - Note 5', note7Val, `Reconciliation: Adjusted from ${oldVal} to match Tally Group total`);
    }
    mapping['Note 7 - Other Current Liabilities'] = note7Val;

    mapping['Note 8 - PPE (Fixed Assets)'] = Math.abs(data['Fixed Assets'] || 0);
    trace('Note 8', 'balance_sheet_response.xml', 'Fixed Assets', mapping['Note 8 - PPE (Fixed Assets)'], 'Top-level group total');

    mapping['Note 9 - Long Term Loans & Advances'] = advancesInfo ? advancesInfo.total : Math.abs(data['Advances & Deposit'] || 0);
    if (advancesInfo) trace('Note 9', 'current_assets_response.xml', 'Advances & Deposit (Detailed)', advancesInfo.total, 'Parsed from detailed Current Assets Group Summary', advancesInfo.items);

    mapping['Note 10 - Investments'] = investmentsInfo ? investmentsInfo.total : Math.abs(data['Deposits (Assets)'] || 0);
    if (investmentsInfo) trace('Note 10', 'current_assets_response.xml', 'Deposits (Assets/Asset)', investmentsInfo.total, 'Parsed from detailed Current Assets Group Summary', investmentsInfo.ledgers);

    mapping['Note 11 - Sundry Debtors'] = sundryDebtorsInfo ? sundryDebtorsInfo.total : Math.abs(data['Sundry Debtors'] || 0);
    if (sundryDebtorsInfo) trace('Note 11', 'sundry_debtors_response.xml', 'Sundry Debtors (Detailed)', sundryDebtorsInfo.total, 'Parsed from detailed Group Summary', sundryDebtorsInfo.ledgers);

    mapping['Note 12 - Cash & Cash Equivalents'] = cashInfo ? cashInfo.total : (Math.abs(data['Cash-in-Hand'] || 0) + Math.abs(data['Bank Accounts'] || 0));
    if (cashInfo) trace('Note 12', 'current_assets_response.xml', 'Bank Accounts + Cash-in-hand', cashInfo.total, 'Parsed from detailed Group Summary', { 'Bank Accounts': cashInfo.bankAccounts, 'Cash-in-Hand': cashInfo.cashInHand });

    mapping['Note 17 - Depreciation'] = Math.abs(data['Depreciation'] || 0);
    trace('Note 17', 'profit_loss_response.xml', 'Depreciation', mapping['Note 17 - Depreciation'], 'Extracted from P&L');

    // ---------------------------------------------------------------
    // FINAL RECONCILIATION anchored to Tally's authoritative summary totals.
    // Note 13 absorbs any gap between Tally's total assets and notes 8-12.
    // ---------------------------------------------------------------
    const summaryTotals = parseSummaryTotals(bsSummaryXml);
    const tallyLiabTotal  = summaryTotals.liabilitiesTotal;
    const tallyAssetsTotal = summaryTotals.assetsTotal;

    // LIABILITIES SIDE: Override Note 7 as the balancing note so that
    // Notes 1-7 sum exactly to tallyLiabTotal (Tally's authoritative total).
    if (tallyLiabTotal > 0) {
        const knownLiabs = (mapping['Note 1 - Share Capital'] || 0)
            + (mapping['Note 2 - Profit/Loss'] || 0)
            + (loansInfo ? loansInfo.total : 0)
            + (mapping['Note 4 - '] || 0)
            + (provisionsInfo ? provisionsInfo.total : 0)
            + (mapping['Note 6 - Trade Payables'] || 0);
        mapping['Note 7 - Other Current Liabilities'] = Math.max(0, tallyLiabTotal - knownLiabs);
    }

    const computedLiabTotal = 
          (mapping['Note 1 - Share Capital'] || 0) +
          (mapping['Note 2 - Profit/Loss'] || 0) +
          (loansInfo ? loansInfo.total : 0) +
          (mapping['Note 4 - '] || 0) +
          (provisionsInfo ? provisionsInfo.total : 0) +
          (mapping['Note 6 - Trade Payables'] || 0) +
          (mapping['Note 7 - Other Current Liabilities'] || 0);
    const totalLiabilitiesSide = (tallyLiabTotal > 0) ? tallyLiabTotal : computedLiabTotal;

    // ASSETS SIDE: Note 13 absorbs the gap between Tally's total assets and notes 8-12.
    const otherAssetsSum = 
          (mapping['Note 8 - PPE (Fixed Assets)'] || 0) +
          (advancesInfo ? advancesInfo.total : 0) +
          (investmentsInfo ? investmentsInfo.total : 0) +
          (sundryDebtorsInfo ? sundryDebtorsInfo.total : 0) +
          (cashInfo ? cashInfo.total : 0);
    const anchorTotal = (tallyAssetsTotal > 0) ? tallyAssetsTotal : totalLiabilitiesSide;

    mapping['Note 13 - Other Current Assets'] = Math.max(0, anchorTotal - otherAssetsSum);
    trace('Note 13', 'Balancing Calculation', 'Tally Assets Total - (Notes 8-12)', mapping['Note 13 - Other Current Assets'], 'Anchored to balance_sheet_summary.xml authoritative total');

    mapping['Note 14 - Revenue from Operations'] = Math.abs(data['Sales Accounts'] || 0) + Math.abs(data['Direct Incomes'] || 0);
    trace('Note 14', 'profit_loss_response.xml', 'Sales Accounts + Direct Incomes', mapping['Note 14 - Revenue from Operations'], 'Core operations revenue');

    mapping['Note 15 - Other Income'] = indirectIncomesInfo ? indirectIncomesInfo.total : Math.abs(data['Indirect Incomes'] || 0);
    if (indirectIncomesInfo) trace('Note 15', 'profit_loss_response.xml', 'Indirect Incomes (Detailed)', indirectIncomesInfo.total, 'Parsed from detailed P&L Indirect Incomes', indirectIncomesInfo.items);

    mapping['Note 16 - Employee Benefit Expenses'] = note16Info ? note16Info.total : (Math.abs(data['Employment Benefit Expenses'] || 0) + Math.abs(data['Staff Welfare Expenses'] || 0));
    if (note16Info) trace('Note 16', 'indirect_expenses_response.xml', 'Staff Welfare + Salary (Detailed)', note16Info.total, 'Parsed from detailed Indirect Expenses Group', note16Info.items);

    const directExp = Math.abs(data['Direct Expenses'] || 0);
    const allIndirectExp = Math.abs(data['Indirect Expenses'] || 0);
    const note16 = mapping['Note 16 - Employee Benefit Expenses'];
    const note17 = mapping['Note 17 - Depreciation'];

    mapping['Note 18 - Other Expenses'] = Math.max(0, allIndirectExp - note16 - note17);
    trace('Note 18', 'Calculated', 'Total Indirect Exp - Note 16 - Note 17', mapping['Note 18 - Other Expenses'], 'Residual expenses');

    function fmt(num) {
        return num.toLocaleString('en-IN', { maximumFractionDigits: 2, minimumFractionDigits: 2 });
    }

    log("------- SCHEDULE 3 MAPPING (NOTE 1 TO 18) -------");
    for (let i = 1; i <= 18; i++) {
        const key = Object.keys(mapping).find(k => k.startsWith(`Note ${i} -`));
        if (i === 3 && loansInfo) {
            log("-----------------------------------------------------------------");
            log(`Note 3 - Long-term Borrowings                      : ₹ ${fmt(loansInfo.total)}`);

            log(`  (A) Secured Loans                                : ₹ ${fmt(loansInfo.secTotal)}`);
            for (let [name, val] of Object.entries(loansInfo.secured)) {
                log(`      - ${name.padEnd(39)}: ₹ ${fmt(val)}`);
            }

            log(`  (B) Unsecured Loans                              : ₹ ${fmt(loansInfo.unsecTotal)}`);
            for (let [name, val] of Object.entries(loansInfo.unsecured)) {
                log(`      - ${name.padEnd(39)}: ₹ ${fmt(val)}`);
            }

            log(`  Total Note 3                                     : ₹ ${fmt(loansInfo.total)}`);
            log("-----------------------------------------------------------------");
        } else if (i === 5 && provisionsInfo) {
            log("-----------------------------------------------------------------");
            log(`Note 5 - Long-Term Provisions and Short-Term Provisions`);

            let alphaCode = 65; // 'A'
            for (let [name, val] of Object.entries(provisionsInfo.shortTermItems)) {
                log(`  (${String.fromCharCode(alphaCode++)}) ${name.padEnd(42)}: ₹ ${fmt(val)}`);
            }
            if (provisionsInfo.provisionForTaxation > 0) {
                log(`  (${String.fromCharCode(alphaCode++)}) Provision for Taxation                    : ₹ ${fmt(provisionsInfo.provisionForTaxation)}`);
            }

            log(`  Total Note 5                                     : ₹ ${fmt(provisionsInfo.total)}`);
            log("-----------------------------------------------------------------");
        } else if (i === 7 && otherCurrLiabInfo) {
            log("-----------------------------------------------------------------");
            log(`Note 7 - Other Current Liabilities`);
            log();

            let alphaCode = 65; // 'A'
            for (let [name, val] of Object.entries(otherCurrLiabInfo.items)) {
                log(`  (${String.fromCharCode(alphaCode++)}) ${name.padEnd(42)}: ₹ ${fmt(val)}`);
            }

            log(`  ${'Total'.padEnd(45)}: ₹ ${fmt(otherCurrLiabInfo.total)}`);
            log("-----------------------------------------------------------------");
        } else if (i === 9 && advancesInfo) {
            log("-----------------------------------------------------------------");
            log(`Note 9 - Long Term Loans & Advances`);
            log();
            log(`  Unsecured advances : Considered Good`);
            log();

            let alphaCode = 97; // 'a'
            for (let [name, val] of Object.entries(advancesInfo.items)) {
                const valStr = val === 0 ? '-' : `₹ ${fmt(val)}`;
                log(`  (${String.fromCharCode(alphaCode++)}) ${name.padEnd(42)}: ${valStr}`);
            }

            log();
            log(`  ${'Total Unsecured Advances'.padEnd(45)}: ₹ ${fmt(advancesInfo.total)}`);
            log(`  ${'Total Advances'.padEnd(45)}: ₹ ${fmt(advancesInfo.total)}`);
            console.log("-----------------------------------------------------------------");
        } else if (i === 12 && cashInfo) {
            log("-----------------------------------------------------------------");
            log(`Note 12 - Cash & Cash Equivalents`);
            log();

            const dash = '-';
            log(`  ${'Balances with Banks'.padEnd(44)}: ₹ ${fmt(cashInfo.bankAccounts)}`);
            log(`  ${'Cheques, drafts on hand'.padEnd(44)}: ${dash}`);
            log(`  ${'Cash on Hand'.padEnd(44)}: ₹ ${fmt(cashInfo.cashInHand)}`);
            log(`  ${'Others ( Specify nature )'.padEnd(44)}: ${dash}`);
            log();
            log(`  ${'Total'.padEnd(44)}: ₹ ${fmt(cashInfo.total)}`);
            log("-----------------------------------------------------------------");
        } else if (i === 15 && indirectIncomesInfo) {
            log("-----------------------------------------------------------------");
            log(`Note 15 - Other Income`);
            log();

            for (let [name, val] of Object.entries(indirectIncomesInfo.items)) {
                const valStr = val === 0 ? '-' : `₹ ${fmt(val)}`;
                log(`  ${name.padEnd(46)}: ${valStr}`);
            }

            log();
            log(`  ${'Total'.padEnd(46)}: ₹ ${fmt(indirectIncomesInfo.total)}`);
            log("-----------------------------------------------------------------");
        } else if (i === 16 && note16Info) {
            log("-----------------------------------------------------------------");
            log(`Note 16 - Employee Benefit Expenses`);
            log();
            log(`  ${'Particulars'.padEnd(46)}  ${'For the year ended 31 March, 2025'.padStart(20)}`);
            log(`  ${''.padEnd(46)}  ${'Rs.'.padStart(20)}`);
            log(`  ${'-'.repeat(68)}`);

            for (let [name, val] of Object.entries(note16Info.items)) {
                log(`  ${name.padEnd(46)}: ₹ ${fmt(val)}`);
            }

            log(`  ${'-'.repeat(68)}`);
            log(`  ${'Total'.padEnd(46)}: ₹ ${fmt(note16Info.total)}`);
            log("-----------------------------------------------------------------");
        } else if (key) {
            log(`${key.padEnd(50)} : ₹ ${fmt(mapping[key])}`);
        } else if (i !== 3 && i !== 5 && i !== 7 && i !== 9 && i !== 12 && i !== 15 && i !== 16) {
            log(`Note ${i} -                                        : ₹ 0.00`);
        }
    }
    const cacheDir = path.join(__dirname, 'cache');
    if (!fs.existsSync(cacheDir)) fs.mkdirSync(cacheDir, { recursive: true });

    const companyClean = process.env.COMPANY_NAME ? process.env.COMPANY_NAME.replace(/[^a-zA-Z0-9]/g, '_') : '';
    const companySuffix = sessionId ? (companyClean ? `_${companyClean}_${sessionId}` : `_${sessionId}`) : (process.env.COMPANY_NAME ? `_${process.env.COMPANY_NAME}` : '');
    const txtFileName = `schedule3_output${companySuffix}.txt`;
    fs.writeFileSync(
        path.join(cacheDir, txtFileName),
        output,
        'utf8'
    );
    console.log(`\nSaved to cache/${txtFileName}`);

    // Save Trace Log
    const traceFileName = `mapping_trace${companySuffix}.txt`;
    const traceContent = traceLog.map(t => {
        let text = `[${t.Note}] Source: ${t.Source}\n     Var: ${t.Tally_Variable}\n     Val: ${t.Value}\n     Log: ${t.Logic}\n`;
        if (t.Breakdown) {
            text += `     Breakdown:\n`;
            for (const [k, v] of Object.entries(t.Breakdown)) {
                text += `       - ${k}: ${v}\n`;
            }
        }
        return text;
    }).join('\n' + '-'.repeat(80) + '\n');
    fs.writeFileSync(path.join(cacheDir, traceFileName), traceContent, 'utf8');
    fs.writeFileSync(path.join(cacheDir, `mapping_trace${companySuffix}.json`), JSON.stringify(traceLog, null, 2), 'utf8');
    console.log(`Saved trace log to cache/${traceFileName}`);

    const finalData = {
        period: { fromDate: process.env.FROM_DATE || '', toDate: process.env.TO_DATE || '' },
        company: process.env.COMPANY_NAME || '',
        sourceFiles: ['balance_sheet_response.xml','balance_sheet_summary.xml','profit_loss_response.xml','loans_liability_response.xml','current_liabilities_response.xml','current_assets_response.xml','indirect_expenses_response.xml','sundry_debtors_response.xml','sundry_creditors_response.xml'],
        mapping, loansInfo, provisionsInfo, sundryCredInfo, sundryDebtorsInfo, investmentsInfo, otherCurrLiabInfo, fixedAssetsInfo,
        advancesInfo, cashInfo, gstItcAmount, indirectIncomesInfo, note16Info,
        tallyLiabTotal, tallyAssetsTotal, summaryGroups: summaryTotals.groups
    };
    const jsonFileName = `schedule3_output${companySuffix}.json`;
    fs.writeFileSync(
        path.join(cacheDir, jsonFileName),
        JSON.stringify(finalData, null, 2),
        'utf8'
    );

    // ---------------------------------------------------------------
    // Build Excel output
    // ---------------------------------------------------------------
    await buildExcel({
        mapping, loansInfo, provisionsInfo, sundryCredInfo, sundryDebtorsInfo, investmentsInfo, otherCurrLiabInfo,
        advancesInfo, cashInfo, gstItcAmount, indirectIncomesInfo, note16Info, fixedAssetsInfo, fmt
    });
}

// ---------------------------------------------------------------
// Excel Builder
// ---------------------------------------------------------------
async function buildExcel({
    mapping, loansInfo, provisionsInfo, sundryCredInfo, sundryDebtorsInfo, investmentsInfo, otherCurrLiabInfo,
    advancesInfo, cashInfo, indirectIncomesInfo, note16Info, fixedAssetsInfo, fmt
}) {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Tally Schedule 3 Extractor';
    const ws = workbook.addWorksheet('Schedule 3', {
        pageSetup: { fitToPage: true, fitToWidth: 1 }
    });

    // Column widths
    ws.columns = [
        { key: 'particulars', width: 52 },
        { key: 'current', width: 22 },
    ];

    // ---- Helper styles ----
    const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F3864' } };
    const SECTION_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD6E4F0' } };
    const SUB_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F7FB' } };
    const TOTAL_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCE6F1' } };
    const NUM_FMT = '[>=10000000]##\\,##\\,##\\,##0.00;[>=100000]##\\,##\\,##0.00;##,##0.00';
    const BORDER_THIN = { style: 'thin', color: { argb: 'FFB0C4D8' } };
    const BORDER_MED = { style: 'medium', color: { argb: 'FF1F3864' } };

    function allBorder(row, style = BORDER_THIN) {
        row.eachCell({ includeEmpty: true }, cell => {
            cell.border = { top: style, left: style, bottom: style, right: style };
        });
    }

    // ---- Global header ----
    const titleRow = ws.addRow(['Schedule 3 – Notes to Financial Statements', 'Current Year (₹)']);
    titleRow.height = 28;
    ws.mergeCells(`A${titleRow.number}:A${titleRow.number}`);
    titleRow.eachCell({ includeEmpty: true }, (cell, col) => {
        cell.fill = HEADER_FILL;
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: col === 1 ? 13 : 11, name: 'Calibri' };
        cell.alignment = { vertical: 'middle', horizontal: col === 1 ? 'left' : 'center', wrapText: true };
    });
    allBorder(titleRow, BORDER_MED);

    // ---- Helper: add a note section title row ----
    function addNoteHeader(label) {
        const r = ws.addRow([label, '']);
        r.height = 20;
        r.eachCell({ includeEmpty: true }, cell => {
            cell.fill = SECTION_FILL;
            cell.font = { bold: true, size: 11, name: 'Calibri', color: { argb: 'FF1F3864' } };
            cell.alignment = { vertical: 'middle', wrapText: true };
        });
        allBorder(r);
        return r;
    }

    // ---- Helper: add a data row ----
    function addDataRow(label, value, options = {}) {
        const r = ws.addRow([label, value === null ? '' : value]);
        r.height = options.height || 18;
        const [labelCell, valCell] = [r.getCell(1), r.getCell(2)];

        labelCell.font = { size: 10, name: 'Calibri', italic: options.italic || false, bold: options.bold || false };
        labelCell.alignment = { indent: options.indent || 0, wrapText: true };

        if (typeof value === 'number') {
            valCell.numFmt = NUM_FMT;
            valCell.alignment = { horizontal: 'right' };
            valCell.font = { size: 10, name: 'Calibri', bold: options.bold || false };
        } else {
            valCell.value = value === null ? '' : value;
            valCell.alignment = { horizontal: value === '-' ? 'center' : 'right' };
            valCell.font = { size: 10, name: 'Calibri', bold: options.bold || false };
        }

        if (options.fill) {
            r.eachCell({ includeEmpty: true }, cell => { cell.fill = options.fill; });
        }
        if (options.topBorder) {
            r.eachCell({ includeEmpty: true }, cell => { cell.border = { ...cell.border, top: BORDER_MED }; });
        }
        allBorder(r);
        return r;
    }

    // ---- Helper: add a total row ----
    function addTotalRow(label, value) {
        return addDataRow(label, value, { bold: true, fill: TOTAL_FILL, topBorder: true });
    }

    // ---- Helper: blank separator ----
    function addBlank() {
        const r = ws.addRow(['', '']);
        r.height = 6;
    }

    // ================================================================
    // NOTE 1 – Share Capital
    addNoteHeader('Note 1 – Share Capital');
    addTotalRow('Total Share Capital', mapping['Note 1 - Share Capital'] || 0);
    addBlank();

    // NOTE 2 – Profit / Loss
    addNoteHeader('Note 2 – Profit / Loss');
    addTotalRow('Net Profit / (Loss)', mapping['Note 2 - Profit/Loss'] || 0);
    addBlank();

    // NOTE 3 – Long-term Borrowings
    addNoteHeader('Note 3 – Long-Term Borrowings');
    if (loansInfo) {
        addDataRow('(A) Secured Loans', loansInfo.secTotal, { bold: true, fill: SUB_FILL });
        for (const [name, val] of Object.entries(loansInfo.secured)) {
            addDataRow(name, val, { indent: 2 });
        }
        addDataRow('(B) Unsecured Loans', loansInfo.unsecTotal, { bold: true, fill: SUB_FILL });
        for (const [name, val] of Object.entries(loansInfo.unsecured)) {
            addDataRow(name, val, { indent: 2 });
        }
        addTotalRow('Total Note 3', loansInfo.total);
    } else {
        addTotalRow('Total Note 3', 0);
    }
    addBlank();

    // NOTE 4
    addNoteHeader('Note 4 – Deferred Tax Liabilities (Net)');
    addTotalRow('Total Note 4', mapping['Note 4 - '] || 0);
    addBlank();

    // NOTE 5 – Provisions
    addNoteHeader('Note 5 – Short-Term Provisions');
    if (provisionsInfo) {
        let code = 65;
        for (const [name, val] of Object.entries(provisionsInfo.shortTermItems)) {
            addDataRow(`(${String.fromCharCode(code++)}) ${name}`, val, { indent: 1 });
        }
        if (provisionsInfo.provisionForTaxation > 0) {
            addDataRow(`(${String.fromCharCode(code++)}) Provision for Taxation`, provisionsInfo.provisionForTaxation, { indent: 1 });
        }
        addTotalRow('Total Note 5', provisionsInfo.total);
    } else {
        addTotalRow('Total Note 5', 0);
    }
    addBlank();

    // NOTE 6 – Trade Payables (Sundry Creditors breakdown)
    addNoteHeader('Note 6 – Trade Payables');
    if (sundryCredInfo && Object.keys(sundryCredInfo.ledgers).length > 0) {
        for (const [name, val] of Object.entries(sundryCredInfo.ledgers)) {
            addDataRow(name, val, { indent: 1 });
        }
        addTotalRow('Total Trade Payables', sundryCredInfo.total);
    } else {
        addTotalRow('Total Trade Payables', mapping['Note 6 - Trade Payables'] || 0);
    }
    addBlank();

    // NOTE 7 – Other Current Liabilities
    addNoteHeader('Note 7 – Other Current Liabilities');
    if (otherCurrLiabInfo) {
        let code = 65;
        for (const [name, val] of Object.entries(otherCurrLiabInfo.items)) {
            addDataRow(`(${String.fromCharCode(code++)}) ${name}`, val, { indent: 1 });
        }
        addTotalRow('Total Note 7', otherCurrLiabInfo.total);
    } else {
        addTotalRow('Total Note 7', mapping['Note 7 - Other Current Liabilities'] || 0);
    }
    addBlank();

    // NOTE 8 – PPE
    addNoteHeader('Note 8 – Property, Plant & Equipment (Fixed Assets)');
    if (fixedAssetsInfo && Object.keys(fixedAssetsInfo.ledgers).length > 0) {
        for (const [name, value] of Object.entries(fixedAssetsInfo.ledgers)) {
            addDataRow(name, value, { indent: 1 });
        }
    }
    addTotalRow('Total Fixed Assets', mapping['Note 8 - PPE (Fixed Assets)'] || 0);
    addBlank();

    // NOTE 9 – Advances & Deposits
    addNoteHeader('Note 9 – Long Term Loans & Advances');
    if (advancesInfo) {
        addDataRow('Unsecured advances : Considered Good', null, { italic: true });
        let code = 97;
        for (const [name, val] of Object.entries(advancesInfo.items)) {
            addDataRow(`(${String.fromCharCode(code++)}) ${name}`, val === 0 ? '-' : val, { indent: 1 });
        }
        addTotalRow('Total Advances', advancesInfo.total);
    } else {
        addTotalRow('Total Advances', mapping['Note 9 - Long Term Loans & Advances'] || 0);
    }
    addBlank();

    // NOTE 10 – Investments
    addNoteHeader('Note 10 – Investments');
    if (investmentsInfo && Object.keys(investmentsInfo.ledgers).length > 0) {
        for (const [name, val] of Object.entries(investmentsInfo.ledgers)) {
            addDataRow(name, Math.abs(val), { indent: 1 });
        }
        addTotalRow('Total Investments', investmentsInfo.total);
    } else {
        addTotalRow('Total Investments', mapping['Note 10 - Investments'] || 0);
    }
    addBlank();

    // NOTE 11 – Sundry Debtors
    addNoteHeader('Note 11 – Sundry Debtors / Trade Receivables');
    if (sundryDebtorsInfo && Object.keys(sundryDebtorsInfo.ledgers).length > 0) {
        for (const [name, val] of Object.entries(sundryDebtorsInfo.ledgers)) {
            addDataRow(name, Math.abs(val), { indent: 1 });
        }
        addTotalRow('Total Sundry Debtors', sundryDebtorsInfo.total);
    } else {
        addTotalRow('Total Sundry Debtors', mapping['Note 11 - Sundry Debtors'] || 0);
    }
    addBlank();

    // NOTE 12 – Cash & Equivalents
    addNoteHeader('Note 12 – Cash & Cash Equivalents');
    if (cashInfo) {
        addDataRow('Balances with Banks', cashInfo.bankAccounts, { indent: 1 });
        addDataRow('Cheques, drafts on hand', '-', { indent: 1 });
        addDataRow('Cash on Hand', cashInfo.cashInHand, { indent: 1 });
        addDataRow('Others (Specify nature)', '-', { indent: 1 });
        addTotalRow('Total Note 12', cashInfo.total);
    } else {
        addTotalRow('Total Note 12', mapping['Note 12 - Cash & Cash Equivalents'] || 0);
    }
    addBlank();

    // NOTE 13 – Other Current Assets
    addNoteHeader('Note 13 – Other Current Assets');
    addTotalRow('Total Other Current Assets', mapping['Note 13 - Other Current Assets'] || 0);
    addBlank();

    // NOTE 14 – Revenue from Operations
    addNoteHeader('Note 14 – Revenue from Operations');
    addTotalRow('Total Revenue from Operations', mapping['Note 14 - Revenue from Operations'] || 0);
    addBlank();

    // NOTE 15 – Other Income
    addNoteHeader('Note 15 – Other Income');
    if (indirectIncomesInfo) {
        for (const [name, val] of Object.entries(indirectIncomesInfo.items)) {
            addDataRow(name, val === 0 ? '-' : val, { indent: 1 });
        }
        addTotalRow('Total Other Income', indirectIncomesInfo.total);
    } else {
        addTotalRow('Total Other Income', mapping['Note 15 - Other Income'] || 0);
    }
    addBlank();

    // NOTE 16 – Employee Benefit Expenses
    addNoteHeader('Note 16 – Employee Benefit Expenses');
    if (note16Info) {
        for (const [name, val] of Object.entries(note16Info.items)) {
            addDataRow(name, val, { indent: 1 });
        }
        addTotalRow('Total Note 16', note16Info.total);
    } else {
        addTotalRow('Total Note 16', mapping['Note 16 - Employee Benefit Expenses'] || 0);
    }
    addBlank();

    // NOTE 17 – Depreciation
    addNoteHeader('Note 17 – Depreciation & Amortisation');
    addTotalRow('Total Depreciation', mapping['Note 17 - Depreciation'] || 0);
    addBlank();

    // NOTE 18 – Other Expenses
    addNoteHeader('Note 18 – Other Expenses');
    addTotalRow('Total Other Expenses', mapping['Note 18 - Other Expenses'] || 0);
    addBlank();

    // ---- Save ----
    // ---- Save ----
    const companySuffix = process.env.COMPANY_NAME ? `_${process.env.COMPANY_NAME}` : '';
    const outFileName = `schedule3_output${companySuffix}.xlsx`;
    const outPath = path.join(__dirname, 'cache', outFileName);
    try {
        await workbook.xlsx.writeFile(outPath);
        console.log(`Saved to ${outFileName}`);
    } catch (err) {
        if (err.code === 'EBUSY') {
            console.error(`\n[WARNING] The file ${outFileName} is open in Excel or another program. Could not update the Excel file, but text mapping was completed.`);
        } else {
            console.error('\n[ERROR] Failed to save Excel file:', err);
        }
    }
}

// ---- Entry point ----
(async () => {
    await extractMapping();
})();
