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

    // Preserve the Fixed Assets ledger hierarchy for Excel export.
    // The portal has this information in the detailed Balance Sheet XML,
    // but the old final JSON discarded it, which made PPE breakup impossible.
    function parseFixedAssets(xmlStr) {
        if (!xmlStr) return { ledgers: {}, total: 0 };

        // Tally can emit whitespace/newline/attribute variations. Work on the
        // Fixed Assets block instead of relying on an exact newline pattern.
        const markerRe = /<DSPDISPNAME>\s*Fixed Assets\s*<\/DSPDISPNAME>/i;
        const marker = markerRe.exec(xmlStr);
        if (!marker) return { ledgers: {}, total: 0 };

        const start = marker.index;
        const tail = xmlStr.slice(start + marker[0].length);

        // Stop at the next top-level balance-sheet group. The individual PPE
        // ledger blocks are between the Fixed Assets group and that group.
        const nextGroupRe = /<DSPDISPNAME>\s*(?:Current Assets|Current Liabilities|Loans \(Liability\)|Capital Account|Reserves & Surplus|Reserves and Surplus|Suspense A\/c)\s*<\/DSPDISPNAME>/i;
        const next = nextGroupRe.exec(tail);
        const section = next ? tail.slice(0, next.index) : tail;

        const ledgers = {};

        // Extract each ledger name and the first balance amount belonging to it.
        // Do not restrict the amount regex to [-0-9.]+ because Tally may return
        // commas, Dr/Cr suffixes, parentheses, or spaces.
        const blockRe = /<DSPDISPNAME>([\s\S]*?)<\/DSPDISPNAME>([\s\S]*?)(?=<DSPDISPNAME>|$)/gi;
        let match;
        while ((match = blockRe.exec(section))) {
            const name = match[1]
                .replace(/<!\[CDATA\[|\]\]>/g, '')
                .replace(/&#13;|&#10;/g, ' ')
                .replace(/&amp;/g, '&')
                .trim();

            if (!name || /^Fixed Assets$/i.test(name)) continue;

            const body = match[2];
            const amountMatch = body.match(/<(?:BSSUBAMT|BSMAINAMT)>([\s\S]*?)<\/(?:BSSUBAMT|BSMAINAMT)>/i);
            if (!amountMatch) continue;

            const value = Math.abs(normalizeAmount(amountMatch[1]));
            if (value > 0) {
                ledgers[name] = (ledgers[name] || 0) + value;
            }
        }

        return {
            ledgers,
            total: Object.values(ledgers).reduce((a, b) => a + b, 0)
        };
    }


    // Explicit Note 18 parser. Never create Note 18 from the balance-sheet
    // residual; collect the actual Other Current Assets group from Tally.
    function parseOtherCurrentAssets(xmlStr) {
        if (!xmlStr) return { items: {}, total: 0 };
        const blocks = xmlStr.split('<DSPACCNAME>');
        const items = {};
        let collecting = false;
        for (let i = 1; i < blocks.length; i++) {
            const block = blocks[i];
            const nameMatch = block.match(/<DSPDISPNAME>([^<]+)<\/DSPDISPNAME>/);
            if (!nameMatch) continue;
            const name = nameMatch[1].trim().replace(/&amp;/g, '&');
            if (name === 'Other Current Assets') {
                collecting = true;
                continue;
            }
            if (collecting && CURRENT_ASSET_GROUPS.has(name)) break;
            if (!collecting) continue;
            const infoStart = block.indexOf('<DSPACCINFO>');
            if (infoStart === -1) continue;
            const info = block.slice(infoStart);
            const drMatch = info.match(/<DSPCLDRAMTA>([^<]+)<\/DSPCLDRAMTA>/);
            const crMatch = info.match(/<DSPCLCRAMTA>([^<]+)<\/DSPCLCRAMTA>/);
            const dr = drMatch ? Math.abs(normalizeAmount(drMatch[1])) : 0;
            const cr = crMatch ? Math.abs(normalizeAmount(crMatch[1])) : 0;
            const value = dr - cr;
            if (value !== 0) items[name] = (items[name] || 0) + value;
        }
        return { items, total: Object.values(items).reduce((a, b) => a + b, 0) };
    }

    function parseOtherExpenses(xmlStr) {
        if (!xmlStr) return { items: {}, total: 0 };
        const blocks = xmlStr.split('<DSPACCNAME>');
        const items = {};
        const excluded = name => isEmployeeBenefitLedger(name) || /depreciation|amortization/i.test(name) || /bank\s*charges?|interest|finance\s*cost/i.test(name);
        for (let i = 1; i < blocks.length; i++) {
            const block = blocks[i];
            const nameMatch = block.match(/<DSPDISPNAME>([^<]*)<\/DSPDISPNAME>/i);
            if (!nameMatch) continue;
            const name = nameMatch[1].trim().replace(/&amp;/g, '&');
            if (!name || excluded(name)) continue;
            const infoStart = block.indexOf('<DSPACCINFO>');
            if (infoStart === -1) continue;
            const info = block.slice(infoStart);
            const drMatch = info.match(/<DSPCLDRAMTA>([^<]+)<\/DSPCLDRAMTA>/i);
            const crMatch = info.match(/<DSPCLCRAMTA>([^<]+)<\/DSPCLCRAMTA>/i);
            const dr = drMatch ? Math.abs(normalizeAmount(drMatch[1])) : 0;
            const cr = crMatch ? Math.abs(normalizeAmount(crMatch[1])) : 0;
            const value = Math.abs(dr - cr);
            if (value !== 0) items[name] = (items[name] || 0) + value;
        }
        return { items, total: Object.values(items).reduce((a, b) => a + b, 0) };
    }

    const fixedAssetsInfo = parseFixedAssets(bsXml);

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
    const otherExpensesInfo   = parseOtherExpenses(indirectExpXml);
    const otherCurrentAssetsInfo = parseOtherCurrentAssets(currAssetsXml);

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
    mapping["Note 1 - Brief about the entity"] = 0;
    mapping["Note 2 - Significant Accounting Policies"] = 0;
    mapping["Note 3 - Owner's capital Account"] = shareCapitalVal;

    // Note 4 - Reserves & Surplus (closing balance from BS XML)
    const plNetProfit = (data['Sales Accounts'] || 0)
        + (data['Direct Incomes'] || 0)
        + (data['Indirect Incomes'] || 0)
        + (data['Add: Purchase Accounts'] || 0)
        + (data['Direct Expenses'] || 0)
        + (data['Indirect Expenses'] || 0);
    
    let note4Val = 0;
    let note4Var = '';
    if (Object.prototype.hasOwnProperty.call(data, 'Reserves & Surplus')) { note4Val = data['Reserves & Surplus']; note4Var = 'Reserves & Surplus'; }
    else if (Object.prototype.hasOwnProperty.call(data, 'Profit & Loss A/c')) { note4Val = data['Profit & Loss A/c']; note4Var = 'Profit & Loss A/c'; }
    else if (Object.prototype.hasOwnProperty.call(data, 'Excess of income over expenditure')) { note4Val = data['Excess of income over expenditure']; note4Var = 'Excess of income over expenditure'; }
    else { note4Val = plNetProfit; note4Var = 'Calculated Net Profit (Sales + Incomes - Expenses)'; }
    
    mapping["Note 4 - Reserves and Surplus"] = note4Val;
    if (data['Suspense A/c']) {
        mapping["Note 4 - Reserves and Surplus"] += data['Suspense A/c'];
    }

    mapping["Note 5 - Borrowings"] = loansInfo ? loansInfo.total : 0;
    mapping["Note 6 - Deffered tax Liabilities"] = 0;
    mapping["Note 7 - Other long term Liabilities"] = 0;
    
    // Note 8: Provisions
    const provisionLedgers = ['Provision for Accounting Charges', 'Provision for Audit Fees', 'Provision for Secretarial Charges', 'Provision for Taxation'];
    let catchAllProv = 0;
    for (let pName of provisionLedgers) {
        if (data[pName]) {
            catchAllProv += Math.abs(data[pName]);
        }
    }
    if (catchAllProv > 0 && (!provisionsInfo || provisionsInfo.total === 0)) {
        provisionsInfo = { shortTermItems: { 'Other Provisions': catchAllProv }, total: catchAllProv, provisionForTaxation: 0 };
    }
    mapping["Note 8 - Provisions"] = provisionsInfo ? provisionsInfo.total : 0;

    const bsSundryCreditors = Math.abs(data['Sundry Creditors'] || 0);
    mapping["Note 9 - Trade Payables"] = (sundryCredInfo && Math.abs(sundryCredInfo.total - bsSundryCreditors) < 10) ? sundryCredInfo.total : bsSundryCreditors;

    // Note 10 is sourced from the actual Other Current Liabilities group.
    // It is deliberately NOT overwritten with a balance-sheet residual.
    const note10Val = otherCurrLiabInfo
        ? otherCurrLiabInfo.total
        : (Math.abs(data['Other Current Liabilities'] || 0) + Math.abs(data['TDS Payable - Other Services'] || 0) + Math.abs(data['Duties & Taxes'] || 0));
    mapping["Note 10 - Other current Liabilities"] = note10Val;

    mapping["Note 11 - PPE"] = Math.abs(data['Fixed Assets'] || 0);
    mapping["Note 12 - Investments - Non current and current"] = investmentsInfo ? investmentsInfo.total : Math.abs(data['Deposits (Assets)'] || 0);
    mapping["Note 13 - Loans and Advances"] = advancesInfo ? advancesInfo.total : Math.abs(data['Advances & Deposit'] || 0);
    mapping["Note 14 - Non current assets"] = 0;
    mapping["Note 15 - Inventory"] = Math.abs(data['Closing Stock'] || 0);
    mapping["Note 16 - Trade Receivables"] = sundryDebtorsInfo ? sundryDebtorsInfo.total : Math.abs(data['Sundry Debtors'] || 0);
    mapping["Note 17 - Cash and Bank Balance"] = cashInfo ? cashInfo.total : (Math.abs(data['Cash-in-Hand'] || 0) + Math.abs(data['Bank Accounts'] || 0));

    // Authoritative Tally totals are used for reconciliation/diagnostics only.
    // They are never used to manufacture a Schedule III note.
    const summaryTotals = parseSummaryTotals(bsSummaryXml);
    const tallyLiabTotal = summaryTotals.liabilitiesTotal;
    const tallyAssetsTotal = summaryTotals.assetsTotal;
    const computedLiabTotal = [3,4,5,6,7,8,9,10].reduce((sum, n) => sum + Number(mapping[`Note ${n} - ${n === 3 ? "Owner's capital Account" : n === 4 ? 'Reserves and Surplus' : n === 5 ? 'Borrowings' : n === 6 ? 'Deffered tax Liabilities' : n === 7 ? 'Other long term Liabilities' : n === 8 ? 'Provisions' : n === 9 ? 'Trade Payables' : 'Other current Liabilities'}`] || 0), 0);
    const otherCurrentAssets = otherCurrentAssetsInfo ? otherCurrentAssetsInfo.total : 0;
    mapping["Note 18 - Other current assets"] = otherCurrentAssets;
    const computedAssetTotal = [11,12,13,14,15,16,17,18].reduce((sum, n) => sum + Number(mapping[`Note ${n} - ${n === 11 ? 'PPE' : n === 12 ? 'Investments - Non current and current' : n === 13 ? 'Loans and Advances' : n === 14 ? 'Non current assets' : n === 15 ? 'Inventory' : n === 16 ? 'Trade Receivables' : n === 17 ? 'Cash and Bank Balance' : 'Other current assets'}`] || 0), 0);
    const liabilityReconciliationGap = tallyLiabTotal - computedLiabTotal;
    const assetReconciliationGap = tallyAssetsTotal - computedAssetTotal;

    mapping["Note 19 - Revenue from Operations"] = Math.abs(data['Sales Accounts'] || 0) + Math.abs(data['Direct Incomes'] || 0);
    mapping["Note 20 - other income"] = indirectIncomesInfo ? indirectIncomesInfo.total : Math.abs(data['Indirect Incomes'] || 0);
    mapping["Note 21 - Cost of goods sold"] = Math.max(0, Math.abs(data['Purchase Accounts'] || 0) + Math.abs(data['Direct Expenses'] || 0) + Math.abs(data['Opening Stock'] || 0) - Math.abs(data['Closing Stock'] || 0));
    mapping["Note 22 - Employee benefits expense"] = note16Info ? note16Info.total : (Math.abs(data['Employment Benefit Expenses'] || 0) + Math.abs(data['Staff Welfare Expenses'] || 0));
    mapping["Note 23 - Finance cost"] = Math.abs(data['Bank Charges'] || 0);
    mapping["Note 24 - Depreciation and amortization expense"] = Math.abs(data['Depreciation'] || 0);

    mapping["Note 25 - Other Expenses"] = otherExpensesInfo ? otherExpensesInfo.total : 0;


    function fmt(num) {
        return num.toLocaleString('en-IN', { maximumFractionDigits: 2, minimumFractionDigits: 2 });
    }

    log("------- SCHEDULE 3 MAPPING (NOTE 1 TO 25) -------");
    for (let i = 1; i <= 25; i++) {
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
    console.log(`\nSaved to ${cacheDir}/${txtFileName}`);

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
    console.log(`Saved trace log to ${cacheDir}/${traceFileName}`);

    const finalData = {
        period: { fromDate: process.env.FROM_DATE || '', toDate: process.env.TO_DATE || '' },
        company: process.env.COMPANY_NAME || '',
        sourceFiles: ['balance_sheet_response.xml','balance_sheet_summary.xml','profit_loss_response.xml','loans_liability_response.xml','current_liabilities_response.xml','current_assets_response.xml','indirect_expenses_response.xml','sundry_debtors_response.xml','sundry_creditors_response.xml'],
        mapping, loansInfo, provisionsInfo, sundryCredInfo, sundryDebtorsInfo, investmentsInfo, otherCurrLiabInfo,
        advancesInfo, cashInfo, gstItcAmount, indirectIncomesInfo, note16Info, fixedAssetsInfo,
        // Keep the normalized raw report values. The Excel generator uses these
        // as a fallback when a legacy mapping key is absent/zero.
        rawData: data,
        otherExpensesInfo, otherCurrentAssetsInfo,
        reconciliation: {
            tallyLiabTotal, computedLiabTotal, liabilityReconciliationGap,
            tallyAssetsTotal, computedAssetTotal, assetReconciliationGap,
            balancedWithinTolerance: Math.abs(liabilityReconciliationGap) <= 1 && Math.abs(assetReconciliationGap) <= 1
        },
        tallyLiabTotal, tallyAssetsTotal, summaryGroups: summaryTotals.groups
    };
    const jsonFileName = `schedule3_output${companySuffix}.json`;
    fs.writeFileSync(
        path.join(cacheDir, jsonFileName),
        JSON.stringify(finalData, null, 2),
        'utf8'
    );

    // ---------------------------------------------------------------
    // NOTE: A legacy single-company Excel builder used to run here
    // (buildExcel(), writing schedule3_output_<company>.xlsx). It used an
    // old 1-18 note key scheme that no longer matches the `mapping` object
    // built above (25-note Schedule III scheme), so most of its cells wrote
    // 0 for keys like "Note 8 - PPE (Fixed Assets)" that don't exist in
    // `mapping`. It was never the file the Merge & Download button serves
    // (that comes from generateMergedExcelPartnership -> Tally-Mapped-*.xlsx),
    // so removing it does not change any user-facing output — it only stops
    // a confusing, partially-wrong file from being written to /cache on
    // every mapping run. The JSON written above (schedule3_output_*.json)
    // remains the single source of truth consumed by the webpage and by
    // the merge/Excel-generation step.
    // ---------------------------------------------------------------
}


// ---- Entry point ----
(async () => {
    await extractMapping();
})();
