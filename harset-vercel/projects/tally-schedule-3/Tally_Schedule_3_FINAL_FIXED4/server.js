require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { exec } = require('child_process');
const path = require('path');
const fs = require('fs');
const http = require('http');
const { discoverTally, postXml } = require('./fetch_tally');
const os = require('os');
const { generateMergedExcel } = require('./generate_merged_excel');
const { generateMergedExcelPartnership } = require('./generate_merged_excel_partnership');
const { generateMergedExcelRwaTrust } = require('./generate_merged_excel_rwa_trust');
const { validateGeneratedWorkbook } = require('./excel_validator');
const { canonicalCompanyName, canonicalCompanyType, expectedNoteCount, snapshot, validateMappingPayload } = require('./schedule3_model');

const app = express();
const port = parseInt(process.env.PORT || '8000', 10);
const host = process.env.HOST || '0.0.0.0';

function getLocalLanAddresses() {
    const ifaces = os.networkInterfaces();
    const addresses = [];
    Object.values(ifaces).forEach(entries => {
        entries.forEach(iface => {
            if (iface.family === 'IPv4' && !iface.internal) {
                addresses.push(iface.address);
            }
        });
    });
    return addresses;
}

// ── Tally config (shared) ──────────────────────────────────────────────────

// ── Fetch company list from Tally ──────────────────────────────────────────
async function fetchCompanyList(tallyHost, tallyPort) {
    const host = String(tallyHost || process.env.TALLY_HOST || 'localhost').trim();
    const requestedPort = Number(tallyPort || process.env.TALLY_PORT || 9000);
    try {
                const xml = `<ENVELOPE>
        <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>MyCompanyList</ID></HEADER>
        <BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>
            <TDL><TDLMESSAGE><COLLECTION NAME="MyCompanyList"><TYPE>Company</TYPE><FETCH>Name</FETCH></COLLECTION></TDLMESSAGE></TDL>
        </DESC></BODY>
</ENVELOPE>`;
        const port = await discoverTally(host, requestedPort);
        const data = await postXml(host, port, xml, undefined, true);
        const names = [];
        const re = /<NAME(?:\s+[^>]*)?>([\s\S]*?)<\/NAME>/gi;
        let m;
        while ((m = re.exec(data))) {
            const name = String(m[1]).replace(/<!\[CDATA\[|\]\]>/g, '').replace(/&amp;/g,'&').trim();
            if (name && !names.includes(name)) names.push(name);
        }
        console.log(`[INFO] Retrieved ${names.length} companies from Tally on ${host}:${port}`);
        if (names.length === 0) {
            console.warn(`[WARN] Tally responded successfully but returned no company names from ${host}:${port}`);
        }
        return names;
    } catch (error) {
        console.error(`[ERROR] Company list fetch failed for ${host}:${requestedPort}: ${error.message}`);
        throw error;
    }
}

app.use(cors());
app.use(express.json());
app.use(express.static('public', { etag: false, maxAge: 0, setHeaders: (res) => { res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate'); } }));

// ── Company list endpoints ─────────────────────────────────────────────────
app.post('/api/companies', async (req, res) => {
    try {
        const { tallyHost, tallyPort } = req.body || {};
        console.log(`[INFO] Company discovery requested for ${tallyHost || process.env.TALLY_HOST || 'localhost'}:${tallyPort || process.env.TALLY_PORT || 9000}`);
        const companies = await fetchCompanyList(tallyHost, tallyPort);
        res.json({ success: true, companies });
    } catch (error) {
        res.status(503).json({ success: false, companies: [], error: error.message });
    }
});

    app.get('/api/local-company-list', (req, res) => {
        const listPath = path.join(__dirname, 'company_list.json');
        if (fs.existsSync(listPath)) {
            const data = fs.readFileSync(listPath, 'utf8');
            res.json(JSON.parse(data));
        } else {
            res.json([]);
        }
    });

const importedLogsDir = path.resolve(__dirname, process.env.TALLYEDIT_LOG_DIR || path.join('cache', 'imported-logs'));
function safeLogFilename(filename) {
    return typeof filename === 'string' && /^[a-zA-Z0-9._-]+\.(json|log|txt)$/i.test(filename)
        ? filename
        : null;
}

app.get('/api/logs', (req, res) => {
    fs.mkdirSync(importedLogsDir, { recursive: true });
    const logs = fs.readdirSync(importedLogsDir)
        .filter(safeLogFilename)
        .map(filename => ({ filename, url: `/api/logs/${encodeURIComponent(filename)}` }));
    res.json({ success: true, logs });
});

app.get('/api/logs/:filename', (req, res) => {
    const filename = safeLogFilename(req.params.filename);
    if (!filename) return res.status(400).json({ success: false, error: 'Invalid log filename' });
    const filePath = path.join(importedLogsDir, filename);
    if (!fs.existsSync(filePath)) return res.status(404).json({ success: false, error: 'Log not found' });
    if (req.query.view === '1') return res.type(path.extname(filename)).sendFile(filePath);
    res.download(filePath, filename);
});

app.post('/api/logs/import', (req, res) => {
    const filename = safeLogFilename(req.query.filename);
    if (!filename) return res.status(400).json({ success: false, error: 'filename must end in .json, .log, or .txt' });
    fs.mkdirSync(importedLogsDir, { recursive: true });
    const filePath = path.join(importedLogsDir, filename);
    fs.writeFileSync(filePath, JSON.stringify(req.body, null, 2), 'utf8');
    res.status(201).json({ success: true, filename, url: `/api/logs/${encodeURIComponent(filename)}` });
});

    
app.get('/api/health', async (req, res) => {
    const tallyHost = req.query.host || process.env.TALLY_HOST || 'localhost';
    const tallyPort = req.query.port || process.env.TALLY_PORT || 9000;
    try {
        const discoveredPort = await discoverTally(tallyHost, tallyPort);
        res.json({ success: true, tallyHost, tallyPort: discoveredPort, message: `Tally reachable at ${tallyHost}:${discoveredPort}` });
    } catch (error) {
        res.status(503).json({ success: false, tallyHost, tallyPort, error: error.message });
    }
});

app.post('/api/run-mapping', async (req, res) => {
        const { companyName, fromDate, toDate, pvtLtdInfo, companyType, tallyHost, tallyPort } = req.body;

        if (!companyName || !fromDate || !toDate) {
            return res.status(400).json({ error: 'Missing required parameters' });
        }

        try {
            console.log(`Starting mapping for ${companyName} from ${fromDate} to ${toDate}`);

            const sessionId = Date.now().toString() + '_' + Math.random().toString(36).substr(2, 9);
            const envVars = {
                ...process.env,
                COMPANY_NAME: companyName,
                FROM_DATE: fromDate,
                TO_DATE: toDate,
                COMPANY_TYPE: companyType || '',
                SESSION_ID: sessionId,
                TALLY_HOST: tallyHost || process.env.TALLY_HOST || 'localhost',
                TALLY_PORT: tallyPort || process.env.TALLY_PORT || ''
            };

            // Run fetch_tally.js
            console.log('Running fetch_tally.js...');
            await runCommand('node fetch_tally.js', envVars);

            // Run mapping script based on company type
            if (companyType === 'Private Limited') {
                console.log('Running test_mapping.js...');
                await runCommand('node test_mapping.js', envVars);
            } else if (companyType === 'RWA' || companyType === 'Trust') {
                console.log('Running test_mapping_rwa.js...');
                await runCommand('node test_mapping_rwa.js', envVars);
            } else {
                console.log('Running test_mapping_partnership.js...');
                await runCommand('node test_mapping_partnership.js', envVars);
            }

            // Read output
            const companyClean = companyName ? companyName.replace(/[^a-zA-Z0-9]/g, '_') : '';
            const uniqueSuffix = companyClean ? `_${companyClean}_${sessionId}` : `_${sessionId}`;
            const jsonPath = path.join(__dirname, 'cache', `schedule3_output${uniqueSuffix}.json`);
            let result = null;
            if (fs.existsSync(jsonPath)) {
                result = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));
            } else {
                result = { error: 'No output file found. The process may have failed silently.' };
            }

            // Inject pvtLtdInfo so the frontend can render Note 1 breakdown
            if (pvtLtdInfo && result && !result.error) {
                result.pvtLtdInfo = pvtLtdInfo;
            }
            if (result && !result.error) {
                result.companyType = companyType;
                result.sessionId = sessionId;
            }

            const excelFileName = `schedule3_output${companyName ? `_${companyName}` : ''}.xlsx`;
            res.json({ success: true, result, excelFileName });
        } catch (error) {
            console.error('Mapping failed:', error);
            res.status(500).json({ success: false, error: error.message || 'Mapping failed' });
        }
    });

    function runCommand(command, env) {
        return new Promise((resolve, reject) => {
            exec(command, { env, cwd: __dirname }, (error, stdout, stderr) => {
                if (error) {
                    console.error(`Error executing ${command}:`, error);
                    const errMsg = stderr || error.message;
                    return reject(new Error(errMsg));
                }
                console.log(`Stdout from ${command}: ${stdout}`);
                if (stderr) console.error(`Stderr from ${command}: ${stderr}`);
                resolve(stdout);
            });
        });
    }

    app.post('/api/merge-mapping', async (req, res) => {
        /**
         * Merge contract:
         *  - the two browser payloads are the only source data used for Excel;
         *  - same company, same entity type and consecutive financial years are required;
         *  - every expected Schedule III note must exist and be numeric;
         *  - the generated workbook is read back and compared for BOTH years before download.
         */
        try {
            const body = req.body || {};
            const leftData = body.leftData;
            const rightData = body.rightData;
            const leftName = normalizeCompanyName(body.leftName, leftData);
            const rightName = normalizeCompanyName(body.rightName, rightData);
            const leftYear = Number(body.leftYear);
            const rightYear = Number(body.rightYear);
            const requestedType = canonicalCompanyType(body.companyType || '');
            const leftType = canonicalCompanyType(body.leftCompanyType || leftData?.companyType || requestedType);
            const rightType = canonicalCompanyType(body.rightCompanyType || rightData?.companyType || requestedType);

            if (!leftData || !rightData) {
                return res.status(400).json({ success: false, error: 'Both current-year and previous-year mappings are required. Generate both reports before merging.' });
            }
            if (!Number.isInteger(leftYear) || !Number.isInteger(rightYear)) {
                return res.status(400).json({ success: false, error: 'Invalid financial years supplied to Excel merge.' });
            }
            if (leftYear !== rightYear + 1) {
                return res.status(400).json({ success: false, error: `Financial years must be consecutive. Current year=${leftYear}, previous year=${rightYear}.` });
            }
            if (!leftType || !rightType || leftType !== rightType) {
                return res.status(400).json({ success: false, error: `Company type mismatch. Current year=${leftType || 'unknown'}, previous year=${rightType || 'unknown'}.` });
            }
            if (requestedType && requestedType !== leftType) {
                return res.status(400).json({ success: false, error: `Selected company type (${requestedType}) does not match the generated mapping (${leftType}). Regenerate the reports using the same company type.` });
            }
            if (canonicalCompanyName(leftName) !== canonicalCompanyName(rightName)) {
                return res.status(400).json({ success: false, error: `Company mismatch. Current year=${leftName}; previous year=${rightName}. Both years must belong to the same company.` });
            }

            const payloadErrors = [
                ...validateMappingPayload(leftData, 'Current-year', leftType),
                ...validateMappingPayload(rightData, 'Previous-year', rightType)
            ];
            if (payloadErrors.length) {
                return res.status(400).json({ success: false, error: `Invalid mapping payload:\n${payloadErrors.join('\n')}` });
            }

            const noteCount = expectedNoteCount(leftType);
            const currentSnapshot = snapshot(leftData, noteCount);
            const previousSnapshot = snapshot(rightData, noteCount);

            // This is the exact regression that previously produced a correct portal
            // and a zero-filled Excel file. Do not generate a workbook from this state.
            if (currentSnapshot.nonZero === 0 && previousSnapshot.nonZero > 0) {
                return res.status(409).json({
                    success: false,
                    error: 'Current-year mapping contains only zero values while the previous year contains data. Excel generation was stopped to prevent a zero-filled workbook. Regenerate the current-year mapping and retry.'
                });
            }

            console.log('[EXCEL] Browser payload accepted as authoritative source.');
            console.log(JSON.stringify({ current: currentSnapshot.notes, previous: previousSnapshot.notes }, null, 2));

            const safeCompanyName = leftName.replace(/[\\/?%*:|"<>]/g, '').trim() || 'Schedule3';
            const outFileName = `Tally-Mapped-${safeCompanyName}-${leftYear}-${rightYear}.xlsx`;
            const outPath = path.join(__dirname, 'cache', outFileName);
            fs.mkdirSync(path.join(__dirname, 'cache'), { recursive: true });

            if (leftType === 'RWA' || leftType === 'Trust') {
                await generateMergedExcelRwaTrust(leftName, rightName, leftData, rightData, outPath, leftYear, rightYear);
            } else if (leftType === 'Private Limited') {
                await generateMergedExcel(leftName, leftData, rightName, rightData, outPath, leftYear, rightYear, leftType, body.leftPvtInfo, body.rightPvtInfo);
            } else {
                await generateMergedExcelPartnership(leftName, leftData, rightName, rightData, outPath, leftYear, rightYear);
            }

            const validation = await validateGeneratedWorkbook(outPath, leftData, rightData, leftYear, rightYear, leftType);
            res.json({ success: true, excelFileName: outFileName, source: 'browser-report-snapshot', validation });
        } catch (error) {
            console.error('[EXCEL][ERROR] Merge failed:', error);
            res.status(500).json({ success: false, error: error.message || 'Excel generation failed' });
        }
    });

    function normalizeCompanyName(value, data) {
        if (typeof value === 'string' && value.trim()) return value.trim();
        if (typeof data?.company === 'string' && data.company.trim()) return data.company.trim();
        if (value && typeof value === 'object') {
            const candidate = value.name || value.label || value.value || value.companyName;
            if (typeof candidate === 'string' && candidate.trim()) return candidate.trim();
        }
        return '';
    }

    app.get('/api/download/:filename', (req, res) => {
        const filename = req.params.filename;
        if (filename.includes('..') || !filename.endsWith('.xlsx')) {
            return res.status(400).send('Invalid filename');
        }
        const filePath = path.join(__dirname, 'cache', filename);
        if (fs.existsSync(filePath)) {
            res.download(filePath);
        } else {
            res.status(404).send('File not found');
        }
    });

app.listen(port, host, () => {
    console.log(`Server running at http://localhost:${port}`);
    if (host === '0.0.0.0') {
        const localIPs = getLocalLanAddresses();
        if (localIPs.length > 0) {
            localIPs.forEach(ip => console.log(`Accessible on LAN at http://${ip}:${port}`));
        } else {
            console.log(`Accessible on LAN at http://<your-machine-ip>:${port}`);
        }
    } else {
        console.log(`Listening on host ${host}. Access at http://${host}:${port}`);
    }
});
