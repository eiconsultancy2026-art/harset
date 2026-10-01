'use strict';

require('dotenv').config();
const http = require('http');
const fs = require('fs');
const path = require('path');

const DISCOVERY_PORTS = [9000, 9001, 9002, 9003, 9004, 9005];
const DEFAULT_HOST = process.env.TALLY_HOST || 'localhost';
const REQUEST_TIMEOUT_MS = Number(process.env.TALLY_TIMEOUT_MS || 20000);
const MAX_RETRIES = 3;
const RETRY_BASE_DELAY_MS = Number(process.env.TALLY_RETRY_DELAY_MS || 500);
const activeServers = new Map(); // host -> port

const info = m => console.log(`[INFO] ${m}`);
const warn = m => console.warn(`[WARN] ${m}`);
const errorLog = m => console.error(`[ERROR] ${m}`);

function escapeXml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function normalizeHost(host) {
  const value = String(host || DEFAULT_HOST).trim();
  if (!value) return DEFAULT_HOST;
  return value.replace(/^https?:\/\//i, '').replace(/\/$/, '');
}

function normalizePort(port) {
  const n = Number(port);
  return Number.isInteger(n) && n > 0 && n <= 65535 ? n : null;
}

function generateRequestXML(reportName, groupName = null, isDetailed = true, options = {}) {
  const companyName = options.companyName ?? process.env.COMPANY_NAME ?? '';
  const fromDate = options.fromDate ?? process.env.FROM_DATE ?? '';
  const toDate = options.toDate ?? process.env.TO_DATE ?? '';
  const group = groupName ? `<GROUPNAME>${escapeXml(groupName)}</GROUPNAME>` : '';
  const company = companyName ? `<SVCURRENTCOMPANY>${escapeXml(companyName)}</SVCURRENTCOMPANY>` : '';
  const dates = fromDate && toDate
    ? `<SVFROMDATE>${escapeXml(fromDate)}</SVFROMDATE><SVTODATE>${escapeXml(toDate)}</SVTODATE>`
    : '';
  return `<ENVELOPE>
  <HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER>
  <BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>${escapeXml(reportName)}</REPORTNAME>
  <STATICVARIABLES>
    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
    ${company}${dates}${group}
    <EXPLODEFLAG>${isDetailed ? 'Yes' : 'No'}</EXPLODEFLAG>
    <ISDETAILED>${isDetailed ? 'Yes' : 'No'}</ISDETAILED>
    <DSPSHOWALLLEDGERS>Yes</DSPSHOWALLLEDGERS>
    <EXPLODEALLLEVELS>Yes</EXPLODEALLLEVELS>
    <DSPDISPLAYBILLS>Yes</DSPDISPLAYBILLS>
  </STATICVARIABLES></REQUESTDESC></EXPORTDATA></BODY>
</ENVELOPE>`;
}

function postXml(host, port, xmlData, timeoutMs = REQUEST_TIMEOUT_MS, simpleHeaders = false) {
  host = normalizeHost(host);
  port = normalizePort(port);
  if (!port) return Promise.reject(new Error('Invalid Tally port'));
  return new Promise((resolve, reject) => {
    const req = http.request({
      hostname: host, port, path: '/', method: 'POST',
      headers: simpleHeaders
        ? { 'Content-Type': 'text/xml', 'Content-Length': Buffer.byteLength(xmlData) }
        : {
          'Content-Type': 'text/xml; charset=utf-8',
          Accept: 'application/xml,text/xml,*/*',
          'Content-Length': Buffer.byteLength(xmlData),
          Connection: 'close'
        }
    }, res => {
      let body = '';
      res.setEncoding('utf8');
      res.on('data', c => { body += c; });
      res.on('end', () => {
        if (res.statusCode < 200 || res.statusCode >= 300)
          return reject(new Error(`Tally returned HTTP ${res.statusCode}`));
        if (!body.trim()) return reject(new Error('Tally returned an empty response'));
        if (!/<[A-Za-z_][^>]*>/i.test(body)) return reject(new Error('Tally response is not XML'));
        resolve(body);
      });
    });
    req.setTimeout(timeoutMs, () => req.destroy(new Error(`Tally request timed out after ${timeoutMs}ms`)));
    req.on('error', reject);
    try { req.write(xmlData); req.end(); } catch (e) { reject(e); }
  });
}

async function testTallyPort(host, port) {
  const probe = `<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>CompanyProbe</ID></HEADER><BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES><TDL><TDLMESSAGE><COLLECTION NAME="CompanyProbe"><TYPE>Company</TYPE><FETCH>Name</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>`;
  try { return /<ENVELOPE\b/i.test(await postXml(host, port, probe, 3000)); }
  catch (_) { return false; }
}

async function discoverTally(host = DEFAULT_HOST, preferredPort = null) {
  host = normalizeHost(host);
  const configured = normalizePort(preferredPort ?? process.env.TALLY_PORT);
  if (configured) {
    activeServers.set(host, configured);
    info(`Using configured Tally XML server at ${host}:${configured}`);
    return configured;
  }
  const cached = activeServers.get(host);
  const ports = [];
  if (configured) ports.push(configured);
  if (cached && !ports.includes(cached)) ports.push(cached);
  for (const p of DISCOVERY_PORTS) if (!ports.includes(p)) ports.push(p);

  info(`Searching for Tally XML server at ${host}:${ports.join(', ')}`);
  for (const port of ports) {
    try {
      if (await testTallyPort(host, port)) {
        activeServers.set(host, port);
        info(`Connected to Tally on ${host}:${port}`);
        return port;
      }
      info(`No Tally response on ${host}:${port}`);
    } catch (e) { warn(`Port ${port} check failed: ${e.message}`); }
  }
  throw new Error(`Tally XML Server Unreachable at ${host}. Checked ports ${ports.join(', ')}. Verify TallyPrime is running, HTTP/XML is enabled, the company is loaded, and Windows Firewall allows the selected port.`);
}

async function requestWithRetry(host, port, xmlData, requestName = 'Tally request') {
  let lastError;
  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    try {
      if (attempt > 0) {
        const delay = RETRY_BASE_DELAY_MS * (2 ** (attempt - 1));
        warn(`${requestName}: retry ${attempt}/${MAX_RETRIES} after ${delay}ms`);
        await new Promise(r => setTimeout(r, delay));
      }
      info(`${requestName}: requesting ${normalizeHost(host)}:${port}`);
      return await postXml(host, port, xmlData);
    } catch (e) {
      lastError = e;
      warn(`${requestName}: attempt ${attempt + 1} failed - ${e.message}`);
    }
  }
  throw new Error(`${requestName} failed after ${MAX_RETRIES + 1} attempts: ${lastError?.message || 'Unknown network error'}`);
}

function saveResponse(fileName, xml) {
  if (!/^[a-zA-Z0-9._-]+\.xml$/.test(fileName)) throw new Error(`Invalid XML output filename: ${fileName}`);
  const sessionId = process.env.SESSION_ID || '';
  const dir = path.join(__dirname, 'responses', sessionId || '');
  fs.mkdirSync(dir, { recursive: true });
  const target = path.join(dir, path.basename(fileName));
  fs.writeFileSync(target, xml, 'utf8');
  info(`Saved XML response: ${target}`);
  return target;
}

async function fetchFromTally(reportName, outputFileName, groupName = null, isDetailed = true) {
  const host = normalizeHost(process.env.TALLY_HOST || DEFAULT_HOST);
  try {
    const port = await discoverTally(host, process.env.TALLY_PORT);
    const xml = generateRequestXML(reportName, groupName, isDetailed);
    const response = await requestWithRetry(host, port, xml, groupName ? `${reportName} (${groupName})` : reportName);
    saveResponse(outputFileName, response);
    return response;
  } catch (e) {
    errorLog(`${reportName}: ${e.message}`);
    throw e;
  }
}

function generateCurrentAssetsXML() { return generateRequestXML('Group Summary', 'Current Assets', true); }
async function fetchCurrentAssets() {
  const host = normalizeHost(process.env.TALLY_HOST || DEFAULT_HOST);
  try {
    const port = await discoverTally(host, process.env.TALLY_PORT);
    const response = await requestWithRetry(host, port, generateCurrentAssetsXML(), 'Group Summary (Current Assets)');
    saveResponse('current_assets_response.xml', response);
    return response;
  } catch (e) { errorLog(`Current Assets: ${e.message}`); throw e; }
}

async function main() {
  info('Starting Tally Schedule III extraction...');
  try {
    await discoverTally(process.env.TALLY_HOST || DEFAULT_HOST, process.env.TALLY_PORT);
    await fetchFromTally('Balance Sheet', 'balance_sheet_response.xml');
    await fetchFromTally('Balance Sheet', 'balance_sheet_summary.xml', null, false);
    await fetchFromTally('Profit and Loss', 'profit_loss_response.xml');
    await fetchFromTally('Group Summary', 'loans_liability_response.xml', 'Loans (Liability)');
    await fetchFromTally('Group Summary', 'current_liabilities_response.xml', 'Current Liabilities');
    await fetchFromTally('Group Summary', 'indirect_expenses_response.xml', 'Indirect Expenses');
    await fetchFromTally('Group Summary', 'sundry_debtors_response.xml', 'Sundry Debtors');
    await fetchFromTally('Group Summary', 'sundry_creditors_response.xml', 'Sundry Creditors');
    await fetchCurrentAssets();
    info('All Tally financial data extracted successfully.');
  } catch (e) { errorLog(`Extraction failed: ${e.message}`); process.exitCode = 1; }
}

module.exports = { discoverTally, fetchFromTally, fetchCurrentAssets, generateRequestXML, generateCurrentAssetsXML, postXml, requestWithRetry, normalizeHost, normalizePort };
if (require.main === module) main();
