# Tally Schedule III Mapper – Fixed Build v2.0.0

A Node.js/Express application that connects to TallyPrime's XML server, maps financial data to Schedule III notes, displays the mapping in the web portal, and generates comparative Excel workbooks.

## What was fixed

### 1. Browser → Excel regression
The merge endpoint now uses the exact current/previous mapping objects submitted by the browser. It does not reload an older cached JSON file.

If the current-year mapping is entirely zero while the previous year contains data, Excel generation is stopped rather than producing a misleading zero-filled workbook.

### 2. Full Excel read-back validation
Every generated workbook contains a `Tally Data Audit` sheet. Before the download response is returned, the server reopens the workbook and compares every expected Schedule III note for **both years** against the browser payload.

This catches:
- current-year values becoming zero;
- previous-year values becoming zero;
- missing notes;
- incorrect note values;
- incomplete audit coverage.

### 3. Merge integrity
Merge now requires:
- the same company in both periods;
- the same company type in both periods;
- consecutive financial years;
- complete numeric note mappings;
- successful Excel validation.

### 4. Mapping integrity
Partnership/Proprietorship/HUF mapping no longer uses the balance-sheet total to silently manufacture Note 10 or Note 18. Tally totals are retained as reconciliation diagnostics.

Note 18 is sourced from Tally's `Other Current Assets` group. Note 25 is based on explicit indirect-expense ledger classification.

### 5. Tally connection
The default host is now `localhost`. A LAN address can still be supplied through `.env`.

## Supported entity types

- Private Limited
- Partnership Firm
- Proprietorship
- HUF
- RWA
- Trust

## Installation

Requirements:
- Node.js 18+ recommended
- TallyPrime with XML/HTTP server enabled
- TallyPrime company loaded

```powershell
cd Tally_Schedule_3_FINAL_FIXED4
npm install
copy .env.example .env
```

Edit `.env` if Tally is on another machine:

```env
PORT=8000
HOST=0.0.0.0
TALLY_HOST=localhost
TALLY_PORT=9000
TALLY_TIMEOUT_MS=20000
```

Start:

```powershell
npm start
```

Open `http://localhost:8000`.

## Validation commands

Syntax and static checks:

```powershell
npm test
```

Excel pipeline regression test:

```powershell
npm run test:pipeline
```

The pipeline test requires `npm install` first because it uses ExcelJS.

## Project flow

```text
TallyPrime
   ↓ XML/HTTP
fetch_tally.js
   ↓ raw XML
XML parsing / normalization
   ↓
test_mapping*.js
   ↓ canonical mapping JSON
Web portal
   ↓ exact browser payload
/api/merge-mapping
   ↓
Excel generator
   ↓
Tally Data Audit
   ↓ read-back validator
Validated .xlsx download
```

## Important accounting behavior

The application does not intentionally hide an unexplained Tally-to-Schedule III difference by assigning it to `Other Current Assets` or `Other Current Liabilities`. Such differences are exposed through the `reconciliation` object in the mapping JSON and must be investigated in the Tally mapping.

## Key new files

- `schedule3_model.js` – canonical note extraction and merge contract.
- `excel_validator.js` – Excel read-back validation for both years.
- `test_excel_pipeline.js` – Excel regression test.
- `FIXES_APPLIED.md` – detailed change list.
