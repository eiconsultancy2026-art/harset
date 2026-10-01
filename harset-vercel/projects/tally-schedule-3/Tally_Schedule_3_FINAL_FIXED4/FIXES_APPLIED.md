# Tally Schedule III – Fixes Applied (v2.0.0)

## Merge / Excel boundary
- Excel generation now uses only the exact browser payload sent by the merge request.
- Merge rejects a current-year payload that is all zero while the previous year contains data. This prevents the known zero-filled Excel regression.
- Generated Excel is read back and validated before the download response is returned.
- Validation checks **every Schedule III note for both current and previous years**, not only selected current-year balance-sheet cells.
- The validation reads the `Tally Data Audit` sheet written by every supported Excel generator, avoiding brittle label matching.

## Merge integrity checks
- Current and previous year must be consecutive.
- Company names must match case-insensitively after whitespace normalization.
- Company type must match on both years.
- The selected company type must match the generated mappings.
- Every expected note must exist and be numeric before Excel generation.

## Mapping integrity
- Partnership/Proprietorship/HUF mapping no longer overwrites Note 10 with a balance-sheet residual.
- Note 18 is sourced from Tally's `Other Current Assets` group instead of absorbing the asset-side difference.
- Note 25 is built from explicit indirect-expense ledger classification, excluding employee benefits, depreciation/amortisation and finance-cost ledgers, instead of silently absorbing an unexplained residual.
- Tally balance-sheet totals are retained as reconciliation diagnostics rather than being used to manufacture Schedule III note values.
- Note 4 selection treats an actual zero as a valid value instead of falling through to a calculated fallback.
- Reconciliation diagnostics are stored in the mapping JSON so unresolved Tally-vs-Schedule III differences are visible.

## Tally connectivity
- Default Tally host is now `localhost` unless `TALLY_HOST` is explicitly configured.
- Existing configured LAN/IP hosts continue to work.

## New modules
- `schedule3_model.js` – canonical note extraction, company/type normalization and payload validation.
- `excel_validator.js` – read-back validation of generated Excel against both browser datasets.
- `test_excel_pipeline.js` – regression test for previous-year/current-year Excel mismatches.

## Important operational rule
If the reconciliation diagnostics report a non-zero gap, the application does **not** hide the difference by putting it into `Other Current Assets` or `Other Current Liabilities`. The gap must be investigated in the Tally mapping instead.
