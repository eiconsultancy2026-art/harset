# Merge → Excel Fix (v2.0.0)

The merge endpoint treats the browser mapping objects as the authoritative source. It never substitutes an older cached JSON file during Excel generation.

Before merge:
1. Both mappings must exist.
2. Company names must match.
3. Company types must match.
4. Years must be consecutive.
5. Every expected note must be present and numeric.
6. Current-year all-zero / previous-year non-zero payloads are rejected.

After generation:
1. The workbook is reopened.
2. The `Tally Data Audit` sheet is read.
3. Every Schedule III note is compared for both current and previous years.
4. The download is returned only when validation passes.
