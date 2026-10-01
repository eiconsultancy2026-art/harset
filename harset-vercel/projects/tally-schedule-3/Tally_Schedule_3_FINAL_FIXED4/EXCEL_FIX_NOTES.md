# Excel Fix Notes (v2.0.0)

The old failure mode was a mismatch between the data displayed by the browser and the data written to Excel. The fixed implementation passes the browser objects directly to the generator and validates the resulting workbook by reading it back.

The validator does not depend on fragile row labels. It uses the `Tally Data Audit` sheet emitted by each generator and validates both year columns for every supported Schedule III note.

A workbook with a missing or mismatched note is rejected and is not offered as a successful download.
