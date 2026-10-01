# Bank Statement -> Tally Accounting Voucher Converter

This application converts bank statements into the supplied TallyPrime Accounting Voucher Excel structure.

## Target structure
The output uses the exact headers from `AccountingVouchers_Tally_Template.xlsx`:
Voucher Date, Voucher Type Name, Voucher Number, Buyer/Supplier - Address,
Buyer/Supplier - Pincode, Ledger Name, Ledger Amount, Ledger Amount Dr/Cr,
Item Name, Billed Quantity, Item Rate, Item Rate per, Item Amount, Change Mode.

## Accounting logic
- Bank withdrawal/debit -> Payment voucher: counter ledger Dr, bank ledger Cr.
- Bank deposit/credit -> Receipt voucher: bank ledger Dr, counter ledger Cr.
- Unmatched transactions use Suspense by default.
- Optional Ledger Master supports Keyword, Ledger columns.
- Each transaction creates two ledger rows with the same Voucher Number.
- A Conversion Review sheet preserves source narration/reference and matching status.

## Run
python -m pip install -r requirements.txt
python -m streamlit run app.py

Or double-click `run_windows.bat`.

For password-protected Excel, enter the workbook password in the sidebar before uploading.
