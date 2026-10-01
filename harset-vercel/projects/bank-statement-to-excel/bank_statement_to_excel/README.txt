BANK STATEMENT TO ACCOUNTING VOUCHER EXCEL
==========================================

Files:
- app.py                         Python/Streamlit application
- AccountingVouchers(2).xlsx     Your Excel template
- requirements.txt               Required Python packages

INSTALL
-------
Open PowerShell in this folder and run:

    pip install -r requirements.txt

RUN
---

    streamlit run app.py

The browser will open automatically. Upload the bank statement PDF and click
Generate Excel.

MAPPING
-------
Voucher Date             <- Transaction Date
Voucher Type Name        <- CR = Receipt, DR = Payment
Voucher Number          <- S.No
Ledger Name              <- Particulars
Ledger Amount            <- Amount(INR)
Ledger Amount Dr/Cr      <- CR/DR

Fields that are present in the Excel template but are not supplied by the
bank statement are intentionally left blank.

The original Accounting Voucher header and Read Me sheet are retained.
