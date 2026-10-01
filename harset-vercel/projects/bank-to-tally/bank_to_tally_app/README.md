# Bank Statement → Tally Bank Import Converter

This version keeps the original UI/workflow and uses the supplied **Bank Import Template-imp.xls** as the output format reference.

## Tally import output

The generated Excel follows the same seven columns and row structure as the supplied template:

1. Date
2. Narration
3. Chq./Ref.No.
4. Value Dt
5. Withdrawal Amt.
6. Deposit Amt.
7. Closing Balance

The first three rows are also reproduced from the supplied template structure. Amounts use `-` on the side with no transaction, matching the sample template.

### Closing Balance
If the source parser does not have a trustworthy closing balance, the converter leaves it blank. It never invents a balance from an unknown opening balance.

### XLS and XLSX
The application always creates a Tally-format XLSX. If LibreOffice/soffice is installed and available on PATH, it also offers a legacy `.xls` download. This is useful when the Tally environment specifically requires the older XLS format.

## Excel password support

The **Excel Password** field is retained. Password-protected Office Excel files are decrypted in memory using `msoffcrypto-tool` before parsing.

The parser also tries legacy `.xls`/BIFF using `xlrd`, including bank files whose extension is incorrectly `.xlsx`.

## Run on Windows

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py --server.address 0.0.0.0 --server.port 8502
```

Open:

`http://localhost:8502`

## Original workflow preserved

- Bank Ledger Name
- Default Contra Ledger (Suspense)
- Excel Password
- Bank Statement upload
- Automatic transaction detection
- Review & Edit
- Tally Bank Import Preview
- Tally Import XLSX
- Tally Import XLS when LibreOffice is available
- Tally XML

No Ledger Master screen was added.
