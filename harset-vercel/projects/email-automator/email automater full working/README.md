# Client Excel Request Email Automater

This version replaces the old Advance Tax/DOCX workflow.

It sends **one common Excel template as an attachment** to every trust/client listed in a second Excel file. The email greeting is personalized using the trust name.

## Required input files

Put exactly these two `.xlsx` files in `input/`.

### 1. Trust master Excel

Required columns:

| S.No | Name of Trust | email |
|---:|---|---|
| 1 | ABC Trust | abc@example.com |
| 2 | XYZ Trust | xyz@example.com |

The program identifies this file automatically from the column names, so the filename does not matter.

### 2. Client-input Excel template

Required columns:

| Sl. No. | Pre Acknowledgement Number | ID Code | Unique Registration Number (URN) | Name of donor | Address of donor | Donation Type | Mode of receipt | Amount of donation (Indian rupees) |
|---:|---|---|---|---|---|---|---|---:|

This is the **same Excel file that will be attached to every client email**. The program does not alter or personalize the attachment.

## Run on Windows PowerShell

Open PowerShell in this folder and run:

```powershell
py -m pip install -r requirements.txt
py email_automation.py
```

If `py` is unavailable, use:

```powershell
python -m pip install -r requirements.txt
python email_automation.py
```

## Gmail setup

Use a Gmail **App Password**, not your normal Gmail password. The program asks for it at runtime and does not save it.

## What happens

1. The program finds and validates both Excel files.
2. It reads `S.No`, `Name of Trust`, and `email` from the master file.
3. It validates the client-input template columns.
4. It asks for the sender Gmail and App Password.
5. It asks for the email subject and body.
6. `{trust_name}` can be used in the body and is replaced automatically for each client.
7. It displays a recipient preview.
8. Nothing is sent until you type `SEND` exactly.
9. The same client-input Excel template is attached to every email.
10. Failed emails do not stop the remaining emails.
11. A delivery log is written to `output/delivery_log.csv`.

## Example email

```text
Dear {trust_name},

Please find attached the Excel format for providing the required donation details.

Kindly fill in the required details in the attached Excel sheet and send the completed sheet back to us at your earliest convenience.

Please ensure that all applicable fields are filled in accurately.

Regards,
[Your Name]
[Company/Firm Name]
```

For `ABC Trust`, `{trust_name}` becomes `ABC Trust`.

## Important

Close the Excel files in Microsoft Excel before running the program if Windows reports that a file is locked.


## Desktop GUI

Run `run_gui.bat` or `py email_automation_gui.py`. The GUI lets you edit Sender Gmail, Subject, and Email Body before sending. Use `{trust_name}` to insert the trust name automatically.
