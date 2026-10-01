TEMPLE ACCOUNTING - LAN MULTI-TEMPLE WITH SUPER ADMIN

ACCESS MODEL
- Super Admin is a separate system account.
- Only Super Admin can:
  1. Register/add a temple.
  2. Create the temple username and password.
  3. Modify a temple name/address/login username.
  4. Reset a temple password.
  5. Modify existing accounting transactions.
  6. Delete accounting transactions.
  7. View the complete cross-temple audit log.
- Temple users can only access their own temple's records.
- Temple users can create new income/expense entries but cannot edit or delete an existing entry.
- Authorization is enforced on the Flask server routes, so hiding UI buttons is not the only protection.

FIRST-RUN SUPER ADMIN
Username: superadmin
Password: SuperAdmin@123

IMPORTANT: Change this password in a production deployment. The application stores passwords as hashes.

LEGACY DEMO ACCOUNT
The original demo temple remains available after migration:
Username: admin
Password: admin123

RUNNING
1. Install Python 3.11+.
2. Open PowerShell in this folder.
3. Create/activate a virtual environment if needed:
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
4. Install dependencies:
   python -m pip install -r requirements.txt
5. Start:
   python app.py
6. Open on host PC: http://127.0.0.1:8000
7. For another device on the same Wi-Fi, run ipconfig and open:
   http://<HOST-PC-WIFI-IP>:8000

DATABASE
- temple_accounts.db is migrated automatically on first start.
- Existing temple users are retained.
- Existing plain-text demo passwords are converted to secure password hashes.

SECURITY NOTES
- Set TEMPLE_APP_SECRET to a long random secret before real deployment.
- Do not expose port 8000 to the public Internet.
- Restrict Windows Firewall to the intended private LAN.
- Back up temple_accounts.db regularly.
