# Tally Schedule III - Setup

## 1. Install Node.js
Use a current Node.js LTS release.

## 2. Install dependencies
From this folder:

```powershell
npm install
```

## 3. Configure Tally
Copy `.env.example` to `.env`. Tally is auto-discovered on ports 9000 through 9005. If Tally is on another host, set `TALLY_HOST`. If you know the port, set `TALLY_PORT`; otherwise leave it blank.

## 4. Check the refactored modules

```powershell
npm run check
```

## 5. Test Tally directly

```powershell
node fetch_tally.js
```

## 6. Start the web application

```powershell
npm start
```

The web server uses the same robust Tally discovery layer for company-list requests and mapping extraction.

## 7. Important
The supplied Excel templates are intentionally preserved. Generated `.xlsx` files, XML responses, cache files and logs are ignored by Git.


## Correct Tally server configuration

If TallyPrime is running on the same computer used for this project, configure `TALLY_HOST=172.16.1.13` and `TALLY_PORT=9000`, or enter the same values in **Server Settings**. The browser settings are passed to every company-list and mapping request.

The company list is independent of financial year. Select the same company on both sides and use different date ranges, for example `01-Apr-2025` to `31-Mar-2026` versus `01-Apr-2024` to `31-Mar-2025`. Tally's XML request receives the selected company and period for each mapping run.

## Troubleshooting

Run `Test-NetConnection 172.16.1.13 -Port 9000` and require `TcpTestSucceeded : True`. Then run `node fetch_tally.js`. Generated workbooks include a **Tally Data Audit** sheet showing each Schedule III note and detailed ledger breakdowns returned by Tally.
