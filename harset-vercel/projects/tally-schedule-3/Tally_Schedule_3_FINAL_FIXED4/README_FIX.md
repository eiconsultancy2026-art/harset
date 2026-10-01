# Fixed Build – Tally Schedule III

### Start

```powershell
cd Tally_Schedule_3_FINAL_FIXED4
npm install
copy .env.example .env
npm start
```

Open:

```text
http://localhost:8000
```

For LAN access, use the machine's LAN IP with port `8000`.

### Tally configuration

Default:

```env
TALLY_HOST=localhost
TALLY_PORT=9000
```

If TallyPrime is running on another machine, set the actual LAN address in `.env`.

### Validation

```powershell
npm test
npm run test:pipeline
```

`npm run test:pipeline` requires dependencies to be installed first.

### Merge requirements

The Merge & Download operation requires:

- same company in both periods;
- same company type in both periods;
- consecutive financial years;
- complete numeric mapping for every expected note;
- non-zero current-year data when the previous year contains data;
- successful Excel read-back validation for both years.
