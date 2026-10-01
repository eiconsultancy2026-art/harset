const ExcelJS = require('exceljs');
const fs = require('fs');

async function processTemplate() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('Format for trust and RWAs.xlsx');

    const sheets = ['Balance Sheet', 'Statement of P&L'];

    for (const sheetName of sheets) {
        const ws = wb.getWorksheet(sheetName);
        if (!ws) continue;

        ws.eachRow(row => {
            // Replace year headers
            row.eachCell(cell => {
                if (typeof cell.value === 'string') {
                    if (cell.value.includes('31 March 20XX')) {
                        if (cell.address.includes('D')) {
                            cell.value = cell.value.replace('31 March 20XX', '31 March {LY}');
                        } else if (cell.address.includes('F')) {
                            cell.value = cell.value.replace('31 March 20XX', '31 March {RY}');
                        }
                    }
                    if (cell.value.includes('Name of the Non-Corporate Entity')) {
                        cell.value = 'Name of the Non-Corporate Entity: {LNAME}';
                    }
                }
            });

            // Find note references
            const noteNumCell = row.getCell(3); // Column C
            if (noteNumCell.value && typeof noteNumCell.value === 'number') {
                const n = noteNumCell.value;
                const dCell = row.getCell(4); // Column D
                const fCell = row.getCell(6); // Column F
                
                // Only overwrite if there is a formula or 0
                if (dCell.type === ExcelJS.ValueType.Formula || dCell.value === 0 || dCell.value === null) {
                    dCell.value = `{L${n}}`;
                }
                if (fCell.type === ExcelJS.ValueType.Formula || fCell.value === 0 || fCell.value === null) {
                    fCell.value = `{R${n}}`;
                }
            }
        });
    }

    await wb.xlsx.writeFile('Format for trust and RWAs.xlsx');
    console.log('Successfully injected placeholders into template.');
}

processTemplate().catch(console.error);
