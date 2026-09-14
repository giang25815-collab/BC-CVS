const XLSX = require('xlsx');
const wb = XLSX.readFile('Team_CamGiang_Report.xlsx');
const s = wb.Sheets['Chi tiết SKU FamilyMart'];

console.log('Calculating sum of static SKU columns (I to BB) for each row:');
let grandTotalStatic = 0;
for (let r = 6; r <= 40; r++) {
    const addr = s[`A${r}`] ? s[`A${r}`].v : '';
    let rowSum = 0;
    // Iterate over all cell keys for this row from I to BB
    for (let c = 8; c <= 53; c++) { // Col 8 is I, Col 53 is BB
        const colLetter = XLSX.utils.encode_col(c);
        const cell = s[`${colLetter}${r}`];
        if (cell && typeof cell.v === 'number') {
            rowSum += cell.v;
        }
    }
    grandTotalStatic += rowSum;
    console.log(`Row ${r}: sum=${rowSum} | ${addr}`);
}

console.log('\nGrand Total across 35 stores from static SKU columns (I..BB):', grandTotalStatic);
