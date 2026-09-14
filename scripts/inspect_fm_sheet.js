const XLSX = require('xlsx');
const wb = XLSX.readFile('Team_CamGiang_Report.xlsx');
const s = wb.Sheets['Chi tiết SKU FamilyMart'];
const data = XLSX.utils.sheet_to_json(s, {header: 1, defval: ''});

let sumCol6 = 0;
for (let r = 5; r < data.length - 1; r++) {
    const row = data[r];
    const addr = row[0];
    const target = row[5];
    const actual = row[6];
    sumCol6 += (typeof actual === 'number' ? actual : 0);
    console.log(`Row ${r + 1}: actual=${actual} | ${addr}`);
}
console.log('\nSum of col 6 (r5..r39):', sumCol6);
const lastRow = data[data.length - 1];
console.log('Bottom total row:', lastRow.slice(0, 10));

// Also let's check formulas in sheet cell objects
for (let r = 6; r <= 41; r++) {
    const cellG = s[`G${r}`];
    if (cellG && cellG.f) {
        console.log(`G${r} formula:`, cellG.f, 'val:', cellG.v);
    }
}
const totalCellG = s['G41'];
if (totalCellG) {
    console.log('G41 formula:', totalCellG.f, 'val:', totalCellG.v);
}
