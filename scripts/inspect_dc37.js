const XLSX = require('xlsx');
const fs = require('fs');

const files = fs.readdirSync('.').filter(f => f.includes('HNTRINH_KD6'));
const wb = XLSX.readFile(files[0]);
const s = wb.Sheets['NPP'];
const rows = XLSX.utils.sheet_to_json(s, {header: 1, defval: ''});

console.log('Searching for DC37 or Vietsing in NPP sheet:');
let totalDC37 = 0;
for (let i = 1; i <= 2597; i++) {
    const r = rows[i];
    const ocode = String(r[4] || '');
    const oname = String(r[5] || '');
    const oaddr = String(r[7] || '');
    const icode = String(r[8] || '');
    const iname = String(r[9] || '');
    const amt = Number(r[12]) || 0;
    const cat = String(r[10] || '');

    if (oaddr.toLowerCase().includes('vietsing') || oaddr.toLowerCase().includes('dc37') || oname.toLowerCase().includes('dc37')) {
        totalDC37 += amt;
        console.log(`Row ${i} | code: ${ocode} | name: ${oname} | item: ${icode} - ${iname} | amt: ${amt} | addr: ${oaddr}`);
    }
}
console.log('Total DC37 amount:', totalDC37);
