const XLSX = require('xlsx');
const fs = require('fs');

const files = fs.readdirSync('.').filter(f => f.includes('HNTRINH_KD6'));
const wb = XLSX.readFile(files[0]);
const s = wb.Sheets['NPP'];
const rows = XLSX.utils.sheet_to_json(s, {header: 1, defval: ''});

const stores = {};
let grandTotal = 0;

for (let i = 1; i <= 2597; i++) {
    const r = rows[i];
    const ocode = String(r[4] || '').trim();
    const oname = String(r[5] || '').trim();
    const oaddr = String(r[7] || '').trim();
    const amt = Number(r[12]) || 0; // Amount
    grandTotal += amt;

    const key = oaddr;
    if (!stores[key]) {
        stores[key] = { ocode, oname, oaddr, amt: 0, count: 0 };
    }
    stores[key].amt += amt;
    stores[key].count++;
}

console.log('Grand Total Amount of rows 1..2597:', grandTotal);
console.log('Total distinct store addresses:', Object.keys(stores).length);

const list = Object.values(stores).sort((a, b) => b.amt - a.amt);
list.forEach((st, idx) => {
    console.log(`${idx + 1}. [${st.ocode}] ${st.amt} đ | ${st.oname} | ${st.oaddr}`);
});
