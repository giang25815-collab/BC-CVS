const XLSX = require('xlsx');
const fs = require('fs');

// Read SO sheet from KD6
const files = fs.readdirSync('.').filter(f => f.includes('HNTRINH_KD6'));
console.log('Loading KD6 file:', files[0]);
const wb = XLSX.readFile(files[0]);
const soSheet = wb.Sheets['SO'];
const soRows = XLSX.utils.sheet_to_json(soSheet, {header: 1, defval: ''});

// Let's find columns
const header = soRows[0];
let colMaso = -1, colTen = -1, colDiachi = -1, colTtien = -1;
for (let c = 0; c < header.length; c++) {
    const h = String(header[c]).toLowerCase();
    if (h.includes('mã khách') || h.includes('mã kh')) colMaso = c;
    if (h.includes('tên khách') || h.includes('tên kh')) colTen = c;
    if (h.includes('địa chỉ') || h.includes('dia chi')) colDiachi = c;
    if (h.includes('tiền') || h.includes('thành tiền') || h.includes('doanh số')) colTtien = c;
}

let ckKhoKhoAmt = 0;
const ckStoresSo = {}; // storeKey -> total amount

for (let r = 1; r < soRows.length; r++) {
    const row = soRows[r];
    const ma = String(row[colMaso] || '').trim();
    const ten = String(row[colTen] || '').trim();
    const diachi = String(row[colDiachi] || '').trim();
    const ttien = Number(row[colTtien]) || 0;

    if (ma === 'VT4050' || ma === 'VT3013' || ma === 'VT3014' || ma === 'VT3015') {
        const addrLower = diachi.toLowerCase();
        if (addrLower.includes('nam tân uyên') || addrLower.includes('tân uyên') || addrLower.includes('lô g1-9') || ten.toLowerCase().includes('kho khô') || diachi.toLowerCase().includes('kho khô')) {
            ckKhoKhoAmt += ttien;
        } else {
            const key = diachi || ten;
            ckStoresSo[key] = (ckStoresSo[key] || 0) + ttien;
        }
    }
}

const totalCkStoresSystem = Object.keys(ckStoresSo).length;
const tCk = totalCkStoresSystem > 0 ? Math.round(ckKhoKhoAmt / totalCkStoresSystem) : 0;
console.log('Kho Kho Amt:', ckKhoKhoAmt);
console.log('Total CK Stores System:', totalCkStoresSystem);
console.log('tCk (Kho Kho per store):', tCk);

// Compare with Team_CamGiang_Report.xlsx Circle K stores
const wbReport = XLSX.readFile('Team_CamGiang_Report.xlsx');
const ckSheet = wbReport.Sheets['Circle K'];
if (ckSheet) {
    const ckData = XLSX.utils.sheet_to_json(ckSheet, {header: 1, defval: ''});
    console.log('\n--- Team_CamGiang_Report.xlsx [Circle K] sheet ---');
    let reportTotal = 0;
    for (let r = 0; r < ckData.length; r++) {
        const row = ckData[r];
        const rep = row[1];
        const code = row[2];
        const addr = row[3];
        const actual = row[4];
        if (typeof actual === 'number' && actual > 0) {
            console.log(`${rep} | ${code} | ${actual} | ${addr}`);
            reportTotal += actual;
        }
    }
    console.log('Total Circle K in Team_CamGiang_Report.xlsx:', reportTotal);
}
