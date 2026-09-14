const XLSX = require('xlsx');
const fs = require('fs');

const files = fs.readdirSync('.').filter(f => f.includes('HNTRINH_KD6'));
const wb = XLSX.readFile(files[0]);
const s = wb.Sheets['NPP'];
const rows = XLSX.utils.sheet_to_json(s, {header: 1, defval: ''});

// Collect all rows by OutletAddress and OutletCode
const nppByAddr = {};
for (let i = 1; i <= 2597; i++) {
    const r = rows[i];
    const addr = String(r[7] || '').trim();
    const amt = Number(r[12]) || 0;
    nppByAddr[addr] = (nppByAddr[addr] || 0) + amt;
}

// Read master data DEFAULT_FM_STORES
let content = fs.readFileSync('master_data.js', 'utf8');
content += '\nmodule.exports = { FM_CATEGORIES, DEFAULT_FM_STORES, MASTER_DATA };';
const path = require('path');
const mdPath = path.join(__dirname, '..', 'temp_test_md.js');
fs.writeFileSync(mdPath, content, 'utf8');
const md = require(mdPath);
const fmStores = md.DEFAULT_FM_STORES;

console.log('Total FM stores in master_data.js:', fmStores.length);

function norm(t) {
    return String(t || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

let sumMatched = 0;
let matchedCount = 0;
const results = [];

fmStores.forEach(st => {
    const n = norm(st.addr);
    let amt = 0;
    let matchKey = '';
    for (let a in nppByAddr) {
        if (norm(a) === n || (norm(a).length > 10 && n.includes(norm(a))) || (n.length > 10 && norm(a).includes(n))) {
            amt = nppByAddr[a];
            matchKey = a;
            break;
        }
    }
    if (amt > 0) matchedCount++;
    sumMatched += amt;
    results.push({
        rep: st.rep,
        target: st.target,
        actual: amt,
        addr: st.addr,
        matchKey: matchKey
    });
    console.log(`${st.rep} | target: ${st.target} | actual: ${amt} | ${st.addr}`);
});

console.log('\nMatched stores count:', matchedCount);
console.log('Total Actual Sum:', sumMatched);

fs.unlinkSync('temp_test_md.js');
