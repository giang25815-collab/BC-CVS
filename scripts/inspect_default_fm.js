const fs = require('fs');
const path = require('path');

let content = fs.readFileSync('master_data.js', 'utf8');
content += '\nmodule.exports = { FM_CATEGORIES, DEFAULT_FM_STORES, MASTER_DATA };';
const mdPath = path.resolve(__dirname, 'temp_md_runner.js');
fs.writeFileSync(mdPath, content, 'utf8');
const md = require(mdPath);

let countWithVal = 0;
let totalVal = 0;
md.DEFAULT_FM_STORES.forEach((s, idx) => {
    if (s.actual > 0) countWithVal++;
    totalVal += s.actual;
    console.log(`${idx + 1}. [${s.ten_pg}] actual: ${s.actual} | target: ${s.target} | ${s.addr.slice(0, 40)}`);
});
console.log(`DEFAULT_FM_STORES count with val: ${countWithVal}, Total val: ${totalVal}`);
fs.unlinkSync(mdPath);
