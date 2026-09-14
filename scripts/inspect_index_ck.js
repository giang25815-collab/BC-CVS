const fs = require('fs');
const html = fs.readFileSync('index.html', 'utf8');
const regex = /\{"employee_name":\s*"([^"]+)",\s*"chain":\s*"Circle K",\s*"store_code":\s*"([^"]+)",\s*"store_address":\s*"([^"]+)",\s*"actual":\s*([0-9]+)\}/g;
let match;
let sum = 0;
while ((match = regex.exec(html)) !== null) {
  sum += parseInt(match[4]);
  console.log(`${match[1]} | ${match[2]} | actual: ${match[4]} | ${match[3].slice(0, 30)}`);
}
console.log('Total CK in index.html:', sum);
