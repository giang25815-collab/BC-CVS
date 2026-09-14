const fs = require('fs');

const exactValues = JSON.parse(fs.readFileSync('scripts/circlek_exact_values.json', 'utf8'));

// 1. Update master_data.js
let mdContent = fs.readFileSync('master_data.js', 'utf8');
let mdUpdated = 0;
for (const code in exactValues) {
    const val = exactValues[code].actual;
    // Look for store_code: "code" followed by actual: ...
    const pattern = new RegExp(`("store_code":\\s*"${code.replace(/\./g, '\\.')}"[\\s\\S]*?"actual":\\s*)[0-9]+`, 'g');
    if (pattern.test(mdContent)) {
        mdContent = mdContent.replace(pattern, `$1${val}`);
        mdUpdated++;
    }
}
fs.writeFileSync('master_data.js', mdContent, 'utf8');
console.log(`Updated master_data.js: ${mdUpdated} stores`);

// 2. Update index.html
let indexContent = fs.readFileSync('index.html', 'utf8');
let indexUpdated = 0;
for (const code in exactValues) {
    const val = exactValues[code].actual;
    const pattern = new RegExp(`("store_code":\\s*"${code.replace(/\./g, '\\.')}"[\\s\\S]*?"actual":\\s*)[0-9]+`, 'g');
    if (pattern.test(indexContent)) {
        indexContent = indexContent.replace(pattern, `$1${val}`);
        indexUpdated++;
    }
}
fs.writeFileSync('index.html', indexContent, 'utf8');
console.log(`Updated index.html: ${indexUpdated} stores`);

// 3. Update MasterData.gs
let gsContent = fs.readFileSync('MasterData.gs', 'utf8');
let gsUpdated = 0;
for (const code in exactValues) {
    const val = exactValues[code].actual;
    const pattern = new RegExp(`(["']store_code["']:\\s*['"]${code.replace(/\./g, '\\.')}['"][\\s\\S]*?["']actual["']:\\s*)[0-9]+`, 'g');
    if (pattern.test(gsContent)) {
        gsContent = gsContent.replace(pattern, `$1${val}`);
        gsUpdated++;
    }
}
fs.writeFileSync('MasterData.gs', gsContent, 'utf8');
console.log(`Updated MasterData.gs: ${gsUpdated} stores`);

// 4. Update CRypto.html if exists
if (fs.existsSync('CRypto.html')) {
    let crContent = fs.readFileSync('CRypto.html', 'utf8');
    let crUpdated = 0;
    for (const code in exactValues) {
        const val = exactValues[code].actual;
        const pattern = new RegExp(`("store_code":\\s*"${code.replace(/\./g, '\\.')}"[\\s\\S]*?"actual":\\s*)[0-9]+`, 'g');
        if (pattern.test(crContent)) {
            crContent = crContent.replace(pattern, `$1${val}`);
            crUpdated++;
        }
    }
    fs.writeFileSync('CRypto.html', crContent, 'utf8');
    console.log(`Updated CRypto.html: ${crUpdated} stores`);
}
