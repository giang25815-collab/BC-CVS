const fs = require('fs');

function inspectFile(filename) {
    if (!fs.existsSync(filename)) return;
    const content = fs.readFileSync(filename, 'utf8');
    const regex = /"store_code":\s*"([^"]+)",\s*"store_address":\s*"([^"]+)",\s*"actual":\s*([0-9]+)/g;
    let match;
    let sum = 0;
    let count = 0;
    while ((match = regex.exec(content)) !== null) {
        if (match[1].startsWith('CVS_CIRCLEK')) {
            count++;
            sum += parseInt(match[3]);
        }
    }
    if (count === 0) {
        // Try alternate format
        const regex2 = /store_code:\s*['"]([^'"]+)['"],\s*store_address:\s*['"]([^'"]+)['"],\s*actual:\s*([0-9]+)/g;
        while ((match = regex2.exec(content)) !== null) {
            if (match[1].startsWith('CVS_CIRCLEK')) {
                count++;
                sum += parseInt(match[3]);
            }
        }
    }
    console.log(`${filename} -> Circle K stores: ${count}, Total Actual: ${sum}`);
}

['master_data.js', 'index.html', 'MasterData.gs', 'CRypto.html', 'android/app/src/main/assets/www/index.html', 'android/app/src/main/assets/www/master_data.js'].forEach(inspectFile);
