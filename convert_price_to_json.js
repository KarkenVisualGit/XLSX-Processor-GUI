const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');

const inputXlsxPath = path.join(__dirname, 'data', 'kaspi_price.xlsx');
const outputJsonPath = path.join(__dirname, 'data', 'kaspi_price.json');

function convertPriceXLSXtoJSON() {
    const workbook = xlsx.readFile(inputXlsxPath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { defval: '' });

    fs.mkdirSync(path.dirname(outputJsonPath), { recursive: true });
    fs.writeFileSync(outputJsonPath, JSON.stringify(data, null, 2), 'utf8');

    console.log(`✅ Успешно сохранено: ${outputJsonPath}`);
}

convertPriceXLSXtoJSON();
