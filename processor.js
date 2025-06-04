const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

function parseXLSX(filepath) {
    const workbook = XLSX.readFile(filepath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet, { header: 1 });
}

function loadPriceData(priceJsonPath) {
    const raw = fs.readFileSync(priceJsonPath, 'utf-8');
    return JSON.parse(raw);
}

function saveXLSX(data, headers, outputPath) {
    const output = [headers, ...data];

    for (let i = 1; i < output.length; i++) {
        const value = output[i][5];

        if (typeof value === 'string') {
            const clean = value.replace(/\s/g, '').replace(',', '.');
            const num = parseFloat(clean);
            output[i][5] = isNaN(num) ? value : num;
        }
    }
    const ws = XLSX.utils.aoa_to_sheet(output);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Processed_data');
    XLSX.writeFile(wb, outputPath);
}

function getPriceMap(priceData) {
    const map = {};
    for (let i = 0; i < priceData.length; i++) {
        const article = priceData[i]["Артикул"];
        const price = parseFloat(priceData[i]["Цена"]);
        if (article && !isNaN(price)) {
            map[article] = price;
        }
    }
    return map;
}

async function processLatestXLSX(inputPath, priceJsonPath) {
    if (!inputPath || !priceJsonPath) {
        return 'Оба файла должны быть выбраны!';
    }
    if (!fs.existsSync(inputPath)) return 'Файл заказов не найден.';
    if (!fs.existsSync(priceJsonPath)) return 'Файл прайса не найден.';

    console.log('Передаём в processXLSX:', inputPath, priceJsonPath);

    const salesData = parseXLSX(inputPath);
    const priceData = loadPriceData(priceJsonPath);
    const priceMap = getPriceMap(priceData);


    let updates = [];
    let removedRows = [];
    let originalTotal = 0;
    let newTotal = 0;

    const headers = salesData[0];
    const data = salesData.slice(1);


    for (let i = data.length - 1; i >= 0; i--) {
        const row = data[i];
        if (row.every(cell => cell === undefined || cell === null || cell === "")) continue;
        const article = row[4];
        const status = row[9];
        const zakaz = row[0];
        const quantity = parseInt(row[21]);
        const currentSum = parseFloat(row[5]);

        if (!isNaN(currentSum)) {
            originalTotal += currentSum;
        }

        if (status !== 'Выдан' && status !== 'Возврат') {
            if (article || status || zakaz) {
                removedRows.push(
                    `Артикул: ${article || 'Не указан'}, 
                Статус: ${status || 'Пусто'}, 
                Номер заказа: ${zakaz || 'Не указан'}`);
            }
            data.splice(i, 1); // удалить строку
            continue;
        }

        if (!isNaN(quantity) && priceMap[article] !== undefined) {
            const correctSum = quantity * priceMap[article];
            newTotal += correctSum;

            if (Math.abs(correctSum - currentSum) > 0.01) {
                data[i][5] = correctSum;
                updates.push(
                    `Артикул: ${article}, 
                    Кол-во: ${quantity}, 
                    Старая сумма: ${currentSum}, 
                    Новая сумма: ${correctSum}`);
            }
        } else {
            if (!isNaN(currentSum)) {
                newTotal += currentSum;
            }
        }
    }

    const outputDir = path.dirname(inputPath);
    const baseName = path.basename(inputPath, '.xlsx');
    const outputName = `${baseName}_2.xlsx`;
    const outputPath = path.join(outputDir, outputName);
    saveXLSX(data, headers, outputPath);

    let message = '';
    if (updates.length) {
        message += 'Обновлённые строки:\n' + updates.join('\n') + '\n\n';
    }
    if (removedRows.length) {
        message += 'Удалённые строки:\n' + removedRows.join('\n') + '\n\n';
    }

    message += '\n' + `Изначальная сумма: ${originalTotal.toFixed(2)}\n`;
    message += '\n' + `Новая сумма после проверки: ${newTotal.toFixed(2)}\n`;
    message += '\n' + `Результат сохранён в: ${outputPath}`;
    console.log(updates.join('\n') + '\n\n');
    console.log(removedRows.join('\n') + '\n\n');

    return {
        updates,
        removedRows,
        originalTotal: originalTotal.toFixed(2),
        newTotal: newTotal.toFixed(2),
        outputPath
    };
}

async function processBatchXLSX(directoryPath, priceJsonPath) {
    const allFiles = fs.readdirSync(directoryPath);
    const matchedFiles = allFiles.filter(file =>
        /^ArchiveOrders-\d{6}\.xlsx$/i.test(file) && !/_\d+\.xlsx$/i.test(file)
    );

    if (!matchedFiles.length) {
        console.log('Нет подходящих файлов для обработки.');
        return 'Нет подходящих файлов.';
    }

    let allResults = [];
    let combinedRows = [];
    let combinedHeaders = null;

    for (const fileName of matchedFiles) {
        const fullPath = path.join(directoryPath, fileName);
        console.log(`\n--- Обработка файла: ${fileName} ---\n`);

        try {
            // Обрабатываем файл (удаляем строки, пересчитываем суммы и т.п.)
            const result = await processLatestXLSX(fullPath, priceJsonPath);
            allResults.push({ file: fileName, result });

            // Повторно читаем обработанный файл (он уже сохранён как *_2.xlsx)
            const processedPath = result.outputPath;
            const salesData = parseXLSX(processedPath);

            const headers = salesData[0];
            const rows = salesData.slice(1);

            if (!combinedHeaders) combinedHeaders = headers;
            combinedRows.push(...rows);
        } catch (err) {
            console.log(`Ошибка при обработке ${fileName}: ${err.message}`);
        }
    }

    // Сохраняем объединённый файл
    if (combinedRows.length && combinedHeaders) {
        const outputPath = path.join(directoryPath, 'Combined_Processed_Data.xlsx');
        saveXLSX(combinedRows, combinedHeaders, outputPath);
        console.log(`Общий файл сохранён: ${outputPath}`);
    }

    return allResults;
}


module.exports = { processLatestXLSX, processBatchXLSX };