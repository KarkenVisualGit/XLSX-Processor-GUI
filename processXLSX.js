/* const folderPath = path.join(__dirname, 'data');
const fileRegex = /^ArchiveOrders-(\d{2})(\d{2})(\d{2})\.xlsx$/i;
const priceFile = 'kaspi_price.xlsx'; */

async function processLatestXLSX(inputPath, pricePath) {
    try {
        console.log('Загрузка входного файла:', inputPath);
        console.log('Загрузка прайс-файла:', pricePath);

        const inputWb = XLSX.readFile(inputPath);
        const inputWs = inputWb.Sheets[inputWb.SheetNames[0]];
        const salesData = XLSX.utils.sheet_to_json(inputWs, { header: 1 });

        const priceWb = XLSX.readFile(pricePath);
        const priceWs = priceWb.Sheets[priceWb.SheetNames[0]];
        const priceData = XLSX.utils.sheet_to_json(priceWs, { header: 1 });

        const priceMap = {};
        for (let i = 1; i < priceData.length; i++) {
            const article = priceData[i][1]; // Колонка B
            const price = parseFloat(priceData[i][2]); // Колонка C
            if (article && !isNaN(price)) priceMap[article] = price;
        }
        const headers = salesData[0];
        const output = [headers];
        let updates = [];
        let removedRows = [];
        let originalTotal = 0;
        let newTotal = 0;

        for (let i = 1; i < salesData.length; i++) {
            const row = salesData[i];
            if (row.every(cell => cell === undefined || cell === null || cell === "")) continue;
            const article = row[4]; // E
            const status = row[9];  // J
            const zakaz = row[0];
            const qty = parseInt(row[21]); // V
            const sum = parseFloat(row[5]); // F

            if (!isNaN(sum)) originalTotal += sum;

            if (status !== 'Выдан' && status !== 'Возврат') {
                if (article || status || zakaz) {
                    removedRows.push(
                        `Артикул: ${article || 'Не указан'}, 
                        Статус: ${status || 'Пусто'}, 
                        Номер заказа: ${zakaz || 'Не указан'}`);
                }
                continue;
            }

            if (!isNaN(qty) && priceMap[article] !== undefined) {
                const correctSum = qty * priceMap[article];
                if (Math.abs(correctSum - sum) > 0.01) {
                    row[5] = correctSum;
                    updates.push(
                        `Артикул: ${article}, 
                        Кол-во: ${qty}, 
                        Старая сумма: ${sum}, 
                        Новая сумма: ${correctSum}`);
                }
                if (!isNaN(correctSum)) newTotal += correctSum;
            } else {
                newTotal += sum;
            }

            output.push(row);
        }

        for (let i = 1; i < output.length; i++) {
            const value = output[i][5];

            // Преобразуем только если value существует
            if (typeof value === 'string') {
                // Удалим пробелы, заменим запятые на точки
                const clean = value.replace(/\s/g, '').replace(',', '.');

                const num = parseFloat(clean);
                output[i][5] = isNaN(num) ? value : num; // если не число — оставим как есть
            }
        }
        // Запись нового файла


        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(output);
        XLSX.utils.book_append_sheet(wb, ws, 'ОбработанныеДанные');
        const dir = path.dirname(inputPath);
        const extName = path.extname(inputPath);
        const baseName = path.basename(inputPath, extName);
        const outPath = path.join(dir, `${baseName}_2${extName}`);
        XLSX.writeFile(wb, outPath);
        // Вывод статистики
        console.log("Обновлённые строки:\n", updates.join("\n"));
        console.log("Удалённые строки:\n", removedRows.join("\n"));
        console.log(`Изначальная сумма: ${originalTotal.toFixed(2)}`);
        console.log(`Новая сумма после проверки: ${newTotal.toFixed(2)}`);

        return 'Готово. Выходной файл: ' + outPath;
    } catch (err) {
        console.log('Ошибка обработки:', err);
        return 'Ошибка: ' + err.message;
    }
};