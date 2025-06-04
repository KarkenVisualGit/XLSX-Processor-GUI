let inputFilePath = null;
let priceFilePath = null;
let lastOutputPath = '';

const selectBatchFolderBtn = document.getElementById('select-batch-folder');
const batchFolderPathSpan = document.getElementById('batch-folder-path');
const runBatchProcessingBtn = document.getElementById('run-batch-processing');
const batchOutput = document.getElementById('batch-output');

let selectedBatchFolder = '';
let selectedPriceFile = '';

selectBatchFolderBtn.addEventListener('click', async () => {
    const result = await window.electronAPI.selectFolder();
    if (!result.canceled && result.filePaths.length > 0) {
        selectedBatchFolder = result.filePaths[0];
        batchFolderPathSpan.textContent = selectedBatchFolder;
    }
});

runBatchProcessingBtn.addEventListener('click', async () => {
    if (!selectedBatchFolder) {
        batchOutput.textContent = 'Пожалуйста, выберите папку с Excel-файлами.';
        return;
    }

    const result = await window.electronAPI.selectPriceFile();
    if (!result) {
        batchOutput.textContent = 'Пожалуйста, выберите файл прайса.';
        return;
    }

    selectedPriceFile = result;

    batchOutput.textContent = 'Обработка файлов...';

    const processingResults = await window.electronAPI.processBatchXLSX(selectedBatchFolder, selectedPriceFile);

    if (!processingResults || !Array.isArray(processingResults)) {
        batchOutput.textContent = 'Ошибка при обработке файлов.';
        return;
    }

    let logText = '';
    for (const fileResult of processingResults) {
        logText += `📄 Файл: ${fileResult.file}\n`;
        const res = fileResult.result;

        if (typeof res === 'string') {
            logText += `❌ ${res}\n\n`;
            continue;
        }

        if (res.updates.length) {
            logText += `🔧 Обновления:\n${res.updates.join('\n')}\n`;
        }

        if (res.removedRows.length) {
            logText += `🗑 Удалено:\n${res.removedRows.join('\n')}\n`;
        }

        logText += `💰 Сумма до: ${res.originalTotal}\n`;
        logText += `💰 Сумма после: ${res.newTotal}\n`;
        logText += `📁 Сохранено в: ${res.outputPath}\n`;
        logText += `-----------------------------\n\n`;
    }

    batchOutput.textContent = logText;
});


document.getElementById('selectInputBtn').addEventListener('click', async () => {
    const result = await window.electronAPI.selectInputFile();
    if (result) {
        inputFilePath = result;
        document.getElementById('inputFileName').textContent = result;
    }
});

document.getElementById('processBtn').addEventListener('click', async () => {
    const status = document.getElementById('status');
    const log = document.getElementById('log');
    log.innerHTML = '';

    if (!inputFilePath) {
        status.textContent = 'Пожалуйста, выберите файл заказов';
        return;
    }
    status.textContent = 'Обработка...';

    const priceFilePath = await window.electronAPI.selectPriceFile();

    const result = await window.electronAPI.processXLSX(inputFilePath, priceFilePath);

    if (!result || typeof result !== 'object') {
        status.textContent = 'Ошибка при обработке данных';
        return;
    }

    lastOutputPath = result.outputPath;

    status.textContent = 'Обработка завершена';

});

window.electronAPI.onLogMessage((result) => {
    const logEl = document.getElementById('log');
    logEl.innerHTML = '';

    const section = (title, items) => {
        const container = document.createElement('div');
        container.style.display = 'flex';
        container.style.flexDirection = 'column';
        container.style.gap = '4px';
        container.style.marginBottom = '20px';

        const header = document.createElement('h3');
        header.textContent = title;
        container.appendChild(header);

        items.forEach(text => {
            const line = document.createElement('div');
            line.textContent = text;
            line.style.borderBottom = '1px solid #ccc';
            container.appendChild(line);
        });

        return container;
    };

    if (result.updates?.length) {
        logEl.appendChild(section('Обновлённые строки', result.updates));
    }

    if (result.removedRows?.length) {
        logEl.appendChild(section('Удалённые строки', result.removedRows));
    }

    const totals = document.createElement('div');
    totals.style.display = 'flex';
    totals.style.flexDirection = 'column';
    totals.innerHTML = `
        <div>Изначальная сумма: ${result.originalTotal}</div>
        <div>Новая сумма после проверки: ${result.newTotal}</div>
        <div>Результат сохранён в: ${result.outputPath}</div>
    `;
    logEl.appendChild(totals);
    logEl.scrollTop = logEl.scrollHeight;
});

document.getElementById('saveBtn').addEventListener('click', async () => {
    if (!lastOutputPath) {
        alert('Сначала обработайте файл!');
        return;
    }

    const result = await window.electronAPI.saveXLSX(lastOutputPath);
    document.getElementById('status').textContent = result.message;
});