let inputFilePath = null;
let priceFilePath = null;
let lastOutputPath = '';


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