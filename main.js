const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { processLatestXLSX, processBatchXLSX } = require('./processor');


const ENABLE_GUI_LOGS = true;
const originalLog = console.log;

function createWindow() {
    const win = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true,
            sandbox: false,
        }
    });

    win.loadFile('index.html');
}

app.whenReady().then(createWindow);

function sendLogToWindow(log) {
    const win = BrowserWindow.getAllWindows()[0];
    if (win && win.webContents) {
        win.webContents.send('log-message', log);
    }
}

console.log = (...args) => {
    const message = args.map(arg => (typeof arg === 'object'
        ? JSON.stringify(arg, null, 2)
        : String(arg))).join(' ');
    originalLog(message);
    if (ENABLE_GUI_LOGS) {
        sendLogToWindow(message);
    }
};

ipcMain.handle('select-input-file', async () => {
    const { canceled, filePaths } = await dialog.showOpenDialog({
        properties: ['openFile'],
        filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
    });
    return canceled ? null : filePaths[0];
});


ipcMain.handle('select-price-file', async () => {
    const staticJsonPath = path.join(__dirname, 'data', 'kaspi_price.json');
    return staticJsonPath;
});

ipcMain.handle('process-xlsx', async (_event, inputPath, staticJsonPath) => {
    const result = await processLatestXLSX(inputPath, staticJsonPath);
    sendLogToWindow(result);
    return result;
});

ipcMain.handle('process-batch-xlsx', async (_event, directoryPath, priceJsonPath) => {
    const result = await processBatchXLSX(directoryPath, priceJsonPath);
    return result;
});

ipcMain.handle('select-folder', async () => {
    const result = await dialog.showOpenDialog({
        properties: ['openDirectory']
    });
    return result;
});