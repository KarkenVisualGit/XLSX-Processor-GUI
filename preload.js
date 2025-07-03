const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    selectInputFile: () => ipcRenderer.invoke('select-input-file'),
    selectPriceFile: () => ipcRenderer.invoke('select-price-file'),
    getCustomSavePath: (defaultFileName) => ipcRenderer.invoke('get-custom-save-path', defaultFileName),
    processBatchXLSX: (dirPath, priceJsonPath, customSavePath) =>
        ipcRenderer.invoke('process-batch-xlsx', dirPath, priceJsonPath,customSavePath),
    selectFolder: () => ipcRenderer.invoke('select-folder'),
    processXLSX: (inputPath, priceJsonPath) =>
        ipcRenderer.invoke('process-xlsx', inputPath, priceJsonPath),
    onLogMessage: (callback) =>
        ipcRenderer.on('log-message', (_event, log) => callback(log))
});