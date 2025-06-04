const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    selectInputFile: () => ipcRenderer.invoke('select-input-file'),
    selectPriceFile: () => ipcRenderer.invoke('select-price-file'),
    processBatchXLSX: (dirPath, priceJsonPath) =>
        ipcRenderer.invoke('process-batch-xlsx', dirPath, priceJsonPath),
    selectFolder: () => ipcRenderer.invoke('select-folder'),
    processXLSX: (inputPath, priceJsonPath) =>
        ipcRenderer.invoke('process-xlsx', inputPath, priceJsonPath),
    onLogMessage: (callback) =>
        ipcRenderer.on('log-message', (_event, log) => callback(log))
});