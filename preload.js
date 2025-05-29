const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    selectInputFile: () => ipcRenderer.invoke('select-input-file'),
    selectPriceFile: () => ipcRenderer.invoke('select-price-file'),
    processXLSX: (inputPath, pricePath) =>
        ipcRenderer.invoke('process-xlsx', inputPath, pricePath),
    saveXLSX: (defaultPath) => ipcRenderer.invoke('save-xlsx', defaultPath),
    onLogMessage: (callback) =>
        ipcRenderer.on('log-message', (_event, log) => callback(log))
});