const { contextBridge, ipcRenderer } = require('electron');


contextBridge.exposeInMainWorld('api', {
    readClipboard: () => ipcRenderer.invoke('read-clipboard'),
    runPythonScript: (filePath) => ipcRenderer.invoke('run-python-script', filePath),
    openFileDialog: () => ipcRenderer.invoke('open-file-dialog'),
    showMessageBox: (options) => ipcRenderer.invoke('show-message-box', options),
    // 祝日判定関数を追加
    getHolidayName: (date) => ipcRenderer.invoke('check-holiday', date),
    showSaveDialog: (options) => ipcRenderer.invoke('show-save-dialog', options),
    exportToExcel: (data, filePath, columnFields) => ipcRenderer.invoke('export-to-excel', data, filePath, columnFields)
});