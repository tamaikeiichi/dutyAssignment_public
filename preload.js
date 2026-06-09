const { contextBridge, ipcRenderer } = require('electron');


contextBridge.exposeInMainWorld('api', {
    readClipboard: () => ipcRenderer.invoke('read-clipboard'),
    runPythonScript: (filePath) => ipcRenderer.invoke('run-python-script', filePath),
    openFileDialog: () => ipcRenderer.invoke('open-file-dialog'),
    showMessageBox: (options) => ipcRenderer.invoke('show-message-box', options),
    getHolidayName: (date) => ipcRenderer.invoke('check-holiday', date),
    readFileBase64: (filePath) => ipcRenderer.invoke('read-file-base64', filePath),
    writeTempFile: (base64) => ipcRenderer.invoke('write-temp-file', base64),
    openResultWindow: (filePath) => ipcRenderer.invoke('open-result-window', filePath),
    getResultFile: () => ipcRenderer.invoke('get-result-file'),
});