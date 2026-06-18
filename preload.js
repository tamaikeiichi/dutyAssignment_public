const { contextBridge, ipcRenderer, webFrame } = require('electron');


contextBridge.exposeInMainWorld('api', {
    readClipboard: () => ipcRenderer.invoke('read-clipboard'),
    runPythonScript: (filePath) => ipcRenderer.invoke('run-python-script', filePath),
    openFileDialog: () => ipcRenderer.invoke('open-file-dialog'),
    showMessageBox: (options) => ipcRenderer.invoke('show-message-box', options),
    getHolidayName: (date) => ipcRenderer.invoke('check-holiday', date),
    readFileBase64: (filePath) => ipcRenderer.invoke('read-file-base64', filePath),
    writeTempFile: (base64) => ipcRenderer.invoke('write-temp-file', base64),
    openResultWindow: (filePath, year, month) => ipcRenderer.invoke('open-result-window', filePath, year, month),
    getResultFile: () => ipcRenderer.invoke('get-result-file'),
    showSaveDialog: (defaultPath) => ipcRenderer.invoke('show-save-dialog', defaultPath),
    saveFile: (filePath, base64) => ipcRenderer.invoke('save-file', filePath, base64),
    setZoomFactor: (f) => webFrame.setZoomFactor(f),
    getZoomFactor: () => webFrame.getZoomFactor(),
    resizeWindow: (w, h) => ipcRenderer.invoke('resize-window', w, h),
    getResultMeta: () => ipcRenderer.invoke('get-result-meta'),
});