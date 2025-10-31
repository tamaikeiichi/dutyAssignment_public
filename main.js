// main.js

// 必要なモジュールに ipcMain と clipboard を追加
const { app, BrowserWindow, ipcMain, clipboard, dialog } = require('electron');
const { spawn } = require('child_process');
const path = require('path');

const isDev = !app.isPackaged; // アプリがパッケージ化されているかで開発環境かを判断
let mainWindow; // ウィンドウオブジェクトをスコープ外で参照できるようにする

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 400,
        height: 300,
        title: '当直表作成アプリ',
        icon: path.join(__dirname, 'assets', 'icon.png'), // ウィンドウのアイコンを設定
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            contextIsolation: true,
            nodeIntegration: false
        }
    });

    // レンダラープロセスからの 'read-clipboard' イベントを待ち受ける
    // `ipcMain.handle`は、`ipcRenderer.invoke`と対になる
    ipcMain.handle('read-clipboard', () => {
        // クリップボードからテキストを読み取って返す
        return clipboard.readText();
    });

    // レンダラープロセスから 'run-python-script' イベントを受け取る
    ipcMain.handle('run-python-script', async (event, filePath) => {
 
        // 開発時はvenvのPythonを、本番時はPyInstallerのexeを使用
        const scriptPath = path.join(__dirname, 'dutyAssign.py');
        const venvPython = path.join(__dirname, '.venv', 'Scripts', 'python.exe');
        const packagedExe = path.join(process.resourcesPath, 'app', 'dutyAssign.exe');

        const command = isDev ? venvPython : packagedExe;
        const args = isDev ? [scriptPath, filePath] : [filePath];
        const pythonProcess = spawn(command, args, { encoding: 'utf8' });
        
        return new Promise((resolve) => {
            let result = '';
            let error = '';
    
            // spawn自体が失敗した場合のエラーハンドリング (例: コマンドが見つからない)
            pythonProcess.on('error', (err) => {
                resolve({ success: false, message: `Failed to start Python script: ${err.message}` });
            });

            // Pythonからの標準出力を受け取る
            pythonProcess.stdout.on('data', (data) => {
                result += data.toString();
            });
    
            // Pythonからの標準エラー出力を受け取る
            pythonProcess.stderr.on('data', (data) => {
                error += data.toString();
            });
    
            // プロセスの終了を待つ
            pythonProcess.on('close', (code) => {
                if (code !== 0) {
                    resolve({ success: false, message: `Python script exited with code ${code}: ${error}` });
                } else {
                    resolve({ success: true, message: result });
                }
            });
        });
    });

    // ファイル選択ダイアログを開くためのIPCハンドラを追加
    ipcMain.handle('open-file-dialog', async () => {
        const { canceled, filePaths } = await dialog.showOpenDialog({
            properties: ['openFile'],
            filters: [
                { name: 'Excel Files', extensions: ['xlsx', 'xls', 'xlsm'] },
                { name: 'All Files', extensions: ['*'] }
            ]
        });
        if (canceled) {
            return null; // キャンセルされた場合はnullを返す
        }
        return filePaths[0]; // 選択されたファイルのパスを返す
    });

    // メッセージボックスを表示するためのIPCハンドラを追加
    ipcMain.handle('show-message-box', async (event, options) => {
        return await dialog.showMessageBox(mainWindow, options);
    });

    mainWindow.loadFile('index.html');
    //mainWindow.webContents.openDevTools(); // この行をコメントアウトすると、起動時にデベロッパーツールが開かなくなります。
}

app.whenReady().then(() => {
    createWindow();

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});