// main.js

// 必要なモジュールに ipcMain と clipboard を追加
const { app, BrowserWindow, ipcMain, clipboard, dialog } = require('electron');
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const JapaneseHolidays = require('japanese-holidays');
//const { ipcMain } = require('electron');

const isDev = !app.isPackaged; // アプリがパッケージ化されているかで開発環境かを判断
let mainWindow; // ウィンドウオブジェクトをスコープ外で参照できるようにする

function createWindow() {
    // ウィンドウサイズと位置の保存先
    const statePath = path.join(app.getPath('userData'), 'window-state.json');
    let windowState = { width: 800, height: 300 };

    // 保存された状態があれば読み込む
    try {
        if (fs.existsSync(statePath)) {
            const savedState = JSON.parse(fs.readFileSync(statePath, 'utf8'));
            windowState = { ...windowState, ...savedState };
        }
    } catch (e) {
        console.log('Failed to load window state:', e);
    }

    mainWindow = new BrowserWindow({
        width: windowState.width,
        height: windowState.height,
        x: windowState.x,
        y: windowState.y,
        title: '当直表作成アプリ',
        icon: path.join(__dirname, 'assets', 'icon.ico'), // ウィンドウのアイコンを設定
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            contextIsolation: true,
            nodeIntegration: false
        }
    });

    // ウィンドウが閉じられる直前にサイズと位置を保存
    mainWindow.on('close', () => {
        try {
            const bounds = mainWindow.getBounds();
            fs.writeFileSync(statePath, JSON.stringify(bounds));
        } catch (e) {
            console.log('Failed to save window state:', e);
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
        lastKibouFilePath = filePath; // kibouファイルパスを記録（当直不要データ取得用）
 
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

    // ファイルをBase64文字列として読み込む（Buffer転送の不安定さを避けるため文字列で返す）
    ipcMain.handle('read-file-base64', async (event, filePath) => {
        return fs.readFileSync(filePath).toString('base64');
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

    // 一時ファイルへの書き込み（Python入力用）
    ipcMain.handle('write-temp-file', async (event, base64) => {
        const tempPath = path.join(app.getPath('temp'), `duty_input_${Date.now()}.xlsx`);
        fs.writeFileSync(tempPath, Buffer.from(base64, 'base64'));
        return tempPath;
    });

    // 結果ファイルを新ウィンドウで表示
    let lastResultFilePath = null;
    let lastResultYear = null;
    let lastResultMonth = null;
    let lastResultScore = null;
    let lastKibouFilePath = null;
    ipcMain.handle('open-result-window', async (event, filePath, year, month, score) => {
        lastResultFilePath = filePath;
        lastResultYear  = year  ?? null;
        lastResultMonth = month ?? null;
        lastResultScore = score ?? null;
        const resultWindow = new BrowserWindow({
            width: 1100, height: 600,
            title: '当直表の結果',
            webPreferences: {
                preload: path.join(__dirname, 'preload.js'),
                contextIsolation: true,
                nodeIntegration: false
            }
        });
        await resultWindow.loadFile('result.html');
        resultWindow.webContents.openDevTools(); // デバッグ用：DevToolsを自動で開く
    });

    ipcMain.handle('get-result-file', () => lastResultFilePath);
    ipcMain.handle('get-result-meta', () => ({ year: lastResultYear, month: lastResultMonth, score: lastResultScore, kibouFilePath: lastKibouFilePath }));

    ipcMain.handle('resize-window', (event, width, height) => {
        const win = BrowserWindow.fromWebContents(event.sender);
        if (win) win.setContentSize(Math.round(width), Math.round(height));
    });

    ipcMain.handle('show-save-dialog', async (event, originalPath) => {
        const fileName = originalPath ? path.basename(originalPath) : 'result.xlsx';
        const downloadsDir = app.getPath('downloads');
        const { canceled, filePath } = await dialog.showSaveDialog({
            defaultPath: path.join(downloadsDir, fileName),
            filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
        });
        return canceled ? null : filePath;
    });

    ipcMain.handle('save-file', async (event, filePath, base64) => {
        fs.writeFileSync(filePath, Buffer.from(base64, 'base64'));
    });

    mainWindow.loadFile('index.html');
    // mainWindow.webContents.openDevTools();
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

// 祝日判定の依頼を受ける
ipcMain.handle('check-holiday', (event, date) => {
    // 送られてきたデータをDateオブジェクトに変換して判定
    return JapaneseHolidays.isHoliday(new Date(date));
});