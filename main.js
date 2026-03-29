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
    let windowState = { width: 400, height: 300 };

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

    // ファイル保存ダイアログを開くためのIPCハンドラを追加
    ipcMain.handle('show-save-dialog', async (event, options) => {
        const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, options);
        if (canceled) {
            return null; // キャンセルされた場合はnullを返す
        }
        return filePath; // 選択されたファイルのパスを返す
    });

    // ExcelエクスポートのためのIPCハンドラを追加
    ipcMain.handle('export-to-excel', async (event, data, filePath, columnFields) => {
        // 開発時はvenvのPythonを、本番時はPyInstallerのexeを使用
        const scriptPath = path.join(__dirname, 'export_excel.py'); // 新しいPythonスクリプト
        const venvPython = path.join(__dirname, '.venv', 'Scripts', 'python.exe');
        // Note: For packaged app, export_excel.py would also need to be bundled.
        // For simplicity, let's assume it's always run via python for now,
        // or packaged into its own exe if needed.
        // For a packaged app, you might need to bundle export_excel.py separately or include it in dutyAssign.exe.
        // For now, let's just use the python interpreter for both dev and prod for this new script.
        const command = isDev ? venvPython : 'python'; // Assuming python is in PATH for packaged app, or specify full path
        const args = isDev ? [scriptPath, filePath, JSON.stringify(columnFields)] : [scriptPath, filePath, JSON.stringify(columnFields)];

        return new Promise((resolve) => {
            const pythonProcess = spawn(command, args);
            let stdout = '';
            let stderr = '';

            pythonProcess.stdin.write(JSON.stringify(data)); // Send table data via stdin
            pythonProcess.stdin.end();

            pythonProcess.stdout.on('data', (data) => { stdout += data.toString(); });
            pythonProcess.stderr.on('data', (data) => { stderr += data.toString(); });

            pythonProcess.on('close', (code) => {
                if (code !== 0) {
                    resolve({ success: false, message: `Python export script exited with code ${code}: ${stderr}` });
                } else {
                    resolve({ success: true, message: stdout });
                }
            });
        });
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

// 祝日判定の依頼を受ける
ipcMain.handle('check-holiday', (event, date) => {
    // 送られてきたデータをDateオブジェクトに変換して判定
    return JapaneseHolidays.isHoliday(new Date(date));
});