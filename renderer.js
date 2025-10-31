const tableContainer = document.getElementById('table-container');
const openFileButton = document.getElementById('open-file-button');

// Pythonスクリプトを実行し、結果を通知する関数
async function executePythonScript(filePath) {
    if (!filePath) {
        console.log('ファイルパスが指定されていません。');
        return;
    }
    console.log('Pythonスクリプトの実行対象ファイルパス:', filePath);
    // メインプロセスにPythonスクリプトの実行を依頼し、結果を受け取る
    const result = await window.api.runPythonScript(filePath);
    console.log('Python script result:', result);
    // 結果をネイティブのダイアログで表示
    if (result.success) {
        await window.api.showMessageBox({
            type: 'info',
            title: '成功',
            message: 'Pythonの実行に成功しました',
            detail: result.message
        });
    } else {
        await window.api.showMessageBox({
            type: 'error',
            title: 'エラー',
            message: 'Pythonの実行に失敗しました',
            detail: result.message
        });
    }
}

// ファイル選択ボタンの処理
openFileButton.addEventListener('click', async () => {
    // メインプロセスにファイル選択ダイアログの表示を依頼（最も確実な方法）
    const filePath = await window.api.openFileDialog();
    if (filePath) {
        // ファイルが選択されたら、Pythonスクリプトを実行
        await executePythonScript(filePath);
    }
});