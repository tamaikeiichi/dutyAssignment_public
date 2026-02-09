const tableContainer = document.getElementById('table-container');
const openFileButton = document.getElementById('open-file-button');

// 1. カラムの定義をループで回す際の設定
const tableColumns = [
    {
        title: "回数",
        field: "duty_count",
        width: 50,
        frozen: true,
        headerSort: false,
        hozAlign: "center",
        editor: "input",
        editTriggerEvent: "dblclick"
    },
    {
        title: "氏名", 
        field: "name", 
        widthGrow: 1,      // 他の列が短い時に伸びる比率
        minWidth: 100,     // 最低限の幅
        frozen: true, 
        editor: "input",
        editTriggerEvent: "dblclick",
        headerSort: false
    },
];

for (let i = 1; i <= 31; i++) {
    tableColumns.push({
        title: `${i}`,
        field: `day${i}`,
        // width: 45,      // ★ 固定幅を消すか、minWidth に変更する
        minWidth: 40,      // 最低限の幅
        hozAlign: "center",
        editor: "input",
        headerSort: false, // ヘッダーでの並び替えをオフ（誤操作防止）
        // ★クリックではなくダブルクリックで編集開始にする設定
        editTriggerEvent: "dblclick"
    });
}

// 2. Tabulator本体の設定
const table = new Tabulator("#table-container", {
    height: "80vh",
    data: Array.from({ length: 20 }, () => ({ name: "" })),
    columns: tableColumns,
    
    // ★ ここが重要：コンテンツの量に合わせて列幅を自動調整する
    layout: "fitDataFill", // データに合わせて広がり、余白があれば埋める
    editTriggerEvent: "dblclick", // ダブルクリックで編集開始
    // ★ セルを選択可能にし、フォーカスを有効にする
    selectable: true, 
    tabEndNewRow: true, // Tabキーで末尾まで行ったら新しい行を作る（便利機能）
    // 入力（編集）が終わった瞬間に幅を再計算させる
    cellEdited: function(cell){
        cell.getTable().redraw(true); // データの変更に合わせてレイアウトを再描画
    },

    clipboard: true,
    clipboardPasteAction: "replace",
    clipboardPasteParser: "table",
});
// ------------------------------

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

// 曜日の定義
const dayOfWeek = ["日", "月", "火", "水", "木", "金", "土"];

// --- 1. 選択肢（プルダウン）の初期化 ---
function initSelectors() {
    const yearSelect = document.getElementById('select-year');
    const monthSelect = document.getElementById('select-month');
    const now = new Date();
    const currentYear = now.getFullYear();

    // 年：今年を中心に前後1年分を作成
    for (let y = currentYear - 1; y <= currentYear + 1; y++) {
        const opt = document.createElement('option');
        opt.value = y;
        opt.textContent = y;
        if (y === currentYear) opt.selected = true;
        yearSelect.appendChild(opt);
    }

    // 月：1〜12月
    for (let m = 1; m <= 12; m++) {
        const opt = document.createElement('option');
        opt.value = m;
        opt.textContent = m;
        if (m === now.getMonth() + 1) opt.selected = true;
        monthSelect.appendChild(opt);
    }
}

// --- 2. 選択された年月から表のカレンダーを構築する ---
let forcedHolidays = new Set(); // 強制休日設定を保持するセット

async function updateTableStructure() {
    const yearSelect = document.getElementById('select-year');
    const monthSelect = document.getElementById('select-month');
    
    if (!yearSelect || !monthSelect) return;

    const year = parseInt(yearSelect.value);
    const month = parseInt(monthSelect.value);
    const lastDay = new Date(year, month, 0).getDate();

    // 氏名の列
    const newColumns = [
        {
            title: "回数",
            field: "duty_count",
            width: 50,
            frozen: true,
            headerSort: false,
            hozAlign: "center",
            editor: "input",
            editTriggerEvent: "dblclick",
            // ヘッダー情報行と当直不要行は編集不可にする
            editable: (cell) => {
                const rowData = cell.getRow().getData();
                return !(typeof rowData.id === 'string' && (rowData.id.startsWith("header_") || rowData.id === "row_no_duty"));
            }
        },
        { 
            title: "休業日なら☑", 
            field: "name", 
            width: 120, 
            frozen: true, 
            editor: "input",
            headerSort: false,
            // ヘッダー情報行（idがheader_で始まる行）と当直不要行は編集不可にする
            editable: (cell) => {
                const rowData = cell.getRow().getData();
                return !(typeof rowData.id === 'string' && (rowData.id.startsWith("header_") || rowData.id === "row_no_duty"));
            }
        },
    ];

    // ヘッダー情報の行データを作成
    const headerData = {
        no_duty: { id: "row_no_duty", name: "当直不要" },
        date: { id: "header_date", name: "日付" },
        day: { id: "header_day", name: "曜日" },
        holiday: { id: "header_holiday", name: "祝日" },
        noon_night: { id: "header_noon_night", name: "昼夜" }
    };

    for (let d = 1; d <= lastDay; d++) {
        const date = new Date(year, month - 1, d);
        const dayNum = date.getDay();
        // ★ここで dayStr を定義します
        const dayStr = dayOfWeek[dayNum]; 
        
        const holidayName = await window.api.getHolidayName(date);
        const isNaturalRestDay = (dayNum === 0 || dayNum === 6 || holidayName);
        const isRestDay = (isNaturalRestDay || forcedHolidays.has(d));

        // --- 共通設定：カラム構成 ---
        const getCellConfig = (field, cClass) => ({
            title: isNaturalRestDay ? "" : `<input type="checkbox" class="header-checkbox" data-field="${field}" ${forcedHolidays.has(d) ? "checked" : ""}>`,
            field: field,
            minWidth: 50,
            hozAlign: "center",
            headerSort: false,
            cssClass: cClass,
            editor: "input",
            editTriggerEvent: "dblclick",
            // ヘッダー情報行は編集不可にする
            editable: (cell) => {
                const rowData = cell.getRow().getData();
                return !(typeof rowData.id === 'string' && rowData.id.startsWith("header_"));
            }
        });

        if (isRestDay) {
            const rClass = (dayNum === 6 && !holidayName) ? "sat" : "sun";
            
            // データ行に値をセット
            headerData.no_duty[`day${d}_noon`] = "";
            headerData.no_duty[`day${d}_night`] = "";
            headerData.date[`day${d}_noon`] = d;
            headerData.date[`day${d}_night`] = d;
            headerData.day[`day${d}_noon`] = dayStr;
            headerData.day[`day${d}_night`] = dayStr;
            headerData.holiday[`day${d}_noon`] = holidayName || "";
            headerData.holiday[`day${d}_night`] = holidayName || "";
            headerData.noon_night[`day${d}_noon`] = "昼";
            headerData.noon_night[`day${d}_night`] = "夜";

            newColumns.push(getCellConfig(`day${d}_noon`, `${rClass}-cell`));
            newColumns.push(getCellConfig(`day${d}_night`, `${rClass}-cell`));
        } else {
            const rClass = "weekday-cell";
            
            // データ行に値をセット
            headerData.no_duty[`day${d}`] = "";
            headerData.date[`day${d}`] = d;
            headerData.day[`day${d}`] = dayStr;
            headerData.holiday[`day${d}`] = holidayName || "";

            newColumns.push(getCellConfig(`day${d}`, rClass));
        }
    }

    // カラムをセット
    if (typeof table !== 'undefined') {
        table.setColumns(newColumns);

        // データ行の構築
        const tableData = [
            headerData.no_duty,
            headerData.date,
            headerData.day,
            headerData.holiday
        ];
        tableData.push(headerData.noon_night);

        // 通常のデータ行（空行）を追加
        const rowCount = 20;
        for (let i = 0; i < rowCount; i++) {
            tableData.push({ id: i, name: "", duty_count: "" });
        }
        
        table.setData(tableData);
    }
}

// --- 3. 実行指示 ---
initSelectors();

// 2. Tabulatorのセットアップが完了したら、初期描画を行う
table.on("tableBuilt", function(){
    console.log("Tabulatorの準備ができたので、初期描画を実行します");
    updateTableStructure();
});

// 3. 各種イベントリスナーの設定
const updateButton = document.getElementById('update-table-button');
if (updateButton) {
    updateButton.addEventListener('click', updateTableStructure);
}

// プルダウン変更時の自動更新
const yearSel = document.getElementById('select-year');
const monthSel = document.getElementById('select-month');
if (yearSel && monthSel) {
    // 年月変更時は強制休日設定をリセットして再描画
    const onDateChange = () => {
        forcedHolidays.clear();
        updateTableStructure();
    };
    yearSel.addEventListener('change', onDateChange);
    monthSel.addEventListener('change', onDateChange);
}

// ヘッダーのチェックボックス変更時の処理（イベント委譲）
document.addEventListener('change', (e) => {
    if (e.target && e.target.classList.contains('header-checkbox')) {
        const field = e.target.dataset.field;
        const match = field.match(/^day(\d+)/); // day1, day1_noon などから数字を抽出
        if (match) {
            const d = parseInt(match[1], 10);
            if (e.target.checked) {
                forcedHolidays.add(d);
            } else {
                forcedHolidays.delete(d);
            }
            updateTableStructure();
        }
    }
});