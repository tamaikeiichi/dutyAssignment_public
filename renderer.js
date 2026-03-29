const tableContainer = document.getElementById('table-container');
const openFileButton = document.getElementById('open-file-button');

// --- カスタムUndo/Redo機能（通常の編集とペースト両対応） ---
const customHistory = {
    undoStack: [],
    redoStack: [],

    clear: function() {
        this.undoStack = [];
        this.redoStack = [];
    },

    pushEdit: function(row, field, oldVal, newVal) {
        this.undoStack.push({ type: "edit", row: row, field: field, oldVal: oldVal, newVal: newVal });
        this.redoStack = []; // 新しい操作をしたらRedoはクリア
    },

    pushPaste: function(actions) {
        this.undoStack.push({ type: "paste", actions: actions });
        this.redoStack = [];
    },

    undo: function() {
        if (this.undoStack.length === 0) return;
        const action = this.undoStack.pop();
        this.redoStack.push(action);
        this._apply(action, "oldVal");
    },

    redo: function() {
        if (this.redoStack.length === 0) return;
        const action = this.redoStack.pop();
        this.undoStack.push(action);
        this._apply(action, "newVal");
    },

    _apply: function(action, valKey) {
        if (action.type === "edit") {
            action.row.update({ [action.field]: action[valKey] });
        } else if (action.type === "paste") {
            const rowUpdates = new Map();
            action.actions.forEach(a => {
                if (!rowUpdates.has(a.row)) rowUpdates.set(a.row, {});
                rowUpdates.get(a.row)[a.field] = a[valKey];
            });
            rowUpdates.forEach((updateObj, row) => row.update(updateObj));
        }
        if (typeof table !== "undefined") table.redraw(true);
    }
};

// キーボードでのUndo/Redoを監視
document.addEventListener("keydown", function(e) {
    // セル内で文字入力中（編集中）は、ブラウザ標準の文字単位のUndoに任せるため無視する
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;

    if (e.ctrlKey && e.key.toLowerCase() === "z") {
        e.preventDefault();
        customHistory.undo();
    }
    if (e.ctrlKey && e.key.toLowerCase() === "y") {
        e.preventDefault();
        customHistory.redo();
    }
});

// 2. Tabulator本体の設定
const table = new Tabulator("#table-container", {
    height: "calc(100vh - 105px)", // 下側に少し余白を持たせるため、引き算の値を調整
    data: Array.from({ length: 20 }, () => ({ name: "" })),
    columns: [], // 初期設定は空にしておき、起動直後に動的に構築する

    // --- すべての列に対する共通設定 ---
    columnDefaults: {
        minWidth: 15, // Tabulatorのデフォルト制限(40px)を解除し、限界まで狭くできるようにする
        // マウスオーバー時に、セル内に文字が収まりきらず省略されている場合のみツールチップを表示する
        tooltip: function(e, cell) {
            const el = cell.getElement();
            // 要素の中身の幅(scrollWidth)が、実際の表示幅(clientWidth)を超えているか判定
            return el.scrollWidth > el.clientWidth ? cell.getValue() : null;
        }
    },
    tooltipGenerationDelay: 0, // マウスオーバー後、即座に（遅延なしで）ツールチップを表示する
    
    // layout: "fitDataFill", // 無理に余白を埋めて列幅を広げるのを防ぐため無効化
    editTriggerEvent: "dblclick", // ダブルクリックで編集開始
    // ★ セルを選択可能にし、フォーカスを有効にする
    selectable: true, 
    tabEndNewRow: true, // Tabキーで末尾まで行ったら新しい行を作る（便利機能）
    // 入力（編集）が終わった瞬間に幅を再計算させる
    cellEdited: function(cell){
        // 編集完了時に独自のUndo履歴に保存
        customHistory.pushEdit(cell.getRow(), cell.getField(), cell.getOldValue(), cell.getValue());
        cell.getTable().redraw(true); // データの変更に合わせてレイアウトを再描画
    },

    // カスタムペースト処理を使用するため、デフォルトのペーストは無効化（コピーのみ残す）
    clipboard: "copy",
    // clipboardPasteAction: "replace",
    // clipboardPasteParser: "table",

    // --- 右クリックメニュー（コンテキストメニュー）の設定 ---
    rowContextMenu: [
        {
            label: "元に戻す (Ctrl+Z)",
            action: function(e, row) {
                customHistory.undo();
            },
            disabled: function() {
                return customHistory.undoStack.length === 0; // 履歴がない時は無効化
            }
        },
        {
            label: "やり直し (Ctrl+Y)",
            action: function(e, row) {
                customHistory.redo();
            },
            disabled: function() {
                return customHistory.redoStack.length === 0; // やり直し履歴がない時は無効化
            }
        }
    ]
});

// --- ペースト先の基準セルを記憶する処理 ---
let targetPasteCell = null;

// セルにマウスオーバーまたはクリックしたときに、ペーストの始点として記録
table.on("cellMouseEnter", function(e, cell) {
    targetPasteCell = cell;
});
table.on("cellClick", function(e, cell) {
    targetPasteCell = cell;
});

// --- カスタムペースト処理（Excelのような部分ペースト） ---
document.addEventListener("paste", function(e) {
    if (!targetPasteCell) return;

    // セルをダブルクリックして文字入力中（編集モード中）であれば、この処理は無視して通常の文字ペーストをさせる
    if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;

    const clipboardData = e.clipboardData || window.clipboardData;
    const pastedText = clipboardData.getData("text/plain");
    if (!pastedText) return;

    e.preventDefault(); // 画面全体に対する不要なデフォルトペーストをキャンセル

    // 改行で分割して行ごとの配列にし、さらにタブ区切りでセルごとの2次元配列にする
    let rows = pastedText.split(/\r\n|\n|\r/);
    if (rows.length > 0 && rows[rows.length - 1] === "") {
        rows.pop(); // Excelコピー時の末尾の空行を除去
    }
    const dataMatrix = rows.map(row => row.split("\t"));

    const startRow = targetPasteCell.getRow();
    const startColumn = targetPasteCell.getColumn();
    
    // 現在表示されている行・列のリストを取得
    const allRows = table.getRows("active");
    const allColumns = table.getColumns();

    const startRowIndex = allRows.findIndex(r => r === startRow);
    const startColIndex = allColumns.findIndex(c => c === startColumn);

    if (startRowIndex === -1 || startColIndex === -1) return;

    // 起点セルから順に右と下へデータをセットしていく
    const pasteActions = []; // ペースト履歴保存用
    dataMatrix.forEach((rowData, i) => {
        const targetRow = allRows[startRowIndex + i];
        if (!targetRow) return; // ペースト範囲が行数を超える場合は無視

        const updateObj = {};
        rowData.forEach((val, j) => {
            const targetCol = allColumns[startColIndex + j];
            if (!targetCol) return; // ペースト範囲が列数を超える場合は無視

            const field = targetCol.getField();
            const cell = targetRow.getCell(field);
            
            if (cell) {
                // 対象列の編集可否（editable）のルールを確認
                const colDef = targetCol.getDefinition();
                let isEditable = true;
                if (typeof colDef.editable === "function") {
                    isEditable = colDef.editable(cell);
                } else if (colDef.editable === false) {
                    isEditable = false;
                }

                // 編集可能なセルのみ更新（日付のヘッダー行などを上書き破壊から保護する）
                if (isEditable && cell.getValue() !== val) {
                    pasteActions.push({
                        row: targetRow,
                        field: field,
                        oldVal: cell.getValue(),
                        newVal: val
                    });
                    updateObj[field] = val;
                }
            }
        });

        // 対象行のデータを一括更新
        if (Object.keys(updateObj).length > 0) {
            targetRow.update(updateObj);
        }
    });

    // 変更があった場合は履歴に追加
    if (pasteActions.length > 0) {
        customHistory.pushPaste(pasteActions);
        table.redraw(true); // ペースト後にレイアウトを再描画
    }
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
let forcedHolidays = new Set(); // 強制休日設定を保持するセット (YYYY-MM-DD形式)

async function updateTableStructure() {
    customHistory.clear(); // 新しい表を描画するので履歴をリセット

    const yearSelect = document.getElementById('select-year');
    const monthSelect = document.getElementById('select-month');
    
    if (!yearSelect || !monthSelect) return;

    const year = parseInt(yearSelect.value);
    const month = parseInt(monthSelect.value);
    const lastDay = new Date(year, month, 0).getDate();

    // 前月の情報を取得
    const prevMonthDate = new Date(year, month - 1, 0);
    const prevYear = prevMonthDate.getFullYear();
    const prevMonth = prevMonthDate.getMonth() + 1;
    const prevLastDay = prevMonthDate.getDate();

    // 表示する日付リストを作成（前月の最後10日分 ＋ 当月分）
    const displayDays = [];
    // 前月の最後10日
    for (let d = prevLastDay - 9; d <= prevLastDay; d++) {
        displayDays.push({
            year: prevYear,
            month: prevMonth,
            date: d,
            isCurrentMonth: false,
            fieldPrefix: `prev_day${d}`
        });
    }
    // 当月
    for (let d = 1; d <= lastDay; d++) {
        displayDays.push({
            year: year,
            month: month,
            date: d,
            isCurrentMonth: true,
            fieldPrefix: `day${d}`
        });
    }

    // 氏名の列
    const newColumns = [
        {
            title: "仮当直回数",
            field: "duty_count",
            width: 30,
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
            width: 100, 
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

    for (const dayInfo of displayDays) {
        const { year: y, month: m, date: d, isCurrentMonth, fieldPrefix } = dayInfo;
        const dateObj = new Date(y, m - 1, d);
        const dayNum = dateObj.getDay();
        const dayStr = dayOfWeek[dayNum]; 
        
        const holidayName = await window.api.getHolidayName(dateObj);
        const isNaturalRestDay = (dayNum === 0 || dayNum === 6 || holidayName);
        
        // 強制休日用のキー（YYYY-MM-DD）を作成
        const dateKey = `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
        const isRestDay = (isNaturalRestDay || forcedHolidays.has(dateKey));

        // --- 共通設定：カラム構成 ---
        const getCellConfig = (field, cClass) => ({
            title: isNaturalRestDay ? "" : `<input type="checkbox" class="header-checkbox" data-field="${field}" data-date="${dateKey}" ${forcedHolidays.has(dateKey) ? "checked" : ""}>`,
            field: field,
            width: 30, // 横幅を40に完全固定して自動拡張を防ぐ
            hozAlign: "center",
            headerSort: false,
            cssClass: cClass,
            editor: "input",
            editTriggerEvent: "dblclick",
            // ヘッダー情報行は編集不可にする
            editable: (cell) => {
                const rowData = cell.getRow().getData();
                return !(typeof rowData.id === 'string' && (rowData.id.startsWith("header_") || rowData.id === "row_no_duty"));
            },
            // --- セルの表示形式をカスタマイズ ---
            formatter: (cell) => {
                const rowData = cell.getRow().getData();
                const val = cell.getValue();
                
                // 日付の行の場合、幅が狭い時に「前半（月）を省略して後半（日）が見える」ようにする
                if (rowData.id === "header_date" && val != null) {
                    // direction: rtl (右から左) を使って左側に「...」を出し、bdiタグで文字自体の反転を防ぐ
                    // margin を使ってセルの余白(padding)の限界まで表示領域を広げる
                    return `<div style="direction: rtl; text-align: center; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; width: calc(100% + 8px); margin: 0 -4px;"><bdi dir="ltr">${val}</bdi></div>`;
                }
                return val != null ? val : "";
            }
        });

        // 日付の表示内容。前月は月を付けて「m/d」形式にし、当月は「d」のみにする
        const dateDisplay = isCurrentMonth ? d : `${m}/${d}`;

        if (isRestDay) {
            const rClass = (dayNum === 6 && !holidayName) ? "sat" : "sun";
            
            // データ行に値をセット
            headerData.no_duty[`${fieldPrefix}_noon`] = "";
            headerData.no_duty[`${fieldPrefix}_night`] = "";
            headerData.date[`${fieldPrefix}_noon`] = dateDisplay;
            headerData.date[`${fieldPrefix}_night`] = dateDisplay;
            headerData.day[`${fieldPrefix}_noon`] = dayStr;
            headerData.day[`${fieldPrefix}_night`] = dayStr;
            headerData.holiday[`${fieldPrefix}_noon`] = holidayName || "";
            headerData.holiday[`${fieldPrefix}_night`] = holidayName || "";
            headerData.noon_night[`${fieldPrefix}_noon`] = "昼";
            headerData.noon_night[`${fieldPrefix}_night`] = "夜";

            newColumns.push(getCellConfig(`${fieldPrefix}_noon`, `${rClass}-cell`));
            newColumns.push(getCellConfig(`${fieldPrefix}_night`, `${rClass}-cell`));
        } else {
            const rClass = "weekday-cell";
            
            // データ行に値をセット
            headerData.no_duty[fieldPrefix] = "";
            headerData.date[fieldPrefix] = dateDisplay;
            headerData.day[fieldPrefix] = dayStr;
            headerData.holiday[fieldPrefix] = holidayName || "";

            newColumns.push(getCellConfig(fieldPrefix, rClass));
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
        const dateKey = e.target.dataset.date;
        if (dateKey) {
            if (e.target.checked) {
                forcedHolidays.add(dateKey);
            } else {
                forcedHolidays.delete(dateKey);
            }
            updateTableStructure();
        }
    }
});