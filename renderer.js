const tableContainer = document.getElementById('table-container');
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
    if (e.ctrlKey && e.key.toLowerCase() === "c") {
        const ranges = table.getRanges();
        if (ranges.length === 0) return;
        e.preventDefault();
        const tsv = ranges[0].getCells()
            .map(row => row.map(cell => cell.getValue() ?? "").join("\t"))
            .join("\n");
        navigator.clipboard.writeText(tsv);
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
    editTriggerEvent: "click", // ダブルクリックで編集開始
    selectable: false,
    selectableRange: 1,             // ドラッグによるセル範囲選択を有効化
    selectableRangeColumns: true,   // 列ヘッダークリックで列全体を選択
    selectableRangeRows: true,      // 行ヘッダークリックで行全体を選択
    selectableRangeClearCells: false,
    tabEndNewRow: true, // Tabキーで末尾まで行ったら新しい行を作る（便利機能）
    // 入力（編集）が終わった瞬間に幅を再計算させる
    cellEdited: function(cell){
        // 編集完了時に独自のUndo履歴に保存
        customHistory.pushEdit(cell.getRow(), cell.getField(), cell.getOldValue(), cell.getValue());
        cell.getTable().redraw(true); // データの変更に合わせてレイアウトを再描画
        if (cell.getField() === 'duty_count') updateProvisionalDutyCountDisplay();
    },

    clipboard: false,
    // clipboardPasteAction: "replace",
    // clipboardPasteParser: "table",

    // --- 右クリックメニュー（コンテキストメニュー）の設定 ---
    rowFormatter: function(row) {
        if (typeof row.getData().id !== 'string') return;
        const cells = row.getCells();
        if (cells[0]) {
            cells[0].getElement().style.setProperty('background-color', '#ffffff', 'important');
            cells[0].getElement().style.pointerEvents = 'none';
        }
        if (cells[1]) {
            cells[1].getElement().style.setProperty('background-color', '#ffffff', 'important');
            cells[1].getElement().style.pointerEvents = 'none';
        }
    },

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
    customHistory.clear();

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
    const canEdit = (cell) => {
        const id = cell.getRow().getData().id;
        return !(typeof id === 'string' && (id.startsWith("header_") || id === "row_no_duty" || id === "row_holiday_checkbox"));
    };
    const newColumns = [
        {
            title: "　",
            field: "name",
            width: 100,
            frozen: true,
            editor: "input",
            headerSort: false,
            editable: canEdit,
        },
        {
            title: "仮当直回数",
            field: "duty_count",
            width: 50,
            frozen: true,
            headerSort: false,
            hozAlign: "center",
            editor: "input",
            editable: canEdit,
        },
    ];

    // ヘッダー情報の行データを作成
    const headerData = {
        holiday_checkbox: { id: "row_holiday_checkbox", name: "休業日" },
        no_duty: { id: "row_no_duty", name: "当直不要" },
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

        // 日付の表示内容。前月は月を付けて「m/d」形式にし、当月は「d」のみにする
        const dateDisplay = isCurrentMonth ? d : `${m}/${d}`;

        // --- 共通設定：カラム構成 ---
        const getCellConfig = (field, cClass) => ({
            title: `<div style="direction:rtl;text-align:center;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;width:calc(100% + 8px);margin:0 -4px;"><bdi dir="ltr">${dateDisplay}</bdi></div>`,
            field: field,
            width: 30, // 横幅を40に完全固定して自動拡張を防ぐ
            hozAlign: "center",
            headerSort: false,
            cssClass: cClass,
            editor: "input",
            editTriggerEvent: "click",
            // ヘッダー情報行は編集不可にする
            editable: (cell) => {
                const rowData = cell.getRow().getData();
                return !(typeof rowData.id === 'string' && (rowData.id.startsWith("header_") || rowData.id === "row_no_duty" || rowData.id === "row_holiday_checkbox"));
            },
            // --- セルの表示形式をカスタマイズ ---
            formatter: (cell) => {
                const rowData = cell.getRow().getData();
                const val = cell.getValue();

                // 当直不要行はチェックボックスで表示
                if (rowData.id === "row_no_duty") {
                    return `<input type="checkbox" ${val === true ? "checked" : ""} style="cursor:pointer; pointer-events:none;">`;
                }
                // 休業日なら行は平日のみチェックボックスで表示（土日祝は null なので空セル）
                if (rowData.id === "row_holiday_checkbox") {
                    if (val === null || val === undefined) return "";
                    return `<input type="checkbox" ${val === true ? "checked" : ""} style="cursor:pointer; pointer-events:none;">`;
                }

                return val != null ? val : "";
            },
            cellClick: (e, cell) => {
                const rowId = cell.getRow().getData().id;
                if (rowId === "row_no_duty") {
                    cell.setValue(!cell.getValue());
                    updateDutyCountDisplay();
                } else if (rowId === "row_holiday_checkbox") {
                    if (cell.getValue() === null || cell.getValue() === undefined) return;
                    if (forcedHolidays.has(dateKey)) {
                        forcedHolidays.delete(dateKey);
                    } else {
                        forcedHolidays.add(dateKey);
                    }
                    updateTableStructure();
                }
            }
        });

        if (isRestDay) {
            const rClass = (dayNum === 6 && !holidayName) ? "sat" : "sun";
            
            // データ行に値をセット
            // 土日祝は null（チェックボックスなし）、強制休日は true
            const cbVal = isNaturalRestDay ? null : true;
            headerData.holiday_checkbox[`${fieldPrefix}_noon`] = cbVal;
            headerData.holiday_checkbox[`${fieldPrefix}_night`] = cbVal;
            headerData.no_duty[`${fieldPrefix}_noon`] = false;
            headerData.no_duty[`${fieldPrefix}_night`] = false;
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
            headerData.holiday_checkbox[fieldPrefix] = forcedHolidays.has(dateKey);
            headerData.no_duty[fieldPrefix] = false;
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
            headerData.holiday_checkbox,
            headerData.no_duty,
            headerData.day,
            headerData.noon_night,
            headerData.holiday
        ];

        // 通常のデータ行（空行）を追加
        const rowCount = 20;
        for (let i = 0; i < rowCount; i++) {
            tableData.push({ id: i, name: "", duty_count: "" });
        }
        
        await table.setData(tableData);
        updateDutyCountDisplay();
    }
}

function updateDutyCountDisplay() {
    const el = document.getElementById('duty-count-display');
    if (!el) return;
    const noDutyRow = table.getRows().find(r => r.getData().id === 'row_no_duty');
    const noDutyData = noDutyRow ? noDutyRow.getData() : {};
    let total = 0;
    let noDutyCount = 0;
    for (const col of table.getColumns()) {
        const f = col.getField();
        if (f && /^day\d+/.test(f)) {
            total++;
            if (noDutyData[f] === true) noDutyCount++;
        }
    }
    const net = total - noDutyCount;
    el.textContent = total > 0 ? `今月の当直回数: ${net}` : '';
    updateProvisionalDutyCountDisplay();
}

function updateProvisionalDutyCountDisplay() {
    const el = document.getElementById('provisional-duty-count-display');
    if (!el) return;
    const skipIds = new Set(["row_no_duty", "row_holiday_checkbox", "header_day", "header_holiday", "header_noon_night"]);
    let total = 0;
    for (const row of table.getRows()) {
        const data = row.getData();
        if (typeof data.id === 'string' && (data.id.startsWith("header_") || skipIds.has(data.id))) continue;
        const val = parseInt(data.duty_count, 10);
        if (!isNaN(val)) total += val;
    }
    el.textContent = `仮当直回数の合計: ${total}`;
}

// --- 3. 実行指示 ---
initSelectors();

// 2. Tabulatorのセットアップが完了したら、初期描画を行う
table.on("tableBuilt", function(){
    console.log("Tabulatorの準備ができたので、初期描画を実行します");
    updateTableStructure();
});

table.on("cellEdited", function(cell){
    if (cell.getField() === 'duty_count') updateProvisionalDutyCountDisplay();
});



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

// --- Excel エクスポート ---
async function exportToExcel() {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('当直表');
    const columns = table.getColumns();
    const rows = table.getRows();

    // HTML タグを除去してテキストだけ返す
    const tmp = document.createElement('div');
    const stripHtml = (html) => {
        if (!html || !html.includes('<')) return html || '';
        tmp.innerHTML = html;
        return tmp.textContent.trim();
    };

    // "rgb(r,g,b)" / "rgba(r,g,b,a)" → ExcelJS の ARGB 文字列 ("FFrrggbb")
    // 透明 (alpha=0) の場合は null を返す
    const toArgb = (rgb) => {
        const m = rgb.match(/[\d.]+/g);
        if (!m || m.length < 3) return null;
        if (rgb.includes('rgba') && parseFloat(m[3] ?? '1') === 0) return null;
        return 'FF' + [m[0], m[1], m[2]].map(n => parseInt(n).toString(16).padStart(2, '0')).join('').toUpperCase();
    };

    // DOM 要素の背景色を Excel セルに適用し、罫線を白・細線にする
    const whiteBorder = { style: 'thin', color: { argb: 'FFFFFFFF' } };
    const allWhiteBorder = { top: whiteBorder, bottom: whiteBorder, left: whiteBorder, right: whiteBorder };
    const applyStyle = (excelCell, domEl) => {
        excelCell.border = allWhiteBorder;
        if (!domEl) return;
        const argb = toArgb(window.getComputedStyle(domEl).backgroundColor);
        if (argb) excelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
    };

    // 列幅（Tabulator のピクセル幅を文字数に近似換算）
    worksheet.columns = columns.map(col => ({
        width: Math.max(4, Math.round(col.getWidth() / 7))
    }));

    // 1行目：列タイトル（日付など）
    // 列ヘッダー要素には sat/sun 系 CSS が当たらないため、最初のデータ行セルで代用
    const firstDataRow = rows.length > 0 ? rows[0] : null;
    const titleExcelRow = worksheet.addRow(columns.map(col => stripHtml(col.getDefinition().title)));
    columns.forEach((col, j) => {
        applyStyle(titleExcelRow.getCell(j + 1), firstDataRow?.getCell(col.getField())?.getElement());
    });

    // 2行目以降：テーブルの全データ行
    for (const row of rows) {
        const data = row.getData();
        const isNoDutyRow = data.id === 'row_no_duty';
        const excelRow = worksheet.addRow(columns.map(col => {
            const val = data[col.getField()];
            if (val === true)  return isNoDutyRow ? '○' : '✓';
            if (val === false || val === null || val === undefined) return '';
            return val;
        }));
        columns.forEach((col, j) => {
            applyStyle(excelRow.getCell(j + 1), row.getCell(col.getField())?.getElement());
        });
    }

    // ファイルとして保存
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const year = document.getElementById('select-year').value;
    const month = document.getElementById('select-month').value;
    a.download = `当直表_${year}年${month}月.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
}

const exportButton = document.getElementById('export-excel-button');
if (exportButton) {
    exportButton.addEventListener('click', exportToExcel);
}

// ExcelJS のセル値を安全に文字列化する
// リッチテキスト({ richText:[{text:'○'}] })や数式結果({ result:... })も正しく扱う
function cellText(cell) {
    const v = cell.value;
    if (v === null || v === undefined) return '';
    if (typeof v === 'object') {
        if (v.richText) return v.richText.map(r => r.text ?? '').join('');
        if (v.result !== undefined) return String(v.result ?? '');
        if (v.text !== undefined) return String(v.text);
    }
    return String(v);
}

// --- 先月データ読み込み ---
async function loadPrevMonthData() {
    try {
    // ファイル選択
    const filePath = await window.api.openFileDialog();
    if (!filePath) return;
    // ExcelJS でファイルを読み込む（base64経由でArrayBufferに変換）
    const base64 = await window.api.readFileBase64(filePath);
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(bytes.buffer);
    const ws = workbook.worksheets[0];
    if (!ws) { console.error('[load] ワークシートが見つかりません'); return; }
    // ファイル名から読み込んだ年月を検出
    const filename = filePath.split(/[\\/]/).pop();
    const nameMatch = filename.match(/(\d{4})年(\d{1,2})月/);
    let loadedYear, loadedMonth;
    if (nameMatch) {
        loadedYear = parseInt(nameMatch[1]);
        loadedMonth = parseInt(nameMatch[2]);
    } else {
        // ファイル名から取れない場合は現在の選択から前月を推定
        loadedMonth = parseInt(document.getElementById('select-month').value) - 1;
        loadedYear  = parseInt(document.getElementById('select-year').value);
        if (loadedMonth === 0) { loadedMonth = 12; loadedYear--; }
    }

    // Excelの行インデックス（1始まり）
    const ROW_HOLIDAY_CB  = 2;  // 休業日なら☑
    const ROW_NO_DUTY     = 3;  // 当直不要
    const ROW_NOON_NIGHT  = 6;  // 昼夜
    const ROW_DATA_START  = 7;  // 人名データ開始行
    const DATA_ROW_COUNT  = 20;

    // 列タイトルを解析して「当月分（数字のみ）」の列を dayNum ごとにグルーピング
    const colsByDay = new Map(); // dayNum → [{colIdx, noonNight}]
    ws.getRow(1).eachCell({ includeEmpty: false }, (cell, colIdx) => {
        if (colIdx <= 2) return; // 仮当直回数・氏名はスキップ
        const title = String(cell.value ?? '').trim();
        if (!title || title.includes('/')) return; // 前月分（"m/d"形式）はスキップ
        const dayNum = parseInt(title);
        if (isNaN(dayNum) || dayNum <= 0 || dayNum > 31) return;
        const noonNight = cellText(ws.getRow(ROW_NOON_NIGHT).getCell(colIdx)).trim();
        if (!colsByDay.has(dayNum)) colsByDay.set(dayNum, []);
        colsByDay.get(dayNum).push({ colIdx, noonNight });
    });

    // 読み込んだ月の最終日を計算し、最後10日だけ抽出
    const daysInLoadedMonth = new Date(loadedYear, loadedMonth, 0).getDate();
    const last10Start = daysInLoadedMonth - 9;
    const last10Days = [...colsByDay.keys()]
        .filter(d => d >= last10Start)
        .sort((a, b) => a - b);
    // 強制休日を検出して forcedHolidays に登録（テーブル再構築前にセット）
    forcedHolidays.clear();
    for (const dayNum of last10Days) {
        const cols = colsByDay.get(dayNum);
        const cbVal = cellText(ws.getRow(ROW_HOLIDAY_CB).getCell(cols[0].colIdx)).trim();
        if (cbVal === '✓') {
            const dateKey = `${loadedYear}-${String(loadedMonth).padStart(2,'0')}-${String(dayNum).padStart(2,'0')}`;
            forcedHolidays.add(dateKey);
        }
    }

    // 人名・当直回数を収集
    const personNames = [];
    const dutyCounts  = [];
    for (let i = 0; i < DATA_ROW_COUNT; i++) {
        const r = ws.getRow(ROW_DATA_START + i);
        dutyCounts.push(cellText(r.getCell(1)));
        personNames.push(cellText(r.getCell(2)));
    }

    // 最後10日分のセルデータを収集
    const importedDayData = new Map(); // dayNum → { isRestDay, noDuty?, noDutyNoon?, noDutyNight?, rowValues }
    for (const dayNum of last10Days) {
        const cols    = colsByDay.get(dayNum);
        const noonCol = cols.find(c => c.noonNight === '昼');
        const nightCol = cols.find(c => c.noonNight === '夜');
        const isRestDay = !!(noonCol && nightCol);

        if (isRestDay) {
            const noDutyNoon  = cellText(ws.getRow(ROW_NO_DUTY).getCell(noonCol.colIdx)).trim()  === '✓';
            const noDutyNight = cellText(ws.getRow(ROW_NO_DUTY).getCell(nightCol.colIdx)).trim() === '✓';
            const rowValues = [];
            for (let i = 0; i < DATA_ROW_COUNT; i++) {
                const r = ws.getRow(ROW_DATA_START + i);
                rowValues.push({
                    noon:  cellText(r.getCell(noonCol.colIdx)),
                    night: cellText(r.getCell(nightCol.colIdx))
                });
            }
            importedDayData.set(dayNum, { isRestDay: true, noDutyNoon, noDutyNight, rowValues });
        } else {
            const col    = cols[0];
            const noDuty = cellText(ws.getRow(ROW_NO_DUTY).getCell(col.colIdx)).trim() === '✓';
            const rowValues = [];
            for (let i = 0; i < DATA_ROW_COUNT; i++) {
                const r = ws.getRow(ROW_DATA_START + i);
                rowValues.push({ value: cellText(r.getCell(col.colIdx)) });
            }
            importedDayData.set(dayNum, { isRestDay: false, noDuty, rowValues });
        }
    }

    // 年月セレクタを「読み込んだ月の翌月」に設定してテーブルを再構築
    const dispMonth = loadedMonth === 12 ? 1 : loadedMonth + 1;
    const dispYear  = loadedMonth === 12 ? loadedYear + 1 : loadedYear;
    document.getElementById('select-year').value  = String(dispYear);
    document.getElementById('select-month').value = String(dispMonth);
    await updateTableStructure();
    // Tabulatorの内部レンダリングが完了するのを待つ
    await new Promise(r => setTimeout(r, 100));

    // テーブルの行を取得して先月データを書き込む
    const allRows   = table.getRows();
    const noDutyRow = allRows.find(r => r.getData().id === 'row_no_duty');
    const dataRows  = allRows.filter(r => typeof r.getData().id === 'number');


    // 各 updateObj を構築（人名 + 先月日付フィールド）
    const noDutyUpdateObj = {};
    const rowUpdateObjs   = dataRows.map((_, i) => ({
        name: personNames[i] ?? '',
        duty_count: dutyCounts[i] ?? ''
    }));

    for (const [dayNum, dayData] of importedDayData) {
        const base = `prev_day${dayNum}`;
        if (dayData.isRestDay) {
            noDutyUpdateObj[`${base}_noon`]  = dayData.noDutyNoon;
            noDutyUpdateObj[`${base}_night`] = dayData.noDutyNight;
            dayData.rowValues.forEach((v, i) => {
                if (rowUpdateObjs[i]) {
                    rowUpdateObjs[i][`${base}_noon`]  = v.noon;
                    rowUpdateObjs[i][`${base}_night`] = v.night;
                }
            });
        } else {
            noDutyUpdateObj[base] = dayData.noDuty;
            dayData.rowValues.forEach((v, i) => {
                if (rowUpdateObjs[i]) rowUpdateObjs[i][base] = v.value;
            });
        }
    }

    if (noDutyRow) noDutyRow.update(noDutyUpdateObj);
    dataRows.forEach((row, i) => { if (rowUpdateObjs[i]) row.update(rowUpdateObjs[i]); });
    table.redraw(true);

    } catch (err) {
        console.error('[load] エラー:', err);
        await window.api.showMessageBox({
            type: 'error',
            title: '読み込みエラー',
            message: `ファイルの読み込みに失敗しました。\n\n${err.message}`
        });
    }
}

const loadPrevButton = document.getElementById('load-prev-month-button');
if (loadPrevButton) {
    loadPrevButton.addEventListener('click', loadPrevMonthData);
}

async function loadKibouSheet() {
    try {
        const filePath = await window.api.openFileDialog();
        if (!filePath) return;

        const base64 = await window.api.readFileBase64(filePath);
        const binary = atob(base64);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(bytes.buffer);
        const ws = workbook.worksheets[0];
        if (!ws) {
            await window.api.showMessageBox({ type: 'error', title: 'エラー', message: 'ワークシートが見つかりません' });
            return;
        }

        // 1行目をヘッダーとして列を再構築、2行目以降をデータとして読み込む
        // 背景色を維持するため、既存列の cssClass を位置順で引き継ぐ
        const colCount = ws.columnCount;
        const existingCssClasses = table.getColumns().map(col => col.getDefinition().cssClass || '');
        const newColumns = [];
        for (let c = 1; c <= colCount; c++) {
            const title = cellText(ws.getRow(1).getCell(c));
            newColumns.push({
                title: title,
                field: `col${c}`,
                headerSort: false,
                editor: 'input',
                hozAlign: 'center',
                width: Math.max(40, title.length * 12),
                cssClass: existingCssClasses[c - 1] || '',
            });
        }

        const newData = [];
        for (let r = 2; r <= ws.rowCount; r++) {
            const rowObj = {};
            for (let c = 1; c <= colCount; c++) {
                rowObj[`col${c}`] = cellText(ws.getRow(r).getCell(c));
            }
            newData.push(rowObj);
        }

        table.setColumns(newColumns);
        await table.setData(newData);

    } catch (err) {
        await window.api.showMessageBox({ type: 'error', title: '読み込みエラー', message: err.message });
    }
}

const kibouButton = document.getElementById('load-kibou-button');
if (kibouButton) kibouButton.addEventListener('click', loadKibouSheet);

async function runDutyAssignment() {
    try {
        const columns = table.getColumns();
        const fields  = columns.map(c => c.getField());
        const allRows = table.getRows();
        if (columns.length === 0 || allRows.length === 0) {
            await window.api.showMessageBox({ type: 'warning', title: '注意', message: 'データがありません。' });
            return;
        }

        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        const normalizeVal = v => {
            if (v === null || v === undefined || v === false) return '';
            if (v === true) return '〇';
            const s = String(v);
            return s === '○' ? '〇' : s;
        };

        const isKibouMode = fields[0]?.startsWith('col');

        if (isKibouMode) {
            // ── キボウモード ──────────────────────────────────────────
            // アプリの exportToExcel() で出力した Excel に希望を書き込んだものを読み込んだ状態。
            // 列タイトル = 日付表示文字列 ("12/22" / "1" 等)、データ行 = 各種ヘッダー行 + 人名行。
            // これらから Python が期待する入力形式（マーカー行付き）を再構成する。
            const nCols = fields.length;
            const HEADER_NAMES = new Set(['休業日', '当直不要', '曜日', '祝日', '昼夜']);

            // 列の並び順に依存しないよう、"当直不要"/"曜日"/"昼夜" を含む列を氏名列と判定する
            const nameField = [fields[0], fields[1]].find(f =>
                allRows.some(r => HEADER_NAMES.has(r.getData()[f] ?? ''))
            ) ?? fields[0];
            const dutyField = [fields[0], fields[1]].find(f => f !== nameField) ?? fields[1];

            // キー行を nameField の値で特定
            const noDutyData  = allRows.find(r => r.getData()[nameField] === '当直不要')?.getData() ?? {};
            const weekdayData = allRows.find(r => r.getData()[nameField] === '曜日')?.getData()    ?? {};
            const shiftData   = allRows.find(r => r.getData()[nameField] === '昼夜')?.getData()    ?? {};
            const personRows  = allRows
                .filter(r => !HEADER_NAMES.has(r.getData()[nameField] ?? ''))
                .map(r => r.getData());

            // 列タイトルから past / start / end 列インデックスを特定（0-based）
            // 前月日付: "12/22" のように "/" を含む → past 列
            // 当月日付: "1", "2", ..., "31" のように数字のみ → start/end 列
            let pastColIdx = -1, startColIdx = -1, endColIdx = -1;
            for (let i = 2; i < nCols; i++) {
                const title = columns[i].getDefinition().title.trim();
                if (title.includes('/') && pastColIdx === -1) pastColIdx = i;
                if (/^\d+$/.test(title)) {
                    if (startColIdx === -1) startColIdx = i;
                    endColIdx = i;
                }
            }

            if (pastColIdx === -1 || startColIdx === -1) {
                await window.api.showMessageBox({ type: 'error', title: 'エラー',
                    message: '列タイトルから日付列を認識できません。\n前月列（例: "12/22"）または当月列が見つかりません。' });
                return;
            }

            const mkRow = () => Array(nCols).fill('');

            // Row 0: 列マーカー (past / start / end)
            const r0 = mkRow();
            r0[pastColIdx]  = 'past';
            r0[startColIdx] = 'start';
            r0[endColIdx]   = 'end';
            ws.addRow(r0);

            // Row 1: 応援医師（当直不要チェック）
            const r1 = mkRow(); r1[1] = '応援医師';
            for (let i = 2; i < nCols; i++) {
                const v = noDutyData[fields[i]];
                if (v === '○' || v === '〇' || v === true) r1[i] = '〇';
            }
            ws.addRow(r1);

            // Row 2: 日付数字（列タイトルから抽出）
            const r2 = mkRow(); r2[1] = '日';
            for (let i = 2; i < nCols; i++) {
                const t = columns[i].getDefinition().title.trim();
                if (t.includes('/')) r2[i] = parseInt(t.split('/')[1]);
                else if (/^\d+$/.test(t)) r2[i] = parseInt(t);
            }
            ws.addRow(r2);

            // Row 3: 曜日
            const r3 = mkRow(); r3[1] = '曜日';
            for (let i = 2; i < nCols; i++) r3[i] = weekdayData[fields[i]] || '';
            ws.addRow(r3);

            // Row 4: 昼夜 + "start" 行マーカー（Python: start_row = 5）
            const r4 = mkRow(); r4[1] = 'start';
            for (let i = 2; i < nCols; i++) r4[i] = shiftData[fields[i]] || '';
            ws.addRow(r4);

            // Rows 5+: 人名・希望データ
            // Python は COL_REQUIRED_SHIFTS=0(A列), COL_NAMES=1(B列) を期待するため
            // 表示列順に依存せず dutyField/nameField から正しい位置に書く
            for (const data of personRows) {
                const er = mkRow();
                er[0] = data[dutyField] || '';  // Python COL_REQUIRED_SHIFTS (A列)
                er[1] = data[nameField] || '';  // Python COL_NAMES (B列)
                for (let i = 2; i < nCols; i++) er[i] = normalizeVal(data[fields[i]]);
                ws.addRow(er);
            }

            // 終端行: "end" マーカー
            const re = mkRow(); re[1] = 'end';
            ws.addRow(re);
        } else {
            // ── 当直表モード ──────────────────────────────────────────
            // prev_day*/day* 列構造から Python 入力 Excel を組み立てる。
            const dateFields = fields.slice(2);
            let pastColIdx = -1, startColIdx = -1, endColIdx = -1;
            dateFields.forEach((f, i) => {
                if (/^prev_day\d+/.test(f) && pastColIdx === -1) pastColIdx = i;
                if (/^day\d+/.test(f)) { if (startColIdx === -1) startColIdx = i; endColIdx = i; }
            });
            if (pastColIdx === -1 || startColIdx === -1) {
                await window.api.showMessageBox({ type: 'error', title: 'エラー',
                    message: '列の構造が認識できません。「表を更新」後に再度お試しください。' });
                return;
            }

            const noDutyData = allRows.find(r => r.getData().id === 'row_no_duty')?.getData() ?? {};
            const dayData    = allRows.find(r => r.getData().id === 'header_day')?.getData()    ?? {};
            const n = fields.length;
            const O = 2;

            const getDateNum   = f => { const m = f.match(/\d+/); return m ? parseInt(m[0]) : 0; };
            const getShiftType = f => f.endsWith('_noon') ? '昼' : '夜';

            // Row 0: past / start / end 列マーカー
            const r0 = Array(n).fill('');
            r0[O + pastColIdx] = 'past'; r0[O + startColIdx] = 'start'; r0[O + endColIdx] = 'end';
            ws.addRow(r0);

            // Row 1: 応援医師（当直不要）
            const r1 = Array(n).fill(''); r1[1] = '応援医師';
            dateFields.forEach((f, i) => { if (noDutyData[f] === true) r1[O + i] = '〇'; });
            ws.addRow(r1);

            // Row 2: 日付数字
            const r2 = Array(n).fill(''); r2[1] = '日';
            dateFields.forEach((f, i) => { r2[O + i] = getDateNum(f); });
            ws.addRow(r2);

            // Row 3: 曜日
            const r3 = Array(n).fill(''); r3[1] = '曜日';
            dateFields.forEach((f, i) => { r3[O + i] = dayData[f] || ''; });
            ws.addRow(r3);

            // Row 4: 昼夜 + "start" マーカー（names[4]="start" → start_row=5）
            const r4 = Array(n).fill(''); r4[1] = 'start';
            dateFields.forEach((f, i) => { r4[O + i] = getShiftType(f); });
            ws.addRow(r4);

            // Rows 5–24: 人名・希望
            for (let pi = 0; pi < 20; pi++) {
                const data = allRows.find(r => r.getData().id === pi)?.getData() ?? {};
                const er = Array(n).fill('');
                er[0] = data.duty_count ?? '';
                er[1] = data.name ?? '';
                dateFields.forEach((f, i) => { er[O + i] = normalizeVal(data[f]); });
                ws.addRow(er);
            }

            // Row 25: "end" マーカー
            const re = Array(n).fill(''); re[1] = 'end';
            ws.addRow(re);
        }

        // バッファ → base64 → 一時ファイル
        const buffer = await wb.xlsx.writeBuffer();
        const bytes  = new Uint8Array(buffer);
        let binary = '';
        for (let i = 0; i < bytes.length; i += 8192)
            binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 8192));
        const tempPath = await window.api.writeTempFile(btoa(binary));
        if (!tempPath) {
            await window.api.showMessageBox({ type: 'error', title: 'エラー', message: '一時ファイルの作成に失敗しました。' });
            return;
        }

        // Python 実行
        const result = await window.api.runPythonScript(tempPath);
        if (!result.success) {
            await window.api.showMessageBox({ type: 'error', title: 'Python エラー', message: result.message });
            return;
        }

        // 結果ウィンドウを開く
        const pathMatch = result.message.match(/'([^']+\.xlsx)'/);
        if (pathMatch) {
            await window.api.openResultWindow(pathMatch[1]);
        } else {
            const debugInfo = `\n\n【デバッグ用】\n入力Excel: ${tempPath}\nログ: %USERPROFILE%\\Documents\\DutyAssignmentLogs\\duty_assign.log`;
            await window.api.showMessageBox({
                type: 'warning', title: '解なし',
                message: result.message + debugInfo
            });
        }

    } catch (err) {
        await window.api.showMessageBox({ type: 'error', title: 'エラー', message: err.message });
    }
}

const runDutyButton = document.getElementById('run-duty-button');
if (runDutyButton) runDutyButton.addEventListener('click', runDutyAssignment);

