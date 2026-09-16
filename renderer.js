const tableContainer = document.getElementById('table-container');
let currentDutyCountField = 'duty_count'; // kibouモード時は 'col1' or 'col2' に切り替わる
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
    height: false,
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
    editTriggerEvent: "dblclick",
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
        if (cell.getField() === currentDutyCountField) updateProvisionalDutyCountDisplay();
    },

    clipboard: false,
    // clipboardPasteAction: "replace",
    // clipboardPasteParser: "table",

    // --- 右クリックメニュー（コンテキストメニュー）の設定 ---
    rowFormatter: function(row) {
        if (typeof row.getData().id !== 'string') return;
        // id が文字列の行（休業日・当直不要・曜日・祝日・昼夜）はヘッダー扱い
        row.getElement().classList.add('is-header-row');
        const cells = row.getCells();
        if (cells[0]) cells[0].getElement().style.pointerEvents = 'none';
        if (cells[1]) cells[1].getElement().style.pointerEvents = 'none';
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
let dutyCountEnterPressed = false; // Enter確定で下セルへ移動するためのフラグ

function showLoading() {
    const el = document.getElementById('loading-overlay');
    if (el) el.style.display = 'flex';
}
function hideLoading() {
    const el = document.getElementById('loading-overlay');
    if (el) el.style.display = 'none';
}

// クリックしたセルをペーストの始点として記録
table.on("cellClick", function(e, cell) {
    targetPasteCell = cell;
});

// --- カスタムペースト処理（Excelのような部分ペースト） ---
document.addEventListener("paste", async function(e) {
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
    let allRows = table.getRows("active");
    const allColumns = table.getColumns();

    const startRowIndex = allRows.findIndex(r => r === startRow);
    const startColIndex = allColumns.findIndex(c => c === startColumn);

    if (startRowIndex === -1 || startColIndex === -1) return;

    // 名前列（1列目）へのペーストで行数が足りない場合、自動で行を追加
    const PASTE_HIDS = new Set(['row_holiday_checkbox', 'row_no_duty', 'header_day', 'header_holiday', 'header_noon_night']);
    const isNameCol = startColIndex === 0;
    if (isNameCol) {
        const personRows = allRows.filter(r => {
            const id = r.getData().id;
            return !(typeof id === 'string' && (id.startsWith('header_') || PASTE_HIDS.has(id)));
        });
        const startPersonIdx = personRows.findIndex(r => r === startRow);
        if (startPersonIdx !== -1) {
            const extraNeeded = startPersonIdx + dataMatrix.length - personRows.length;
            if (extraNeeded > 0) {
                const existingIds = personRows.map(r => r.getData().id).filter(id => typeof id === 'number');
                const maxId = existingIds.length > 0 ? Math.max(...existingIds) : personRows.length - 1;
                for (let j = 0; j < extraNeeded; j++) {
                    await table.addRow({ id: maxId + 1 + j, name: "", duty_count: "" }, false);
                }
                allRows = table.getRows("active");
            }
        }
    }

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
        updateProvisionalDutyCountDisplay();
        if (isNameCol) {
            autoSizeNameColumn('name');
            // 末尾に常に空白行を1行確保
            const personRowsNow = table.getRows("active").filter(r => {
                const id = r.getData().id;
                return !(typeof id === 'string' && (id.startsWith('header_') || PASTE_HIDS.has(id)));
            });
            const lastPerson = personRowsNow[personRowsNow.length - 1];
            if (lastPerson && lastPerson.getData().name) {
                const eIds = personRowsNow.map(r => r.getData().id).filter(id => typeof id === 'number');
                const nId = eIds.length > 0 ? Math.max(...eIds) : personRowsNow.length - 1;
                await table.addRow({ id: nId + 1, name: "", duty_count: "" }, false);
            }
        }
    }
});
// ------------------------------

// Pythonスクリプトを実行し、結果を通知する関数
async function executePythonScript(filePath) {
    if (!filePath) {
        return;
    }
    // メインプロセスにPythonスクリプトの実行を依頼し、結果を受け取る
    const result = await window.api.runPythonScript(filePath);
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

// --- 1. 日付範囲入力の初期化 ---
// 日付を <input type="date"> 用の "YYYY-MM-DD" 文字列に変換
function formatDateInput(date) {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
}
// "YYYY-MM-DD" 文字列をローカル日付として Date に変換（new Date(str) のUTC解釈を避ける）
function parseDateInput(str) {
    const [y, m, d] = str.split('-').map(Number);
    return new Date(y, m - 1, d);
}
// フィールド名に埋め込むための "YYYYMMDD" 文字列
function formatYYYYMMDD(date) {
    return formatDateInput(date).replace(/-/g, '');
}
// フィールド名（prev_day/day + YYYYMMDD + 任意の_noon/_night）から実日付を復元
function dateFromField(field) {
    const m = field && field.match(/^(?:prev_day|day)(\d{8})/);
    if (!m) return null;
    const s = m[1];
    return new Date(parseInt(s.slice(0, 4)), parseInt(s.slice(4, 6)) - 1, parseInt(s.slice(6, 8)));
}

function initDateRangeInputs() {
    const startInput = document.getElementById('select-start-date');
    const endInput = document.getElementById('select-end-date');
    const today = new Date();
    const oneMonthLater = new Date(today.getFullYear(), today.getMonth() + 1, today.getDate());
    startInput.value = formatDateInput(today);
    endInput.value = formatDateInput(oneMonthLater);
}

// --- 2. 選択された年月から表のカレンダーを構築する ---
let forcedHolidays = new Set(); // 強制休日設定を保持するセット (YYYY-MM-DD形式)

// 「昼勤務の翌日も夜勤務可能」の特別条件を適用する人（氏名で管理。デフォルトは尾崎泰）
let specialRuleNames = new Set(['尾崎泰']);

async function updateTableStructure() {
    customHistory.clear();
    currentDutyCountField = 'duty_count';

    const startInput = document.getElementById('select-start-date');
    const endInput = document.getElementById('select-end-date');

    if (!startInput || !endInput || !startInput.value || !endInput.value) return;

    const startDate = parseDateInput(startInput.value);
    let endDate = parseDateInput(endInput.value);
    if (endDate < startDate) {
        endDate = new Date(startDate);
        endInput.value = startInput.value;
    }

    // 表示する日付リストを作成（開始日の前10日分＝参考表示のみ ＋ 開始日〜終了日＝当直決めの対象）
    const displayDays = [];
    // 開始日の前10日（当直回数の計算対象外）
    for (let i = 10; i >= 1; i--) {
        const d = new Date(startDate);
        d.setDate(d.getDate() - i);
        displayDays.push({ dateObj: d, fieldPrefix: `prev_day${formatYYYYMMDD(d)}` });
    }
    // 開始日〜終了日（当直決めの対象）
    for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
        displayDays.push({ dateObj: new Date(d), fieldPrefix: `day${formatYYYYMMDD(d)}` });
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
        const { dateObj, fieldPrefix } = dayInfo;
        const y = dateObj.getFullYear(), m = dateObj.getMonth() + 1, d = dateObj.getDate();
        const dayNum = dateObj.getDay();
        const dayStr = dayOfWeek[dayNum];

        const holidayName = await window.api.getHolidayName(dateObj);
        const isNaturalRestDay = (dayNum === 0 || dayNum === 6 || holidayName);

        // 強制休日用のキー（YYYY-MM-DD）を作成
        const dateKey = `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
        const isRestDay = (isNaturalRestDay || forcedHolidays.has(dateKey));

        // 日付の表示内容。前10日分（参考表示）は「m/<br>d」形式、対象範囲は「d」のみ
        // （Excel再読込時にこの表記の違いで両者を区別しているため維持している）
        const isLookback = fieldPrefix.startsWith('prev_day');
        const dateDisplay = isLookback ? `${m}/<br>${d}` : d;

        // --- 共通設定：カラム構成 ---
        const getCellConfig = (field, cClass) => ({
            title: `<div style="text-align:center;white-space:normal;line-height:1.1;">${dateDisplay}</div>`,
            field: field,
            width: 23, // 横幅を40に完全固定して自動拡張を防ぐ
            hozAlign: "center",
            headerSort: false,
            cssClass: cClass,
            editor: "list",
            editorParams: {
                values: ['', '〇', '×', '輪番'],
                autocomplete: false,
                clearable: false,
                listOnEmpty: true,
                itemFormatter: (label, value, item, element) => {
                    const STYLE = {
                        '':   { bg:'#f5f5f5', color:'#888', text:'空白' },
                        '〇':  { bg:'#e6f4ea', color:'#2a7d4f', text:'〇' },
                        '×':  { bg:'#fdecea', color:'#c0392b', text:'×' },
                        '輪番': { bg:'#e8eaf6', color:'#3949ab', text:'輪番' },
                    };
                    const s = STYLE[value] ?? { bg:'#fff', color:'#333', text: String(label) };
                    if (element) element.style.background = s.bg;
                    return `<span style="color:${s.color};font-weight:600;font-size:13px;">${s.text}</span>`;
                },
            },
            editTriggerEvent: "dblclick",
            // ヘッダー情報行は編集不可にする
            editable: (cell) => {
                const rowData = cell.getRow().getData();
                return !(typeof rowData.id === 'string' && (rowData.id.startsWith("header_") || rowData.id === "row_no_duty" || rowData.id === "row_holiday_checkbox"));
            },
            // --- セルの表示形式をカスタマイズ ---
            formatter: (cell) => {
                const rowData = cell.getRow().getData();
                const field = cell.getColumn().getField();
                const val = cell.getValue();

                // 当直不要行はチェックボックスで表示
                if (rowData.id === "row_no_duty") {
                    return `<div style="display:flex;justify-content:center;align-items:center;height:100%;"><input type="checkbox" ${val === true ? "checked" : ""} style="cursor:pointer; pointer-events:none;"></div>`;
                }
                // 休業日なら行は平日のみチェックボックスで表示（土日祝は null なので空セル）
                if (rowData.id === "row_holiday_checkbox") {
                    if (val === null || val === undefined) return "";
                    return `<div style="display:flex;justify-content:center;align-items:center;height:100%;"><input type="checkbox" ${val === true ? "checked" : ""} style="cursor:pointer; pointer-events:none;"></div>`;
                }

                // 当直不要が true の列を灰色（休業日・当直不要行自身は除く）
                // （CSSの !important を上書きするため setProperty を使用）
                const GRAY_SKIP = new Set(['row_holiday_checkbox', 'row_no_duty']);
                if (!GRAY_SKIP.has(rowData.id)) {
                    const noDutyRow = table.getRows().find(r => r.getData().id === 'row_no_duty');
                    const el = cell.getElement();
                    if (noDutyRow && noDutyRow.getData()[field] === true) {
                        const pos = cell.getRow().getPosition();
                        const gray = (pos % 2 === 1) ? '#d0d0d0' : '#bcbcbc';
                        el.style.setProperty('background-color', gray, 'important');
                        el.style.setProperty('color', '#888', 'important');
                    } else {
                        el.style.removeProperty('background-color');
                        el.style.removeProperty('color');
                    }
                }

                return val != null ? val : "";
            },
            cellClick: (e, cell) => {
                const rowId = cell.getRow().getData().id;
                if (rowId === "row_no_duty") {
                    const newVal = !cell.getValue();
                    cell.setValue(newVal);
                    updateDutyCountDisplay();
                    // 同列の全行（休業日・当直不要行自身を除く）を直接DOM操作で即時グレー化
                    const f = cell.getColumn().getField();
                    const GRAY_SKIP_IDS = new Set(['row_holiday_checkbox', 'row_no_duty']);
                    table.getRows().forEach(row => {
                        if (GRAY_SKIP_IDS.has(row.getData().id)) return;
                        const c = row.getCell(f);
                        if (!c) return;
                        const el = c.getElement();
                        if (!el) return;
                        if (newVal) {
                            const pos = row.getPosition();
                            const gray = (pos % 2 === 1) ? '#d0d0d0' : '#bcbcbc';
                            el.style.setProperty('background-color', gray, 'important');
                            el.style.setProperty('color', '#888', 'important');
                        } else {
                            el.style.removeProperty('background-color');
                            el.style.removeProperty('color');
                        }
                    });
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
        autoSizeNameColumn('name');
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

function autoSizeNameColumn(field = 'name') {
    const col = table.getColumn(field);
    if (!col) return;
    let maxLen = 0;
    for (const row of table.getRows()) {
        const val = String(row.getData()[field] ?? '');
        if (val.length > maxLen) maxLen = val.length;
    }
    col.updateDefinition({ width: Math.max(70, maxLen * 14 + 10) });
}

function updateProvisionalDutyCountDisplay() {
    const el = document.getElementById('provisional-duty-count-display');
    if (!el) return;
    const skipIds = new Set(["row_no_duty", "row_holiday_checkbox", "header_day", "header_holiday", "header_noon_night"]);
    let total = 0;
    for (const row of table.getRows()) {
        const data = row.getData();
        if (typeof data.id === 'string' && (data.id.startsWith("header_") || skipIds.has(data.id))) continue;
        const val = parseInt(data[currentDutyCountField], 10);
        if (!isNaN(val)) total += val;
    }
    el.textContent = `仮当直回数の合計: ${total}`;
    // 今月の当直回数と比較して色を切り替える
    const dutyEl = document.getElementById('duty-count-display');
    if (dutyEl) {
        const m = dutyEl.textContent.match(/(\d+)/);
        const net = m ? parseInt(m[1], 10) : null;
        el.style.color = (net !== null && total !== net) ? 'red' : '#333';
    }
    updateSpecialRuleButtonLabel();
}

// --- 「昼勤務の翌日も夜勤務可能」特別条件のプルダウン ---
function getCurrentPersonNames() {
    return table.getRows()
        .filter(r => typeof r.getData().id === 'number')
        .map(r => (r.getData().name || '').trim())
        .filter(name => name !== '');
}

function escapeHtml(s) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function updateSpecialRuleButtonLabel() {
    const btn = document.getElementById('special-rule-button');
    if (!btn) return;
    const checked = [...new Set(getCurrentPersonNames())].filter(n => specialRuleNames.has(n));
    const names = checked.length > 0 ? checked.join('、') : 'なし';
    btn.textContent = `昼夜連続勤務可：${names}`;
    btn.title = names;
}

function renderSpecialRulePanel() {
    const list = document.getElementById('special-rule-list');
    if (!list) return;
    const names = [...new Set(getCurrentPersonNames())];
    if (names.length === 0) {
        list.innerHTML = '<div style="font-size:12px; color:#999;">名前が入力されていません</div>';
        return;
    }
    list.innerHTML = names.map(name => {
        const checked = specialRuleNames.has(name) ? 'checked' : '';
        const escaped = escapeHtml(name);
        return `<label style="display:flex; align-items:center; gap:6px; padding:3px 2px; cursor:pointer; white-space:nowrap;">
            <input type="checkbox" class="special-rule-checkbox" data-name="${escaped}" ${checked}>
            <span>${escaped}</span>
        </label>`;
    }).join('');
}

(function initSpecialRuleDropdown() {
    const button = document.getElementById('special-rule-button');
    const panel = document.getElementById('special-rule-panel');
    if (!button || !panel) return;

    button.addEventListener('click', (e) => {
        e.stopPropagation();
        if (panel.style.display === 'block') {
            panel.style.display = 'none';
        } else {
            renderSpecialRulePanel();
            panel.style.display = 'block';
        }
    });
    panel.addEventListener('click', (e) => e.stopPropagation());
    panel.addEventListener('change', (e) => {
        const cb = e.target.closest('.special-rule-checkbox');
        if (!cb) return;
        if (cb.checked) specialRuleNames.add(cb.dataset.name);
        else specialRuleNames.delete(cb.dataset.name);
        updateSpecialRuleButtonLabel();
    });
    document.addEventListener('click', () => { panel.style.display = 'none'; });

    updateSpecialRuleButtonLabel();
})();

// --- 3. 実行指示 ---
initDateRangeInputs();

// tableholderの横スクロールをviewport下部の擬似スクロールバーと同期する
function setupFakeHScrollbar() {
    const tableholder = document.querySelector('.tabulator-tableholder');
    const proxy = document.getElementById('hscroll-proxy');
    const inner = document.getElementById('hscroll-proxy-inner');
    const container = document.getElementById('table-container');
    if (!tableholder || !proxy || !inner) return;

    function syncWidth() {
        inner.style.width = tableholder.scrollWidth + 'px';
        if (container) {
            const tbl = tableholder.querySelector('.tabulator-table');
            const tableW = tbl ? tbl.offsetWidth : 0;
            // テーブルがコンテナより狭いときはコンテナをテーブル幅に縮める
            if (tableW > 0 && tableW <= tableholder.clientWidth) {
                container.style.width = tableW + 'px';
            } else {
                container.style.width = '';
            }
        }
    }
    syncWidth();

    let fromProxy = false;
    let fromTable = false;
    proxy.addEventListener('scroll', () => {
        if (fromTable) return;
        fromProxy = true;
        tableholder.scrollLeft = proxy.scrollLeft;
        fromProxy = false;
    }, { passive: true });
    tableholder.addEventListener('scroll', () => {
        if (fromProxy) return;
        fromTable = true;
        proxy.scrollLeft = tableholder.scrollLeft;
        fromTable = false;
    }, { passive: true });

    new ResizeObserver(syncWidth).observe(tableholder);

    // ウィンドウリサイズ時にコンテナ幅をリセットして再計算
    window.addEventListener('resize', () => {
        if (container) container.style.width = '';
        syncWidth();
    });
}

// 2. Tabulatorのセットアップが完了したら、初期描画を行う
table.on("tableBuilt", function(){
    updateTableStructure();
    setupFakeHScrollbar();
});

// 仮当直回数列：現在編集中のセルを追跡（クリック・プログラム両対応）
let editingDutyCountRow = null;
table.on("cellEditing", function(cell) {
    editingDutyCountRow = (cell.getField() === currentDutyCountField) ? cell.getRow() : null;
});
table.on("cellEditCancelled", function(cell) {
    if (cell.getField() === currentDutyCountField) editingDutyCountRow = null;
});

// Tabulatorより先にEnterを検知（キャプチャフェーズ）
document.addEventListener('keydown', function(e) {
    if (e.key === 'Enter' && editingDutyCountRow) dutyCountEnterPressed = true;
}, true);

// 仮当直回数列はシングルクリックで編集開始、エディタ全選択
table.on("cellClick", function(e, cell) {
    if (cell.getField() !== currentDutyCountField) return;
    if (typeof cell.getRow().getData().id !== 'number') return;
    setTimeout(() => {
        cell.edit(true);
        setTimeout(() => {
            const input = cell.getElement().querySelector('input');
            if (input) { input.focus(); input.select(); }
        }, 10);
    }, 0);
});

table.on("cellEdited", function(cell){
    if (cell.getField() === currentDutyCountField) {
        // 全角数字→半角に自動変換
        const raw = String(cell.getValue() ?? '');
        const converted = raw.replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
        if (converted !== raw) cell.setValue(converted, false);
        updateProvisionalDutyCountDisplay();

        // Enterで確定した場合、すぐ下の人名行に移動して編集開始
        if (dutyCountEnterPressed) {
            dutyCountEnterPressed = false;
            const currentRowId = cell.getRow().getData().id;
            const allRows = table.getRows();
            let foundCurrent = false;
            for (const row of allRows) {
                const id = row.getData().id;
                if (!foundCurrent) {
                    if (id === currentRowId) foundCurrent = true;
                    continue;
                }
                if (typeof id === 'number') {
                    const nextCell = row.getCell(currentDutyCountField);
                    if (nextCell) {
                        setTimeout(() => {
                            nextCell.edit(true);
                            setTimeout(() => {
                                const inp = nextCell.getElement().querySelector('input');
                                if (inp) { inp.focus(); inp.select(); }
                            }, 10);
                        }, 0);
                    }
                    break;
                }
            }
        }
    }
});

// 固定列（名前・仮当直回数）の最終行に手打ちしたら自動で1行追加
table.on("cellEdited", async function(cell) {
    // 名前欄はどの行を編集しても列幅を再計算する（最終行以外の編集で幅が更新されず省略されるのを防ぐ）
    if (cell.getField() === 'name') { autoSizeNameColumn('name'); updateSpecialRuleButtonLabel(); }
    if (!cell.getColumn().getDefinition().frozen || !cell.getValue()) return;
    const HIDS = new Set(['row_holiday_checkbox', 'row_no_duty', 'header_day', 'header_holiday', 'header_noon_night']);
    const personRows = table.getRows().filter(r => {
        const id = r.getData().id;
        return !(typeof id === 'string' && (id.startsWith('header_') || HIDS.has(id)));
    });
    if (personRows.length === 0) return;
    const lastRow = personRows[personRows.length - 1];
    if (cell.getRow().getIndex() !== lastRow.getIndex()) return;
    const existingIds = personRows.map(r => r.getData().id).filter(id => typeof id === 'number');
    const maxId = existingIds.length > 0 ? Math.max(...existingIds) : personRows.length - 1;
    await table.addRow({ id: maxId + 1, name: "", duty_count: "" }, false);
    autoSizeNameColumn('name');
});

table.on("cellEditing", function(cell) {
    if (cell.getField() !== currentDutyCountField) return;
    const input = cell.getElement().querySelector('input');
    if (!input) return;
    input.addEventListener('input', () => {
        const el = document.getElementById('provisional-duty-count-display');
        if (!el) return;
        const skipIds = new Set(["row_no_duty", "row_holiday_checkbox", "header_day", "header_holiday", "header_noon_night"]);
        let total = 0;
        for (const row of table.getRows()) {
            const data = row.getData();
            if (typeof data.id === 'string' && (data.id.startsWith("header_") || skipIds.has(data.id))) continue;
            const val = row === cell.getRow()
                ? parseInt(input.value, 10)
                : parseInt(data[currentDutyCountField], 10);
            if (!isNaN(val)) total += val;
        }
        el.textContent = `仮当直回数の合計: ${total}`;
        // 今月の当直回数と比較して色を切り替える
        const dutyEl = document.getElementById('duty-count-display');
        if (dutyEl) {
            const m = dutyEl.textContent.match(/(\d+)/);
            const net = m ? parseInt(m[1], 10) : null;
            el.style.color = (net !== null && total !== net) ? 'red' : '#333';
        }
    });
});



// 日付範囲変更時の自動更新
const startSel = document.getElementById('select-start-date');
const endSel = document.getElementById('select-end-date');
if (startSel && endSel) {
    const HEADER_IDS = new Set(['row_no_duty', 'row_holiday_checkbox', 'header_day', 'header_holiday', 'header_noon_night']);
    const isPersonRow = (r) => {
        const id = r.getData().id;
        return !(typeof id === 'string' && (id.startsWith('header_') || HEADER_IDS.has(id)));
    };

    const onDateChange = async () => {
        // 年月変更前の名前・当直回数を保存
        const saved = table.getRows()
            .filter(isPersonRow)
            .map(r => ({ name: r.getData().name ?? '', duty_count: r.getData().duty_count ?? '' }));

        forcedHolidays.clear();
        await updateTableStructure();

        // 名前・当直回数を復元
        const personRows = table.getRows().filter(isPersonRow);
        for (let i = 0; i < Math.min(saved.length, personRows.length); i++) {
            if (saved[i].name !== '' || saved[i].duty_count !== '') {
                personRows[i].update({ name: saved[i].name, duty_count: saved[i].duty_count });
            }
        }
        updateProvisionalDutyCountDisplay();
    };
    startSel.addEventListener('change', onDateChange);
    endSel.addEventListener('change', onDateChange);
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

    // 当月列（day\d+ 形式）の列番号（1ベース）を収集
    const currentMonthColNums = [];
    columns.forEach((col, j) => {
        const f = col.getField();
        if (f && /^day\d+/.test(f)) currentMonthColNums.push(j + 1);
    });

    // 1行目：列タイトル（日付など）
    // 列ヘッダー要素には sat/sun 系 CSS が当たらないため、最初のデータ行セルで代用
    const firstDataRow = rows.length > 0 ? rows[0] : null;
    const titleExcelRow = worksheet.addRow(columns.map(col => stripHtml(col.getDefinition().title)));
    columns.forEach((col, j) => {
        applyStyle(titleExcelRow.getCell(j + 1), firstDataRow?.getCell(col.getField())?.getElement());
    });
    titleExcelRow.eachCell({ includeEmpty: true }, cell => {
        cell.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
    });

    // 2行目以降：テーブルの全データ行（人名行のExcel行番号を記録）
    const personExcelRowNums = [];
    let excelRowNum = 2;
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
        excelRow.height = 22; // 通常の行の高さ(20)より10%高くする
        if (typeof data.id === 'number') personExcelRowNums.push(excelRowNum);
        excelRowNum++;
    }

    // 人名行 × 当月列のセルを編集可能にしてドロップダウンを設定（他は保護でロック）
    for (const rowNum of personExcelRowNums) {
        for (const colNum of currentMonthColNums) {
            const cell = worksheet.getCell(rowNum, colNum);
            cell.protection = { locked: false };
            cell.dataValidation = {
                type: 'list',
                allowBlank: true,
                formulae: ['"　,〇,×,輪番"'],
                showErrorMessage: true,
                errorStyle: 'stop',
                errorTitle: '入力エラー',
                error: '「〇」「×」「輪番」または空白のみ入力できます。'
            };
        }
    }

    // シートを保護（パスワードなし・セル選択は両方許可）
    await worksheet.protect('', {
        selectLockedCells: true,
        selectUnlockedCells: true
    });

    // ファイルとして保存
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const startStr = document.getElementById('select-start-date').value.replaceAll('-', '');
    const endStr = document.getElementById('select-end-date').value.replaceAll('-', '');
    a.download = `${startStr}-${endStr}_当直表.xlsx`;
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
    const filePath = await window.api.openFileDialog();
    if (!filePath) return;
    showLoading();
    // ExcelJS でファイルを読み込む（base64経由でArrayBufferに変換）
    const base64 = await window.api.readFileBase64(filePath);
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(bytes.buffer);
    const ws = workbook.worksheets[0];
    if (!ws) { return; }
    // ファイル名から読み込んだ年月を検出
    const filename = filePath.split(/[\\/]/).pop();
    const nameMatch = filename.match(/(\d{4})年(\d{1,2})月/);
    let loadedYear, loadedMonth;
    if (nameMatch) {
        loadedYear = parseInt(nameMatch[1]);
        loadedMonth = parseInt(nameMatch[2]);
    } else {
        // ファイル名から取れない場合はA1セルから検出（Python出力・保存済み形式）
        const a1 = cellText(ws.getRow(1).getCell(1));
        const a1Match = a1.match(/(\d{4})年(\d{1,2})月/);
        if (a1Match) {
            loadedYear = parseInt(a1Match[1]);
            loadedMonth = parseInt(a1Match[2]);
        } else {
            // A1にもなければ現在の開始日から前月を推定
            const curStart = parseDateInput(document.getElementById('select-start-date').value);
            loadedMonth = curStart.getMonth(); // 0始まりなので、これで「開始月-1」になる
            loadedYear  = curStart.getFullYear();
            if (loadedMonth === 0) { loadedMonth = 12; loadedYear--; }
        }
    }

    // 行ラベルを動的に探す（フォーマットが異なるExcelにも対応）
    const PREV_HEADER_LABELS = new Set([
        '', '　', '名前', '仮当直回数', '曜日', '祝日', '昼夜',
        '休業日', '当直不要', 'start', 'end', '応援医師', 'past', '日',
    ]);
    let ROW_HOLIDAY_CB = 2, ROW_NO_DUTY = 3, ROW_NOON_NIGHT = 5;
    const personRowIndices = [];
    for (let r = 1; r <= ws.rowCount; r++) {
        const label = cellText(ws.getRow(r).getCell(1)).trim();
        if      (label === '休業日')   ROW_HOLIDAY_CB = r;
        else if (label === '当直不要') ROW_NO_DUTY    = r;
        else if (label === '昼夜')     ROW_NOON_NIGHT = r;
        else if (label && !PREV_HEADER_LABELS.has(label) && !/^\d{4}年\d{1,2}月/.test(label)) personRowIndices.push(r);
    }
    // ROW_NOON_NIGHT がラベル検出で見つからなかった（Python形式はラベルが 'start'）場合、
    // 実際に '昼'/'夜' 値を持つ行を確認・修正する。
    // 保存済みファイルでは save handler が '祝日'行を row 5 に挿入するため
    // デフォルト 5 が '祝日'行を指してしまう問題を防ぐ。
    const maxCol = Math.max(ws.columnCount, ws.actualColumnCount || 0, 60);
    if (ROW_NOON_NIGHT === 5) {
        let confirmed = false;
        for (let c = 2; c <= maxCol; c++) {
            const v = cellText(ws.getRow(5).getCell(c)).trim();
            if (v === '昼' || v === '夜') { confirmed = true; break; }
        }
        if (!confirmed) {
            for (let r = 1; r <= Math.min(ws.rowCount, 12); r++) {
                let found = false;
                for (let c = 2; c <= maxCol; c++) {
                    const v = cellText(ws.getRow(r).getCell(c)).trim();
                    if (v === '昼' || v === '夜') { found = true; break; }
                }
                if (found) { ROW_NOON_NIGHT = r; break; }
            }
        }
    }
    // '日'行（col A = '日'）が存在すれば、そこから日番号 → 列インデックスを読む（Python出力形式）
    // 存在しなければ row 1 を pandas ヘッダー行として読む（通常の先月データ形式）
    let dayLabelRow = -1;
    for (let r = 1; r <= ws.rowCount; r++) {
        if (cellText(ws.getRow(r).getCell(1)) === '日') { dayLabelRow = r; break; }
    }
    const colsByDay = new Map(); // dayNum → [{colIdx, noonNight}]
    if (dayLabelRow > 0) {
        // Python 出力形式: '日'行から日番号 → 列インデックスをマッピング
        const dayRow = ws.getRow(dayLabelRow);
        for (let c = 2; c <= maxCol; c++) {
            const dayNum = parseInt(cellText(dayRow.getCell(c)));
            if (isNaN(dayNum) || dayNum < 1 || dayNum > 31) continue;
            const noonNight = cellText(ws.getRow(ROW_NOON_NIGHT).getCell(c)).trim();
            if (!colsByDay.has(dayNum)) colsByDay.set(dayNum, []);
            colsByDay.get(dayNum).push({ colIdx: c, noonNight });
        }
    } else {
        // 通常形式: row 1 を Date 型 / 数値タイトルで列マッピング
        ws.getRow(1).eachCell({ includeEmpty: false }, (cell, colIdx) => {
            if (colIdx <= 1) return;
            let dayNum;
            if (cell.value instanceof Date) {
                if (cell.value.getMonth() + 1 !== loadedMonth) return;
                dayNum = cell.value.getDate();
            } else {
                const title = String(cell.value ?? '').trim();
                if (!title || title.includes('/')) return;
                dayNum = parseInt(title);
            }
            if (isNaN(dayNum) || dayNum <= 0 || dayNum > 31) return;
            const noonNight = cellText(ws.getRow(ROW_NOON_NIGHT).getCell(colIdx)).trim();
            if (!colsByDay.has(dayNum)) colsByDay.set(dayNum, []);
            colsByDay.get(dayNum).push({ colIdx, noonNight });
        });
    }

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
    for (const rowIdx of personRowIndices) {
        personNames.push(cellText(ws.getRow(rowIdx).getCell(1)));
        // '日'行あり = Python形式 = duty_count列なし。なければ通常形式でcol2を読む
        dutyCounts.push(dayLabelRow <= 0 ? cellText(ws.getRow(rowIdx).getCell(2)) : '');
    }

    // 最後10日分のセルデータを収集
    const importedDayData = new Map(); // dayNum → { isRestDay, noDuty?, noDutyNoon?, noDutyNight?, rowValues }
    for (const dayNum of last10Days) {
        const cols    = colsByDay.get(dayNum);
        // '昼'/'夜'が明示されていれば優先。
        // saved fileでは'start'行が削除されて昼夜情報がないため、
        // 列数2=休日(昼+夜)、列数1=平日 で判定するフォールバックを使う。
        const noonCol  = cols.find(c => c.noonNight === '昼') ?? (cols.length >= 2 ? cols[0] : null);
        const nightCol = cols.find(c => c.noonNight === '夜') ?? (cols.length >= 2 ? cols[1] : null);
        const isRestDay = !!(noonCol && nightCol);

        if (isRestDay) {
            const isNoDutyMark = v => v === '○' || v === '〇' || v === '◯';
            const noDutyNoon  = isNoDutyMark(cellText(ws.getRow(ROW_NO_DUTY).getCell(noonCol.colIdx)).trim());
            const noDutyNight = isNoDutyMark(cellText(ws.getRow(ROW_NO_DUTY).getCell(nightCol.colIdx)).trim());
            const rowValues = [];
            for (const rowIdx of personRowIndices) {
                rowValues.push({
                    noon:  cellText(ws.getRow(rowIdx).getCell(noonCol.colIdx)),
                    night: cellText(ws.getRow(rowIdx).getCell(nightCol.colIdx))
                });
            }
            importedDayData.set(dayNum, { isRestDay: true, noDutyNoon, noDutyNight, rowValues });
        } else {
            const col    = cols[0];
            const isNoDutyMark2 = v => v === '○' || v === '〇' || v === '◯';
            const noDuty = isNoDutyMark2(cellText(ws.getRow(ROW_NO_DUTY).getCell(col.colIdx)).trim());
            const rowValues = [];
            for (const rowIdx of personRowIndices) {
                rowValues.push({ value: cellText(ws.getRow(rowIdx).getCell(col.colIdx)) });
            }
            importedDayData.set(dayNum, { isRestDay: false, noDuty, rowValues });
        }
    }

    // 開始日を「読み込んだ月の翌月の1日」に設定（これで前10日分＝読み込んだ月の最後10日と一致する）。
    // 期間の長さ（日数）は現在選択されている長さを維持する。
    const dispMonth = loadedMonth === 12 ? 1 : loadedMonth + 1;
    const dispYear  = loadedMonth === 12 ? loadedYear + 1 : loadedYear;
    const newStartDate = new Date(dispYear, dispMonth - 1, 1);
    const curStartDate = parseDateInput(document.getElementById('select-start-date').value);
    const curEndDate   = parseDateInput(document.getElementById('select-end-date').value);
    const spanDays = Math.round((curEndDate - curStartDate) / 86400000);
    const newEndDate = new Date(newStartDate);
    newEndDate.setDate(newEndDate.getDate() + spanDays);
    document.getElementById('select-start-date').value = formatDateInput(newStartDate);
    document.getElementById('select-end-date').value   = formatDateInput(newEndDate);
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
        const realDate = new Date(loadedYear, loadedMonth - 1, dayNum);
        const base = `prev_day${formatYYYYMMDD(realDate)}`;
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

    if (noDutyRow) await noDutyRow.update(noDutyUpdateObj);
    await Promise.all(dataRows.map((row, i) => rowUpdateObjs[i] ? row.update(rowUpdateObjs[i]) : null));
    table.redraw(true);
    updateProvisionalDutyCountDisplay();
    autoSizeNameColumn('name');

    } catch (err) {
        console.error('[load] エラー:', err);
        await window.api.showMessageBox({
            type: 'error',
            title: '読み込みエラー',
            message: `ファイルの読み込みに失敗しました。\n\n${err.message}`
        });
    } finally {
        hideLoading();
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
        showLoading();

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
        // 列構造はそのまま維持し、データのみ更新する

        // テーブルの開始日から前月・当月を取得（フォールバック用）
        const tableStartDate = parseDateInput(document.getElementById('select-start-date').value);
        const tableYear       = tableStartDate.getFullYear();
        const tableMonth      = tableStartDate.getMonth() + 1;
        const tablePrevMonth  = new Date(tableYear, tableMonth - 1, 0).getMonth() + 1;

        // セル値を正規化するヘルパー（数式セルはresultを使用）
        const getRaw = (cell) => {
            let v = cell.value;
            if (v && typeof v === 'object' && 'formula' in v) v = v.result ?? v;
            return v;
        };

        // スキップするラベル（ヘッダー行）
        const SKIP_LABELS = new Set([
            '', '　', '名前', '仮当直回数',
            '曜日', '祝日', '昼夜', '休業日', '当直不要',
            'start', 'end', '応援医師',
        ]);

        const hdrRow = ws.getRow(1);
        const hdrColCount = ws.columnCount;

        // ── Pass 0: データ検証（〇×輪番リスト）から「計算範囲」列を検出 ──
        // メイン画面のExcel出力では、計算対象の列（当月列）の人物行セルにのみ
        // 「　,〇,×,輪番」のリスト入力規則を設定しているため、これを手がかりに
        // 前月分（参考表示）と当月分（計算範囲）の境目を判定する。
        const hasCalcRangeValidation = (cell) => {
            const dv = cell.dataValidation;
            if (!dv || dv.type !== 'list') return false;
            const f = Array.isArray(dv.formulae) ? dv.formulae[0] : '';
            return typeof f === 'string' && f.includes('〇') && f.includes('×') && f.includes('輪番');
        };
        let samplePersonRow = -1;
        for (let r = 1; r <= ws.rowCount; r++) {
            const col1 = cellText(ws.getRow(r).getCell(1));
            if (col1 && !SKIP_LABELS.has(col1)) { samplePersonRow = r; break; }
        }
        const isTargetCol = new Set();
        if (samplePersonRow > 0) {
            const pr = ws.getRow(samplePersonRow);
            for (let c = 3; c <= hdrColCount; c++) {
                if (hasCalcRangeValidation(pr.getCell(c))) isTargetCol.add(c);
            }
        }

        // ── Pass 1: 前月分の最終列から、計算範囲が開始する月を正確に判定する ──
        // 前月分（参考表示）は必ず"m/d"表記という出力規約を利用し、計算範囲の直前列
        // （＝前月分の最終列）の月日を読み取って、その翌日が属する月を計算範囲の開始月とする。
        // 計算範囲は「前月分の翌月」とは限らない点に注意（前月分は開始日の前10日固定であり、
        // 月末を跨がなければ計算範囲は前月分と同じ月から始まることもある）。
        const tryParseMD = (v) => {
            if (v instanceof Date) return { m: v.getMonth() + 1, d: v.getDate() };
            if (v !== null && v !== undefined) {
                const s = String(v).trim();
                if (s.includes('/')) {
                    const parts = s.split('/');
                    const m = parseInt(parts[0]), d = parseInt(parts[1]);
                    if (!isNaN(m) && m >= 1 && m <= 12 && !isNaN(d)) return { m, d };
                }
            }
            return null;
        };
        let excelPrevMonth = null, excelPrevDay = null;
        if (isTargetCol.size > 0) {
            // 計算範囲の直前列から遡って、最初に見つかる"m/d"表記のセルを探す
            // （週末昼夜ペアの2列目は空欄のためスキップされる）
            for (let ci = Math.min(...isTargetCol) - 1; ci >= 3; ci--) {
                const parsed = tryParseMD(getRaw(hdrRow.getCell(ci)));
                if (parsed) { excelPrevMonth = parsed.m; excelPrevDay = parsed.d; break; }
            }
        }
        if (excelPrevMonth === null) {
            // isTargetCol が使えない場合は、従来通り1行目を先頭から走査する
            for (let ci = 3; ci <= hdrColCount; ci++) {
                const parsed = tryParseMD(getRaw(hdrRow.getCell(ci)));
                if (parsed) { excelPrevMonth = parsed.m; excelPrevDay = parsed.d; break; }
            }
        }
        // 検出できなかった場合はテーブルのセレクタを使用
        if (excelPrevMonth === null) excelPrevMonth = tablePrevMonth;

        let excelCurMonth;
        if (excelPrevDay !== null) {
            const daysInPrevMonth = new Date(tableYear, excelPrevMonth, 0).getDate();
            excelCurMonth = (excelPrevDay >= daysInPrevMonth)
                ? (excelPrevMonth === 12 ? 1 : excelPrevMonth + 1)
                : excelPrevMonth;
        } else {
            excelCurMonth = (excelPrevMonth % 12) + 1;
        }
        // ── Excelの月とテーブルの月が異なれば自動切り替え ──
        // （開始日の日にち・期間の長さ（日数）は維持したまま、月だけをExcelに合わせる）
        if (excelCurMonth !== tableMonth) {
            let targetYear = tableYear;
            const diff = excelCurMonth - tableMonth;
            if (diff < -6) targetYear = tableYear + 1; // 例: table=12月, excel=1月 → 翌年
            if (diff > 6)  targetYear = tableYear - 1; // 例: table=1月, excel=12月 → 前年

            const startInput = document.getElementById('select-start-date');
            const endInput   = document.getElementById('select-end-date');
            const curStart = parseDateInput(startInput.value);
            const curEnd   = parseDateInput(endInput.value);
            const spanDays = Math.round((curEnd - curStart) / 86400000);
            const daysInTargetMonth = new Date(targetYear, excelCurMonth, 0).getDate();
            const newStart = new Date(targetYear, excelCurMonth - 1, Math.min(curStart.getDate(), daysInTargetMonth));
            const newEnd = new Date(newStart);
            newEnd.setDate(newEnd.getDate() + spanDays);

            startInput.value = formatDateInput(newStart);
            endInput.value   = formatDateInput(newEnd);
            await updateTableStructure();
        }

        // ── Pass 2: Tabulator フィールドの month_day 逆引きマップ構築 ──
        // キー例: "7_25_night" → "prev_day{実日付}_night", "8_1" → "day{実日付}"
        // テーブル切り替え後に table.getColumns() を呼ぶので正しい構造が反映される。
        // 前月分・計算範囲は連続した重複のない期間なので、実際の「月_日」だけをキーにすれば
        // prev_day/day の区別なく一意に引ける（計算範囲が複数月にまたがっても問題ない）。
        const fieldByDate = new Map();
        table.getColumns().filter(c => !c.getDefinition().frozen).forEach(c => {
            const field = c.getField();
            const dateObj = dateFromField(field);
            if (!dateObj) return;
            const month = dateObj.getMonth() + 1;
            const day = dateObj.getDate();
            const m = field.match(/(_noon|_night)$/);
            fieldByDate.set(`${month}_${day}${m ? m[1] : ''}`, field);
        });

        // Excel から昼夜行を動的に探す
        let noonNightRowIdx = -1;
        for (let r = 1; r <= ws.rowCount; r++) {
            if (cellText(ws.getRow(r).getCell(1)) === '昼夜') { noonNightRowIdx = r; break; }
        }

        // ── Pass 3: colIdx → Tabulator field マッピングを構築 ──
        // 全列を明示的にループし、空ヘッダーセルは「直前の日付を継承」する。
        // これにより週末昼夜ペアの夜列（row1が空）も正しくマッピングできる。
        const colToField = new Map();
        let lastM = null, lastD = null; // 直前の非空セルの月・日
        // "/"を含まない裸の数字（＝計算範囲の列）の月を推定するためのカウンタ。
        // 計算範囲が複数月にまたがる場合、日が前の列より小さくなった時点で月が繰り上がったとみなす。
        let runningMonth = excelPrevMonth;

        for (let colIdx = 3; colIdx <= hdrColCount; colIdx++) {
            const v = getRaw(hdrRow.getCell(colIdx));
            let cellMonth, cellDay;
            let inherited = false; // 直前の日付を引き継いだか

            if (v === null || v === undefined || String(v).trim() === '') {
                // 空セル → 直前の日付を引き継ぐ（昼夜ペアの夜列など）
                if (lastM === null) continue;
                cellMonth = lastM;
                cellDay   = lastD;
                inherited = true;
            } else if (v instanceof Date) {
                cellMonth = v.getMonth() + 1;
                cellDay   = v.getDate();
                runningMonth = cellMonth;
            } else {
                const header = String(v).trim();
                if (header.includes('/')) {
                    const parts = header.split('/');
                    cellMonth = parseInt(parts[0]);
                    cellDay   = parseInt(parts[1]);
                    runningMonth = cellMonth;
                } else {
                    cellDay = parseInt(header);
                    if (!isNaN(cellDay) && lastD !== null && cellDay < lastD) {
                        // 前の列より日が小さくなった＝月が繰り上がった（例: 30 → 1）
                        runningMonth = runningMonth === 12 ? 1 : runningMonth + 1;
                    }
                    cellMonth = runningMonth;
                }
            }

            if (isNaN(cellMonth) || isNaN(cellDay) || cellDay < 1 || cellDay > 31) continue;
            if (!inherited) { lastM = cellMonth; lastD = cellDay; }

            const nn = noonNightRowIdx > 0
                ? cellText(ws.getRow(noonNightRowIdx).getCell(colIdx)).trim()
                : '';
            const suffix = nn === '昼' ? '_noon' : nn === '夜' ? '_night' : '';
            const key    = `${cellMonth}_${cellDay}${suffix}`;
            const field  = fieldByDate.get(key);
            if (field) colToField.set(colIdx, field);
        }
        const allTabRows = table.getRows();
        let noDutyUpdate = null;
        const personList = [];

        for (let r = 1; r <= ws.rowCount; r++) {
            const col1 = cellText(ws.getRow(r).getCell(1));
            const col2 = cellText(ws.getRow(r).getCell(2));

            if (col1 === '当直不要' || col1 === '応援医師') {
                // 当直不要行: ○/〇 → true
                const update = {};
                for (const [colIdx, field] of colToField) {
                    const v = cellText(ws.getRow(r).getCell(colIdx));
                    update[field] = (v === '○' || v === '〇' || v === 'true');
                }
                noDutyUpdate = update;
            } else if (col1 && !SKIP_LABELS.has(col1)) {
                // 人物行: 日付マッピングを使用してフィールドに値をセット
                const pData = { name: col1, duty_count: col2 };
                for (const [colIdx, field] of colToField) {
                    pData[field] = cellText(ws.getRow(r).getCell(colIdx));
                }
                personList.push(pData);
            }
        }

        // 当直不要行を更新
        if (noDutyUpdate) {
            const noDutyRow = allTabRows.find(r => r.getData().id === 'row_no_duty');
            if (noDutyRow) await noDutyRow.update(noDutyUpdate);
        }

        // 既存の人物行を削除してから再追加
        const existingPersonRows = allTabRows.filter(r => typeof r.getData().id === 'number');
        for (const row of existingPersonRows) await row.delete();

        let nextId = 0;
        for (const pData of personList) {
            await table.addRow({ id: nextId++, ...pData }, false);
        }
        // 末尾に空白行
        await table.addRow({ id: nextId, name: '', duty_count: '' }, false);

        table.redraw(true);
        currentDutyCountField = 'duty_count';
        updateDutyCountDisplay();
        updateProvisionalDutyCountDisplay();
        autoSizeNameColumn('name');

    } catch (err) {
        await window.api.showMessageBox({ type: 'error', title: '読み込みエラー', message: err.message });
    } finally {
        hideLoading();
    }
}

const kibouButton = document.getElementById('load-kibou-button');
if (kibouButton) kibouButton.addEventListener('click', loadKibouSheet);

async function runDutyAssignment() {
    showLoading();
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
        // 対象範囲（day*列）の実日付一覧（"YYYY-MM-DD"）。結果ウィンドウでの祝日判定・保存時ラベルに使用
        let mainDates = null;

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
            const stripTitleHtml = (h) => (h || '').replace(/<[^>]*>/g, '').trim();
            let pastColIdx = -1, startColIdx = -1, endColIdx = -1;
            for (let i = 2; i < nCols; i++) {
                const title = stripTitleHtml(columns[i].getDefinition().title);
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
                const t = stripTitleHtml(columns[i].getDefinition().title);
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
            mainDates = dateFields.slice(startColIdx, endColIdx + 1)
                .map(f => formatDateInput(dateFromField(f)));

            const noDutyData = allRows.find(r => r.getData().id === 'row_no_duty')?.getData() ?? {};
            const dayData    = allRows.find(r => r.getData().id === 'header_day')?.getData()    ?? {};
            const dateN = fields.length;
            const specialRuleColIdx = dateN; // 末尾に追加する「特別条件」列（昼勤務の翌日も夜勤務可能）
            const n = dateN + 1;
            const O = 2;

            const getDateNum   = f => { const d = dateFromField(f); return d ? d.getDate() : 0; };
            const getShiftType = f => f.endsWith('_noon') ? '昼' : '夜';

            // Row 0: past / start / end / special_rule 列マーカー
            const r0 = Array(n).fill('');
            r0[O + pastColIdx] = 'past'; r0[O + startColIdx] = 'start'; r0[O + endColIdx] = 'end';
            r0[specialRuleColIdx] = 'special_rule';
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
                er[specialRuleColIdx] = specialRuleNames.has((data.name ?? '').trim()) ? '〇' : '';
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
            const scoreText = result.message.split('\n')
                .filter(l => !l.startsWith('勤務表を') && l.trim() !== '')
                .join('\n').trim() || null;
            await window.api.openResultWindow(pathMatch[1], mainDates, scoreText);
        } else {
            const debugInfo = `\n\n【デバッグ用】\n入力Excel: ${tempPath}\nログ: %USERPROFILE%\\Documents\\DutyAssignmentLogs\\duty_assign.log`;
            await window.api.showMessageBox({
                type: 'warning', title: '解なし',
                message: result.message + debugInfo
            });
        }

    } catch (err) {
        await window.api.showMessageBox({ type: 'error', title: 'エラー', message: err.message });
    } finally {
        hideLoading();
    }
}

const runDutyButton = document.getElementById('run-duty-button');
if (runDutyButton) runDutyButton.addEventListener('click', runDutyAssignment);

const openFileButton = document.getElementById('open-file-button');
if (openFileButton) openFileButton.addEventListener('click', async () => {
    const filePath = await window.api.openFileDialog();
    if (filePath) {
        showLoading();
        try {
            await executePythonScript(filePath);
        } finally {
            hideLoading();
        }
    }
});

// ── ズーム機能 ────────────────────────────────────────────────
(function initZoom() {
    const MIN = 0.4, MAX = 2.5, STEP = 0.1;
    let hideTimer = null;

    function badge() {
        let el = document.getElementById('zoom-badge');
        if (!el) {
            el = document.createElement('div');
            el.id = 'zoom-badge';
            Object.assign(el.style, {
                position: 'fixed', bottom: '22px', right: '18px',
                background: 'rgba(30,30,30,0.75)', color: '#fff',
                padding: '4px 14px', borderRadius: '999px',
                fontSize: '13px', fontWeight: '600', letterSpacing: '0.04em',
                zIndex: '9999', pointerEvents: 'none',
                transition: 'opacity 0.35s ease',
                backdropFilter: 'blur(4px)',
            });
            document.body.appendChild(el);
        }
        return el;
    }

    function applyZoom(delta) {
        const next = parseFloat(
            Math.min(MAX, Math.max(MIN, window.api.getZoomFactor() + delta)).toFixed(2)
        );
        window.api.setZoomFactor(next);

        const el = badge();
        el.textContent = `${Math.round(next * 100)} %`;
        el.style.opacity = '1';
        clearTimeout(hideTimer);
        hideTimer = setTimeout(() => { el.style.opacity = '0'; }, 1800);
    }

    // Ctrl + ホイール
    window.addEventListener('wheel', (e) => {
        if (!e.ctrlKey) return;
        e.preventDefault();
        applyZoom(e.deltaY < 0 ? STEP : -STEP);
    }, { passive: false });

    // Ctrl + Plus / Minus / 0
    window.addEventListener('keydown', (e) => {
        if (!e.ctrlKey) return;
        if (e.key === '=' || e.key === '+') { e.preventDefault(); applyZoom(STEP); }
        else if (e.key === '-')             { e.preventDefault(); applyZoom(-STEP); }
        else if (e.key === '0')             { e.preventDefault(); applyZoom(1 - window.api.getZoomFactor()); }
    });
}());

