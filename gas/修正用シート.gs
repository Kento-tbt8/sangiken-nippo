/**
 * 【修正用】日報シート：送った日報を直して、記録を上書きし、PDFを再発行する
 *
 * 使い方
 *   1. 管理No.（J1）に直したい日報の番号を入れる
 *   2. 「このデータを修正する」ボタン → 報告データ・材料データから中身を読み込む
 *   3. シート上で直す
 *   4. 「PDF再発行」ボタン →
 *        ・PDFを「日報PDF／元請け名」フォルダに【修正_日付】付きで保存
 *        ・報告データ・材料データの同じ管理No.の行を、直した内容で上書き
 *          （行を薄い黄色にして、管理No.のセルに「修正版」のメモを付ける）
 *        ・上書き前の内容は「修正履歴」シートに残す
 *        ・修正用シートの記入欄を全部消す
 *
 * ボタンに割り当てる関数名は今までと同じです（copyDataToDailyReport / saveAndResetDailyReport）。
 * ※ 前の copyDataToDailyReport・saveAndResetDailyReport は、必ず消してからこのファイルを貼ってください
 *    （同じ名前の関数が2つあると、どちらが動くか分からなくなります）。
 */

const FIX_SHEET_NAME = '【修正用】日報シート';
const FIX_NO_CELL = 'J1';            // 管理No. を入れるセル
const FIX_PRINT_RANGE = 'A4:X41';    // PDFにする範囲
const FIX_HISTORY_SHEET = '修正履歴';
const FIX_MARK_COLOR = '#FFF2CC';    // 修正した行の色（薄い黄色）

// 修正用シートのセル ⇔ 報告データの列（0 = A列）
const FIX_FIELDS = [
  { cell: 'V4',  col: 2 },   // 記入者
  { cell: 'R4',  col: 3 },   // 作業日
  { cell: 'I8',  col: 4 },   // 開始時間
  { cell: 'I10', col: 5 },   // 終了時間
  { cell: 'I6',  col: 6 },   // 天気
  { cell: 'G4',  col: 7 },   // 現場名
  { cell: 'C8',  col: 8 },   // 監督者
  { cell: 'C6',  col: 9 },   // 元請け
  { cell: 'C10', col: 10 },  // 営業担当者
  { cell: 'C12', col: 11 },  // 人工
  { cell: 'I12', col: 12 },  // 作業者
  { cell: 'C14', col: 13 },  // 応援者
  { cell: 'A36', col: 14 }   // 備考
];
const FIX_LOC_FIRST_COL = 15;                    // 施工箇所1 が始まる列（P列）
const FIX_LOC_ROWS = [17, 20, 23, 26, 29, 32];   // 施工箇所の行（A=箇所, J=数量, L=単位）
const FIX_MAT_BLOCKS = [                          // 材料（名前の列, 数量の列）
  { name: 'M', qty: 'Q' },
  { name: 'S', qty: 'W' }
];
const FIX_MAT_FIRST_ROW = 7;
const FIX_MAT_LAST_ROW = 37;

/**
 * 「このデータを修正する」ボタン：管理No.の日報を修正用シートに読み込む
 */
function copyDataToDailyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(FIX_SHEET_NAME);
  if (!sheet) { ui.alert('「' + FIX_SHEET_NAME + '」シートが見つかりません。'); return; }

  const no = readFixNo_(sheet);
  if (!no) { ui.alert('管理No.（' + FIX_NO_CELL + '）に番号を入れてください。'); return; }

  const report = findRowByNo_(ss.getSheetByName('報告データ'), no);
  if (!report) { ui.alert('報告データに 管理No.' + no + ' が見つかりません。'); return; }
  const material = findRowByNo_(ss.getSheetByName('材料データ'), no);

  clearFixInputs_(sheet, false);

  // 基本情報
  FIX_FIELDS.forEach(f => sheet.getRange(f.cell).setValue(cellValue_(report.values[f.col])));

  // 施工箇所（3列ずつ：箇所・数量・単位）
  FIX_LOC_ROWS.forEach((row, i) => {
    const c = FIX_LOC_FIRST_COL + i * 3;
    sheet.getRange('A' + row).setValue(cellValue_(report.values[c]));
    sheet.getRange('J' + row).setValue(cellValue_(report.values[c + 1]));
    sheet.getRange('L' + row).setValue(cellValue_(report.values[c + 2]));
  });
  const extraLocs = countLocsBeyond_(report.values);

  // 材料（D列から 名前・数量 の順に並んでいる）
  const pairs = [];
  if (material) {
    for (let i = 3; i < material.values.length; i += 2) {
      if (material.values[i] === '' && material.values[i + 1] === '') continue;
      pairs.push([cellValue_(material.values[i]), cellValue_(material.values[i + 1])]);
    }
  }
  const perBlock = FIX_MAT_LAST_ROW - FIX_MAT_FIRST_ROW + 1;
  FIX_MAT_BLOCKS.forEach((b, bi) => {
    const chunk = pairs.slice(bi * perBlock, (bi + 1) * perBlock);
    const names = [], qtys = [];
    for (let r = 0; r < perBlock; r++) {
      names.push([chunk[r] ? chunk[r][0] : '']);
      qtys.push([chunk[r] ? chunk[r][1] : '']);
    }
    sheet.getRange(b.name + FIX_MAT_FIRST_ROW + ':' + b.name + FIX_MAT_LAST_ROW).setValues(names);
    sheet.getRange(b.qty + FIX_MAT_FIRST_ROW + ':' + b.qty + FIX_MAT_LAST_ROW).setValues(qtys);
  });

  // どの管理No.を読み込んだかを覚えておく（PDF再発行の時に、番号の入れ違いを防ぐ）
  PropertiesService.getDocumentProperties().setProperty('FIX_LOADED_NO', String(no));

  let msg = '管理No.' + no + ' を読み込みました。直したら「PDF再発行」を押してください。';
  if (extraLocs > 0) msg += '\n※ 施工箇所が7つ以上あります。7つ目以降はシートに出ませんが、記録には残したままにします。';
  if (pairs.length > perBlock * FIX_MAT_BLOCKS.length) msg += '\n※ 材料が多く、シートに入りきらない分があります。入りきらない分は記録から消えるので注意してください。';
  ss.toast(msg, '読み込み完了', 5);
}

/**
 * 「PDF再発行」ボタン：PDFを作り直し、記録を上書きし、シートを空にする
 */
function saveAndResetDailyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName(FIX_SHEET_NAME);
  if (!sheet) { ui.alert('「' + FIX_SHEET_NAME + '」シートが見つかりません。'); return; }

  const no = readFixNo_(sheet);
  const loadedNo = PropertiesService.getDocumentProperties().getProperty('FIX_LOADED_NO');
  if (!no) { ui.alert('管理No.（' + FIX_NO_CELL + '）が空です。'); return; }
  if (String(no) !== loadedNo) {
    ui.alert('管理No.' + no + ' はまだ読み込まれていません。\n先に「このデータを修正する」を押して読み込んでから直してください。');
    return;
  }

  const ok = ui.alert('PDF再発行',
    '管理No.' + no + ' の記録を、今の内容で上書きしてPDFを作り直します。\n' +
    '（上書き前の内容は「' + FIX_HISTORY_SHEET + '」シートに残ります）\n\nよろしいですか？',
    ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  let fileName;
  try {
    SpreadsheetApp.flush();
    const reportSheet = ss.getSheetByName('報告データ');
    const materialSheet = ss.getSheetByName('材料データ');
    const report = findRowByNo_(reportSheet, no);
    if (!report) throw new Error('報告データに 管理No.' + no + ' が見つかりません。');
    const material = findRowByNo_(materialSheet, no);

    const v = cell => sheet.getRange(cell).getValue();
    const site = v('G4'), prime = v('C6'), record = v('V4');

    // 1. PDF（先に作る。失敗したら記録は何も変えない）
    const dateStr = formatFixDate_(v('R4'));
    const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    fileName = '【' + (prime || '名称なし') + '】' + dateStr + '_' + (site || '現場名不明') + '【修正_' + today + '】.pdf';
    exportFixPdf_(ss, sheet, prime, fileName);

    const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

    // 2. 報告データを上書き（管理No.と送信日時はそのまま）
    const tz = ss.getSpreadsheetTimeZone();
    const newReport = report.values.slice(0, FIX_LOC_FIRST_COL);
    FIX_FIELDS.forEach(f => { newReport[f.col] = v(f.cell); });
    // 時刻はシートから読むと日付付きになるので「9:31」の形に戻す
    [4, 5].forEach(c => {
      if (newReport[c] instanceof Date) newReport[c] = Utilities.formatDate(newReport[c], tz, 'H:mm');
    });
    FIX_LOC_ROWS.forEach(row => {
      const loc = v('A' + row), qty = v('J' + row), unit = v('L' + row);
      if (loc !== '' || qty !== '') newReport.push(loc, qty, unit);
    });
    // 7つ目以降の施工箇所（シートに出ていない分）は元のまま残す
    const tailStart = FIX_LOC_FIRST_COL + FIX_LOC_ROWS.length * 3;
    newReport.push(...trimEnd_(report.values.slice(tailStart)));

    saveHistory_(ss, '報告データ', no, report.values, stamp);
    overwriteRow_(reportSheet, report.row, report.values.length, newReport, stamp);

    // 3. 材料データを上書き（材料の行が無かった日報に材料を足した時は、行を足す）
    const newMat = [];
    FIX_MAT_BLOCKS.forEach(b => {
      for (let r = FIX_MAT_FIRST_ROW; r <= FIX_MAT_LAST_ROW; r++) {
        const name = v(b.name + r);
        if (name !== '') newMat.push(name, v(b.qty + r));
      }
    });
    if (material) {
      const matRow = material.values.slice(0, 3).concat(newMat);
      saveHistory_(ss, '材料データ', no, material.values, stamp);
      overwriteRow_(materialSheet, material.row, material.values.length, matRow, stamp);
    } else if (newMat.length > 0 && materialSheet) {
      materialSheet.appendRow([no, report.values[1], record].concat(newMat));
      markRow_(materialSheet, materialSheet.getLastRow(), newMat.length + 3, stamp);
    }

    // 4. 修正用シートを空にする
    clearFixInputs_(sheet, true);
    PropertiesService.getDocumentProperties().deleteProperty('FIX_LOADED_NO');
    SpreadsheetApp.flush();
  } catch (err) {
    ui.alert('PDF再発行が途中で止まりました。\n\n' + err.message);
    return;
  } finally {
    lock.releaseLock();
  }

  ui.alert('完了', '管理No.' + no + ' を修正しました。\n\n・PDF：' + fileName +
    '\n・報告データ／材料データの行を上書き（黄色の行）\n・修正用シートを空にしました', ui.ButtonSet.OK);
}

/* ---------------- ここから下は補助の関数 ---------------- */

function readFixNo_(sheet) {
  const n = Number(String(sheet.getRange(FIX_NO_CELL).getValue()).trim());
  return n > 0 ? n : null;
}

/** A列が管理No.の行を探す（見出しの1行目は飛ばす） */
function findRowByNo_(sheet, no) {
  if (!sheet || sheet.getLastRow() < 2) return null;
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  for (let i = 0; i < data.length; i++) {
    if (typeof data[i][0] === 'number' && data[i][0] === no) {
      return { row: i + 2, values: data[i] };
    }
  }
  return null;
}

/** 記録の値をシートに書く形にする（undefined を空にするだけ） */
function cellValue_(v) {
  return v === undefined || v === null ? '' : v;
}

function countLocsBeyond_(values) {
  const tail = values.slice(FIX_LOC_FIRST_COL + FIX_LOC_ROWS.length * 3);
  return tail.some(x => x !== '') ? 1 : 0;
}

function trimEnd_(arr) {
  let end = arr.length;
  while (end > 0 && arr[end - 1] === '') end--;
  return arr.slice(0, end);
}

function formatFixDate_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(d || '日付なし').replace(/\//g, '-');
}

/** 行を新しい内容で上書きし、修正版の印を付ける */
function overwriteRow_(sheet, row, oldLength, newValues, stamp) {
  const width = Math.max(oldLength, newValues.length);
  if (width > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), width - sheet.getMaxColumns());
  sheet.getRange(row, 1, 1, width).clearContent();
  sheet.getRange(row, 1, 1, newValues.length).setValues([newValues]);
  markRow_(sheet, row, width, stamp);
}

function markRow_(sheet, row, width, stamp) {
  sheet.getRange(row, 1, 1, width).setBackground(FIX_MARK_COLOR);
  const a = sheet.getRange(row, 1);
  const old = a.getNote();
  a.setNote('修正版（' + stamp + ' 修正）' + (old ? '\n' + old : ''));
}

/** 上書き前の行を「修正履歴」シートに残す */
function saveHistory_(ss, sheetName, no, values, stamp) {
  let h = ss.getSheetByName(FIX_HISTORY_SHEET);
  if (!h) {
    h = ss.insertSheet(FIX_HISTORY_SHEET);
    h.appendRow(['修正日時', '元のシート', '管理No.', '↓ 修正する前の内容（元のシートのA列から順に）']);
    h.setFrozenRows(1);
  }
  h.appendRow([stamp, sheetName, no].concat(trimEnd_(values)));
}

/** 修正用シートの記入欄を消す（withNo が true なら管理No.も消す） */
function clearFixInputs_(sheet, withNo) {
  const cells = FIX_FIELDS.map(f => f.cell);
  FIX_LOC_ROWS.forEach(r => cells.push('A' + r, 'J' + r, 'L' + r));
  FIX_MAT_BLOCKS.forEach(b => {
    cells.push(b.name + FIX_MAT_FIRST_ROW + ':' + b.name + FIX_MAT_LAST_ROW);
    cells.push(b.qty + FIX_MAT_FIRST_ROW + ':' + b.qty + FIX_MAT_LAST_ROW);
  });
  if (withNo) cells.push(FIX_NO_CELL);
  sheet.getRangeList(cells).clearContent();
}

/** 修正用シートの A4:X41 だけをA4横でPDFにして、元請けのフォルダに保存する */
function exportFixPdf_(ss, sheet, prime, fileName) {
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' +
    'exportFormat=pdf&format=pdf&size=A4&portrait=false&fitw=true' +
    '&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false' +
    '&gid=' + sheet.getSheetId() +
    '&range=' + encodeURIComponent(FIX_PRINT_RANGE);
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('PDFの書き出しに失敗しました（' + res.getResponseCode() + '）');
  }
  const blob = res.getBlob().setName(fileName);
  return getPrimeFolder_(getPdfFolder_(ss.getId()), prime).createFile(blob);
}
