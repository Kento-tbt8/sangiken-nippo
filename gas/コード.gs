/** @OnlyCurrentDoc */

/**
 * PDFの保存先フォルダ
 *
 * 空のまま（''）にすると、このスプレッドシートと同じフォルダの中に
 * 「日報PDF」フォルダを自動で作って保存します。
 * その中は元請けごとのフォルダに分かれます（例：日報PDF/△△建設/【△△建設】2026-11-24_〇〇ビル.pdf）。
 * 元請けが空の日報は「名称なし」フォルダに入ります。
 * → スプシをお客様のドライブへ移したりコピーしたりすると、保存先も自動でお客様側になります。
 *
 * 決まったフォルダに保存したい時だけ、フォルダのID（URLの folders/ の後ろ）を入れてください。
 */
const PDF_FOLDER_ID = '';
const PDF_FOLDER_NAME = '日報PDF';

/**
 * 1. フォームからのデータを受け取り、各シートへ振り分け、PDFを作成する
 */
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const p = JSON.parse(e.postData.contents);

    // スタッフ登録・一時保存（「スタッフと下書き」ファイルで処理）
    if (p.action) return handleNippouAction_(p);

    // 同時に送信された時に、日報シート・管理番号・PDFが混ざらないよう1件ずつ処理する
    const lock = LockService.getScriptLock();
    lock.waitLock(60000);
    try {
      const now = new Date();

      // --- 1. 「報告データ」シート：蓄積用（データベース） ---
      const reportSheet = ss.getSheetByName("報告データ");
      if (reportSheet) {

        let lastRow = reportSheet.getLastRow();
        let managementNo = 1000 + lastRow; // 1行目がヘッダーなら、2行目に書き込むときは「1」になります

        // 基本データ
        let reportRow = [managementNo, now, p.record, p.date, p.startTime, p.endTime, p.weather, p.site, p.director, p.primeContractor, p.salesStaff, p.count, p.name, p.supportMember, p.notes];

        // 施工箇所（場所, 数量, 単位）を横に追加
        if (p.locationArray) {
          p.locationArray.forEach(item => {
            if (item.loc) {
              reportRow.push(item.loc, item.qty, item.unit);
            }
          });
        }
        reportSheet.appendRow(reportRow);
      }

      // --- 2. 「材料データ」シート：蓄積用（データベース） ---
      const materialSheet = ss.getSheetByName("材料データ");
      if (materialSheet) {

        let lastRow = reportSheet.getLastRow();
        let managementsNo = 999 + lastRow; // 1行目がヘッダーなら、2行目に書き込むときは「1」になります

        let materialRow = [managementsNo, now, p.record];
        // 通常材料
        if (p.materialArray && p.materialArray.length > 0) {
          p.materialArray.forEach(item => { if (item.name) materialRow.push(item.name, item.qty); });
        }
        // その他項目（経費など）を同じ行の続きに追加
        if (p.extraArray && p.extraArray.length > 0) {
          p.extraArray.forEach(extra => { if (extra.item) materialRow.push(extra.item, extra.qty); });
        }
        // 材料もその他項目も無い日報は、材料データに行を作らない
        if (materialRow.length > 3) materialSheet.appendRow(materialRow);
      }

      // --- 3. 「日報シート」：指定セルへの個別書き込み（印刷用） ---
      const dailySheet = ss.getSheetByName("日報シート");
      if (dailySheet) {
        // 基本情報
        dailySheet.getRange("G1").setValue(p.site); // 現場名
        dailySheet.getRange("R1").setValue(p.date); // 日付
        dailySheet.getRange("V1").setValue(p.record); // 記入者
        dailySheet.getRange("C3").setValue(p.primeContractor); // 元請け
        dailySheet.getRange("I3").setValue(p.weather); // 天気
        dailySheet.getRange("C5").setValue(p.director); // 現場監督
        dailySheet.getRange("C7").setValue(p.salesStaff); // 営業担当
        dailySheet.getRange("I5").setValue(p.startTime); // 開始
        dailySheet.getRange("I7").setValue(p.endTime); // 終了
        dailySheet.getRange("C9").setValue(p.count); // 人数
        dailySheet.getRange("I9").setValue(p.name); // 作業員名
        dailySheet.getRange("C11").setValue(p.supportMember || p.support || ""); // 応援者
        dailySheet.getRange("A33").setValue(p.notes); // 備考（全体）

        // 施工内容の書き込み（14行目〜29行目）
        // A列:内容, J列:数量, L列:単位
        dailySheet.getRange("A14:A29").clearContent();
        dailySheet.getRange("J14:J29").clearContent();
        dailySheet.getRange("L14:L29").clearContent();

        if (p.locationArray && p.locationArray.length > 0) {
          p.locationArray.forEach((item, index) => {
            if (index < 16) { // 最大16行まで
              let row = 14 + (index * 3);
              if (item.loc || item.qty) {
                dailySheet.getRange("A" + row).setValue(item.loc);
                dailySheet.getRange("J" + row).setValue(item.qty);
                dailySheet.getRange("L" + row).setValue(item.unit);
              }
            }
          });
        }

        // 使用材料・その他費用の書き込み（M/S列 2列分散）
        const columns = [
          { name: "M", qty: "Q", start: 4, end: 34 },
          { name: "S", qty: "W", start: 4, end: 34 }
        ];
        // 材料エリアを一旦クリア
        columns.forEach(col => {
          dailySheet.getRange(col.name + col.start + ":" + col.name + col.end).clearContent();
          dailySheet.getRange(col.qty + col.start + ":" + col.qty + col.end).clearContent();
        });

        // 通常材料とその他項目を一つのリストにまとめる
        let allItems = [];
        if (p.materialArray) allItems = allItems.concat(p.materialArray);
        if (p.extraArray) {
          const extras = p.extraArray.map(ex => ({ name: ex.item, qty: ex.qty }));
          allItems = allItems.concat(extras);
        }

        if (allItems.length > 0) {
          let colIdx = 0;
          let currentRow = columns[colIdx].start;
          allItems.forEach(item => {
            if (item.name && colIdx < columns.length) {
              let col = columns[colIdx];
              dailySheet.getRange(col.name + currentRow).setValue(item.name);
              dailySheet.getRange(col.qty + currentRow).setValue(item.qty);
              currentRow++;
              if (currentRow > col.end) {
                colIdx++;
                if (colIdx < columns.length) currentRow = columns[colIdx].start;
              }
            }
          });
        }

        SpreadsheetApp.flush(); // 反映確定

        createPdf(dailySheet, p.date, p.site, p.primeContractor); // PDF作成
      }
    } finally {
      lock.releaseLock();
    }

    return ContentService.createTextOutput(JSON.stringify({result: 'success'})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({result: 'error', message: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 2. 日報シートをPDF化してフォルダに保存する
 */
function createPdf(sheet, date, site, primeContractor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();
  const sheetId = sheet.getSheetId();
  const fileName = "【" + (primeContractor || "名称なし") + "】" + date + "_" + site + ".pdf";

  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?" +
    "exportFormat=pdf&format=pdf" +
    "&size=A4" +
    "&portrait=false" +
    "&fitw=true" +
    "&sheetnames=false&printtitle=false&pagenumbers=false" +
    "&gridlines=false" +
    "&gid=" + sheetId;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { "Authorization": "Bearer " + token },
    muteHttpExceptions: true
  });

  const blob = response.getBlob().setName(fileName);
  getPrimeFolder_(getPdfFolder_(ssId), primeContractor).createFile(blob);
}

/** PDFの保存先フォルダを返す（PDF_FOLDER_ID が空なら、スプシと同じ場所の「日報PDF」フォルダ） */
function getPdfFolder_(ssId) {
  if (PDF_FOLDER_ID) return DriveApp.getFolderById(PDF_FOLDER_ID);

  const parents = DriveApp.getFileById(ssId).getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  return getOrCreateFolder_(parent, PDF_FOLDER_NAME);
}

/** 保存先の中に、元請けごとのフォルダを返す（無ければ作る） */
function getPrimeFolder_(baseFolder, primeContractor) {
  const name = String(primeContractor || '').replace(/[\/\\]/g, '／').trim() || '名称なし';
  return getOrCreateFolder_(baseFolder, name);
}

function getOrCreateFolder_(parent, name) {
  const found = parent.getFoldersByName(name);
  return found.hasNext() ? found.next() : parent.createFolder(name);
}

/**
 * 3. 現場・材料リスト取得用（HTML表示時に呼び出される）
 */
function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 現場リスト
  const siteSheet = ss.getSheetByName("現場リスト");
  let siteData = [];
  if (siteSheet && siteSheet.getLastRow() >= 2) {
    const rows = siteSheet.getRange(2, 1, siteSheet.getLastRow() - 1, 6).getValues();
    siteData = rows.map(r => ({
      name: r[0], director: r[1], prime: r[2], staff: r[3], Count: r[4], Name: r[5]
    }));
  }

  // 材料リスト
  const matSheet = ss.getSheetByName("材料リスト");
  let matData = [];
  if (matSheet && matSheet.getLastRow() >= 2) {
    matData = matSheet.getRange(2, 1, matSheet.getLastRow() - 1, 1).getValues().map(r => r[0]);
  }

  return ContentService.createTextOutput(JSON.stringify({sites: siteData, materials: matData}))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 4. 承認用のテスト関数
 *    コードを貼り替えたら、デプロイ前にこれを1回実行して権限を許可してください。
 */
function testPermissions() {
  console.log("承認用のログ出力です。この関数を実行して許可を与えてください。");
  // PDF保存先フォルダの確認（無ければ作られます）
  console.log("PDF保存先: " + getPdfFolder_(SpreadsheetApp.getActiveSpreadsheet().getId()).getName());
  // LINEのログイン確認用（外部への接続）の権限もここで許可される
  UrlFetchApp.fetch("https://api.line.me/oauth2/v2.1/verify", { method: "post", payload: { id_token: "x", client_id: "0" }, muteHttpExceptions: true });
}
