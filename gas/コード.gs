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
 * 決まったフォルダに保存したい時だけ、フォルダのリンク（またはID）を入れてください。
 *   例：const PDF_FOLDER_ID = 'https://drive.google.com/drive/folders/xxxxxxxx';
 * その中に元請けごとのフォルダが作られます。
 * ※ GASを実行するアカウント（デプロイした人）が、そのフォルダの「編集者」である必要があります。
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

    // 送った本人のLINEユーザーID（PDFをLINEで送るのに使う）。確認できなくても日報の記録は続ける
    let lineUserId = null;
    let lineStatus = '';
    try {
      if (p.idToken) lineUserId = verifyLineIdToken_(p.idToken);
      else lineStatus = 'フォームからLINEのログイン情報が届いていません（フォームが古い版の可能性）';
    } catch (err) {
      lineStatus = 'LINEの本人確認に失敗しました（' + err.message + '）';
      console.warn(lineStatus);
    }

    let pdfFile = null;

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

        pdfFile = createPdf(dailySheet, p.date, p.site, p.primeContractor); // PDF作成
      }
    } finally {
      lock.releaseLock();
    }

    // 送った本人に、PDFのリンクをLINEで送る（失敗しても日報の記録はそのまま）
    if (!pdfFile && !lineStatus) lineStatus = 'PDFが作られていません（「日報シート」が見つからない可能性）';
    if (pdfFile && lineUserId) {
      try {
        lineStatus = sendPdfToLine_(lineUserId, p, pdfFile);
      } catch (err) {
        lineStatus = 'LINEへの送信に失敗しました（' + err.message + '）';
        console.error(lineStatus);
      }
    }

    return ContentService.createTextOutput(JSON.stringify({result: 'success', line: lineStatus})).setMimeType(ContentService.MimeType.JSON);
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

  if (response.getResponseCode() !== 200) {
    throw new Error('PDFの書き出しに失敗しました（' + response.getResponseCode() + '）');
  }
  const blob = response.getBlob().setName(fileName);
  return getPrimeFolder_(getPdfFolder_(ssId), primeContractor).createFile(blob);
}

/**
 * 送った本人に、日報PDFのリンクをLINEで送る
 *
 * 必要な設定：Apps Script の「プロジェクトの設定」→「スクリプト プロパティ」に
 *   LINE_CHANNEL_ACCESS_TOKEN … 公式LINE（Messaging API チャネル）のチャネルアクセストークン（長期）
 * を登録する。未登録なら何もしない。
 */
function sendPdfToLine_(userId, p, file) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token) {
    const msg = 'スクリプト プロパティに LINE_CHANNEL_ACCESS_TOKEN が登録されていません';
    console.warn(msg);
    return msg;
  }

  // リンクを知っている人だけが見られるようにする（LINEのトークから開けるようにするため）
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const message = buildReportFlex_(p, file.getUrl());

  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to: userId, messages: [message] }),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code === 200) return 'sent';
  if (code === 401) throw new Error('チャネルアクセストークンが正しくありません');
  if (code === 400 && /to/.test(res.getContentText())) {
    throw new Error('送り先のユーザーが見つかりません。LINEログインチャネルとMessaging APIチャネルが同じプロバイダーか確認してください');
  }
  throw new Error('LINE送信エラー ' + code + ' ' + res.getContentText());
}

/** PDFの保存先フォルダを返す（PDF_FOLDER_ID が空なら、スプシと同じ場所の「日報PDF」フォルダ） */
function getPdfFolder_(ssId) {
  if (PDF_FOLDER_ID) {
    // フォルダのリンクをそのまま貼っても、ID部分だけ取り出して使う
    const m = String(PDF_FOLDER_ID).match(/folders\/([\w-]+)/);
    return DriveApp.getFolderById(m ? m[1] : String(PDF_FOLDER_ID).trim());
  }

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
 * LINEに送る「日報受付」カード（Flex Message）
 * 上：緑の帯に「日報を受け付けました」／中：日報の要点／下：「PDFを確認する」ボタン
 */
function buildReportFlex_(p, pdfUrl) {
  const val = (v) => (v === undefined || v === null || String(v).trim() === '') ? '―' : String(v);

  // 2026-11-24 → 2026年11月24日（火）
  let dateLabel = val(p.date);
  const dm = String(p.date || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (dm) {
    const d = new Date(Number(dm[1]), Number(dm[2]) - 1, Number(dm[3]));
    dateLabel = dm[1] + '年' + Number(dm[2]) + '月' + Number(dm[3]) + '日（' + '日月火水木金土'[d.getDay()] + '）';
  }

  const time = (p.startTime || p.endTime) ? (val(p.startTime) + ' 〜 ' + val(p.endTime)) : '―';
  const workers = val(p.name) + (p.count ? '（' + p.count + '名）' : '');
  const locations = (p.locationArray || [])
    .filter(l => l.loc)
    .map(l => '・' + l.loc + (l.qty ? '　' + l.qty + (l.unit || '') : ''))
    .join('\n');

  const row = (label, value) => ({
    type: 'box', layout: 'baseline', spacing: 'md',
    contents: [
      { type: 'text', text: label, size: 'sm', color: '#8C8C8C', flex: 2 },
      { type: 'text', text: value, size: 'sm', color: '#333333', flex: 5, wrap: true }
    ]
  });

  const body = [
    { type: 'text', text: val(p.site), weight: 'bold', size: 'lg', wrap: true, color: '#222222' },
    { type: 'text', text: dateLabel, size: 'sm', color: '#666666', margin: 'sm' },
    { type: 'separator', margin: 'lg' },
    {
      type: 'box', layout: 'vertical', margin: 'lg', spacing: 'sm',
      contents: [
        row('元請け', val(p.primeContractor)),
        row('監督', val(p.director)),
        row('作業時間', time),
        row('天気', val(p.weather)),
        row('作業者', workers),
        row('記入者', val(p.record))
      ]
    }
  ];
  if (locations) {
    body.push(
      { type: 'separator', margin: 'lg' },
      { type: 'text', text: '施工内容', size: 'sm', color: '#8C8C8C', margin: 'lg' },
      { type: 'text', text: locations, size: 'sm', color: '#333333', wrap: true, margin: 'sm' }
    );
  }

  return {
    type: 'flex',
    altText: '日報を受け付けました（' + val(p.site) + '）',
    contents: {
      type: 'bubble',
      size: 'mega',
      header: {
        type: 'box', layout: 'vertical', backgroundColor: '#28A745', paddingAll: '16px',
        contents: [
          { type: 'text', text: '✓ 日報を受け付けました', color: '#FFFFFF', weight: 'bold', size: 'md' }
        ]
      },
      body: { type: 'box', layout: 'vertical', paddingAll: '16px', contents: body },
      footer: {
        type: 'box', layout: 'vertical', paddingAll: '12px',
        contents: [
          {
            type: 'button', style: 'primary', color: '#007BFF', height: 'sm',
            action: { type: 'uri', label: 'PDFを確認する', uri: pdfUrl }
          }
        ]
      }
    }
  };
}

/**
 * LINE送信の確認用：エディタで実行すると、「スタッフ」シートの一番下の人に
 * テストのカードをLINEで送り、結果を実行ログに出す
 */
function testLinePush() {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  console.log('① トークン: ' + (token ? '登録あり（' + token.length + '文字）' : '★未登録★ プロジェクトの設定 → スクリプト プロパティ に LINE_CHANNEL_ACCESS_TOKEN を登録してください'));
  if (!token) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('スタッフ');
  if (!sheet || sheet.getLastRow() < 2) {
    console.log('② ★「スタッフ」シートに登録がありません。LINEからフォームを開いてスタッフ登録してください');
    return;
  }
  const last = sheet.getRange(sheet.getLastRow(), 1, 1, 2).getValues()[0];
  console.log('② 送り先: ' + last[1] + '（ユーザーID ' + String(last[0]).slice(0, 6) + '…）');

  // 友だちかどうか（プロフィールが取れれば、この公式LINEの友だち）
  const prof = UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/' + last[0], {
    headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
  });
  console.log('③ 友だち確認: ' + prof.getResponseCode() + ' ' + prof.getContentText());
  if (prof.getResponseCode() === 401) console.log('   ★トークンが正しくありません（コピー漏れ・別チャネルのトークンの可能性）');
  if (prof.getResponseCode() === 404) console.log('   ★この人は公式LINEの友だちではない、またはLINEログインチャネルと公式LINEのプロバイダーが違います');

  const msg = buildReportFlex_({
    date: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'), site: 'テスト送信', record: last[1],
    locationArray: [{ loc: 'これはテストです', qty: '', unit: '' }]
  }, 'https://drive.google.com/');
  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ to: last[0], messages: [msg] }),
    muteHttpExceptions: true
  });
  console.log('④ 送信結果: ' + res.getResponseCode() + ' ' + res.getContentText());
  if (res.getResponseCode() === 200) console.log('   → 送信成功。LINEにテストのカードが届いているはずです');
}

/** カードの見た目確認用：実行するとログにカードのJSONが出る（LINEには送らない） */
function testReportFlex() {
  const sample = {
    date: '2026-11-24', startTime: '08:00', endTime: '17:00', weather: '晴れ',
    site: '〇〇ビル', director: '山田太郎', primeContractor: 'A社', record: '山田太郎',
    name: '酒井、森、長谷部、馬場', count: '4',
    locationArray: [{ loc: '基礎', qty: '12', unit: '㎡' }, { loc: '外壁', qty: '30', unit: 'm' }]
  };
  console.log(JSON.stringify(buildReportFlex_(sample, 'https://drive.google.com/'), null, 2));
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
