/**
 * 工事日報 Ver5.1 追加分：スタッフ登録・一時保存（下書き）
 *
 * このファイルは既存の Apps Script プロジェクトに「新しいファイル」として追加します。
 * 既存の doPost の「const p = JSON.parse(e.postData.contents);」のすぐ下に、
 * 次の1行を足してください（それ以外の既存コードは触りません）。
 *
 *   if (p.action) return handleNippouAction_(p);
 *
 * ※ 必ず「報告データ」への書き込みより前に置くこと。後ろに置くと、
 *    一時保存のたびに空の日報が記録されてしまいます。
 *
 * シート「スタッフ」「下書き」は、無ければ自動で作られます。
 */

// LINE Developers の LINEログインチャネルのチャネルID（LIFF ID のハイフンより前の数字）
const NIPPOU_LINE_CHANNEL_ID = '2012765015';

const NIPPOU_STAFF_SHEET = 'スタッフ';
const NIPPOU_DRAFT_SHEET = '下書き';
const NIPPOU_STAFF_HEADER = ['LINEユーザーID', '名前', '企業名', '電話番号', '登録日時', '更新日時'];
const NIPPOU_DRAFT_HEADER = ['LINEユーザーID', '名前', '保存日時', '下書きデータ'];

/** フォームからの action 付きリクエストをさばく */
function handleNippouAction_(body) {
  try {
    const userId = verifyLineIdToken_(body.idToken);
    let result;
    switch (body.action) {
      case 'getMe':       result = getMe_(userId); break;
      case 'registerStaff': result = registerStaff_(userId, body.staff || {}); break;
      case 'saveDraft':   result = saveDraft_(userId, body.draft); break;
      case 'deleteDraft': result = deleteDraft_(userId); break;
      default: throw new Error('unknown_action');
    }
    return nippouJson_(Object.assign({ ok: true }, result));
  } catch (err) {
    return nippouJson_({ ok: false, error: String(err && err.message || err) });
  }
}

/** LIFF の IDトークンを LINE に照会して、本人の LINEユーザーID を取り出す */
function verifyLineIdToken_(idToken) {
  if (!idToken) throw new Error('no_token');
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post',
    payload: { id_token: idToken, client_id: NIPPOU_LINE_CHANNEL_ID },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('invalid_token');
  const info = JSON.parse(res.getContentText());
  if (!info.sub) throw new Error('invalid_token');
  return info.sub;
}

/** 登録情報と下書きをまとめて返す（フォームを開いた時に1回呼ぶ） */
function getMe_(userId) {
  const staff = findStaff_(userId);
  const draft = findDraft_(userId);
  return { staff: staff, draft: draft };
}

function registerStaff_(userId, s) {
  const name = String(s.name || '').trim();
  const company = String(s.company || '').trim();
  const phone = String(s.phone || '').trim();
  if (!name || !company || !phone) throw new Error('missing_fields');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = nippouSheet_(NIPPOU_STAFF_SHEET, NIPPOU_STAFF_HEADER);
    sheet.getRange('D:D').setNumberFormat('@'); // 電話番号の先頭の0を消さない
    const now = new Date();
    const row = findRowByUserId_(sheet, userId);
    if (row) {
      sheet.getRange(row, 2, 1, 3).setValues([[name, company, phone]]);
      sheet.getRange(row, 6).setValue(now);
    } else {
      sheet.appendRow([userId, name, company, phone, now, now]);
    }
  } finally {
    lock.releaseLock();
  }
  return { staff: { name: name, company: company, phone: phone } };
}

function saveDraft_(userId, draft) {
  if (!draft) throw new Error('no_draft');
  const json = JSON.stringify(draft);
  if (json.length > 45000) throw new Error('draft_too_large'); // セル上限(5万字)の手前で止める

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = nippouSheet_(NIPPOU_DRAFT_SHEET, NIPPOU_DRAFT_HEADER);
    const staff = findStaff_(userId);
    const now = new Date();
    const values = [userId, staff ? staff.name : '', now, json];
    const row = findRowByUserId_(sheet, userId);
    if (row) sheet.getRange(row, 1, 1, values.length).setValues([values]);
    else sheet.appendRow(values);
    return { savedAt: Utilities.formatDate(now, 'Asia/Tokyo', 'M/d HH:mm') };
  } finally {
    lock.releaseLock();
  }
}

function deleteDraft_(userId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = nippouSheet_(NIPPOU_DRAFT_SHEET, NIPPOU_DRAFT_HEADER);
    const row = findRowByUserId_(sheet, userId);
    if (row) sheet.deleteRow(row);
  } finally {
    lock.releaseLock();
  }
  return {};
}

function findStaff_(userId) {
  const sheet = nippouSheet_(NIPPOU_STAFF_SHEET, NIPPOU_STAFF_HEADER);
  const row = findRowByUserId_(sheet, userId);
  if (!row) return null;
  const v = sheet.getRange(row, 2, 1, 3).getValues()[0];
  return { name: v[0], company: v[1], phone: String(v[2]) };
}

function findDraft_(userId) {
  const sheet = nippouSheet_(NIPPOU_DRAFT_SHEET, NIPPOU_DRAFT_HEADER);
  const row = findRowByUserId_(sheet, userId);
  if (!row) return null;
  const v = sheet.getRange(row, 3, 1, 2).getValues()[0];
  try {
    return {
      savedAt: Utilities.formatDate(new Date(v[0]), 'Asia/Tokyo', 'M/d HH:mm'),
      data: JSON.parse(v[1])
    };
  } catch (e) {
    return null;
  }
}

/** A列（LINEユーザーID）が一致する行番号を返す。無ければ 0 */
function findRowByUserId_(sheet, userId) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === userId) return i + 2;
  }
  return 0;
}

function nippouSheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function nippouJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
