// ===== 設定 =====
const SPREADSHEET_ID = '1sWrRdtkgQTjKPHO6iP42VscGOM_fHl9XFkVpnTGmt9Y';
const SHEET_NAME = 'シート1';

// ★ここに現在使用中のチャネルアクセストークンを貼り付け
const LINE_CHANNEL_ACCESS_TOKEN = 'ここにLINE Messaging APIチャネルアクセストークンを入れてください';
const LINE_TO_ID = 'ここに通知先のUser IDまたはGroup IDを入れてください';
const LINE_API_URL = 'https://api.line.me/v2/bot/message/push';

// A列からM列までの見出し
const EXPECTED_HEADERS = [
  '登録日時','姓','名','フリガナ(姓)','フリガナ(名)',
  '郵便番号','住所','電話番号','性別','生年月日',
  '暗証番号','DM','台番号'
];

function doPost(e){
  try{
    const sheet = getTargetSheet_();
    const payload = extractPayload_(e);
    const data = normalizePayload_(payload);
    ensureHeaderRow_(sheet);
    const row = buildRow_(data);
    sheet.appendRow(row);
    sendLineNotification(data);
    return jsonResponse_({status:'success',message:'saved'});
  } catch (err) {
    return jsonResponse_({status:'error',message:err && err.message ? err.message : 'Unknown error'});
  }
}

function doGet(){
  return jsonResponse_({status:'ok',message:'This endpoint accepts POST requests.'});
}

function doOptions(){
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

function extractPayload_(e){
  if (e && e.parameter && Object.keys(e.parameter).length > 0) {
    return e.parameter;
  }
  const postData = e && e.postData && e.postData.contents ? e.postData.contents : '';
  if (!postData) return {};
  const params = {};
  postData.split('&').forEach(pair => {
    if (!pair) return;
    const [rawKey, ...rawValue] = pair.split('=');
    const key = decodeURIComponent(rawKey.replace(/\+/g,' '));
    const value = decodeURIComponent((rawValue.join('=')||'').replace(/\+/g,' '));
    params[key] = value;
  });
  return params;
}

function getTargetSheet_(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}

function buildRow_(data){
  return EXPECTED_HEADERS.map(h => data[h] ?? '');
}

function getValue_(data, keys){
  for (const key of keys) {
    const value = String(data[key] ?? '').trim();
    if (value) return value;
  }
  return '';
}

function ensureHeaderRow_(sheet){
  const cols = sheet.getMaxColumns();
  const hc = EXPECTED_HEADERS.length;
  if (cols > hc) sheet.deleteColumns(hc + 1, cols - hc);
  if (cols < hc) sheet.insertColumnsAfter(cols, hc - cols);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(EXPECTED_HEADERS);
    return;
  }
  sheet.getRange(1,1,1,hc).setValues([EXPECTED_HEADERS]);
}

function normalizePayload_(payload){
  const d = {};
  d['姓'] = String(payload['姓'] ?? payload['familyName'] ?? '').trim();
  d['名'] = String(payload['名'] ?? payload['givenName'] ?? '').trim();
  d['フリガナ(姓)'] = String(payload['フリガナ(姓)'] ?? payload['familyNameKana'] ?? '').trim();
  d['フリガナ(名)'] = String(payload['フリガナ(名)'] ?? payload['givenNameKana'] ?? '').trim();
  d['郵便番号'] = String(payload['郵便番号'] ?? payload['postalCode'] ?? '').trim();
  d['住所'] = String(payload['住所'] ?? payload['address'] ?? '').trim();
  let phone = String(payload['電話番号'] ?? payload['phone'] ?? '').trim();
  if (phone && /^[0-9]+$/.test(phone) && phone.startsWith('0')) phone = "'" + phone;
  d['電話番号'] = phone;
  d['性別'] = String(payload['性別'] ?? payload['gender'] ?? '').trim();
  const era = String(payload['birthEra'] || payload['era'] || '').trim();
  const y = String(payload['birthYear'] || '').trim();
  const m = String(payload['birthMonth'] || '').trim();
  const day = String(payload['birthDay'] || '').trim();
  if (era && y && m && day) {
    d['生年月日'] = `${era}${y}年${m.padStart(2,'0')}月${day.padStart(2,'0')}日`;
  } else {
    d['生年月日'] = String(payload['生年月日'] ?? payload['birthDate'] ?? '').trim();
  }
  d['暗証番号'] = String(payload['暗証番号'] ?? payload['pin'] ?? '').trim();
  d['DM'] = String(payload['DM'] ?? payload['dm'] ?? '').trim();
  d['台番号'] = String(payload['台番号'] ?? payload['machineNumber'] ?? '').trim();
  d['登録日時'] = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  return d;
}

function sendLineNotification(data){
  if (!LINE_CHANNEL_ACCESS_TOKEN || LINE_CHANNEL_ACCESS_TOKEN.startsWith('ここに')) return;
  if (!LINE_TO_ID || LINE_TO_ID.startsWith('ここに')) return;
  const payload = {
    to: LINE_TO_ID,
    messages: [{
      type: 'text',
      text: `新規会員登録\n\n氏名：${data['姓']} ${data['名']}\nフリガナ：${data['フリガナ(姓)']} ${data['フリガナ(名)']}\n郵便番号：${data['郵便番号']}\n住所：${data['住所']}\n電話番号：${data['電話番号']}\n性別：${data['性別']}\n生年月日：${data['生年月日']}\n暗証番号：${data['暗証番号']}\nDM：${data['DM']}\n台番号：${data['台番号']}\n登録日時：${data['登録日時']}`
    }]
  };
  UrlFetchApp.fetch(LINE_API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

function jsonResponse_(data){
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
