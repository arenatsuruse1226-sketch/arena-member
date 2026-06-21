const SPREADSHEET_ID = '1sWrRdtkgQTjKPHO6iP42VscGOM_fHl9XFkVpnTGmt9Y';
const SHEET_NAME = 'シート1';

// A列からM列までの正しい見出しの順番を定義
const EXPECTED_HEADERS = [
  '登録日時', '姓', '名', 'フリガナ(姓)', 'フリガナ(名)', 
  '郵便番号', '住所', '電話番号', '性別', '生年月日', '暗証番号', 'DM', '台番号'
];

function doPost(e) {
  try {
    const sheet = getTargetSheet_();
    
    // 1. 送られてきたデータを綺麗な日本語キーに統一・和暦に変換する
    const data = normalizePayload_(e && e.parameter ? e.parameter : {});
    
    // 2. 見出し行の数をA〜M列（13列）に矯正し、余分な列（N列以降）を削除する
    ensureHeaderRow_(sheet);
    
    // 3. 整理された日本語データから、A〜M列の順番通りに並んだ行データを作る
    const row = buildRow_(data);
    sheet.appendRow(row);
    
    return jsonResponse_({
      status: 'success',
      message: 'saved'
    });
  } catch (error) {
    return jsonResponse_({
      status: 'error',
      message: error && error.message ? error.message : 'Unknown error'
    });
  }
}

function doGet() {
  return jsonResponse_({
    status: 'ok',
    message: 'This endpoint accepts POST requests.'
  });
}

function getTargetSheet_() {
  if (SPREADSHEET_ID === 'ここにスプレッドシートIDを入れてください') {
    throw new Error('SPREADSHEET_ID を設定してください');
  }
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
  return sheet;
}

/**
 * 修正：整理された日本語データから、EXPECTED_HEADERSの順番通りに値を並べる
 */
function buildRow_(data) {
  return EXPECTED_HEADERS.map(header => data[header] ?? '');
}

function getValue_(data, keys) {
  for (const key of keys) {
    const value = String(data[key] ?? '').trim();
    if (value) {
      return value;
    }
  }
  return '';
}

/**
 * シートの見出しをA〜M列（13列）に固定し、不要な右側の列を削除する
 */
function ensureHeaderRow_(sheet) {
  const currentColumns = sheet.getMaxColumns();
  const headerCount = EXPECTED_HEADERS.length;

  if (currentColumns > headerCount) {
    sheet.deleteColumns(headerCount + 1, currentColumns - headerCount);
  } else if (currentColumns < headerCount) {
    sheet.insertColumnsAfter(currentColumns, headerCount - currentColumns);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(EXPECTED_HEADERS);
    return;
  }

  const firstRow = sheet.getRange(1, 1, 1, headerCount).getValues()[0];
  const firstRowText = firstRow.map(value => String(value || '').trim());
  const shouldReplace = EXPECTED_HEADERS.some((header, index) => header !== firstRowText[index]);

  if (shouldReplace) {
    sheet.getRange(1, 1, 1, headerCount).setValues([EXPECTED_HEADERS]);
  }
}

/**
 * 外部からのデータを日本語に変換し、生年月日を和暦に補正する
 */
function normalizePayload_(payload) {
  const normalized = {};

  // 各項目の文字列化と表記揺れの吸収
  normalized['姓'] = String(payload['姓'] ?? payload['familyName'] ?? '').trim();
  normalized['名'] = String(payload['名'] ?? payload['givenName'] ?? '').trim();
  normalized['フリガナ(姓)'] = String(payload['フリガナ(姓)'] ?? payload['familyNameKana'] ?? payload['familyNameKanji'] ?? '').trim();
  normalized['フリガナ(名)'] = String(payload['フリガナ(名)'] ?? payload['givenNameKana'] ?? payload['givenNameKanji'] ?? '').trim();
  normalized['郵便番号'] = String(payload['郵便番号'] ?? payload['postalCode'] ?? '').trim();
  normalized['住所'] = String(payload['住所'] ?? payload['address'] ?? '').trim();
  let phoneValue = String(payload['電話番号'] ?? payload['phone'] ?? '').trim();
  if (phoneValue && /^[0-9]+$/.test(phoneValue) && phoneValue.startsWith('0')) {
    phoneValue = `'${phoneValue}`;
  }
  normalized['電話番号'] = phoneValue;
  normalized['性別'] = String(payload['性別'] ?? payload['gender'] ?? '').trim();
  
  // --- 生年月日の和暦自動フォーマット処理 ---
  const bEra = String(payload['birthEra'] || payload['era'] || '').trim();
  const bYear = String(payload['birthYear'] || '').trim();
  const bMonth = String(payload['birthMonth'] || '').trim();
  const bDay = String(payload['birthDay'] || '').trim();

  if (bEra && bYear && bMonth && bDay) {
    const month = bMonth.padStart(2, '0');
    const day = bDay.padStart(2, '0');
    normalized['生年月日'] = `${bEra}${bYear}年${month}月${day}日`;
  } else {
    const rawBirth = getValue_(payload, ['生年月日', 'birthDate', 'birthday']);
    if (rawBirth) {
      const dateObj = new Date(rawBirth);
      if (!isNaN(dateObj.getTime())) {
        normalized['生年月日'] = Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy/MM/dd');
      } else {
        normalized['生年月日'] = rawBirth;
      }
    } else {
      normalized['生年月日'] = '';
    }
  }
  // -----------------------------------------

  normalized['暗証番号'] = String(payload['暗証番号'] ?? payload['pin'] ?? '').trim();
  normalized['DM'] = String(payload['DM'] ?? payload['dm'] ?? '').trim();
  normalized['台番号'] = String(payload['台番号'] ?? payload['machineNumber'] ?? '').trim();

  // 日時の自動補完
  if (!normalized['登録日時']) {
    const rawDate = payload['登録日時'] ?? payload['submittedAt'];
    if (rawDate) {
      normalized['登録日時'] = String(rawDate).trim();
    } else {
      normalized['登録日時'] = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    }
  }

  return normalized;
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}