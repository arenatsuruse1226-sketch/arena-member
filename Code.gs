const SPREADSHEET_ID = '1sWrRdtkgQTjKPHO6iP42VscGOM_fHl9XFkVpnTGmt9Y';
const SHEET_NAME = 'シート1';

function doPost(e) {
  try {
    const sheet = getTargetSheet_();
    const data = normalizePayload_(e && e.parameter ? e.parameter : {});

    const headers = getHeaders_();
    ensureHeaderRow_(sheet, headers);

    const row = buildRow_(data);
    sheet.appendRow(row);

    return jsonResponse_({
      status: 'success',
      message: 'saved',
    });
  } catch (error) {
    return jsonResponse_({
      status: 'error',
      message: error && error.message ? error.message : 'Unknown error',
    });
  }
}

function doGet() {
  return jsonResponse_({
    status: 'ok',
    message: 'This endpoint accepts POST requests.',
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

function getHeaders_() {
  return [
    'submittedAt',
    '登録日時',
    '姓',
    '名',
    'フリガナ(姓)',
    'フリガナ(名)',
    '郵便番号',
    '住所',
    '電話番号',
    '性別',
    '生年月日',
    '暗証番号',
    'DM',
    '台番号',
    '利用規約同意',
    'familyName',
    'givenName',
    'familyNameKana',
    'givenNameKana',
    'postalCode',
    'phone',
    'address',
    'gender',
    'birthEra',
    'birthYear',
    'birthMonth',
    'birthDay',
    'pin',
    'dm',
    'machineNumber',
    'agree',
  ];
}

function buildRow_(data) {
  return [
    getValue_(data, ['submittedAt']),
    getValue_(data, ['登録日時']),
    getValue_(data, ['姓', 'familyName']),
    getValue_(data, ['名', 'givenName']),
    getValue_(data, ['フリガナ(姓)', 'familyNameKana']),
    getValue_(data, ['フリガナ(名)', 'givenNameKana']),
    getValue_(data, ['郵便番号', 'postalCode']),
    getValue_(data, ['住所', 'address']),
    getValue_(data, ['電話番号', 'phone']),
    getValue_(data, ['性別', 'gender']),
    getValue_(data, ['生年月日']),
    getValue_(data, ['暗証番号', 'pin']),
    getValue_(data, ['DM', 'dm']),
    getValue_(data, ['台番号', 'machineNumber']),
    getValue_(data, ['利用規約同意', 'agree']),
    getValue_(data, ['familyName', '姓']),
    getValue_(data, ['givenName', '名']),
    getValue_(data, ['familyNameKana', 'フリガナ(姓)']),
    getValue_(data, ['givenNameKana', 'フリガナ(名)']),
    getValue_(data, ['postalCode', '郵便番号']),
    getValue_(data, ['phone', '電話番号']),
    getValue_(data, ['address', '住所']),
    getValue_(data, ['gender', '性別']),
    getValue_(data, ['birthEra']),
    getValue_(data, ['birthYear']),
    getValue_(data, ['birthMonth']),
    getValue_(data, ['birthDay']),
    getValue_(data, ['pin', '暗証番号']),
    getValue_(data, ['dm', 'DM']),
    getValue_(data, ['machineNumber', '台番号']),
    getValue_(data, ['agree', '利用規約同意']),
  ];
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

function ensureHeaderRow_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return;
  }

  const firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const firstRowText = firstRow.map((value) => String(value || '').trim());
  const shouldReplace = headers.some((header, index) => firstRowText[index] !== header);

  if (shouldReplace) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function normalizePayload_(payload) {
  const normalized = {};
  Object.keys(payload).forEach((key) => {
    normalized[key] = String(payload[key] ?? '');
  });

  if (!normalized['台番号'] && normalized.machineNumber) {
    normalized['台番号'] = normalized.machineNumber;
  }

  if (!normalized.machineNumber && normalized['台番号']) {
    normalized.machineNumber = normalized['台番号'];
  }

  if (!normalized['登録日時']) {
    normalized['登録日時'] = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  }

  if (!normalized['submittedAt']) {
    normalized['submittedAt'] = new Date().toISOString();
  }

  return normalized;
}

function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
