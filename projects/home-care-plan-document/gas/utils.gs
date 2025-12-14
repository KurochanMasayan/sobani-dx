/**
 * ログ/共通ユーティリティ
 */
function logInfo(message, data) {
  if (data) {
    Logger.log(`[INFO] ${message} ${JSON.stringify(data)}`);
  } else {
    Logger.log(`[INFO] ${message}`);
  }
}

function logError(functionName, error) {
  Logger.log(`[ERROR] ${functionName}: ${error && error.stack ? error.stack : error}`);
}

function notify(message) {
  Logger.log(`[NOTICE] ${message}`);
}

function normalizeJapaneseDate(text) {
  if (!text) return '';
  const eraMap = { 令和: 2018, 平成: 1988, 昭和: 1925, 大正: 1911 };
  const match = text.match(/(令和|平成|昭和|大正)(\d+)年(\d+)月(\d+)日/);
  if (match) {
    const year = eraMap[match[1]] + parseInt(match[2], 10);
    const month = (`0${match[3]}`).slice(-2);
    const day = (`0${match[4]}`).slice(-2);
    return `${year}-${month}-${day}`;
  }
  return text;
}

function formatZipCode(zip) {
  const digits = (zip || '').replace(/[^0-9]/g, '');
  if (digits.length === 7) {
    return `〒${digits.slice(0, 3)}-${digits.slice(3)}`;
  }
  return zip;
}

function extractSpreadsheetId(url) {
  if (!url) return null;
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

function sanitizeMultiline(text) {
  return (text || '').replace(/\\n/g, '\n').trim();
}

function normalizeRomanNumerals(text) {
  if (!text) return text;
  return text
    .replace(/�W/g, 'Ⅳ')
    .replace(/�V/g, 'Ⅲ')
    .replace(/�U/g, 'Ⅱ')
    .replace(/�T/g, 'Ⅰ');
}

