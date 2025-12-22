/**
 * ログ/共通ユーティリティ
 */

/**
 * アクティブスプレッドシートを取得（なければエラー）
 * @returns {Spreadsheet} アクティブスプレッドシート
 * @throws {Error} スプレッドシートから実行されていない場合
 */
function getActiveSpreadsheetOrThrow() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('テンプレートスプレッドシートから実行してください');
  }
  return active;
}

/**
 * 文字列を正規化（null/undefined対応、トリム）
 * @param {*} value - 正規化する値
 * @returns {string} 正規化された文字列
 */
function normalizeString(value) {
  return String(value || '').trim();
}

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

/**
 * キューを施設名でソートする
 * 同じ施設の患者が連続して処理されるように並び替える
 * @param {string[]} queue - 患者IDの配列
 * @param {Map} patientMap - 患者ID -> 患者情報のMap
 * @returns {string[]} 施設名でソートされたキュー
 */
function sortQueueByFacility(queue, patientMap) {
  return queue.slice().sort((a, b) => {
    const patientA = patientMap.get(a);
    const patientB = patientMap.get(b);
    const facilityA = normalizeFacilityName(patientA ? patientA.facility : '') || '';
    const facilityB = normalizeFacilityName(patientB ? patientB.facility : '') || '';
    return facilityA.localeCompare(facilityB, 'ja');
  });
}

function normalizeRomanNumerals(text) {
  if (!text) return text;
  return text
    .replace(/�W/g, 'Ⅳ')
    .replace(/�V/g, 'Ⅲ')
    .replace(/�U/g, 'Ⅱ')
    .replace(/�T/g, 'Ⅰ');
}

