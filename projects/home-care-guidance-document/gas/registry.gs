/**
 * 事業所ファイルURLレジストリ
 *
 * 列構成: 年月, 居宅支援事業所名, URL, 最終更新
 * 同じ月の処理は同じファイルに追記される
 */
function getRegistrySheet() {
  const active = getActiveSpreadsheetOrThrow();
  const sheet = active.getSheetByName(BASE_CONFIG.registrySheetName);
  if (!sheet) {
    throw new Error(`「${BASE_CONFIG.registrySheetName}」シートが見つかりません。ヘッダー付きで作成してください。`);
  }
  return sheet;
}

/**
 * 事業所+年月でレコードを検索
 * @param {Sheet} sheet - 台帳シート
 * @param {string} facilityName - 居宅支援事業所名
 * @param {string} yearMonth - 年月（例: 2025-12）
 * @returns {Object|null} 見つかった場合は { rowIndex, url, fileId }
 */
function lookupFacilityRecord(sheet, facilityName, yearMonth) {
  const normalizedFacility = normalizeString(facilityName);
  const normalizedYearMonth = normalizeString(yearMonth);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    // 年月は日付型として解釈される可能性があるため、文字列に変換して年月形式に正規化
    const rawYearMonth = values[i][0];
    const storedYearMonth = normalizeYearMonth(rawYearMonth);
    const storedFacility = normalizeString(values[i][1]);

    // 年月 + 施設名 の2条件で検索
    if (storedYearMonth === normalizedYearMonth &&
        storedFacility === normalizedFacility) {
      logInfo('既存レコードを発見', { rowIndex: i + 1, storedYearMonth, storedFacility });
      const url = values[i][2] || '';
      return {
        rowIndex: i + 1,
        url,
        fileId: extractSpreadsheetId(url)
      };
    }
  }
  logInfo('既存レコードなし', { facilityName: normalizedFacility, yearMonth: normalizedYearMonth });
  return null;
}

/**
 * 年月を正規化（日付型や様々な形式に対応）
 * @param {*} value - 年月の値（文字列、日付型など）
 * @returns {string} 正規化された年月（例: 2026-01）
 */
function normalizeYearMonth(value) {
  if (!value) return '';

  // 日付型の場合
  if (value instanceof Date) {
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM');
  }

  // 文字列の場合、すでに正しい形式ならそのまま返す
  const str = String(value).trim();
  if (/^\d{4}-\d{2}$/.test(str)) {
    return str;
  }

  // その他の形式は試行錯誤
  const date = new Date(str);
  if (!isNaN(date.getTime())) {
    return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM');
  }

  return str;
}

/**
 * 事業所ファイルを台帳に登録
 * @param {Sheet} sheet - 台帳シート
 * @param {string} facilityName - 居宅支援事業所名
 * @param {string} yearMonth - 年月
 * @param {string} fileId - スプレッドシートID
 * @returns {number} 登録された行番号
 */
function recordFacilityFile(sheet, facilityName, yearMonth, fileId) {
  const normalizedFacility = normalizeString(facilityName);
  const normalizedYearMonth = normalizeString(yearMonth);
  const url = `https://docs.google.com/spreadsheets/d/${fileId}`;
  const rowIndex = sheet.getLastRow() + 1;
  sheet.getRange(rowIndex, 1, 1, 4).setValues([[
    normalizedYearMonth,
    normalizedFacility,
    url,
    new Date()
  ]]);
  // 書き込みを即座に反映（continueGuidanceGenerationで再開時に確実に検索できるように）
  SpreadsheetApp.flush();
  logInfo('台帳に事業所ファイルを登録しました', {
    facilityName,
    yearMonth,
    rowIndex,
    url
  });
  return rowIndex;
}

/**
 * 事業所ファイルの最終更新日時を更新
 * @param {Sheet} sheet - 台帳シート
 * @param {number} rowIndex - 行番号
 */
function updateFacilityLastTouched(sheet, rowIndex) {
  sheet.getRange(rowIndex, 4).setValue(new Date());  // 4列目: 最終更新
}
