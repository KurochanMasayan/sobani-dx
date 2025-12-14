/**
 * 事業所ファイルURLレジストリ
 *
 * 変更履歴:
 * - 患者単位から年月+居宅支援事業所単位に変更
 * - 列構成: 年月, 居宅支援事業所名, URL, 最終更新
 */
function getRegistrySheet() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('テンプレートスプレッドシートから実行してください');
  }
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
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const storedYearMonth = String(values[i][0] || '').trim();
    const storedFacility = String(values[i][1] || '').trim();
    if (storedYearMonth === yearMonth && storedFacility === facilityName) {
      const url = values[i][2] || '';
      return {
        rowIndex: i + 1,
        url,
        fileId: extractSpreadsheetId(url)
      };
    }
  }
  return null;
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
  const url = `https://docs.google.com/spreadsheets/d/${fileId}`;
  const rowIndex = sheet.getLastRow() + 1;
  sheet.getRange(rowIndex, 1, 1, 4).setValues([[
    yearMonth,
    facilityName,
    url,
    new Date()
  ]]);
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
  sheet.getRange(rowIndex, 4).setValue(new Date());
}
