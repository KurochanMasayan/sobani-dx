/**
 * 事業所ファイルURLレジストリ
 *
 * 変更履歴:
 * - 患者単位から年月+居宅支援事業所単位に変更
 * - 列構成: 年月, 居宅支援事業所名, URL, 実行ID, フォルダID, 最終更新
 * - 実行IDを追加（6分制限による中断再開で同じファイルに追記するため）
 * - フォルダIDを追加（再開時に確実に同じフォルダを使用するため）
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
 * 事業所+年月+実行IDでレコードを検索
 * @param {Sheet} sheet - 台帳シート
 * @param {string} facilityName - 居宅支援事業所名
 * @param {string} yearMonth - 年月（例: 2025-12）
 * @param {string} runId - 実行ID（6分制限対応）
 * @returns {Object|null} 見つかった場合は { rowIndex, url, fileId, folderId }
 */
function lookupFacilityRecord(sheet, facilityName, yearMonth, runId) {
  const normalizedFacility = normalizeString(facilityName);
  const normalizedYearMonth = normalizeString(yearMonth);
  const normalizedRunId = normalizeString(runId);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const storedYearMonth = normalizeString(values[i][0]);
    const storedFacility = normalizeString(values[i][1]);
    const storedRunId = normalizeString(values[i][3]);
    // 年月 + 施設名 + 実行ID の3条件で検索
    if (storedYearMonth === normalizedYearMonth &&
        storedFacility === normalizedFacility &&
        storedRunId === normalizedRunId) {
      const url = values[i][2] || '';
      const folderId = values[i][4] || '';
      return {
        rowIndex: i + 1,
        url,
        fileId: extractSpreadsheetId(url),
        folderId: folderId
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
 * @param {string} runId - 実行ID
 * @param {string} folderId - フォルダID
 * @returns {number} 登録された行番号
 */
function recordFacilityFile(sheet, facilityName, yearMonth, fileId, runId, folderId) {
  const normalizedFacility = normalizeString(facilityName);
  const normalizedYearMonth = normalizeString(yearMonth);
  const normalizedRunId = normalizeString(runId);
  const normalizedFolderId = normalizeString(folderId);
  const url = `https://docs.google.com/spreadsheets/d/${fileId}`;
  const rowIndex = sheet.getLastRow() + 1;
  sheet.getRange(rowIndex, 1, 1, 6).setValues([[
    normalizedYearMonth,
    normalizedFacility,
    url,
    normalizedRunId,
    normalizedFolderId,
    new Date()
  ]]);
  // 書き込みを即座に反映（continuePlanGenerationで再開時に確実に検索できるように）
  SpreadsheetApp.flush();
  logInfo('台帳に事業所ファイルを登録しました', {
    facilityName,
    yearMonth,
    runId,
    folderId,
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
  sheet.getRange(rowIndex, 6).setValue(new Date());  // 6列目: 最終更新
}
