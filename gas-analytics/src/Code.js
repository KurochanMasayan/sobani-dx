/**
 * 訪問診療データ分析 - メインスクリプト
 *
 * 機能：
 * 1. 患者情報シート: CSVから患者情報を上書き更新
 * 2. 履歴シート: 要介護度・在/施データを追記（履歴管理）
 */

// ===== 設定 =====
const CONFIG = {
  // 各CSVファイルが保存されているフォルダID（Google Drive）
  FOLDER_IDS: {
    PATIENT_INFO: '1Mp-LiUij5_4wqoAJQPxwBw23yPwQipre',      // 患者情報CSV
    CARE_LEVEL: '1eXZDz7CzsyzsQ6jF0KcnsO0tz5RZeFqz',          // 要介護度CSV
    FACILITY_TYPE: '1OwP3XmBdDPGL4Dvy_lxZEflMEcQ0Id58',  // 在医総施医総CSV
    INSURANCE_CLAIM: '185TNcOYXmbL0-Szffgu1DH99fYR_yyHu' // 保険請求リストCSV
  },

  // 出力先スプレッドシートID（空の場合は新規作成）
  SPREADSHEET_ID: '1ycvmjzLNYzxY0OLuKZSL9Z66do4J-CakSqp1zVngjnI',

  // シート名
  SHEET_NAMES: {
    PATIENT_INFO: '患者情報',
    HISTORY: '履歴'
  },

  // CSVカラム名マッピング
  CSV_COLUMNS: {
    PATIENT_INFO: {
      PATIENT_ID: '患者ID',
      PATIENT_NAME: '患者名',
      PATIENT_NAME_KANA: '患者名カナ',
      FACILITY_NAME: '施設名'
    },
    CARE_LEVEL: {
      PATIENT_ID: '利用者コード',
      CARE_LEVEL: '要介護度'
    },
    FACILITY_TYPE: {
      PATIENT_ID: '患者ID',
      MANAGEMENT_FEE: '管理料設定（総合管理料）'
    },
    INSURANCE_CLAIM: {
      BILLING_MONTH: '請求年月',
      PATIENT_ID: '患者番号',
      PATIENT_NAME: '患者氏名',
      POINTS: '点数'
    }
  }
};

/**
 * メニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('訪問診療データ')
    .addItem('患者情報を更新', 'updatePatientInfo')
    .addItem('履歴を追加', 'addHistory')
    .addSeparator()
    .addItem('すべて実行', 'runAll')
    .addToUi();
}

/**
 * すべての処理を実行
 */
function runAll() {
  updatePatientInfo();
  addHistory();
  SpreadsheetApp.getActiveSpreadsheet().toast('すべての処理が完了しました', '完了', 5);
}

/**
 * 患者情報シートを上書き更新
 */
function updatePatientInfo() {
  const ss = getOrCreateSpreadsheet();
  const sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAMES.PATIENT_INFO);

  // 患者情報CSVファイルを取得
  const patientCsvFile = getCsvFileFromFolder(CONFIG.FOLDER_IDS.PATIENT_INFO);
  if (!patientCsvFile) {
    SpreadsheetApp.getActiveSpreadsheet().toast('患者情報CSVファイルが見つかりません', 'エラー', 5);
    return;
  }

  // 患者情報CSVを解析
  const patientData = parseCsvFile(patientCsvFile, 'Shift_JIS');
  if (patientData.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast('患者情報CSVデータが空です', 'エラー', 5);
    return;
  }

  // カラムインデックスマップを作成
  const cols = CONFIG.CSV_COLUMNS.PATIENT_INFO;
  const columnMap = getColumnIndexMap(patientData[0]);

  // ヘッダー行を設定
  const headers = ['患者ID', '患者名', '患者名カナ', '施設名', '更新日'];

  // データを変換
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const outputData = [headers];

  // CSVの1行目はヘッダーなのでスキップ
  for (let i = 1; i < patientData.length; i++) {
    const row = patientData[i];

    const patientId = String(getColumnValue(row, columnMap, cols.PATIENT_ID)).trim();
    const patientName = getColumnValue(row, columnMap, cols.PATIENT_NAME);
    const patientNameKana = getColumnValue(row, columnMap, cols.PATIENT_NAME_KANA);
    const facilityName = getColumnValue(row, columnMap, cols.FACILITY_NAME);

    if (!patientId) continue;

    outputData.push([patientId, patientName, patientNameKana, facilityName, timestamp]);
  }

  // シートをクリアして書き込み
  sheet.clear();
  if (outputData.length > 0) {
    sheet.getRange(1, 1, outputData.length, headers.length).setValues(outputData);

    // ヘッダー行の書式設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4a86e8');
    headerRange.setFontColor('#ffffff');

  }

  // 処理完了後、使用済みCSVファイルを削除
  deleteProcessedFile(patientCsvFile);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `患者情報を${outputData.length - 1}件更新しました`,
    '完了',
    5
  );
}

/**
 * 履歴シートにデータを追加
 */
function addHistory() {
  const ss = getOrCreateSpreadsheet();
  const sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAMES.HISTORY);

  // 保険請求リストCSVファイルを取得（ベースデータ）
  const insuranceClaimCsvFile = getCsvFileFromFolder(CONFIG.FOLDER_IDS.INSURANCE_CLAIM);
  if (!insuranceClaimCsvFile) {
    SpreadsheetApp.getActiveSpreadsheet().toast('保険請求リストCSVファイルが見つかりません', 'エラー', 5);
    return;
  }

  // 要介護度CSVファイルを取得（必須）
  const careLevelCsvFile = getCsvFileFromFolder(CONFIG.FOLDER_IDS.CARE_LEVEL);
  if (!careLevelCsvFile) {
    SpreadsheetApp.getActiveSpreadsheet().toast('要介護度CSVファイルが見つかりません', 'エラー', 5);
    return;
  }

  // 在医総施医総CSVファイルを取得（必須）
  const facilityTypeCsvFile = getCsvFileFromFolder(CONFIG.FOLDER_IDS.FACILITY_TYPE);
  if (!facilityTypeCsvFile) {
    SpreadsheetApp.getActiveSpreadsheet().toast('在医総施医総CSVファイルが見つかりません', 'エラー', 5);
    return;
  }

  // 要介護度マップを作成（患者ID → 要介護度）
  const careLevelCols = CONFIG.CSV_COLUMNS.CARE_LEVEL;
  const careLevelMap = {};
  const careLevelData = parseCsvFile(careLevelCsvFile, 'UTF-8');
  const careLevelColumnMap = getColumnIndexMap(careLevelData[0]);
  for (let i = 1; i < careLevelData.length; i++) {
    const row = careLevelData[i];
    const patientId = String(getColumnValue(row, careLevelColumnMap, careLevelCols.PATIENT_ID)).trim();
    const careLevel = getColumnValue(row, careLevelColumnMap, careLevelCols.CARE_LEVEL);
    if (patientId) {
      careLevelMap[patientId] = careLevel;
    }
  }

  // 在/施マップを作成（患者ID → 在/施）
  const facilityTypeCols = CONFIG.CSV_COLUMNS.FACILITY_TYPE;
  const facilityTypeMap = {};
  const facilityTypeData = parseCsvFile(facilityTypeCsvFile, 'Shift_JIS');
  const facilityTypeColumnMap = getColumnIndexMap(facilityTypeData[0]);
  for (let i = 1; i < facilityTypeData.length; i++) {
    const row = facilityTypeData[i];
    const patientId = String(getColumnValue(row, facilityTypeColumnMap, facilityTypeCols.PATIENT_ID)).trim();
    const managementFee = getColumnValue(row, facilityTypeColumnMap, facilityTypeCols.MANAGEMENT_FEE);
    // 1文字目を取得（「在重」→「在」、「施２」→「施」）
    if (patientId) {
      facilityTypeMap[patientId] = managementFee ? managementFee.charAt(0) : '';
    }
  }

  // 保険請求リストCSVを解析（Shift-JIS）
  const insuranceClaimData = parseCsvFile(insuranceClaimCsvFile, 'Shift_JIS');
  if (insuranceClaimData.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast('保険請求リストCSVデータが空です', 'エラー', 5);
    return;
  }

  // カラムインデックスマップを作成
  const insuranceCols = CONFIG.CSV_COLUMNS.INSURANCE_CLAIM;
  const insuranceColumnMap = getColumnIndexMap(insuranceClaimData[0]);

  // 同じ患者・同じ月のデータを集約するためのマップを作成
  // キー: 請求年月_患者ID、値: { billingMonth, patientId, patientName, totalPoints }
  const aggregatedDataMap = {};
  let aggregatedCount = 0; // 集約された（重複していた）レコード数

  // CSVの1行目はヘッダーなのでスキップ
  for (let i = 1; i < insuranceClaimData.length; i++) {
    const row = insuranceClaimData[i];

    const billingMonth = getColumnValue(row, insuranceColumnMap, insuranceCols.BILLING_MONTH);
    const patientIdRaw = String(getColumnValue(row, insuranceColumnMap, insuranceCols.PATIENT_ID)).trim();
    // ゼロパディングを除去して患者IDを正規化（000797 → 797）
    const patientId = patientIdRaw ? String(parseInt(patientIdRaw, 10)) : '';
    const patientName = getColumnValue(row, insuranceColumnMap, insuranceCols.PATIENT_NAME);
    const pointsRaw = getColumnValue(row, insuranceColumnMap, insuranceCols.POINTS);
    const points = parseInt(pointsRaw, 10) || 0;

    if (!patientId) continue;

    const key = `${billingMonth}_${patientId}`;

    if (aggregatedDataMap[key]) {
      // 既に同じ患者・同じ月のデータがある場合は点数を合計
      aggregatedDataMap[key].totalPoints += points;
      aggregatedCount++;
    } else {
      // 新規の患者・月の組み合わせ
      aggregatedDataMap[key] = {
        billingMonth: billingMonth,
        patientId: patientId,
        patientName: patientName,
        totalPoints: points
      };
    }
  }

  // 集約結果をログ出力
  if (aggregatedCount > 0) {
    Logger.log(`同一患者・同一月のデータを集約: ${aggregatedCount}件のレコードが合計されました`);
  }

  // ヘッダー行を確認・設定
  const headers = ['該当月', '患者ID', '患者名', '要介護度', '在/施', '診療報酬点数', '記入日'];
  const lastRow = sheet.getLastRow();

  // 既存データを読み込んで「該当月+患者ID」→行番号のマップを作成
  const existingDataMap = {};
  if (lastRow > 1) {
    const existingData = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // 該当月と患者IDのみ
    for (let i = 0; i < existingData.length; i++) {
      const key = `${existingData[i][0]}_${existingData[i][1]}`; // 該当月_患者ID
      existingDataMap[key] = i + 2; // 行番号（1始まり、ヘッダー除く）
    }
  }

  if (lastRow === 0) {
    // 新規シートの場合、ヘッダーを追加
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // ヘッダー行の書式設定
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#6aa84f');
    headerRange.setFontColor('#ffffff');
  }

  // データを変換
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const newData = [];
  const updateData = []; // 上書き用: { row: 行番号, data: データ配列 }

  // 集約されたデータを処理
  for (const key in aggregatedDataMap) {
    const data = aggregatedDataMap[key];
    const careLevel = careLevelMap[data.patientId] || '';
    const facilityType = facilityTypeMap[data.patientId] || '';

    const rowData = [
      data.billingMonth,
      data.patientId,
      data.patientName,
      careLevel,
      facilityType,
      data.totalPoints,
      timestamp
    ];

    if (existingDataMap[key]) {
      // 既存データがある場合は上書き
      updateData.push({ row: existingDataMap[key], data: rowData });
    } else {
      // 新規データは追加
      newData.push(rowData);
    }
  }

  // 既存データを上書き
  for (const update of updateData) {
    sheet.getRange(update.row, 1, 1, headers.length).setValues([update.data]);
  }

  // 新規データを末尾に追加
  if (newData.length > 0) {
    const startRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(startRow, 1, newData.length, headers.length).setValues(newData);
  }

  // 処理完了後、使用済みCSVファイルを削除
  deleteProcessedFile(insuranceClaimCsvFile);
  deleteProcessedFile(careLevelCsvFile);
  deleteProcessedFile(facilityTypeCsvFile);

  // 完了メッセージの組み立て
  let message = `履歴: 新規${newData.length}件、上書き${updateData.length}件`;
  if (aggregatedCount > 0) {
    message += ` (同一患者の${aggregatedCount}件を合計)`;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(message, '完了', 5);
}

/**
 * スプレッドシートを取得または作成
 */
function getOrCreateSpreadsheet() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * シートを取得または作成
 */
function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * 指定フォルダ内の唯一のCSVファイルを取得
 * @param {string} folderId - Google DriveフォルダID
 * @returns {File|null} - CSVファイル、またはエラー時はnull
 */
function getCsvFileFromFolder(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();

  const csvFiles = [];
  while (files.hasNext()) {
    const file = files.next();
    if (isCsvFile(file)) {
      csvFiles.push(file);
    }
  }

  if (csvFiles.length === 0) {
    return null;
  }

  if (csvFiles.length > 1) {
    throw new Error(`フォルダ内に複数のCSVファイルがあります（${csvFiles.length}件）。1つだけにしてください。`);
  }

  return csvFiles[0];
}

/**
 * ファイルがCSVかどうかを判定
 * @param {File} file - Google Driveファイル
 * @returns {boolean} - CSVファイルの場合true
 */
function isCsvFile(file) {
  const mimeType = file.getMimeType();
  const fileName = file.getName();

  // MIMEタイプまたは拡張子で判定
  return mimeType === 'text/csv' ||
         mimeType === 'text/plain' ||
         mimeType === 'application/vnd.ms-excel' ||
         fileName.endsWith('.csv');
}

/**
 * CSVファイルを解析
 */
function parseCsvFile(file, encoding) {
  const blob = file.getBlob();
  let content;

  if (encoding === 'Shift_JIS') {
    // Shift-JISの場合
    content = blob.getDataAsString('Shift_JIS');
  } else {
    // UTF-8の場合（BOM除去）
    content = blob.getDataAsString('UTF-8');
    if (content.charCodeAt(0) === 0xFEFF) {
      content = content.substring(1);
    }
  }

  return Utilities.parseCsv(content);
}

/**
 * ヘッダー行からカラム名のインデックスマップを作成
 * @param {Array} headerRow - CSVのヘッダー行
 * @returns {Object} - カラム名 → インデックスのマップ
 */
function getColumnIndexMap(headerRow) {
  const map = {};
  for (let i = 0; i < headerRow.length; i++) {
    map[headerRow[i].trim()] = i;
  }
  return map;
}

/**
 * カラム名からデータを取得
 * @param {Array} row - データ行
 * @param {Object} columnMap - カラムインデックスマップ
 * @param {string} columnName - 取得したいカラム名
 * @returns {string} - カラムの値（見つからない場合は空文字）
 */
function getColumnValue(row, columnMap, columnName) {
  const index = columnMap[columnName];
  if (index === undefined || index >= row.length) {
    return '';
  }
  return row[index];
}

/**
 * 処理済みCSVファイルを削除（ゴミ箱へ移動）
 * @param {File} file - 削除するGoogle Driveファイル
 */
function deleteProcessedFile(file) {
  if (file) {
    const fileName = file.getName();
    file.setTrashed(true);
    Logger.log(`処理済みファイルを削除しました: ${fileName}`);
  }
}

