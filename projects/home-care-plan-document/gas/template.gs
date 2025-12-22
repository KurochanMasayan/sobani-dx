/**
 * テンプレートシートと患者ファイルの生成
 *
 * 変更履歴:
 * - スプレッドシートの単位を「患者」から「年月+居宅支援事業所」に変更
 * - シート名を「訪問日」から「患者名+日付」に変更
 */
let FACILITY_SPREADSHEET_CACHE = {};

function processPatientVisits(patient, visitPlan, registrySheet, outputFolderId, yearMonth, runId) {
  // 居宅支援事業所ごと・年月ごと・実行IDごとにスプレッドシートを取得または作成
  const facilityName = normalizeFacilityName(patient.facility) || '事業所未設定';
  const spreadsheet = getOrCreateFacilitySpreadsheet(facilityName, yearMonth, registrySheet, outputFolderId, runId);

  visitPlan.forEach(visit => {
    const sheetName = buildPatientVisitSheetName(patient, visit);
    const sheet = duplicateTemplateSheet(spreadsheet, sheetName);
    fillTemplateSheet(sheet, patient, visit);
  });

  // 訪問シート追加後、デフォルトのシート（Sheet1など）を削除
  cleanupDefaultSheets(spreadsheet);
}

// 実行ID別フォルダのキャッシュ
let RUN_FOLDER_CACHE = {};

/**
 * 実行ID用のサブフォルダを取得または作成
 * @param {string} outputFolderId - 出力フォルダID
 * @param {string} runId - 実行ID
 * @returns {Folder} 実行ID用フォルダ
 */
function getOrCreateRunFolder(outputFolderId, runId) {
  if (RUN_FOLDER_CACHE[runId]) {
    return RUN_FOLDER_CACHE[runId];
  }

  const outputFolder = DriveApp.getFolderById(outputFolderId);
  // フォルダ名: 実行日時のみ（例: 2025-12-20-103000）
  const folderName = runId.split('_')[0].replace(/T/g, '-').replace(/:/g, '');

  // 既存フォルダを検索
  const existingFolders = outputFolder.getFoldersByName(folderName);
  if (existingFolders.hasNext()) {
    const folder = existingFolders.next();
    RUN_FOLDER_CACHE[runId] = folder;
    logInfo('既存の実行フォルダを使用します', { folderName, folderId: folder.getId() });
    return folder;
  }

  // 新規フォルダを作成
  const newFolder = outputFolder.createFolder(folderName);
  RUN_FOLDER_CACHE[runId] = newFolder;
  logInfo('実行フォルダを新規作成しました', { folderName, folderId: newFolder.getId() });
  return newFolder;
}

/**
 * 年月+居宅支援事業所+実行IDごとにスプレッドシートを取得または作成
 * @param {string} facilityName - 居宅支援事業所名
 * @param {string} yearMonth - 年月（例: 2025-12）
 * @param {Sheet} registrySheet - 台帳シート
 * @param {string} outputFolderId - 出力フォルダID
 * @param {string} runId - 実行ID（6分制限対応）
 * @returns {Spreadsheet} スプレッドシート
 */
function getOrCreateFacilitySpreadsheet(facilityName, yearMonth, registrySheet, outputFolderId, runId) {
  // キャッシュキーを作成（同一実行内で複数患者が同じ事業所に所属する場合のパフォーマンス向上）
  // 実行IDを含めることで、同一バッチ内でのみキャッシュを使用
  const cacheKey = `${runId}_${yearMonth}_${facilityName}`;
  if (FACILITY_SPREADSHEET_CACHE[cacheKey]) {
    return FACILITY_SPREADSHEET_CACHE[cacheKey];
  }

  // 既存のファイルを検索（年月 + 施設名 + 実行IDで検索）
  const existingRecord = lookupFacilityRecord(registrySheet, facilityName, yearMonth, runId);
  if (existingRecord) {
    logInfo('既存事業所ファイルを使用します', {
      facilityName,
      yearMonth,
      runId,
      fileId: existingRecord.fileId,
      folderId: existingRecord.folderId,
      rowIndex: existingRecord.rowIndex
    });
    updateFacilityLastTouched(registrySheet, existingRecord.rowIndex);
    if (!existingRecord.fileId) {
      throw new Error(`URL からファイルIDを取得できませんでした: facility=${facilityName}, yearMonth=${yearMonth}, runId=${runId}`);
    }
    // 台帳に保存されているフォルダIDをキャッシュに登録（他の施設ファイル作成時に使用）
    if (existingRecord.folderId) {
      RUN_FOLDER_CACHE[runId] = DriveApp.getFolderById(existingRecord.folderId);
    }
    const spreadsheet = SpreadsheetApp.openById(existingRecord.fileId);
    FACILITY_SPREADSHEET_CACHE[cacheKey] = spreadsheet;
    return spreadsheet;
  }

  // 実行ID用のサブフォルダを取得または作成
  const runFolder = getOrCreateRunFolder(outputFolderId, runId);

  // 新規スプレッドシートを作成
  const newFileName = sanitizeName(facilityName);

  const newSpreadsheet = SpreadsheetApp.create(newFileName);
  DriveApp.getFileById(newSpreadsheet.getId()).moveTo(runFolder);

  // 台帳に実行ID・フォルダIDも含めて登録
  const rowIndex = recordFacilityFile(registrySheet, facilityName, yearMonth, newSpreadsheet.getId(), runId, runFolder.getId());
  updateFacilityLastTouched(registrySheet, rowIndex);
  logInfo('事業所ファイルを新規作成しました', {
    facilityName,
    yearMonth,
    runId,
    fileId: newSpreadsheet.getId(),
    rowIndex,
    folderId: runFolder.getId()
  });
  FACILITY_SPREADSHEET_CACHE[cacheKey] = newSpreadsheet;
  return newSpreadsheet;
}

function duplicateTemplateSheet(spreadsheet, sheetName) {
  // マスターテンプレート（GASが動いているスプレッドシート）から常に最新の原本を取得
  const masterTemplate = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = masterTemplate.getSheetByName(BASE_CONFIG.sheetTemplateName) || masterTemplate.getSheets()[0];

  // 患者ファイルに最新のテンプレートをコピー
  const newSheet = templateSheet.copyTo(spreadsheet);
  newSheet.setName(generateUniqueSheetName(spreadsheet, sheetName));
  newSheet.activate();
  return newSheet;
}

function generateUniqueSheetName(spreadsheet, baseName) {
  let name = baseName;
  let counter = 1;
  while (spreadsheet.getSheetByName(name)) {
    counter += 1;
    name = `${baseName}_${counter}`;
  }
  return name;
}

/**
 * 患者名+日付のシート名を生成
 * @param {Object} patient - 患者情報
 * @param {Object} visit - 訪問情報
 * @returns {string} シート名（例: テスト香夕子_2025-12-01）
 */
function buildPatientVisitSheetName(patient, visit) {
  const patientName = sanitizeName(patient.name || '名無し').replace(/\s+/g, '');
  const visitDate = visit.dateText || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const rawName = `${patientName}_${visitDate}`;
  return rawName.replace(/[\\/?*\[\]]/g, '-');
}

function fillTemplateSheet(sheet, patient, visit) {
  const map = BASE_CONFIG.cellMap;
  const setField = (key, value) => {
    if (!map[key]) return;
    setValueEvenIfEmpty(sheet, map[key], value || '');
  };

  setField('createdDate', visit.dateText);
  setField('patientName', patient.name);
  setField('sex', patient.sex);
  setField('birthDate', patient.birth);
  setField('age', patient.age);
  setField('zipCode', patient.zip);
  setField('address', patient.address);
  setField('diagnosis', patient.diagnosis);
  setField('visitLocation', normalizeVisitLocation(patient.facility));
  setField('carePlan', patient.carePlan);
  setField('cautionNote', patient.caution);
  setField('bedriddenLevel', patient.bedridden);
  setField('dementiaLevel', patient.dementia);
  setField('careLevel', patient.careLevel);
}

function setValueEvenIfEmpty(sheet, rangeA1, value) {
  if (!rangeA1) return;
  sheet.getRange(rangeA1).setValue(value || '');
}

function sanitizeName(value) {
  return (value || '').replace(/[\\/:*?"<>|]/g, '_');
}

/**
 * 施設名を正規化
 * - 前後の空白を削除
 * - 「（在）」「（施）」などの接頭辞を削除
 * @param {string} facilityName - 施設名
 * @returns {string} 正規化された施設名
 */
function normalizeFacilityName(facilityName) {
  return (facilityName || '')
    .trim()
    .replace(/^（在）|^（施）|^\(在\)|^\(施\)/g, '')
    .trim();
}

/**
 * 診療場所の表記を正規化
 * - 「（在）」「（施）」などの接頭辞を削除
 * - 「個人宅」→「自宅」に変換
 * @param {string} location - 診療場所
 * @returns {string} 正規化された診療場所
 */
function normalizeVisitLocation(location) {
  // まず施設名の接頭辞を削除
  const normalized = normalizeFacilityName(location);
  if (normalized === '個人宅') {
    return '自宅';
  }
  return normalized;
}

function cleanupDefaultSheets(spreadsheet) {
  const allSheets = spreadsheet.getSheets();
  if (allSheets.length <= 1) return; // 最低1シート必要

  // 削除対象シートを先に抽出
  const sheetsToDelete = allSheets.filter(sheet => {
    const sheetName = sheet.getName();
    return sheetName.match(/^(Sheet|シート)\d*$/i) || sheetName === BASE_CONFIG.sheetTemplateName;
  });

  // 逆順で削除（配列の添字ズレを防ぐ）
  for (let i = sheetsToDelete.length - 1; i >= 0; i--) {
    if (spreadsheet.getSheets().length > 1) {
      spreadsheet.deleteSheet(sheetsToDelete[i]);
    }
  }
}
