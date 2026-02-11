/**
 * テンプレートシートと患者ファイルの生成
 *
 * - スプレッドシートの単位: 年月+居宅支援事業所
 * - シート名: 患者名+日付
 * - 同じ月の処理は同じファイルに追記される
 */
let FACILITY_SPREADSHEET_CACHE = {};

// 月別フォルダのキャッシュ
let MONTH_FOLDER_CACHE = {};

function processPatientVisits(patient, visitPlan, registrySheet, outputFolderId, yearMonth) {
  // 居宅支援事業所ごと・年月ごとにスプレッドシートを取得または作成
  const facilityName = normalizeFacilityName(patient.facility) || BASE_CONFIG.fallbackFacilityName;
  const spreadsheet = getOrCreateFacilitySpreadsheet(facilityName, yearMonth, registrySheet, outputFolderId);

  visitPlan.forEach(visit => {
    const sheetName = buildPatientVisitSheetName(patient, visit);
    const sheet = duplicateTemplateSheet(spreadsheet, sheetName);
    fillTemplateSheet(sheet, patient, visit);
  });

  // 訪問シート追加後、デフォルトのシート（Sheet1など）を削除
  cleanupDefaultSheets(spreadsheet);
}

/**
 * 月別のサブフォルダを取得または作成
 * @param {string} outputFolderId - 出力フォルダID
 * @param {string} yearMonth - 年月（例: 2025-12）
 * @returns {Folder} 月別フォルダ
 */
function getOrCreateMonthFolder(outputFolderId, yearMonth) {
  if (MONTH_FOLDER_CACHE[yearMonth]) {
    return MONTH_FOLDER_CACHE[yearMonth];
  }

  const outputFolder = DriveApp.getFolderById(outputFolderId);

  // 既存フォルダを検索
  const existingFolders = outputFolder.getFoldersByName(yearMonth);
  if (existingFolders.hasNext()) {
    const folder = existingFolders.next();
    MONTH_FOLDER_CACHE[yearMonth] = folder;
    logInfo('既存の月別フォルダを使用します', { yearMonth, folderId: folder.getId() });
    return folder;
  }

  // 新規フォルダを作成
  const newFolder = outputFolder.createFolder(yearMonth);
  MONTH_FOLDER_CACHE[yearMonth] = newFolder;
  logInfo('月別フォルダを新規作成しました', { yearMonth, folderId: newFolder.getId() });
  return newFolder;
}

/**
 * 年月+居宅支援事業所ごとにスプレッドシートを取得または作成
 * @param {string} facilityName - 居宅支援事業所名
 * @param {string} yearMonth - 年月（例: 2025-12）
 * @param {Sheet} registrySheet - 台帳シート
 * @param {string} outputFolderId - 出力フォルダID
 * @returns {Spreadsheet} スプレッドシート
 */
function getOrCreateFacilitySpreadsheet(facilityName, yearMonth, registrySheet, outputFolderId) {
  logInfo('getOrCreateFacilitySpreadsheet開始', { facilityName, yearMonth, outputFolderId });

  // キャッシュキーを作成（同一実行内で複数患者が同じ事業所に所属する場合のパフォーマンス向上）
  const cacheKey = `${yearMonth}_${facilityName}`;
  if (FACILITY_SPREADSHEET_CACHE[cacheKey]) {
    logInfo('キャッシュから取得', { cacheKey });
    return FACILITY_SPREADSHEET_CACHE[cacheKey];
  }

  // 既存のファイルを検索（年月 + 施設名で検索）
  logInfo('台帳検索開始', { facilityName, yearMonth });
  const existingRecord = lookupFacilityRecord(registrySheet, facilityName, yearMonth);
  if (existingRecord) {
    logInfo('既存事業所ファイルを使用します', {
      facilityName,
      yearMonth,
      fileId: existingRecord.fileId,
      url: existingRecord.url,
      rowIndex: existingRecord.rowIndex
    });
    updateFacilityLastTouched(registrySheet, existingRecord.rowIndex);
    if (!existingRecord.fileId) {
      throw new Error(`URL からファイルIDを取得できませんでした: facility=${facilityName}, yearMonth=${yearMonth}`);
    }
    try {
      const spreadsheet = SpreadsheetApp.openById(existingRecord.fileId);
      logInfo('既存ファイルを開きました', { fileId: existingRecord.fileId });
      FACILITY_SPREADSHEET_CACHE[cacheKey] = spreadsheet;
      return spreadsheet;
    } catch (openError) {
      logError('既存ファイルを開けませんでした。新規作成します', openError);
      // 開けない場合は新規作成にフォールバック（台帳のエントリは残るが、新しいファイルを作成）
    }
  }

  // 月別のサブフォルダを取得または作成
  logInfo('新規ファイル作成開始', { facilityName, yearMonth });
  const monthFolder = getOrCreateMonthFolder(outputFolderId, yearMonth);

  // 新規スプレッドシートを作成（ファイル名: 年月_施設名）
  const newFileName = `${yearMonth}_${sanitizeName(facilityName)}`;
  logInfo('スプレッドシート作成', { newFileName });

  const newSpreadsheet = SpreadsheetApp.create(newFileName);
  DriveApp.getFileById(newSpreadsheet.getId()).moveTo(monthFolder);

  // 台帳に登録
  const rowIndex = recordFacilityFile(registrySheet, facilityName, yearMonth, newSpreadsheet.getId());
  updateFacilityLastTouched(registrySheet, rowIndex);
  logInfo('事業所ファイルを新規作成しました', {
    facilityName,
    yearMonth,
    fileId: newSpreadsheet.getId(),
    rowIndex
  });
  FACILITY_SPREADSHEET_CACHE[cacheKey] = newSpreadsheet;
  return newSpreadsheet;
}


function duplicateTemplateSheet(spreadsheet, sheetName) {
  // マスターテンプレート（GASが動いているスプレッドシート）から常に最新の原本を取得
  const masterTemplate = getActiveSpreadsheetOrThrow();
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
  setField('facilityName', patient.facility);
  setField('patientName', patient.name);
  setField('sex', patient.sex);
  setField('birthDate', patient.birth);
  setField('age', patient.age);
  setField('zipCode', patient.zip);
  setField('address', patient.address);
  setField('visitDate', visit.dateText);
  setField('diagnosis', patient.diagnosis);
  setField('visitNote', getVisitNote(patient, visit));
  setField('careLevel', patient.careLevel);
  setField('bedriddenLevel', patient.bedridden);
  setField('dementiaLevel', patient.dementia);
  setField('cautionNote', patient.caution);
}

function setValueEvenIfEmpty(sheet, rangeA1, value) {
  if (!rangeA1) return;
  sheet.getRange(rangeA1).setValue(value || '');
}

function getVisitNote(patient, visit) {
  if (!patient.visitText) return '';
  return patient.visitText[visit.sequence] || '';
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

