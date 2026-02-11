/**
 * 居宅療養管理指導システム - リファクタリング版
 *
 * 主な改善点:
 * - tmpシート不要（メモリ内処理）
 * - 単一関数で完結
 * - 設定の一元管理
 * - 最新API使用
 * - エラーハンドリング強化
 */

// =====================================
// 設定管理
// =====================================

const CONFIG = {
  // フォルダID設定
  folders: {
    // 月次データ投入先（モバカルCSV）
    monthData: PropertiesService.getScriptProperties().getProperty('MONTH_DATA_FOLDER_ID'),

    // 患者マスタCSV
    patientMaster: PropertiesService.getScriptProperties().getProperty('FIXDATA_FOLDER_ID'),

    // メイン出力先（患者別ファイル）
    output: '1wxFMS3SmC03PqvPk9WJ-Qi-QBOV6jB0O',

    // 施設別整理先（ショートカット）
    organized: PropertiesService.getScriptProperties().getProperty('SHORT_FOLDER_ID'),

    // PDF出力先
    pdfOutput: '1r0CfXFfsFozN_Mf0BtICGJgccdwY2lrq'
  },

  // テンプレート設定
  template: {
    id: '1JLVlR1JnbRrY3Jpc7I0iuGOAmVu3gl_tLuJwo5z8HFo',
    cellMapping: {
      visitDate: 'O2',
      facilityName: 'H4',
      patientName: 'G7',
      sex: 'T7',
      dateOfBirth: 'G8',
      age: 'R8',
      zipCode: 'H9',
      address: 'G10',
      visitDate2: 'F11',
      diagnosis: 'F13',
      caregiving: 'G29',
      bedridden: 'G30',
      dementia: 'G31'
    }
  },

  // CSV設定
  csv: {
    encoding: 'Shift_JIS',
    monthDataColumns: {
      patientId: 'orcaID',
      date: '日付',
      status: '定期訪問'
    },
    patientMasterColumns: {
      patientId: '患者ID',
      patientName: '患者名',
      dateOfBirth: '生年月日',
      age: '年齢',
      sex: '性別',
      zipCode: '郵便番号',
      address: '住所',
      diagnosis: '診断名',
      facilityName: '居宅介護支援事業所',
      bedridden: '寝たきり度',
      dementia: '認知度',
      caregiving: '介護認定'
    }
  }
};

// =====================================
// ユーティリティ関数
// =====================================

/**
 * エラーログを記録
 */
function logError(functionName, error, context = {}) {
  const message = `[ERROR] ${functionName}: ${error.message || error.toString()}`;
  Logger.log(message);
  Logger.log('Context:', JSON.stringify(context));
  return message;
}

/**
 * 情報ログを記録
 */
function logInfo(message, data = null) {
  Logger.log(`[INFO] ${message}`);
  if (data) {
    Logger.log(JSON.stringify(data));
  }
}

/**
 * ローマ数字変換（文字化け対策）
 */
function convertToRomanNumerals(text) {
  if (!text) return text;
  return text
    .replace(/Ⅳ/g, 'Ⅳ')
    .replace(/Ⅲ/g, 'Ⅲ')
    .replace(/Ⅱ/g, 'Ⅱ')
    .replace(/Ⅰ/g, 'Ⅰ')
    .replace(/�W/g, 'Ⅳ')
    .replace(/�V/g, 'Ⅲ')
    .replace(/�U/g, 'Ⅱ')
    .replace(/�T/g, 'Ⅰ');
}

/**
 * 日付フォーマット
 */
function formatDate(date) {
  return Utilities.formatDate(new Date(date), 'Asia/Tokyo', 'yyyy-MM-dd');
}

// =====================================
// データ取得モジュール
// =====================================

/**
 * CSVファイルを読み込む
 * @param {string} folderId - フォルダID
 * @param {boolean} deleteAfterRead - 読み込み後に削除するか
 * @returns {Array} CSVデータの配列
 */
function readCSVFiles(folderId, deleteAfterRead = false) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.CSV);
    const allCsvData = [];

    if (!files.hasNext()) {
      logInfo('CSVファイルが見つかりません', { folderId });
      return [];
    }

    while (files.hasNext()) {
      const file = files.next();
      try {
        const csvContent = file.getBlob().getDataAsString(CONFIG.csv.encoding);
        const csvData = Utilities.parseCsv(csvContent);

        if (csvData.length > 1) {
          allCsvData.push(...csvData);
        }

        if (deleteAfterRead) {
          file.setTrashed(true);
          logInfo('CSVファイルを削除しました', { fileName: file.getName() });
        }
      } catch (e) {
        logError('readCSVFiles', e, { fileName: file.getName() });
      }
    }

    return allCsvData;

  } catch (e) {
    logError('readCSVFiles', e, { folderId });
    return [];
  }
}

/**
 * 月次データから定期訪問のデータを抽出
 * @param {Array} monthData - 月次データ
 * @returns {Array} 抽出されたデータ
 */
function extractRegularVisits(monthData) {
  try {
    if (!Array.isArray(monthData) || monthData.length === 0) {
      logInfo('有効なデータがありません');
      return [];
    }

    const header = monthData[0];
    const cols = CONFIG.csv.monthDataColumns;
    const patientIdIndex = header.indexOf(cols.patientId);
    const dateIndex = header.indexOf(cols.date);
    const statusIndex = header.indexOf(cols.status);

    if (patientIdIndex === -1 || dateIndex === -1 || statusIndex === -1) {
      logError('extractRegularVisits', '必要な列が見つかりません', { header });
      return [];
    }

    const extractedData = [];
    monthData.slice(1).forEach(row => {
      if (row && row[statusIndex] === '定期訪問') {
        extractedData.push({
          date: row[dateIndex],
          patientId: row[patientIdIndex]
        });
      }
    });

    logInfo('定期訪問データを抽出しました', { count: extractedData.length });
    return extractedData;

  } catch (e) {
    logError('extractRegularVisits', e);
    return [];
  }
}

/**
 * 患者マスタと月次データを統合
 * @param {Array} monthExtractedData - 抽出された月次データ
 * @param {Array} patientMasterData - 患者マスタデータ
 * @returns {Array} 統合されたデータ
 */
function mergePatientData(monthExtractedData, patientMasterData) {
  try {
    const dataMap = {};
    const header = patientMasterData[0];
    const cols = CONFIG.csv.patientMasterColumns;

    // 列インデックスを取得
    const indices = {};
    Object.keys(cols).forEach(key => {
      indices[key] = header.indexOf(cols[key]);
    });

    // 患者マスタをマップに格納
    patientMasterData.slice(1).forEach(row => {
      const patientId = row[indices.patientId];
      if (!patientId) return;

      dataMap[patientId] = {
        patientId,
        patientName: row[indices.patientName],
        dateOfBirth: row[indices.dateOfBirth],
        age: row[indices.age],
        sex: row[indices.sex],
        zipCode: row[indices.zipCode],
        address: row[indices.address],
        diagnosis: row[indices.diagnosis],
        facilityName: row[indices.facilityName],
        bedridden: row[indices.bedridden],
        dementia: convertToRomanNumerals(row[indices.dementia]),
        caregiving: row[indices.caregiving],
        dates: []
      };
    });

    // 月次データの日付を追加
    monthExtractedData.forEach(data => {
      const patientId = data.patientId;
      if (dataMap[patientId]) {
        dataMap[patientId].dates.push(data.date);
      } else {
        // マスタにない患者の場合
        logInfo('マスタに存在しない患者', { patientId });
      }
    });

    const mergedData = Object.values(dataMap).filter(patient => patient.dates.length > 0);
    logInfo('データ統合完了', { patientCount: mergedData.length });

    return mergedData;

  } catch (e) {
    logError('mergePatientData', e);
    return [];
  }
}

// =====================================
// ファイル生成モジュール
// =====================================

/**
 * スプレッドシートを取得または作成
 * @param {string} fileName - ファイル名
 * @param {string} folderId - フォルダID
 * @returns {Object} スプレッドシートオブジェクト
 */
function getOrCreateSpreadsheet(fileName, folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByName(fileName);

    if (files.hasNext()) {
      // 既存ファイルを開く
      const file = files.next();
      return SpreadsheetApp.openById(file.getId());
    } else {
      // 新規ファイルを作成
      const templateFile = DriveApp.getFileById(CONFIG.template.id);
      const newFile = templateFile.makeCopy(fileName, folder);
      logInfo('新規ファイルを作成しました', { fileName });
      return SpreadsheetApp.openById(newFile.getId());
    }
  } catch (e) {
    logError('getOrCreateSpreadsheet', e, { fileName, folderId });
    throw e;
  }
}

/**
 * シートに患者情報を入力
 * @param {Sheet} sheet - シートオブジェクト
 * @param {Object} patientData - 患者データ
 * @param {string} visitDate - 訪問日
 */
function fillPatientData(sheet, patientData, visitDate) {
  try {
    const mapping = CONFIG.template.cellMapping;
    const cells = [
      mapping.visitDate,
      mapping.facilityName,
      mapping.patientName,
      mapping.sex,
      mapping.dateOfBirth,
      mapping.age,
      mapping.zipCode,
      mapping.address,
      mapping.visitDate2,
      mapping.diagnosis,
      mapping.caregiving,
      mapping.bedridden,
      mapping.dementia
    ];

    const range = sheet.getRangeList(cells);
    const ranges = range.getRanges();

    ranges[0].setValue(visitDate);
    ranges[1].setValue(patientData.facilityName);
    ranges[2].setValue(patientData.patientName);
    ranges[3].setValue(patientData.sex);
    ranges[4].setValue(patientData.dateOfBirth);
    ranges[5].setValue(patientData.age);
    ranges[6].setValue(patientData.zipCode);
    ranges[7].setValue(patientData.address);
    ranges[8].setValue(visitDate);
    ranges[9].setValue(patientData.diagnosis);
    ranges[10].setValue(patientData.caregiving);
    ranges[11].setValue(patientData.bedridden);
    ranges[12].setValue(patientData.dementia);

  } catch (e) {
    logError('fillPatientData', e, { patientName: patientData.patientName, visitDate });
    throw e;
  }
}

/**
 * 患者データからスプレッドシートを生成
 * @param {Array} patientsData - 患者データの配列
 */
function generateSpreadsheets(patientsData) {
  try {
    logInfo('スプレッドシート生成開始', { patientCount: patientsData.length });

    patientsData.forEach((patient, index) => {
      try {
        const fileName = `${patient.facilityName}-${patient.patientName}`;
        const spreadsheet = getOrCreateSpreadsheet(fileName, CONFIG.folders.output);

        // 各訪問日でシートを作成
        patient.dates.forEach(date => {
          const formattedDate = formatDate(date);

          // シートが既に存在するかチェック
          if (spreadsheet.getSheetByName(formattedDate)) {
            logInfo('シートは既に存在します', { fileName, date: formattedDate });
            return;
          }

          // テンプレートシートをコピー
          const templateSheet = spreadsheet.getSheets()[0];
          const newSheet = templateSheet.copyTo(spreadsheet).setName(formattedDate);

          // 患者情報を入力
          fillPatientData(newSheet, patient, formattedDate);

          logInfo('シートを作成しました', { fileName, date: formattedDate });
        });

        logInfo(`患者処理完了 (${index + 1}/${patientsData.length})`, { patientName: patient.patientName });

      } catch (e) {
        logError('generateSpreadsheets', e, { patientName: patient.patientName });
      }
    });

    logInfo('スプレッドシート生成完了');

  } catch (e) {
    logError('generateSpreadsheets', e);
    throw e;
  }
}

// =====================================
// フォルダ整理モジュール
// =====================================

/**
 * 施設別フォルダにショートカットを作成
 * @param {string} outputFolderId - 出力フォルダID
 * @param {string} organizedFolderId - 整理先フォルダID
 */
function organizeByFacility(outputFolderId, organizedFolderId) {
  try {
    logInfo('フォルダ整理開始');

    const outputFolder = DriveApp.getFolderById(outputFolderId);
    const organizedFolder = DriveApp.getFolderById(organizedFolderId);
    const files = outputFolder.getFilesByType(MimeType.GOOGLE_SHEETS);

    if (!files.hasNext()) {
      logInfo('整理対象のファイルがありません');
      return;
    }

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();

      // ファイル名から施設名を取得
      const facilityName = fileName.split('-')[0] || '自宅';

      // 施設フォルダを取得または作成
      const facilityFolders = organizedFolder.getFoldersByName(facilityName);
      let facilityFolder;

      if (facilityFolders.hasNext()) {
        facilityFolder = facilityFolders.next();
      } else {
        facilityFolder = organizedFolder.createFolder(facilityName);
        logInfo('施設フォルダを作成しました', { facilityName });
      }

      // ショートカット名
      const shortcutName = `${fileName} ショートカット`;

      // 既存のショートカットチェック
      const existingShortcuts = facilityFolder.getFilesByName(shortcutName);
      if (existingShortcuts.hasNext()) {
        logInfo('ショートカットは既に存在します', { shortcutName });
        continue;
      }

      // ショートカットを作成して移動
      try {
        DriveApp.createShortcut(file.getId())
          .setName(shortcutName)
          .moveTo(facilityFolder);

        logInfo('ショートカットを作成しました', { facilityName, fileName });
      } catch (e) {
        logError('organizeByFacility - shortcut creation', e, { fileName });
      }
    }

    logInfo('フォルダ整理完了');

  } catch (e) {
    logError('organizeByFacility', e);
    throw e;
  }
}

// =====================================
// メイン処理関数
// =====================================

/**
 * 月次データ処理のメイン関数
 * @param {Object} options - オプション設定
 */
function processMonthlyData(options = {}) {
  const startTime = new Date();
  logInfo('=== 月次データ処理開始 ===');

  try {
    // デフォルトオプション
    const opts = {
      organizeFiles: true,  // ファイル整理を行うか
      deleteCSV: true,      // CSVを削除するか
      ...options
    };

    // 1. 月次データ読み込み
    logInfo('ステップ1: 月次データ読み込み');
    const monthData = readCSVFiles(CONFIG.folders.monthData, opts.deleteCSV);
    if (monthData.length === 0) {
      logInfo('処理対象のデータがありません');
      return;
    }

    // 2. 定期訪問データ抽出
    logInfo('ステップ2: 定期訪問データ抽出');
    const extractedData = extractRegularVisits(monthData);
    if (extractedData.length === 0) {
      logInfo('定期訪問データがありません');
      return;
    }

    // 3. 患者マスタ読み込み
    logInfo('ステップ3: 患者マスタ読み込み');
    const patientMaster = readCSVFiles(CONFIG.folders.patientMaster, false);
    if (patientMaster.length === 0) {
      logError('processMonthlyData', '患者マスタが見つかりません');
      return;
    }

    // 4. データ統合
    logInfo('ステップ4: データ統合');
    const mergedData = mergePatientData(extractedData, patientMaster);
    if (mergedData.length === 0) {
      logInfo('統合データがありません');
      return;
    }

    // 5. スプレッドシート生成
    logInfo('ステップ5: スプレッドシート生成');
    generateSpreadsheets(mergedData);

    // 6. ファイル整理（オプション）
    if (opts.organizeFiles && CONFIG.folders.organized) {
      logInfo('ステップ6: ファイル整理');
      organizeByFacility(CONFIG.folders.output, CONFIG.folders.organized);
    }

    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    logInfo('=== 月次データ処理完了 ===', {
      duration: `${duration}秒`,
      patientCount: mergedData.length
    });

  } catch (e) {
    logError('processMonthlyData', e);
    throw e;
  }
}

// =====================================
// テスト・ユーティリティ関数
// =====================================

/**
 * 設定確認用関数
 */
function checkConfiguration() {
  logInfo('=== 設定確認 ===');

  const checks = {
    monthDataFolder: CONFIG.folders.monthData,
    patientMasterFolder: CONFIG.folders.patientMaster,
    outputFolder: CONFIG.folders.output,
    organizedFolder: CONFIG.folders.organized,
    templateId: CONFIG.template.id
  };

  Object.keys(checks).forEach(key => {
    const value = checks[key];
    const status = value ? '✓' : '✗';
    logInfo(`${status} ${key}: ${value || '未設定'}`);
  });
}

/**
 * ドライ実行（ファイル生成なし）
 */
function dryRun() {
  logInfo('=== ドライ実行開始 ===');

  try {
    // データ読み込みと統合のみ
    const monthData = readCSVFiles(CONFIG.folders.monthData, false);
    const extractedData = extractRegularVisits(monthData);
    const patientMaster = readCSVFiles(CONFIG.folders.patientMaster, false);
    const mergedData = mergePatientData(extractedData, patientMaster);

    logInfo('処理対象データ', {
      monthDataRows: monthData.length,
      regularVisits: extractedData.length,
      patients: mergedData.length
    });

    // サンプル表示
    if (mergedData.length > 0) {
      logInfo('サンプルデータ（1件目）', mergedData[0]);
    }

  } catch (e) {
    logError('dryRun', e);
  }
}
