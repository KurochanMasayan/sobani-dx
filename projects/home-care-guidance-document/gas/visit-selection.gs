/**
 * 訪問選択シートの管理
 * 3件以上の訪問がある患者について、どの訪問回を書類①②に使用するか選択する
 */

/**
 * 訪問選択シートを取得または作成
 * @returns {Sheet} 訪問選択シート
 */
function getVisitSelectionSheet() {
  const active = getActiveSpreadsheetOrThrow();
  let sheet = active.getSheetByName(BASE_CONFIG.visitSelectionSheetName);
  if (!sheet) {
    sheet = active.insertSheet(BASE_CONFIG.visitSelectionSheetName);
    setupVisitSelectionSheetHeader(sheet);
  }
  return sheet;
}

/**
 * 訪問選択シートのヘッダーを設定
 * @param {Sheet} sheet - シート
 */
function setupVisitSelectionSheetHeader(sheet) {
  const headers = [
    '患者ID',
    '患者名',
    '施設名',
    '訪問回数',
    '訪問日一覧',
    '書類①に使用する回',
    '書類②に使用する回',
    'ステータス',
    '備考'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  // ヘッダーの書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4a86e8');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // 列幅の設定
  sheet.setColumnWidth(1, 100);  // 患者ID
  sheet.setColumnWidth(2, 120);  // 患者名
  sheet.setColumnWidth(3, 180);  // 施設名
  sheet.setColumnWidth(4, 80);   // 訪問回数
  sheet.setColumnWidth(5, 300);  // 訪問日一覧
  sheet.setColumnWidth(6, 140);  // 書類①
  sheet.setColumnWidth(7, 140);  // 書類②
  sheet.setColumnWidth(8, 100);  // ステータス
  sheet.setColumnWidth(9, 200);  // 備考
}

/**
 * 3件以上の訪問がある患者を訪問選択シートに登録
 * @param {Object} multipleVisitsPatients - 3件以上の訪問がある患者データ
 * @param {Map} patientMap - 患者情報のMap
 * @returns {number} 登録した患者数
 */
function registerMultipleVisitsPatients(multipleVisitsPatients, patientMap) {
  const sheet = getVisitSelectionSheet();

  // 既存データをクリア（ヘッダー以外）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 9).clearContent();
    // データ入力規則もクリア
    sheet.getRange(2, 6, lastRow - 1, 2).clearDataValidations();
  }

  const patientIds = Object.keys(multipleVisitsPatients);
  if (patientIds.length === 0) {
    logInfo('3件以上の訪問がある患者はいません');
    return 0;
  }

  const rows = [];
  patientIds.forEach(patientId => {
    const visits = multipleVisitsPatients[patientId];
    const patient = patientMap.get(patientId);
    const visitCount = visits.length;

    // 訪問日一覧を作成
    const visitDates = visits.map(v => `${v.visitNumber}回目: ${v.dateText}`).join('\n');

    // デフォルト値を設定
    let defaultVisit1, defaultVisit2;
    if (visitCount === 3) {
      defaultVisit1 = BASE_CONFIG.defaultVisitPattern.visits3[0];
      defaultVisit2 = BASE_CONFIG.defaultVisitPattern.visits3[1];
    } else {
      defaultVisit1 = BASE_CONFIG.defaultVisitPattern.visits4[0];
      defaultVisit2 = BASE_CONFIG.defaultVisitPattern.visits4[1];
    }

    rows.push([
      patientId,
      patient ? patient.name : '',
      patient ? normalizeFacilityName(patient.facility) : '',
      visitCount,
      visitDates,
      `${defaultVisit1}回目`,
      `${defaultVisit2}回目`,
      BASE_CONFIG.status.unprocessed,
      ''  // 備考
    ]);
  });

  // データを書き込み
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 9).setValues(rows);

    // プルダウンを設定
    patientIds.forEach((patientId, index) => {
      const rowIndex = index + 2;
      const visits = multipleVisitsPatients[patientId];
      const visitCount = visits.length;

      // 選択肢を作成（1回目〜N回目）
      const options = [];
      for (let i = 1; i <= visitCount; i++) {
        options.push(`${i}回目`);
      }

      // データ入力規則を設定
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(options, true)
        .setAllowInvalid(false)
        .build();

      sheet.getRange(rowIndex, 6).setDataValidation(rule);  // 書類①
      sheet.getRange(rowIndex, 7).setDataValidation(rule);  // 書類②
    });

    // セルの書式設定（訪問日一覧は折り返し表示）
    sheet.getRange(2, 5, rows.length, 1).setWrap(true);
  }

  SpreadsheetApp.flush();
  logInfo('訪問選択シートに登録しました', { count: rows.length });
  return rows.length;
}

/**
 * 訪問選択シートから選択内容を取得
 * バリデーションエラーがある場合は備考に理由を記載してスキップ
 * @returns {Object} { patientId: { visit1Number, visit2Number } }
 */
function getVisitSelections() {
  const sheet = getVisitSelectionSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return {};

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  const selections = {};

  data.forEach((row, index) => {
    const rowIndex = index + 2;  // シート上の行番号（1行目はヘッダー）
    // 患者IDを6桁ゼロパディングに正規化（スプレッドシートで数値として解釈される対策）
    const patientId = String(row[0]).trim().padStart(6, '0');
    if (!patientId || patientId === '000000') return;

    const status = row[7];
    if (status === BASE_CONFIG.status.processed) return;  // 処理済みはスキップ

    const visit1Text = row[5];  // "1回目" などの形式
    const visit2Text = row[6];

    // "N回目" から数字を抽出
    const visit1Number = parseInt(visit1Text, 10);
    const visit2Number = parseInt(visit2Text, 10);

    // バリデーション: 訪問回が正しく選択されているか
    if (isNaN(visit1Number) || isNaN(visit2Number)) {
      const errorMessage = '訪問回の選択が不正です';
      sheet.getRange(rowIndex, 9).setValue(errorMessage);
      logInfo(errorMessage, { patientId, visit1Text, visit2Text });
      return;
    }

    // バリデーション: 書類①と書類②に同じ訪問回が選択されていないか
    if (visit1Number === visit2Number) {
      const errorMessage = `書類①②に同じ訪問回（${visit1Number}回目）が選択されています`;
      sheet.getRange(rowIndex, 9).setValue(errorMessage);
      logInfo(errorMessage, { patientId, visit1Number, visit2Number });
      return;
    }

    // バリデーション成功: 備考をクリア
    sheet.getRange(rowIndex, 9).setValue('');

    selections[patientId] = {
      visit1Number,
      visit2Number
    };
  });

  return selections;
}

/**
 * 訪問選択シートの患者を処理済みにマーク
 * @param {string} patientId - 患者ID
 */
function markVisitSelectionAsProcessed(patientId) {
  const sheet = getVisitSelectionSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  // 比較用に6桁ゼロパディング
  const normalizedPatientId = String(patientId).trim().padStart(6, '0');

  for (let i = 0; i < data.length; i++) {
    const sheetPatientId = String(data[i][0]).trim().padStart(6, '0');
    if (sheetPatientId === normalizedPatientId) {
      sheet.getRange(i + 2, 8).setValue(BASE_CONFIG.status.processed);
      break;
    }
  }
}

/**
 * 訪問選択シートの未処理件数を取得
 * @returns {number} 未処理の患者数
 */
function getPendingVisitSelectionCount() {
  const sheet = getVisitSelectionSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  const data = sheet.getRange(2, 8, lastRow - 1, 1).getValues();
  return data.filter(row => row[0] !== BASE_CONFIG.status.processed).length;
}

/**
 * 訪問選択シートをクリア
 */
function clearVisitSelectionSheet() {
  try {
    const sheet = getVisitSelectionSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 9).clearContent();
      sheet.getRange(2, 6, lastRow - 1, 2).clearDataValidations();
    }
  } catch (e) {
    logError('clearVisitSelectionSheet', e);
  }
}
