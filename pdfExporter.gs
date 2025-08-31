/**
 * PDF出力機能
 */

/**
 * 現在の年月を取得
 * @return {string} YYYY年MM月形式の文字列
 */
function getCurrentYearMonth() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  return `${year}年${month}月`;
}

/**
 * 施設カレンダーシートから単一のPDFを作成
 * @param {string} facilityName - 施設名
 * @return {File} 作成されたPDFファイル
 */
function createSinglePdf(facilityName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarSheet = ss.getSheetByName(CONFIG.SHEETS.FACILITY_CALENDAR);
  
  if (!calendarSheet) {
    throw new Error(`シート「${CONFIG.SHEETS.FACILITY_CALENDAR}」が見つかりません。`);
  }
  
  // A5セルに施設名を設定
  calendarSheet.getRange(CONFIG.FACILITY.NAME_CELL).setValue(facilityName);
  SpreadsheetApp.flush(); // 変更を即座に反映
  
  // PDFファイル名を作成
  const yearMonth = getCurrentYearMonth();
  const fileName = `${facilityName}_${yearMonth}.pdf`;
  
  // PDFエクスポートのURL構築
  const spreadsheetId = ss.getId();
  const sheetId = calendarSheet.getSheetId();
  
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?` +
    `format=pdf&` +
    `size=${CONFIG.PDF.SIZE === 'A4' ? '1' : '0'}&` + // 1=A4, 0=Letter
    `portrait=${CONFIG.PDF.ORIENTATION === 'portrait' ? 'true' : 'false'}&` +
    `fitw=${CONFIG.PDF.SCALE === 'fit' ? 'true' : 'false'}&` +
    `sheetnames=false&` +
    `printtitle=false&` +
    `pagenumbers=false&` +
    `gridlines=false&` +
    `fzr=false&` +
    `gid=${sheetId}&` +
    `range=${CONFIG.PDF.EXPORT_RANGE}&` +
    `top_margin=${CONFIG.PDF.MARGINS.top}&` +
    `bottom_margin=${CONFIG.PDF.MARGINS.bottom}&` +
    `left_margin=${CONFIG.PDF.MARGINS.left}&` +
    `right_margin=${CONFIG.PDF.MARGINS.right}`;
  
  // PDFをBlobとして取得
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });
  
  const blob = response.getBlob().setName(fileName);
  
  // Google Driveに保存
  const folderId = getPdfFolderId();
  const folder = DriveApp.getFolderById(folderId);
  
  // 同じ施設の過去のPDFファイルをすべて削除
  // ファイル名のパターン: 施設名_YYYY年MM月.pdf
  const searchPattern = `${facilityName}_`;
  const allFiles = folder.getFiles();
  
  while (allFiles.hasNext()) {
    const file = allFiles.next();
    const existingFileName = file.getName();
    
    // 同じ施設名で始まり、.pdfで終わるファイルを削除
    if (existingFileName.startsWith(searchPattern) && existingFileName.endsWith('.pdf')) {
      console.log(`削除: ${existingFileName}`);
      file.setTrashed(true);
    }
  }
  
  // 新しいPDFを保存
  const pdfFile = folder.createFile(blob);
  
  console.log(`PDF作成完了: ${fileName}`);
  return pdfFile;
}

/**
 * 全施設のPDFを作成
 */
function createAllFacilityPdfs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.SHEETS.FACILITY_DATA);
  
  if (!dataSheet) {
    throw new Error(`シート「${CONFIG.SHEETS.FACILITY_DATA}」が見つかりません。`);
  }
  
  // 施設名を取得（D2から下）
  const startRow = CONFIG.FACILITY.DATA_START_ROW;
  const column = CONFIG.FACILITY.DATA_COLUMN;
  const lastRow = dataSheet.getLastRow();
  
  if (lastRow < startRow) {
    throw new Error('施設データが見つかりません。');
  }
  
  const facilityRange = dataSheet.getRange(`${column}${startRow}:${column}${lastRow}`);
  const facilityValues = facilityRange.getValues().flat();
  
  // ユニークな施設名を取得（個人宅を除外）
  const uniqueFacilities = [...new Set(facilityValues)]
    .filter(name => name && !name.includes(CONFIG.FACILITY.EXCLUDE_PATTERN))
    .sort();
  
  if (uniqueFacilities.length === 0) {
    throw new Error('処理対象の施設が見つかりません。');
  }
  
  console.log(`処理対象施設数: ${uniqueFacilities.length}`);
  
  const results = [];
  const errors = [];
  
  // 各施設のPDFを作成
  uniqueFacilities.forEach((facilityName, index) => {
    try {
      console.log(`処理中 (${index + 1}/${uniqueFacilities.length}): ${facilityName}`);
      const pdfFile = createSinglePdf(facilityName);
      results.push({
        facility: facilityName,
        fileName: pdfFile.getName(),
        fileId: pdfFile.getId(),
        status: 'success'
      });
      
      // API制限を避けるため少し待機
      Utilities.sleep(500);
      
    } catch (error) {
      console.error(`エラー: ${facilityName} - ${error.message}`);
      errors.push({
        facility: facilityName,
        error: error.message,
        status: 'error'
      });
    }
  });
  
  // 結果サマリー
  console.log('\n=== PDF作成完了 ===');
  console.log(`成功: ${results.length}件`);
  console.log(`エラー: ${errors.length}件`);
  
  if (errors.length > 0) {
    console.log('\n=== エラー詳細 ===');
    errors.forEach(e => {
      console.log(`${e.facility}: ${e.error}`);
    });
  }
  
  return {
    success: results,
    errors: errors,
    summary: {
      total: uniqueFacilities.length,
      succeeded: results.length,
      failed: errors.length
    }
  };
}

/**
 * 指定された施設のPDFを作成（UI用）
 */
function createPdfForFacility() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'PDF作成',
    '施設名を入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const facilityName = response.getResponseText().trim();
  if (!facilityName) {
    ui.alert('施設名が入力されていません。');
    return;
  }
  
  try {
    const pdfFile = createSinglePdf(facilityName);
    ui.alert(
      'PDF作成完了',
      `PDFを作成しました。\n\nファイル名: ${pdfFile.getName()}\nファイルID: ${pdfFile.getId()}`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert(
      'エラー',
      `PDF作成中にエラーが発生しました。\n\n${error.message}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 全施設のPDFを作成（UI用）
 */
function createAllPdfsWithConfirmation() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '全施設PDF作成',
    '全施設のPDFを作成します。\n処理には時間がかかる場合があります。\n続行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) {
    return;
  }
  
  try {
    ui.alert('処理を開始します。完了まで画面を閉じないでください。');
    const results = createAllFacilityPdfs();
    
    let message = `PDF作成が完了しました。\n\n`;
    message += `処理対象: ${results.summary.total}件\n`;
    message += `成功: ${results.summary.succeeded}件\n`;
    message += `失敗: ${results.summary.failed}件`;
    
    if (results.summary.failed > 0) {
      message += '\n\n詳細はログを確認してください。';
    }
    
    ui.alert('完了', message, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert(
      'エラー',
      `処理中にエラーが発生しました。\n\n${error.message}`,
      ui.ButtonSet.OK
    );
  }
}