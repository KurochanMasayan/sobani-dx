/**
 * PDF出力機能
 */

/**
 * 施設カレンダーシートのA2セルから年月を取得
 * @param {Sheet} calendarSheet - 施設カレンダーシート
 * @return {string} YYYY年MM月形式の文字列
 */
function getYearMonthFromSheet(calendarSheet: GoogleAppsScript.Spreadsheet.Sheet): string {
  const dateValue = calendarSheet.getRange('A2').getValue();
  
  if (!dateValue) {
    // A2セルに日付がない場合は現在の年月を使用
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    return `${year}年${month}月`;
  }
  
  // 日付から年月を取得
  const date = new Date(dateValue);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  return `${year}年${month}月`;
}

/**
 * 施設カレンダーシートから単一のPDFを作成
 * @param {string} facilityName - 施設名
 * @return {File} 作成されたPDFファイル
 */
function createSinglePdf(facilityName: string): GoogleAppsScript.Drive.File {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarSheet = ss.getSheetByName(CONFIG.SHEETS.FACILITY_CALENDAR);
  
  if (!calendarSheet) {
    throw new Error(`シート「${CONFIG.SHEETS.FACILITY_CALENDAR}」が見つかりません。`);
  }
  
  // A5セルに施設名を設定
  calendarSheet.getRange(CONFIG.FACILITY.NAME_CELL).setValue(facilityName);
  SpreadsheetApp.flush(); // 変更を即座に反映
  
  // A2セルから年月を取得してPDFファイル名を作成
  const yearMonth = getYearMonthFromSheet(calendarSheet);
  const fileName = `${facilityName}_${yearMonth}.pdf`;
  
  // PDFエクスポートのURL構築
  const spreadsheetId = ss.getId();
  const sheetId = calendarSheet.getSheetId();
  
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?` +
    `format=pdf&` +
    `size=${CONFIG.PDF.SIZE === 'A4' ? '1' : '0'}&` + // 1=A4, 0=Letter
    `portrait=${CONFIG.PDF.ORIENTATION === 'landscape' ? 'false' : 'true'}&` +
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
function createAllFacilityPdfs(): {
  success: Array<{
    facility: string;
    fileName: string;
    fileId: string;
    status: string;
  }>;
  errors: Array<{
    facility: string;
    error: string;
    status: string;
  }>;
  summary: {
    total: number;
    succeeded: number;
    failed: number;
  };
} {
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
  
  const results: Array<{
    facility: string;
    fileName: string;
    fileId: string;
    status: string;
  }> = [];
  const errors: Array<{
    facility: string;
    error: string;
    status: string;
  }> = [];
  
  // 各施設のPDFを作成
  uniqueFacilities.forEach((facilityName, index) => {
    try {
      console.log(`処理中 (${index + 1}/${uniqueFacilities.length}): ${facilityName}`);
      const pdfFile = createSinglePdf(facilityName as string);
      results.push({
        facility: facilityName as string,
        fileName: pdfFile.getName(),
        fileId: pdfFile.getId(),
        status: 'success'
      });
      
      // API制限を避けるため少し待機
      Utilities.sleep(500);
      
    } catch (error: any) {
      console.error(`エラー: ${facilityName} - ${error.message}`);
      errors.push({
        facility: facilityName as string,
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

