/**
 * メインファイル
 * メニューの作成と基本機能の管理
 */

/**
 * ボタンから実行する関数
 */

/**
 * 単一施設のPDF作成（ボタン用）
 */
function createSinglePdfButton(): GoogleAppsScript.Drive.File {
  try {
    // 施設カレンダーシートのA5セルから施設名を取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const calendarSheet = ss.getSheetByName(CONFIG.SHEETS.FACILITY_CALENDAR);
    
    if (!calendarSheet) {
      throw new Error(`シート「${CONFIG.SHEETS.FACILITY_CALENDAR}」が見つかりません`);
    }
    
    const facilityName = calendarSheet.getRange(CONFIG.FACILITY.NAME_CELL).getValue();
    
    if (!facilityName) {
      throw new Error(`${CONFIG.SHEETS.FACILITY_CALENDAR}シートの${CONFIG.FACILITY.NAME_CELL}セルに施設名が設定されていません`);
    }
    
    console.log(`施設名「${facilityName}」のPDFを作成します`);
    const pdfFile = createSinglePdf(facilityName);
    console.log(`PDF作成完了: ${pdfFile.getName()}`);
    
    return pdfFile;
  } catch (error: any) {
    console.error(`PDF作成エラー: ${error.message}`);
    throw error;
  }
}

/**
 * 全施設のPDF作成（ボタン用）
 */
function createAllPdfsButton(): any {
  try {
    console.log('全施設のPDF作成を開始します...');
    const results = createAllFacilityPdfs();
    
    console.log('=== PDF作成結果 ===');
    console.log(`成功: ${results.summary.succeeded}件`);
    console.log(`失敗: ${results.summary.failed}件`);
    
    return results;
  } catch (error: any) {
    console.error(`PDF一括作成エラー: ${error.message}`);
    throw error;
  }
}

/**
 * カレンダーデータ同期（ボタン用）
 * 通常カレンダーと歯科カレンダーの両方を同期
 */
function syncCalendarButton(): any {
  try {
    console.log('カレンダーデータの同期を開始します...');
    const result = syncCalendarData();
    console.log(`同期完了: 通常カレンダー ${result.eventCount}件、歯科カレンダー ${result.dentalEventCount}件`);
    return result;
  } catch (error: any) {
    console.error(`カレンダー同期エラー: ${error.message}`);
    throw error;
  }
}

/**
 * カレンダーデータをCSV出力してリンクを貼る（ボタン用）
 */
function exportCalendarCSVButton(): any {
  try {
    console.log('カレンダーデータのCSV出力を開始します...');
    const result = exportCalendarToCSVWithLink();

    if (result) {
      console.log(`CSV出力完了: ${result.fileName}`);
      console.log(`イベント数: ${result.eventCount}件`);
    } else {
      console.log('CSVデータがありません');
    }

    return result;
  } catch (error: any) {
    console.error(`CSV出力エラー: ${error.message}`);
    throw error;
  }
}

