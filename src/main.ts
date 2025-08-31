// カレンダー取得とCSV/PDF出力を一括実行する関数
function executeAll(): { csv: any; facilityPdf: any; generalPdf: any } {
  try {
    console.log('===== 処理開始 =====');
    
    // 1. カレンダーデータを取得してスプレッドシートに出力
    console.log('1. カレンダーデータを取得中...');
    exportCalendarToSpreadsheet();
    
    // 2. CSVファイルを生成
    console.log('2. CSVファイルを生成中...');
    const csvResult = exportCsvSheetToFile();
    
    // 3. 施設カレンダーPDFファイルを生成
    console.log('3. 施設カレンダーPDFファイルを生成中...');
    const facilityPdfResult = exportFacilityCalendarToPDF();
    
    // 4. 汎用PDFファイルを生成
    console.log('4. 汎用PDFファイルを生成中...');
    const generalPdfResult = exportSpreadsheetToPDF();
    
    console.log('===== 処理完了 =====');
    if (csvResult) {
      console.log('CSVファイルのダウンロードURL:');
      console.log(csvResult.downloadUrl);
    }
    console.log('施設カレンダーPDFのダウンロードURL:');
    console.log(facilityPdfResult.downloadUrl);
    console.log('汎用PDFのダウンロードURL:');
    console.log(generalPdfResult.downloadUrl);
    
    return {
      csv: csvResult,
      facilityPdf: facilityPdfResult,
      generalPdf: generalPdfResult
    };
    
  } catch (error) {
    console.error('一括実行エラー:', error);
    throw error;
  }
}