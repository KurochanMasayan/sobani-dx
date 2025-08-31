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
function createSinglePdfButton() {
  try {
    const facilityName = '施設A'; // ここを適切に変更するか、別途入力を受け付ける
    const pdfFile = createSinglePdf(facilityName);
    console.log(`PDF作成完了: ${pdfFile.getName()}`);
  } catch (error) {
    console.error(`PDF作成エラー: ${error.message}`);
  }
}

/**
 * 全施設のPDF作成（ボタン用）
 */
function createAllPdfsButton() {
  try {
    console.log('全施設のPDF作成を開始します...');
    const results = createAllFacilityPdfs();
    
    console.log('=== PDF作成結果 ===');
    console.log(`成功: ${results.summary.succeeded}件`);
    console.log(`失敗: ${results.summary.failed}件`);
    
    return results;
  } catch (error) {
    console.error(`PDF一括作成エラー: ${error.message}`);
    throw error;
  }
}

/**
 * CSVダウンロード（ボタン用）
 */
function downloadCsvButton() {
  try {
    // 施設データをCSV出力
    const result = exportSheetToCsv(CONFIG.SHEETS.FACILITY_DATA);
    console.log(`CSV作成完了: ${result.fileName}`);
    console.log(`ダウンロードURL: ${result.downloadUrl}`);
    return result;
  } catch (error) {
    console.error(`CSVエクスポートエラー: ${error.message}`);
    throw error;
  }
}

/**
 * テスト用：設定値の確認
 */
function testConfiguration() {
  console.log('=== 設定確認 ===');
  console.log('施設データシート:', CONFIG.SHEETS.FACILITY_DATA);
  console.log('施設カレンダーシート:', CONFIG.SHEETS.FACILITY_CALENDAR);
  console.log('PDF出力範囲:', CONFIG.PDF.EXPORT_RANGE);
  console.log('PDFサイズ:', CONFIG.PDF.SIZE);
  console.log('PDF向き:', CONFIG.PDF.ORIENTATION);
  
  try {
    const folderId = getPdfFolderId();
    console.log('PDF保存先フォルダID:', folderId);
    const folder = DriveApp.getFolderById(folderId);
    console.log('PDF保存先フォルダ名:', folder.getName());
  } catch (error) {
    console.error('PDF保存先エラー:', error.message);
  }
}

/**
 * テスト用：単一施設のPDF作成テスト
 */
function testSinglePdfCreation() {
  const testFacilityName = 'テスト施設';
  try {
    const result = createSinglePdf(testFacilityName);
    console.log('テストPDF作成成功:', result.getName());
  } catch (error) {
    console.error('テストPDF作成失敗:', error.message);
  }
}