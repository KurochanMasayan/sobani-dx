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
 * カレンダーデータ同期（ボタン用）
 */
function syncCalendarButton() {
  try {
    console.log('カレンダーデータの同期を開始します...');
    const result = syncCalendarData();
    console.log(`同期完了: ${result.eventCount}件のイベントを取得`);
    return result;
  } catch (error) {
    console.error(`カレンダー同期エラー: ${error.message}`);
    throw error;
  }
}

/**
 * カレンダーデータをCSV出力してリンクを貼る（ボタン用）
 */
function exportCalendarCSVButton() {
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
  } catch (error) {
    console.error(`CSV出力エラー: ${error.message}`);
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