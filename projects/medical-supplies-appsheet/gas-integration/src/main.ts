/**
 * メイン実行関数群
 * GASエディタやトリガーから呼び出すためのラッパー
 */

/** 全シート一括CSV出力を実行 */
function exportAllSheets() {
  return exportAllSheetsToCSV();
}

/** AppSheetスキーマに基づいてデータを処理 */
function processWithAppSheetSchema() {
  return processAppSheetData();
}

/** 元データをAppSheet用にフォーマット */
function formatDataForAppSheet() {
  return importDataWithSchema();
}

/** Drive上のCSV/TSVをスプレッドシートへ同期 */
function syncDriveCsvToSpreadsheet() {
  updateAllCsvFiles();
}

/** エディタ実行向けのDrive同期 */
function syncDriveCsvToSpreadsheetFromEditor() {
  return updateAllCsvFilesFromEditor();
}
