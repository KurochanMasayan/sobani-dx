/**
 * 全シート一括CSV出力
 * シンプルで軽量な実装
 */
function exportAllSheetsToCSV() {
  // 固定のスプレッドシートIDと出力先フォルダID
  const SPREADSHEET_ID = '1-0MC9HCgSozf9oW_c-0fdeamgB0WGtspV03ypUaPTgI';
  const OUTPUT_FOLDER_ID = '1D9WRmTX1T2alkM50IAmAjqjyHU_4lVNA';

  // 指定のスプレッドシートを開く
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmm');

  // 指定のフォルダ内にサブフォルダを作成
  const parentFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
  const folderName = `CSV出力_${timestamp}`;
  const folder = parentFolder.createFolder(folderName);

  console.log(`${ss.getName()} - ${sheets.length}シートのCSV出力を開始...`);

  sheets.forEach((sheet, index) => {
    const sheetName = sheet.getName();
    console.log(`${index + 1}/${sheets.length}: ${sheetName}`);

    // データ取得
    const data = sheet.getDataRange().getValues();

    // 空のシートはスキップ
    if (data.length === 0) {
      console.log(`  スキップ（データなし）`);
      return;
    }

    // CSV形式に変換（簡単なエスケープ処理付き）
    const csvContent = data.map(row =>
      row.map(cell => {
        const value = String(cell || '');
        // カンマや改行が含まれる場合はダブルクォートで囲む
        return value.includes(',') || value.includes('\n') || value.includes('"')
          ? `"${value.replace(/"/g, '""')}"`
          : value;
      }).join(',')
    ).join('\n');

    // BOM付きUTF-8でファイル作成（Excel対応）
    const bom = '\uFEFF';
    const fileName = `${sheetName.replace(/[/\\?%*:|"<>]/g, '_')}.csv`;
    const blob = Utilities.newBlob(bom + csvContent, 'text/csv', fileName);

    folder.createFile(blob);
  });

  const folderUrl = folder.getUrl();
  console.log(`\n✅ CSV出力完了！\n📁 出力先: ${folderUrl}`);

  return folderUrl;
}