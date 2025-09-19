/**
 * 医療用品管理システム用 CSV出力機能
 * スプレッドシートの全シートを一括でCSVファイルとして出力
 */

/**
 * スプレッドシートの全シートを一括でCSV出力
 * @param {string} folderName - 出力先フォルダ名（オプション）
 * @param {string} spreadsheetId - スプレッドシートID（オプション、指定がなければアクティブシートを使用）
 * @return {Object} 出力結果情報
 */
function exportAllSheetsToCSV(folderName: string = '医療用品データ出力', spreadsheetId?: string): {
  success: boolean;
  totalSheets: number;
  processedSheets: number;
  folderId: string;
  folderUrl: string;
  spreadsheetName: string;
  files: Array<{
    sheetName: string;
    fileName: string;
    fileId: string;
    downloadUrl: string;
    rowCount: number;
    columnCount: number;
    dataRange: string;
  }>;
  errors: Array<{
    sheetName: string;
    error: string;
  }>;
} {
  // スプレッドシートを取得
  const ss = spreadsheetId
    ? SpreadsheetApp.openById(spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error('スプレッドシートが見つかりません');
  }

  const sheets = ss.getSheets();
  const spreadsheetName = ss.getName();
  const timestamp = new Date().toISOString().slice(0, 16).replace(/:/g, '-');
  const fullFolderName = `${folderName}_${spreadsheetName}_${timestamp}`;

  console.log(`スプレッドシート「${spreadsheetName}」の全${sheets.length}シートCSV出力を開始...`);

  // 出力用フォルダを作成
  const folder = DriveApp.createFolder(fullFolderName);
  const folderId = folder.getId();
  const folderUrl = `https://drive.google.com/drive/folders/${folderId}`;

  const files: Array<{
    sheetName: string;
    fileName: string;
    fileId: string;
    downloadUrl: string;
    rowCount: number;
    columnCount: number;
    dataRange: string;
  }> = [];

  const errors: Array<{
    sheetName: string;
    error: string;
  }> = [];

  let processedSheets = 0;

  for (const sheet of sheets) {
    try {
      const sheetName = sheet.getName();
      console.log(`処理中: ${sheetName}`);

      // データ範囲を取得
      const dataRange = sheet.getDataRange();
      if (!dataRange || dataRange.getNumRows() === 0) {
        console.warn(`スキップ: ${sheetName} (データなし)`);
        continue;
      }

      const data = dataRange.getValues();
      const rowCount = data.length;
      const columnCount = data[0] ? data[0].length : 0;

      // CSVファイル名を生成（シート名を安全なファイル名に変換）
      const safeSheetName = sheetName.replace(/[/\\?%*:|"<>]/g, '_');
      const fileName = `${safeSheetName}.csv`;

      // CSV形式に変換
      const csvContent = data.map(row => {
        return row.map(cell => {
          // null, undefined, 空文字列の処理
          if (cell === null || cell === undefined) {
            return '';
          }

          const value = String(cell);

          // 日付の場合は適切にフォーマット
          if (cell instanceof Date) {
            return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
          }

          // カンマ、改行、ダブルクォートが含まれる場合はエスケープ
          if (value.includes(',') || value.includes('\n') || value.includes('"')) {
            return '"' + value.replace(/"/g, '""') + '"';
          }

          return value;
        }).join(',');
      }).join('\n');

      // BOM付きUTF-8でファイル作成（Excel対応）
      const bom = '\uFEFF';
      const blob = Utilities.newBlob(bom + csvContent, 'text/csv', fileName);

      // フォルダ内にファイルを作成
      const csvFile = folder.createFile(blob);
      const fileId = csvFile.getId();
      const downloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;

      files.push({
        sheetName: sheetName,
        fileName: fileName,
        fileId: fileId,
        downloadUrl: downloadUrl,
        rowCount: rowCount,
        columnCount: columnCount,
        dataRange: dataRange.getA1Notation()
      });

      processedSheets++;
      console.log(`完了: ${sheetName} (${rowCount}行 × ${columnCount}列, 範囲: ${dataRange.getA1Notation()})`);

    } catch (error: any) {
      const sheetName = sheet.getName();
      console.error(`エラー: ${sheetName} - ${error.message}`);
      errors.push({
        sheetName: sheetName,
        error: error.message
      });
    }
  }

  const result = {
    success: errors.length === 0,
    totalSheets: sheets.length,
    processedSheets: processedSheets,
    folderId: folderId,
    folderUrl: folderUrl,
    spreadsheetName: spreadsheetName,
    files: files,
    errors: errors
  };

  // 結果をログ出力
  console.log('=== CSV出力完了 ===');
  console.log(`スプレッドシート: ${spreadsheetName}`);
  console.log(`処理結果: ${processedSheets}/${sheets.length}シート成功`);
  console.log(`出力先フォルダ: ${folderUrl}`);

  if (errors.length > 0) {
    console.warn(`${errors.length}件のエラーが発生:`);
    errors.forEach(err => console.warn(`- ${err.sheetName}: ${err.error}`));
  }

  // 成功したファイルの一覧
  if (files.length > 0) {
    console.log(`\n出力されたファイル (${files.length}件):`);
    files.forEach(file => {
      console.log(`- ${file.fileName} (${file.rowCount}行)`);
    });
  }

  return result;
}

/**
 * 特定のシートのみをCSV出力
 * @param {string[]} sheetNames - 出力対象のシート名配列
 * @param {string} folderName - 出力先フォルダ名（オプション）
 * @param {string} spreadsheetId - スプレッドシートID（オプション）
 */
function exportSelectedSheetsToCSV(
  sheetNames: string[],
  folderName: string = '医療用品選択データ出力',
  spreadsheetId?: string
): {
  success: boolean;
  totalSheets: number;
  processedSheets: number;
  folderId: string;
  folderUrl: string;
  spreadsheetName: string;
  files: Array<{
    sheetName: string;
    fileName: string;
    fileId: string;
    downloadUrl: string;
    rowCount: number;
  }>;
  errors: Array<{
    sheetName: string;
    error: string;
  }>;
} {
  const ss = spreadsheetId
    ? SpreadsheetApp.openById(spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();

  const timestamp = new Date().toISOString().slice(0, 16).replace(/:/g, '-');
  const fullFolderName = `${folderName}_${timestamp}`;

  // 出力用フォルダを作成
  const folder = DriveApp.createFolder(fullFolderName);
  const folderId = folder.getId();
  const folderUrl = `https://drive.google.com/drive/folders/${folderId}`;

  const files: any[] = [];
  const errors: any[] = [];
  let processedSheets = 0;

  console.log(`選択シート${sheetNames.length}件のCSV出力を開始...`);

  for (const sheetName of sheetNames) {
    try {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        errors.push({
          sheetName: sheetName,
          error: 'シートが見つかりません'
        });
        continue;
      }

      const data = sheet.getDataRange().getValues();
      if (data.length === 0) {
        console.warn(`スキップ: ${sheetName} (データなし)`);
        continue;
      }

      const fileName = `${sheetName.replace(/[/\\?%*:|"<>]/g, '_')}.csv`;
      const csvContent = data.map(row => {
        return row.map(cell => {
          const value = cell === null || cell === undefined ? '' : String(cell);
          if (value.includes(',') || value.includes('\n') || value.includes('"')) {
            return '"' + value.replace(/"/g, '""') + '"';
          }
          return value;
        }).join(',');
      }).join('\n');

      const bom = '\uFEFF';
      const blob = Utilities.newBlob(bom + csvContent, 'text/csv', fileName);
      const csvFile = folder.createFile(blob);

      files.push({
        sheetName: sheetName,
        fileName: fileName,
        fileId: csvFile.getId(),
        downloadUrl: `https://drive.google.com/uc?export=download&id=${csvFile.getId()}`,
        rowCount: data.length
      });

      processedSheets++;
      console.log(`完了: ${sheetName} (${data.length}行)`);

    } catch (error: any) {
      errors.push({
        sheetName: sheetName,
        error: error.message
      });
    }
  }

  return {
    success: errors.length === 0,
    totalSheets: sheetNames.length,
    processedSheets: processedSheets,
    folderId: folderId,
    folderUrl: folderUrl,
    spreadsheetName: ss.getName(),
    files: files,
    errors: errors
  };
}