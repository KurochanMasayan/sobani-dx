/**
 * CSV処理機能
 * CSVのダウンロードと処理を行う
 */


/**
 * CSVデータをダウンロード可能な形式に変換
 * @param {Array<Array>} data - 2次元配列のデータ
 * @param {string} fileName - ファイル名
 * @return {Object} ダウンロード情報
 */
function createCsvDownload(data: any[][], fileName: string = 'export.csv'): {
  fileId: string;
  fileName: string;
  downloadUrl: string;
  mimeType: string;
  size: number;
} {
  if (!data || data.length === 0) {
    throw new Error('データが空です。');
  }
  
  // データをCSV形式に変換
  const csvContent = data.map(row => {
    return row.map(cell => {
      // セルの値を文字列に変換
      const value = cell === null || cell === undefined ? '' : String(cell);
      
      // カンマ、改行、ダブルクォートが含まれる場合はダブルクォートで囲む
      if (value.includes(',') || value.includes('\n') || value.includes('"')) {
        return '"' + value.replace(/"/g, '""') + '"';
      }
      return value;
    }).join(',');
  }).join('\n');
  
  // BOM付きUTF-8で作成（Excelで文字化けを防ぐ）
  const bom = '\uFEFF';
  const blob = Utilities.newBlob(bom + csvContent, 'text/csv', fileName);
  
  // Google Driveに一時保存
  const tempFile = DriveApp.createFile(blob);
  const fileId = tempFile.getId();
  
  // ダウンロードURLを生成
  const downloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
  
  return {
    fileId: fileId,
    fileName: fileName,
    downloadUrl: downloadUrl,
    mimeType: 'text/csv',
    size: blob.getBytes().length
  };
}

/**
 * スプレッドシートからCSVをエクスポート
 * @param {string} sheetName - シート名
 * @param {string} range - 範囲（オプション）
 * @return {Object} ダウンロード情報
 */
function exportSheetToCsv(sheetName: string, range: string | null = null): {
  fileId: string;
  fileName: string;
  downloadUrl: string;
  mimeType: string;
  size: number;
} {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error(`シート「${sheetName}」が見つかりません。`);
  }
  
  let data;
  if (range) {
    data = sheet.getRange(range).getValues();
  } else {
    data = sheet.getDataRange().getValues();
  }
  
  const fileName = `${sheetName}_${new Date().getTime()}.csv`;
  return createCsvDownload(data, fileName);
}

/**
 * CSVファイルをインポート
 * @param {string} fileId - Google DriveのファイルID
 * @param {string} targetSheetName - インポート先のシート名
 * @param {boolean} clearSheet - シートをクリアするかどうか
 */
function importCsvToSheet(fileId: string, targetSheetName: string, clearSheet: boolean = true): {
  success: boolean;
  rowCount: number;
  columnCount: number;
  sheetName: string;
} {
  try {
    // ファイルを取得
    const file = DriveApp.getFileById(fileId);
    const csvData = file.getBlob().getDataAsString('UTF-8');
    
    // CSVをパース
    const parsedData = Utilities.parseCsv(csvData);
    
    if (!parsedData || parsedData.length === 0) {
      throw new Error('CSVデータが空です。');
    }
    
    // シートを取得または作成
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(targetSheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(targetSheetName);
    } else if (clearSheet) {
      sheet.clear();
    }
    
    // データを書き込み
    const range = sheet.getRange(1, 1, parsedData.length, parsedData[0].length);
    range.setValues(parsedData);
    
    console.log(`CSVインポート完了: ${parsedData.length}行を「${targetSheetName}」シートに書き込みました。`);
    
    return {
      success: true,
      rowCount: parsedData.length,
      columnCount: parsedData[0].length,
      sheetName: targetSheetName
    };
    
  } catch (error: any) {
    console.error(`CSVインポートエラー: ${error.message}`);
    throw error;
  }
}
