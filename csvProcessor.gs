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
function createCsvDownload(data, fileName = 'export.csv') {
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
function exportSheetToCsv(sheetName, range = null) {
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
function importCsvToSheet(fileId, targetSheetName, clearSheet = true) {
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
    
  } catch (error) {
    console.error(`CSVインポートエラー: ${error.message}`);
    throw error;
  }
}

/**
 * 施設データをCSVとしてエクスポート（UI用）
 */
function exportFacilityDataToCsv() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const result = exportSheetToCsv(CONFIG.SHEETS.FACILITY_DATA);
    
    const htmlOutput = HtmlService.createHtmlOutput(`
      <div style="padding: 20px; font-family: Arial, sans-serif;">
        <h3>CSVエクスポート完了</h3>
        <p>ファイル名: ${result.fileName}</p>
        <p>サイズ: ${Math.round(result.size / 1024)} KB</p>
        <p style="margin-top: 20px;">
          <a href="${result.downloadUrl}" target="_blank" 
             style="background-color: #4CAF50; color: white; padding: 10px 20px; 
                    text-decoration: none; border-radius: 4px;">
            ダウンロード
          </a>
        </p>
        <p style="margin-top: 20px; font-size: 12px; color: #666;">
          ※ ダウンロードリンクは一時的なものです。<br>
          必要に応じて保存してください。
        </p>
      </div>
    `)
    .setWidth(400)
    .setHeight(250);
    
    ui.showModalDialog(htmlOutput, 'CSVエクスポート');
    
  } catch (error) {
    ui.alert('エラー', `CSVエクスポート中にエラーが発生しました。\n\n${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * 施設カレンダーをCSVとしてエクスポート（UI用）
 */
function exportCalendarToCsv() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const result = exportSheetToCsv(CONFIG.SHEETS.FACILITY_CALENDAR, CONFIG.PDF.EXPORT_RANGE);
    
    const htmlOutput = HtmlService.createHtmlOutput(`
      <div style="padding: 20px; font-family: Arial, sans-serif;">
        <h3>カレンダーCSVエクスポート完了</h3>
        <p>ファイル名: ${result.fileName}</p>
        <p>範囲: ${CONFIG.PDF.EXPORT_RANGE}</p>
        <p>サイズ: ${Math.round(result.size / 1024)} KB</p>
        <p style="margin-top: 20px;">
          <a href="${result.downloadUrl}" target="_blank" 
             style="background-color: #2196F3; color: white; padding: 10px 20px; 
                    text-decoration: none; border-radius: 4px;">
            ダウンロード
          </a>
        </p>
      </div>
    `)
    .setWidth(400)
    .setHeight(250);
    
    ui.showModalDialog(htmlOutput, 'カレンダーCSVエクスポート');
    
  } catch (error) {
    ui.alert('エラー', `CSVエクスポート中にエラーが発生しました。\n\n${error.message}`, ui.ButtonSet.OK);
  }
}