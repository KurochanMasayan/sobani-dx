/**
 * AppSheet スキーマ処理機能
 * スキーマ情報を読み込んでデータを適切に整理
 */

/**
 * AppSheetスキーマシートから情報を取得して処理
 */
function processAppSheetData() {
  const SPREADSHEET_ID = '1-0MC9HCgSozf9oW_c-0fdeamgB0WGtspV03ypUaPTgI';
  const OUTPUT_FOLDER_ID = '1D9WRmTX1T2alkM50IAmAjqjyHU_4lVNA';

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // AppSheet Schemaシートを探す
  const schemaSheet = ss.getSheetByName('AppSheet Schema');
  if (!schemaSheet) {
    console.error('AppSheet Schemaシートが見つかりません');
    return;
  }

  // スキーマ情報を読み込み
  const schemaData = schemaSheet.getDataRange().getValues();
  const headers = schemaData[0]; // ヘッダー行

  // テーブル名とカラム定義を抽出
  const tableDefinitions: {[key: string]: any[]} = {};
  let currentTable = '';

  for (let i = 1; i < schemaData.length; i++) {
    const row = schemaData[i];
    const tableName = row[0]; // 最初のカラムがテーブル名と仮定

    if (tableName && tableName !== '') {
      currentTable = String(tableName);
      if (!tableDefinitions[currentTable]) {
        tableDefinitions[currentTable] = [];
      }
      tableDefinitions[currentTable].push(row);
    }
  }

  console.log(`スキーマ解析完了: ${Object.keys(tableDefinitions).length}個のテーブル定義を発見`);

  // 各テーブルに対応するシートからデータを取得して処理
  const processedData: {[key: string]: any[][]} = {};

  for (const tableName in tableDefinitions) {
    console.log(`処理中: ${tableName}`);

    // 対応するデータシートを探す
    const dataSheet = ss.getSheetByName(tableName);
    if (!dataSheet) {
      console.warn(`データシート「${tableName}」が見つかりません`);
      continue;
    }

    // データを取得
    const rawData = dataSheet.getDataRange().getValues();
    if (rawData.length <= 1) {
      console.log(`  ${tableName}: データなし`);
      continue;
    }

    // スキーマに基づいてデータを整理
    const processedRows = [];
    const dataHeaders = rawData[0];

    // ヘッダー行を追加
    processedRows.push(dataHeaders);

    // データ行を処理
    for (let i = 1; i < rawData.length; i++) {
      const row = rawData[i];
      const processedRow = [];

      for (let j = 0; j < row.length; j++) {
        let value = row[j];

        // スキーマ定義に基づいてデータを変換
        // 例: 日付フォーマット、数値変換、テキストクレンジングなど
        if (value instanceof Date) {
          value = Utilities.formatDate(value, 'JST', 'yyyy-MM-dd HH:mm:ss');
        } else if (typeof value === 'number') {
          // 数値はそのまま
          value = value;
        } else {
          // テキストはトリミング
          value = String(value || '').trim();
        }

        processedRow.push(value);
      }

      processedRows.push(processedRow);
    }

    processedData[tableName] = processedRows;
    console.log(`  ${tableName}: ${processedRows.length - 1}行のデータを処理`);
  }

  // 処理済みデータを新しいシートに出力
  exportProcessedData(ss, processedData, OUTPUT_FOLDER_ID);

  return processedData;
}

/**
 * 処理済みデータを新しいシートとCSVに出力
 */
function exportProcessedData(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  processedData: {[key: string]: any[][]},
  outputFolderId: string
) {
  const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmm');

  // 出力用フォルダを作成
  const parentFolder = DriveApp.getFolderById(outputFolderId);
  const folderName = `AppSheet処理済み_${timestamp}`;
  const folder = parentFolder.createFolder(folderName);

  console.log('処理済みデータを出力中...');

  // 新しいスプレッドシートを作成
  const newSpreadsheet = SpreadsheetApp.create(`AppSheet処理済みデータ_${timestamp}`);
  const newSpreadsheetFile = DriveApp.getFileById(newSpreadsheet.getId());

  // フォルダに移動
  folder.addFile(newSpreadsheetFile);
  DriveApp.getRootFolder().removeFile(newSpreadsheetFile);

  // 最初のデフォルトシートを取得
  const sheets = newSpreadsheet.getSheets();
  let sheetIndex = 0;

  for (const tableName in processedData) {
    const data = processedData[tableName];

    if (data.length === 0) continue;

    // シートを作成または既存のシートを使用
    let sheet;
    if (sheetIndex < sheets.length) {
      sheet = sheets[sheetIndex];
      sheet.setName(tableName);
    } else {
      sheet = newSpreadsheet.insertSheet(tableName);
    }

    // データを書き込み
    if (data.length > 0 && data[0].length > 0) {
      const range = sheet.getRange(1, 1, data.length, data[0].length);
      range.setValues(data);

      // ヘッダー行を太字に
      const headerRange = sheet.getRange(1, 1, 1, data[0].length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f0f0f0');
    }

    // CSVファイルも作成
    const csvContent = data.map(row =>
      row.map(cell => {
        const value = String(cell || '');
        return value.includes(',') || value.includes('\n') || value.includes('"')
          ? `"${value.replace(/"/g, '""')}"`
          : value;
      }).join(',')
    ).join('\n');

    const bom = '\uFEFF';
    const fileName = `${tableName}.csv`;
    const blob = Utilities.newBlob(bom + csvContent, 'text/csv', fileName);
    folder.createFile(blob);

    console.log(`  ${tableName}: シートとCSVを作成`);
    sheetIndex++;
  }

  // 不要な空のシートを削除
  const allSheets = newSpreadsheet.getSheets();
  for (let i = sheetIndex; i < allSheets.length; i++) {
    if (allSheets.length > 1) { // 最低1シートは必要
      newSpreadsheet.deleteSheet(allSheets[i]);
    }
  }

  console.log(`\n✅ 処理完了！`);
  console.log(`📊 スプレッドシート: ${newSpreadsheet.getUrl()}`);
  console.log(`📁 フォルダ: ${folder.getUrl()}`);

  return {
    spreadsheetUrl: newSpreadsheet.getUrl(),
    folderUrl: folder.getUrl(),
    tableCount: Object.keys(processedData).length
  };
}

/**
 * 元データから適切なフォーマットでデータを入力
 */
function importDataWithSchema() {
  const SPREADSHEET_ID = '1-0MC9HCgSozf9oW_c-0fdeamgB0WGtspV03ypUaPTgI';

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // データシートの命名規則を推測（例: "原始データ"、"raw_data"、"data_"で始まるシート）
  const sheets = ss.getSheets();
  const dataSheets = sheets.filter(sheet => {
    const name = sheet.getName().toLowerCase();
    return name.includes('data') || name.includes('原始') || name.includes('raw') ||
           name.includes('元データ') || name.includes('source');
  });

  console.log(`${dataSheets.length}個のデータシートを発見`);

  // AppSheet向けにフォーマット
  const formattedData: {[key: string]: any[][]} = {};

  dataSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    console.log(`フォーマット中: ${sheetName}`);

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;

    // AppSheet用のカラムヘッダーを正規化
    const headers = data[0].map(header => {
      return String(header)
        .trim()
        .replace(/\s+/g, '_')           // スペースをアンダースコアに
        .replace(/[^\w\u3040-\u309f\u30a0-\u30ff\u4e00-\u9faf]/g, '') // 特殊文字を削除
        .substring(0, 50);               // 最大50文字に制限
    });

    const formattedRows: (string | number)[][] = [headers];

    // データ行をフォーマット
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const formattedRow = row.map((cell, index) => {
        // AppSheet用にデータをクレンジング
        if (cell === null || cell === undefined || cell === '') {
          return '';
        }

        // 日付の処理
        if (cell instanceof Date) {
          return Utilities.formatDate(cell, 'JST', 'yyyy-MM-dd');
        }

        // 数値の処理
        if (typeof cell === 'number') {
          // 大きすぎる数値の処理
          if (Math.abs(cell) > 1e15) {
            return String(cell);
          }
          return cell;
        }

        // テキストの処理
        let text = String(cell).trim();

        // 改行をスペースに置換
        text = text.replace(/[\r\n]+/g, ' ');

        // 最大文字数制限（AppSheetのText型は通常8000文字まで）
        if (text.length > 8000) {
          text = text.substring(0, 7997) + '...';
        }

        return text;
      });

      formattedRows.push(formattedRow);
    }

    // クリーンなテーブル名を生成
    const cleanTableName = sheetName
      .replace(/[^\w\u3040-\u309f\u30a0-\u30ff\u4e00-\u9faf]/g, '_')
      .substring(0, 30);

    formattedData[cleanTableName] = formattedRows;
    console.log(`  ${cleanTableName}: ${formattedRows.length - 1}行をフォーマット`);
  });

  // フォーマット済みデータを出力
  const OUTPUT_FOLDER_ID = '1D9WRmTX1T2alkM50IAmAjqjyHU_4lVNA';
  return exportProcessedData(ss, formattedData, OUTPUT_FOLDER_ID);
}