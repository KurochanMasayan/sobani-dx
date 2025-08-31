function exportCalendarToSpreadsheet() {
  
  try {
    // スプレッドシートを取得
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 使い方シートから日付を取得
    let usageSheet = spreadsheet.getSheetByName('使い方');
    if (!usageSheet) {
      throw new Error('「使い方」シートが見つかりません');
    }
    
    // B2とB3から開始日と終了日を取得
    const startDateValue = usageSheet.getRange('B2').getValue();
    const endDateValue = usageSheet.getRange('B3').getValue();
    
    if (!startDateValue || !endDateValue) {
      throw new Error('使い方シートのB2（開始日）またはB3（終了日）が設定されていません');
    }
    
    // 日付をJSTで処理（時刻を0:00:00に設定）
    const startDate = new Date(startDateValue);
    startDate.setHours(0, 0, 0, 0);
    
    // 終了日は23:59:59.999に設定して、その日の終わりまで含める
    const endDate = new Date(endDateValue);
    endDate.setHours(23, 59, 59, 999);
    
    // カレンダーを取得
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    let sheet = spreadsheet.getSheetByName("カレンダーデータ");
    
    // シートが存在しない場合は作成
    if (!sheet) {
      sheet = spreadsheet.insertSheet("カレンダーデータ");
    }
    
    // A~D列のみクリア（他の列のデータは保持）
    const lastRow = sheet.getMaxRows();
    if (lastRow > 0) {
      sheet.getRange(2, 1, lastRow, 5).clear();
    }
    
    // ヘッダー行を設定
    sheet.getRange(1, 1, 1, 5).setValues([['日付', '開始時刻', '終了時刻', 'タイトル', 'コメント']]);
    
    // イベントを取得
    const events = calendar.getEvents(startDate, endDate);
    
    if (events.length === 0) {
      sheet.getRange(2, 1).setValue('指定期間にイベントはありません');
      console.log('指定期間にイベントはありません');
      return;
    }
    
    // データ行を準備
    const data = [];
    events.forEach(event => {
      const startTime = event.getStartTime();
      const endTime = event.getEndTime();
      
      data.push([
        Utilities.formatDate(startTime, 'JST', 'yyyy/M/d'),
        event.isAllDayEvent() ? '終日' : Utilities.formatDate(startTime, 'JST', 'H:mm:ss'),
        event.isAllDayEvent() ? '終日' : Utilities.formatDate(endTime, 'JST', 'H:mm:ss'),
        event.getTitle(),
        event.getDescription() || ''  // 説明をE列に直接記入
      ]);
    });
    
    // データを一括でスプレッドシートに書き込み
    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, 5).setValues(data);
    }
    
    console.log(`${events.length}件のイベントをエクスポートしました`);
    console.log(`スプレッドシート: ${spreadsheet.getName()}`);
    console.log(`シート: ${"カレンダーデータ"}`);
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    throw error;
  }
}


// CSVシートをCSVファイルとしてエクスポートする関数
function exportCsvSheetToFile() {
  try {
    // スプレッドシートを取得
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const csvSheet = spreadsheet.getSheetByName('csv');
    
    if (!csvSheet) {
      throw new Error('「csv」シートが見つかりません');
    }
    
    // データが存在する最終行を取得
    const lastRow = csvSheet.getLastRow();
    if (lastRow === 0) {
      console.log('csvシートにデータがありません');
      return null;
    }
    
    // A〜N列（1〜14列）のデータを取得
    const range = csvSheet.getRange(1, 1, lastRow, 14); // A1:N[lastRow]
    const values = range.getValues();
    
    // CSV形式に変換（改行やカンマを含む値を適切にエスケープ）
    let csvContent = '';
    values.forEach((row, index) => {
      const csvRow = row.map((cell, colIndex) => {
        // セルの値を文字列に変換
        let cellValue = cell === null || cell === undefined ? '' : String(cell);
        
        // A列（0）、B列（1）、C列（2）の日付・時刻データを文字列として扱う
        if (colIndex <= 2 && cell instanceof Date) {
          // 日付オブジェクトの場合、適切にフォーマット
          if (colIndex === 0) {
            // A列: 日付を yyyy/M/d 形式で文字列化
            cellValue = Utilities.formatDate(cell, 'JST', 'yyyy/M/d');
          } else {
            // B列、C列: 時刻を H:mm:ss 形式で文字列化
            cellValue = Utilities.formatDate(cell, 'JST', 'H:mm:ss');
          }
        }
        
        // ヘッダー行（1行目）以外は全てダブルクォートで囲む
        if (index > 0) {
          // ダブルクォートをエスケープ（""に変換）
          cellValue = cellValue.replace(/"/g, '""');
          // ダブルクォートで囲む
          cellValue = `"${cellValue}"`;
        } else {
          // ヘッダー行でもカンマや改行、ダブルクォートを含む場合はダブルクォートで囲む
          if (cellValue.includes(',') || cellValue.includes('\n') || cellValue.includes('\r') || cellValue.includes('"')) {
            // ダブルクォートをエスケープ（""に変換）
            cellValue = cellValue.replace(/"/g, '""');
            // ダブルクォートで囲む
            cellValue = `"${cellValue}"`;
          }
        }
        
        return cellValue;
      }).join(',');
      
      csvContent += csvRow;
      // 最終行以外は改行を追加
      if (index < values.length - 1) {
        csvContent += '\n';
      }
    });
    
    // ファイル名を生成
    const fileName = `export_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}.csv`;
    
    // UTF-8でCSVを作成（文字コード変換は手動で行う）
    const blob = Utilities.newBlob(csvContent, 'text/csv', fileName);
    
    // Googleドライブに保存
    const file = DriveApp.createFile(blob);
    
    console.log('===========================================');
    console.log(`CSVファイルを作成しました: ${fileName}`);
    console.log(`ファイルURL: ${file.getUrl()}`);
    console.log(`ダウンロードURL: ${file.getDownloadUrl()}`);
    console.log('===========================================');
    console.log('ダウンロードするには以下のURLをブラウザで開いてください:');
    console.log(file.getDownloadUrl());
    console.log('===========================================');
    console.log('※ 文字コード: UTF-8');
    console.log('※ Shift-JISが必要な場合は、メモ帳で開いて「ANSI」形式で保存してください');
    
    // 結果を使い方シートに書き込む
    const resultSheet = spreadsheet.getSheetByName('使い方');
    if (resultSheet) {
      // CSVファイル名（12行目）
      resultSheet.getRange('B12').setValue(fileName);
      
      // ダウンロードリンク（13行目）
      const richText = SpreadsheetApp.newRichTextValue()
        .setText('📥 クリックしてダウンロード')
        .setLinkUrl(file.getDownloadUrl())
        .build();
      resultSheet.getRange('B13').setRichTextValue(richText);
      
      // ダウンロードリンクのスタイル設定
      resultSheet.getRange('B13').setFontColor('#1a73e8').setFontWeight('bold');
    }
    
    return {
      fileName: file.getName(),
      fileUrl: file.getUrl(),
      downloadUrl: file.getDownloadUrl()
    };
    
  } catch (error) {
    console.error('CSV出力エラー:', error);
    throw error;
  }
}


// カレンダー取得とCSV出力を一括実行する関数
function executeAll() {
  try {
    console.log('===== 処理開始 =====');
    
    // 1. カレンダーデータを取得してスプレッドシートに出力
    console.log('1. カレンダーデータを取得中...');
    exportCalendarToSpreadsheet();
    
    // 2. CSVファイルを生成
    console.log('2. CSVファイルを生成中...');
    const result = exportCsvSheetToFile();
    
    console.log('===== 処理完了 =====');
    console.log('CSVファイルのダウンロードURL:');
    console.log(result.downloadUrl);
    
    return result;
    
  } catch (error) {
    console.error('一括実行エラー:', error);
    throw error;
  }
}
