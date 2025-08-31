// CSVシートをCSVファイルとしてエクスポートする関数
function exportCsvSheetToFile(): { fileName: string; fileUrl: string; downloadUrl: string } | null {
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