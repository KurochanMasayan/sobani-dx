// 施設カレンダーシートの特定範囲をPDFとしてエクスポートする関数
function exportFacilityCalendarToPDF(): { fileName: string; fileUrl: string; downloadUrl: string } {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = spreadsheet.getId();
    
    // 施設カレンダーシートを取得
    const targetSheet = spreadsheet.getSheetByName('施設カレンダー');
    if (!targetSheet) {
      throw new Error('「施設カレンダー」シートが見つかりません');
    }
    
    const sheetId = targetSheet.getSheetId();
    
    // PDFエクスポートのパラメータ（横向き、C1:P25の範囲指定）
    const params: { [key: string]: string | number | boolean } = {
      'format': 'pdf',
      'size': 'A4',
      'portrait': false,  // false = 横向き
      'fitw': true,
      'sheetnames': false,
      'printtitle': false,
      'pagenumbers': false,
      'gridlines': true,
      'fzr': false,
      'gid': sheetId,
      // 範囲指定: C1:P25
      'r1': 0,   // 開始行（0ベース）
      'c1': 2,   // 開始列（0ベース、C列 = 2）
      'r2': 24,  // 終了行（0ベース、25行目 = 24）
      'c2': 15   // 終了列（0ベース、P列 = 15）
    };
    
    // URLパラメータを作成
    const queryString = Object.entries(params)
      .map(([key, value]) => `${key}=${value}`)
      .join('&');
    
    // PDFエクスポートURL
    const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?${queryString}`;
    
    // PDFを取得
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });
    
    // ファイル名を生成
    const fileName = `施設カレンダー_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}.pdf`;
    
    // PDFをDriveに保存
    const blob = response.getBlob().setName(fileName);
    const file = DriveApp.createFile(blob);
    
    console.log('===========================================');
    console.log(`PDFファイルを作成しました: ${fileName}`);
    console.log(`範囲: C1:P25（横向き）`);
    console.log(`ファイルURL: ${file.getUrl()}`);
    console.log(`ダウンロードURL: ${file.getDownloadUrl()}`);
    console.log('===========================================');
    
    // 結果を使い方シートに書き込む
    const usageSheet = spreadsheet.getSheetByName('使い方');
    if (usageSheet) {
      // PDFファイル名（15行目）
      usageSheet.getRange('B15').setValue(fileName);
      
      // ダウンロードリンク（16行目）
      const richText = SpreadsheetApp.newRichTextValue()
        .setText('📄 PDFをダウンロード')
        .setLinkUrl(file.getDownloadUrl())
        .build();
      usageSheet.getRange('B16').setRichTextValue(richText);
      
      // ダウンロードリンクのスタイル設定
      usageSheet.getRange('B16').setFontColor('#d93025').setFontWeight('bold');
    }
    
    return {
      fileName: file.getName(),
      fileUrl: file.getUrl(),
      downloadUrl: file.getDownloadUrl()
    };
    
  } catch (error) {
    console.error('PDF出力エラー:', error);
    throw error;
  }
}

// スプレッドシートをPDFとしてエクスポートする関数（汎用版）
function exportSpreadsheetToPDF(): { fileName: string; fileUrl: string; downloadUrl: string } {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetId = spreadsheet.getId();
    
    // 使い方シートから設定を取得
    const usageSheet = spreadsheet.getSheetByName('使い方');
    let sheetName = 'カレンダーデータ';
    
    if (usageSheet) {
      const customSheetName = usageSheet.getRange('B5').getValue();
      if (customSheetName) {
        sheetName = customSheetName as string;
      }
    }
    
    const targetSheet = spreadsheet.getSheetByName(sheetName);
    if (!targetSheet) {
      throw new Error(`「${sheetName}」シートが見つかりません`);
    }
    
    const sheetId = targetSheet.getSheetId();
    
    // PDFエクスポートのパラメータ
    const params: { [key: string]: string | number | boolean } = {
      'format': 'pdf',
      'size': 'A4',
      'portrait': true,
      'fitw': true,
      'sheetnames': false,
      'printtitle': false,
      'pagenumbers': false,
      'gridlines': true,
      'fzr': false,
      'gid': sheetId
    };
    
    // URLパラメータを作成
    const queryString = Object.entries(params)
      .map(([key, value]) => `${key}=${value}`)
      .join('&');
    
    // PDFエクスポートURL
    const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?${queryString}`;
    
    // PDFを取得
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });
    
    // ファイル名を生成
    const fileName = `${sheetName}_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}.pdf`;
    
    // PDFをDriveに保存
    const blob = response.getBlob().setName(fileName);
    const file = DriveApp.createFile(blob);
    
    console.log('===========================================');
    console.log(`PDFファイルを作成しました: ${fileName}`);
    console.log(`ファイルURL: ${file.getUrl()}`);
    console.log(`ダウンロードURL: ${file.getDownloadUrl()}`);
    console.log('===========================================');
    
    // 結果を使い方シートに書き込む
    if (usageSheet) {
      // PDFファイル名（17行目）
      usageSheet.getRange('B17').setValue(fileName);
      
      // ダウンロードリンク（18行目）
      const richText = SpreadsheetApp.newRichTextValue()
        .setText('📄 汎用PDFをダウンロード')
        .setLinkUrl(file.getDownloadUrl())
        .build();
      usageSheet.getRange('B18').setRichTextValue(richText);
      
      // ダウンロードリンクのスタイル設定
      usageSheet.getRange('B18').setFontColor('#188038').setFontWeight('bold');
    }
    
    return {
      fileName: file.getName(),
      fileUrl: file.getUrl(),
      downloadUrl: file.getDownloadUrl()
    };
    
  } catch (error) {
    console.error('PDF出力エラー:', error);
    throw error;
  }
}