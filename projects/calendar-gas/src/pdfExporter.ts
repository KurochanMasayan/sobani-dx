/**
 * PDF出力機能
 */

/**
 * 施設カレンダーシートのA2セルから年月を取得
 * @param {Sheet} calendarSheet - 施設カレンダーシート
 * @return {string} YYYY年MM月形式の文字列
 */
function getYearMonthFromSheet(calendarSheet: GoogleAppsScript.Spreadsheet.Sheet): string {
  const dateValue = calendarSheet.getRange('A2').getValue();
  
  if (!dateValue) {
    // A2セルに日付がない場合は現在の年月を使用
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    return `${year}年${month}月`;
  }
  
  // 日付から年月を取得
  const date = new Date(dateValue);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  return `${year}年${month}月`;
}

/**
 * 施設カレンダーシートから単一のPDFを作成
 * @param {string} facilityName - 施設名
 * @return {File} 作成されたPDFファイル
 */
function createSinglePdf(facilityName: string): GoogleAppsScript.Drive.File {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarSheet = ss.getSheetByName(CONFIG.SHEETS.FACILITY_CALENDAR);
  
  if (!calendarSheet) {
    throw new Error(`シート「${CONFIG.SHEETS.FACILITY_CALENDAR}」が見つかりません。`);
  }
  
  // A5セルに施設名を設定
  calendarSheet.getRange(CONFIG.FACILITY.NAME_CELL).setValue(facilityName);
  SpreadsheetApp.flush(); // 変更を即座に反映
  
  // A2セルから年月を取得してPDFファイル名を作成
  const yearMonth = getYearMonthFromSheet(calendarSheet);
  const fileName = `${facilityName}_${yearMonth}.pdf`;
  
  // PDFエクスポートのURL構築
  const spreadsheetId = ss.getId();
  const sheetId = calendarSheet.getSheetId();
  
  // scale: 1=標準(100%), 2=幅に合わせる, 3=高さに合わせる, 4=ページに合わせる
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?` +
    `format=pdf&` +
    `size=${CONFIG.PDF.SIZE === 'A4' ? '1' : '0'}&` + // 1=A4, 0=Letter
    `portrait=${CONFIG.PDF.ORIENTATION === 'landscape' ? 'false' : 'true'}&` +
    `scale=3&` + // 高さに合わせる
    `sheetnames=false&` +
    `printtitle=false&` +
    `pagenumbers=false&` +
    `gridlines=false&` +
    `fzr=false&` +
    `gid=${sheetId}&` +
    `range=${CONFIG.PDF.EXPORT_RANGE}&` +
    `top_margin=${CONFIG.PDF.MARGINS.top}&` +
    `bottom_margin=${CONFIG.PDF.MARGINS.bottom}&` +
    `left_margin=${CONFIG.PDF.MARGINS.left}&` +
    `right_margin=${CONFIG.PDF.MARGINS.right}`;
  
  // PDFをBlobとして取得
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });
  
  const blob = response.getBlob().setName(fileName);
  
  // Google Driveに保存
  const folderId = getPdfFolderId();
  const folder = DriveApp.getFolderById(folderId);
  
  // 同じ施設の過去のPDFファイルをすべて削除
  // ファイル名のパターン: 施設名_YYYY年MM月.pdf
  const searchPattern = `${facilityName}_`;
  const allFiles = folder.getFiles();
  
  while (allFiles.hasNext()) {
    const file = allFiles.next();
    const existingFileName = file.getName();
    
    // 同じ施設名で始まり、.pdfで終わるファイルを削除
    if (existingFileName.startsWith(searchPattern) && existingFileName.endsWith('.pdf')) {
      console.log(`削除: ${existingFileName}`);
      file.setTrashed(true);
    }
  }
  
  // 新しいPDFを保存
  const pdfFile = folder.createFile(blob);
  
  console.log(`PDF作成完了: ${fileName}`);
  return pdfFile;
}

/**
 * 施設別訪問一覧から施設データを取得
 * 設定に基づいて、空白や特定のパターンを除外したユニークな値を取得
 * @return {Array<string>} 施設データの配列
 */
function getFacilityData(): string[] {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.FACILITY.VISIT_LIST.SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`${CONFIG.FACILITY.VISIT_LIST.SHEET_NAME}シートが見つかりません。`);
  }
  
  // 設定に基づいてデータを取得
  const lastRow = sheet.getLastRow();
  const startRow = CONFIG.FACILITY.VISIT_LIST.START_ROW;
  const dataColumn = CONFIG.FACILITY.VISIT_LIST.DATA_COLUMN;
  
  if (lastRow < startRow) {
    return [];
  }
  
  const range = sheet.getRange(startRow, dataColumn, lastRow - startRow + 1, 1);
  const values = range.getValues();
  
  // ユニークな施設データを格納するSet
  const facilitySet = new Set<string>();
  
  // 各値をチェックして条件に合うものを追加
  values.forEach(row => {
    const value = String(row[0]).trim();
    
    // 空白をスキップ
    if (!value) {
      return;
    }
    
    // 設定されたスキップパターンをチェック
    const shouldSkip = CONFIG.FACILITY.VISIT_LIST.SKIP_PATTERNS.some(pattern => 
      value.includes(pattern)
    );
    
    if (shouldSkip) {
      return;
    }
    
    // 条件を満たす値を追加
    facilitySet.add(value);
  });
  
  // Setを配列に変換してソートして返す
  return Array.from(facilitySet).sort();
}

/**
 * 処理の進捗を保存・取得するためのプロパティサービス
 */
function getProgressProperty() {
  return PropertiesService.getScriptProperties();
}

/**
 * 実行時間制限対応のトリガーから再開する関数
 */
function resumePdfCreationFromTrigger(): void {
  try {
    console.log('トリガーから処理を再開します...');
    createAllFacilityPdfs();
    
    // 処理完了後、トリガーを削除
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'resumePdfCreationFromTrigger') {
        ScriptApp.deleteTrigger(trigger);
        console.log('トリガーを削除しました');
      }
    });
  } catch (error: any) {
    console.error('トリガー実行エラー:', error.message);
  }
}

/**
 * 全施設のPDFを作成（実行時間制限対応版）
 */
function createAllFacilityPdfs(): {
  success: Array<{
    facility: string;
    fileName: string;
    fileId: string;
    status: string;
  }>;
  errors: Array<{
    facility: string;
    error: string;
    status: string;
  }>;
  summary: {
    total: number;
    succeeded: number;
    failed: number;
  };
} {
  const startTime = new Date().getTime();
  const properties = getProgressProperty();
  
  // 前回の進捗を取得
  const savedProgress = properties.getProperty('pdfCreationProgress');
  let startIndex = 0;
  let results: Array<{
    facility: string;
    fileName: string;
    fileId: string;
    status: string;
  }> = [];
  let errors: Array<{
    facility: string;
    error: string;
    status: string;
  }> = [];
  
  // 保存された進捗がある場合は復元
  if (savedProgress) {
    const progress = JSON.parse(savedProgress);
    startIndex = progress.nextIndex;
    results = progress.results || [];
    errors = progress.errors || [];
    console.log(`前回の処理を続行します（${startIndex + 1}件目から）`);
  }
  
  // getFacilityData関数を使用して施設リストを取得
  const uniqueFacilities = getFacilityData();
  
  if (uniqueFacilities.length === 0) {
    throw new Error('処理対象の施設が見つかりません。');
  }
  
  console.log(`処理対象施設数: ${uniqueFacilities.length}`);
  
  // 各施設のPDFを作成
  for (let index = startIndex; index < uniqueFacilities.length; index++) {
    const facilityName = uniqueFacilities[index];
    
    // 処理前に実行時間をチェック（5分 = 300秒）
    const elapsedTime = (new Date().getTime() - startTime) / 1000;
    if (elapsedTime > 280) { // 4分40秒で中断
      // 進捗を保存（次に処理するインデックスを保存）
      const progress = {
        nextIndex: index,  // 現在のインデックスを次回の開始位置として保存
        results: results,
        errors: errors,
        totalFacilities: uniqueFacilities.length
      };
      properties.setProperty('pdfCreationProgress', JSON.stringify(progress));
      
      // トリガーを作成して1分後に再開
      ScriptApp.newTrigger('resumePdfCreationFromTrigger')
        .timeBased()
        .after(60 * 1000) // 1分後
        .create();
      
      console.log(`実行時間制限のため処理を中断します。1分後に自動再開します。（${index}/${uniqueFacilities.length}完了）`);
      
      return {
        success: results,
        errors: errors,
        summary: {
          total: index,
          succeeded: results.length,
          failed: errors.length
        }
      };
    }
    // 12件ごとに長めの休憩を入れる
    if (index > 0 && index % 12 === 0) {
      console.log('レート制限回避のため30秒待機中...');
      Utilities.sleep(30000); // 30秒待機
    }
    
    let retryCount = 0;
    const maxRetries = 3;
    let success = false;
    
    while (!success && retryCount < maxRetries) {
      try {
        console.log(`処理中 (${index + 1}/${uniqueFacilities.length}): ${facilityName}`);
        const pdfFile = createSinglePdf(facilityName as string);
        results.push({
          facility: facilityName as string,
          fileName: pdfFile.getName(),
          fileId: pdfFile.getId(),
          status: 'success'
        });
        success = true;
        
        // 成功したら進捗を保存（念のため）
        const progressUpdate = {
          nextIndex: index + 1,  // 次に処理するインデックス
          results: results,
          errors: errors,
          totalFacilities: uniqueFacilities.length
        };
        properties.setProperty('pdfCreationProgress', JSON.stringify(progressUpdate));
        
        // 通常の処理間隔（2秒）
        if (index < uniqueFacilities.length - 1) {
          Utilities.sleep(2000);
        }
        
      } catch (error: any) {
        retryCount++;
        
        // 429エラー（レート制限）の場合
        if (error.message && error.message.includes('429')) {
          if (retryCount < maxRetries) {
            const waitTime = retryCount * 20000; // 20秒、40秒、60秒と増やす
            console.log(`レート制限エラー。${waitTime/1000}秒後にリトライします... (${retryCount}/${maxRetries})`);
            Utilities.sleep(waitTime);
          } else {
            console.error(`エラー: ${facilityName} - レート制限のため処理をスキップ`);
            errors.push({
              facility: facilityName as string,
              error: 'レート制限エラー（リトライ回数超過）',
              status: 'error'
            });
          }
        } else {
          // その他のエラーの場合
          console.error(`エラー: ${facilityName} - ${error.message}`);
          errors.push({
            facility: facilityName as string,
            error: error.message,
            status: 'error'
          });
          break; // リトライしない
        }
      }
    }
  }
  
  // 処理が全て完了したら進捗をクリア
  properties.deleteProperty('pdfCreationProgress');
  console.log('全施設の処理が完了しました。進捗情報をクリアしました。');
  
  // 結果サマリー
  console.log('\n=== PDF作成完了 ===');
  console.log(`成功: ${results.length}件`);
  console.log(`エラー: ${errors.length}件`);
  
  if (errors.length > 0) {
    console.log('\n=== エラー詳細 ===');
    errors.forEach(e => {
      console.log(`${e.facility}: ${e.error}`);
    });
  }
  
  return {
    success: results,
    errors: errors,
    summary: {
      total: uniqueFacilities.length,
      succeeded: results.length,
      failed: errors.length
    }
  };
}

