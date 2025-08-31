/**
 * Googleカレンダー連携機能
 */

/**
 * スプレッドシートから日付を取得するヘルパー関数
 */
function getDateFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.CALENDAR.DATE_SOURCE.SHEET_NAME);
  
  if (!sourceSheet) {
    throw new Error(`「${CONFIG.CALENDAR.DATE_SOURCE.SHEET_NAME}」シートが見つかりません`);
  }
  
  const startDateValue = sourceSheet.getRange(CONFIG.CALENDAR.DATE_SOURCE.START_DATE_CELL).getValue();
  const endDateValue = sourceSheet.getRange(CONFIG.CALENDAR.DATE_SOURCE.END_DATE_CELL).getValue();
  
  if (!startDateValue || !endDateValue) {
    throw new Error(`${CONFIG.CALENDAR.DATE_SOURCE.SHEET_NAME}シートの${CONFIG.CALENDAR.DATE_SOURCE.START_DATE_CELL}（開始日）または${CONFIG.CALENDAR.DATE_SOURCE.END_DATE_CELL}（終了日）が設定されていません`);
  }
  
  // 日付をJSTで処理（時刻を0:00:00に設定）
  const startDate = new Date(startDateValue);
  startDate.setHours(0, 0, 0, 0);
  
  // 終了日は23:59:59.999に設定して、その日の終わりまで含める
  const endDate = new Date(endDateValue);
  endDate.setHours(23, 59, 59, 999);
  
  return { startDate, endDate };
}

/**
 * Googleカレンダーからイベントデータを取得
 * @param {string} calendarId - カレンダーID
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Array} イベントデータの配列
 */
function getCalendarEvents(calendarId, startDate, endDate) {
  try {
    // カレンダーを取得
    const calendar = CalendarApp.getCalendarById(calendarId);
    
    if (!calendar) {
      throw new Error(`カレンダー（ID: ${calendarId}）が見つかりません。`);
    }
    
    // イベントを取得
    const events = calendar.getEvents(startDate, endDate);
    
    // イベントデータを整形
    const eventData = events.map(event => {
      return {
        title: event.getTitle(),
        startTime: event.getStartTime(),
        endTime: event.getEndTime(),
        location: event.getLocation(),
        description: event.getDescription(),
        allDay: event.isAllDayEvent(),
        creators: event.getCreators().join(', '),
        id: event.getId()
      };
    });
    
    console.log(`${events.length}件のイベントを取得しました。`);
    return eventData;
    
  } catch (error) {
    console.error(`カレンダーイベント取得エラー: ${error.message}`);
    throw error;
  }
}

/**
 * カレンダーイベントをスプレッドシートにエクスポート
 * @param {string} targetSheetName - エクスポート先のシート名
 * @param {string} calendarId - カレンダーID
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 */
function exportCalendarToSheet(targetSheetName, calendarId, startDate, endDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // シートを取得または作成
  let sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(targetSheetName);
  }
  
  // A~E列のみクリア（他の列のデータは保持）
  const lastRow = sheet.getMaxRows();
  if (lastRow > 0) {
    sheet.getRange(2, 1, lastRow, 5).clear();
  }
  
  // イベントデータを取得
  const events = getCalendarEvents(calendarId, startDate, endDate);
  
  if (events.length === 0) {
    sheet.getRange(2, 1).setValue('指定期間にイベントはありません');
    console.log('指定期間にイベントはありません');
    return {
      sheetName: targetSheetName,
      eventCount: 0
    };
  }
  
  // ヘッダーを設定（元のコードに合わせて変更）
  const headers = ['日付', '開始時刻', '終了時刻', 'タイトル', 'コメント'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // データを書き込み（元のコードのフォーマットに合わせる）
  const data = events.map(event => [
    Utilities.formatDate(event.startTime, 'JST', 'yyyy/M/d'),
    event.allDay ? '終日' : Utilities.formatDate(event.startTime, 'JST', 'H:mm:ss'),
    event.allDay ? '終日' : Utilities.formatDate(event.endTime, 'JST', 'H:mm:ss'),
    event.title,
    event.description || ''  // 説明をE列に直接記入
  ]);
  
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  
  console.log(`${events.length}件のイベントをシート「${targetSheetName}」にエクスポートしました。`);
  
  return {
    sheetName: targetSheetName,
    eventCount: events.length
  };
}

/**
 * カレンダーデータを同期（ボタン用）
 */
function syncCalendarData() {
  try {
    // 設定からカレンダーIDを取得
    const calendarIds = CONFIG.CALENDAR.CALENDAR_IDS;
    const outputSheet = CONFIG.CALENDAR.OUTPUT_SHEET;
    
    // シートから日付を取得
    const { startDate, endDate } = getDateFromSheet();
    
    console.log(`期間: ${Utilities.formatDate(startDate, 'JST', 'yyyy/MM/dd')} - ${Utilities.formatDate(endDate, 'JST', 'yyyy/MM/dd')}`);
    
    if (calendarIds.length === 1) {
      // 単一カレンダーの場合
      const result = exportCalendarToSheet(outputSheet, calendarIds[0], startDate, endDate);
      console.log(`カレンダー同期完了: ${result.eventCount}件のイベントを取得`);
      console.log(`スプレッドシート: ${SpreadsheetApp.getActiveSpreadsheet().getName()}`);
      console.log(`シート: ${outputSheet}`);
      return result;
    } else {
      // 複数カレンダーの場合
      const allEvents = [];
      
      calendarIds.forEach(calendarId => {
        try {
          const events = getCalendarEvents(calendarId, startDate, endDate);
          events.forEach(event => {
            event.calendarId = calendarId;
            allEvents.push(event);
          });
        } catch (error) {
          console.error(`カレンダー ${calendarId} の取得に失敗: ${error.message}`);
        }
      });
      
      // 開始時刻でソート
      allEvents.sort((a, b) => a.startTime - b.startTime);
      
      console.log(`複数カレンダー同期完了: ${allEvents.length}件のイベントを取得`);
      
      // スプレッドシートに書き込み
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName(outputSheet);
      if (!sheet) {
        sheet = ss.insertSheet(outputSheet);
      }
      
      // A~F列のみクリア
      const lastRow = sheet.getMaxRows();
      if (lastRow > 0) {
        sheet.getRange(2, 1, lastRow, 6).clear();
      }
      
      // ヘッダーとデータを書き込み
      if (allEvents.length > 0) {
        const headers = ['日付', '開始時刻', '終了時刻', 'タイトル', 'コメント', 'カレンダーID'];
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        
        const data = allEvents.map(event => [
          Utilities.formatDate(event.startTime, 'JST', 'yyyy/M/d'),
          event.allDay ? '終日' : Utilities.formatDate(event.startTime, 'JST', 'H:mm:ss'),
          event.allDay ? '終日' : Utilities.formatDate(event.endTime, 'JST', 'H:mm:ss'),
          event.title,
          event.description || '',
          event.calendarId
        ]);
        
        sheet.getRange(2, 1, data.length, headers.length).setValues(data);
      } else {
        sheet.getRange(2, 1).setValue('指定期間にイベントはありません');
      }
      
      return {
        sheetName: outputSheet,
        eventCount: allEvents.length,
        calendarCount: calendarIds.length
      };
    }
    
  } catch (error) {
    console.error(`同期エラー: ${error.message}`);
    throw error;
  }
}

/**
 * 一括実行機能（元のコードのexecuteAllに相当）
 */
function executeAll() {
  try {
    console.log('===== 処理開始 =====');
    
    // 1. カレンダーデータを取得してスプレッドシートに出力
    console.log('1. カレンダーデータを取得中...');
    const calendarResult = syncCalendarData();
    
    // 2. CSV出力
    console.log('2. CSVファイルを生成中...');
    const csvResult = downloadCsvButton();
    
    // 3. 施設カレンダーPDF作成
    console.log('3. 施設カレンダーPDFファイルを生成中...');
    // createSinglePdfButtonまたはcreateAllPdfsButtonを実行
    
    console.log('===== 処理完了 =====');
    console.log(`カレンダーイベント: ${calendarResult.eventCount}件`);
    if (csvResult) {
      console.log(`CSVファイル: ${csvResult.fileName}`);
    }
    
    return {
      calendar: calendarResult,
      csv: csvResult
    };
    
  } catch (error) {
    console.error('一括実行エラー:', error);
    throw error;
  }
}