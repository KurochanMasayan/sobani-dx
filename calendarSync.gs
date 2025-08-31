/**
 * Googleカレンダー連携機能
 */

/**
 * Googleカレンダーからイベントデータを取得
 * @param {string} calendarId - カレンダーID（省略時はデフォルトカレンダー）
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Array} イベントデータの配列
 */
function getCalendarEvents(calendarId = 'primary', startDate = null, endDate = null) {
  // デフォルトの期間を設定（今月）
  if (!startDate) {
    startDate = new Date();
    startDate.setDate(1);
    startDate.setHours(0, 0, 0, 0);
  }
  
  if (!endDate) {
    endDate = new Date(startDate);
    endDate.setMonth(endDate.getMonth() + 1);
    endDate.setDate(0);
    endDate.setHours(23, 59, 59, 999);
  }
  
  try {
    // カレンダーを取得
    const calendar = calendarId === 'primary' 
      ? CalendarApp.getDefaultCalendar()
      : CalendarApp.getCalendarById(calendarId);
    
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
function exportCalendarToSheet(targetSheetName = 'カレンダーデータ', calendarId = 'primary', startDate = null, endDate = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // シートを取得または作成
  let sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(targetSheetName);
  } else {
    sheet.clear();
  }
  
  // イベントデータを取得
  const events = getCalendarEvents(calendarId, startDate, endDate);
  
  if (events.length === 0) {
    console.log('イベントが見つかりませんでした。');
    return;
  }
  
  // ヘッダーを設定
  const headers = [
    'タイトル',
    '開始日時',
    '終了日時',
    '場所',
    '説明',
    '終日',
    '作成者',
    'イベントID'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  
  // データを書き込み
  const data = events.map(event => [
    event.title,
    event.startTime,
    event.endTime,
    event.location || '',
    event.description || '',
    event.allDay ? 'はい' : 'いいえ',
    event.creators || '',
    event.id
  ]);
  
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  
  // 列幅を自動調整
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  console.log(`${events.length}件のイベントをシート「${targetSheetName}」にエクスポートしました。`);
  
  return {
    sheetName: targetSheetName,
    eventCount: events.length,
    period: {
      start: startDate || '今月初',
      end: endDate || '今月末'
    }
  };
}

/**
 * 複数のカレンダーからイベントを統合して取得
 * @param {Array<string>} calendarIds - カレンダーIDの配列
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @return {Array} 統合されたイベントデータ
 */
function getMultipleCalendarEvents(calendarIds, startDate = null, endDate = null) {
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
  
  return allEvents;
}

/**
 * カレンダーデータをCSVとしてエクスポート（ボタン用）
 */
function exportCalendarToCsvFile() {
  try {
    // 今月のイベントを取得
    const events = getCalendarEvents();
    
    if (events.length === 0) {
      SpreadsheetApp.getUi().alert('イベントが見つかりませんでした。');
      return;
    }
    
    // CSVデータを作成
    const csvData = [
      ['タイトル', '開始日時', '終了日時', '場所', '説明', '終日'],
      ...events.map(event => [
        event.title,
        Utilities.formatDate(event.startTime, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'),
        Utilities.formatDate(event.endTime, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'),
        event.location || '',
        event.description || '',
        event.allDay ? 'はい' : 'いいえ'
      ])
    ];
    
    // CSVファイルとして保存
    const fileName = `カレンダーエクスポート_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd')}.csv`;
    const csvContent = csvData.map(row => row.map(cell => {
      const value = String(cell);
      if (value.includes(',') || value.includes('\n') || value.includes('"')) {
        return '"' + value.replace(/"/g, '""') + '"';
      }
      return value;
    }).join(',')).join('\n');
    
    const blob = Utilities.newBlob('\uFEFF' + csvContent, 'text/csv', fileName);
    const file = DriveApp.createFile(blob);
    
    console.log(`カレンダーデータをCSVファイルとして保存しました: ${fileName}`);
    
    // ダウンロードリンクを表示
    const downloadUrl = file.getDownloadUrl();
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'エクスポート完了',
      `カレンダーデータをCSVファイルとして保存しました。\n\nファイル名: ${fileName}\n\nダウンロードURL:\n${downloadUrl}`,
      ui.ButtonSet.OK
    );
    
    return file;
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`エラー: ${error.message}`);
    throw error;
  }
}

/**
 * 施設ごとのカレンダーデータを取得（ボタン用）
 */
function syncCalendarData() {
  try {
    const result = exportCalendarToSheet();
    
    SpreadsheetApp.getUi().alert(
      '同期完了',
      `カレンダーデータを同期しました。\n\nシート名: ${result.sheetName}\nイベント数: ${result.eventCount}件`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert(`同期エラー: ${error.message}`);
  }
}