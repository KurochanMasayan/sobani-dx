// カレンダーデータをスプレッドシートにエクスポートする関数
function exportCalendarToSpreadsheet(): void {
  try {
    // スプレッドシートを取得
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 使い方シートから日付を取得
    const usageSheet = spreadsheet.getSheetByName('使い方');
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
    const startDate = new Date(startDateValue as string);
    startDate.setHours(0, 0, 0, 0);
    
    // 終了日は23:59:59.999に設定して、その日の終わりまで含める
    const endDate = new Date(endDateValue as string);
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
    const data: any[][] = [];
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