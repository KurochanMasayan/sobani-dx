/**
 * ステータスシートの更新
 */
function getStatusSheet() {
  const active = getActiveSpreadsheetOrThrow();
  let sheet = active.getSheetByName(BASE_CONFIG.statusSheetName);
  if (!sheet) {
    sheet = active.insertSheet(BASE_CONFIG.statusSheetName);
    sheet.getRange(1, 1, 1, 4).setValues([[
      '更新時刻', 'ステータス', '今回処理人数', '残り人数'
    ]]);
  }
  return sheet;
}

function updateStatus({ message, processed, remaining, totalQueue }) {
  try {
    const sheet = getStatusSheet();
    sheet.getRange(2, 1, 1, 4).setValues([[
      new Date(),
      message || '',
      processed != null ? processed : '',
      remaining != null ? remaining : ''
    ]]);
    if (typeof totalQueue === 'number') {
      sheet.getRange(3, 1, 1, 2).setValues([[ '総キュー数', totalQueue ]]);
    }
  } catch (err) {
    logError('updateStatus', err);
  }
}
