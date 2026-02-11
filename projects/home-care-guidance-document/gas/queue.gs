/**
 * キューとトリガー管理
 * キュー情報はスプレッドシートに保存（デバッグ容易化・容量制限回避）
 */

/**
 * 処理キューシートを取得または作成
 * @returns {Sheet} 処理キューシート
 */
function getQueueSheet() {
  const active = getActiveSpreadsheetOrThrow();
  let sheet = active.getSheetByName(BASE_CONFIG.queueSheetName);
  if (!sheet) {
    sheet = active.insertSheet(BASE_CONFIG.queueSheetName);
    sheet.getRange(1, 1, 1, 4).setValues([[
      '順番', '患者ID', '施設名', 'ステータス'
    ]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function scheduleContinuation() {
  logInfo('scheduleContinuation開始');
  try {
    const remaining = getPendingQueue();
    logInfo('getPendingQueue結果', { count: remaining ? remaining.length : 0 });
    if (!remaining || remaining.length === 0) {
      logInfo('残りキューが空のためトリガーをキャンセル');
      cancelContinuationTriggers();
      return;
    }
    cancelContinuationTriggers();
    ScriptApp.newTrigger('continueGuidanceGeneration')
      .timeBased()
      .after(60 * 1000)
      .create();
    logInfo('次回実行をスケジュールしました', { remaining: remaining.length });
  } catch (e) {
    logError('scheduleContinuation', e);
  }
}

function cancelContinuationTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'continueGuidanceGeneration') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * 未処理の患者IDリストを取得（スプレッドシートから）
 * @returns {string[]} 未処理の患者IDの配列
 */
function getPendingQueue() {
  try {
    const sheet = getQueueSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; // ヘッダーのみ

    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    // ステータスが「待機」の患者IDのみ抽出（6桁ゼロパディングで正規化）
    const pending = data
      .filter(row => row[3] === BASE_CONFIG.status.pending)
      .map(row => String(row[1]).padStart(6, '0')); // 患者ID（B列）
    return pending;
  } catch (e) {
    logError('getPendingQueue', e);
    return [];
  }
}

/**
 * ソート済みキューをスプレッドシートに保存
 * @param {string[]} queue - 患者IDの配列（ソート済み）
 * @param {Map} patientMap - 患者ID -> 患者情報のMap（オプション、初回保存時のみ必要）
 */
function savePendingQueue(queue, patientMap) {
  if (!queue || queue.length === 0) {
    clearProcessingQueue();
    return;
  }

  const sheet = getQueueSheet();

  // patientMapが渡された場合は新規保存（全件書き込み）
  if (patientMap) {
    // 既存データをクリア（ヘッダー以外）
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
    }

    // キューデータを書き込み
    const rows = queue.map((patientId, index) => {
      const patient = patientMap.get(patientId);
      const facilityName = patient ? normalizeFacilityName(patient.facility) : '';
      return [index + 1, patientId, facilityName, BASE_CONFIG.status.pending];
    });
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    logInfo('キューをスプレッドシートに保存しました', { count: rows.length });
  } else {
    // patientMapがない場合は、処理済みの患者をマーク
    // remainingQueueに含まれない患者ID = 処理済み
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    // 患者IDを文字列として比較するためSetも文字列化
    const remainingSet = new Set(queue.map(id => String(id)));

    const updatedData = data.map(row => {
      // スプレッドシートから読み込んだ値は数値になる可能性があるため、6桁ゼロパディングで正規化
      const patientId = String(row[1]).padStart(6, '0');
      if (!remainingSet.has(patientId) && row[3] === BASE_CONFIG.status.pending) {
        return [row[0], row[1], row[2], BASE_CONFIG.status.processed];
      }
      return row;
    });

    sheet.getRange(2, 1, updatedData.length, 4).setValues(updatedData);
  }
  SpreadsheetApp.flush();
}

/**
 * 処理キューをクリア
 */
function clearProcessingQueue() {
  try {
    const sheet = getQueueSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
    }
  } catch (e) {
    logError('clearProcessingQueue', e);
  }
  SCRIPT_PROPS.deleteProperty(BASE_CONFIG.yearMonthPropertyKey);
}

/**
 * 処理中の年月を取得
 * @returns {string|null} 保存されている年月、なければnull
 */
function getSavedYearMonth() {
  return SCRIPT_PROPS.getProperty(BASE_CONFIG.yearMonthPropertyKey) || null;
}

/**
 * 処理中の年月を保存
 * @param {string} yearMonth - 年月（例: 2025-12）
 */
function saveYearMonth(yearMonth) {
  if (yearMonth) {
    SCRIPT_PROPS.setProperty(BASE_CONFIG.yearMonthPropertyKey, yearMonth);
  }
}
