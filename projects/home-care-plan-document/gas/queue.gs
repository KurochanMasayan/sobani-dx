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
  const remaining = getPendingQueue();
  if (!remaining || remaining.length === 0) {
    cancelContinuationTriggers();
    return;
  }
  cancelContinuationTriggers();
  ScriptApp.newTrigger('continuePlanGeneration')
    .timeBased()
    .after(60 * 1000)
    .create();
  logInfo('次回実行をスケジュールしました', { remaining: remaining.length });
}

function cancelContinuationTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'continuePlanGeneration') {
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
    // ステータスが「待機」の患者IDのみ抽出
    const pending = data
      .filter(row => row[3] === '待機')
      .map(row => row[1]); // 患者ID（B列）
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
      return [index + 1, patientId, facilityName, '待機'];
    });
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    logInfo('キューをスプレッドシートに保存しました', { count: rows.length });
  } else {
    // patientMapがない場合は、処理済みの患者をマーク
    // remainingQueueに含まれない患者ID = 処理済み
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    const remainingSet = new Set(queue);

    const updatedData = data.map(row => {
      const patientId = row[1];
      if (!remainingSet.has(patientId) && row[3] === '待機') {
        return [row[0], row[1], row[2], '処理済み'];
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
  SCRIPT_PROPS.deleteProperty(BASE_CONFIG.runIdPropertyKey);
}

/**
 * 実行IDを生成
 * 同一バッチ処理を識別するための一意のID
 * @returns {string} 実行ID（例: 2025-12-20T10:30:00_a1b2c3d4）
 */
function generateRunId() {
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd\'T\'HH:mm:ss');
  const random = Utilities.getUuid().substring(0, 8);
  return `${timestamp}_${random}`;
}

/**
 * 実行IDを保存
 * @param {string} runId - 実行ID
 */
function saveRunId(runId) {
  if (runId) {
    SCRIPT_PROPS.setProperty(BASE_CONFIG.runIdPropertyKey, runId);
  }
}

/**
 * 保存されている実行IDを取得
 * @returns {string|null} 実行ID、なければnull
 */
function getSavedRunId() {
  return SCRIPT_PROPS.getProperty(BASE_CONFIG.runIdPropertyKey) || null;
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
