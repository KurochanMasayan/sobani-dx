/**
 * 歯科診療スケジュール管理システム - GAS
 * AppSheet Call a script 連携
 */

// ===========================================
// AppSheet から呼び出す関数
// ===========================================

/**
 * テスト用関数
 * @param {Object} params - AppSheetから渡されるパラメータ
 * @return {Object} 結果
 */
function testFunction(params) {
  Logger.log('testFunction called: ' + JSON.stringify(params));
  return {
    success: true,
    message: 'Test OK',
    timestamp: new Date().toISOString()
  };
}

/**
 * サンプル：スケジュール更新時の処理
 * @param {Object} params - { scheduleId, patientId, ... }
 * @return {Object} 結果
 */
function onScheduleUpdate(params) {
  Logger.log('onScheduleUpdate: ' + JSON.stringify(params));

  // 必要な処理をここに追加

  return {
    success: true,
    scheduleId: params.scheduleId
  };
}

/**
 * サンプル：算定フラグ更新時の処理
 * @param {Object} params - { scheduleId, billingType, ... }
 * @return {Object} 結果
 */
function onBillingUpdate(params) {
  Logger.log('onBillingUpdate: ' + JSON.stringify(params));

  // 必要な処理をここに追加

  return {
    success: true,
    billingType: params.billingType
  };
}
