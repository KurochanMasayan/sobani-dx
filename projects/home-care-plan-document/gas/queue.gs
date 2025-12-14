/**
 * キューとトリガー管理
 */
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

function getPendingQueue() {
  const value = SCRIPT_PROPS.getProperty(BASE_CONFIG.queuePropertyKey);
  if (!value) return [];
  try {
    const parsed = JSON.parse(value);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    logError('getPendingQueue', e);
    return [];
  }
}

function savePendingQueue(queue) {
  if (!queue || queue.length === 0) {
    clearProcessingQueue();
    return;
  }
  SCRIPT_PROPS.setProperty(BASE_CONFIG.queuePropertyKey, JSON.stringify(queue));
}

function clearProcessingQueue() {
  SCRIPT_PROPS.deleteProperty(BASE_CONFIG.queuePropertyKey);
}
