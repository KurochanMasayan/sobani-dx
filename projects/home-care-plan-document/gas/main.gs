/**
 * 在宅診療計画書 自動作成エントリーポイント
 *
 * Script Properties で指定する値:
 *  - PATIENT_CSV_FOLDER_ID: 患者マスタ CSV を配置するフォルダ
 *  - OUTPUT_FOLDER_ID: 事業所別スプレッドシートを保存するフォルダ
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('在宅診療計画書')
    .addItem('計画書生成（開始）', 'startPlanGeneration')
    .addToUi();
}

function startPlanGeneration() {
  cancelContinuationTriggers();
  clearProcessingQueue();
  runPlanGeneration({ resetQueue: true });
}

function continuePlanGeneration() {
  runPlanGeneration({ resetQueue: false });
}

function runPlanGeneration(options = { resetQueue: false }) {
  try {
    const folders = getFolderIds();
    const context = buildProcessingContext(folders);

    if (context.patientMap.size === 0) {
      notify('患者CSVが読み込めませんでした。フォルダと文字コードをご確認ください。');
      updateStatus({ message: '患者CSV未読込', processed: 0, remaining: 0, totalQueue: 0 });
      return;
    }

    if (Object.keys(context.visitsByPatient).length === 0) {
      clearProcessingQueue();
      notify('処理対象の患者データが見つかりませんでした。');
      updateStatus({ message: '患者データ無し', processed: 0, remaining: 0, totalQueue: 0 });
      return;
    }

    let queue = getPendingQueue();
    let yearMonth = getSavedYearMonth();
    let runId = getSavedRunId();

    if (queue.length === 0 || options.resetQueue) {
      queue = Object.keys(context.visitsByPatient);
      if (queue.length === 0) {
        notify('処理対象の患者が見つかりませんでした。');
        updateStatus({ message: '処理対象無し', processed: 0, remaining: 0, totalQueue: 0 });
        cancelContinuationTriggers();
        return;
      }
      // 施設名でソート（同じ施設の患者が連続して処理されるように）
      queue = sortQueueByFacility(queue, context.patientMap);
      // 新規開始時は現在の年月と実行IDを保存
      yearMonth = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
      runId = generateRunId();
      // 初回保存時はpatientMapを渡してスプレッドシートに施設名も記録
      savePendingQueue(queue, context.patientMap);
      saveYearMonth(yearMonth);
      saveRunId(runId);
      logInfo('新規実行を開始しました', { runId, yearMonth, queueSize: queue.length });
    } else {
      queue = queue.filter(id => context.visitsByPatient[id]);
      if (queue.length === 0) {
        clearProcessingQueue();
        notify('キューに残っていた患者が見つからなかったため、処理をリセットしました。');
        updateStatus({ message: 'キュー無効のためリセット', processed: 0, remaining: 0, totalQueue: 0 });
        cancelContinuationTriggers();
        return;
      }
      // 再開時に年月・実行IDが保存されていなければ生成（フォールバック）
      if (!yearMonth) {
        yearMonth = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
        saveYearMonth(yearMonth);
      }
      if (!runId) {
        runId = generateRunId();
        saveRunId(runId);
        logInfo('実行IDが見つからなかったため新規生成しました', { runId });
      }
      logInfo('処理を再開しました', { runId, yearMonth, remainingQueue: queue.length });
    }

    const registrySheet = getRegistrySheet();
    const { processedIds, remainingQueue } = processQueue(
      context,
      queue,
      registrySheet,
      folders.output,
      yearMonth,
      runId
    );

    savePendingQueue(remainingQueue);

    if (processedIds.length === 0) {
      notify(`処理可能な患者が見つかりませんでした。残り: ${remainingQueue.length} 名`);
      updateStatus({
        message: '対象なしまたは上限到達',
        processed: 0,
        remaining: remainingQueue.length,
        totalQueue: queue.length
      });
      if (remainingQueue.length > 0) {
        scheduleContinuation();
      } else {
        clearProcessingQueue();
        cancelContinuationTriggers();
      }
    } else if (remainingQueue.length > 0) {
      notify(`今回 ${processedIds.length} 名を処理しました。残り ${remainingQueue.length} 名です。`);
      updateStatus({
        message: '処理途中',
        processed: processedIds.length,
        remaining: remainingQueue.length,
        totalQueue: queue.length
      });
      scheduleContinuation();
    } else {
      notify(`全 ${processedIds.length} 名の処理が完了しました。`);
      updateStatus({
        message: '完了',
        processed: processedIds.length,
        remaining: 0,
        totalQueue: queue.length
      });
      clearProcessingQueue();
      cancelContinuationTriggers();
      cleanupSourceCsv(folders);
    }
  } catch (error) {
    logError('runPlanGeneration', error);
    notify(`エラーが発生しました: ${error.message}`);
    updateStatus({ message: `エラー: ${error.message}`, processed: 0, remaining: 0, totalQueue: 0 });
    cancelContinuationTriggers();
  }
}

function processQueue(context, queue, registrySheet, outputFolderId, yearMonth, runId) {
  const start = Date.now();
  const processedIds = [];
  const workingQueue = queue.slice();

  while (workingQueue.length > 0) {
    if (Date.now() - start >= BASE_CONFIG.maxExecutionMs - 15 * 1000) {
      logInfo('実行時間上限が近いため一時停止', {
        processed: processedIds.length,
        remaining: workingQueue.length,
        runId
      });
      break;
    }

    const patientId = workingQueue.shift();
    const patient = context.patientMap.get(patientId);
    const visits = context.visitsByPatient[patientId];

    if (!patient || !visits || visits.length === 0) {
      logInfo('対象データが無いためスキップ', { patientId });
      continue;
    }

    // 個別患者処理（1人の失敗で全体を止めない）
    try {
      const facilityName = normalizeFacilityName(patient.facility) || '不明';
      logInfo('患者処理開始', {
        patientId,
        patientName: patient.name || '名前不明',
        facility: facilityName,
        visitCount: visits.length
      });

      // 保存された年月と実行IDを使用（6分制限対策）
      processPatientVisits(patient, visits, registrySheet, outputFolderId, yearMonth, runId);
      processedIds.push(patientId);
    } catch (patientError) {
      logError('processQueue:患者個別処理', patientError);

      // クォータ制限エラーの場合は処理を中断（翌日再開可能にする）
      if (patientError.message && patientError.message.includes('実行した回数が多すぎます')) {
        logInfo('クォータ制限に達したため処理を中断します', {
          patientId,
          patientName: patient.name || '名前不明'
        });
        // 現在の患者をキューに戻す
        workingQueue.unshift(patientId);
        break;
      }

      logInfo('患者処理をスキップして続行', {
        patientId,
        patientName: patient.name || '名前不明',
        error: patientError.message
      });
      // その他のエラーは次の患者の処理を続行
    }
  }

  return { processedIds, remainingQueue: workingQueue };
}

function cleanupSourceCsv(folders) {
  try {
    const folderId = folders.patientCsv;
    if (!folderId) return;
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.CSV);
    while (files.hasNext()) {
      const file = files.next();
      logInfo('CSVファイルを削除します', { fileName: file.getName() });
      file.setTrashed(true);
    }
  } catch (error) {
    logError('cleanupSourceCsv', error);
  }
}
