/**
 * 居宅療養管理指導書 自動作成エントリーポイント
 *
 * Script Properties で指定する値:
 *  - PATIENT_CSV_FOLDER_ID: Ⅰ の患者マスタ CSV を配置するフォルダ
 *  - CALENDAR_CSV_FOLDER_ID: Ⅱ のカレンダー CSV を配置するフォルダ
 *  - OUTPUT_FOLDER_ID: 事業所別スプレッドシートを保存するフォルダ
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('居療管指導書')
    .addItem('指導書生成（開始）', 'startGuidanceGeneration')
    .addSeparator()
    .addItem('選択患者を処理', 'processSelectedPatients')
    .addToUi();
}

function startGuidanceGeneration() {
  cancelContinuationTriggers();
  clearProcessingQueue();
  runGuidanceGeneration({ resetQueue: true, source: 'manual' });
}

function continueGuidanceGeneration() {
  runGuidanceGeneration({ resetQueue: false, source: 'trigger' });
}

function runGuidanceGeneration(options = { resetQueue: false, source: 'manual' }) {
  try {
    const folders = getFolderIds();
    const context = buildProcessingContext(folders);

    if (context.patientMap.size === 0) {
      notify('患者CSVが読み込めませんでした。フォルダと文字コードをご確認ください。');
      updateStatus({ message: '患者CSV未読込', processed: 0, remaining: 0, totalQueue: 0 });
      return;
    }

    // 3件以上の訪問がある患者を訪問選択シートに登録
    const multipleVisitsCount = Object.keys(context.multipleVisitsPatients || {}).length;
    if (multipleVisitsCount > 0 && options.resetQueue) {
      registerMultipleVisitsPatients(context.multipleVisitsPatients, context.patientMap);
      notify(`${multipleVisitsCount} 名の患者は訪問が3件以上あります。「訪問選択」シートで訪問回を選択し、「選択患者を処理」を実行してください。`);
    }

    if (Object.keys(context.visitsByPatient).length === 0 && multipleVisitsCount === 0) {
      clearProcessingQueue();
      notify('定期訪問の対象データが見つかりませんでした。');
      updateStatus({ message: '定期訪問データ無し', processed: 0, remaining: 0, totalQueue: 0 });
      return;
    }

    // 2件以下の患者のみ処理
    if (Object.keys(context.visitsByPatient).length === 0) {
      notify('2件以下の定期訪問がある患者はいません。「訪問選択」シートで選択後、「選択患者を処理」を実行してください。');
      updateStatus({ message: '訪問選択待ち', processed: 0, remaining: multipleVisitsCount, totalQueue: multipleVisitsCount });
      return;
    }

    let queue = getPendingQueue();
    let yearMonth = getSavedYearMonth();

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
      // 新規開始時はカレンダーCSVの年月を使用（取得できなければ現在日付）
      yearMonth = context.calendarYearMonth || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
      // 初回保存時はpatientMapを渡してスプレッドシートに施設名も記録
      savePendingQueue(queue, context.patientMap);
      saveYearMonth(yearMonth);
      logInfo('新規実行を開始しました', { yearMonth, calendarYearMonth: context.calendarYearMonth, queueSize: queue.length });
    } else {
      queue = queue.filter(id => context.visitsByPatient[id]);
      if (queue.length === 0) {
        clearProcessingQueue();
        notify('キューに残っていた患者が見つからなかったため、処理をリセットしました。');
        updateStatus({ message: 'キュー無効のためリセット', processed: 0, remaining: 0, totalQueue: 0 });
        cancelContinuationTriggers();
        return;
      }
      // 再開時に年月が保存されていなければカレンダーCSVから取得（フォールバック）
      if (!yearMonth) {
        yearMonth = context.calendarYearMonth || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
        saveYearMonth(yearMonth);
      }
      logInfo('処理を再開しました', { yearMonth, calendarYearMonth: context.calendarYearMonth, remainingQueue: queue.length });
    }

    const registrySheet = getRegistrySheet();
    const { processedIds, remainingQueue } = processQueue(
      context,
      queue,
      registrySheet,
      folders.output,
      yearMonth
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
      let message = `全 ${processedIds.length} 名の処理が完了しました。`;
      if (multipleVisitsCount > 0) {
        message += ` 訪問選択待ち: ${multipleVisitsCount} 名`;
      }
      notify(message);
      updateStatus({
        message: multipleVisitsCount > 0 ? '完了（訪問選択待ちあり）' : '完了',
        processed: processedIds.length,
        remaining: multipleVisitsCount,
        totalQueue: queue.length + multipleVisitsCount
      });
      clearProcessingQueue();
      cancelContinuationTriggers();
      // 訪問選択待ちがなければCSVを削除
      if (multipleVisitsCount === 0) {
        cleanupSourceCsv(folders);
      }
    }
  } catch (error) {
    logError('runGuidanceGeneration', error);
    notify(`エラーが発生しました: ${error.message}`);
    updateStatus({ message: `エラー: ${error.message}`, processed: 0, remaining: 0, totalQueue: 0 });
    cancelContinuationTriggers();
  }
}

/**
 * 訪問選択シートで選択された患者を処理
 */
function processSelectedPatients() {
  try {
    const folders = getFolderIds();
    const context = buildProcessingContext(folders);

    if (context.patientMap.size === 0) {
      notify('患者CSVが読み込めませんでした。フォルダと文字コードをご確認ください。');
      return;
    }

    // 訪問選択シートから選択内容を取得
    const selections = getVisitSelections();
    const selectionCount = Object.keys(selections).length;

    if (selectionCount === 0) {
      notify('訪問選択シートに未処理の患者がいません。');
      return;
    }

    // 3件以上の患者データを取得
    const multipleVisitsPatients = context.multipleVisitsPatients || {};

    // 選択に基づいて訪問プランを構築
    const selectedQueue = [];
    const selectedVisitsByPatient = {};

    Object.keys(selections).forEach(patientId => {
      const allVisits = multipleVisitsPatients[patientId];
      if (!allVisits) {
        logInfo('訪問データが見つかりません', { patientId });
        return;
      }

      const { visit1Number, visit2Number } = selections[patientId];
      logInfo('訪問選択情報', { patientId, visit1Number, visit2Number, allVisitsCount: allVisits.length });
      const selectedVisits = selectVisitsForDocument(allVisits, visit1Number, visit2Number);
      logInfo('選択された訪問', { patientId, selectedVisitsCount: selectedVisits.length });

      if (selectedVisits.length > 0) {
        selectedVisitsByPatient[patientId] = selectedVisits;
        selectedQueue.push(patientId);
      }
    });

    if (selectedQueue.length === 0) {
      notify('処理対象の患者が見つかりませんでした。訪問選択シートを確認してください。');
      return;
    }

    // 施設名でソート
    const sortedQueue = sortQueueByFacility(selectedQueue, context.patientMap);

    // 年月を取得または生成（カレンダーCSVの年月を優先）
    let yearMonth = getSavedYearMonth();
    if (!yearMonth) {
      yearMonth = context.calendarYearMonth || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM');
      saveYearMonth(yearMonth);
    }
    logInfo('processSelectedPatients: 年月設定', { yearMonth, calendarYearMonth: context.calendarYearMonth });

    const registrySheet = getRegistrySheet();
    logInfo('processSelectedPatients: 台帳シート取得', { sheetName: registrySheet.getName() });

    // 処理実行
    const start = Date.now();
    const processedIds = [];

    for (const patientId of sortedQueue) {
      if (isNearExecutionLimit(start)) {
        logInfo('実行時間上限が近いため一時停止', {
          processed: processedIds.length,
          remaining: sortedQueue.length - processedIds.length
        });
        break;
      }

      const patient = context.patientMap.get(patientId);
      const visits = selectedVisitsByPatient[patientId];

      if (!patient || !visits || visits.length === 0) {
        logInfo('対象データが無いためスキップ', { patientId });
        continue;
      }

      try {
        const facilityName = normalizeFacilityName(patient.facility) || BASE_CONFIG.fallbackFacilityName;
        logInfo('選択患者処理開始', {
          patientId,
          patientName: patient.name || '名前不明',
          facility: facilityName,
          rawFacility: patient.facility,
          selectedVisits: visits.map(v => `${v.visitNumber}回目`),
          yearMonth,
          outputFolder: folders.output
        });

        processPatientVisits(patient, visits, registrySheet, folders.output, yearMonth);
        logInfo('選択患者処理完了', { patientId, facilityName });
        processedIds.push(patientId);

        // 訪問選択シートを処理済みにマーク
        markVisitSelectionAsProcessed(patientId);
      } catch (patientError) {
        logError('processSelectedPatients:患者個別処理', patientError);
        logInfo('患者処理をスキップして続行', {
          patientId,
          error: patientError.message
        });
      }
    }

    const remainingCount = getPendingVisitSelectionCount();

    if (processedIds.length === 0) {
      notify('処理可能な患者が見つかりませんでした。');
    } else if (remainingCount > 0) {
      notify(`${processedIds.length} 名を処理しました。残り ${remainingCount} 名は再度実行してください。`);
    } else {
      notify(`全 ${processedIds.length} 名の処理が完了しました。`);
      // 全て完了したらCSVと訪問選択シートをクリア
      cleanupSourceCsv(folders);
      clearVisitSelectionSheet();
    }

    updateStatus({
      message: remainingCount > 0 ? '選択患者処理中' : '選択患者処理完了',
      processed: processedIds.length,
      remaining: remainingCount,
      totalQueue: selectionCount
    });
  } catch (error) {
    logError('processSelectedPatients', error);
    notify(`エラーが発生しました: ${error.message}`);
    updateStatus({ message: `エラー: ${error.message}`, processed: 0, remaining: 0, totalQueue: 0 });
  }
}

function processQueue(context, queue, registrySheet, outputFolderId, yearMonth) {
  const start = Date.now();
  const processedIds = [];
  const workingQueue = queue.slice();

  while (workingQueue.length > 0) {
    if (isNearExecutionLimit(start)) {
      logInfo('実行時間上限が近いため一時停止', {
        processed: processedIds.length,
        remaining: workingQueue.length
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
      const facilityName = normalizeFacilityName(patient.facility) || BASE_CONFIG.fallbackFacilityName;
      logInfo('患者処理開始', {
        patientId,
        patientName: patient.name || '名前不明',
        facility: facilityName,
        visitCount: visits.length
      });

      // 保存された年月を使用（月ごとに同じファイルに追記）
      processPatientVisits(patient, visits, registrySheet, outputFolderId, yearMonth);
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
    ['patientCsv', 'calendarCsv'].forEach(key => {
      const folderId = folders[key];
      if (!folderId) return;
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFilesByType(MimeType.CSV);
      while (files.hasNext()) {
        const file = files.next();
        logInfo('CSVファイルを削除します', { folderKey: key, fileName: file.getName() });
        file.setTrashed(true);
      }
    });
  } catch (error) {
    logError('cleanupSourceCsv', error);
  }
}
