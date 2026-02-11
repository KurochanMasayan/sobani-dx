/**
 * CSVデータの読み込みと加工
 */
function buildProcessingContext(folders) {
  const patientRows = loadCsvRows(folders.patientCsv);
  const patientMap = buildPatientMap(patientRows);
  const calendarRows = loadCsvRows(folders.calendarCsv);
  const { visitsByPatient, multipleVisitsPatients, calendarYearMonth } = buildVisitPlan(calendarRows, patientMap);
  return { patientMap, visitsByPatient, multipleVisitsPatients, calendarYearMonth };
}

function loadCsvRows(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.CSV);
  let latestFile = null;
  while (files.hasNext()) {
    const file = files.next();
    if (!latestFile || file.getLastUpdated().getTime() > latestFile.getLastUpdated().getTime()) {
      latestFile = file;
    }
  }
  if (!latestFile) {
    logInfo('CSVファイルが見つかりません', { folderId });
    return [];
  }
  logInfo('CSVを読み込みます', { fileName: latestFile.getName() });
  const content = latestFile.getBlob().getDataAsString(BASE_CONFIG.encoding);
  return Utilities.parseCsv(content);
}

function buildPatientMap(rows) {
  const map = new Map();
  if (rows.length === 0) return map;

  const header = rows[0];
  const normalizedHeader = header.map(value => normalizeCircledNumbers(value || ''));
  const index = name => normalizedHeader.indexOf(normalizeCircledNumbers(name));

  const col = {
    patientId: index('患者ID'),
    name: index('患者名'),
    birth: index('生年月日'),
    age: index('年齢'),
    sex: index('性別'),
    zip: index('郵便番号'),
    address: index('住所'),
    diagnosis: index('診断名'),
    facility: index('居宅介護支援事業所'),
    bedridden: index('寝たきり度'),
    dementia: index('認知度'),
    careLevel: index('介護認定'),
    caution: index('日常生活や介護サービスを利用する上での留意点'),
    visit1Stable: index('定期訪問①（著変なし）'),
    visit1Change: index('定期訪問①（変更あり）'),
    visit2Stable: index('定期訪問②（著変なし）'),
    visit2Change: index('定期訪問②（変更あり）')
  };

  rows.slice(1).forEach(row => {
    const rawPatientId = (row[col.patientId] || '').trim();
    if (!rawPatientId) return;
    // 患者IDを6桁ゼロパディングに正規化（CSVパース時にNumber型変換で先頭ゼロが消える対策）
    const patientId = rawPatientId.padStart(6, '0');
    map.set(patientId, {
      patientId,
      name: row[col.name] || '',
      birth: normalizeJapaneseDate(row[col.birth] || ''),
      age: row[col.age] || '',
      sex: row[col.sex] || '',
      zip: formatZipCode(row[col.zip] || ''),
      address: row[col.address] || '',
      diagnosis: sanitizeMultiline(row[col.diagnosis] || ''),
      facility: row[col.facility] || '',
      bedridden: normalizeRomanNumerals(row[col.bedridden] || ''),
      dementia: normalizeRomanNumerals(row[col.dementia] || ''),
      careLevel: normalizeRomanNumerals(row[col.careLevel] || ''),
      caution: sanitizeMultiline(row[col.caution] || ''),
      visitText: {
        1: selectVisitText(row[col.visit1Change], row[col.visit1Stable]),
        2: selectVisitText(row[col.visit2Change], row[col.visit2Stable])
      }
    });
  });

  return map;
}

/**
 * 訪問プランを構築する
 * @param {Array} rows - カレンダーCSVの行
 * @param {Map} patientMap - 患者情報のMap
 * @returns {Object} { visitsByPatient: 2件以下の患者, multipleVisitsPatients: 3件以上の患者, calendarYearMonth: カレンダーの年月 }
 */
function buildVisitPlan(rows, patientMap) {
  const allVisits = {};
  let calendarYearMonth = null;
  if (rows.length === 0) return { visitsByPatient: {}, multipleVisitsPatients: {}, calendarYearMonth: null };

  const header = rows[0];
  const normalizedHeader = header.map(value => normalizeCircledNumbers(value || ''));
  const index = name => normalizedHeader.indexOf(normalizeCircledNumbers(name));

  const col = {
    date: index('日付'),
    start: index('開始時間'),
    end: index('終了時間'),
    doctor: index('医師名'),
    patientId: index('orcaID'),
    status: index('定期訪問')
  };

  // 最初のデータ行（2行目）から年月を取得
  if (rows.length > 1 && col.date >= 0) {
    const firstDate = parseCalendarDate(rows[1][col.date]);
    if (firstDate) {
      calendarYearMonth = Utilities.formatDate(firstDate, 'Asia/Tokyo', 'yyyy-MM');
      logInfo('カレンダーCSVの年月を取得', { calendarYearMonth, firstDate: rows[1][col.date] });
    }
  }

  rows.slice(1).forEach(row => {
    const status = (row[col.status] || '').trim();
    if (status !== '定期訪問') return;

    const rawPatientId = (row[col.patientId] || '').trim();
    if (!rawPatientId) return;
    // orcaIDを6桁ゼロパディングに変換（患者マスタの形式に合わせる）
    const patientId = rawPatientId.padStart(6, '0');
    if (!patientMap.has(patientId)) return;

    const visitDate = parseCalendarDate(row[col.date]);
    if (!visitDate) return;

    allVisits[patientId] = allVisits[patientId] || [];
    allVisits[patientId].push({
      patientId,
      date: visitDate,
      dateText: Utilities.formatDate(visitDate, 'Asia/Tokyo', 'yyyy-MM-dd'),
      startTime: row[col.start] || '',
      endTime: row[col.end] || '',
      doctor: row[col.doctor] || ''
    });
  });

  // 訪問を日付でソートし、2件以下と3件以上を分離
  const visitsByPatient = {};
  const multipleVisitsPatients = {};

  Object.keys(allVisits).forEach(patientId => {
    const visits = allVisits[patientId];
    visits.sort((a, b) => a.date - b.date);

    // 訪問回番号を付与（1から順番）
    visits.forEach((visit, index) => {
      visit.visitNumber = index + 1;  // 何回目の訪問か
    });

    if (visits.length <= 2) {
      // 2件以下: 通常処理対象
      visits.forEach((visit, index) => {
        visit.sequence = index + 1;  // 書類の①②
      });
      visitsByPatient[patientId] = visits;
    } else {
      // 3件以上: 訪問選択シートで選択が必要
      multipleVisitsPatients[patientId] = visits;
      logInfo('定期訪問が3件以上のため訪問選択が必要', {
        patientId,
        count: visits.length,
        dates: visits.map(v => v.dateText)
      });
    }
  });

  return { visitsByPatient, multipleVisitsPatients, calendarYearMonth };
}

/**
 * 選択された訪問回に基づいて訪問プランを構築
 * @param {Array} allVisits - 全訪問データ
 * @param {number} visit1Number - 書類①に使用する訪問回（1〜N）
 * @param {number} visit2Number - 書類②に使用する訪問回（1〜N）
 * @returns {Array} 選択された2件の訪問（sequence付き）
 */
function selectVisitsForDocument(allVisits, visit1Number, visit2Number) {
  const selectedVisits = [];

  // 訪問回番号で検索
  const visit1 = allVisits.find(v => v.visitNumber === visit1Number);
  const visit2 = allVisits.find(v => v.visitNumber === visit2Number);

  if (visit1) {
    selectedVisits.push({ ...visit1, sequence: 1 });
  }
  if (visit2) {
    selectedVisits.push({ ...visit2, sequence: 2 });
  }

  return selectedVisits;
}
