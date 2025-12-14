/**
 * CSVデータの読み込みと加工
 * 在宅診療計画書用（カレンダーCSV不要、患者CSVのみで処理）
 */
function buildProcessingContext(folders) {
  const patientRows = loadCsvRows(folders.patientCsv);
  const patientMap = buildPatientMap(patientRows);
  // 患者CSVのみで処理（カレンダーCSVは使用しない）
  const visitsByPatient = buildVisitPlanFromPatients(patientMap);
  return { patientMap, visitsByPatient };
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

  // CSV列マッピング（在宅診療計画書用）
  // A:患者名, B:生年月日, C:年齢, D:性別, E:郵便番号, F:住所, G:診断名,
  // H:施設名, I:寝たきり度, J:認知度, K:介護認定, L:在宅療養計画の要点, M:留意点
  const col = {
    name: 0,       // A列: 患者名
    birth: 1,      // B列: 生年月日
    age: 2,        // C列: 年齢
    sex: 3,        // D列: 性別
    zip: 4,        // E列: 郵便番号
    address: 5,    // F列: 住所
    diagnosis: 6,  // G列: 診断名
    facility: 7,   // H列: 施設名（事業所として使用）
    bedridden: 8,  // I列: 寝たきり度
    dementia: 9,   // J列: 認知度
    careLevel: 10, // K列: 介護認定
    carePlan: 11,  // L列: 在宅療養計画の要点
    caution: 12    // M列: 日常生活や介護サービスを利用する上での留意点
  };

  rows.slice(1).forEach(row => {
    const name = (row[col.name] || '').trim();
    if (!name) return;
    // 患者IDがないため、患者名をキーとして使用
    const patientId = name;
    map.set(patientId, {
      patientId,
      name: name,
      birth: normalizeJapaneseDate(row[col.birth] || ''),
      age: (row[col.age] || '').replace(/歳$/, ''),
      sex: row[col.sex] || '',
      zip: formatZipCode(row[col.zip] || ''),
      address: row[col.address] || '',
      diagnosis: sanitizeMultiline(row[col.diagnosis] || ''),
      facility: row[col.facility] || '',
      bedridden: normalizeRomanNumerals(row[col.bedridden] || ''),
      dementia: normalizeRomanNumerals(row[col.dementia] || ''),
      careLevel: normalizeRomanNumerals(row[col.careLevel] || ''),
      carePlan: sanitizeMultiline(row[col.carePlan] || ''),
      caution: sanitizeMultiline(row[col.caution] || '')
    });
  });

  return map;
}

/**
 * 患者マップから訪問プランを生成（カレンダーCSV不要）
 * 各患者に対して現在日付で1件の訪問データを作成
 */
function buildVisitPlanFromPatients(patientMap) {
  const plan = {};
  const now = new Date();
  const dateText = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');

  patientMap.forEach((patient, patientId) => {
    plan[patientId] = [{
      date: now,
      dateText: dateText
    }];
  });

  return plan;
}
