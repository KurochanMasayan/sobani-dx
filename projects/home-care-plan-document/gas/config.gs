/**
 * 共通設定と環境変数
 * 在宅診療計画書用
 */
const SCRIPT_PROPS = PropertiesService.getScriptProperties();

const BASE_CONFIG = {
  encoding: 'cp932',
  sheetTemplateName: '原本',
  registrySheetName: '事業所ファイル台帳',
  statusSheetName: '進行管理',
  queueSheetName: '処理キュー',            // ソート済みキューを保存するシート
  yearMonthPropertyKey: 'PLAN_YEAR_MONTH',  // 処理中の年月を保持（月またぎ対策）
  runIdPropertyKey: 'PLAN_RUN_ID',          // 実行ID（6分制限による中断再開で同じファイルに追記するため）
  maxExecutionMs: 5 * 60 * 1000, // 5分で一旦中断し、再実行で続きから処理
  cellMap: {
    createdDate: 'H2',        // 作成日（出力日）
    patientName: 'B4',        // 患者氏名（CSV A列）
    sex: 'K4',                // 性別（CSV D列）
    birthDate: 'B5',          // 生年月日（CSV B列）
    age: 'J5',                // 年齢（CSV C列）
    zipCode: 'B6',            // 郵便番号（CSV E列）
    address: 'B7',            // 住所（CSV F列）
    diagnosis: 'B8',          // 診断名（CSV G列）
    visitLocation: 'B10',     // 診療場所（CSV H列）※10行目
    carePlan: 'B12',          // 在宅療養計画の要点（CSV L列）
    cautionNote: 'B16',       // 日常生活や介護サービスを利用する上での留意点（CSV M列）
    bedriddenLevel: 'D19',    // 寝たきり度（CSV I列）
    dementiaLevel: 'G19',     // 認知度（CSV J列）
    careLevel: 'J19'          // 介護認定（CSV K列）
  }
};

function getFolderIds() {
  return {
    patientCsv: getRequiredProp('PATIENT_CSV_FOLDER_ID'),
    output: getRequiredProp('OUTPUT_FOLDER_ID')
  };
}

function getRequiredProp(key) {
  const value = SCRIPT_PROPS.getProperty(key);
  if (!value) {
    throw new Error(`Script Property ${key} が設定されていません`);
  }
  return value;
}
