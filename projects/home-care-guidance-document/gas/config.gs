/**
 * 共通設定と環境変数
 */
const SCRIPT_PROPS = PropertiesService.getScriptProperties();

const BASE_CONFIG = {
  encoding: 'cp932',
  sheetTemplateName: '原本',
  registrySheetName: '事業所ファイル台帳',
  statusSheetName: '進行管理',
  queueSheetName: '処理キュー',                       // ソート済みキューを保存するシート
  visitSelectionSheetName: '訪問選択',               // 3件以上訪問がある患者の訪問回選択シート
  yearMonthPropertyKey: 'GUIDANCE_YEAR_MONTH',        // 処理中の年月を保持（月またぎ対策）
  maxExecutionMs: 5 * 60 * 1000, // 5分で一旦中断し、再実行で続きから処理
  fallbackFacilityName: '事業所未設定',              // 施設名が空の場合のフォールバック
  // キュー・訪問選択シートのステータス文字列
  status: {
    pending: '待機',
    processed: '処理済み',
    unprocessed: '未処理'
  },
  // 訪問回選択のデフォルト設定
  defaultVisitPattern: {
    visits3: [1, 2],    // 全3回の場合: 1回目と2回目を使用
    visits4: [1, 3]     // 全4回の場合: 1回目と3回目を使用（デフォルト）
  },
  cellMap: {
    createdDate: 'R2',
    facilityName: 'H4',
    patientName: 'G7',
    sex: 'T7',
    birthDate: 'G8',
    age: 'R8',
    zipCode: 'G9',
    address: 'G10',
    visitDate: 'F11',
    diagnosis: 'F13',
    visitNote: 'C22',
    careLevel: 'G27',
    bedriddenLevel: 'G28',
    dementiaLevel: 'G29',
    cautionNote: 'B31'
  }
};

function getFolderIds() {
  return {
    patientCsv: getRequiredProp('PATIENT_CSV_FOLDER_ID'),
    calendarCsv: getRequiredProp('CALENDAR_CSV_FOLDER_ID'),
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
