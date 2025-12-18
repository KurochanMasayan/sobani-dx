/**
 * 設定ファイル
 */

const CONFIG = {
  // スプレッドシートの設定
  SHEETS: {
    FACILITY_CALENDAR: '施設カレンダー'
  },
  
  // PDFエクスポートの設定
  PDF: {
    // PDF保存先のGoogle DriveフォルダID
    // フォルダIDは、Google Driveでフォルダを開いた時のURLから取得できます
    // 例: https://drive.google.com/drive/folders/XXXXX ← このXXXXXの部分
    FOLDER_ID: '1FvHFGRSGEj29i60Wj315DZv0W2BmhPV0',
    
    // PDFエクスポート範囲
    EXPORT_RANGE: 'C1:R37',
    
    // PDFの設定
    ORIENTATION: 'landscape' as const, // 横向き
    SIZE: 'A4' as const,
    SCALE: 'fit' as const, // ページに合わせる
    MARGINS: {
      top: 0.1,
      bottom: 0.1,
      left: 0.1,
      right: 0.1
    }
  },
  
  // 施設データの設定
  FACILITY: {
    // 施設名を取得するセル（施設カレンダーシート）
    NAME_CELL: 'A5',
    
    // 合算訪問一覧から施設データを取得する際の設定
    VISIT_LIST: {
      // シート名
      SHEET_NAME: '合算訪問一覧',
      // データ開始行（A2から）
      START_ROW: 2,
      // データ列（A列）
      DATA_COLUMN: 1,
      // スキップする値のパターン
      SKIP_PATTERNS: ['【関数】', '施設名']
    }
  },
  
  // Googleカレンダーの設定
  CALENDAR: {
    // カレンダーID（'primary'はデフォルトカレンダー）
    // 複数のカレンダーを使用する場合は配列で指定
    CALENDAR_IDS: ['c_c89f5fd72defb5ef85547e43d4c019687042f5882de1ec7fdfd2edbb6e2064b4@group.calendar.google.com'],

    // カレンダーデータを出力するシート名
    OUTPUT_SHEET: 'カレンダーデータ',

    // 日付を取得するシートと範囲
    DATE_SOURCE: {
      SHEET_NAME: '使い方',
      START_DATE_CELL: 'B2',  // 開始日のセル
      END_DATE_CELL: 'B3'      // 終了日のセル
    }
  },

  // 歯科カレンダーの設定
  DENTAL_CALENDAR: {
    // 歯科用カレンダーID
    CALENDAR_ID: 'c_4b2c180413cb0c727836433b8f62d492af5d7f282a61d411c0cf3e713bb3f937@group.calendar.google.com',

    // 歯科カレンダーデータを出力するシート名
    OUTPUT_SHEET: '歯科カレンダーデータ',

    // 日付を取得するシートと範囲（通常のカレンダーと同じ設定を使用）
    DATE_SOURCE: {
      SHEET_NAME: '使い方',
      START_DATE_CELL: 'B2',  // 開始日のセル
      END_DATE_CELL: 'B3'      // 終了日のセル
    }
  }
};

/**
 * PDFフォルダIDの検証と設定
 * @return {string} 有効なフォルダID
 */
function getPdfFolderId(): string {
  const folderId = CONFIG.PDF.FOLDER_ID;
  
  if (folderId === 'YOUR_FOLDER_ID_HERE' || !folderId) {
    throw new Error('PDFの保存先フォルダIDが設定されていません。config.gsのCONFIG.PDF.FOLDER_IDを設定してください。');
  }
  
  // フォルダの存在確認
  try {
    DriveApp.getFolderById(folderId);
    return folderId;
  } catch (e) {
    throw new Error(`指定されたフォルダID（${folderId}）が見つかりません。正しいフォルダIDを設定してください。`);
  }
}
