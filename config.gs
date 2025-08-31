/**
 * 設定ファイル
 */

const CONFIG = {
  // スプレッドシートの設定
  SHEETS: {
    FACILITY_DATA: '施設データ',
    FACILITY_CALENDAR: '施設カレンダー'
  },
  
  // PDFエクスポートの設定
  PDF: {
    // PDF保存先のGoogle DriveフォルダID
    // フォルダIDは、Google Driveでフォルダを開いた時のURLから取得できます
    // 例: https://drive.google.com/drive/folders/XXXXX ← このXXXXXの部分
    FOLDER_ID: 'YOUR_FOLDER_ID_HERE',
    
    // PDFエクスポート範囲
    EXPORT_RANGE: 'C1:P25',
    
    // PDFの設定
    ORIENTATION: 'landscape', // 横向き
    SIZE: 'A4',
    SCALE: 'fit', // ページに合わせる
    MARGINS: {
      top: 0.5,
      bottom: 0.5,
      left: 0.5,
      right: 0.5
    }
  },
  
  // 施設データの設定
  FACILITY: {
    // 施設名を取得するセル（施設カレンダーシート）
    NAME_CELL: 'A5',
    
    // 施設データシートの施設名列の開始位置
    DATA_START_ROW: 2,
    DATA_COLUMN: 'D',
    
    // 除外する施設名のパターン
    EXCLUDE_PATTERN: '個人宅'
  }
};

/**
 * PDFフォルダIDの検証と設定
 * @return {string} 有効なフォルダID
 */
function getPdfFolderId() {
  const folderId = CONFIG.PDF.FOLDER_ID;
  
  if (folderId === 'YOUR_FOLDER_ID_HERE' || !folderId) {
    throw new Error('PDFの保存先フォルダIDが設定されていません。config.gsのCONFIG.PDF.FOLDER_IDを設定してください。');
  }
  
  // フォルダの存在確認
  try {
    const folder = DriveApp.getFolderById(folderId);
    return folderId;
  } catch (e) {
    throw new Error(`指定されたフォルダID（${folderId}）が見つかりません。正しいフォルダIDを設定してください。`);
  }
}