/**
 * メインファイル
 * メニューの作成と基本機能の管理
 */

/**
 * スプレッドシートを開いた時に実行される関数
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // カスタムメニューを作成
  ui.createMenu('施設管理')
    .addSubMenu(ui.createMenu('PDF作成')
      .addItem('単一施設のPDF作成', 'createPdfForFacility')
      .addItem('全施設のPDF一括作成', 'createAllPdfsWithConfirmation')
      .addSeparator()
      .addItem('PDF保存先の設定確認', 'checkPdfFolderSettings'))
    .addSeparator()
    .addSubMenu(ui.createMenu('CSVエクスポート')
      .addItem('施設データをCSV出力', 'exportFacilityDataToCsv')
      .addItem('施設カレンダーをCSV出力', 'exportCalendarToCsv'))
    .addSeparator()
    .addItem('設定', 'showSettings')
    .addItem('ヘルプ', 'showHelp')
    .addToUi();
}

/**
 * PDF保存先の設定を確認
 */
function checkPdfFolderSettings() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const folderId = getPdfFolderId();
    const folder = DriveApp.getFolderById(folderId);
    
    ui.alert(
      'PDF保存先設定',
      `現在のPDF保存先:\n\nフォルダ名: ${folder.getName()}\nフォルダID: ${folderId}\n\n設定を変更する場合は、config.gsのCONFIG.PDF.FOLDER_IDを編集してください。`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert(
      'PDF保存先未設定',
      `PDF保存先が設定されていません。\n\n${error.message}\n\n1. Google Driveで保存先フォルダを作成\n2. フォルダのURLからIDをコピー\n3. config.gsのCONFIG.PDF.FOLDER_IDに設定`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 設定画面を表示
 */
function showSettings() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="padding: 20px; font-family: Arial, sans-serif;">
      <h2>設定</h2>
      
      <h3>現在の設定値</h3>
      <table style="width: 100%; border-collapse: collapse;">
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd;"><b>施設データシート</b></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${CONFIG.SHEETS.FACILITY_DATA}</td>
        </tr>
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd;"><b>施設カレンダーシート</b></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${CONFIG.SHEETS.FACILITY_CALENDAR}</td>
        </tr>
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd;"><b>PDF出力範囲</b></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${CONFIG.PDF.EXPORT_RANGE}</td>
        </tr>
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd;"><b>PDF用紙サイズ</b></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${CONFIG.PDF.SIZE} (${CONFIG.PDF.ORIENTATION})</td>
        </tr>
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd;"><b>施設名セル</b></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${CONFIG.FACILITY.NAME_CELL}</td>
        </tr>
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd;"><b>除外パターン</b></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${CONFIG.FACILITY.EXCLUDE_PATTERN}</td>
        </tr>
      </table>
      
      <p style="margin-top: 20px; padding: 10px; background-color: #f0f0f0; border-radius: 4px;">
        <b>設定の変更方法:</b><br>
        スクリプトエディタでconfig.gsファイルを編集してください。
      </p>
    </div>
  `)
  .setWidth(500)
  .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '設定');
}

/**
 * ヘルプを表示
 */
function showHelp() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="padding: 20px; font-family: Arial, sans-serif; line-height: 1.6;">
      <h2>施設管理システム ヘルプ</h2>
      
      <h3>📁 PDF作成機能</h3>
      <ul>
        <li><b>単一施設のPDF作成:</b> 指定した施設のカレンダーをPDF化</li>
        <li><b>全施設のPDF一括作成:</b> 全施設分のPDFを自動生成</li>
        <li>PDFファイル名: 施設名_年月.pdf</li>
        <li>既存ファイルは自動的に上書きされます</li>
      </ul>
      
      <h3>📊 CSV出力機能</h3>
      <ul>
        <li><b>施設データ:</b> 施設データシート全体をCSV出力</li>
        <li><b>施設カレンダー:</b> カレンダーの指定範囲をCSV出力</li>
        <li>文字コード: UTF-8（BOM付き）</li>
      </ul>
      
      <h3>⚙️ 初期設定</h3>
      <ol>
        <li>Google DriveでPDF保存用フォルダを作成</li>
        <li>フォルダを開き、URLからIDをコピー<br>
            例: drive.google.com/drive/folders/<b>XXXXX</b></li>
        <li>config.gsのCONFIG.PDF.FOLDER_IDに設定</li>
      </ol>
      
      <h3>📝 注意事項</h3>
      <ul>
        <li>「個人宅」を含む施設は自動的に除外されます</li>
        <li>大量のPDF作成時は処理に時間がかかります</li>
        <li>エラーが発生した場合はログを確認してください</li>
      </ul>
      
      <div style="margin-top: 20px; padding: 10px; background-color: #e3f2fd; border-radius: 4px;">
        <b>お問い合わせ:</b> システム管理者までご連絡ください
      </div>
    </div>
  `)
  .setWidth(600)
  .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ヘルプ');
}

/**
 * テスト用：設定値の確認
 */
function testConfiguration() {
  console.log('=== 設定確認 ===');
  console.log('施設データシート:', CONFIG.SHEETS.FACILITY_DATA);
  console.log('施設カレンダーシート:', CONFIG.SHEETS.FACILITY_CALENDAR);
  console.log('PDF出力範囲:', CONFIG.PDF.EXPORT_RANGE);
  console.log('PDFサイズ:', CONFIG.PDF.SIZE);
  console.log('PDF向き:', CONFIG.PDF.ORIENTATION);
  
  try {
    const folderId = getPdfFolderId();
    console.log('PDF保存先フォルダID:', folderId);
    const folder = DriveApp.getFolderById(folderId);
    console.log('PDF保存先フォルダ名:', folder.getName());
  } catch (error) {
    console.error('PDF保存先エラー:', error.message);
  }
}

/**
 * テスト用：単一施設のPDF作成テスト
 */
function testSinglePdfCreation() {
  const testFacilityName = 'テスト施設';
  try {
    const result = createSinglePdf(testFacilityName);
    console.log('テストPDF作成成功:', result.getName());
  } catch (error) {
    console.error('テストPDF作成失敗:', error.message);
  }
}