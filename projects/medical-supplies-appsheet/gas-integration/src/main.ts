/**
 * 医療用品管理システム メイン実行ファイル
 * ボタンから実行可能な関数群
 */

/**
 * 全シート一括CSV出力（ボタン用）
 * アクティブなスプレッドシートの全シートをCSV出力
 */
function exportAllSheetsButton(): any {
  try {
    console.log('医療用品データの全シート一括CSV出力を開始...');
    const result = exportAllSheetsToCSV();

    console.log('=== CSV出力結果 ===');
    console.log(`スプレッドシート: ${result.spreadsheetName}`);
    console.log(`処理シート数: ${result.processedSheets}/${result.totalSheets}`);
    console.log(`出力フォルダ: ${result.folderUrl}`);

    if (result.errors.length > 0) {
      console.warn(`エラー: ${result.errors.length}件`);
      result.errors.forEach(err => console.warn(`- ${err.sheetName}: ${err.error}`));
    }

    // 成功時のメッセージ
    if (result.success) {
      console.log(`\n✅ CSV出力が正常に完了しました！`);
      console.log(`📁 出力先: ${result.folderUrl}`);
      console.log(`📊 出力ファイル数: ${result.files.length}件`);

      // ファイル一覧をログ出力
      if (result.files.length > 0) {
        console.log('\n📋 出力されたファイル:');
        result.files.forEach((file, index) => {
          console.log(`${index + 1}. ${file.fileName} (${file.rowCount}行 × ${file.columnCount}列)`);
        });
      }
    } else {
      console.error('❌ CSV出力中にエラーが発生しました。上記のエラー詳細を確認してください。');
    }

    return result;
  } catch (error: any) {
    console.error(`全シートCSV出力エラー: ${error.message}`);
    console.error('スタックトレース:', error.stack);
    throw error;
  }
}

/**
 * 指定シートのみCSV出力（ボタン用）
 * 医療用品管理でよく使用される主要シートのみを出力
 */
function exportMainSheetsButton(): any {
  try {
    // 一般的な医療用品管理シート名
    const mainSheets = [
      '在庫管理',
      '入庫記録',
      '出庫記録',
      '発注管理',
      '商品マスタ',
      '供給業者',
      'ダッシュボード',
      '月次集計',
      '年次集計'
    ];

    console.log('主要シートのCSV出力を開始...');
    const result = exportSelectedSheetsToCSV(mainSheets, '医療用品主要データ');

    console.log('=== 主要シートCSV出力結果 ===');
    console.log(`処理シート数: ${result.processedSheets}/${result.totalSheets}`);
    console.log(`出力フォルダ: ${result.folderUrl}`);

    return result;
  } catch (error: any) {
    console.error(`主要シートCSV出力エラー: ${error.message}`);
    throw error;
  }
}

/**
 * カスタムシート選択CSV出力（ボタン用）
 * ユーザーがシート名を指定して出力
 */
function exportCustomSheetsButton(): any {
  try {
    // ユーザーにシート名の入力を求める
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'CSV出力するシート選択',
      'CSV出力したいシート名をカンマ区切りで入力してください:\n例: 在庫管理,入庫記録,出庫記録',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
      console.log('CSV出力がキャンセルされました');
      return { cancelled: true };
    }

    const input = response.getResponseText().trim();
    if (!input) {
      throw new Error('シート名が入力されていません');
    }

    // シート名を配列に変換
    const sheetNames = input.split(',').map(name => name.trim()).filter(name => name.length > 0);

    if (sheetNames.length === 0) {
      throw new Error('有効なシート名が入力されていません');
    }

    console.log(`指定シート${sheetNames.length}件のCSV出力を開始...`);
    console.log(`対象シート: ${sheetNames.join(', ')}`);

    const result = exportSelectedSheetsToCSV(sheetNames, 'カスタム選択データ');

    console.log('=== カスタム選択CSV出力結果 ===');
    console.log(`処理シート数: ${result.processedSheets}/${result.totalSheets}`);
    console.log(`出力フォルダ: ${result.folderUrl}`);

    // 成功時のメッセージダイアログ
    if (result.success && result.processedSheets > 0) {
      ui.alert(
        'CSV出力完了',
        `${result.processedSheets}件のシートをCSV出力しました。\n\n出力先: Google Driveの「${result.folderUrl}」\n\nログで詳細を確認してください。`,
        ui.ButtonSet.OK
      );
    }

    return result;
  } catch (error: any) {
    console.error(`カスタムシートCSV出力エラー: ${error.message}`);
    const ui = SpreadsheetApp.getUi();
    ui.alert('エラー', `CSV出力中にエラーが発生しました:\n${error.message}`, ui.ButtonSet.OK);
    throw error;
  }
}

/**
 * スプレッドシート情報を取得（デバッグ用）
 */
function getSpreadsheetInfo(): any {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();

    const info = {
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      spreadsheetUrl: ss.getUrl(),
      totalSheets: sheets.length,
      sheets: sheets.map(sheet => ({
        name: sheet.getName(),
        index: sheet.getIndex(),
        rowCount: sheet.getLastRow(),
        columnCount: sheet.getLastColumn(),
        dataRange: sheet.getDataRange().getA1Notation()
      }))
    };

    console.log('=== スプレッドシート情報 ===');
    console.log(`名前: ${info.spreadsheetName}`);
    console.log(`ID: ${info.spreadsheetId}`);
    console.log(`総シート数: ${info.totalSheets}`);
    console.log('\nシート一覧:');
    info.sheets.forEach((sheet, index) => {
      console.log(`${index + 1}. ${sheet.name} (${sheet.rowCount}行 × ${sheet.columnCount}列, 範囲: ${sheet.dataRange})`);
    });

    return info;
  } catch (error: any) {
    console.error(`スプレッドシート情報取得エラー: ${error.message}`);
    throw error;
  }
}

/**
 * CSV出力フォルダを整理（古いフォルダを削除）
 */
function cleanupOldExportFolders(): any {
  try {
    const folders = DriveApp.getFolders();
    const exportFolders: GoogleAppsScript.Drive.Folder[] = [];
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - 30); // 30日前

    // CSV出力フォルダを検索
    while (folders.hasNext()) {
      const folder = folders.next();
      const folderName = folder.getName();

      if (folderName.includes('医療用品データ出力') || folderName.includes('医療用品選択データ出力') || folderName.includes('カスタム選択データ')) {
        if (folder.getDateCreated() < cutoffDate) {
          exportFolders.push(folder);
        }
      }
    }

    console.log(`${exportFolders.length}個の古いCSV出力フォルダを削除します...`);

    let deletedCount = 0;
    for (const folder of exportFolders) {
      try {
        console.log(`削除中: ${folder.getName()}`);
        folder.setTrashed(true);
        deletedCount++;
      } catch (deleteError: any) {
        console.warn(`フォルダ削除失敗: ${folder.getName()} - ${deleteError.message}`);
      }
    }

    console.log(`クリーンアップ完了: ${deletedCount}個のフォルダを削除しました`);

    return {
      totalFound: exportFolders.length,
      deleted: deletedCount,
      cutoffDate: cutoffDate
    };
  } catch (error: any) {
    console.error(`フォルダクリーンアップエラー: ${error.message}`);
    throw error;
  }
}