const DRIVE_SYNC_CONFIG = {
  FOLDER_ID: '1LGjfnZdSyWDuz1gjvDHa00HSFtnxNi2Q',
  SPREADSHEET_ID: '1_1BIGayuq-pWTLYLhTOoMUcNk-xy_w91NlepqvM7Ydg'
} as const;

type FileEntry = {
  file: GoogleAppsScript.Drive.File;
  pathSegments: string[];
};

/**
 * Google ドライブのフォルダにある CSV/TSV を再帰的に探し、対象スプレッドシートへ反映する
 */
function updateAllCsvFiles(): void {
  try {
    const rootFolder = DriveApp.getFolderById(DRIVE_SYNC_CONFIG.FOLDER_ID);
    const spreadsheet = SpreadsheetApp.openById(DRIVE_SYNC_CONFIG.SPREADSHEET_ID);

    const fileEntries: FileEntry[] = [];
    collectCsvAndTsvFiles(rootFolder, [], fileEntries);

    if (fileEntries.length === 0) {
      notifyUi('更新対象なし', 'CSV/TSVファイルが見つかりませんでした。');
      return;
    }

    const usedSheetNames = new Set<string>();
    let updatedCount = 0;
    const results: string[] = [];

    fileEntries.forEach(({ file, pathSegments }) => {
      const fileName = file.getName();
      const baseSheetName = buildSheetName(fileName);
      const sheetName = makeUniqueSheetName(baseSheetName, usedSheetNames);
      const isTsv = fileName.toLowerCase().endsWith('.tsv');

      console.log(`処理開始: ${sheetName} ← ${fileName} (${isTsv ? 'TSV' : 'CSV'})`);

      try {
        updateSheetFromCsv(spreadsheet, file, sheetName, isTsv);
        updatedCount++;
        results.push(`✓ ${sheetName}: ${fileName} を反映しました`);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        console.error(`✗ ${sheetName}: ${message}`);
        results.push(`✗ ${sheetName}: エラー - ${message}`);
      }
    });

    notifyUi(
      '更新完了',
      `${updatedCount}件のCSV/TSVファイルを更新しました。\n\n${results.join('\n')}`
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error('フォルダまたはスプレッドシートへのアクセスに失敗しました:', message);
    notifyUi('エラー', `フォルダ/スプレッドシートへのアクセスに失敗しました: ${message}`);
    throw error;
  }
}

/**
 * エディタやトリガーから UI なしで実行するためのヘルパー
 */
function updateAllCsvFilesFromEditor(): {
  success: boolean;
  updatedCount: number;
  results: string[];
} {
  const rootFolder = DriveApp.getFolderById(DRIVE_SYNC_CONFIG.FOLDER_ID);
  const spreadsheet = SpreadsheetApp.openById(DRIVE_SYNC_CONFIG.SPREADSHEET_ID);

  const fileEntries: FileEntry[] = [];
  collectCsvAndTsvFiles(rootFolder, [], fileEntries);

  if (fileEntries.length === 0) {
    console.log('CSV/TSVファイルが見つかりませんでした。');
    return { success: true, updatedCount: 0, results: [] };
  }

  const usedSheetNames = new Set<string>();
  let updatedCount = 0;
  const results: string[] = [];

  fileEntries.forEach(({ file, pathSegments }) => {
    const fileName = file.getName();
    const baseSheetName = buildSheetName(fileName);
    const sheetName = makeUniqueSheetName(baseSheetName, usedSheetNames);
    const isTsv = fileName.toLowerCase().endsWith('.tsv');

    console.log(`処理開始: ${sheetName} ← ${fileName} (${isTsv ? 'TSV' : 'CSV'})`);

    try {
      updateSheetFromCsv(spreadsheet, file, sheetName, isTsv);
      updatedCount++;
      results.push(`✓ ${sheetName}: ${fileName} を反映しました`);
      console.log(`✓ ${sheetName}: 更新成功`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      results.push(`✗ ${sheetName}: エラー - ${message}`);
      console.error(`✗ ${sheetName}: エラー - ${message}`);
    }
  });

  console.log(`更新完了: ${updatedCount}件のCSV/TSVファイルを更新しました。`);
  return { success: true, updatedCount, results };
}

function collectCsvAndTsvFiles(
  folder: GoogleAppsScript.Drive.Folder,
  pathSegments: string[],
  collector: FileEntry[]
): void {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const lowerName = file.getName().toLowerCase();
    const mimeType = file.getMimeType();

    const isCsv = mimeType === MimeType.CSV || lowerName.endsWith('.csv');
    const isTsv = mimeType === 'text/tab-separated-values' || lowerName.endsWith('.tsv');

    if (isCsv || isTsv) {
      collector.push({ file, pathSegments });
    }
  }

  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    const subFolder = subFolders.next();
    collectCsvAndTsvFiles(subFolder, pathSegments.concat([subFolder.getName()]), collector);
  }
}

function buildSheetName(fileName: string): string {
  const baseName = sanitizeSheetNamePart(removeExtension(fileName)) || 'Sheet';
  return baseName.slice(0, 100) || 'Sheet';
}

function sanitizeSheetNamePart(name: string): string {
  return name
    .replace(/[\\\/\*\[\]\?:]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function makeUniqueSheetName(baseName: string, usedNames: Set<string>): string {
  let candidate = baseName;
  let counter = 2;
  while (usedNames.has(candidate)) {
    candidate = `${baseName}_${counter}`.slice(0, 100);
    counter++;
  }
  usedNames.add(candidate);
  return candidate;
}

function removeExtension(fileName: string): string {
  return fileName.replace(/\.(csv|tsv)$/i, '');
}

function updateSheetFromCsv(
  spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet,
  csvFile: GoogleAppsScript.Drive.File,
  sheetName: string,
  isTsv: boolean
): void {
  const csvContent = csvFile.getBlob().getDataAsString('UTF-8');

  const preview = csvContent.split('\n').slice(0, 5).join('\n');
  console.log(`Preview (${sheetName}):\n${preview}`);

  const parsed = parseDelimited(csvContent, isTsv ? '\t' : ',');
  if (parsed.length === 0) {
    throw new Error('CSV/TSVファイルが空です');
  }

  let sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  ensureSheetCapacity(sheet, parsed.length, parsed[0]?.length ?? 0);
  sheet.getRange(1, 1, parsed.length, parsed[0].length).setValues(parsed);
}

function parseDelimited(content: string, delimiter: string): string[][] {
  try {
    return Utilities.parseCsv(content, delimiter) as string[][];
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.warn(`Utilities.parseCsv でエラー: ${message} -> フォールバックを使用します`);
    return manualParse(content, delimiter);
  }
}

function manualParse(content: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  let currentRow: string[] = [];
  let currentField = '';
  let inQuotes = false;
  let fieldStart = true;

  for (let i = 0; i < content.length; i++) {
    const char = content[i];
    const nextChar = content[i + 1];

    if (fieldStart && char === '"') {
      inQuotes = true;
      fieldStart = false;
      continue;
    }

    if (inQuotes && char === '"') {
      if (nextChar === '"') {
        currentField += '"';
        i++;
      } else if (
        nextChar === delimiter ||
        nextChar === '\r' ||
        nextChar === '\n' ||
        nextChar === undefined
      ) {
        inQuotes = false;
      } else {
        currentField += char;
      }
      continue;
    }

    if (!inQuotes && char === delimiter) {
      currentRow.push(currentField);
      currentField = '';
      fieldStart = true;
      continue;
    }

    if (!inQuotes && (char === '\r' || char === '\n')) {
      if (char === '\r' && nextChar === '\n') {
        i++;
      }
      currentRow.push(currentField);
      if (currentRow.length > 0) {
        rows.push(currentRow);
      }
      currentRow = [];
      currentField = '';
      fieldStart = true;
      continue;
    }

    currentField += char;
    fieldStart = false;
  }

  if (currentField !== '' || currentRow.length > 0) {
    currentRow.push(currentField);
    if (currentRow.length > 0) {
      rows.push(currentRow);
    }
  }

  if (rows.length === 0) {
    throw new Error('CSV/TSVファイルが空です');
  }

  const maxColumns = rows.reduce((max, row) => Math.max(max, row.length), 0);
  return rows.map(row => {
    while (row.length < maxColumns) {
      row.push('');
    }
    return row;
  });
}

function ensureSheetCapacity(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  rowsNeeded: number,
  colsNeeded: number
): void {
  if (rowsNeeded > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), rowsNeeded - sheet.getMaxRows());
  }
  if (colsNeeded > sheet.getMaxColumns()) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), colsNeeded - sheet.getMaxColumns());
  }
}

function notifyUi(title: string, message: string): void {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
  } catch (error) {
    console.log(`${title}: ${message.replace(/\n/g, ' ')}`);
  }
}
