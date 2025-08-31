"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const config_1 = require("./config");
function exportCalendarToSpreadsheet() {
    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const usageSheet = spreadsheet.getSheetByName('使い方');
        if (!usageSheet) {
            throw new Error('「使い方」シートが見つかりません');
        }
        const startDateValue = usageSheet.getRange('B2').getValue();
        const endDateValue = usageSheet.getRange('B3').getValue();
        if (!startDateValue || !endDateValue) {
            throw new Error('使い方シートのB2（開始日）またはB3（終了日）が設定されていません');
        }
        const startDate = new Date(startDateValue);
        startDate.setHours(0, 0, 0, 0);
        const endDate = new Date(endDateValue);
        endDate.setHours(23, 59, 59, 999);
        const calendar = CalendarApp.getCalendarById(config_1.CALENDAR_ID);
        let sheet = spreadsheet.getSheetByName("カレンダーデータ");
        if (!sheet) {
            sheet = spreadsheet.insertSheet("カレンダーデータ");
        }
        const lastRow = sheet.getMaxRows();
        if (lastRow > 0) {
            sheet.getRange(2, 1, lastRow, 5).clear();
        }
        sheet.getRange(1, 1, 1, 5).setValues([['日付', '開始時刻', '終了時刻', 'タイトル', 'コメント']]);
        const events = calendar.getEvents(startDate, endDate);
        if (events.length === 0) {
            sheet.getRange(2, 1).setValue('指定期間にイベントはありません');
            console.log('指定期間にイベントはありません');
            return;
        }
        const data = [];
        events.forEach(event => {
            const startTime = event.getStartTime();
            const endTime = event.getEndTime();
            data.push([
                Utilities.formatDate(startTime, 'JST', 'yyyy/M/d'),
                event.isAllDayEvent() ? '終日' : Utilities.formatDate(startTime, 'JST', 'H:mm:ss'),
                event.isAllDayEvent() ? '終日' : Utilities.formatDate(endTime, 'JST', 'H:mm:ss'),
                event.getTitle(),
                event.getDescription() || ''
            ]);
        });
        if (data.length > 0) {
            sheet.getRange(2, 1, data.length, 5).setValues(data);
        }
        console.log(`${events.length}件のイベントをエクスポートしました`);
        console.log(`スプレッドシート: ${spreadsheet.getName()}`);
        console.log(`シート: ${"カレンダーデータ"}`);
    }
    catch (error) {
        console.error('エラーが発生しました:', error);
        throw error;
    }
}
function exportCsvSheetToFile() {
    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const csvSheet = spreadsheet.getSheetByName('csv');
        if (!csvSheet) {
            throw new Error('「csv」シートが見つかりません');
        }
        const lastRow = csvSheet.getLastRow();
        if (lastRow === 0) {
            console.log('csvシートにデータがありません');
            return null;
        }
        const range = csvSheet.getRange(1, 1, lastRow, 14);
        const values = range.getValues();
        let csvContent = '';
        values.forEach((row, index) => {
            const csvRow = row.map((cell, colIndex) => {
                let cellValue = cell === null || cell === undefined ? '' : String(cell);
                if (colIndex <= 2 && cell instanceof Date) {
                    if (colIndex === 0) {
                        cellValue = Utilities.formatDate(cell, 'JST', 'yyyy/M/d');
                    }
                    else {
                        cellValue = Utilities.formatDate(cell, 'JST', 'H:mm:ss');
                    }
                }
                if (index > 0) {
                    cellValue = cellValue.replace(/"/g, '""');
                    cellValue = `"${cellValue}"`;
                }
                else {
                    if (cellValue.includes(',') || cellValue.includes('\n') || cellValue.includes('\r') || cellValue.includes('"')) {
                        cellValue = cellValue.replace(/"/g, '""');
                        cellValue = `"${cellValue}"`;
                    }
                }
                return cellValue;
            }).join(',');
            csvContent += csvRow;
            if (index < values.length - 1) {
                csvContent += '\n';
            }
        });
        const fileName = `export_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}.csv`;
        const blob = Utilities.newBlob(csvContent, 'text/csv', fileName);
        const file = DriveApp.createFile(blob);
        console.log('===========================================');
        console.log(`CSVファイルを作成しました: ${fileName}`);
        console.log(`ファイルURL: ${file.getUrl()}`);
        console.log(`ダウンロードURL: ${file.getDownloadUrl()}`);
        console.log('===========================================');
        console.log('ダウンロードするには以下のURLをブラウザで開いてください:');
        console.log(file.getDownloadUrl());
        console.log('===========================================');
        console.log('※ 文字コード: UTF-8');
        console.log('※ Shift-JISが必要な場合は、メモ帳で開いて「ANSI」形式で保存してください');
        const resultSheet = spreadsheet.getSheetByName('使い方');
        if (resultSheet) {
            resultSheet.getRange('B12').setValue(fileName);
            const richText = SpreadsheetApp.newRichTextValue()
                .setText('📥 クリックしてダウンロード')
                .setLinkUrl(file.getDownloadUrl())
                .build();
            resultSheet.getRange('B13').setRichTextValue(richText);
            resultSheet.getRange('B13').setFontColor('#1a73e8').setFontWeight('bold');
        }
        return {
            fileName: file.getName(),
            fileUrl: file.getUrl(),
            downloadUrl: file.getDownloadUrl()
        };
    }
    catch (error) {
        console.error('CSV出力エラー:', error);
        throw error;
    }
}
function exportFacilityCalendarToPDF() {
    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const spreadsheetId = spreadsheet.getId();
        const targetSheet = spreadsheet.getSheetByName('施設カレンダー');
        if (!targetSheet) {
            throw new Error('「施設カレンダー」シートが見つかりません');
        }
        const sheetId = targetSheet.getSheetId();
        const params = {
            'format': 'pdf',
            'size': 'A4',
            'portrait': false,
            'fitw': true,
            'sheetnames': false,
            'printtitle': false,
            'pagenumbers': false,
            'gridlines': true,
            'fzr': false,
            'gid': sheetId,
            'r1': 0,
            'c1': 2,
            'r2': 24,
            'c2': 15
        };
        const queryString = Object.entries(params)
            .map(([key, value]) => `${key}=${value}`)
            .join('&');
        const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?${queryString}`;
        const response = UrlFetchApp.fetch(url, {
            headers: {
                'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
            }
        });
        const fileName = `施設カレンダー_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}.pdf`;
        const blob = response.getBlob().setName(fileName);
        const file = DriveApp.createFile(blob);
        console.log('===========================================');
        console.log(`PDFファイルを作成しました: ${fileName}`);
        console.log(`範囲: C1:P25（横向き）`);
        console.log(`ファイルURL: ${file.getUrl()}`);
        console.log(`ダウンロードURL: ${file.getDownloadUrl()}`);
        console.log('===========================================');
        const usageSheet = spreadsheet.getSheetByName('使い方');
        if (usageSheet) {
            usageSheet.getRange('B15').setValue(fileName);
            const richText = SpreadsheetApp.newRichTextValue()
                .setText('📄 PDFをダウンロード')
                .setLinkUrl(file.getDownloadUrl())
                .build();
            usageSheet.getRange('B16').setRichTextValue(richText);
            usageSheet.getRange('B16').setFontColor('#d93025').setFontWeight('bold');
        }
        return {
            fileName: file.getName(),
            fileUrl: file.getUrl(),
            downloadUrl: file.getDownloadUrl()
        };
    }
    catch (error) {
        console.error('PDF出力エラー:', error);
        throw error;
    }
}
function exportSpreadsheetToPDF() {
    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const spreadsheetId = spreadsheet.getId();
        const usageSheet = spreadsheet.getSheetByName('使い方');
        let sheetName = 'カレンダーデータ';
        if (usageSheet) {
            const customSheetName = usageSheet.getRange('B5').getValue();
            if (customSheetName) {
                sheetName = customSheetName;
            }
        }
        const targetSheet = spreadsheet.getSheetByName(sheetName);
        if (!targetSheet) {
            throw new Error(`「${sheetName}」シートが見つかりません`);
        }
        const sheetId = targetSheet.getSheetId();
        const params = {
            'format': 'pdf',
            'size': 'A4',
            'portrait': true,
            'fitw': true,
            'sheetnames': false,
            'printtitle': false,
            'pagenumbers': false,
            'gridlines': true,
            'fzr': false,
            'gid': sheetId
        };
        const queryString = Object.entries(params)
            .map(([key, value]) => `${key}=${value}`)
            .join('&');
        const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?${queryString}`;
        const response = UrlFetchApp.fetch(url, {
            headers: {
                'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
            }
        });
        const fileName = `${sheetName}_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}.pdf`;
        const blob = response.getBlob().setName(fileName);
        const file = DriveApp.createFile(blob);
        console.log('===========================================');
        console.log(`PDFファイルを作成しました: ${fileName}`);
        console.log(`ファイルURL: ${file.getUrl()}`);
        console.log(`ダウンロードURL: ${file.getDownloadUrl()}`);
        console.log('===========================================');
        if (usageSheet) {
            usageSheet.getRange('B17').setValue(fileName);
            const richText = SpreadsheetApp.newRichTextValue()
                .setText('📄 汎用PDFをダウンロード')
                .setLinkUrl(file.getDownloadUrl())
                .build();
            usageSheet.getRange('B18').setRichTextValue(richText);
            usageSheet.getRange('B18').setFontColor('#188038').setFontWeight('bold');
        }
        return {
            fileName: file.getName(),
            fileUrl: file.getUrl(),
            downloadUrl: file.getDownloadUrl()
        };
    }
    catch (error) {
        console.error('PDF出力エラー:', error);
        throw error;
    }
}
function executeAll() {
    try {
        console.log('===== 処理開始 =====');
        console.log('1. カレンダーデータを取得中...');
        exportCalendarToSpreadsheet();
        console.log('2. CSVファイルを生成中...');
        const csvResult = exportCsvSheetToFile();
        console.log('3. 施設カレンダーPDFファイルを生成中...');
        const facilityPdfResult = exportFacilityCalendarToPDF();
        console.log('4. 汎用PDFファイルを生成中...');
        const generalPdfResult = exportSpreadsheetToPDF();
        console.log('===== 処理完了 =====');
        if (csvResult) {
            console.log('CSVファイルのダウンロードURL:');
            console.log(csvResult.downloadUrl);
        }
        console.log('施設カレンダーPDFのダウンロードURL:');
        console.log(facilityPdfResult.downloadUrl);
        console.log('汎用PDFのダウンロードURL:');
        console.log(generalPdfResult.downloadUrl);
        return {
            csv: csvResult,
            facilityPdf: facilityPdfResult,
            generalPdf: generalPdfResult
        };
    }
    catch (error) {
        console.error('一括実行エラー:', error);
        throw error;
    }
}
global.exportCalendarToSpreadsheet = exportCalendarToSpreadsheet;
global.exportCsvSheetToFile = exportCsvSheetToFile;
global.exportFacilityCalendarToPDF = exportFacilityCalendarToPDF;
global.exportSpreadsheetToPDF = exportSpreadsheetToPDF;
global.executeAll = executeAll;
