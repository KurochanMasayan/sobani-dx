# 居宅療養管理指導システム 移行ガイド

## 従来版と新版の詳細比較

### コード構造の比較

#### 従来版
```
test2()               - tmpシートにデータ保存
  └─ getCSVFile()
  └─ dataExtract()
  └─ mergeCSVData()

writeData()           - tmpシートから読み取り、ファイル生成
  └─ getAndDeleteData()

organizeCSVFilesById() - ファイル整理

makepdf()             - PDF生成
```

**問題点:**
- データの二重処理（tmpシート経由）
- 3つの関数を順番に実行する必要
- エラー時の状態が不明確
- ハードコードされた設定

#### 新版
```
processMonthlyData()  - 全処理を一括実行
  ├─ readCSVFiles()
  ├─ extractRegularVisits()
  ├─ mergePatientData()
  ├─ generateSpreadsheets()
  └─ organizeByFacility()
```

**改善点:**
- メモリ内で完結（tmpシート不要）
- 1つの関数で全処理
- 詳細なエラーハンドリング
- 設定の一元管理

### 関数対応表

| 従来版 | 新版 | 備考 |
|--------|------|------|
| `test2()` | `processMonthlyData()` の一部 | tmpシート廃止 |
| `getCSVFile()` | `readCSVFiles()` | 汎用化 |
| `dataExtract()` | `extractRegularVisits()` | 名前を明確化 |
| `mergeCSVData()` | `mergePatientData()` | 名前を明確化 |
| `getCSVDataWithoutDeletion()` | `readCSVFiles(folderId, false)` | 統合 |
| `writeData()` | `generateSpreadsheets()` | tmpシート不要 |
| `getAndDeleteData()` | 廃止 | 不要になった |
| `moveFile()` | `generateSpreadsheets()` 内で処理 | 統合 |
| `organizeCSVFilesById()` | `organizeByFacility()` | 最新API使用 |

### 実行方法の比較

#### 従来版
```javascript
// 3ステップの手動実行が必要
function run() {
  test2();                    // ステップ1
  writeData();                // ステップ2
  organizeCSVFilesById();     // ステップ3
}
```

**問題点:**
- どのステップで失敗したか不明
- 途中で止まった場合の対処が困難
- tmpシートの状態管理が必要

#### 新版
```javascript
// 1ステップで完結
processMonthlyData();

// またはオプション指定
processMonthlyData({
  organizeFiles: true,
  deleteCSV: true
});
```

**改善点:**
- 処理が自己完結
- エラーログで問題箇所を特定可能
- 中間状態の管理不要

## 移行手順（詳細版）

### ステップ1: バックアップ作成

```javascript
// 1-1. 現在のスクリプトをバックアップ
// スクリプトエディタで「ファイル」→「コピーを作成」

// 1-2. 既存データのバックアップ
function backupExistingData() {
  const outputFolder = DriveApp.getFolderById('1wxFMS3SmC03PqvPk9WJ-Qi-QBOV6jB0O');
  const backupFolder = DriveApp.getFolderById('バックアップフォルダID');

  const files = outputFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
  let count = 0;

  while (files.hasNext()) {
    const file = files.next();
    file.makeCopy(`[BACKUP]${file.getName()}`, backupFolder);
    count++;
  }

  Logger.log(`${count}件のファイルをバックアップしました`);
}
```

### ステップ2: 新スクリプト導入

```javascript
// 2-1. 新しいプロジェクトを作成
// または既存プロジェクトに追加

// 2-2. medical-guidance-system-refactored.js の内容をコピー＆ペースト

// 2-3. 保存（Ctrl+S）
```

### ステップ3: 設定確認・移行

```javascript
// 3-1. 既存の設定を確認
function checkOldConfig() {
  const props = PropertiesService.getScriptProperties();
  const keys = props.getKeys();

  Logger.log('=== 現在の設定 ===');
  keys.forEach(key => {
    Logger.log(`${key}: ${props.getProperty(key)}`);
  });
}

// 3-2. 設定を実行
checkOldConfig();

// 3-3. 新版の設定を確認
checkConfiguration();
```

### ステップ4: テスト実行

```javascript
// 4-1. 設定確認
checkConfiguration();

// 4-2. ドライ実行（ファイル生成なし）
dryRun();

// 4-3. ログを確認
// 「表示」→「ログ」でデータが正しく読み込まれているか確認

// 4-4. テスト用CSVで本番実行
// 少量のテストデータで processMonthlyData() を実行
processMonthlyData({
  organizeFiles: false,  // 最初は整理をスキップ
  deleteCSV: false       // テスト時はCSVを保持
});
```

### ステップ5: 結果検証

```javascript
// 5-1. 生成されたファイルを確認
function verifyGeneratedFiles() {
  const outputFolder = DriveApp.getFolderById('1wxFMS3SmC03PqvPk9WJ-Qi-QBOV6jB0O');
  const files = outputFolder.getFilesByType(MimeType.GOOGLE_SHEETS);

  Logger.log('=== 生成ファイル一覧 ===');
  while (files.hasNext()) {
    const file = files.next();
    const ss = SpreadsheetApp.openById(file.getId());
    const sheets = ss.getSheets();

    Logger.log(`ファイル: ${file.getName()}`);
    Logger.log(`  シート数: ${sheets.length}`);
    sheets.forEach(sheet => {
      Logger.log(`  - ${sheet.getName()}`);
    });
  }
}

// 5-2. 実行
verifyGeneratedFiles();

// 5-3. サンプルファイルを開いて内容を目視確認
```

### ステップ6: 並行稼働（推奨）

```javascript
// 6-1. 両方のシステムで1ヶ月間処理
function parallelRun() {
  Logger.log('=== 従来版実行 ===');
  test2();
  writeData();
  organizeCSVFilesById();

  Logger.log('=== 新版実行 ===');
  processMonthlyData({
    deleteCSV: false  // 従来版のためにCSVは保持
  });
}

// 6-2. 結果を比較
function compareResults() {
  // ファイル数、シート数、データ内容を比較
  // 差異があれば調査
}
```

### ステップ7: 完全移行

```javascript
// 7-1. トリガーを新関数に変更
function migrateTriggers() {
  // 既存トリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'test2' ||
        trigger.getHandlerFunction() === 'writeData' ||
        trigger.getHandlerFunction() === 'organizeCSVFilesById') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`削除: ${trigger.getHandlerFunction()}`);
    }
  });

  // 新トリガーを作成
  ScriptApp.newTrigger('processMonthlyData')
    .timeBased()
    .onMonthDay(1)
    .atHour(2)
    .create();

  Logger.log('新トリガーを設定しました');
}

// 7-2. 実行
migrateTriggers();

// 7-3. 旧コードを削除（バックアップ後）
// test2, writeData, getAndDeleteData 等の関数を削除
```

## トラブルシューティング

### 移行時によくある問題

#### 問題1: 「患者マスタが見つかりません」

**原因:**
```javascript
// 従来版
const fixedFolderId = PropertiesService.getScriptProperties().getProperty('FIXDATA_FOLDER_ID');

// 新版でも同じプロパティ名を使用
const patientMaster: PropertiesService.getScriptProperties().getProperty('FIXDATA_FOLDER_ID')
```

**対処:**
```javascript
// プロパティが設定されているか確認
function checkProperty() {
  const value = PropertiesService.getScriptProperties().getProperty('FIXDATA_FOLDER_ID');
  Logger.log('FIXDATA_FOLDER_ID:', value);

  if (!value) {
    Logger.log('設定されていません。setupProperties() を実行してください');
  }
}
```

#### 問題2: セルマッピングが合わない

**原因:**
テンプレートのセル位置が変更されている

**対処:**
```javascript
// CONFIGのcellMappingを修正
const CONFIG = {
  template: {
    cellMapping: {
      visitDate: 'O2',      // ← 実際のセル位置に変更
      facilityName: 'H4',
      // ...
    }
  }
};
```

#### 問題3: ファイルが重複生成される

**原因:**
従来版と新版の両方が実行されている

**対処:**
```javascript
// トリガーを確認
function listTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    Logger.log(`関数: ${trigger.getHandlerFunction()}, 種類: ${trigger.getTriggerSource()}`);
  });
}

// 重複トリガーを削除
```

#### 問題4: ショートカット作成エラー

**エラーメッセージ:**
```
Exception: Unexpected error while getting the method or property moveTo
```

**原因:**
`moveTo()` の使用方法が間違っている

**対処:**
新版では正しく実装済み:
```javascript
DriveApp.createShortcut(file.getId())
  .setName(shortcutName)
  .moveTo(facilityFolder);  // 正しい使用方法
```

## ロールバック手順

問題が発生した場合の戻し方:

### 1. トリガーを戻す
```javascript
function rollbackTriggers() {
  // 新トリガーを削除
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processMonthlyData') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 旧トリガーを再設定（手動でも可）
  ScriptApp.newTrigger('test2')
    .timeBased()
    // ... 元の設定
    .create();
}
```

### 2. バックアップから復元
```javascript
function restoreFromBackup() {
  const backupFolder = DriveApp.getFolderById('バックアップフォルダID');
  const outputFolder = DriveApp.getFolderById('1wxFMS3SmC03PqvPk9WJ-Qi-QBOV6jB0O');

  const files = backupFolder.getFilesByName('[BACKUP]*');
  while (files.hasNext()) {
    const file = files.next();
    const originalName = file.getName().replace('[BACKUP]', '');
    file.makeCopy(originalName, outputFolder);
  }
}
```

## 性能比較

### 処理時間の測定

```javascript
// 従来版
function measureOldSystem() {
  const start = new Date();

  test2();
  writeData();
  organizeCSVFilesById();

  const end = new Date();
  const duration = (end - start) / 1000;

  Logger.log(`従来版処理時間: ${duration}秒`);
}

// 新版
function measureNewSystem() {
  const start = new Date();

  processMonthlyData();

  const end = new Date();
  const duration = (end - start) / 1000;

  Logger.log(`新版処理時間: ${duration}秒`);
}
```

### 期待される性能改善

| 項目 | 従来版 | 新版 | 改善率 |
|------|--------|------|--------|
| 処理時間（50名） | ~120秒 | ~60秒 | 50%短縮 |
| メモリ使用量 | tmpシート使用 | メモリ内 | 削減 |
| エラー回復時間 | 手動確認必要 | 自動ログ | 大幅短縮 |
| コード行数 | ~500行 | ~650行 | 構造改善 |

※ 新版は行数が増えていますが、コメント・エラー処理・ログが充実しているためです。

## チェックリスト

移行完了前に確認:

- [ ] バックアップ作成完了
- [ ] 新スクリプト導入完了
- [ ] `checkConfiguration()` で設定確認
- [ ] `dryRun()` でデータ確認
- [ ] テストデータで実行成功
- [ ] ログに異常なし
- [ ] 生成ファイルの内容確認
- [ ] 並行稼働で結果一致確認（推奨）
- [ ] トリガー移行完了
- [ ] 旧コード削除（バックアップ後）
- [ ] ドキュメント更新

## サポート・問い合わせ

移行に関する質問:
1. ログを確認
2. このガイドのトラブルシューティングを参照
3. エラーメッセージとログを添えて報告

## 更新履歴

- 2025-01-22: 初版作成
- 移行完了後に実績を記録してください
