# 居宅療養管理指導システム - リファクタリング版

## 概要

月次の訪問データと患者マスタを統合し、患者別・日付別のスプレッドシートを自動生成するシステムです。

## 主な改善点

### 従来版との比較

| 項目 | 従来版 | 新版 |
|------|--------|------|
| 関数数 | 10+ | 主要5関数 |
| 実行ステップ | 3ステップ | 1ステップ |
| tmpシート | 必須 | 不要 |
| エラー処理 | 不十分 | 強化 |
| 保守性 | 低 | 高 |

### 改善内容

1. **tmpシート廃止** - メモリ内で直接処理
2. **設定の一元管理** - CONFIG オブジェクトで集中管理
3. **単一関数で完結** - `processMonthlyData()` のみで全処理
4. **最新API使用** - 非推奨メソッドを排除
5. **詳細ログ** - 処理状況を詳細に記録

## セットアップ

### 1. スクリプトプロパティの設定

```javascript
function setupProperties() {
  PropertiesService.getScriptProperties().setProperties({
    'MONTH_DATA_FOLDER_ID': 'xxxxxxxxxx',    // 月次データフォルダID
    'FIXDATA_FOLDER_ID': 'xxxxxxxxxx',        // 患者マスタフォルダID
    'SHORT_FOLDER_ID': 'xxxxxxxxxx'           // 整理先フォルダID（オプション）
  });
}
```

### 2. フォルダ構成

```
Google Drive
├── 居宅療養管理指導/
│   ├── 入力/
│   │   ├── モバカル_2024-10.csv (月次データ - 処理後削除)
│   │   └── 基本情報/
│   │       └── 患者情報.csv (マスタ - 削除しない)
│   └── 出力/
│       ├── ケアプランセンターわかば-佐々木 信夫.xlsx
│       └── デイサービス太陽-田中 花子.xlsx
```

### 3. CSV形式

#### 月次データ (モバカル)
必須列: `orcaID`, `日付`, `定期訪問`

#### 患者マスタ
必須列: `患者ID`, `患者名`, `生年月日`, `年齢`, `性別`, `郵便番号`, `住所`, `診断名`, `居宅介護支援事業所`, `寝たきり度`, `認知度`, `介護認定`

## 使い方

### 基本的な使用方法

#### 1. 設定確認
```javascript
checkConfiguration();
```

実行結果例:
```
[INFO] === 設定確認 ===
[INFO] ✓ monthDataFolder: 1abc...
[INFO] ✓ patientMasterFolder: 2def...
[INFO] ✓ outputFolder: 1wxF...
[INFO] ✓ organizedFolder: 3ghi...
[INFO] ✓ templateId: 1JLV...
```

#### 2. ドライ実行（テスト）
```javascript
dryRun();
```

データ読み込みと統合のみを行い、ファイルは生成しません。データ確認用。

#### 3. 本番実行
```javascript
processMonthlyData();
```

全処理を実行します。

#### 4. オプション指定実行
```javascript
processMonthlyData({
  organizeFiles: false,  // ファイル整理をスキップ
  deleteCSV: false       // CSVファイルを削除しない
});
```

### トリガー設定

#### 時間ベーストリガーの設定例

```javascript
function setupMonthlyTrigger() {
  // 既存のトリガーを削除
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processMonthlyData') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎月1日の午前2時に実行
  ScriptApp.newTrigger('processMonthlyData')
    .timeBased()
    .onMonthDay(1)
    .atHour(2)
    .create();

  Logger.log('月次トリガーを設定しました');
}
```

または手動で設定:
1. スクリプトエディタを開く
2. 左メニューの「トリガー」をクリック
3. 「トリガーを追加」
4. 関数: `processMonthlyData`
5. イベントのソース: 時間主導型
6. 頻度を設定（例: 毎月、1日、午前2時）

## 処理フロー

```
1. 月次CSVファイル読み込み
   ↓
2. 定期訪問データ抽出
   ↓
3. 患者マスタ読み込み
   ↓
4. データ統合（患者ID でマッチング）
   ↓
5. 患者別スプレッドシート生成
   - 既存ファイル: シート追加
   - 新規患者: ファイル新規作成
   ↓
6. 施設別フォルダ整理（オプション）
```

## 生成されるファイル

### ファイル名形式
```
{施設名}-{患者名}.xlsx
```

例: `ケアプランセンターわかば-佐々木 信夫.xlsx`

### シート名形式
```
yyyy-MM-dd
```

例: `2024-10-15`

### データ配置

テンプレートの以下のセルに自動入力:
- O2: 訪問日
- H4: 施設名
- G7: 患者名
- T7: 性別
- G8: 生年月日
- R8: 年齢
- H9: 郵便番号
- G10: 住所
- F11: 訪問日（重複）
- F13: 診断名
- G29: 介護度
- G30: 寝たきり度
- G31: 認知症度

## トラブルシューティング

### エラー: 「CSVファイルが見つかりません」

**原因**:
- 月次CSVファイルが投入されていない
- フォルダIDが間違っている

**対処法**:
1. `checkConfiguration()` で設定確認
2. フォルダ内にCSVファイルが存在するか確認

### エラー: 「患者マスタが見つかりません」

**原因**:
- 患者マスタCSVが配置されていない
- FIXDATA_FOLDER_ID が未設定

**対処法**:
1. 患者情報.csv を基本情報フォルダに配置
2. スクリプトプロパティを確認

### エラー: 「必要な列が見つかりません」

**原因**:
- CSVファイルの列名が想定と異なる

**対処法**:
CONFIG オブジェクトの列名設定を確認:
```javascript
monthDataColumns: {
  patientId: 'orcaID',      // ← この列名がCSVに存在するか
  date: '日付',
  status: '定期訪問'
}
```

### 処理が遅い

**原因**:
- 大量のデータ処理
- APIクォータ制限

**対処法**:
1. データを分割して処理
2. バッチサイズを調整

## ログの確認

### 実行ログの表示

1. スクリプトエディタを開く
2. 「表示」→「ログ」または Ctrl+Enter
3. 処理状況を確認

### ログの例

```
[INFO] === 月次データ処理開始 ===
[INFO] ステップ1: 月次データ読み込み
[INFO] CSVファイルを削除しました {"fileName":"モバカル_2024-10.csv"}
[INFO] ステップ2: 定期訪問データ抽出
[INFO] 定期訪問データを抽出しました {"count":45}
[INFO] ステップ3: 患者マスタ読み込み
[INFO] ステップ4: データ統合
[INFO] データ統合完了 {"patientCount":38}
[INFO] ステップ5: スプレッドシート生成
[INFO] スプレッドシート生成開始 {"patientCount":38}
[INFO] 新規ファイルを作成しました {"fileName":"ケアプランA-田中太郎"}
[INFO] シートを作成しました {"fileName":"ケアプランA-田中太郎","date":"2024-10-15"}
...
[INFO] === 月次データ処理完了 === {"duration":"45.2秒","patientCount":38}
```

## CONFIG 設定詳細

### フォルダID設定

```javascript
folders: {
  monthData: 'プロパティから取得',
  patientMaster: 'プロパティから取得',
  output: '1wxFMS3SmC03PqvPk9WJ-Qi-QBOV6jB0O',  // 固定
  organized: 'プロパティから取得',
  pdfOutput: '1r0CfXFfsFozN_Mf0BtICGJgccdwY2lrq'  // 固定
}
```

### セルマッピングのカスタマイズ

テンプレートのセル位置が変わった場合:

```javascript
cellMapping: {
  visitDate: 'O2',      // 訪問日のセル
  facilityName: 'H4',   // 施設名のセル
  // ... 他のセル位置
}
```

## 移行ガイド

### 従来版から新版への移行手順

1. **バックアップ作成**
   - 現在のスクリプトをコピー保存
   - 既存のスプレッドシートをバックアップ

2. **新スクリプト導入**
   - `medical-guidance-system-refactored.js` をコピー
   - Google Apps Script エディタに貼り付け

3. **設定移行**
   - 既存のスクリプトプロパティを確認
   - 新スクリプト用に設定

4. **テスト実行**
   - `checkConfiguration()` で設定確認
   - `dryRun()` でデータ確認
   - テスト用CSVで `processMonthlyData()` 実行

5. **並行稼働**
   - 1ヶ月間、両方のシステムで処理
   - 結果を比較検証

6. **完全移行**
   - 問題なければ旧システムを削除
   - トリガーを新関数に変更

## よくある質問

### Q: 既存の患者ファイルはどうなりますか？

A: 既存ファイルがある場合、新しい訪問日のシートが追加されます。既存のシートは上書きされません。

### Q: CSVファイルは自動で削除されますか？

A: デフォルトでは削除されます。削除したくない場合は `processMonthlyData({ deleteCSV: false })` を実行してください。

### Q: エラーが発生した場合、処理は中断されますか？

A: 患者単位でエラーハンドリングされているため、一部の患者でエラーが発生しても他の患者の処理は継続されます。

### Q: 処理時間はどのくらいですか？

A: 患者数によりますが、50名程度で1-2分程度です。

## サポート

問題が発生した場合:

1. ログを確認
2. `checkConfiguration()` で設定確認
3. `dryRun()` でデータ確認
4. エラーメッセージを記録して報告

## ライセンス・履歴

- 作成日: 2025-01-22
- バージョン: 2.0 (リファクタリング版)
- 従来版からの主な変更: tmpシート廃止、処理統合、エラーハンドリング強化
