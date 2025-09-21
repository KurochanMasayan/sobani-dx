# 医材管理 AppSheet プロジェクト

医療機器・医療材料の在庫管理、発注記録、使用履歴管理を行うAppSheetアプリケーションプロジェクトです。

## プロジェクト構造

```
medical-supplies-appsheet/
├── appsheet/                   # AppSheet関連ファイル
│   ├── app-definition.json     # AppSheetアプリ定義
│   ├── table-schemas/          # テーブルスキーマ定義
│   │   ├── supplies.json       # 医材マスタ
│   │   ├── inventory.json      # 在庫管理
│   │   ├── orders.json         # 発注記録
│   │   └── usage-logs.json     # 使用履歴
│   ├── workflows/              # ワークフロー定義
│   │   ├── order-approval.json # 発注承認フロー
│   │   └── stock-alert.json    # 在庫アラート
│   └── views/                  # ビュー定義
│       ├── dashboard.json      # ダッシュボード
│       ├── inventory-list.json # 在庫一覧
│       └── order-form.json     # 発注フォーム
├── gas-integration/            # GAS連携スクリプト
│   ├── src/
│   │   ├── csvExporter.ts      # CSV一括出力機能
│   │   ├── main.ts             # メイン実行関数
│   │   └── appsscript.json     # GAS設定
│   ├── package.json
│   ├── tsconfig.json
│   └── .clasp.json
├── documents/                  # プロジェクトドキュメント
│   ├── requirements.md         # 要件定義
│   ├── data-model.md          # データモデル設計
│   ├── user-manual.md         # ユーザーマニュアル
│   └── deployment.md          # デプロイ手順
├── scripts/                    # 運用スクリプト
│   ├── setup.sh               # 初期セットアップ
│   ├── backup.sh              # バックアップ
│   └── deploy.sh              # デプロイ
└── README.md                   # このファイル
```

## 主要機能

### 1. 医材マスタ管理
- 医療機器・材料の基本情報管理
- カテゴリ分類、メーカー情報
- 単価、供給業者情報

### 2. 在庫管理
- リアルタイム在庫数量管理
- 在庫切れアラート
- 適正在庫レベル設定

### 3. 発注記録
- 発注申請・承認ワークフロー
- 発注履歴管理
- 納期管理

### 4. 使用履歴管理
- 医材使用記録
- 患者情報との紐付け
- 使用量分析

### 5. データ出力機能
- スプレッドシート全シート一括CSV出力
- 主要シートのみ選択出力
- カスタムシート選択出力
- データバックアップとリポジトリ取り込み対応

## セットアップ手順

1. AppSheetでの新規アプリ作成
2. テーブルスキーマの設定
3. ワークフローの設定
4. ビューのカスタマイズ
5. GAS連携の設定（必要に応じて）

詳細は `documents/deployment.md` を参照してください。

## CSV一括出力機能

### 機能概要

医療用品管理スプレッドシートの全シート（約30シート）を一括でCSVファイルとして出力する機能です。データのバックアップやリポジトリへの取り込みに最適です。

### 使用方法

#### 1. 基本セットアップ

```bash
cd gas-integration
npm install
npx clasp login
# .clasp.jsonのscriptIdを設定
npm run push
```

#### 2. GASエディタでの実行

**全シート一括出力**:
```javascript
exportAllSheetsButton()
```

**主要シートのみ出力**:
```javascript
exportMainSheetsButton()
```

**カスタム選択出力**:
```javascript
exportCustomSheetsButton()
```

#### 3. 出力結果

- **自動フォルダ作成**: `医療用品データ出力_スプレッドシート名_YYYY-MM-DDTHH-MM`
- **個別ファイル**: 各シートが `シート名.csv` として保存
- **BOM付きUTF-8**: Excel対応、文字化けなし
- **詳細ログ**: 処理結果、エラー、フォルダURLを出力

#### 4. 戻り値の構造

```typescript
{
  success: boolean;             // 全体の成功可否
  totalSheets: number;          // 全シート数
  processedSheets: number;      // 処理成功シート数
  folderId: string;            // 出力先フォルダID
  folderUrl: string;           // 出力先フォルダURL
  spreadsheetName: string;     // スプレッドシート名
  files: Array<{               // 出力ファイル情報
    sheetName: string;
    fileName: string;
    fileId: string;
    downloadUrl: string;
    rowCount: number;
    columnCount: number;
    dataRange: string;
  }>;
  errors: Array<{              // エラー情報
    sheetName: string;
    error: string;
  }>;
}
```

#### 5. データ活用例

1. **リポジトリ取り込み**: CSVファイルをダウンロードして `data/` フォルダに配置
2. **データ分析**: Python/Rでの分析用データソース
3. **定期バックアップ**: スプレッドシートデータの自動バックアップ
4. **システム連携**: 他システムへのデータインポート

### 利用可能な関数

- `exportAllSheetsButton()` - 全シート一括出力（UI付き）
- `exportMainSheetsButton()` - 主要シートのみ出力
- `exportCustomSheetsButton()` - ユーザー選択シート出力
- `getSpreadsheetInfo()` - スプレッドシート情報取得（デバッグ用）
- `cleanupOldExportFolders()` - 古いCSV出力フォルダの削除

### 注意事項

- Google Apps Scriptの実行時間制限（6分）内での処理
- 大量データの場合は分割実行を推奨
- 出力フォルダは30日後に自動削除可能（cleanupOldExportFolders使用）

## 開発ガイドライン

- AppSheetの設定変更は必ずJSONファイルに反映すること
- データモデルの変更時は `documents/data-model.md` を更新すること
- 新機能追加時は要件定義を更新すること

## 関連プロジェクト

- `../calendar-gas/` - カレンダー同期GASプロジェクト（連携可能）
