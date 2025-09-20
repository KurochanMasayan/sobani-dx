# Sobani プロジェクトワークスペース

複数のプロジェクトを統合管理するワークスペースです。Google Apps ScriptとAppSheetを中心とした業務効率化ツールを開発・管理しています。

## プロジェクト構成

### 📅 Calendar GAS Project (`projects/calendar-gas/`)
- **概要**: Google Apps Scriptを使用したカレンダー同期システム
- **言語**: TypeScript
- **機能**: カレンダー同期、CSV処理、PDF出力
- **詳細**: [Calendar GAS README](projects/calendar-gas/README.md)

### 🏥 Medical Supplies AppSheet (`projects/medical-supplies-appsheet/`)
- **概要**: 医療機器・医療材料の管理AppSheetアプリケーション
- **プラットフォーム**: AppSheet + Google Apps Script
- **機能**: 在庫管理、発注管理、使用履歴管理
- **詳細**: [Medical Supplies README](projects/medical-supplies-appsheet/README.md)

## ワークスペース管理

### セットアップ
```bash
# 全プロジェクトの依存関係をインストール
npm run install:all

# 個別プロジェクトのセットアップ
npm run install:calendar
npm run medical:setup
```

### 開発コマンド

#### カレンダーGASプロジェクト
```bash
npm run calendar:build    # TypeScriptをビルド
npm run calendar:push     # GASにデプロイ
npm run calendar:watch    # ファイル変更を監視
```

#### 医材管理AppSheetプロジェクト
```bash
npm run medical:setup     # 初期セットアップ
npm run medical:backup    # データバックアップ
npm run medical:deploy    # AppSheetデプロイ
```

#### 全体管理
```bash
npm run build:all         # 全プロジェクトをビルド
npm run lint:all          # 全プロジェクトをlint
npm run clean             # node_modulesを削除
```

## フォルダ構造

```
sobani/
├── projects/                           # プロジェクト群
│   ├── calendar-gas/                   # カレンダー同期GAS
│   │   ├── src/                        # TypeScriptソース
│   │   ├── package.json                # 依存関係
│   │   └── README.md                   # プロジェクト詳細
│   └── medical-supplies-appsheet/      # 医材管理AppSheet
│       ├── appsheet/                   # AppSheet設定
│       ├── gas-integration/            # GAS連携
│       ├── documents/                  # ドキュメント
│       ├── scripts/                    # 運用スクリプト
│       └── README.md                   # プロジェクト詳細
├── shared/                             # 共通リソース
│   ├── configs/                        # 共通設定
│   ├── templates/                      # テンプレート
│   └── utils/                          # 共通ユーティリティ
├── .claude/                            # Claude設定
├── package.json                        # ワークスペース設定
└── README.md                           # このファイル
```

## 開発環境要件

- **Node.js**: >= 18.0.0
- **npm**: >= 9.0.0
- **Google Apps Script CLI**: clasp
- **Google Workspace**: AppSheet利用のため

### 認証設定

#### Google Apps Script
```bash
cd projects/calendar-gas
npx clasp login
```

#### AppSheet
- Google Workspaceアカウントでの認証が必要
- AppSheetコンソールでのアプリ設定

## プロジェクト追加ガイド

新しいプロジェクトを追加する場合：

1. `projects/` 配下に新しいディレクトリを作成
2. プロジェクト固有の `package.json` を作成
3. ルートの `package.json` にスクリプトを追加
4. `README.md` を更新

## 貢献ガイドライン

- 各プロジェクトの開発ガイドラインに従う
- TypeScriptプロジェクトは型安全性を重視
- AppSheetプロジェクトは設定をJSONで管理
- 変更は適切にドキュメント化する

## ライセンス

MIT License

## サポート

- プロジェクト固有の問題は各プロジェクトのREADMEを参照
- ワークスペース全体に関する問題はIssueを作成