# Google Apps Script TypeScript プロジェクト

このプロジェクトはTypeScriptで記述され、claspを使用してGoogle Apps Scriptにデプロイします。

## セットアップ

### 1. 依存関係のインストール
```bash
npm install
```

### 2. claspの認証
```bash
npx clasp login
```

### 3. プロジェクトのビルド
```bash
npm run build
```

### 4. Google Apps Scriptへのデプロイ
```bash
npm run push
```

## プロジェクト構造

```
.
├── src/                  # TypeScriptソースコード
│   ├── main.ts          # メイン処理
│   ├── config.ts        # 設定
│   ├── calendarSync.ts  # カレンダー同期
│   ├── csvProcessor.ts  # CSV処理
│   └── pdfExporter.ts   # PDF出力
├── dist/                # コンパイル済みJavaScript（自動生成）
├── .clasp.json          # claspプロジェクト設定
├── .claspignore         # claspアップロード除外設定
├── tsconfig.json        # TypeScript設定
└── package.json         # npm設定
```

## 開発コマンド

- `npm run build` - TypeScriptをJavaScriptにコンパイル
- `npm run push` - ビルド後、GASにアップロード
- `npm run watch` - ファイル変更を監視して自動ビルド

## 注意事項

- `.clasprc.json`は個人の認証情報を含むため、絶対にコミットしないでください
- `dist/`フォルダは自動生成されるため、コミット不要です
- GAS固有のグローバル変数（SpreadsheetApp等）は`@types/google-apps-script`で型定義されています

## 変換前のGSファイルとの互換性

すべての関数は元のGSファイルと同じ動作をするよう正確に変換されています。
TypeScript化により型安全性とIDEサポートが向上しています。