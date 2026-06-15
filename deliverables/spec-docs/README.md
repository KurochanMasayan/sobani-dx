# そばに DX 仕様書（納品物）

そばに各ツールの概要仕様書（PDF）と、その生成元 Markdown・ビルド環境をまとめたフォルダ。
内製運用への引き継ぎを目的とした「簡単な仕様書」。

## 構成

```
deliverables/spec-docs/
├── pdf/      ← 納品用 PDF（これを共有する）
├── src/      ← 各仕様書の Markdown（編集はこちら）
└── build/    ← 生成スクリプト（build.mjs / style.css / preview.mjs）
```

## ドキュメント一覧

| No. | ファイル | 対象ツール |
| --- | --- | --- |
| 00 | Sobani-DX_全体概要 | （全体俯瞰） |
| 01 | カレンダー同期PDF出力システム | calendar-gas |
| 02 | 医材在庫管理システム | medical-supplies-appsheet |
| 03 | 歯科訪問スケジュール管理システム | dental-schedule-appsheet |
| 04 | 歯科検診報告書システム | dental-examination-report |
| 05 | 居宅療養管理指導書 自動作成システム | home-care-guidance-document |
| 06 | 在宅診療計画書 自動作成システム | home-care-plan-document |

## PDF の再生成

```bash
cd deliverables/spec-docs/build
npm install              # 初回のみ（markdown-it / puppeteer-core）
node build.mjs           # src/ の全 .md を pdf/ へ
node build.mjs 02_医材在庫管理システム.md   # 1冊だけ
```

- ローカルの Google Chrome を使って PDF 化する（`build.mjs` の `CHROME_CANDIDATES` 参照）。
- 各 `.md` 冒頭の `<!--META ... -->` で表紙のタイトル・版数・発行日・提出先を設定する。
- 章番号・目次・ヘッダ/フッタ・ページ番号は自動付与。

## 方針メモ

- **想定読者**: ビルド等の開発作業はできない前提の担当者が、もらって勉強する用途。大枠と「どこをどう直すか」が分かることを目的とする。
- **粒度**: 日々の運用に必要な範囲の概要仕様。内部実装には深入りしない。現仕様との一致確認は最低限。
- **修正方法**: 各冊に「修正のしかた」を記載。clasp/ビルドは使わず、**GAS はエディタにコピー&ペースト**／**AppSheet は編集画面**で直す方針で統一。
- **AppSheet ツール（02・03・04）**: 実機とドキュメントが食い違い得るため、細部に踏み込まず「実機が正」と明記。
- **GAS ツール（01・05・06）**: ソースコードが正。
