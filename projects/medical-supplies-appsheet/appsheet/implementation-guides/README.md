# AppSheet実装ガイド

このフォルダには、AppSheetアプリケーションの具体的な実装方法やアクション設定のガイドラインを格納します。

## ガイド一覧

### 1. 01-batch-distribution-action-guide.md
配布テンプレートから複数レコードを選択し、一括で配布記録に複製するアクションの実装方法。

- ビューでの複数選択機能
- 選択状態でのみ表示されるアクション
- 実装手順とトラブルシューティング

### 2. 02-inventory-tracking-integration.md
配布記録から在庫変動履歴への自動連携実装方法。

- Automationボットによる統一管理
- フォーム入力・アクション作成両方に対応
- 在庫数量の自動更新とエラーハンドリング

### 3. 03-inventory-edit-delete-handling.md
入荷・棚卸・配布記録の編集・削除時の在庫管理更新方法。

- 在庫変動履歴の削除・再作成方式
- 各テーブルごとのAutomation設定
- 差分計算と在庫調整

### 4. 04-update-vs-recreate-comparison.md
在庫変動履歴の更新方式比較（上書き vs 削除→新規作成）。

- 両方式のメリット・デメリット比較
- アクション再利用の重要性
- 推奨実装パターンの解説

## フォルダ構造

```
/appsheet/
  ├── implementation-guides/  # 実装ガイド（このフォルダ）
  │   ├── README.md
  │   ├── 01-batch-distribution-action-guide.md
  │   ├── 02-inventory-tracking-integration.md
  │   ├── 03-inventory-edit-delete-handling.md
  │   └── 04-update-vs-recreate-comparison.md
  ├── table-schemas/         # テーブルスキーマ定義
  ├── views/                # ビュー設定
  └── workflows/            # ワークフローとオートメーション
```

## 関連ドキュメント

- データモデル定義: `/data-model/documentation/`
- プロジェクト要件: `/documents/requirements.md`
- データモデル概要: `/documents/data-model.md`