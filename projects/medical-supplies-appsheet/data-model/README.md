# 訪問診療 医材在庫管理システム データモデル

このディレクトリには、訪問診療における医材在庫管理システムのデータベース設計書が含まれています。

## ディレクトリ構成

```
data-model/
├── README.md               # このファイル
├── documentation/          # テーブル設計書
│   ├── 00_システム概要.md   # システム全体の概要
│   ├── 01_施設マスタ.md    # facilities テーブル定義
│   ├── 02_患者マスタ.md    # patients テーブル定義
│   ├── 03_商品マスタ.md    # products テーブル定義
│   ├── 04_発注先マスタ.md  # suppliers テーブル定義
│   ├── 05_配布テンプレート.md # distribution_templates テーブル定義
│   ├── 06_在庫管理.md      # inventory テーブル定義
│   ├── 07_配布記録.md      # distribution_records テーブル定義
│   ├── 08_発注管理.md      # purchase_orders テーブル定義
│   └── 09_在庫変動履歴.md  # inventory_transactions テーブル定義
└── csv-data/              # CSVデータファイル（実データ格納用）
```

## テーブル概要

| テーブル名 | 説明 | 主キー |
|-----------|------|--------|
| facilities | 施設マスタ | facility_id |
| patients | 患者マスタ | patient_id |
| products | 商品マスタ | product_id |
| suppliers | 発注先マスタ | supplier_id |
| distribution_templates | 配布テンプレート | template_id |
| inventory | 在庫管理 | product_id |
| distribution_records | 配布記録 | distribution_id |
| purchase_orders | 発注管理 | order_id |
| inventory_transactions | 在庫変動履歴 | transaction_id |

## 主要機能

1. **施設・患者管理**: 訪問先施設と患者情報の管理
2. **医材マスタ管理**: 取り扱い医材の基本情報管理
3. **配布テンプレート管理**: 患者別の定期配布パターン管理（偶数月・奇数月対応）
4. **医材配布記録**: 施設への配布実績管理
5. **在庫管理**: 箱単位・個数単位での在庫管理
6. **発注管理**: 発注から入荷までの管理
7. **在庫変動履歴**: 全ての在庫変動を記録

## データ形式

- 文字コード: UTF-8（BOM付き）
- CSVヘッダー: 日本語カラム名を使用
- 日付形式: YYYY-MM-DD
- 日時形式: YYYY-MM-DD HH:MM:SS
- ID形式:
  - 施設ID: faci-XXX
  - 患者ID: 6桁数字
  - 商品ID: PRD-XXX
  - 発注先ID: store-XX
  - テンプレートID: TMPL-XXX
  - その他ID: UniqueID（自動生成）

## AppSheet実装手順

1. 各テーブル設計書（documentation/）を参照
2. Google スプレッドシートに各テーブルを作成
3. AppSheetでテーブルを追加
4. カラム型とリレーションを設定
5. 計算式と初期値を設定
6. ビューとアクションを作成

## 更新履歴

- 2024-09-20: 要件定義書に基づく全面改訂
- 2024-09-19: 初版作成