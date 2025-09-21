# データモデル設計書

## 1. テーブル設計概要

医材管理システムは以下の5つの主要テーブルで構成されています：

1. **商品マスタ** - 医材マスタ
2. **在庫管理** - 在庫管理
3. **発注管理** - 発注管理
4. **入荷管理** - 入荷記録
5. **使用履歴** - 使用履歴

## 2. テーブル詳細設計

### 2.1 商品マスタ

| フィールド名 | データ型 | 必須 | 説明 | 例 |
|-------------|---------|------|------|---|
| supply_id | Text | ✓ | 医材ID（主キー） | SUP001 |
| supply_name | Text | ✓ | 医材名 | 使い捨て手袋 |
| pieces_per_box | Number | ✓ | 箱入数 | 10 |
| unit | Text | ✓ | 単位 | 箱/個/本 |
| supplier_id | Text | ✓ | 発注先ID | store-01 |
| shelf_number | Text |  | 棚番号 | 1 |
| shelf_level | Text |  | 段数 | 5 |
| fixed_stock_flag | Yes/No |  | 定数在庫フラグ | TRUE |
| fixed_stock_boxes | Number |  | 定数在庫（箱単位） | 5 |

### 2.2 在庫管理

| フィールド名 | データ型 | 必須 | 説明 | 例 |
|-------------|---------|------|------|---|
| inventory_id | Text | ✓ | 在庫ID（主キー） | INV001 |
| supply_id | Ref(商品マスタ) | ✓ | 医材ID（外部キー） | SUP001 |
| stock_boxes | Number | ✓ | 在庫箱数 | 20 |
| stock_pieces | Number |  | 在庫個数（端数） | 5 |
| allocated_boxes | Number |  | 引当済在庫（箱） | 2 |
| available_boxes | Number | ✓ | 利用可能在庫（箱） | 18 |
| lot_number | Text |  | ロット番号 | LOT2024001 |
| expiry_date | Date |  | 有効期限 | 2024-12-31 |
| location | Text |  | 保管場所 | 倉庫A-1 |
| last_updated | DateTime | ✓ | 最終更新日時 | 2024-01-02 12:34:56 |
| updated_by | Text | ✓ | 更新者 | user@example.com |

### 2.3 発注管理

| フィールド名 | データ型 | 必須 | 説明 | 例 |
|-------------|---------|------|------|---|
| order_id | Text | ✓ | 発注ID（主キー） | ORD001 |
| supply_id | Ref(商品マスタ) | ✓ | 医材ID（外部キー） | SUP001 |
| order_quantity_boxes | Number | ✓ | 発注数量（箱単位） | 100 |
| unit_price | Number | ✓ | 単価 | 50 |
| total_amount | Number | ✓ | 合計金額 | 5000 |
| order_date | Date | ✓ | 発注日 | 2024-01-01 |
| requested_by | Text | ✓ | 申請者 | user@example.com |
| department | Text | ✓ | 申請部署 | 外科 |
| supplier | Text | ✓ | 発注先 | ××商事 |
| status | Enum | ✓ | ステータス | 申請中/承認済/発注済/一部入荷/入荷完了/キャンセル |
| approved_by | Text |  | 承認者 | manager@example.com |
| approved_date | DateTime |  | 承認日時 | 2024-01-02 09:00:00 |
| expected_delivery | Date |  | 納期予定 | 2024-01-10 |
| notes | LongText |  | 備考 | 緊急発注 |
| created_at | DateTime | ✓ | 作成日時 | 2024-01-01 10:00:00 |

### 2.4 入荷管理

| フィールド名 | データ型 | 必須 | 説明 | 例 |
|-------------|---------|------|------|---|
| receipt_id | Text | ✓ | 入荷ID（主キー） | RCV001 |
| order_id | Ref(orders) | ✓ | 紐づく発注ID | ORD001 |
| received_quantity_boxes | Number | ✓ | 入荷数量（箱単位） | 40 |
| delivery_date | Date | ✓ | 入荷日 | 2024-01-09 |
| lot_number | Text |  | 入荷ロット番号 | LOT2024002 |
| inspected_by | Text |  | 検品者 | staff@example.com |
| notes | LongText |  | 備考 | 分割入荷（第1便） |
| created_at | DateTime | ✓ | 作成日時 | 2024-01-09 10:00:00 |
| supply_id | Ref(商品マスタ) |  | 医材ID（発注から参照） | SUP001 |

### 2.5 使用履歴

| フィールド名 | データ型 | 必須 | 説明 | 例 |
|-------------|---------|------|------|---|
| usage_id | Text | ✓ | 使用ID（主キー） | USE001 |
| supply_id | Ref(商品マスタ) | ✓ | 医材ID（外部キー） | SUP001 |
| usage_quantity | Number | ✓ | 使用数量 | 10 |
| usage_date | Date | ✓ | 使用日 | 2024-01-01 |
| used_by | Text | ✓ | 使用者 | doctor@example.com |
| department | Text | ✓ | 使用部署 | 外科 |
| patient_id | Text |  | 患者ID | P001 |
| procedure_type | Text |  | 処置種別 | 手術 |
| lot_number | Text |  | 使用ロット番号 | LOT2024001 |
| cost_center | Text |  | コストセンター | CC001 |
| notes | LongText |  | 備考 | 緊急手術での使用 |
| created_at | DateTime | ✓ | 作成日時 | 2024-01-01 18:00:00 |

## 3. リレーションシップ

```
商品マスタ (1) ←→ (n) 在庫管理
商品マスタ (1) ←→ (n) 発注管理
発注管理 (1) ←→ (n) 入荷管理
商品マスタ (1) ←→ (n) 入荷管理
商品マスタ (1) ←→ (n) 使用履歴
```

- 1つの医材に対して複数の在庫レコード（ロット管理）
- 1つの医材に対して複数の発注レコード
- 1つの発注に対して複数の入荷レコード（分割入荷対応）
- 1つの医材に対して複数の入荷レコード
- 1つの医材に対して複数の使用履歴レコード

## 4. インデックス設計

### 4.1 検索性能向上のためのインデックス

- **商品マスタ**: supply_id, pieces_per_box, supplier_id
- **在庫管理**: supply_id, expiry_date, stock_boxes
- **発注管理**: supply_id, status, order_date, department
- **入荷管理**: order_id, delivery_date
- **使用履歴**: supply_id, usage_date, department

## 5. データ整合性ルール

### 5.1 在庫管理ルール
- available_stock = current_stock - allocated_stock
- current_stock >= 0
- allocated_stock >= 0

### 5.2 発注管理ルール
- total_amount = order_quantity_boxes × unit_price
- approved_date >= order_date
- status = "入荷完了" → SUM(SELECT(入荷管理[received_quantity_boxes], [order_id] = [_THISROW].[order_id])) = order_quantity_boxes

### 5.3 入荷管理ルール
- SUM(received_quantity_boxes) <= order_quantity_boxes
- delivery_date >= order_date
- 入荷登録時に 在庫管理[stock_boxes] と [stock_pieces] を更新し、端数がマイナスの場合は箱を調整

### 5.4 使用履歴ルール
- usage_quantity > 0
- usage_date <= 今日の日付

## 6. AppSheet固有の設定

### 6.1 Key設定
- 商品マスタ: supply_id
- 在庫管理: inventory_id
- 発注管理: order_id
- 入荷管理: receipt_id
- 使用履歴: usage_id

### 6.2 Ref設定
- 在庫管理.supply_id → 商品マスタ.supply_id
- 発注管理.supply_id → 商品マスタ.supply_id
- 入荷管理.order_id → 発注管理.order_id（IsPartOf）
- 使用履歴.supply_id → 商品マスタ.supply_id

### 6.3 Virtual Column例
- 商品マスタ.total_stock_pieces: SUM(SELECT(在庫管理[total_stock_pieces], [supply_id] = [_THISROW].[supply_id]))
- 商品マスタ.pending_orders: SUM(SELECT(発注管理[order_quantity_boxes], [supply_id] = [_THISROW].[supply_id], IN([status], {"申請中", "承認済", "発注済", "一部入荷"})))

## 7. データ移行計画

### 7.1 初期データ投入順序
1. 商品マスタ（医材マスタ）
2. 在庫管理（初期在庫）
3. 過去の発注データ（必要に応じて）
4. 使用履歴（必要に応じて）

### 7.2 データ検証
- 参照整合性チェック
- 数値範囲チェック
- 必須項目チェック
