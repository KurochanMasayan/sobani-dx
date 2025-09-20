# データモデル設計書

## 1. テーブル設計概要

医材管理システムは以下の4つの主要テーブルで構成されています：

1. **supplies** - 医材マスタ
2. **inventory** - 在庫管理
3. **orders** - 発注管理
4. **usage_logs** - 使用履歴

## 2. テーブル詳細設計

### 2.1 supplies（医材マスタ）

| フィールド名 | データ型 | 必須 | 説明 | 例 |
|-------------|---------|------|------|---|
| supply_id | Text | ✓ | 医材ID（主キー） | SUP001 |
| supply_name | Text | ✓ | 医材名 | 使い捨て手袋 |
| category | Enum | ✓ | カテゴリ | 消耗品/機器/薬剤 |
| subcategory | Text |  | サブカテゴリ | 感染防護用品 |
| manufacturer | Text |  | メーカー名 | ○○医療器械 |
| model_number | Text |  | 型番 | GLV-001 |
| jan_code | Text |  | JANコード | 4901234567890 |
| unit_price | Number |  | 単価 | 50 |
| unit | Text | ✓ | 単位 | 箱/個/本 |
| supplier | Text |  | 供給業者 | ××商事 |
| safety_stock | Number | ✓ | 安全在庫数 | 100 |
| reorder_point | Number | ✓ | 発注点 | 150 |
| storage_location | Text |  | 保管場所 | 倉庫A-1 |
| expiry_managed | Boolean |  | 有効期限管理 | TRUE/FALSE |
| description | LongText |  | 説明 | 滅菌済み使い捨て手袋 |
| is_active | Boolean | ✓ | 有効フラグ | TRUE |
| created_at | DateTime | ✓ | 作成日時 | 2024-01-01 10:00:00 |
| updated_at | DateTime | ✓ | 更新日時 | 2024-01-01 10:00:00 |

### 2.2 inventory（在庫管理）

| フィールド名 | データ型 | 必須 | 説明 | 例 |
|-------------|---------|------|------|---|
| inventory_id | Text | ✓ | 在庫ID（主キー） | INV001 |
| supply_id | Ref(supplies) | ✓ | 医材ID（外部キー） | SUP001 |
| current_stock | Number | ✓ | 現在庫数 | 250 |
| allocated_stock | Number |  | 引当済在庫 | 50 |
| available_stock | Number | ✓ | 利用可能在庫 | 200 |
| lot_number | Text |  | ロット番号 | LOT2024001 |
| expiry_date | Date |  | 有効期限 | 2024-12-31 |
| location | Text |  | 保管場所 | 倉庫A-1 |
| last_updated | DateTime | ✓ | 最終更新日時 | 2024-01-01 15:30:00 |
| updated_by | Text | ✓ | 更新者 | user@example.com |

### 2.3 orders（発注管理）

| フィールド名 | データ型 | 必須 | 説明 | 例 |
|-------------|---------|------|------|---|
| order_id | Text | ✓ | 発注ID（主キー） | ORD001 |
| supply_id | Ref(supplies) | ✓ | 医材ID（外部キー） | SUP001 |
| order_quantity | Number | ✓ | 発注数量 | 100 |
| unit_price | Number | ✓ | 単価 | 50 |
| total_amount | Number | ✓ | 合計金額 | 5000 |
| order_date | Date | ✓ | 発注日 | 2024-01-01 |
| requested_by | Text | ✓ | 申請者 | user@example.com |
| department | Text | ✓ | 申請部署 | 外科 |
| supplier | Text | ✓ | 発注先 | ××商事 |
| status | Enum | ✓ | ステータス | 申請中/承認済/発注済/入荷済/完了 |
| approved_by | Text |  | 承認者 | manager@example.com |
| approved_date | DateTime |  | 承認日時 | 2024-01-02 09:00:00 |
| expected_delivery | Date |  | 納期予定 | 2024-01-10 |
| actual_delivery | Date |  | 実際の納期 | 2024-01-09 |
| notes | LongText |  | 備考 | 緊急発注 |
| created_at | DateTime | ✓ | 作成日時 | 2024-01-01 10:00:00 |
| updated_at | DateTime | ✓ | 更新日時 | 2024-01-01 10:00:00 |

### 2.4 usage_logs（使用履歴）

| フィールド名 | データ型 | 必須 | 説明 | 例 |
|-------------|---------|------|------|---|
| usage_id | Text | ✓ | 使用ID（主キー） | USE001 |
| supply_id | Ref(supplies) | ✓ | 医材ID（外部キー） | SUP001 |
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
supplies (1) ←→ (n) inventory
supplies (1) ←→ (n) orders
supplies (1) ←→ (n) usage_logs
```

- 1つの医材に対して複数の在庫レコード（ロット管理）
- 1つの医材に対して複数の発注レコード
- 1つの医材に対して複数の使用履歴レコード

## 4. インデックス設計

### 4.1 検索性能向上のためのインデックス

- **supplies**: supply_id, category, is_active
- **inventory**: supply_id, expiry_date, current_stock
- **orders**: supply_id, status, order_date, department
- **usage_logs**: supply_id, usage_date, department

## 5. データ整合性ルール

### 5.1 在庫管理ルール
- available_stock = current_stock - allocated_stock
- current_stock >= 0
- allocated_stock >= 0

### 5.2 発注管理ルール
- total_amount = order_quantity × unit_price
- approved_date >= order_date
- actual_delivery >= order_date

### 5.3 使用履歴ルール
- usage_quantity > 0
- usage_date <= 今日の日付

## 6. AppSheet固有の設定

### 6.1 Key設定
- supplies: supply_id
- inventory: inventory_id
- orders: order_id
- usage_logs: usage_id

### 6.2 Ref設定
- inventory.supply_id → supplies.supply_id
- orders.supply_id → supplies.supply_id
- usage_logs.supply_id → supplies.supply_id

### 6.3 Virtual Column例
- supplies.total_inventory: SUM(inventory[current_stock])
- supplies.pending_orders: SUM(SELECT(orders[order_quantity], [supply_id] = [_THISROW].[supply_id], [status] = "発注済"))

## 7. データ移行計画

### 7.1 初期データ投入順序
1. supplies（医材マスタ）
2. inventory（初期在庫）
3. 過去の発注データ（必要に応じて）
4. 使用履歴（必要に応じて）

### 7.2 データ検証
- 参照整合性チェック
- 数値範囲チェック
- 必須項目チェック