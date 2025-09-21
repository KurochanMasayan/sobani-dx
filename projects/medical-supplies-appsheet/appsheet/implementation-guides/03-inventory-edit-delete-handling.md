# 入荷・棚卸の編集・削除時の在庫管理実装ガイド

## 概要

入荷記録・棚卸記録が編集または削除された際の在庫管理更新方法について説明します。

## 推奨方式：在庫変動履歴の削除・再作成

### 基本方針

1. **編集時**: 古い在庫変動履歴レコードを削除 → 新しい値で再作成
2. **削除時**: 関連する在庫変動履歴レコードを削除
3. **在庫数**: 差分を計算して更新

### メリット

- 実装がシンプル（複雑な差分計算が不要）
- 履歴が明確（1取引 = 1レコード）
- エラーが起きにくい
- データの整合性を保ちやすい

## 実装方法

### 1. 入荷記録の編集時

#### Automationボット設定

**Bot名**: `Handle_Receipt_Edit`

**イベント設定**:
```
Event Type: Data Change
Table: 入荷記録
Event: UPDATES
Condition: TRUE
```

**プロセス設定**:

##### Step 1: 古い在庫変動履歴を削除
```
Action Type: Delete rows
Table: 在庫変動履歴
Rows to delete:
  FILTER(
    "在庫変動履歴",
    AND(
      [入荷ID] = [_THISROW].[入荷ID],
      [変動区分] = "入荷"
    )
  )
```

##### Step 2: 新しい在庫変動履歴を作成
```
Action Type: Add a row
Table: 在庫変動履歴
Set these columns:
  - 変動日: [_THISROW].[入荷日]
  - 商品ID: [_THISROW].[商品ID]
  - 変動区分: "入荷"
  - 変動数量_箱: [_THISROW].[入荷数量_箱]
  - 変動数量_個: [_THISROW].[入荷数量_個]
  - 入荷ID: [_THISROW].[入荷ID]
  - 変動理由: CONCATENATE("入荷修正 - ", [_THISROW].[備考])
  - 記録者: USEREMAIL()
```

##### Step 3: 在庫数を更新
```
Action Type: Update row
Table: 在庫管理
Row to update:
  FILTER(
    "在庫管理",
    [商品ID] = [_THISROW].[商品ID]
  )
Set these columns:
  - 現在庫_箱:
    [現在庫_箱]
    - [_THISROW_BEFORE].[入荷数量_箱]
    + [_THISROW].[入荷数量_箱]
  - 現在庫_個:
    [現在庫_個]
    - [_THISROW_BEFORE].[入荷数量_個]
    + [_THISROW].[入荷数量_個]
  - 最終更新日: NOW()
```

### 2. 入荷記録の削除時

**Bot名**: `Handle_Receipt_Delete`

**イベント設定**:
```
Event Type: Data Change
Table: 入荷記録
Event: DELETES
Condition: TRUE
```

**プロセス設定**:

##### Step 1: 在庫変動履歴を削除
```
Action Type: Delete rows
Table: 在庫変動履歴
Rows to delete:
  FILTER(
    "在庫変動履歴",
    [入荷ID] = [_THISROW].[入荷ID]
  )
```

##### Step 2: 在庫数を減算
```
Action Type: Update row
Table: 在庫管理
Row to update:
  FILTER(
    "在庫管理",
    [商品ID] = [_THISROW].[商品ID]
  )
Set these columns:
  - 現在庫_箱: [現在庫_箱] - [_THISROW].[入荷数量_箱]
  - 現在庫_個: [現在庫_個] - [_THISROW].[入荷数量_個]
  - 最終更新日: NOW()
```

### 3. 棚卸記録の編集・削除

#### 棚卸編集時のボット

**Bot名**: `Handle_Inventory_Edit`

同様の構造で実装：
1. 古い在庫変動履歴（変動区分="棚卸調整"）を削除
2. 新しい調整値で在庫変動履歴を作成
3. 在庫数を差分で更新

```
在庫調整値の計算:
調整数_箱 = [_THISROW].[実在庫_箱] - [_THISROW].[理論在庫_箱]
調整数_個 = [_THISROW].[実在庫_個] - [_THISROW].[理論在庫_個]
```

## 代替案：監査ログ付き実装

### 変更履歴テーブルの追加

監査要件が厳しい場合は、別途「変更履歴」テーブルを作成：

**テーブル名**: `在庫変動変更履歴`

| カラム名 | データ型 | 説明 |
|---------|----------|------|
| 変更ID | UniqueID | 一意識別子 |
| 変更日時 | DateTime | 変更実行日時 |
| 対象テーブル | Text | 入荷記録/棚卸記録 |
| 対象ID | Text | 入荷ID/棚卸ID |
| 変更種別 | Enum | 編集/削除 |
| 変更前値 | Text | JSON形式で保存 |
| 変更後値 | Text | JSON形式で保存 |
| 変更者 | Email | 実行者 |

これにより完全な監査証跡を残しつつ、シンプルな実装が可能。

## 実装上の注意点

### 1. トランザクション処理

AppSheetはトランザクション処理が限定的なため：
- 各ステップでエラーチェック
- 必要に応じてロールバック処理を追加

### 2. 同時実行の考慮

複数ユーザーが同時に編集する可能性：
- タイムスタンプで競合検出
- 楽観的ロック機構の実装

### 3. パフォーマンス

大量データの場合：
- バッチ処理の検討
- インデックスの最適化

## テスト手順

1. **入荷編集テスト**
   - 入荷数量を変更 → 在庫が正しく更新されるか確認
   - 在庫変動履歴が置き換わっているか確認

2. **入荷削除テスト**
   - 入荷レコードを削除 → 在庫が減算されるか確認
   - 関連する在庫変動履歴が削除されているか確認

3. **棚卸編集テスト**
   - 実在庫数を変更 → 調整値が正しく反映されるか確認

4. **エラーケーステスト**
   - 在庫がマイナスになるケース
   - 同時編集のケース

## 4. 配布記録の編集・削除

### 配布記録編集時のボット

**Bot名**: `Handle_Distribution_Edit`

**イベント設定**:
```
Event Type: Data Change
Table: 配布記録
Event: UPDATES
Condition: TRUE
```

**プロセス設定**:

##### Step 1: 古い在庫変動履歴を削除
```
Action Type: Delete rows
Table: 在庫変動履歴
Rows to delete:
  FILTER(
    "在庫変動履歴",
    AND(
      [配布ID] = [_THISROW].[配布ID],
      [変動区分] = "配布"
    )
  )
```

##### Step 2: 新しい在庫変動履歴を作成
```
Action Type: Add a row
Table: 在庫変動履歴
Set these columns:
  - 変動日: [_THISROW].[配布日]
  - 商品ID: [_THISROW].[商品ID]
  - 変動区分: "配布"
  - 変動数量_個: -[_THISROW].[配布数量]  # マイナス値で記録
  - 施設ID: [_THISROW].[施設ID]
  - 患者ID: [_THISROW].[患者ID]
  - 配布ID: [_THISROW].[配布ID]
  - 変動理由: CONCATENATE("配布修正 - ", [_THISROW].[備考])
  - 記録者: USEREMAIL()
```

##### Step 3: 在庫数を更新（差分計算）
```
Action Type: Update row
Table: 在庫管理
Row to update:
  FILTER(
    "在庫管理",
    [商品ID] = [_THISROW].[商品ID]
  )
Set these columns:
  - 現在庫_個:
    [現在庫_個]
    + [_THISROW_BEFORE].[配布数量]  # 古い配布分を戻す
    - [_THISROW].[配布数量]          # 新しい配布分を引く
  - 最終更新日: NOW()
```

##### Step 4: 商品IDが変更された場合の処理
```
Condition:
  [_THISROW_BEFORE].[商品ID] <> [_THISROW].[商品ID]

Action:
  # 旧商品の在庫を戻す
  - 旧商品の在庫_個 += [_THISROW_BEFORE].[配布数量]
  # 新商品から引く
  - 新商品の在庫_個 -= [_THISROW].[配布数量]
```

### 配布記録削除時のボット

**Bot名**: `Handle_Distribution_Delete`

**イベント設定**:
```
Event Type: Data Change
Table: 配布記録
Event: DELETES
Condition: TRUE
```

**プロセス設定**:

##### Step 1: 在庫変動履歴を削除
```
Action Type: Delete rows
Table: 在庫変動履歴
Rows to delete:
  FILTER(
    "在庫変動履歴",
    [配布ID] = [_THISROW].[配布ID]
  )
```

##### Step 2: 在庫数を戻す（配布分を加算）
```
Action Type: Update row
Table: 在庫管理
Row to update:
  FILTER(
    "在庫管理",
    [商品ID] = [_THISROW].[商品ID]
  )
Set these columns:
  - 現在庫_個: [現在庫_個] + [_THISROW].[配布数量]  # 配布分を戻す
  - 最終更新日: NOW()
```

### 配布記録の特別な考慮事項

1. **在庫不足チェック**
   - 編集時に配布数量を増やす場合、在庫が足りるか確認
   - 不足時は警告通知を送信

2. **箱→個の自動補充との連携**
   - 配布削除で在庫が戻る際、箱数の調整は不要
   - 配布編集で数量増加時は自動補充を考慮

3. **施設・患者の変更**
   - 配送先が変更された場合の処理
   - 在庫変動履歴にも反映

## まとめ

入荷記録、棚卸記録、配布記録すべてで同じ「削除・再作成方式」を採用することで：
- 実装の一貫性が保たれる
- シンプルで保守しやすい
- 差分計算の複雑さを回避
- データの整合性を保つ

この方式により、どのテーブルの編集・削除でも確実に在庫管理と同期が取れます。