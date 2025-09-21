# AppSheetボットでADDS/UPDATESイベントを判別する方法

## 重要な事実確認

### _THISROW_BEFOREの実際の動作

**公式ドキュメントに基づく事実**：
- **ADDSイベント時**：`_THISROW_BEFORE`は**使用不可**（新規レコードには「以前の状態」が存在しない）
- **UPDATESイベント時**：`_THISROW_BEFORE`には更新前の値が入る
- **DELETESイベント時**：`_THISROW`に削除される値が入る

## 推奨実装方式

### 方式1: 別々のボットを作成（最も確実）

```yaml
Bot 1: Handle_New_Records
  Event:
    - Table: 入荷記録
    - Event Type: ADDS only
  Process:
    - Action: Create_Inventory_History
    - Action: Update_Stock_Add

Bot 2: Handle_Updates
  Event:
    - Table: 入荷記録
    - Event Type: UPDATES only
  Process:
    - Action: Delete_Old_History
    - Action: Create_New_History
    - Action: Update_Stock_Diff
```

**メリット**：
- 明確で理解しやすい
- エラーが起きにくい
- 各イベントに最適化した処理

**デメリット**：
- ボット数が増える（管理の複雑さ）

### 方式2: 1つのボットで両イベントを処理

```yaml
Bot: Handle_All_Changes
  Event:
    - Table: 入荷記録
    - Event Type: ADDS, UPDATES

  Process:
    Step 1: Branch on a condition
      # タイムスタンプや作成日時フィールドを使った判定
      Condition: ISBLANK([作成日時])

      TRUE Branch (新規):
        - Action: Set_Creation_Timestamp
        - Action: Create_Inventory_History
        - Action: Update_Stock_Add

      FALSE Branch (更新):
        - Action: Delete_Old_History
        - Action: Create_New_History
        - Action: Update_Stock_Diff
```

### 方式3: カスタムフラグを使用

テーブルに`処理タイプ`フィールドを追加：

```yaml
Process:
  Step 1: Set Processing Type
    - IF(ISBLANK([処理タイプ]), "新規", "更新")

  Step 2: Branch on Processing Type
    Condition: [処理タイプ] = "新規"

    TRUE Branch:
      - 新規追加処理
    FALSE Branch:
      - 更新処理
```

## 実装上の注意点

### 1. _THISROW_BEFOREは使えない場合がある

```
誤った実装例：
Condition: ISBLANK([_THISROW_BEFORE].[ID])  # ADDSイベントでエラーになる可能性

正しい実装例：
Condition: ISBLANK([作成日時])  # 独自のフィールドで判定
```

### 2. イベントタイプの組み合わせ制限

- ADDS + UPDATES：同一ボットで処理可能
- DELETES：必ず別ボットが必要（異なる処理フロー）

### 3. パフォーマンスへの影響

複数イベントを1つのボットで処理すると：
- 条件判定のオーバーヘッド
- 不要な分岐処理の実行可能性

## ベストプラクティス

### 推奨アプローチの選択基準

| ケース | 推奨方式 | 理由 |
|--------|----------|------|
| シンプルな処理 | 別々のボット | 明確で保守しやすい |
| 共通処理が多い | 1つのボット＋分岐 | コードの重複を避ける |
| 監査要件が厳しい | カスタムフラグ | 処理種別を明示的に記録 |

### 実装例：在庫管理システムの場合

```yaml
# 新規入荷専用ボット
Bot: New_Receipt_Handler
  Event: 入荷記録.ADDS
  Process:
    - Add row to 在庫変動履歴
    - Update 在庫管理 (加算)

# 入荷編集専用ボット
Bot: Edit_Receipt_Handler
  Event: 入荷記録.UPDATES
  Process:
    - Delete old 在庫変動履歴
    - Add new 在庫変動履歴
    - Update 在庫管理 (差分)

# 入荷削除専用ボット
Bot: Delete_Receipt_Handler
  Event: 入荷記録.DELETES
  Process:
    - Delete 在庫変動履歴
    - Update 在庫管理 (減算)
```

## まとめ

1. **_THISROW_BEFOREはADDSイベントでは使用不可**
2. **最も確実な方法は別々のボットを作成**
3. **1つのボットで処理する場合は、カスタムフィールドで判定**
4. **処理の複雑さと保守性のバランスを考慮して選択**

この実装により、各イベントタイプに応じた適切な処理が確実に実行されます。