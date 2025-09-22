# 施設テンプレート一括複製機能 実装ガイド

## 概要
施設を選択して、その施設に関連する全ての配布テンプレートを配布記録に一括で複製する機能の実装方法。

## 実装方法

### 1. 配布テンプレートテーブルに選択用カラムを追加

**テーブル: 配布テンプレート**
- カラム名: `選択フラグ`
- Type: Yes/No
- Initial value: FALSE
- Show: TRUE（施設別ビューでのみ表示）

### 2. 施設選択用のスライスを作成

**スライス名**: `施設別テンプレート`
- Source table: 配布テンプレート
- Row filter condition:
```
AND(
  [施設ID] = CONTEXT("SelectedFacility"),
  [有効] = TRUE
)
```

### 3. 施設選択ビューの作成

**ビュー名**: `施設選択`
- Type: Deck または Gallery
- For this data: 施設マスタ
- View Options:
  - Group by: なし
  - Sort by: 施設名

**表示設定**:
- Primary header: [施設名]
- Secondary header:
```
CONCATENATE(
  "テンプレート数: ",
  COUNT(SELECT(配布テンプレート[テンプレートID],
    [施設ID] = [_THISROW].[施設ID]))
)
```

### 4. 施設別テンプレート一覧ビューの作成

**ビュー名**: `施設別テンプレート一覧`
- Type: Table
- For this data: 施設別テンプレート（スライス）
- Allow QuickEdit: TRUE

**カラム設定**:
- 選択フラグ: QuickEdit有効
- テンプレートID: 表示のみ
- 患者ID: 表示のみ
- 商品ID: 表示のみ
- 奇数月数量: 表示
- 偶数月数量: 表示
- 配布頻度_日: 表示

**グループ化**:
- Group by: 患者ID
- Group aggregate: なし

### 5. メインアクション：施設選択から複製まで

#### 5.1 施設選択アクション
**Action name**: `施設を選択して複製`
- For a record of this table: 施設マスタ
- Do this: App: go to another view within this app
- Target:
```
LINKTOROW([施設ID], "施設別テンプレート一覧")
```
- Prominence: Display prominently
- Icon: copy_all

#### 5.2 全選択/選択解除アクション

**全選択アクション**
- Action name: `全テンプレート選択`
- For a record of this table: 配布テンプレート
- Do this: Data: execute an action on a set of rows
- Referenced Table: 配布テンプレート
- Referenced Rows:
```
SELECT(配布テンプレート[テンプレートID],
  [施設ID] = CONTEXT("SelectedFacility"))
```
- Referenced Action: `選択フラグON`（個別アクション）

**個別選択アクション**
- Action name: `選択フラグON`
- For a record of this table: 配布テンプレート
- Do this: Data: set the values of some columns in this row
- Set these columns:
  - 選択フラグ = TRUE

#### 5.3 配布記録への複製アクション

**複製実行アクション**
- Action name: `選択テンプレートを配布記録に複製`
- For a record of this table: 配布テンプレート
- Do this: Data: add a new row to another table using values from this row
- Table to add to: 配布記録
- Set these columns:
```
配布ID = CONCATENATE(
  "DIST-",
  TEXT(NOW(), "YYYYMMDDHHmmss"),
  RIGHT(CONCATENATE("00", RANDBETWEEN(1, 99)), 2)
)

配布日 = TODAY()

施設ID = [施設ID]

患者ID = [患者ID]

商品ID = [商品ID]

配布数量_箱 = IF(
  OR([奇数月数量] >= 10, [偶数月数量] >= 10),
  FLOOR(MAX([奇数月数量], [偶数月数量]) / 10),
  0
)

配布数量_個 = IF(
  MOD(MONTH(TODAY()), 2) = 1,
  [奇数月数量],
  [偶数月数量]
)

配布タイプ = IF(
  ISNOTBLANK([患者ID]),
  "個別配布",
  "施設配布"
)

状態 = "準備中"

配布者 = USEREMAIL()

備考 = CONCATENATE(
  TEXT(TODAY(), "YYYY/MM"),
  " 定期配布（テンプレート一括複製）"
)
```

**バッチ複製アクション（グループ）**
- Action name: `選択済みテンプレート一括複製`
- For a record of this table: 配布テンプレート
- Do this: Grouped: execute a sequence of actions
- Actions:
  1. Data: execute an action on a set of rows
     - Referenced Rows:
     ```
     SELECT(配布テンプレート[テンプレートID],
       AND(
         [施設ID] = CONTEXT("SelectedFacility"),
         [選択フラグ] = TRUE
       ))
     ```
     - Referenced Action: `選択テンプレートを配布記録に複製`

  2. 選択フラグリセット（全てFALSEに）

  3. 完了通知表示

### 6. ボタン配置とUI

#### 施設選択ビューのアクション配置
- `施設を選択して複製`ボタンを各施設カードに表示

#### 施設別テンプレート一覧ビューのアクション配置

**ヘッダーアクション**:
1. `全選択`ボタン
2. `選択解除`ボタン
3. `選択済みを複製実行`ボタン（確認ダイアログ付き）
4. `施設選択に戻る`ボタン

### 7. 確認ダイアログの実装

**複製前の確認**
- Condition for confirmation:
```
COUNT(SELECT(配布テンプレート[テンプレートID],
  AND(
    [施設ID] = CONTEXT("SelectedFacility"),
    [選択フラグ] = TRUE
  ))) > 0
```

- Confirmation message:
```
CONCATENATE(
  "選択された ",
  COUNT(SELECT(配布テンプレート[テンプレートID],
    AND(
      [施設ID] = CONTEXT("SelectedFacility"),
      [選択フラグ] = TRUE
    ))),
  " 件のテンプレートを配布記録に複製します。",
  CHAR(10),
  "よろしいですか？"
)
```

### 8. 重複チェックと検証

**Valid If（配布記録テーブル）**:
```
NOT(IN(
  CONCATENATE([配布日], [施設ID], [患者ID], [商品ID]),
  SELECT(
    配布記録[CONCATENATE([配布日], [施設ID], [患者ID], [商品ID])],
    [配布ID] <> [_THISROW].[配布ID]
  )
))
```

**Error message**:
```
"同じ日付・施設・患者・商品の組み合わせが既に存在します"
```

### 9. 完了後の処理

**複製完了後のアクション**:
1. 選択フラグを全てFALSEにリセット
2. 成功メッセージを表示
3. 配布記録一覧ビューに遷移（フィルタ付き）

**遷移先のフィルタ**:
```
AND(
  [配布日] = TODAY(),
  [配布者] = USEREMAIL(),
  CONTAINS([備考], "テンプレート一括複製")
)
```

### 10. 権限制御

**アクション表示条件**:
```
AND(
  USERSETTINGS("権限") IN LIST("管理者", "看護師"),
  OR(
    USERSETTINGS("権限") = "管理者",
    [施設ID] IN USERSETTINGS("担当施設")
  )
)
```

## 実装手順

1. **Step 1**: 配布テンプレートテーブルに`選択フラグ`カラムを追加
2. **Step 2**: 施設別テンプレートスライスを作成
3. **Step 3**: 施設選択ビューを作成
4. **Step 4**: 施設別テンプレート一覧ビューを作成
5. **Step 5**: 個別アクション（選択、複製）を作成
6. **Step 6**: グループアクション（一括処理）を作成
7. **Step 7**: ビューにアクションボタンを配置
8. **Step 8**: 確認ダイアログと検証ルールを設定
9. **Step 9**: 権限制御を実装
10. **Step 10**: テスト実行と調整

## 注意事項

- 大量のテンプレートを一度に複製する場合、パフォーマンスに影響する可能性があります
- 配布IDの一意性を保証するため、タイムスタンプとランダム値を組み合わせています
- 月の奇数・偶数に応じて配布数量を自動判定します
- 患者IDがない場合は「施設配布」として扱います

## テスト項目

- [ ] 施設選択が正しく動作するか
- [ ] テンプレートの選択/選択解除が正しく動作するか
- [ ] 選択したテンプレートのみが複製されるか
- [ ] 配布記録のIDが重複しないか
- [ ] 奇数月/偶数月の数量判定が正しいか
- [ ] 権限に応じたアクション表示制御が機能するか
- [ ] 複製後の選択フラグリセットが動作するか
- [ ] 完了後のビュー遷移とフィルタが正しく動作するか