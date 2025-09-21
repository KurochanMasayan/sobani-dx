# 配布テンプレート複数選択アクション実装ガイド

## 概要
配布テンプレートから複数のレコードを選択し、一括で配布記録に複製するアクションの実装方法を説明します。

## 実装方式

### 方式1: ビューでの複数選択＋バルクアクション方式（推奨・最もシンプル）

別テーブルを作成せず、AppSheetの標準的な複数選択機能とアクションを組み合わせる方式です。

#### 1. ビューの設定

**ビュー名**: `配布テンプレート_選択`
**設定手順**:
1. テーブルビューまたはギャラリービューを作成
2. **View Options**で以下を設定：
   - `Selection mode`: ON にする
   - `Multi-select`: ON にする
   - `Show action bar`: ON にする
3. **Row Selected**の動作設定：
   - 選択時にチェックボックスを表示

#### 2. アクションの作成

##### Action 1: 単一配布記録作成
```
Action Name: Create_Distribution_From_Template
Action Type: Data: add a new row to another table
Target Table: 配布記録
Set these columns:
  - 配布日: TODAY()
  - 施設ID: [_THISROW].[施設ID]
  - 患者ID: [_THISROW].[患者ID]
  - 商品ID: [_THISROW].[商品ID]
  - 配布数量: [_THISROW].[現在月の配布数量]
  - 配布者: USEREMAIL()
  - テンプレートID: [_THISROW].[テンプレートID]
  - 備考: CONCATENATE("一括作成 ", TEXT(NOW()))
```

##### Action 2: 複数選択アクション（メイン）
```
Action Name: Bulk_Create_Distributions
Action Type: Data: execute an action on a set of rows
Referenced Table: 配布テンプレート
Referenced Rows:
  FILTER(
    "配布テンプレート",
    IN([テンプレートID], SPLIT(CONTEXT("SelectedRows"), ","))
  )
Referenced Action: Create_Distribution_From_Template
Prominence: Display prominently
Icon: ic_playlist_add
Display Name: "選択したテンプレートから配布作成"
Only if this condition is true:
  COUNT(CONTEXT("SelectedRows")) > 0
```

#### 3. アクションの表示制御

##### 選択状態でのみアクションを表示する設定

```
Display condition (Only if this condition is true):
  COUNT(CONTEXT("SelectedRows")) > 0
```

または、より詳細な制御：

```
Display condition:
  AND(
    COUNT(CONTEXT("SelectedRows")) > 0,
    CONTEXT("ViewName") = "配布テンプレート_選択"
  )
```

##### アクション無効化の代替案

表示ではなく、実行可能性を制御する場合：

```
Action condition (Only if this condition is true):
  COUNT(FILTER(
    "配布テンプレート",
    IN([テンプレートID], SPLIT(CONTEXT("SelectedRows"), ","))
  )) > 0
```

#### 4. 実装のポイント

- **CONTEXT("SelectedRows")**: ビューで選択された行のキー値リストを取得
- **Display prominently**: アクションバーに表示
- **Confirmation**: "X件の配布記録を作成します"のような確認メッセージを設定
- **選択状態チェック**: アクションは選択された行がある場合のみ表示/有効化

### 方式2: Automation機能を使った自動処理方式

テーブルビューで複数選択した後、Automationで自動的に配布記録を作成する方式です。

#### 1. カスタムアクションの作成

```
Action Name: Mark_Templates_For_Processing
Action Type: Data: set the values of some columns in this row
Set these columns:
  - 処理フラグ: TRUE
  - 処理日時: NOW()
Condition: [有効フラグ] = TRUE
```

#### 2. 複数選択用アクション

```
Action Name: Mark_Multiple_Templates
Action Type: Data: execute an action on a set of rows
Referenced Table: 配布テンプレート
Referenced Rows: LINKTOFILTEREDVIEW("配布テンプレート_選択")で選択された行
Referenced Action: Mark_Templates_For_Processing
```

#### 3. Automationボットの設定

**Bot名**: `Process_Marked_Templates`

**イベント設定**:
- Event Type: Data Change
- Table: 配布テンプレート
- Condition:
  ```
  AND(
    [処理フラグ] = TRUE,
    [_THISROW_BEFORE].[処理フラグ] = FALSE
  )
  ```

**プロセス設定**:
1. 配布記録を作成
2. 処理フラグをFALSEにリセット

### 方式3: 施設単位での一括作成（Automation活用）

#### 1. Automationボットの作成

**Bot名**: `Auto_Create_Distributions`

**イベント設定**:
- Event Type: Data Change
- Table: 一括配布制御
- Condition: [処理状態] = "準備中"

**プロセス設定**:
1. **Step 1**: ForEachループ
   - List: SPLIT([選択テンプレート], ",")
   - 変数名: template_id

2. **Step 2**: データ追加アクション
   - Target Table: 配布記録
   - 値の設定: 上記Action 1と同様

3. **Step 3**: ステータス更新
   - 処理状態 を "完了" に更新

## 推奨実装手順（方式1: ビューでの複数選択＋バルクアクション）

### ステップ1: ビューの準備
1. AppSheetエディタで「UX」タブを開く
2. 新規ビューを作成：
   - View name: `配布テンプレート_選択`
   - For this data: `配布テンプレート`
   - View type: `Table` または `Gallery`
3. View Optionsで設定：
   - Selection mode: ON
   - Multi-select: ON
   - Show action bar: ON

### ステップ2: アクション作成
1. 「Behavior」タブを開く
2. Actionsセクションで新規アクション作成
3. 単一配布記録作成アクションを設定
4. バルクアクション（複数選択用）を設定：
   - Display prominentlyをONにして、アクションバーに表示
   - Only if this condition is true: `COUNT(CONTEXT("SelectedRows")) > 0`
   - これにより選択時のみアクションが表示される

### ステップ3: テスト
1. プレビューモードでビューを開く
2. 複数のテンプレートをタップして選択
3. アクションバーの「選択したテンプレートから配布作成」をタップ
4. 配布記録テーブルで作成されたレコードを確認

### ステップ4: 最適化（オプション）
1. 条件付き表示の設定（施設や患者でフィルタ）
2. 確認ダイアログのカスタマイズ
3. 作成後の通知設定

## 使用方法（方式1: ビューでの複数選択）

### 操作手順
1. **配布テンプレート選択ビューを開く**
   - メニューから「配布テンプレート_選択」を選択

2. **テンプレートを複数選択**
   - 各行の左側のチェックボックスをタップ
   - または長押しして選択モードに入る
   - 必要なテンプレートをすべて選択

3. **アクションを実行**
   - 画面下部のアクションバーが表示される
   - 「選択したテンプレートから配布作成」ボタンをタップ

4. **確認ダイアログ**
   - 「X件の配布記録を作成しますか？」と表示
   - 「OK」をタップ

5. **完了**
   - 選択したテンプレート数分の配布記録が自動作成
   - 配布記録一覧に自動遷移

## 注意事項

### データ整合性
- テンプレートIDの参照整合性を保持
- 配布数量は現在月（奇数/偶数）の設定値を自動取得
- 重複チェック機能の実装を推奨

### パフォーマンス
- 一度に処理するレコード数は50件以下を推奨
- 大量処理の場合はAutomation機能で非同期処理

### エラーハンドリング
- 必須フィールドの検証
- 在庫不足チェック（オプション）
- トランザクション処理の考慮

## カスタマイズ例

### 施設単位での一括選択
```
Valid If式:
SELECT(
  配布テンプレート[テンプレートID],
  AND(
    [施設ID] = [_THISROW].[facility_id],
    [有効フラグ] = TRUE
  )
)
```

### 患者単位でのグループ化
```
Group By: [患者ID]
Display: CONCATENATE([患者名], " - ", COUNT([商品ID]), "品目")
```

### 配布頻度による自動選択
```
初期値式:
SELECT(
  配布テンプレート[テンプレートID],
  AND(
    [施設ID] = [_THISROW].[facility_id],
    MOD(DAY(TODAY()), [配布頻度_日]) = 0
  )
)
```

## トラブルシューティング

### 問題: テンプレートが選択できない
**解決方法**:
- Valid If条件を確認
- 参照テーブルの設定を確認
- データ型（Ref/EnumList）の設定を確認

### 問題: 配布記録が作成されない
**解決方法**:
- アクションの条件式を確認
- 必須フィールドの値設定を確認
- アクションの実行権限を確認

### 問題: パフォーマンスが遅い
**解決方法**:
- Virtual Columnの使用を最小限に
- Automation機能での非同期処理を検討
- インデックス設定の最適化

## まとめ

この実装により、配布テンプレートから効率的に複数の配布記録を作成できます。AppSheetの標準機能のみを使用しているため、追加コストなく実装可能です。