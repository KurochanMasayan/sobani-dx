# AppSheet セットアップガイド

## 1. Googleスプレッドシートの準備

### 1.1 新規スプレッドシート作成

1. Google Driveで新規スプレッドシートを作成
2. ファイル名: `歯科診療スケジュール管理`

### 1.2 シートの作成

以下の5つのシートを作成し、CSVデータをインポート:

| シート名 | CSVファイル |
|---------|------------|
| 施設マスタ | csv-data/施設マスタ.csv |
| 患者マスタ | csv-data/患者マスタ.csv |
| 居宅介護支援事業所マスタ | csv-data/居宅介護支援事業所マスタ.csv |
| スタッフマスタ | csv-data/スタッフマスタ.csv |
| 診療スケジュール | csv-data/診療スケジュール.csv |

---

## 2. AppSheetアプリの作成

### 2.1 AppSheetにアクセス

1. https://www.appsheet.com/ にアクセス
2. Googleアカウントでログイン
3. 「Create」→「App」→「Start with existing data」

### 2.2 データソース接続

1. 「Google Sheets」を選択
2. 作成したスプレッドシートを選択
3. 全てのシートを追加

---

## 3. テーブル設定

### 3.1 施設マスタ

| カラム | 型 | Key | 設定 |
|--------|-----|-----|------|
| 施設ID | Text | ✓ | |
| 施設名 | Text | | |
| 郵便番号 | Text | | |
| 住所 | Text | | |
| 電話番号 | Phone | | |
| FAX番号 | Phone | | |
| 管理者/担当者 | Text | | |
| 備考 | LongText | | |

### 3.2 患者マスタ

| カラム | 型 | Key | 設定 |
|--------|-----|-----|------|
| ID | Number | ✓ | |
| 氏名 | Text | | |
| カナ | Text | | |
| 生年月日 | Date | | |
| 性別 | Enum | | Values: 男, 女 |
| 施設ID | Ref | | → 施設マスタ |
| 部屋番号 | Text | | |
| 保険種別 | Enum | | Values: 生活保護, 後期高齢, 国保 |
| 負担割合 | Enum | | Values: 0割, 1割, 2割 |
| 重度医療 | Text | | |
| 要介護度 | Enum | | Values: 非該当, 要介護1, 要介護2, 要介護3, 要介護4, 要介護5 |
| 介護負担割合 | Enum | | Values: 0割, 1割 |
| 自己負担有無 | Text | | |
| 訪問頻度 | Enum | | Values: 月1, 月2 |
| 曜日 | Text | | |
| 目的 | Text | | |
| 事業所ID | Ref | | → 居宅介護支援事業所マスタ |
| 備考 | LongText | | |

### 3.3 居宅介護支援事業所マスタ

| カラム | 型 | Key | 設定 |
|--------|-----|-----|------|
| 事業所ID | Text | ✓ | |
| 事業所名 | Text | | |
| 電話番号 | Phone | | |
| 担当者 | Text | | |
| 備考 | LongText | | |

### 3.4 スタッフマスタ

| カラム | 型 | Key | 設定 |
|--------|-----|-----|------|
| スタッフID | Text | ✓ | |
| スタッフ名 | Text | | |
| 役職 | Enum | | Values: 歯科医師, 歯科衛生士 |
| 電話番号 | Phone | | |
| 稼働曜日 | Text | | |
| 備考 | LongText | | |

### 3.5 診療スケジュール

| カラム | 型 | Key | 設定 |
|--------|-----|-----|------|
| スケジュールID | Text | ✓ | Initial: CONCATENATE("SCH-", TEXT(TODAY(), "YYYYMMDD"), "-", TEXT(RANDBETWEEN(1,999), "000")) |
| 訪問日 | Date | | |
| 訪問開始時間 | Time | | |
| 訪問終了時間 | Time | | |
| 患者ID | Ref | | → 患者マスタ |
| 施設ID | Ref | | → 施設マスタ |
| 担当医師ID | Ref | | → スタッフマスタ (Valid If: [役職]="歯科医師") |
| 担当衛生士ID | Ref | | → スタッフマスタ (Valid If: [役職]="歯科衛生士") |
| 診療内容 | Text | | |
| ステータス | Enum | | Values: 予定, 完了, キャンセル / Default: 予定 |
| 医師算定 | Yes/No | | Default: FALSE |
| 衛生士算定 | Yes/No | | Default: FALSE |
| 衛生士滞在時間 | Number | | |
| 診療メモ | LongText | | |
| 備考 | LongText | | |

---

## 4. Virtual Columns（仮想列）の設定

### 4.1 患者マスタに追加

#### 当月訪問回数
```
COUNT(
  SELECT(
    診療スケジュール[スケジュールID],
    AND(
      [患者ID] = [_THISROW].[ID],
      MONTH([訪問日]) = MONTH(TODAY()),
      YEAR([訪問日]) = YEAR(TODAY()),
      [ステータス] <> "キャンセル"
    )
  )
)
```

#### 当月医師算定回数
```
COUNT(
  SELECT(
    診療スケジュール[スケジュールID],
    AND(
      [患者ID] = [_THISROW].[ID],
      [医師算定] = TRUE,
      MONTH([訪問日]) = MONTH(TODAY()),
      YEAR([訪問日]) = YEAR(TODAY())
    )
  )
)
```

#### 当月衛生士算定回数
```
COUNT(
  SELECT(
    診療スケジュール[スケジュールID],
    AND(
      [患者ID] = [_THISROW].[ID],
      [衛生士算定] = TRUE,
      MONTH([訪問日]) = MONTH(TODAY()),
      YEAR([訪問日]) = YEAR(TODAY())
    )
  )
)
```

#### 医師算定残り
```
MAX(0, 2 - [当月医師算定回数])
```

#### 衛生士算定残り
```
MAX(0, 4 - [当月衛生士算定回数])
```

#### 算定ステータス
```
IF(
  AND([医師算定残り] = 0, [衛生士算定残り] = 0),
  "✓ 算定完了",
  CONCATENATE(
    IF([医師算定残り] > 0, "医師残" & [医師算定残り] & "回 ", ""),
    IF([衛生士算定残り] > 0, "衛生士残" & [衛生士算定残り] & "回", "")
  )
)
```

### 4.2 診療スケジュールに追加

#### 患者名（表示用）
```
[患者ID].[氏名]
```

#### 施設名（表示用）
```
[施設ID].[施設名]
```

#### 施設住所（表示用）
```
[施設ID].[住所]
```

#### 医師名（表示用）
```
[担当医師ID].[スタッフ名]
```

#### 衛生士名（表示用）
```
[担当衛生士ID].[スタッフ名]
```

#### 20分警告
```
IF(
  AND(
    ISNOTBLANK([担当衛生士ID]),
    [衛生士滞在時間] < 20,
    [衛生士滞在時間] > 0
  ),
  "⚠️ 20分未満",
  ""
)
```

#### 患者の医師算定残り
```
[患者ID].[医師算定残り]
```

#### 患者の衛生士算定残り
```
[患者ID].[衛生士算定残り]
```

---

## 5. ビュー（View）の設定

### 5.1 本日のスケジュール

- Type: Table または Deck
- Data: 診療スケジュール
- Filter: [訪問日] = TODAY()
- Sort: [訪問開始時間] ASC
- Columns: 訪問開始時間, 患者名, 施設名, 診療内容, ステータス

### 5.2 週間スケジュール

- Type: Calendar
- Data: 診療スケジュール
- Start: [訪問日]
- End: [訪問日]
- Description: [患者名] & " - " & [施設名]

### 5.3 患者一覧

- Type: Table
- Data: 患者マスタ
- Sort: [カナ] ASC
- Columns: 氏名, 施設ID, 要介護度, 訪問頻度, 算定ステータス

### 5.4 算定アラート

- Type: Deck
- Data: 患者マスタ
- Filter: OR([医師算定残り] > 0, [衛生士算定残り] > 0)
- Sort: [氏名] ASC
- Primary: [氏名]
- Secondary: [算定ステータス]

### 5.5 施設別スケジュール

- Type: Table
- Data: 診療スケジュール
- Group by: [施設ID]
- Sort: [訪問日], [訪問開始時間]

---

## 6. アクション（Action）の設定

### 6.1 医師算定ON

- Action name: 医師算定をONにする
- For a record of: 診療スケジュール
- Do this: Data: set the values of some columns
- Set: 医師算定 = TRUE
- Only if: AND([医師算定] = FALSE, [患者の医師算定残り] > 0)

### 6.2 衛生士算定ON

- Action name: 衛生士算定をONにする
- For a record of: 診療スケジュール
- Do this: Data: set the values of some columns
- Set: 衛生士算定 = TRUE
- Only if: AND([衛生士算定] = FALSE, [患者の衛生士算定残り] > 0, [衛生士滞在時間] >= 20)

### 6.3 完了にする

- Action name: 完了にする
- For a record of: 診療スケジュール
- Do this: Data: set the values of some columns
- Set: ステータス = "完了"
- Only if: [ステータス] = "予定"

### 6.4 キャンセルにする

- Action name: キャンセルにする
- For a record of: 診療スケジュール
- Do this: Data: set the values of some columns
- Set: ステータス = "キャンセル"
- Only if: [ステータス] = "予定"

---

## 7. ナビゲーション設定

メニュー構成:
1. 本日のスケジュール（Primary）
2. 週間カレンダー
3. 患者一覧
4. 算定アラート
5. 施設別スケジュール
6. マスタ管理（サブメニュー）
   - 施設マスタ
   - スタッフマスタ
   - 事業所マスタ
