# AppSheet セットアップガイド

## 1. Googleスプレッドシートの準備

### 1.1 新規スプレッドシート作成

1. Google Driveで新規スプレッドシートを作成
2. ファイル名: `歯科診療スケジュール管理`

### 1.2 シートの作成

以下の7つのシートを作成し、CSVデータをインポート:

| シート名 | CSVファイル |
|---------|------------|
| 施設マスタ | csv-data/施設マスタ.csv |
| 患者マスタ | csv-data/患者マスタ.csv |
| 居宅介護支援事業所マスタ | csv-data/居宅介護支援事業所マスタ.csv |
| スタッフマスタ | csv-data/スタッフマスタ.csv |
| 診療スケジュール | csv-data/診療スケジュール.csv |
| 患者不在時間 | csv-data/患者不在時間.csv |
| 日別訪問計画 | csv-data/日別訪問計画.csv |

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
| 作成日時 | DateTime | | Initial: NOW() / Editable: FALSE |

### 3.6 患者不在時間

| カラム | 型 | Key | 設定 |
|--------|-----|-----|------|
| 不在時間ID | Text | ✓ | Initial: CONCATENATE("ABS-", TEXT(RANDBETWEEN(1,999), "000")) |
| 患者ID | Ref | | → 患者マスタ |
| 曜日リスト | EnumList | | Values: 月, 火, 水, 木, 金, 土, 日 |
| 終日フラグ | Yes/No | | Default: FALSE |
| 開始時間 | Time | | Show_If: [終日フラグ] = FALSE |
| 終了時間 | Time | | Show_If: [終日フラグ] = FALSE |
| 不在理由 | Text | | |
| 備考 | LongText | | |

※ 曜日リストはEnumList型のため、AppSheet上ではチェックボックスで複数曜日を選択可能

#### バリデーションルール

**時間帯の必須チェック（開始時間・終了時間に設定）:**
```
IF(
  [終日フラグ] = FALSE,
  ISNOTBLANK([_THIS]),
  TRUE
)
```

**時間の前後関係チェック（終了時間に設定）:**
```
IF(
  [終日フラグ] = FALSE,
  [開始時間] < [終了時間],
  TRUE
)
```

**曜日リストの必須チェック（曜日リストに設定）:**
```
ISNOTBLANK([_THIS])
```

### 3.7 日別訪問計画

| カラム | 型 | Key | 設定 |
|--------|-----|-----|------|
| 計画ID | Text | ✓ | Initial: CONCATENATE("PLAN-", TEXT(TODAY(), "YYYYMMDD")) |
| 訪問日 | Date | | |
| ステータス | Enum | | Values: 作成済み, 印刷済み, 提出済み / Default: 作成済み |
| PDF作成日時 | DateTime | | Initial: NOW() / Editable: FALSE |
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

#### 不在時間一覧（参照用・改行区切り）
```
SUBSTITUTE(
  CONCATENATE(
    SELECT(
      患者不在時間[不在時間サマリ],
      [患者ID] = [_THISROW].[ID]
    )
  ),
  " , ",
  CHAR(10)
)
```
→ PDFテンプレートで改行表示されるよう、カンマ区切りを改行に変換

### 4.2 患者不在時間に追加

#### 患者名（表示用）
```
[患者ID].[氏名]
```

#### 時間帯表示
```
IF(
  [終日フラグ],
  "終日",
  TEXT([開始時間], "HH:MM") & " - " & TEXT([終了時間], "HH:MM")
)
```

#### 曜日短縮表示（連続曜日を範囲表示）

連続する曜日を「月-金」のように短縮表示する。長いパターンから順にSUBSTITUTEで置換。

```
SUBSTITUTE(
  SUBSTITUTE(
    SUBSTITUTE(
      SUBSTITUTE(
        SUBSTITUTE(
          SUBSTITUTE(
            SUBSTITUTE(
              SUBSTITUTE(
                SUBSTITUTE(
                  SUBSTITUTE(
                    SUBSTITUTE(
                      SUBSTITUTE(
                        SUBSTITUTE(
                          SUBSTITUTE(
                            SUBSTITUTE(
                              SUBSTITUTE(
                                [曜日リスト],
                                "月 , 火 , 水 , 木 , 金 , 土 , 日", "月-日"),
                              "月 , 火 , 水 , 木 , 金 , 土", "月-土"),
                            "火 , 水 , 木 , 金 , 土 , 日", "火-日"),
                          "月 , 火 , 水 , 木 , 金", "月-金"),
                        "火 , 水 , 木 , 金 , 土", "火-土"),
                      "水 , 木 , 金 , 土 , 日", "水-日"),
                    "月 , 火 , 水 , 木", "月-木"),
                  "火 , 水 , 木 , 金", "火-金"),
                "水 , 木 , 金 , 土", "水-土"),
              "木 , 金 , 土 , 日", "木-日"),
            "月 , 火 , 水", "月-水"),
          "火 , 水 , 木", "火-木"),
        "水 , 木 , 金", "水-金"),
      "木 , 金 , 土", "木-土"),
    "金 , 土 , 日", "金-日"),
  " , ", "・")
```

#### 不在時間サマリ（表示用）
```
"(" & [曜日短縮表示] & ") " & [時間帯表示]
```
→ 表示例: 「(月-金) 14:00 - 15:00」「(火・金) 終日」

### 4.3 診療スケジュールに追加

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

#### 訪問日の曜日
```
SWITCH(
  WEEKDAY([訪問日]),
  1, "日",
  2, "月",
  3, "火",
  4, "水",
  5, "木",
  6, "金",
  7, "土"
)
```

#### 訪問No（日別通し番号）
```
COUNT(
  SELECT(
    診療スケジュール[スケジュールID],
    AND(
      [訪問日] = [_THISROW].[訪問日],
      [ステータス] <> "キャンセル",
      OR(
        [訪問開始時間] < [_THISROW].[訪問開始時間],
        AND(
          [訪問開始時間] = [_THISROW].[訪問開始時間],
          [作成日時] < [_THISROW].[作成日時]
        )
      )
    )
  )
) + 1
```
→ 同じ訪問日の中で訪問開始時間順の通し番号。PDFテンプレートのNo.列に使用

#### 患者自己負担有無（PDF用）
```
[患者ID].[自己負担有無]
```

#### 患者要介護度（PDF用）
```
[患者ID].[要介護度]
```

#### 患者不在時間一覧（PDF用）
```
[患者ID].[不在時間一覧]
```

#### 不在時間チェック
```
IF(
  ISNOTBLANK(
    SELECT(
      患者不在時間[不在時間ID],
      AND(
        [患者ID] = [_THISROW].[患者ID],
        IN([_THISROW].[訪問日の曜日], [曜日リスト]),
        OR(
          [終日フラグ] = TRUE,
          AND(
            [開始時間] < [_THISROW].[訪問終了時間],
            [終了時間] > [_THISROW].[訪問開始時間]
          )
        )
      )
    )
  ),
  "⚠️ 不在時間と重複",
  ""
)
```
→ IN() で訪問日の曜日がEnumListに含まれるかチェック

### 4.4 日別訪問計画に追加

#### AM件数
```
COUNT(
  SELECT(
    診療スケジュール[スケジュールID],
    AND(
      [訪問日] = [_THISROW].[訪問日],
      [訪問開始時間] < "12:00:00",
      [ステータス] <> "キャンセル"
    )
  )
)
```

#### PM件数
```
COUNT(
  SELECT(
    診療スケジュール[スケジュールID],
    AND(
      [訪問日] = [_THISROW].[訪問日],
      [訪問開始時間] >= "12:00:00",
      [ステータス] <> "キャンセル"
    )
  )
)
```

#### 合計件数
```
[AM件数] + [PM件数]
```

#### スケジュール概要（一覧表示用）
```
TEXT([訪問日], "YYYY/MM/DD") & "（" &
SWITCH(
  WEEKDAY([訪問日]),
  1, "日", 2, "月", 3, "火", 4, "水", 5, "木", 6, "金", 7, "土"
) & "） " & [合計件数] & "件"
```
→ 表示例: 「2025/01/20（月） 3件」

#### AM訪問（List型・PDFテンプレート用）
```
SELECT(
  診療スケジュール[スケジュールID],
  AND(
    [訪問日] = [_THISROW].[訪問日],
    [訪問開始時間] < "12:00:00",
    [ステータス] <> "キャンセル"
  )
)
```
→ PDFテンプレートの `<<Start: ORDERBY([AM訪問], ...)>>` で使用

#### PM訪問（List型・PDFテンプレート用）
```
SELECT(
  診療スケジュール[スケジュールID],
  AND(
    [訪問日] = [_THISROW].[訪問日],
    [訪問開始時間] >= "12:00:00",
    [ステータス] <> "キャンセル"
  )
)
```
→ PDFテンプレートの `<<Start: ORDERBY([PM訪問], ...)>>` で使用

---

## 5. ビュー（View）の設定

### 5.1 本日のスケジュール

- Type: Table または Deck
- Data: 診療スケジュール
- Filter: [訪問日] = TODAY()
- Sort: [訪問開始時間] ASC
- Columns: 訪問開始時間, 患者名, 施設名, 診療内容, ステータス, 不在時間チェック

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

### 5.6 訪問計画一覧

- Type: Table
- Data: 日別訪問計画
- Sort: [訪問日] DESC
- Columns: 訪問日, スケジュール概要, AM件数, PM件数, 合計件数, ステータス

### 5.7 訪問計画カレンダー

- Type: Calendar
- Data: 日別訪問計画
- Start: [訪問日]
- End: [訪問日]
- Description: [スケジュール概要]

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

### 6.5 別の時間帯を追加（患者不在時間の複製）

- Action name: 別の時間帯を追加
- For a record of: 患者不在時間
- Do this: Data: add a new row to this table using values from this row
- Set columns:
  - 不在時間ID = CONCATENATE("ABS-", TEXT(RANDBETWEEN(1,999), "000"))
  - 患者ID = [患者ID]
  - 曜日リスト = [曜日リスト]
  - 終日フラグ = FALSE
  - 開始時間 = ""
  - 終了時間 = ""
  - 不在理由 = ""
  - 備考 = ""
- Prominence: Display prominently

→ 同じ患者・同じ曜日パターンで別の時間帯を追加。曜日リストがコピーされるので、時間帯だけ変更すればOK

### 6.6 タイムスケジュールPDF生成（レコード追加時に自動実行）

レコード追加時にPDFを自動生成するイベントアクション。

- Event: Adds only
- Table: 日別訪問計画
- Do this: App: Create a new file with a template
- Template: タイムスケジュールテンプレート（Google Docs）
- File name prefix: `CONCATENATE("タイムスケジュール_", TEXT([訪問日], "YYYYMMDD"))`
- File folder path: `"タイムスケジュールPDF/"`

※ テンプレートの作成手順は [03_PDFテンプレート設計書](../documents/03_PDFテンプレート設計書.md) を参照

### 6.7 印刷済みにする

- Action name: 印刷済みにする
- For a record of: 日別訪問計画
- Do this: Data: set the values of some columns
- Set: ステータス = "印刷済み"
- Only if: [ステータス] = "作成済み"
- Prominence: Display prominently

### 6.8 提出済みにする

- Action name: 提出済みにする
- For a record of: 日別訪問計画
- Do this: Data: set the values of some columns
- Set: ステータス = "提出済み"
- Only if: [ステータス] = "印刷済み"
- Prominence: Display prominently

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
   - 患者不在時間
7. 訪問計画（PDF生成用）
   - 訪問計画一覧
   - 訪問計画カレンダー
