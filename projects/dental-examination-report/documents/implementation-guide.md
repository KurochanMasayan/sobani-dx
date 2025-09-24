# 歯科検診報告書 実装ガイド

## 1. セットアップ手順

### 1.1 新規スプレッドシート作成

1. **Google Driveを開く**
   - drive.google.comにアクセス
   - 新規 > Google スプレッドシート

2. **シート名を設定**
   - 「歯科検診報告書_テンプレート」

3. **シートの基本設定**
   - シート1を「検診記録」に名前変更
   - グリッド線を表示

### 1.2 レイアウト作成

#### Step 1: ヘッダー部分
```
A1: =「歯科検診報告書」（タイトル、フォント20pt、太字）
A2: 結合してセンタリング（A2:L2）

B3: 患者名:
C3: [入力欄]
E3: 年齢:
F3: [入力欄]
H3: 検診日:
I3: [入力欄]
K3: 担当医:
L3: [入力欄]
```

#### Step 2: 凡例作成
```
B5: === 凡例 ===
C5: ○健全
D5: C虫歯
E5: ×欠損
F5: ✓治療済
G5: CO要観察
```

### 1.3 歯科チャート配置

#### 上顎の設定（行8-12）
```
D7: 【上顎】
E8: 18  F8: 17  G8: 16  H8: 15  I8: 14  J8: 13  K8: 12  L8: 11
E9: ○   F9: ○   G9: ○   H9: ○   I9: ○   J9: ○   K9: ○   L9: ○

（1行空白）

E11: ○   F11: ○   G11: ○   H11: ○   I11: ○   J11: ○   K11: ○   L11: ○
E12: 21  F12: 22  G12: 23  H12: 24  I12: 25  J12: 26  K12: 27  L12: 28
```

#### 下顎の設定（行15-19）
```
D14: 【下顎】
E15: 48  F15: 47  G15: 46  H15: 45  I15: 44  J15: 43  K15: 42  L15: 41
E16: ○   F16: ○   G16: ○   H16: ○   I16: ○   J16: ○   K16: ○   L16: ○

（1行空白）

E18: ○   F18: ○   G18: ○   H18: ○   I18: ○   J18: ○   K18: ○   L18: ○
E19: 31  F19: 32  G19: 33  H19: 34  I19: 35  J19: 36  K19: 37  L19: 38
```

## 2. 書式設定

### 2.1 セルサイズ調整

```javascript
// 手動設定
1. 列E～L を選択
2. 右クリック > 列のサイズを変更
3. 幅を 50 ピクセルに設定

4. 行8,9,11,12,15,16,18,19 を選択
5. 右クリック > 行のサイズを変更
6. 高さを 40 ピクセルに設定
```

### 2.2 罫線設定

1. **歯科チャート全体を選択**（E8:L19）
2. **書式 > 罫線**
3. **外枠**: 太線
4. **内部**: 格子（細線）

### 2.3 セルの配置

1. **歯番号セル**（E8:L8, E12:L12, E15:L15, E19:L19）
   - 中央揃え
   - フォントサイズ: 10pt
   - 背景色: #F0F0F0

2. **状態表示セル**（E9:L9, E11:L11, E16:L16, E18:L18）
   - 中央揃え
   - フォントサイズ: 16pt
   - 太字

## 3. データ入力規則の設定

### 3.1 ドロップダウンリスト作成

1. **状態入力セルを選択**
   - E9:L9, E11:L11, E16:L16, E18:L18

2. **データ > データの入力規則**

3. **条件設定**
   - 条件: リストから選択
   - リスト: ○,C,×,✓,CO
   - 無効なデータの場合: 入力を拒否

4. **外観**
   - セルにドロップダウンリストを表示: チェック

### 3.2 患者情報の入力規則

```
検診日（I3）:
- 条件: 日付
- 無効なデータ: 警告を表示

年齢（F3）:
- 条件: 数値
- 1 ～ 120 の間
```

## 4. 条件付き書式の設定

### 4.1 状態別の色分け

1. **範囲を選択**: E9:L9, E11:L11, E16:L16, E18:L18
2. **書式 > 条件付き書式**

#### 健全歯（○）
```
条件: テキストが次と完全一致
値: ○
書式:
- 背景色: #FFFFFF
- 文字色: #000000
```

#### 虫歯（C）
```
条件: テキストが次と完全一致
値: C
書式:
- 背景色: #FF4444
- 文字色: #FFFFFF
- 太字: ON
```

#### 欠損（×）
```
条件: テキストが次と完全一致
値: ×
書式:
- 背景色: #333333
- 文字色: #FFFFFF
```

#### 治療済（✓）
```
条件: テキストが次と完全一致
値: ✓
書式:
- 背景色: #4444FF
- 文字色: #FFFFFF
```

#### 要観察（CO）
```
条件: テキストが次と完全一致
値: CO
書式:
- 背景色: #FFCC00
- 文字色: #000000
- 斜体: ON
```

## 5. 統計機能の実装

### 5.1 統計表示エリア作成

```
B22: === 検診結果統計 ===
B23: 健全歯:
C23: =COUNTIF({E9:L9,E11:L11,E16:L16,E18:L18},"○")
B24: 虫歯:
C24: =COUNTIF({E9:L9,E11:L11,E16:L16,E18:L18},"C")
B25: 欠損:
C25: =COUNTIF({E9:L9,E11:L11,E16:L16,E18:L18},"×")
B26: 治療済:
C26: =COUNTIF({E9:L9,E11:L11,E16:L16,E18:L18},"✓")
B27: 要観察:
C27: =COUNTIF({E9:L9,E11:L11,E16:L16,E18:L18},"CO")
```

### 5.2 パーセンテージ表示

```
D23: =C23/32*100 & "%"
D24: =C24/32*100 & "%"
D25: =C25/32*100 & "%"
D26: =C26/32*100 & "%"
D27: =C27/32*100 & "%"
```

## 6. Google Apps Script (GAS) 機能追加

### 6.1 スクリプトエディタを開く

1. **拡張機能 > Apps Script**
2. 新規プロジェクト作成

### 6.2 基本機能の実装

#### 初期化関数
```javascript
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('歯科検診ツール')
    .addItem('全て健全に設定', 'setAllHealthy')
    .addItem('選択範囲をクリア', 'clearSelection')
    .addItem('データを保存', 'saveData')
    .addSeparator()
    .addItem('印刷用レイアウト', 'setPrintLayout')
    .addToUi();
}
```

#### 全て健全に設定
```javascript
function setAllHealthy() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ranges = [
    sheet.getRange('E9:L9'),
    sheet.getRange('E11:L11'),
    sheet.getRange('E16:L16'),
    sheet.getRange('E18:L18')
  ];

  ranges.forEach(function(range) {
    range.setValue('○');
  });

  SpreadsheetApp.getUi().alert('全ての歯を健全に設定しました。');
}
```

#### 選択範囲クリア
```javascript
function clearSelection() {
  var range = SpreadsheetApp.getActiveRange();
  range.clearContent();
  SpreadsheetApp.getUi().alert('選択範囲をクリアしました。');
}
```

#### データ保存
```javascript
function saveData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('保存データ');

  if (!dataSheet) {
    dataSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('保存データ');
  }

  // 患者情報取得
  var patientName = sheet.getRange('C3').getValue();
  var age = sheet.getRange('F3').getValue();
  var date = sheet.getRange('I3').getValue();
  var doctor = sheet.getRange('L3').getValue();

  // 歯の状態取得
  var upperRight = sheet.getRange('E9:L9').getValues()[0];
  var upperLeft = sheet.getRange('E11:L11').getValues()[0];
  var lowerRight = sheet.getRange('E16:L16').getValues()[0];
  var lowerLeft = sheet.getRange('E18:L18').getValues()[0];

  // タイムスタンプ付きで保存
  var timestamp = new Date();
  var rowData = [timestamp, patientName, age, date, doctor]
    .concat(upperRight).concat(upperLeft)
    .concat(lowerRight).concat(lowerLeft);

  dataSheet.appendRow(rowData);

  SpreadsheetApp.getUi().alert('データを保存しました。');
}
```

### 6.3 高度な機能

#### キーボードショートカット
```javascript
function onEdit(e) {
  var range = e.range;
  var value = e.value;

  // 数字キーで状態変更
  var statusMap = {
    '1': '○',
    '2': 'C',
    '3': '×',
    '4': '✓',
    '5': 'CO'
  };

  if (statusMap[value]) {
    range.setValue(statusMap[value]);
  }
}
```

## 7. 使用方法

### 7.1 基本操作

1. **患者情報入力**
   - 各項目に必要な情報を入力

2. **歯の状態入力**
   - セルをクリックしてドロップダウンから選択
   - または直接入力（○, C, ×, ✓, CO）

3. **一括操作**
   - 複数セルを選択
   - データ > データの入力規則から一括設定

### 7.2 効率的な入力方法

#### ドラッグ操作
1. 最初のセルに状態を入力
2. セルの右下をドラッグして複製

#### コピー＆ペースト
1. 状態を入力したセルをコピー（Ctrl+C）
2. 範囲を選択してペースト（Ctrl+V）

#### ショートカットキー
- 数字の1～5キーで素早く状態変更（GAS設定後）

## 8. トラブルシューティング

### 8.1 よくある問題と解決方法

#### ドロップダウンが表示されない
- データ > データの入力規則を再設定
- ブラウザのキャッシュをクリア

#### 条件付き書式が適用されない
- 書式ルールの優先順位を確認
- 範囲指定が正しいか確認

#### GASが動作しない
- 承認が必要な場合は許可する
- スクリプトエディタでエラーを確認

### 8.2 パフォーマンス改善

- 不要な計算式を削除
- 画像やグラフを最小限に
- 定期的にシートを最適化

## 9. カスタマイズ例

### 9.1 施設ロゴ追加
```
A1にロゴ画像を挿入
挿入 > 画像 > セル内の画像
```

### 9.2 印刷レイアウト
```javascript
function setPrintLayout() {
  var sheet = SpreadsheetApp.getActiveSheet();

  // 印刷範囲設定
  sheet.getRange('A1:L30').activate();

  // ページ設定
  // ファイル > 印刷 で手動調整
}
```

### 9.3 メール送信機能
```javascript
function sendReport() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var patientName = sheet.getRange('C3').getValue();

  // PDFとして出力
  var pdf = SpreadsheetApp.getActiveSpreadsheet().getAs('application/pdf');

  // メール送信
  MailApp.sendEmail({
    to: 'recipient@example.com',
    subject: '歯科検診報告書 - ' + patientName,
    body: '検診報告書を添付します。',
    attachments: [pdf]
  });
}
```

## 10. 今後の展開

- モバイルアプリ対応
- 画像添付機能
- AIによる診断支援
- 他システムとの連携