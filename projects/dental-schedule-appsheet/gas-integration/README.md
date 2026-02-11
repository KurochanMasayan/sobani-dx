# Google Apps Script - AppSheet Call a script 連携

## セットアップ

### 1. GASプロジェクト作成

1. スプレッドシートで「拡張機能」→「Apps Script」
2. `Code.gs` の内容を貼り付け
3. 保存

### 2. Apps Script API を有効化

1. GASエディタで「プロジェクトの設定」（歯車アイコン）
2. 「Google Cloud Platform (GCP) プロジェクト」でプロジェクト番号を確認
3. [Google Cloud Console](https://console.cloud.google.com/) でApps Script APIを有効化

### 3. AppSheet設定

1. Automation → Bot作成
2. Event: トリガー設定
3. Process → Step追加 → 「Call a script」
4. 設定:
   - Apps Script Project: スプレッドシートに紐づくプロジェクトを選択
   - Function Name: 呼び出す関数名（例: `testFunction`）
   - Function Parameters: 渡すパラメータ

## 利用可能な関数

| 関数名 | 説明 |
|--------|------|
| `testFunction` | 接続テスト |
| `onScheduleUpdate` | スケジュール更新時 |
| `onBillingUpdate` | 算定フラグ更新時 |

## AppSheet設定例

```
Function Name: testFunction
Function Parameters:
  - scheduleId: <<[スケジュールID]>>
  - patientId: <<[患者ID]>>
```
