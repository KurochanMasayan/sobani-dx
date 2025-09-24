# AppSheet アプリケーション管理者

## 管理者一覧

### 主管理者
- **名前**: 川本 達也
- **メールアドレス**: t.kawamoto@hks-sbn.jp
- **役割**: システム管理者
- **登録日**: 2025-01-22

## AppSheetでの管理者設定手順

1. **AppSheetエディタを開く**
   - app.appsheet.comにログイン
   - 該当アプリを選択

2. **Manage > Users and Auth に移動**
   - 左サイドメニューから「Manage」を選択
   - 「Users and Auth」セクションを開く

3. **User settings で管理者を追加**
   - 「User settings」タブを選択
   - 「Add users」をクリック
   - メールアドレス「t.kawamoto@hks-sbn.jp」を入力

4. **権限設定**
   - Role: Admin
   - Can edit app definition: Yes
   - Can view app data: Yes
   - Can edit app data: Yes

5. **保存して同期**
   - 「Save」をクリック
   - 「Deploy」で変更を反映

## 管理者権限の詳細

### Admin権限で可能な操作
- アプリ定義の編集
- 全データの閲覧・編集
- ユーザー管理
- セキュリティ設定の変更
- Automation/Workflowの設定
- データソース接続の管理

## セキュリティ設定推奨事項

1. **Authentication Provider**
   - Google認証を使用（推奨）
   - メールアドレスドメインの制限: @hks-sbn.jp

2. **データアクセス制限**
   - 機密データには追加のセキュリティフィルター設定
   - ユーザーロールベースのビュー制限

3. **監査ログ**
   - Admin操作の記録を有効化
   - 定期的な監査レビュー

## 連絡先

管理者への連絡:
- Email: t.kawamoto@hks-sbn.jp
- システムに関する問い合わせ、権限付与依頼等