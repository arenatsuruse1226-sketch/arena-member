# Google Apps Script 連携手順

1. `Code.gs` を Google Apps Script プロジェクトに貼り付ける。
2. `SPREADSHEET_ID` を反映先のスプレッドシートIDに置き換える。
3. `SHEET_NAME` を必要に応じて変更する。
4. Apps Script を「ウェブアプリ」としてデプロイする。
5. デプロイ後に表示される URL を `index.html` の `fetch` 先に設定する。
6. `Code.gs` の `LINE_CHANNEL_ACCESS_TOKEN` に LINE Messaging API のチャネルアクセストークンを設定する。
7. `Code.gs` の `LINE_TO_ID` に通知先のLINEユーザーIDまたはグループIDを設定する。

## 受信される主な列

- `登録日時`
- `姓` / `名`
- `フリガナ(姓)` / `フリガナ(名)`
- `郵便番号`
- `住所`
- `電話番号`
- `性別`
- `生年月日`
- `暗証番号`
- `DM`
- `台番号`
- `利用規約同意`

`familyName` や `postalCode` などの元のフォーム名も一緒に送信されるので、シート側の列設計に合わせて使えます。
