# 会議室予約・出欠管理システム

Google スプレッドシートをデータベースとして使う、会議室予約システムと会議出欠管理システムの統合版です。  
読み込みはフロントからスプレッドシート CSV を直接取得し、書き込みだけを GAS Web API 経由で行います。

## 画面構成

| ファイル | 内容 |
|---|---|
| `index.html` | トップ画面 |
| `reserve.html` | 会議室予約画面 |
| `status.html` | 予約確認画面 |
| `attendance.html` | 出欠登録画面 |
| `admin.html` | 管理画面 |
| `config.js` | フロント共通設定 |

## GAS 構成

ユーザー要望に合わせて、GAS はほぼ 2 ファイルに寄せています。

| ファイル | 内容 |
|---|---|
| `gas/Setup.gs` | シート作成、初期データ投入、書式設定 |
| `gas/Api.gs` | Web API、予約処理、出欠処理、集計更新 |
| `gas/appsscript.json` | GAS マニフェスト |

## スプレッドシート構成

`setupSpreadsheet()` または `initializeSpreadsheet()` を実行すると、次のシートを自動作成します。

1. `予約`
2. `会議室`
3. `設定`
4. `メンバー一覧`
5. `会議一覧`
6. `出欠回答`
7. `会議別集計`
8. `個人別集計`
9. `部局別集計`
10. `連続欠席チェック`
11. `操作ログ`

## セットアップ

### 1. スプレッドシートと GAS を用意する

1. Google スプレッドシートを新規作成する
2. スプレッドシートから Apps Script を開く
3. `gas/Setup.gs` `gas/Api.gs` `gas/appsscript.json` を貼り付ける
4. `setupSpreadsheet()` を手動実行する
5. 必要なら `createRoomCalendarsAndUpdateRooms()` を実行して会議室ごとの Google カレンダーを作る

### 2. GAS を Web アプリとしてデプロイする

1. Apps Script で「デプロイ」→「新しいデプロイ」
2. 種類は「ウェブアプリ」
3. 実行者は自分
4. アクセスできるユーザーは全員
5. 発行された `exec` URL を控える

### 3. スプレッドシートを読み取り公開する

CSV 読み込みのため、スプレッドシートは「リンクを知っている全員が閲覧可」にしてください。

### 4. `config.js` を更新する

`config.js` にすべて集約してあります。必須は `SPREADSHEET_ID` と `GAS_API_URL` です。  
さらに高速化したい場合だけ `SHEET_GID` を埋めてください。`gid` が入っていない項目は自動でシート名読み込みにフォールバックします。

```js
window.APP_CONFIG = {
  SPREADSHEET_ID: 'スプレッドシートID',
  GAS_API_URL: 'GAS の Web アプリ URL',
  SHEET_GID: {
    reservations: '予約シートの gid',
    rooms: '会議室シートの gid',
    settings: '設定シートの gid',
    members: 'メンバー一覧シートの gid',
    meetings: '会議一覧シートの gid',
    attendanceResponses: '出欠回答シートの gid',
    meetingAggregations: '会議別集計シートの gid',
    memberAggregations: '個人別集計シートの gid',
    departmentAggregations: '部局別集計シートの gid',
    streaks: '連続欠席チェックシートの gid',
    logs: '操作ログシートの gid',
  },
};
```

他のフロントファイルを触る必要はありません。

### 5. GitHub Pages で公開する

1. このリポジトリを GitHub に push する
2. `Settings -> Pages` を開く
3. `main` ブランチの `/ (root)` を公開する

## 管理画面でできること

- 管理者パスワード認証
- 会議登録と更新
- 出欠一覧確認
- 個人別集計確認
- 部局別集計確認
- 連続欠席者確認
- 欠席理由一覧確認
- 設定更新
- 会議室予約の取消
- 集計手動再計算

## 注意

- 出欠管理と会議室予約は同居していますが、データ上は独立しています
- 読み込み高速化のため一覧系は CSV 直 fetch です
- 静的公開なので、管理者パスワードは強固な秘匿にはなりません
- 画面反映は `config.js` の変更だけで足ります
