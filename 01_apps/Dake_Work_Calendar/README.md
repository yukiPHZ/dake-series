# Dake工程カレンダー

## アプリ概要

指定した開始日から終了日までの日付だけを並べる、期間可視化用のPDF作成アプリです。

月間カレンダーを作るアプリではありません。Excelに日付をコピーして紙面を整えていた作業を、A4縦PDFで置き換えるためのDAKEシリーズ用アプリです。

## 使い方

1. `main.py` を起動します。
2. 現場名、支店名、担当者名、電話番号を入力します。
3. 開始日と終了日を確認または入力します。
4. 保存先フォルダを選択します。
5. 「PDF作成」を押します。
6. 保存完了ダイアログのOK後、保存フォルダが開きます。

日付は `2025/11/29`、`2025-11-29`、`2025.11.29`、`20251129` の形式で入力できます。

## 入力項目

- 現場名
- 支店名
- 担当者名
- 電話番号
- 開始日
- 終了日
- 保存先

担当者名、電話番号、支店名、前回保存先は `dake_work_calendar_config.json` に保存され、次回起動時に再利用されます。

## 日付初期値

- 開始日：今日
- 終了日：今日から45日後

## PDF出力仕様

- 出力形式：PDF
- 用紙：A4縦
- 指定期間だけを表示
- 1日 = 1セル
- 日曜始まりの横7列で上から下へ配置
- 期間外セルは空欄で表示
- 前月・翌月の補完日は表示しません。
- 月全体を埋めません。
- 各セルに月/日、曜日、祝日名を表示します。
- 終了日のセルに「完工」を表示します。
- セル内のメモ罫線は表示しません。
- 印刷・スキャンで端が切れにくいよう、外側余白を安全側に取っています。
- 余白を残し、手書きしやすい紙面にしています。
- 日本の祝日表示
  - `holidays` ライブラリが利用できる場合は利用します。
  - 利用できない場合も、内蔵の簡易ルールで日本の祝日を表示します。

## ビルド方法

同じフォルダ内で以下を実行します。

```bat
build.bat
```

`dist` フォルダに `DakeWork_Calendar.exe` が作成されます。

## DAKE共通仕様レビュー結果

- UIフォント：BIZ UDPGothic を最優先、フォールバックは Yu Gothic UI / Meiryo
- ヘッダー：画面内でアプリ名を重複表示せず、「期間カレンダーを作る」＋短い説明に整理
- フッター：DAKE共通仕様の2ブロック構成に準拠し、狭幅時は中央寄せ2段構成へ切替
- UI_TEXT：ボタン、見出し、説明文、状態表示、エラー、完了、フッター、PDF表示文言を集約
- 応答性：PDF生成はワーカースレッドで実行し、処理中は進捗バーとドットループを表示
- スクロール：Canvas + Scrollable Frame にマウスホイール対応を追加
- アイコン：`../../02_assets/dake_icon.ico` を参照
- PDF調整：日曜始まり、タイトル簡略化、ページ表記削除、終了日「完工」表示、印刷安全余白を適用

## 注意事項

このアプリは期間を紙で見える化するためのものです。
日程・表示期間は、必ず利用者自身が確認してください。

初期実装では、工程の自動計算、ガントチャート化、業者別管理、案件データベース化、複数現場管理、クラウド保存、メール送信、印刷直接実行、Excel出力には対応していません。

## DAKE_META

```json
{
  "app_key": "dake_work_calendar",
  "display_name": "Dake工程カレンダー",
  "launcher_title": "工程カレンダー",
  "launcher_description": "指定期間の日付枠をA4縦PDFに並べます。",
  "site_title": "Dake工程カレンダー",
  "site_description": "開始日から終了日までの日付枠を並べ、工程確認用のA4縦PDFを作成できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_Work_Calendar",
  "exe_name": "DakeWork_Calendar.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Work_Calendar_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true,
  "demo_video_path": "release_artifacts/demo.mp4",
  "demo_video_url": "",
  "social_release_path": "release_artifacts/social_release.json"
}
```

## RELEASE_BODY

- 工程カレンダーPDF作成アプリ
- 指定期間の日付枠に対応
- A4縦PDF出力
- Windows向けexe
