# Dake確認印

PDFに名字と日付入りの確認印を1か所だけ押して、別名PDFとして保存するWindowsデスクトップアプリです。

## アプリ概要

- PDFを読み込み、プレビュー上で押印位置をクリックします。
- 確認印は赤い丸枠、縦書き明朝体の名字、円周下部内側の日付で構成します。
- 日付入力は `YYYYMMDD` のまま、印影では `YYYY.MM.DD` 形式で表示します。
- 保存ファイル名は `元ファイル名_確認印.pdf` を初期値にし、元PDFと同じフォルダへ保存します。
- 元PDFは上書きしません。

## 使い方

1. `PDFを選ぶ` からPDFを選択します。
2. PDFプレビュー上で確認印を押したい場所をクリックします。
3. 必要に応じて名字、日付、印影サイズを調整します。
4. `確認印を押して保存` を押します。
5. 完了ダイアログのOK後、保存先フォルダが開きます。

`tkinterdnd2` が利用できる環境では、PDFのドラッグ＆ドロップにも対応します。

## 注意事項

- このアプリはPDF編集アプリではありません。
- このアプリは電子署名・電子契約・本人性証明を目的とするものではありません。
- 印影は社内確認・作業記録用の確認印です。
- 元PDFは上書きせず、必ず別名PDFとして保存します。
- 1PDFにつき1か所のみ押印します。
- 押印対象は現在表示中のページです。

## 必要ライブラリ

- PyMuPDF
- Pillow
- tkinterdnd2（ドラッグ＆ドロップ用。未導入でもPDF選択ボタンは利用可能）
- PyInstaller（exe化用）

## ビルド方法

1. Python環境に必要ライブラリをインストールします。
2. `build.bat` を実行します。
3. `dist/Dake_Check_Stamp.exe` が生成されます。

DAKEシリーズ共通アイコンは `../../02_assets/dake_icon.ico` を参照します。アイコン正本は `02_assets/dake_icon.ico` で、アプリ個別のアイコンは使用しません。

## DAKE_META

```json
{
  "app_key": "dake_pdf_checkstamp",
  "display_name": "Dake確認印",
  "launcher_title": "PDF確認印",
  "launcher_description": "PDFに名字と日付入りの確認印を押します。",
  "site_title": "Dake確認印",
  "site_description": "PDFプレビュー上で位置を選び、名字と日付入りの確認印を1か所押せるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_PDF_CheckStamp",
  "exe_name": "Dake_Check_Stamp.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_CheckStamp_v1.0.0",
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

- PDF確認印アプリ
- 名字と日付入りの印影に対応
- クリック位置へ押印
- Windows向けexe
