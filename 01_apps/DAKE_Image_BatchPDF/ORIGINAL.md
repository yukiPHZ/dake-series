# ORIGINAL.md

## 正本宣言

このファイルは `DAKE_Image_BatchPDF` の真の正本です。

README.md、DAKE_META、release_body.md、booth_product.txt、Store表示などは、このファイルから派生するビューです。

## 基本情報

- app_id: `DAKE_Image_BatchPDF`
- internal_app_name: Dake画像まとめPDF
- display_name: 画像まとめPDF
- window_title: 画像まとめPDF
- folder_name: `01_apps/DAKE_Image_BatchPDF/`
- exe_name: `DakeImage_BatchPDF.exe`
- category: 画像/PDF
- target_platform: Windows
- status: draft
- app_type: market
- completion_goal: formal_release

## 目的

複数の画像ファイルを追加し、順番を確認・変更してから、1つのPDFファイルとして保存するためのDAKEアプリです。

## 実務背景

iPhone等から複数画像のみが送られてきた後の、案件フォルダ整理・共有・閲覧を楽にするため、複数画像を1つのPDFへまとめる。

物件写真、現場写真、設備写真などが画像ファイルのまま大量に届くと、転送、社内共有、案件フォルダ内での確認、後日の資料確認に摩擦が出ます。

このアプリでは、画像を追加し、並べ、配置を選び、PDFとして保存するところまでを最短で終わらせます。

## 対象ユーザー

- iPhone等から届いた複数画像を案件資料としてまとめたい人
- 物件写真、現場写真、設備写真を1つの確認用PDFにしたい人
- 写真管理や画像編集ではなく、共有しやすいPDF化だけを短く終わらせたい人

## 入力

- JPG
- JPEG
- PNG
- BMP
- WEBP
- HEIC
- HEIF

## 処理

1. 複数画像をドラッグ＆ドロップまたはファイル選択で追加する。
2. 追加画像の順番番号、サムネイル、ファイル名を表示する。
3. 選択画像を上へ / 下へ移動して順番を変える。
4. 選択画像を削除、またはすべてクリアする。
5. PDF配置を `1ページに1枚` または `1ページに4枚` から選ぶ。
6. A4縦、白背景、元画像の縦横比維持、切り取りなしでPDFを作成する。

## 出力

- 1つのPDFファイル
- 初期保存名候補: `画像まとめ.pdf`
- 保存先とファイル名はユーザーが選択する
- 保存完了後に完了ダイアログを表示し、保存先フォルダを開く

## 1ページ1枚仕様

- 画像1枚につきPDF 1ページ
- A4縦のページ内に、余白を確保して画像全体を最大表示
- 元画像の縦横比を維持
- 画像を切り取らない
- 画像を引き伸ばさない
- 中央配置

## 1ページ4枚仕様

- A4縦を2列×2行に分ける
- 配置順は左上、右上、左下、右下
- 5枚目以降は次ページへ送る
- 端数枠は空白のままにする
- 端数画像を複製しない
- 端数画像を無理に拡大しない

## iPhone / HEIC対応方針

- `pillow-heif` を利用してHEIC / HEIFをPillowで読み込む
- PyInstaller buildでは `--collect-all=pillow_heif` を指定する
- HEIC / HEIFはコード上だけでなく、exeでの読込確認対象とする
- EXIF Orientationは `ImageOps.exif_transpose()` で補正する

## 透明画像対応

- PNG / WEBP等の透明画像は白背景へ合成する
- PDF出力時に黒背景化しないようRGB化してから配置する

## 非破壊方針

- 元画像を削除しない
- 元画像を移動しない
- 元画像を上書きしない
- PDFは一時ファイルへ生成し、完成後に保存先へ確定する

## UI方針

- DAKE共通UIに合わせる
- 白 / ライトグレー基調
- フラットで静かな実務向けUI
- UI文言は `APP_NAME`、`WINDOW_TITLE`、`COPYRIGHT`、`UI_TEXT` で一元管理する
- フッターはDAKE共通文言を省略しない

## 応答性方針

- サムネイル生成とPDF生成はworker threadで実行する
- UI更新はqueueとafterで行う
- PDF生成中は進捗を表示する
- PDF生成中はキャンセル要求を受け付ける
- キャンセルは画像単位またはページ単位の境界で安全に処理する

## やらないこと / 非ゴール

- 既存の `DAKE_Image_ToPDF` を置き換えない
- 画像編集アプリにしない
- 写真管理アプリにしない
- OCRを入れない
- AIを入れない
- 案件管理機能を入れない
- 余計な設定項目を増やさない

## 技術構成

- GUI: tkinter
- D&D: tkinterdnd2
- 画像処理: Pillow
- HEIC / HEIF: pillow-heif
- PDF生成: reportlab
- build: PyInstaller

## DAKE_META生成用情報

```json
{
  "app_key": "DAKE_Image_BatchPDF",
  "display_name": "画像まとめPDF",
  "launcher_title": "画像まとめPDF",
  "launcher_description": "複数画像を順番どおり1つのPDFにまとめます。",
  "site_title": "Dake画像まとめPDF",
  "site_description": "複数の画像を並べて、1ページ1枚または4枚のPDFにまとめるWindows向けDAKEアプリです。",
  "update_summary": "複数画像を順番確認して1つのPDFへまとめる新規アプリを追加しました。",
  "folder_name": "DAKE_Image_BatchPDF",
  "exe_name": "DakeImage_BatchPDF.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "draft",
  "show_in_launcher": false,
  "show_on_site": false,
  "app_type": "market",
  "completion_goal": "formal_release"
}
```

## release_body生成用情報

- 複数画像を順番どおり1つのPDFにまとめます。
- 1ページ1枚 / 1ページ4枚を選べます。
- JPG / PNG / WEBP / HEIC / HEIFなどに対応します。
- Windows向けexeです。

## Codex作業時の注意

- 既存の単体画像→PDFアプリを変更しない
- 新規独立アプリとして扱う
- build/、dist/、*.spec、*_config.json、__pycache__/、*.pyc、exeを通常commitへ含めない
- GitHub Release公開、BOOTH登録、dakeapp.com公開、Store登録、Stripe変更、SNS投稿は別指示まで行わない
