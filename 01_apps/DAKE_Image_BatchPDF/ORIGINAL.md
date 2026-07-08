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
6. EXIF Orientationを正規化し、必要に応じて透明画像を白背景へ合成する。
7. 配置枠に対する0度と時計回り90度の最大表示面積を比較し、より大きく表示できる向きを採用する。
8. A4縦、白背景、元画像の縦横比維持、切り取りなしでPDFを作成する。

## 出力

- 1つのPDFファイル
- 1ページ1枚の初期保存名: `Dake_画像まとめ_1枚配置_YYYYMMDD_HHMMSS.pdf`
- 1ページ4枚の初期保存名: `Dake_画像まとめ_4枚配置_YYYYMMDD_HHMMSS.pdf`
- 同名ファイルが存在する場合は、拡張子の前へ `_2`、`_3` のような連番を付けて既存PDFを上書きしない
- 保存先とファイル名はユーザーが選択する
- 保存完了後に完了ダイアログを表示し、保存先フォルダを開く

## 回転仕様

各画像は次の順で処理する。

1. 画像を読み込む。
2. EXIF Orientationを正規化する。
3. 必要に応じて透明画像を白背景へ合成する。
4. 実際の配置枠に対する0度配置時の最大表示面積を計算する。
5. 同じ配置枠に対する時計回り90度配置時の最大表示面積を計算する。
6. より大きく表示できる向きを採用する。

- 90度回転を採用する場合は時計回り90度へ統一する
- 同一またはほぼ同一の表示面積では0度を優先する
- 元画像の縦横比を維持する
- 画像をトリミングしない
- 画像を引き伸ばさない
- 配置枠内で最大表示する
- 中央配置する
- PDFの背景は白とする

## 1ページ1枚仕様

- 画像1枚につきPDF 1ページ
- A4縦のページ内に、余白を確保して画像全体を最大表示
- 元画像の縦横比を維持
- 画像を切り取らない
- 画像を引き伸ばさない
- 中央配置
- A4縦の実際の配置可能領域に対して0度と時計回り90度の表示面積を比較する

## 1ページ4枚仕様

- A4縦を2列×2行に分ける
- 配置順は左上、右上、左下、右下
- 5枚目以降は次ページへ送る
- 端数枠は空白のままにする
- 端数画像を複製しない
- 端数画像を無理に拡大しない
- 各配置枠に対して0度と時計回り90度の表示面積を個別に比較する
- 最終ページも2列×2行の固定配置を維持する
- 残数が1枚、2枚、3枚の場合も可変レイアウトへ変更しない
- 画像5枚の場合、2ページ目は左上へ5枚目を配置し、残り3枠を空白とする

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
  "update_summary": "画像と配置枠の比率を比較し、必要な画像を時計回り90度へ回転して大きく配置するよう改善しました。",
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
- 配置枠ごとに0度と時計回り90度を比較し、より大きく表示できる向きを選びます。
- 配置モードと日時を含む初期保存名を使い、同名ファイルは連番で上書きを防ぎます。
- JPG / PNG / WEBP / HEIC / HEIFなどに対応します。
- Windows向けexeです。

## Codex作業時の注意

- 既存の単体画像→PDFアプリを変更しない
- 新規独立アプリとして扱う
- build/、dist/、*.spec、*_config.json、__pycache__/、*.pyc、exeを通常commitへ含めない
- GitHub Release公開、BOOTH登録、dakeapp.com公開、Store登録、Stripe変更、SNS投稿は別指示まで行わない
