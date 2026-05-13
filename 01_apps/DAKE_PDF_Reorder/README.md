# DakePDFページ並べ替え

PDFのページ順だけをドラッグ＆ドロップで並べ替え、別名PDFとして保存するDAKEシリーズの単機能アプリです。

## 使い方

1. PDFをドラッグ＆ドロップ、または「PDFを追加」から読み込みます。
2. 表示されたサムネイルをドラッグして順番を変更します。
3. 必要に応じて「保存先を選ぶ」で出力フォルダを指定します。
4. 「並べ替えて保存」で `元ファイル名_reordered.pdf` を保存します。
5. 完了後、保存先フォルダが開きます。

## やること

- PDFを1ファイルだけ読み込む
- ページごとのサムネイル表示
- ドラッグ＆ドロップによるページ順変更
- 並べ替え後PDFの保存
- 前回保存先の記憶

## やらないこと

- PDF結合
- PDF分割
- ページ削除
- ページ回転
- 文字編集、注釈編集、OCR、圧縮
- 複数PDF同時処理

## ビルド方法

```bat
build.bat
```

必要なライブラリは `requirements.txt` を参照してください。

## 注意事項

- 元PDFは上書きしません。
- 出力先に同名ファイルがある場合は、上書き確認を表示します。
- 保護されたPDFや破損したPDFは処理できない場合があります。

## DAKE共通仕様レビュー結果

- フォントは BIZ UDPGothic を最優先にし、Yu Gothic UI / Meiryo をフォールバックにしています。
- ヘッダーはアプリ名を重複表示せず、機能タイトルと短い説明文だけにしています。
- フッターは「シンプルそれDAKEシリーズ / 止まらない、迷わない、すぐ終わる。」と、戸建買取査定 / Instagram / コピーライトの2ブロック構成です。
- UI文言は APP_NAME、WINDOW_TITLE、COPYRIGHT、UI_TEXT に集約しています。
- `build.bat` は共通アイコン `..\..\02_assets\dake_icon.ico` を指定します。
- 2026-05-06 時点で `build.bat` による exe 生成と `dist\DakePDF_Reorder.exe` の起動確認を実施済みです。

## 2026-05-06 追加修正

- 空状態エリアをクリックしてPDFを追加できるようにしました。
- サムネイルカードのドラッグ位置を座標で判定し、ページ順と内部 `page_order` が同期するようにしました。
- 保存時は `page_order` の順番でPDFを書き出します。
- 空状態文言を「ドラッグ＆ドロップ または クリックして追加」に調整しました。

## DAKE_META

```json
{
  "app_key": "dake_pdf_reorder",
  "display_name": "DakePDFページ並べ替え",
  "launcher_title": "PDF並べ替え",
  "launcher_description": "PDFページ順をドラッグで並べ替えます。",
  "site_title": "DakePDFページ並べ替え",
  "site_description": "PDFのページサムネイルをドラッグ＆ドロップで並べ替え、別名PDFとして保存できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_PDF_Reorder",
  "exe_name": "DakePDF_Reorder.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Reorder_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- PDFページ並べ替えアプリ
- サムネイルのドラッグ操作に対応
- 別名PDFとして保存
- Windows向けexe
