# DakePDFファイル名整理

PDFの中身を変更せず、既存PDFファイル名の前または後ろに任意テキストを追加するDAKEシリーズの単機能アプリです。

## アプリ概要

- PDFの中身は変更しません。
- 既存ファイル名を残したまま、前後どちらかにテキストを追加します。
- 元PDFと同じフォルダ内でリネームします。
- 複数PDFをまとめて追加できます。
- PDF以外のファイルは受け付けません。

## 使い方

1. PDFを画面中央へドロップ、または「PDFを選択」から追加します。
2. 追加位置を「先頭に追加」または「末尾に追加」から選びます。
3. 追加するテキストを入力します。
4. 「ファイル名に足す」を押します。
5. 完了後、最初のPDFがあるフォルダが開きます。

## ファイル名ルール

既存ファイル名の `.pdf` を除いた部分を残し、アンダーバーで連結します。拡張子は `.pdf` に統一します。

```text
先頭に追加: 追加テキスト_base_name.pdf
末尾に追加: base_name_追加テキスト.pdf
```

例：

```text
IMG_1234.pdf -> 山田様_IMG_1234.pdf
IMG_1234.pdf -> IMG_1234_山田様.pdf
```

同名ファイルがある場合は連番を付け、既存ファイルを上書きしません。

```text
IMG_1234_山田様.pdf
IMG_1234_山田様_2.pdf
IMG_1234_山田様_3.pdf
```

## 注意事項

- PDF編集、OCR、PDF解析、自動判定は行いません。
- 既存ファイル名そのものは消しません。
- 保存先の変更、履歴、細かな設定画面はありません。
- ファイル名に使用できない文字は全角文字へ置き換えます。
- PDFが開かれている場合、リネームできないことがあります。

## ビルド方法

```bat
build.bat
```

ビルドに成功すると、`dist\DakePDF_Rename.exe` が作成されます。共通アイコンは `..\..\02_assets\dake_icon.ico` を使用します。

## DAKEシリーズ共通思想

単機能、軽量、迷わない、現場で止まらない。多機能化せず、確実に「ファイル名にテキストを足す」ことだけに絞っています。

## 2026-05-06 仕様変更・確認結果

- 旧仕様の分類プルダウンと固定リストを削除しました。
- 名前専用の入力、自動敬称付与、重複防止を削除しました。
- 追加位置と追加テキストによる前後追加仕様へ変更しました。
- DAKE共通フッター、共通色、フォント優先順、UI_TEXT管理を維持しました。
- `build.bat` によるexeビルドと `dist\DakePDF_Rename.exe` の起動確認を実施しました。

## DAKE_META

```json
{
  "app_key": "dake_pdf_rename",
  "display_name": "DakePDFファイル名整理",
  "launcher_title": "PDF名整理",
  "launcher_description": "PDF名の前後に任意テキストを追加します。",
  "site_title": "DakePDFファイル名整理",
  "site_description": "PDFの中身を変えず、ファイル名の前または後ろに任意テキストを追加できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_PDF_Rename",
  "exe_name": "DakePDF_Rename.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Rename_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- PDFファイル名整理アプリ
- 前後へのテキスト追加に対応
- 複数PDFをまとめて処理
- Windows向けexe
