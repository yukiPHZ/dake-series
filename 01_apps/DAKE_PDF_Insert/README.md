# DakePDF差し込み

元PDFの好きなページ間に、別のPDFを1ファイル差し込んで保存するWindows向け単機能アプリです。

## できること

- 元PDFを1ファイル選択
- 差し込むPDFを1ファイル選択
- 元PDFのサムネイルを見ながら差し込み位置を選択
- 先頭 / 各ページの後 / 末尾へ差し込み
- 新しいPDFとして保存

## 使い方

1. `元PDFを追加` で元になるPDFを選びます。
2. `差し込むPDFを追加` で差し込みたいPDFを選びます。
3. サムネイルのページ間にある `＋`、またはページカードの `この後へ` で位置を選びます。
4. 必要に応じて `保存先を選ぶ` を押します。
5. `差し込んで保存` を押します。

元PDFと差し込むPDFは変更しません。出力名は `元ファイル名_差し込み済み.pdf` です。同名ファイルがある場合は連番を付けます。

## 必要なもの

- Windows
- Pythonでビルドする場合は `requirements.txt` のライブラリ

## ビルド

```bat
build.bat
```

ビルド後のexe名は `dist\DakePDF_Insert.exe` です。DAKE共通アイコン `..\..\02_assets\dake_icon.ico` を使用します。

## 注意

- パスワード保護PDFへの完全対応は行いません。
- 大切なファイルは事前にバックアップを推奨します。
- 初回起動時にWindowsの確認画面が表示される場合があります。

## DAKE_META

```json
{
  "app_key": "dake_pdf_insert",
  "display_name": "DakePDF差し込み",
  "launcher_title": "PDF差し込み",
  "launcher_description": "元PDFの好きなページ間に、別のPDFを差し込みます。",
  "site_title": "DakePDF差し込み",
  "site_description": "元PDFの好きなページ間に、別のPDFを1ファイル差し込んで保存するWindows向け単機能アプリです。",
  "update_summary": "新規作成。元PDF、差し込みPDF、差し込み位置1か所を指定して保存できる初版。",
  "folder_name": "DAKE_PDF_Insert",
  "exe_name": "DakePDF_Insert.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "draft",
  "show_in_launcher": false,
  "show_on_site": false,
  "app_type": "market",
  "completion_goal": "formal_release",
  "payment_status": "unset"
}
```
