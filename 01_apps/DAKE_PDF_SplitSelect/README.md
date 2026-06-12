# PDFページ選択保存

PDFから必要ページだけを見て選び、すばやく抜き出すためのデスクトップアプリです。

## 特徴

- PDF 1ファイル専用の単機能設計
- サムネイル選択を主導線にした直感操作
- `1-3,5,8-10` 形式の範囲入力にも対応
- 読み込み、サムネイル生成、保存は UI を止めずに非同期処理
- 表示中の範囲だけを描画し、サムネイルはキャッシュ

## ファイル構成

- `main.py`
- `requirements.txt`
- `build.bat`
- `README.md`

## 起動

```powershell
python -m pip install -r requirements.txt
python main.py
```

## しまりすくん連携CLI

`--from-shimarisu` を付けた場合のみGUIを開かず、指定ページを1つのPDFとして抽出します。

```powershell
DakePDF_Split_Select.exe --from-shimarisu --inputs "sample.pdf" --pages "1,3-5"
DakePDF_Split_Select.exe --from-shimarisu --inputs "sample.pdf" --pages "1" --output "out.pdf" --silent
DakePDF_Split_Select.exe --from-shimarisu --inputs "sample.pdf" --pages "1-3" --output "C:\Output"
```

- `--inputs` は先頭のPDFだけを使用します。
- `--pages` は `1`、`1,3,5`、`1-3`、`1,3-5` 形式に対応します。
- `--output` がPDFパスならそのPDF名、フォルダならその中へ保存します。
- `--output` 未指定時は元PDFと同じフォルダに `sample_extract_YYYYMMDD_HHMMSS.pdf` 形式で保存します。
- 不正な指定や存在しないページは短いエラーを `stderr` に出して exit code 1 で終了します。

## 使い方

1. `PDF追加` またはドラッグ＆ドロップで PDF を読み込みます。
2. サムネイルをクリックして必要ページを選びます。
3. 必要なときだけ範囲入力を使います。
4. `抽出する` で 1 つの PDF にまとめて保存します。
5. `1ページずつ出力` でページごとに保存します。

## 動作ルール

- サムネイル選択がある場合はそちらを優先します。
- 範囲入力が不正な場合は実行できません。
- 保存先を手動で選ばない場合は、読み込んだ PDF と同じフォルダを使います。
- 出力先に同名ファイルがある場合は、自動で連番を付けて保存します。

## ビルド

```powershell
build.bat
```

生成物は `dist\DakePDF_Split_Select` に出力されます。

## 文言管理

- UI 文言は `main.py` の `APP_NAME` `WINDOW_TITLE` `COPYRIGHT` `UI_TEXT` で管理しています。
- フッターリンク URL は `main.py` の `FOOTER_URLS` で管理しています。
- exe 名は `DakePDF_Split_Select`、表示名は `PDFページ選択保存` として分離しています。

## DAKE_META

```json
{
  "app_key": "dake_pdf_splitselect",
  "display_name": "DakePDF分割Select",
  "launcher_title": "PDF分割Select",
  "launcher_description": "PDFから必要ページだけを選んで保存します。",
  "site_title": "DakePDF分割Select",
  "site_description": "PDFの必要ページだけをサムネイルや範囲入力で選び、1つのPDFとして保存できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_PDF_SplitSelect",
  "exe_name": "DakePDF_Split_Select.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_SplitSelect_v1.0.0",
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

- PDFページ選択保存アプリ
- サムネイル選択に対応
- 範囲入力にも対応
- Windows向けexe
