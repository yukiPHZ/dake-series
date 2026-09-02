# DakePDF俯瞰名前変更

PDFを開いて、閉じて、名前を変える。その往復をなくします。

フォルダ内のPDFを1ページ目のサムネイルで俯瞰しながら、PDFごとに新しい名前を入力し、変更分だけまとめて反映するWindows向けアプリです。

## できること

- 選択フォルダ直下のPDFをサムネイルカードで一覧表示
- PDFごとに異なる新しいファイル名を入力
- 変更分だけを衝突させずに一括反映
- 直前に成功した一括変更を1回だけUndo
- 小・標準・大の表示切替と、一覧上のマウスホイールスクロール

PDF本文の編集、OCR、AI自動命名、サブフォルダの再帰処理は行いません。

## 使い方

1. `フォルダを選ぶ` を押します。
2. 1ページ目のサムネイルを見ながら、各カードの名前欄を編集します。
3. `名前変更を反映 n` を押し、内容を確認します。
4. 必要な場合は `変更を元に戻す` で直前の一括変更を戻します。

`リフレッシュ` は現在のフォルダ選択、カード、Undo履歴、大プレビューを破棄し、起動直後の未選択状態へ戻します。未反映の入力がある場合は先に確認します。

## 安全仕様

- 反映前に全対象を検証し、1件でも問題があればファイル操作を始めません。
- 同名ファイルを上書きしません。
- 入れ替えや循環変更では一時名を挟む2段階リネームを使います。
- 処理途中の失敗時は可能な範囲で全件を元名へ戻します。
- PDF本文は変更しません。
- ネットワーク通信は行いません。

大切なファイルは事前にバックアップし、反映前の確認画面で変更内容を確認してください。

## 動作環境

- Windows
- 配布版は `DakePDF_OverviewRename.exe` 単体で起動
- 開発時は Python 3、Tkinter、Pillow、pypdfium2 が必要

## 開発とビルド

```bat
pip install -r requirements.txt
pytest -q
build.bat
```

`build.bat` はPyInstallerのonefile / noconsoleで `dist\DakePDF_OverviewRename.exe` を作成します。アプリウインドウとexeにはDAKE共通アイコン `02_assets/dake_icon.ico` を使用します。

## 第三者ライセンス

Windows配布版で使う pypdfium2 5.13.0 / PDFium 153.0.7999.0（pdfium-binaries）のwheelに記録されたライセンス文書を、`third_party_licenses/pypdfium2-5.13.0/` に収録しています。配布zipには `THIRD_PARTY_NOTICES.txt` とこのフォルダを含めます。

## 正式版情報

初回正式版はバージョン1.0.0、販売価格500円です。[GitHub Release](https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_OverviewRename_v1.0.0)から正式配布zipをダウンロードできます。[BOOTH](https://peakheadz.booth.pm/items/8798555)でも販売しています。dakeapp.com、Storeの正式出荷ラインにも掲載します。

## DAKE_META

```json
{
  "app_key": "dake_pdf_overview_rename",
  "display_name": "DakePDF俯瞰名前変更",
  "launcher_title": "PDF俯瞰名前変更",
  "launcher_description": "PDFを開かず、サムネイルを見ながら名前を変更します。",
  "site_title": "DakePDF俯瞰名前変更",
  "site_description": "フォルダ内のPDFをサムネイルで俯瞰しながら、PDFごとに名前を変更できるWindows向けアプリです。",
  "update_summary": "初回正式版。PDFの俯瞰表示、安全な一括名前変更、Undoに対応。",
  "folder_name": "DAKE_PDF_OverviewRename",
  "exe_name": "DakePDF_OverviewRename.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_OverviewRename_v1.0.0",
  "app_type": "market",
  "completion_goal": "formal_release",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- PDFの1ページ目サムネイルをフォルダ単位で俯瞰表示
- PDFごとに入力した名前を変更分だけ安全に一括反映
- 衝突事前検証、ロールバック、直前の一括変更のUndoに対応
- 固定workerとPDFium排他制御で一覧操作の応答性を維持
