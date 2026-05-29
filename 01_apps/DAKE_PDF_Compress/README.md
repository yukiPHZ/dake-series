# DakePDF圧縮

PDFを1つ追加して、ファイルサイズを軽くするDAKEのWindowsデスクトップアプリです。

## 使い方

1. PDFをドラッグ＆ドロップします。
2. 表示されたファイル名、元サイズ、保存予定ファイル名を確認します。
3. 「圧縮して保存」を押します。
4. 完了後、保存先フォルダが開きます。

## SHIMARISU連携CLI

SHIMARISUから呼び出す場合だけ、GUIなしでPDF圧縮を実行できます。

```bat
DakePDF_Compress.exe --from-shimarisu --inputs "A.pdf"
```

複数PDFを渡す場合:

```bat
DakePDF_Compress.exe --from-shimarisu --inputs "A.pdf" "B.pdf"
```

- 成功時は exit code `0` で、圧縮後PDFのパスを標準出力に出します。
- 失敗時は exit code `1` で、短いエラーを標準エラーに出します。
- `--help-cli` でCLI仕様を表示して終了します。
- `--from-shimarisu` がない通常起動は、これまで通りGUIで起動します。

## 出力ファイル名

元PDFと同じフォルダに、次の名前で保存します。

- `元ファイル名_compressed.pdf`
- 同名ファイルがある場合は `元ファイル名_compressed_2.pdf` のように自動で連番を付けます。

## 注意事項

- 元PDFは上書きしません。
- PDFの構造によっては、圧縮効果が小さい場合があります。
- 暗号化PDFや破損PDFは処理できない場合があります。
- 初期実装はPDF 1件のみ対応です。

## v2 しっかり圧縮

- v2では Ghostscript を優先して、標準でしっかり圧縮します。
- Ghostscript がある環境では、`/ebook` 相当の圧縮で効果が高くなります。
- Ghostscript が見つからない場合や処理できない場合は、内蔵fallback処理を使います。
- 圧縮後PDFが元PDFより大きい場合は完成扱いにせず、fallbackを試します。
- fallbackでも小さくならない場合は保存せず、圧縮効果がない旨を表示します。
- PDFの構造によっては、圧縮効果が小さい場合があります。
- 元PDFは上書きしません。

## ビルド方法

事前に依存ライブラリをインストールします。

```bat
pip install -r requirements.txt
```

その後、以下を実行します。

```bat
build.bat
```

`dist\DakePDF_Compress.exe` が作成されます。共通アイコン `..\..\02_assets\dake_icon.ico` が存在する場合は自動で使用します。

## DAKEシリーズ表記

シンプルそれDAKEシリーズ  
© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta

## 共通仕様レビュー

- UI文言は `APP_NAME`、`WINDOW_TITLE`、`COPYRIGHT`、`UI_TEXT` に集約しています。
- フォントは BIZ UDPGothic を最優先にし、Yu Gothic UI / Meiryo にフォールバックします。
- ヘッダーは機能タイトルと短い説明のみを表示し、画面内でアプリ名を重複表示しません。
- フッターは DAKE共通仕様に合わせ、広幅時は左右2ブロック、狭幅時は中央寄せ2段構成に切り替えます。
- 共通アイコンは `..\..\02_assets\dake_icon.ico` を参照し、存在しない場合も起動時に落ちないようにしています。
- 圧縮処理は別スレッドで実行し、処理中ステータスとプログレス表示でUIが固まって見えないようにしています。
- 初期ウインドウサイズを `860x740`、最小サイズを `760x720` に調整し、起動直後からフッターが見えるようにしています。

## 起動確認結果

- 2026-05-06: ソース構文確認、圧縮関数の単体確認、重複ファイル名回避、PyInstallerビルドを確認しました。
- 2026-05-06: `dist\DakePDF_Compress.exe` の短時間起動確認を行い、プロセス起動後に停止できることを確認しました。
- 2026-05-06: 共通アイコン参照、初期ウインドウサイズ、最小高さの設定を再確認しました。
- 2026-05-14: SHIMARISU連携用CLIモードを追加し、`--help-cli`、正常圧縮、入力なし、存在しないPDFの確認を行いました。
- 2026-05-29: v2として Ghostscript 優先のしっかり圧縮へ変更し、内蔵fallbackと圧縮効果なし判定を確認しました。
- 2026-05-29: Ghostscript未検出環境でfallback圧縮を確認しました。実PDFは 9,723,881 bytes から 1,118,764 bytes へ圧縮され、削減率は 88.5% でした。低圧縮時の注意表示、`python -m py_compile main.py`、`build.bat`、`dist\DakePDF_Compress.exe` 起動も確認しました。
- 2026-05-30: SHIMARISU から `dist\DakePDF_Compress.exe --from-shimarisu --inputs "対象PDF"` のv2 CLI接続を確認しました。成功時はstdoutへ保存先PDFを出し、失敗時は短いstderrと exit `1` で終了します。

## DAKE_META

```json
{
  "app_key": "dake_pdf_compress",
  "display_name": "DakePDF圧縮",
  "launcher_title": "PDF圧縮",
  "launcher_description": "PDFを追加してファイルサイズを軽くします。",
  "site_title": "DakePDF圧縮",
  "site_description": "PDFを1つ追加し、ファイルサイズを軽くした別名PDFとして保存できるWindows向けアプリです。",
  "update_summary": "READMEメタ情報とRelease本文を整備。スクリーンショットをassets/screenshot.webpに作成。",
  "folder_name": "DAKE_PDF_Compress",
  "exe_name": "DakePDF_Compress.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Compress_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- PDF圧縮アプリ
- ドラッグ＆ドロップ対応
- 保存名の自動連番に対応
- Windows向けexe
