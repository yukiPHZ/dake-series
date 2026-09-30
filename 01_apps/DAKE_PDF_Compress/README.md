# DakePDF圧縮 v1.1.0

PDFを追加するだけで、画質を保ちながら内容に合った方法で軽くするDAKEのWindowsデスクトップアプリです。

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
- GUIはPDF1件、CLIは複数件に対応します。

## 自動で、読みやすく軽く

- PDFの内容に合わせて、画質と容量のバランスがよい結果を自動で選びます。設定は不要です。
- 追加すると「PDFを確認中...」、準備ができると「圧縮できます」と表示します。
- 処理中は現在の作業を表示し、完了するとサイズの変化と削減率が分かります。
- 元PDFは変更せず、同じフォルダへ別名保存します。十分に軽いPDFは保存しない場合があります。
- Adaptive Compression版はHuman Gateを通過した正式版v1.1.0です。

## ビルド方法

Python 3.10以上を使用し、検証済みの依存をインストールします（Windows / Python 3.12.14で検証、PyMuPDF 1.28.2に固定）。

```bat
pip install -r requirements.txt
```

その後、以下を実行します。

```bat
build.bat
```

`dist\DakePDF_Compress.exe` が作成されます。共通アイコン `..\..\02_assets\dake_icon.ico` を埋め込みます。onefileを維持しています。

## DAKEシリーズ表記

シンプルそれDAKEシリーズ  
© 2026 しまりす不動産 — Vibe-Coded by Yukihiko Kikuta

## 共通仕様レビュー

- UI文言は `APP_NAME`、`WINDOW_TITLE`、`COPYRIGHT`、`UI_TEXT` に集約しています。
- フォントは BIZ UDPGothic を最優先にし、Yu Gothic UI / Meiryo にフォールバックします。
- ヘッダーは機能タイトルと短い説明のみを表示し、画面内でアプリ名を重複表示しません。
- フッターは DAKE共通仕様に合わせ、広幅時は左右2ブロック、狭幅時は中央寄せ2段構成に切り替えます。
- 共通アイコンは `..\..\02_assets\dake_icon.ico` を参照し、存在しない場合も起動時に落ちないようにしています。
- PDFの確認と圧縮は別プロセスで実行し、現在の作業を表示します。
- 初期ウインドウサイズを `860x740`、最小サイズを `760x720` に調整し、起動直後からフッターが見えるようにしています。

## 開発検証

2026-09-12: Adaptive Compression、非同期追加、完了表示を改善。検証記録と制限は [開発用レポート](dev/ADAPTIVE_BENCHMARK.md) を参照してください。ログはアプリフォルダの `logs/` に保存します。

## 過去の確認履歴（v2は当時の説明名）

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
  "launcher_description": "PDFを追加して、画質を保ちながらしっかり軽くします。",
  "site_title": "DakePDF圧縮",
  "site_description": "PDFを追加するだけで、画質を保ちながら内容に合った方法で軽くした別名PDFを保存できるWindows向けアプリです。",
  "update_summary": "v1.1.0。PDFごとに圧縮方法を自動選択し、品質確認と操作応答性を改善。",
  "folder_name": "DAKE_PDF_Compress",
  "exe_name": "DakePDF_Compress.exe",
  "version": "1.1.0",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_PDF_Compress_v1.1.0",
  "app_type": "market",
  "completion_goal": "formal_release",
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

- PDFの内容に応じて圧縮方法を自動選択
- 圧縮後の品質確認を強化
- 写真・スキャン・図面系PDFへの対応を改善
- PDF追加時・圧縮中の操作応答性を改善
- 圧縮結果を分かりやすく表示
- SHIMARISU連携CLI互換を維持
