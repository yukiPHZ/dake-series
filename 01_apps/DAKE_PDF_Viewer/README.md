# DakePDF見る

PDFを探して、すぐ見て、すぐ確認するための軽量PDFビューワーです。

保存、編集、注釈、結合、分割、変換はできません。PDFを見ることに絞ったDAKEシリーズのWindowsデスクトップアプリです。

## できること

- フォルダ内のPDFを一覧表示
- PDFを1ページずつ表示
- マウスホイールでスクロール
- Ctrl + マウスホイールで拡大・縮小
- Ctrl + 0 でウインドウ幅に合わせる
- Ctrl + F で文字検索
- Ctrl + P で現在表示中の1ページだけ印刷

画像PDF、スキャンPDFは文字検索できません。OCRは行いません。

## 基本操作

- フォルダ選択: 「フォルダを選ぶ」
- PDF選択: 左のPDF一覧をクリック
- スクロール: マウスホイール / ↑ / ↓
- 前後ページ: ← / → / PageUp / PageDown
- 先頭・最終ページ: Home / End
- 拡大・縮小: Ctrl + マウスホイール
- 幅に合わせる: Ctrl + 0
- 回転: R / L / 0
- 検索: Ctrl + F、Enter / Shift + Enter
- 現在ページ印刷: Ctrl + P

全ページ印刷、ページ範囲指定印刷、保存ボタンはありません。

## ビルド方法

依存ライブラリを入れてからビルドします。

```bat
pip install -r requirements.txt
```

```bat
build.bat
```

PyInstallerで `DakePDF_Viewer.exe` を作成します。

## アイコン

DAKEシリーズ共通アイコン `..\..\02_assets\dake_icon.ico` を使用します。個別アイコンは作成しません。

## DAKE共通仕様レビュー結果

2026-05-06 にDAKE共通仕様へ合わせて横断レビューしました。

- フォントは BIZ UDPGothic を最優先、Yu Gothic UI / Meiryo をフォールバックにしています。
- ヘッダーはアプリ名を重複表示せず、「PDFを見る」と短い説明文に整理しています。
- フッターは「シンプルそれDAKEシリーズ / 止まらない、迷わない、すぐ終わる。」と、戸建買取査定 / Instagram / Copyright の2ブロック構成です。
- フッターリンクは「戸建買取査定」「Instagram」のみクリック可能です。
- UI文言は `APP_NAME` / `WINDOW_TITLE` / `COPYRIGHT` / `UI_TEXT` に集約しています。
- 共通アイコン `..\..\02_assets\dake_icon.ico` を Tkinter 起動時と PyInstaller ビルド時に参照します。
- `build.bat` による `DakePDF_Viewer.exe` のビルドを確認しました。
- `dist\DakePDF_Viewer.exe` の起動確認を実施しました。
