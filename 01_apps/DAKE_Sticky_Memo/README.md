# Dake付箋メモ

アプリウインドウ内に付箋メモを作成し、自由に動かし、不要になったら捨てられる軽量デスクトップアプリです。

保存、分類、検索などの整理機能は初期版では持たせていません。思いついたことをその場で書いて、位置を動かして、終わったら消すだけの小さな道具です。

## アプリ概要

- 表示名: `付箋メモ`
- exe名: `DakeSticky_Memo.exe`
- 実装: Python + Tkinter
- 共通アイコン: `..\..\02_assets\dake_icon.ico`
- 保存: 初期版では保存しません

## できること

- `付箋を追加` で新しい付箋を作成
- 付箋に自由にテキスト入力
- 付箋をドラッグしてウインドウ内で移動
- 付箋右上の `×` で削除
- `全部消す` で全付箋を削除
- ウインドウ外へ出ないよう付箋位置を自動制限
- 画面幅に合わせてDAKE共通フッターを1段/2段で表示

## やらないこと

- 保存
- 色変更
- 検索
- タグ
- 画像添付
- エクスポート
- 分類や高度な整理

## 実行

```bat
python main.py
```

## ビルド

```bat
build.bat
```

`dist\DakeSticky_Memo.exe` が作成されます。

## DAKE_META

```json
{
  "app_key": "DAKE_Sticky_Memo",
  "display_name": "付箋メモ",
  "launcher_title": "付箋メモ",
  "launcher_description": "付箋を書いて、動かして、捨てられる軽量メモです。",
  "site_title": "Dake付箋メモ",
  "site_description": "アプリ内に付箋を作成し、自由に動かし、不要になったら消せるWindows向け軽量メモアプリです。",
  "update_summary": "フッター表示とUI文言をDAKE共通仕様に合わせて調整しました。",
  "folder_name": "DAKE_Sticky_Memo",
  "exe_name": "DakeSticky_Memo.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

```text
- 付箋を追加して、その場でメモできます。
- 付箋をドラッグして自由に並べ替えできます。
- 不要な付箋を個別削除、または全消去できます。
- Windows向けexeです。
```
