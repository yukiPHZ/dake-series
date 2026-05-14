# Dake二人メモ

右の人、左の人。2人で同じメモ空間を使いながら、互いの本文領域を奪わない共同メモです。

チャットではなく、横に存在する作業メモです。相手のメモには短い確認だけを残せます。

## アプリ概要

- 表示名: `Dake二人メモ`
- exe名: `DakeTwoPerson_Memo.exe`
- 実装: Python + Tkinter
- 同期: Python標準ライブラリの `socket` / `threading` / `json`
- 保存: `two_person_memo_data.json`

## できること

- 左の人、右の人を選んで使えます
- 自分側のメモブロックだけ本文編集できます
- 相手側の本文は編集できません
- 相手のメモに `見た`、`確認`、`完了`、`保留` を残せます
- 自分のメモブロックだけ削除できます
- 同じLAN内で簡易同期できます
- ローカルに自動保存します

## 使い方

1. 役割で `左の人` または `右の人` を選びます。
2. 自分側の `メモを追加` を押します。
3. 自分のブロックに本文を書きます。
4. 相手のブロックには確認アクションだけを押します。

## 同期方法

ホスト側:

1. `ポート` を確認します。初期値は `8765` です。
2. `ホスト開始` を押します。
3. 表示された `IPアドレス:ポート` を相手に伝えます。

参加側:

1. `参加先` にホストのIPアドレスを入れます。
2. `ポート` にホストと同じ番号を入れます。
3. `参加` を押します。

同期データはJSONで送受信します。クラウド、ログイン、外部APIは使いません。

## 注意事項

- 同じWi-Fiまたは同じLAN内で使う想定です。
- Windowsのファイアウォールが通信を止める場合があります。
- v1では自動探索やクラウド同期はありません。
- 同じブロックの本文は所有者だけが編集できます。
- 競合時はブロックの `updated_at` と役割に沿って新しい内容を優先します。
- 同期に失敗してもローカル保存は残ります。

## 実行

```bat
python main.py
```

## ビルド

```bat
build.bat
```

`dist\DakeTwoPerson_Memo.exe` が作成されます。

## DAKE_META

```json
{
  "app_key": "dake_two_person_memo",
  "display_name": "Dake二人メモ",
  "launcher_title": "Dake二人メモ",
  "launcher_description": "2人で使う、ぶつからない共同メモ。",
  "site_title": "Dake二人メモ",
  "site_description": "右の人、左の人。互いのメモを奪わず、確認だけ返せる共同メモです。",
  "update_summary": "同一LAN内で使える2人用メモを追加しました。",
  "folder_name": "DAKE_TwoPerson_Memo",
  "exe_name": "DakeTwoPerson_Memo.exe",
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_TwoPerson_Memo_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

```text
- 2人で使う、ぶつからない共同メモです。
- 左右それぞれのメモブロックを追加できます。
- 相手の本文は編集せず、確認だけを残せます。
- 同じLAN内で簡易同期できます。
- Windows向けexeです。
```

## できること

- 左右それぞれのメモブロックを追加できます
- 相手の本文は編集できません
- 相手のメモに「見た」「確認」「完了」「保留」を残せます
- 同じLAN内で簡易同期できます
- ローカルに自動保存します
```
