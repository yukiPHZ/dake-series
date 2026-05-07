# DAKE_UI_TEXT_RULE

DAKEアプリのUI文言管理ルールです。

## 目的

- 日本語UIを横断レビューしやすくする。
- 文字化けや表記ゆれを早く見つける。
- アプリごとの画面文言を1か所で確認できるようにする。

## 必須定義

```python
APP_NAME = "Dakeアプリ名"
WINDOW_TITLE = "短いタイトル"
COPYRIGHT = "© 2026 しまりす不動産 / Vibe-Coded by Yukihiko Kikuta"

UI_TEXT = {
    "main_title": "主見出し",
    "main_description": "短い説明",
}
```

## UI_TEXTに入れるもの

- ヘッダー文言。
- ボタン文言。
- ラベル文言。
- プレースホルダー。
- ステータス。
- エラーメッセージ。
- ダイアログタイトル。
- ダイアログ本文。
- フッター文言。
- ファイル名テンプレート。
- PDF出力用の固定文言。

## 直書き禁止

避ける例:

```python
tk.Label(root, text="PDFを追加してください")
messagebox.showerror("エラー", "処理できませんでした")
```

推奨:

```python
tk.Label(root, text=UI_TEXT["drop_hint"])
messagebox.showerror(UI_TEXT["error_title"], UI_TEXT["error_message"])
```

## APP_NAMEとWINDOW_TITLE

- `APP_NAME` は正式表示名。
- `WINDOW_TITLE` はウインドウタイトル用の短い名前。
- 画面内でアプリ名を繰り返しすぎない。
- ランチャーやサイト用の名前はREADMEの `DAKE_META` で管理する。

## COPYRIGHT

基本:

```python
COPYRIGHT = "© 2026 しまりす不動産 / Vibe-Coded by Yukihiko Kikuta"
```

ルール:

- フッター表示の正本として使う。
- 複数箇所に同じ文字列を直書きしない。

## レビュー基準

- `UI_TEXT` を見れば画面文言が追える。
- 日本語が自然で短い。
- ボタンは動詞中心。
- エラーは責めない表現にする。
- ステータスは短く、現在状態が分かる。
- 文字化けがない。

## 文字化け対策

- ファイルはUTF-8で保存する。
- PowerShellで日本語を生成する場合はUTF-8出力を明示する。
- `????`、`�`、不自然な mojibake を検出したら修正する。
- build後exeの画面でも確認する。
