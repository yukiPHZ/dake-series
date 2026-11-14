# Dake名前フォルダ中身見ながら変更

PDF・画像を見ながら、フォルダ名をまとめて整理するWindowsアプリです。
現在はv1.0.0候補。Human Review待ち・正式公開HOLD。

## 使い方
1. 親フォルダを選択するか、親フォルダ用領域へドロップします。
2. フォルダを選び、右の内容をスクロールして確認します。
3. 変更予定名を入力し、次のフォルダへ進みます。
4. 「変更内容を確認」から旧名・新名を確認して実行します。
5. 必要なら「直前の変更を元に戻す」でUndoします。

直接複数フォルダを入れる場合は専用ドロップ領域を使ってください。
ファイルの中身・名前は変更しません。子フォルダ内部を走査しません。
同名入替・大文字小文字だけの変更・リンク経由の対象は受け付けません。
途中失敗時は成功分だけ変更済みになります。結果と履歴を確認してください。
Undoはセッション内のみ。外部変更や衝突があると停止します。
強制終了後は履歴の新旧パスを確認し、状態不明のまま再実行しないでください。

## 開発
Python 3.12で `python -m pip install -r requirements.txt` 後、`python main.py`。
`python main.py --launch-check` で最小GUI起動確認。
`python build.py` または `build.bat` でクリーンなonefile候補を作成します。
`python -m pytest tests -q` で合成データ安全性試験。
正本と詳細な安全・復旧方針は[ORIGINAL.md](ORIGINAL.md)。

## DAKE_META

```json
{
  "app_key": "dake_folder_overviewrename",
  "display_name": "Dake名前フォルダ中身見ながら変更",
  "launcher_title": "中身を見てフォルダ名変更",
  "launcher_description": "PDF・画像を見ながらフォルダ名をまとめて変更。",
  "site_title": "Dake名前フォルダ中身見ながら変更",
  "site_description": "フォルダを一つずつ開かず、中身を見て名前を整理するWindowsアプリ。",
  "update_summary": "v1.0.0 初期候補。Human Review待ち。",
  "folder_name": "DAKE_Folder_OverviewRename",
  "exe_name": "DakeFolder_OverviewRename.exe",
  "version": "1.0.0",
  "release_url": "",
  "app_type": "market",
  "completion_goal": "formal_release",
  "screenshot_path": "assets/screenshot.webp",
  "status": "draft",
  "show_in_launcher": false,
  "show_on_site": false
}
```

## RELEASE_BODY

- PDF・画像の中身を見ながらフォルダ名を入力。
- フォルダの一括名前変更と検証付きUndo。
- フォルダ内ファイルは閲覧のみ。
- Windows向け単一exe。Human Review待ち・公開HOLD。

## 候補検証

41 tests passed。30フォルダ・406項目のGUI試験で入力保持・末尾到達を確認。
候補exeで正常PDF・画像・ページ数の実画面を確認済み。
実エクスプローラーdrop、日本語IME、候補exeでの連続整理はHuman Review待ち。
詳細・候補ハッシュはORIGINAL.mdの検証記録を参照してください。
