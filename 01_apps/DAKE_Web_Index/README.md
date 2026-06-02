# DAKE Web Index

DAKE Web Index は、サイト群を俯瞰するための静かな索引です。

これは管理ツールではありません。指示しません。提案しません。優先順位を付けません。状況だけを置きます。

見る。気付く。決める。決めるのはユーザーです。

## 役割

- サイト群を一覧表示する
- README正本から自動生成する
- Notion代替として使う
- 手動登録を行わない

## 表示項目

- サイト名
- URL
- GitHub
- README
- 最終更新日

## 表示しないもの

- 状態監視
- KPI
- 通知一覧
- 優先順位
- おすすめ
- 次にやること
- 進捗率
- スコア

## データ取得

既定では以下を読み取ります。

```text
C:\Users\yukiz\devlop
```

各サイトフォルダの `README.md` を正本として読み、`DAKE_WEB_META` を収集します。`DAKE_WEB_META` は README 内の `## DAKE_WEB_META` セクション、または同名ファイルから読み取れます。

取得する値は、サイト名、URL、GitHub URL、README パス、最終更新日のみです。

## UI方針

上部にタイトルと説明を置き、中央に一覧テーブルを置き、下部に再読込ボタンを置きます。

余計なカード、集計、通知、色分けは入れません。

## QPSC連携

QPSC から起動できるよう、通常起動と `--launch-check` に対応します。

Codex 更新後の再スキャンや QPSC からの再読込用に、JSON を返す CLI を用意します。

```powershell
python main.py --launch-check
python main.py --reload-api
python main.py --qpsc-reload
```

`--reload-api` と `--qpsc-reload` は、現在の README 正本を再スキャンして JSON を出力します。おすすめ表示、次にやること表示、優先順位表示、通知カード生成は行いません。

## 操作

- 再読込
- 5分ごとの裏側自動更新
- URL列をダブルクリックしてURLを開く
- GitHub列をダブルクリックしてGitHubを開く
- README列をダブルクリックしてREADMEを開く

GUI起動時と再読込時に GitHub URL 補完用の外部コマンドを呼ぶ場合は、Windowsでコンソールを表示しないように実行します。

## 品質チェック

```powershell
python -m py_compile main.py
python main.py --launch-check
python main.py --reload-api
python main.py --open-check
.\build.bat
.\dist\DakeWeb_Index.exe --launch-check
```

## 内部運用

一般公開、BOOTH販売、dakeapp.com掲載、GitHub Release作成は行いません。

## DAKE_META

```json
{
  "app_key": "DAKE_Web_Index",
  "display_name": "DAKE Web Index",
  "launcher_title": "DAKE Web Index",
  "launcher_description": "README正本からWebサイト索引を自動生成します。",
  "site_title": "",
  "site_description": "",
  "update_summary": "GitHub URL補完の外部コマンドを非表示化し、5分ごとの裏側自動更新を追加しました。",
  "folder_name": "DAKE_Web_Index",
  "exe_name": "DakeWeb_Index.exe",
  "release_url": "",
  "screenshot_path": "",
  "status": "internal",
  "show_in_launcher": true,
  "show_on_site": false
}
```
