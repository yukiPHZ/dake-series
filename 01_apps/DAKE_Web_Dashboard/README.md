# DAKE Web Dashboard

DAKE Web Dashboard は、菊田さん本人の開発アシスト専用に作成した内部アプリです。

一般公開、BOOTH販売、dakeapp.com掲載、GitHub Release作成は行いません。

## 役割

- DAKE / QPSC / PEAKHEADZ / BORINEF / holiday-jinja / JapanMemoryLane / SHIMARISU などのサイト群を一覧で見るサイト台帳
- サイト名、ドメイン / URL、カテゴリ、公開状態、説明、フォルダ、最終更新の確認
- Notionや記憶頼みの管理から、リポジトリ内のREADME / DAKE_WEB_METAを読む運用への移行補助
- Cloudflare Pages / Functions / Git状態 / OpenAI API構成は、詳細ペインで確認する補助情報

## 読み取り対象

```text
C:\Users\yukiz\devlop
```

直下および1〜2階層程度から、`*-site`、`wrangler.toml`、`package.json`、`public/`、`functions/`、`_routes.json`、`.git/` などを手がかりにサイト候補を検出します。

各サイトフォルダでは以下を読み取ります。

- README.md
- DAKE_WEB_META
- wrangler.toml
- package.json
- public/_routes.json または _routes.json
- functions/
- public/
- Git状態

`node_modules`、`.git`、`.wrangler`、`dist`、`build`、`__pycache__`、`venv`、`.venv` は深く読みません。

## 判定内容

- サイト名、説明、カテゴリ、domain / production_url の台帳表示
- 公開中 / 制作中 / URL不明 / 要整理 / 休止・凍結の分類
- README.md と DAKE_WEB_META の有無
- domain / production_url / README内URL / package.json homepage / known domain map からのURL推定
- Gitリポジトリ、Cloudflare Pages、Functions、OpenAI API構成の確認補助

APIキーらしき値は表示しません。`sk-` らしき文字列を見つけた場合も全文は出さず、直書き疑いとして表示します。

## 状態分類

- 公開中
- 制作中
- URL不明
- 要整理
- 休止 / 凍結
- 候補

APIあり、Git未コミット、Functions構成などは主状態ではなく、補助ラベルや詳細ペインで扱います。

## Phase3 サイト台帳モード

Phase3では、画面の主役を監査情報からサイト一覧・サイト台帳へ変更しています。

- 上部カードは、総サイト数、公開中、制作中、休止 / 凍結、URL不明、要整理を表示
- 一覧列は、状態、サイト名、ドメイン / URL、カテゴリ、説明、フォルダ、最終更新、整理メモを優先
- Git / API / Functions / Cloudflare構成は補助情報として右側列または詳細ペインへ移動
- domain / production_url は DAKE_WEB_META、README、package.json、known domain map の順で推定
- GitHub URL、Cloudflare管理画面URL、google.com系URLは本番URLとして誤検出しない
- カテゴリは DAKE、SHIMARISU、PEAKHEADZ、BORINEF、holiday、QPSC、blog、note、other に分類
- 「Codex指示をコピー」で、選択中サイトのREADME / DAKE_WEB_META整備指示をクリップボードへコピー
- 「META候補をコピー」で、選択中サイト用の DAKE_WEB_META JSON 候補をクリップボードへコピー
- DAKE_WEB_META未整備は強い警告ではなく、補助情報を整える候補として扱う

## Phase2.5 判定透明化

Phase2.5 では、状態名だけでなく「なぜその判定になったか」を一覧と詳細ペインで確認できるようにしています。

- 判定理由表示
- サイト検出スコア
- `site` / `candidate` / `ignored_component` の分類
- `public` / `functions` / `sitemap` の誤検出抑制
- QPSC通知カードの理由表示
- 次にやる候補の理由表示

通常の「全部」フィルタには `site` のみを表示し、`candidate` は「候補」フィルタで確認します。`ignored_component` は一覧に出さず、サイト本体の構成要素として扱います。

## 操作

- 再読み込み
- 30秒ごとの自動更新
- watchdog による変更監視
- 監視開始 / 停止
- サイトフォルダを開く
- README.md を開く
- production_url をブラウザで開く
- health_url をブラウザで開く
- README等から取得できたGitHub URLを開く
- 選択中サイト用のCodex指示をコピー
- 選択中サイト用のDAKE_WEB_META候補をコピー
- フィルタ切替
- 検索

`watchdog` が未導入でもアプリは落ちず、30秒自動更新だけで動きます。

## 内部運用

- 一般リリースしません。
- BOOTH対象外です。
- dakeapp.com掲載対象外です。
- GitHub Release作成対象外です。
- Cloudflare / OpenAI API状態の確認補助であり、外部サービスへデプロイや反映は行いません。

## 品質チェック

```powershell
python -m py_compile main.py
python main.py --launch-check
.\build.bat
.\dist\DakeWeb_Dashboard.exe --launch-check
```

## Positioning

This app is a system/operations dashboard, not a market-facing standalone product. Its completion goal is `system_ready`: README, launch-check, build, and dashboard role are documented and working.

## DAKE_META
```json
{
  "app_key": "DAKE_Web_Dashboard",
  "display_name": "DAKE Web Dashboard",
  "launcher_title": "DAKE Web Dashboard",
  "launcher_description": "Internal dashboard for DAKE related sites",
  "site_title": "",
  "site_description": "",
  "update_summary": "Internal tool to review site status from README metadata",
  "folder_name": "DAKE_Web_Dashboard",
  "exe_name": "DakeWeb_Dashboard.exe",
  "release_url": "",
  "screenshot_path": "",
  "status": "internal",
  "app_type": "system",
  "completion_goal": "system_ready",
  "show_in_launcher": false,
  "show_on_site": false
}
```
