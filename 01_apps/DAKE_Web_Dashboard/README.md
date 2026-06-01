# DAKE Web Dashboard

DAKE Web Dashboard は、菊田さん本人の開発アシスト専用に作成した内部アプリです。

一般公開、BOOTH販売、dakeapp.com掲載、GitHub Release作成は行いません。

## 役割

- DAKE / QPSC / PEAKHEADZ / BORINEF / holiday-jinja / JapanMemoryLane / SHIMARISU などのサイト群の開発管制塔
- README正本可視化アプリ
- Cloudflare Pages / Functions / 独自ドメイン / Git状態 / OpenAI API構成の確認補助
- Notionや記憶頼みの管理から、リポジトリ内の正本情報を読む運用への移行補助

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

- README.md と DAKE_WEB_META の有無
- Gitリポジトリ、ブランチ、最新コミット、未コミット変更、未追跡ファイル、push / pull 待ち疑い
- Cloudflare Pagesらしさ、pages_build_output_dir、functions/api、/api/health、_routes.json の /api/* include
- OpenAI API関連の安全性確認
- production_url / domain / cloudflare_project の正本記載状況

APIキーらしき値は表示しません。`sk-` らしき文字列を見つけた場合も全文は出さず、直書き疑いとして表示します。

## 状態分類

- 正常
- 要確認
- API確認
- デプロイ確認
- 内部 / 凍結

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
  "show_in_launcher": false,
  "show_on_site": false
}
```
