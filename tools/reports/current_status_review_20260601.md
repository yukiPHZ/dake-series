# Current Status Review 2026-06-01

調査日時: 2026-06-01 20:08 JST  
対象ルート: `C:\Users\yukiz\devlop`

## 調査方針

- 調査のみ。修正、commit、pushは実施していない。
- `git status`、`git log --oneline --since="2 days ago"`、直近commit、成果物の有無を確認した。
- `shimarisu-fudosan-site` というフォルダは見つからなかったため、実体候補として `shimarisu-site` を確認した。
- push済み判定は、主に `HEAD` と `origin/main` の一致で確認した。

## 全体状況

### 今いちばん進んでいる流れ

- DAKE正式出荷ラインが最も進んでいる。
  - `DAKE_series` では直近2日で Factory shipment / BOOTH ready / exe launch / dashboard / VersionInfo / site detection まわりが集中的に整備された。
  - 最新レポート `tools/reports/dake_factory_shipment_check.md` では、正式出荷対象45件がBOOTH素材3/3、Release未作成0、release_url未設定0、BOOTH登録可能45件になっている。
  - `tools/reports/exe_launch_check.md` では、チェック対象46件が OK 10 / OK_GUI 36 / 問題なし。
- SHIMARISUブランド導線も進んでいる。
  - `SHIMARISU` 本体にPack生成が追加済み。
  - `shimarisu-dakeapp-site` に `SHIMARISU_Pack.zip` が置かれ、ダウンロード導線が追加済み。
  - `dakeapp-site` と `peakheadz-site` 側にも、しまりすくん導線が追加・整理済み。

### 今止まっている流れ

- QPSC / BRAINZ / OIKAWAは「思想と内部UIは進んでいるが、正式出荷ラインにはまだ乗っていない」状態。
  - `DAKE_Brainz_Search`: status `draft`、dist exeあり、release_url空、screenshot/thumbnailなし。
  - `DAKE_Brainz_OIKAWA`: status `experimental`、dist exeあり、release_url空、screenshot/thumbnailなし。
  - `DAKE_Wake_Brainz`: status `draft`、dist exeあり、release_url空、screenshot/thumbnailなし。
  - `DAKE_QPSC_Dashboard`: status `internal`、dist exeあり、README/release_bodyあり。内部ツール扱いのためBOOTH素材なし。
- `pdf_to_jpeg_app` は未コミット差分があり、Redis/RQジョブキュー化の途中に見える。今回のDAKE/SHIMARISU本線とは別枠で要確認。

### 混線している流れ

- SHIMARISUの位置づけ。
  - サイト側では「実務判断レイヤー」として整理が進んでいる。
  - 一方で `SHIMARISU/README.md` にはまだ `SHIMARISU は DAKE シリーズの入口です。` が残っている。
  - `SHIMARISU/01_main_app/ShimarisuKun/main.py` のfooterにも `しまりす不動産` 文脈が残っている。ブランドサイト側の整理方針と合わせるなら要確認。
- 情報源管理。
  - `DAKE_series` に未追跡の `03_docs/DAKE_series_情報源/` がある。
  - `peakheadz-project-index` に `.tmp.driveupload/` と `.wrangler/` が未追跡で残っている。
  - どちらも今回勝手に整理せず、要確認扱い。

### 今日やらなくていいこと

- Cloudflare DNS / Workers削除 / Secrets変更。
- 旧サイトやWorkersの削除。
- BOOTHへ大量登録する作業。
- QPSC / BRAINZ / OIKAWAを無理に正式出荷へ上げる作業。
- `.wrangler/` や `.tmp.driveupload/` の削除。必要性確認後でよい。

### 今日やるなら優先すべきこと TOP5

1. `SHIMARISU/README.md` の古い位置づけ表現を、ブランドサイト側の最新整理に合わせるか確認する。
2. `DAKE_series` の未追跡 `03_docs/DAKE_series_情報源/` を、commit対象にするのか、`peakheadz-project-index` 側へ寄せるのか決める。
3. `dakeapp-site` の未追跡 `.wrangler/` を `.gitignore` 対象として扱うか確認する。
4. QPSC / BRAINZ / OIKAWAの次フェーズを「内部継続」か「スクショ・release_url・BOOTH素材整備へ進む」か決める。
5. `pdf_to_jpeg_app` のRedis/RQ化差分を続行するのか、いったん別作業として隔離するのか決める。

## repo別状況

### DAKE_series

- 状態: 未コミットあり。
- push済みか: `HEAD=e56b004`、`origin/main=e56b004` でpush済み。
- 未コミット差分:
  - 未追跡: `03_docs/DAKE_series_情報源/`
  - 内容はDAKE情報源フォルダで、`00_active/`、`90_archive/`、`99_temp/` を含む。
  - 主なファイル: 開発哲学、共通開発運用ルール、UI原則、ブランド定義、正式出荷ライン、禁止事項、GitHub READMEなど。
- 直近commit:
  - `e56b004 Add DAKE factory shipment product texts`
- 昨日〜今日の主な作業:
  - DAKE Factory shipment report生成。
  - 45件の正式出荷対象アプリの `booth_product.txt` 整備。
  - DAKE dashboard / web dashboard / BOOTH readiness / launch check / VersionInfo / site detection の強化。
  - `DAKE_Mail_AllStaff_v1.0.0`、`DAKE_Mail_Address_Format_v1.0.0` 周辺の出荷整備。
  - `DAKE_QPSC_Dashboard` 追加とブランド名調整。
- 完了していること:
  - `tools/reports/dake_factory_shipment_check.md`: 正式出荷対象45、BOOTH素材3/3が45、Release未作成0、release_url未設定0。
  - `tools/reports/exe_launch_check.md`: checked 46、OK 10、OK_GUI 36、Problems none。
  - `DAKE_BOOTH_Assist`、`DAKE_App_Doko`、`DAKE_Mail_AllStaff`、`DAKE_Mail_Address_Format` は README / release_body / booth_product / screenshot / dist / release_url が揃っている。
- 未完・気になること:
  - `03_docs/DAKE_series_情報源/` が未追跡。情報源として重要そうだが、Git管理するか要確認。
  - QPSC/BRAINZ/OIKAWA系は内部・draft・experimentalのまま。正式出荷ではない。
  - `DAKE_QPSC_Dashboard`: README/release_body/distあり、ただし internal扱いで release_url空、screenshot/booth_productなし。
  - `DAKE_Brainz_Search`: draft、distあり、booth_ready productあり、release_url空、screenshot/thumbnailなし。
  - `DAKE_Brainz_OIKAWA`: experimental、distあり、booth_ready productあり、release_url空、screenshot/thumbnailなし。
  - `DAKE_Wake_Brainz`: draft、distあり、booth_ready productあり、release_url空、screenshot/thumbnailなし。
- 次に見るべきこと:
  - 未追跡情報源フォルダの扱い。
  - QPSC系を内部継続するか、出荷準備へ進めるか。
  - 最新Factory shipment reportを正として、古いBOOTH ready reportとの差分をどう扱うか。

### SHIMARISU

- 状態: clean。
- push済みか: `HEAD=6bcb63c`、`origin/main=6bcb63c` でpush済み。
- 直近commit:
  - `6bcb63c Add SHIMARISU pack generation`
- 昨日〜今日の主な作業:
  - `tools/build_shimarisu_pack.py` 追加。
  - `build_pack.bat` 追加。
  - `01_main_app/ShimarisuKun/main.py` のPack向け調整。
- 完了していること:
  - `dist/SHIMARISU_Pack.zip`、`dist/SHIMARISU_latest_Pack.zip`、`dist/SHIMARISU_v1.0_Pack.zip` が存在。
  - `01_main_app/ShimarisuKun/dist/ShimarisuKun.exe` が存在。
  - release assetsとして screenshot / booth_thumbnail / step画像あり。
- 未完・気になること:
  - `README.md` に `SHIMARISU は DAKE シリーズの入口です。` が残っている。最近のブランド修正方針と矛盾の可能性あり。
  - アプリfooterに `しまりす不動産` 文脈が残っている。ブランドサイト側と統一するか要確認。
- 次に見るべきこと:
  - SHIMARISU READMEとアプリ内コピーの位置づけ整理。
  - Pack zipの中身確認と、ユーザー配布導線の最終確認。

### dakeapp-site

- 状態: 未コミットあり。
- push済みか: `HEAD=49a7a58`、`origin/main=49a7a58` でpush済み。
- 未コミット差分:
  - 未追跡: `.wrangler/`
  - 中身は `.wrangler/cache/pages.json` と `wrangler-account.json`。Cloudflareローカルキャッシュと思われる。
- 直近commit:
  - `49a7a58 Reposition shimarisu footer links`
- 昨日〜今日の主な作業:
  - SHIMARISU導線の位置調整。
  - `public/shimarisu/`、トップの special block、apps一覧の `しまりすくんを見る` 導線。
  - `DAKE_Mail_AllStaff`、`DAKE_Mail_Address_Format`、`DAKE_Mansion_Schedule` の掲載・文言整備。
- 完了していること:
  - 本番 `https://dakeapp.com/` はHTTP 200。
  - `public/index.html` に `https://shimarisu.dakeapp.com/` 導線あり。
  - `public/apps/mail-address-format/` と `public/apps/mail-allstaff/` にDownload導線あり。
- 未完・気になること:
  - `.wrangler/` をcommit対象にしない運用確認。
  - SHIMARISUをdakeapp側でどこまで扱うかは、ブランドサイトとの役割整理を継続確認。
- 次に見るべきこと:
  - `.gitignore` に `.wrangler/` があるか確認し、なければ別途整理候補。
  - Cloudflare Pagesの最新deploy反映確認。

### shimarisu-dakeapp-site

- 状態: clean。
- push済みか: `HEAD=2e36c7a`、`origin/main=2e36c7a` でpush済み。ただし `git status` 表示上は upstream tracking が出ていないため、branch設定は要確認。
- 直近commit:
  - `2e36c7a Add SHIMARISU pack download`
- 昨日〜今日の主な作業:
  - SHIMARISUブランドサイト作成。
  - Cloudflare Pages向けwrangler修正。
  - 「DAKE入口」から「実務判断レイヤー」へのPosition修正。
  - Pack download導線追加。
- 完了していること:
  - `public/downloads/SHIMARISU_Pack.zip` が存在。サイズは約199MB。
  - 本番 `https://shimarisu.dakeapp.com/` はHTTP 200。
  - READMEには、SHIMARISUとDAKEは上下ではなくparallel rolesと明記。
- 未完・気になること:
  - `SHIMARISU`本体README側には古い「入口」表現が残っているため、正本が分かれる可能性あり。
  - `git status`が `## main` でupstream表示なし。pushは済んでいるが、今後の運用では `git branch --set-upstream-to=origin/main main` が必要か要確認。
- 次に見るべきこと:
  - 本体READMEとサイトREADMEの位置づけ統一。
  - Pack downloadの中身・起動確認。

### shimarisu-site

- 状態: clean。
- push済みか: `HEAD=d540010`、`origin/main=d540010` でpush済み。
- 備考:
  - `shimarisu-fudosan-site` フォルダは見つからず、`shimarisu-site` を確認。
- 直近commit:
  - `d540010 Refine Shimarisu special block on DAKE page`
- 昨日〜今日の主な作業:
  - しまりす不動産サイトにSHIMARISUブランド導線を追加。
  - `public/dake.html` とCSSでspecial blockを調整。
- 完了していること:
  - `https://shimarisu-fudosan.com/dake` はHTTP 200。
  - `/dake.html` は `/dake` へ308 redirect。
- 未完・気になること:
  - SHIMARISUブランドサイトとしまりす不動産の導線が混ざりやすい。今後の表現は注意。
- 次に見るべきこと:
  - 不動産サイト側に残すべき導線と、ブランドサイトへ寄せる導線の切り分け。

### peakheadz-site

- 状態: clean。
- push済みか: `HEAD=646191e`、`origin/main=646191e` でpush済み。
- 直近commit:
  - `646191e Add SHIMARISU quietly to PEAKHEADZ`
- 昨日〜今日の主な作業:
  - PEAKHEADZトップに小さな `Now` セクションとして `しまりすくん` を追加。
  - スマホ幅で既存ナビが横へ出る問題を小さく調整。
- 完了していること:
  - 本番 `https://peakheadz.com/` はHTTP 200。
  - 本番HTMLに `https://shimarisu.dakeapp.com/` と `しまりすくん` が反映済み。
- 未完・気になること:
  - なし。現状は静かな導線追加として完了。
- 次に見るべきこと:
  - Cloudflare Deployments上のProduction deploy履歴を管理画面で確認できるなら確認。

### yukihikokikuta-site

- 状態: clean。
- push済みか: `HEAD=6c2fef4`、`origin/main=6c2fef4` でpush済み。
- 直近commit:
  - `6c2fef4 Remove top note trace link`
- 昨日〜今日の主な作業:
  - `git log --since="2 days ago"` では該当なし。
  - 直近commitではトップのnote trace linkを削除。
- 完了していること:
  - トップからnote項目を削除する整理はcommit済み。
- 未完・気になること:
  - 今回の調査範囲では未コミット差分なし。
- 次に見るべきこと:
  - note導線を下層・側の導線として残す方針が維持されているか、必要時に本番表示確認。

### peakheadz-project-index

- 状態: 未コミットあり。
- push済みか: `HEAD=8430fbf`、`origin/main=8430fbf` でpush済み。
- 未コミット差分:
  - 未追跡: `.tmp.driveupload/`
  - 未追跡: `.wrangler/`
- 直近commit:
  - `8430fbf Add OGP image operation rule`
- 昨日〜今日の主な作業:
  - `git log --since="2 days ago"` では該当なし。
  - 直近ではOGP画像運用ルールを追加。
- 完了していること:
  - favicon / OGP / Pages Functions / DAKE Webなどの運用ルール群は整備済み。
- 未完・気になること:
  - `.tmp.driveupload/` は5月12〜13日頃の一時アップロードファイルが多数残存。
  - `.wrangler/` はCloudflareローカルキャッシュ。
  - 削除・ignore整理は今回は未実施。
- 次に見るべきこと:
  - 一時フォルダをGit管理対象外として正式に扱うか、削除してよいか確認。

### pdf_to_jpeg_app

- 状態: 未コミットあり。
- push済みか: `HEAD=e46042c`、`origin/main=e46042c` でpush済み。未コミット差分は未push。
- 未コミット差分:
  - `app.py`
  - `requirements.txt`
- 直近commit:
  - `e46042c ワーカータイムアウト問題修正`
- 昨日〜今日の主な作業:
  - `git log --since="2 days ago"` では該当なし。
- 差分概要:
  - Flask内のThread/Queue処理から Redis + RQ job queue へ変更途中。
  - `/status/<job_id>` と `/download/<job_id>` が追加。
  - `requirements.txt` に `redis` と `rq` が追加され、`poppler-utils` が削除。
- 完了していること:
  - なし。未コミット差分なので作業途中扱い。
- 未完・気になること:
  - Render等でRedis設定が必要になる可能性がある。
  - フロント側が新しいjob APIに対応しているか未確認。
  - `poppler-utils` 削除が環境要件として正しいか要確認。
- 次に見るべきこと:
  - この変更を進めるのか、別作業として保留するのか判断。

## その他repo

以下は今回確認時点で clean、かつ `git log --since="2 days ago"` に該当なし。

- `dake-ai-site`
- `dake-gis-site`
- `dake-labs-site`
- `dake-tools-site`
- `holiday-blue-site`
- `holiday-jinja-site`
- `holiday-side-site`
- `holiday-sky-site`
- `japanmemorylane-site`
- `nicekip-restore`
- `nicekip-site`
- `niceskill-site`
- `soredake-site`
- `wlzphz-site`
- `yukizblog-restore`
- `yukizblog-site`

## サイト反映確認

- `https://peakheadz.com/`: HTTP 200。SHIMARISUリンク反映確認済み。
- `https://dakeapp.com/`: HTTP 200。
- `https://shimarisu.dakeapp.com/`: HTTP 200。
- `https://shimarisu-fudosan.com/dake`: HTTP 200。
- `https://shimarisu-fudosan.com/dake.html`: HTTP 308で `/dake` へredirect。

## 要確認リスト

1. `SHIMARISU/README.md` の「DAKEシリーズの入口」表現を更新するか。
2. `SHIMARISU`アプリfooterの `しまりす不動産` 表記をブランドサイト方針に合わせるか。
3. `DAKE_series/03_docs/DAKE_series_情報源/` をGit管理するか、project-indexへ移すか。
4. `dakeapp-site/.wrangler/`、`peakheadz-project-index/.wrangler/`、`.tmp.driveupload/` の扱い。
5. QPSC / BRAINZ / OIKAWAを内部継続するか、正式出荷準備へ進めるか。
6. `pdf_to_jpeg_app` の未コミット差分を続行するか保留するか。

## ChatGPT貼り付け用 要約版

昨日〜今日の主戦場は、DAKE正式出荷ラインとSHIMARISU導線です。

DAKE_seriesはpush済みだが、未追跡の `03_docs/DAKE_series_情報源/` が残っています。直近ではFactory shipment整備が進み、正式出荷対象45件がBOOTH素材3/3、Release未作成0、release_url未設定0になっています。exe launch checkも問題なしです。

SHIMARISUは本体・Pack生成・ブランドサイト・dakeapp.com導線・peakheadz.com導線まで進んでいます。`shimarisu.dakeapp.com` は200で、Pack zipもサイトに置かれています。ただし、SHIMARISU本体READMEにはまだ「DAKEシリーズの入口」という古い表現が残っており、ブランド整理と合わせるなら次に直す候補です。

QPSC / BRAINZ / OIKAWAは内部的には進んでいますが、正式出荷はまだです。BRAINZはdraft、OIKAWAはexperimental、Wake Brainzはdraftで、dist exeはありますがscreenshotやrelease_urlは未整備です。QPSC Dashboardはinternal扱いです。

サイト系では `peakheadz-site`、`dakeapp-site`、`shimarisu-dakeapp-site`、`shimarisu-site` がpush済み。`peakheadz.com`、`dakeapp.com`、`shimarisu.dakeapp.com`、`shimarisu-fudosan.com/dake` はHTTP 200確認済みです。

未コミットで注意が必要なのは、`DAKE_series` の情報源フォルダ、`dakeapp-site` の `.wrangler/`、`peakheadz-project-index` の `.tmp.driveupload/` と `.wrangler/`、そして `pdf_to_jpeg_app` のRedis/RQ化途中差分です。

今日やるなら、1. SHIMARISU READMEの位置づけ修正、2. DAKE情報源フォルダの扱い決定、3. `.wrangler/`等の一時ファイル整理方針決定、4. QPSC/BRAINZ/OIKAWAの次フェーズ決定、5. pdf_to_jpeg_app差分の扱い決定、の順がよいです。
