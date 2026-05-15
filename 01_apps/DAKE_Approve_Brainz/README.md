# 承認Brainz

承認Brainzは、家PC上の承認待ちをスマホブラウザから確認し、承認 / 却下の結果を保存する小さなDAKEアプリです。

## アプリ概要

- 家PCでローカルWebサーバーを起動します。
- 同一LAN内のスマホから承認待ちを確認できます。
- `data/pending.json` を読み込み、判断結果を `data/result.json` に保存します。
- `pending.json` がない場合は、起動時にサンプル承認待ちを自動生成します。

## v0.1の目的

v0.1では、承認キュー、Web表示、結果保存の型だけを作ります。
Codex本体のapprovalには、まだ直結していません。

## 起動方法

```bat
python main.py
```

ビルド後は以下を起動します。

```bat
dist\DakeApproveBrainz.exe
```

## スマホからのアクセス方法

家PCでアプリを起動したあと、同一LAN内のスマホブラウザで以下を開きます。

```text
http://家PCのIP:8765
```

PC上で確認する場合は以下を開きます。

```text
http://127.0.0.1:8765
```

## v0.1でまだやらないこと

- Codex本体approvalとの自動接続
- Cloudflare Tunnel
- 外部公開URL
- ログイン機能
- WebSocket
- Slack / Discord連携

## 今後の拡張予定

- BRAINZ側の作業報告との橋渡し
- Codex approvalとの安全な連携方式の検討
- 複数承認待ちへの対応
- 承認履歴の保存
- LAN内運用に必要な最小設定の追加

## ビルド

```bat
build.bat
```

出力:

```text
dist\DakeApproveBrainz.exe
```

## DAKE_META

```json
{
  "app_key": "dake_approve_brainz",
  "display_name": "承認Brainz",
  "launcher_title": "承認Brainz",
  "launcher_description": "家PC上の承認待ちをスマホブラウザから確認します。",
  "site_title": "承認Brainz",
  "site_description": "承認だけを取り出す、BRAINZ連携前提の小さな承認パネルです。",
  "update_summary": "スマホブラウザから承認 / 却下できる最小Webパネルを追加しました。",
  "folder_name": "DAKE_Approve_Brainz",
  "exe_name": "DakeApproveBrainz.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

承認Brainz

- スマホブラウザから承認待ちを確認
- 承認 / 却下をresult.jsonへ保存
- BRAINZ連携前提の最小承認パネル
- Windows向けexe

## セキュリティ方針 v0.1

LAN内利用だけを想定します。外部公開はしません。
