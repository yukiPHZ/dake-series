# 承認Brainz

## Frozen

このアプリは現在 frozen 扱いです。削除はしませんが、QPSC本線には含めません。

理由は、Codex公式のモバイル承認 / 遠隔操作導線が進んでおり、承認専用UIとしての役割が重複する可能性が高いためです。

QPSC本線では、承認はCodex公式導線を優先します。BRAINZ / OIKAWA は通知・記憶・検索・提案を担い、承認専用UIは現時点で積極開発しません。将来、公式Codexモバイル承認で代替できない用途が出た場合のみ再開します。

## QPSC v0.1での位置

DAKE_Approve_BrainzはQPSC v0.1では frozen として固定します。

- QPSC本線には含めません
- 新機能追加、UI改修、ビルド更新は行いません
- 承認導線はCodex公式モバイル承認 / 遠隔操作導線を優先します
- BRAINZ / OIKAWA / DAKE_Wake_Brainz の通知・記憶・検索・状態確認の流れとは分離します
- 公式導線で代替できない用途が出た場合のみ再検討します

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

## 凍結中の扱い

- 新機能追加はしません
- UI改修はしません
- QPSC本線の承認導線には使いません
- 必要な場合のみ既存メモとして参照します
- 公式Codexモバイル承認で代替できない用途が出た場合のみ再開します

## ビルド

```bat
build.bat
```

出力:

```text
dist\DakeApproveBrainz.exe
```

## Status

This app is currently frozen.

Saved as an experimental / exploratory DAKE project.
Not part of the current mainline workflow.

## DAKE_META

```json
{
  "app_key": "dake_approve_brainz",
  "display_name": "承認Brainz",
  "launcher_title": "承認Brainz",
  "launcher_description": "家PC上の承認待ちをスマホブラウザから確認します。",
  "site_title": "承認Brainz",
  "site_description": "承認だけを取り出す、BRAINZ連携前提の小さな承認パネルです。",
  "update_summary": "Codex公式モバイル承認導線を優先するため凍結",
  "folder_name": "DAKE_Approve_Brainz",
  "exe_name": "DakeApproveBrainz.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "frozen",
  "show_in_launcher": false,
  "show_on_site": false,
  "demo_video_path": "release_artifacts/demo.mp4",
  "demo_video_url": "",
  "social_release_path": "release_artifacts/social_release.json"
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
