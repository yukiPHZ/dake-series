# DAKE_Wake_Brainz

LIVA Z 常時ON機で動かす、3070Ti メインPC用の小さな補助脳 GATE です。Android スマホのブラウザから LAN/Tailscale 経由でアクセスし、Phase 1 では Wake on LAN と状態確認だけを行います。

## QPSCでの役割

QPSC（Quiet Personal Cognitive System）では、DAKE_Wake_Brainzを起床・遠隔起動・状態確認レイヤーとして扱います。

- BRAINZ = 深層記憶・取り込み・index・設定
- OIKAWA = 検索・原本表示・熾火・熱提案・通知
- DAKE_Wake_Brainz = 起床・遠隔起動・状態確認レイヤー
- Ollama = ローカル読解補助
- OpenClaw = ローカルエージェント / Gateway候補

Phase 1ではBRAINZ本体を操作せず、PC起床と状態確認を担います。BRAINZの起床状態はBRAINZ側のローカル状態ファイル、前面通知はOIKAWA側へ寄せます。

## QPSC v0.1 正本

DAKE_Wake_BrainzはQPSC v0.1で、起床 / 状態確認レイヤーとして固定します。

- LIVA Z 常時ON機から3070Ti PCをWake on LANで起こします
- LAN / Tailscale上のブラウザから状態を確認します
- BRAINZの保存・取り込み・検索・原本表示には踏み込みません
- BRAINZ awake状態はBRAINZ側の `qpsc_brainz_status.json` を正とします
- OIKAWAが前面UI、BRAINZが記憶庫、Wakeが起床確認という分担を保ちます

WakeはQPSCの判断層ではありません。起こす、見る、止めないための入口です。

## QPSC v0.2 現在地

DAKE_Wake_Brainzはv0.2時点でも、起床 / 状態確認レイヤーです。BRAINZやOIKAWAの前面UIではありません。

- LIVA Zやスマホ起点で、3070Ti PCの起床と状態確認を担います
- BRAINZの取り込み、保存、検索、原本表示には踏み込みません
- OIKAWAの通知、熾火、ORBIT、巡回、側に在る表示には踏み込みません
- 将来、スマホ起点のQPSC起動導線として整理する候補です

## 起動

```bat
cd C:\Users\yukiz\devlop\DAKE_series\01_apps\DAKE_Wake_Brainz
pip install -r requirements.txt
python main.py
```

アクセス:

```text
http://localhost:8766
http://LIVAZ_IP:8766
```

`main.py` は標準で `0.0.0.0` に bind します。LAN 内の Android Chrome からは、LIVA Z の IP アドレスを使ってアクセスしてください。

## config.json

`config.example.json` を `config.json` にコピーして、3070Ti PC に合わせて編集します。`config.json` はローカル設定なので Git 管理外です。

```json
{
  "target_mac": "AA:BB:CC:DD:EE:FF",
  "target_ip": "192.168.1.100",
  "web_port": 8766,
  "broadcast_ip": "255.255.255.255",
  "wake_port": 9,
  "ping_timeout_ms": 1000
}
```

項目:

- `target_mac`: Wake する 3070Ti PC の有線 LAN または Wake 対応 NIC の MAC アドレス
- `target_ip`: 状態確認で ping する 3070Ti PC の IP アドレス
- `web_port`: Web UI のポート。既定は `8766`
- `broadcast_ip`: Magic packet の送信先。通常は `255.255.255.255`
- `wake_port`: Wake on LAN の UDP ポート。通常は `9`
- `ping_timeout_ms`: ping 待機時間。既定は `1000`

`config.json` が未作成、壊れている、または MAC/IP が未設定でもアプリは落ちません。UI に設定不足を表示し、可能な範囲で状態確認を続けます。

## Wake on LAN 注意

3070Ti PC 側で Wake on LAN を有効にしてください。

BIOS / UEFI:

- Wake on LAN、Power On By PCI-E、Resume By PCI-E などの項目を有効化
- シャットダウン後も NIC に待機電力が供給される設定にする
- 高速起動や ErP/省電力設定が WOL を妨げる場合は無効化を検討

Windows 側 NIC 設定:

- デバイス マネージャーで対象ネットワークアダプターを開く
- 電源の管理で「このデバイスで、コンピューターのスタンバイ状態を解除できるようにする」を有効化
- 詳細設定で Wake on Magic Packet / Wake on pattern match を有効化
- 可能なら有線 LAN を使い、固定 IP または DHCP 予約で `target_ip` を安定させる

## Web UI

表示内容:

- `DAKE_Wake_Brainz`
- `補助脳 GATE`
- `3070Ti` カード
- 状態: `ONLINE` / `OFFLINE`
- ボタン: `Wake` / `Refresh`

黒ベース、静かな青アクセント、スマホ縦表示向けの 1 カード構成です。横スクロールが出ないように幅を制限しています。

## API

```text
GET  /api/status
POST /api/wake
```

`/api/status` は `target_ip` に ping を実行して `ONLINE` / `OFFLINE` を返します。`/api/wake` は `target_mac` に magic packet を送ります。

## build

PyInstaller で EXE 化します。

```bat
build.bat
```

出力:

```text
dist\DAKE_Wake_Brainz.exe
```

## Phase 1 でやること

- Wake on LAN
- Ping 状態確認
- Android Chrome 向け Web UI
- `config.json` 未設定時の安全なエラー表示

Sleep、Queue、Codex 監視、Slack bridge は将来フェーズで追加します。

## DAKE_META

```json
{
  "app_key": "DAKE_Wake_Brainz",
  "display_name": "DAKE_Wake_Brainz",
  "launcher_title": "Wake Brainz",
  "launcher_description": "LIVA Z から 3070Ti PC を Wake / 状態確認する補助脳 GATE。",
  "site_title": "DAKE_Wake_Brainz",
  "site_description": "Android ブラウザから 3070Ti PC を Wake / 状態確認する LAN 用 DAKE GATE。",
  "update_summary": "QPSC v0.1として起床 / 状態確認レイヤーの役割を正本化しました。",
  "folder_name": "DAKE_Wake_Brainz",
  "exe_name": "DAKE_Wake_Brainz.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "draft",
  "show_in_launcher": false,
  "show_on_site": false
}
```
