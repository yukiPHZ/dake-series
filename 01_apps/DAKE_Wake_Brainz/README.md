# 補助脳 起こす

Wake-on-LANで、家PCの補助脳基地を起動する小さなDAKEアプリです。

このアプリは補助脳BRAINZ本体ではありません。

役割は、BRAINZ PCへWake-on-LAN magic packetを送ることと、簡易的な起動状態を確認することだけです。

## できること

- Wake-on-LAN magic packet送信
- PC Nameへの簡易ping確認
- MAC Address保存
- Broadcast IP設定
- Port設定
- 静かな接続ログ
- PEAKHEADZロゴ表示

## 設定するもの

- PC Name: 起動確認でpingする名前またはIP
- MAC Address: BRAINZ PCの有線LANまたはWake対応NICのMAC Address
- Broadcast IP: 通常は `255.255.255.255`
- Port: 通常は `9`

MAC Address例:

```text
AA:BB:CC:DD:EE:FF
```

## Wake-on-LANの準備

BRAINZ PC側で、BIOS / UEFI と Windows のWake-on-LAN設定を有効にしてください。

一般的には以下を確認します。

- BIOS / UEFIでWake-on-LANを有効化
- Windowsのデバイスマネージャーでネットワークアダプターの電源管理を確認
- マジックパケットでの起動を許可
- 可能なら有線LANで接続
- 同一LAN内で使う

## Phase 1でやらないこと

- VPN越え
- 外部ネットワーク対応
- スマホアプリ化
- BRAINZ本体操作
- Remote Queue送信
- SSH接続
- Remote Desktop接続
- 自動起動ループ

## ビルド

```bat
build.bat
```

出力:

```text
dist/DakeWake_Brainz.exe
```

## DAKE_META

```json
{
  "app_key": "DAKE_Wake_Brainz",
  "display_name": "補助脳 起こす",
  "launcher_title": "補助脳 起こす",
  "launcher_description": "Wake-on-LANで家PCの補助脳基地を起動する小さな道具です。",
  "site_title": "補助脳 起こす",
  "site_description": "Wake-on-LANで補助脳BRAINZ PCを起動するDAKEアプリです。",
  "update_summary": "Wake-on-LAN起動ツールを追加しました。",
  "folder_name": "DAKE_Wake_Brainz",
  "exe_name": "DakeWake_Brainz.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "draft",
  "show_in_launcher": false,
  "show_on_site": false
}
```
