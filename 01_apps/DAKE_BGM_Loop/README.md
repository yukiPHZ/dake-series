# Dake BGM Loop

YouTube作業動画の裏に敷ける、短い“のんきなループBGM”をローカル生成するDAKEアプリです。

完成された曲を作るアプリではありません。15秒 / 30秒 / 60秒の静かなループ音を大量生成し、良かったものだけ保存・お気に入り化するための補助ツールです。

DAKE思想：のんきなループを、静かに作る。

## できること

- 空気を選んで英語プロンプトを内部生成
- 15秒 / 30秒 / 60秒の短いWAVループを生成
- seedを指定、または空欄で自動生成
- 生成後に使用seedを表示
- Preview / Stopで確認
- WAV保存
- お気に入りWAVを `favorites/` へコピー
- 生成metadataを `metadata/` へJSON保存
- 出力フォルダを開く

## Phase 1

Phase 1は、mock生成 + ACE-Step連携準備です。

今回は自前の本格的な音楽生成ロジックを作りません。外部AI音楽生成モデルを呼び出す前提の薄いDAKE UIとして作っています。

実生成部分は `core/generator_adapter.py` のadapter構造で差し替えできます。

- `BaseGeneratorAdapter`
- `MockGeneratorAdapter`
- `AceStepGeneratorAdapter`

`MockGeneratorAdapter` はACE-Step未導入でもアプリの起動・操作確認ができるように、簡単なpad風 / クリック風 / 低音風のWAVを生成します。無音ではありませんが、完成音楽として作り込むものではありません。

`AceStepGeneratorAdapter` はACE-Stepが利用可能な環境だけで呼び出します。未設定の場合、UIには `ACE-Step is not configured. Mock mode only.` と表示され、mock生成で確認できます。

## 空気

- のんき
- 静か
- 作業用
- 神社
- 雨
- 夜
- ミシン
- コード
- 余白

## 出力

生成物はアプリフォルダ内に保存されます。

```text
outputs/
favorites/
metadata/
```

ファイル名例：

```text
DAKE_BGM_Loop_20260515_213015_nonkina_30s_seed418220.wav
```

metadata JSONには以下を保存します。

- `created_at`
- `mood`
- `duration_sec`
- `seed`
- `prompt`
- `model_adapter`
- `output_path`
- `favorite`
- `license_notice`

## 設定

`settings.json` に以下を保存します。

- `generator_mode`: `mock` or `ace_step`
- `output_dir`
- `favorite_dir`
- `last_mood`
- `last_duration_sec`

## 使い方

```powershell
python main.py
```

起動後、空気・長さ・seedを選び、`Generate` を押します。seed欄が空の場合は自動生成されます。

## 注意

- DAW化しません。
- 波形編集UI、MIDI編集、トラック編集は作りません。
- Apple Loops等の第三者素材は同梱しません。
- 生成物の商用利用可否は使用モデルのライセンスに依存します。
- YouTube等で公開する前に、使用モデルのライセンス確認が必要です。
- 商用利用可、または権利上の安全性をこのアプリ単体では断定しません。

## DAKE_META

```json
{
  "app_key": "DAKE_BGM_Loop",
  "display_name": "Dake BGM Loop",
  "launcher_title": "BGM Loop",
  "launcher_description": "のんきなループを、静かに作る。",
  "site_title": "Dake BGM Loop",
  "site_description": "YouTube作業動画の裏に敷く、短いループBGM生成補助ツール。",
  "update_summary": "短いBGMループ生成アプリを追加しました。",
  "folder_name": "DAKE_BGM_Loop",
  "exe_name": "DakeBGM_Loop.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "active",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

- 短いBGMループ生成補助ツール
- Phase 1はmock生成 + ACE-Step連携準備
- seed、metadata、お気に入りコピーに対応
- 生成音声の商用利用可否は使用モデル・素材・公開先の規約に依存

