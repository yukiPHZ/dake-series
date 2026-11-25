# ORIGINAL.md

## 正本宣言

このファイルは `DAKE_Voice_Log` の真の正本です。

README.md、README内DAKE_META、release_body.md、booth_product.txt、Store表示などは、このファイルから派生するビューです。

## 基本情報

- app_id: `dake_voice_log`
- app_name: Dake声ログ
- title: Dake声ログ
- internal_app_name: Dake声ログ
- display_name: 声ログ
- short_title: 声ログ
- folder_name: DAKE_Voice_Log
- exe_name: DakeVoiceLog.exe
- product_type: Windowsデスクトップアプリ
- category: 業務記録 / 音声ログ
- status: draft
- version: 0.1.0
- platform: Windows
- local_first: true
- transcription_engine: faster-whisper / CPU / int8 / ja
- summary_engine: Ollama / qwen3:1.7b / optional
- pricing_status: 未設定
- release_status: 未公開
- source_policy: ORIGINAL.mdを正本とし、README、DAKE_META、release_body、booth_productは派生ビューとして同期する

## 一言説明

PCのそばで交わされた声を、発話時だけ記録し、文字にして、一日ごとにまとめて残すローカルファーストの業務記録アプリです。

## 目的

毎回録音開始ボタンを押すのではなく、アプリ起動中に音声待機を続け、発話を検知した区間だけを安全に保存する。録音、文字起こし、日次統合、日次要約を日付単位で残し、あとから業務の流れを振り返れるようにする。

## 対象ユーザー

- 打ち合わせや日々の会話をあとから確認したい人
- 会議メモや業務メモの取りこぼしを減らしたい人
- 音声と文字起こしをローカルPC内に保存したい人
- DAKEシリーズの静かな実務ツールとして使いたい人

## 中心価値

録音そのものではなく、あとから振り返れること。

```text
聞こえた。
録れた。
文字になった。
あとから見つかった。
```

## 基本機能

- タスクトレイ常駐
- マイク入力監視
- 音声レベル検知
- 発話検知による自動録音開始
- 無音継続による録音終了
- pre-roll付き録音
- 音声ファイル保存
- 日付別保存
- SQLiteによる状態管理
- 未完了一時PCMの復旧
- 文字起こしキュー
- CPUローカル文字起こし
- 文字起こしテキスト保存
- daily_transcript.md生成
- Ollama任意連携
- 段階要約によるdaily_summary.md生成
- 未処理ジョブの再開
- マイク切断時のエラー状態維持と再接続試行
- 待機の一時停止 / 再開
- 保存フォルダを開く
- 今日のログを確認
- 設定画面

## 初版で実装しない機能

- Speaker_1 / Speaker_2 の話者分離
- 声紋による人物継続判定
- 男女推定
- 人物名自動付与
- 電話判定
- 会議判定
- 高度な全文検索
- クラウド同期
- スマートフォン連携
- リアルタイム字幕表示
- OpenAI API連携
- 録音の遠隔操作
- 隠し常駐
- トレイアイコンを隠す機能

## データ保存方針

ユーザー向け記録:

```text
Documents/
└─ Dake声ログ/
   ├─ 2026-07-08/
   │  ├─ audio/
   │  ├─ transcript/
   │  ├─ daily_transcript.md
   │  └─ daily_summary.md
```

アプリ内部状態:

```text
%LOCALAPPDATA%/
└─ DakeVoiceLog/
   ├─ state.db
   ├─ config.json
   ├─ logs/
   ├─ spool/
   └─ models/
```

## 録音方式

- sounddeviceのInputStreamで16kHz / mono / int16を取得
- コールバックは短く保ち、音声チャンクをqueueへ渡す
- 録音中は一時PCMへ追記
- 録音終了時にWAVへfinalize
- finalize成功後に一時PCMとメタ情報を削除
- 起動時にspool内の未完了一時PCMをWAVへ復旧

## VAD方式

v0.1は軽量VADを採用する。

- 音声エネルギー
- 適応ノイズフロア
- ヒステリシス
- 最小発話時間
- 無音継続時間
- pre-rollリングバッファ

初期値:

- pre-roll: 1.5秒
- 無音終了: 30秒
- 長時間ローテーション: 30分

## 文字起こし方針

- faster-whisperを第一実装とする
- CPU-onlyを前提にする
- device: cpu
- compute_type: int8
- language: ja
- 標準: small
- 軽量: base
- モデルはexeに同梱しない
- モデル未導入でも録音を止めない
- 文字起こし失敗時も音声ファイルを消さない

## 要約方針

- Ollamaは任意機能
- Ollama未設定でも録音、保存、文字起こし、daily_transcript.md生成は利用可能
- API到達性は `localhost:11434` を確認
- モデル一覧は `/api/tags`
- モデル取得は `/api/pull`
- 初期候補モデルは `qwen3:1.7b`
- 長い一日はチャンク分割し、中間要約から統合要約を作る
- 要約プロンプトでは推測禁止、元時刻保持、発言者特定禁止を明示する

## 状態設計

録音状態、文字起こし状態、要約状態は分離する。

Capture State:

- INITIALIZING
- WAITING
- RECORDING
- PAUSED
- MIC_ERROR
- STOPPED

Transcription State:

- IDLE
- QUEUED
- PROCESSING
- ERROR

Summary State:

- DISABLED
- WAITING
- PROCESSING
- COMPLETE
- ERROR

## UI方針

- DAKE共通テーマを使う
- 白、ライトグレー基調
- BIZ UDPGothic優先
- 共通アイコンを使う
- UI文言は `APP_NAME`、`WINDOW_TITLE`、`COPYRIGHT`、`UI_TEXT` に集約する
- 音声待機状態とバックグラウンド処理状態を混同しない
- 録音状況を隠さない
- ウィンドウを閉じてもトレイで音声待機を継続する

## DAKE_META生成用情報

```json
{
  "app_key": "dake_voice_log",
  "display_name": "Dake声ログ",
  "launcher_title": "声ログ",
  "launcher_description": "声を、あとから振り返れる記録に。",
  "site_title": "Dake声ログ",
  "site_description": "PCのそばで交わされた声を発話時だけ記録し、文字起こしと日次ログにまとめるローカルファーストの業務記録アプリです。",
  "update_summary": "常時待機と音声検知による録音、文字起こし、日次統合、任意のOllama要約に対応した初版。",
  "folder_name": "DAKE_Voice_Log",
  "exe_name": "DakeVoiceLog.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "draft",
  "show_in_launcher": false,
  "show_on_site": false,
  "app_type": "market",
  "completion_goal": "formal_release",
  "payment_status": "unset"
}
```

## release_body生成用情報

- 発話検知による自動録音
- pre-roll付きで冒頭欠けを抑制
- 日付別にWAV、文字起こし、daily_transcript.mdを保存
- faster-whisper CPU文字起こしキュー
- Ollama任意連携による日次要約
- タスクトレイ常駐、一時停止、再開、完全終了
- Windows向けexe

## booth_product生成用情報

- 商品名: Dake声ログ
- 価格案: 未設定
- 商品紹介文: PCのそばで交わされた声を、発話時だけ記録し、文字起こしと日次ログにまとめます。
- タグ:
  - Windows
  - 音声記録
  - 業務記録
  - 文字起こし
  - ローカルAI
  - 会議メモ
  - 仕事効率化
  - DAKE
- 商品画像: 未設定
- 補助画像: 未設定
- 作品ファイル: 未設定
- GitHub Release: 未設定
- BOOTH URL: 未設定

## Codex作業時の注意

- 触ってよい: DAKE_Voice_Log配下
- 触らない: 既存アプリ、Store/Stripe、BOOTH公開設定、GitHub Release公開設定
- 外部公開しない: 音声、文字起こし、要約、個人情報、APIキー
- 自動操作しない: BOOTH公開、Store決済、GitHub Release更新、SNS投稿

