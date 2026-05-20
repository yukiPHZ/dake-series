# Dakeユキズ稼働中 v0.1.0

動画素材を取り込み、補助脳で投稿準備を整える制作コンソールです。

## できること

- 動画ファイルの取り込み
- FFmpeg / yt-dlp / Ollama / GitHub CLI / Wrangler の環境チェック
- NVENC利用可否の確認
- ffprobeによる動画情報取得
- faster-whisper利用時の文字起こし
- Shorts候補JSONの作成
- プレビュークリップ作成
- タイトル案・説明欄雛形の出力
- 処理ログ保存
- 出力フォルダを開く

## まだやらないこと

- YouTubeへの自動投稿
- 自動公開
- 完全自動編集
- PSD編集
- BGM自動ミックス
- 縦動画の高度な自動クロップ

## Phase 1.5

- Run System Check を追加
- GitHub CLI / Wrangler の認証状態を確認
- Ollama localhost API とモデル一覧を確認
- h264_nvenc / hevc_nvenc と nvidia-smi GPU表示を確認
- 不足がある場合に `data/outputs/system_check/install_guide.txt` を生成

## Phase 1.6

- CLI導入補助モードを追加
- Open Install Guide / Copy Install Commands / Recheck System を追加
- 不足CLIだけの導入候補コマンドをコピー可能に変更
- Wrangler用にNode.js / npm確認を追加
- 自動インストールしない安全方針を明記

## Phase 1.7

- 初回動画テスト導線を追加
- ffprobeによる動画情報取得と `media_info.json` 保存に対応
- 先頭10秒の `test_clip.mp4` 生成に対応
- NVENC優先、失敗時CPU `libx264` フォールバックに対応
- `data/outputs/first_video_test/logs/test_log.txt` に結果を保存

## Phase 1.8

- 投稿パッケージ生成導線を追加
- `data/outputs/packages/...` に投稿準備ファイル一式を出力
- 文字起こし不可時も `transcript_unavailable.txt` で継続
- Shorts候補、タイトル案、説明欄、タグ、確認メモを生成
- Ollama不可時は固定テンプレートで継続
- YouTube自動投稿と元動画変更は行わない

## Phase 2.0

- 補助脳レビュー導線を追加
- 投稿パッケージから `assistant_review.md` を生成
- Ollama利用時はローカルLLMでレビュー生成を試行
- Ollama不可時もテンプレートで継続
- 補助脳は提案のみ行い、最終判断はユーザーが行う方針を明記

## Phase 2.1

- 採用候補整理の `SELECTED OUTPUTS` を追加
- Shorts候補とタイトル候補を選択可能に変更
- `selected/` に選択ドラフト一式を出力
- `selected_summary.md` を生成
- 未選択時は #1 を仮採用し、自動投稿は行わない

## Phase 2.2

- `selected_short.json` からShortsプレビュー生成を追加
- `selected/short_preview.mp4` と `short_preview_log.txt` を出力
- NVENC優先、失敗時CPU `libx264` フォールバックに対応
- 元比率のまま切り出し、9:16クロップは今後対応
- 元動画変更とYouTube自動投稿は行わない

## Phase 2.3

- 9:16 Shorts Export を追加
- `selected/short_vertical_1080x1920.mp4` と `short_vertical_log.txt` を出力
- 背景ぼかし + 前景中央配置で 1080x1920 化
- NVENC 優先、失敗時 CPU `libx264` フォールバックに対応
- 音声 AAC 変換、自動投稿なし

## Phase 2.4

- DAKE_Music_Otooku の Project Box 読み取りを追加
- Project Box の Preset / Suggested Use / BGM 一覧を表示
- BGM Preview Start / Stop Preview を追加
- 選択BGMを `selected/bgm/` へ上書きなしでコピー
- `selected/upload/metadata_draft.txt` を生成
- Ollama READY時は補助脳提案を追記、失敗時はテンプレートで継続

## Phase 3.0

- 補助脳メモリを追加
- `data/memory/memory_index.json` へ制作履歴を追記
- `data/memory/memory_summary.md` を生成
- packageごとの `projects/*.json` / `projects/*.md` を保存
- 投稿package、selected、補助脳レビュー、Project Bridge結果を記録
- Ollama READY時は過去制作傾向の要約を試し、失敗時はテンプレートで継続

## Phase 3.1

- 補助脳リコメンドを追加
- `data/memory/` の履歴からBGM / Preset / タイトル / Shorts Directionを簡易解析
- 現在package向けに `assistant_recommendation.md` を生成
- Ollama READY時は推薦生成を試し、失敗時はテンプレートで継続
- 補助脳は提案のみ行い、最終判断はユーザーが行う

## Phase 3.2

- 上部に STATUS STRIP を追加
- NEXT ACTION で次に押す操作を1つだけ表示
- SYSTEM / VIDEO / OUTPUT / PROJECT BRIDGE / MEMORY のまとまりへUIを整理
- 補助脳ログの画面表示を最新20件に整理
- 既存の処理ロジック、非自動投稿、安全方針は維持

## Phase 3.3

- SEQUENCE BUILDER を追加
- 複数動画の並びを `selected/sequence.json` に保存
- ffmpeg concatで `selected/horizontal_edit.mp4` を生成
- NVENC優先、失敗時は `libx264` へフォールバック
- Ollama READY時は横編集前に短い構成提案を試行
- タイムライン編集ソフト化はせず、静かな構成に限定

## Phase 3.4

- FOCUS MODE を追加
- Current Step / Next Action / Step Progress で制作導線を整理
- Next Action ボタンから既存処理を直接実行
- STATUS STRIP と詳細機能は残しつつ、通常導線では次の一手を優先
- 完了状態では `整っています。` を表示

## Phase 3.5

- LIVE STATUS ENGINE を追加
- RUNNING / COMPLETED / FAILED などの状態を主画面に表示
- 主要処理のフェーズ、進捗バー、ETA / Expected Finish を表示
- 完了時に `Completed` / `整っています。` と出力ファイル名を表示
- 処理中は対象ボタンとNext Actionをdisabledにして二重実行を避ける

## Phase 3.6

- 単体動画向けの `Generate Horizontal Video` を追加
- `selected/horizontal_video.mp4` と `horizontal_video_log.txt` を生成
- 1920x1080 / H.264 / AAC の通常横動画に整形
- NVENC優先、失敗時は `libx264` へフォールバック
- 縦素材や正方形素材は背景ぼかし + 前景中央配置で横動画化
- FOCUS MODE の Export step を Horizontal Video 優先へ更新

## Phase 3.7

- SMART SHORTS PACK を追加
- INTRO / WORK / AFTERGLOW の3本を `selected/shorts_pack/` に生成
- `shorts_pack.json` と `shorts_pack_log.txt` を保存
- 微細なfade in / fade outで呼吸感を整える
- `selected/bgm/` のBGMを小さめに添える処理に対応
- FOCUS MODE の Export step を Shorts Pack まで案内

## Phase 3.8

- SMART HORIZONTAL EDIT を追加
- `selected_short.json` / Shorts Pack / `shorts_candidates.json` から良い区間を選んで横編集版を生成
- `selected/smart_horizontal_edit.mp4` と `smart_horizontal_sequence.json` を保存
- 微細なfade in / fade outで静かな流れを整える
- `selected/bgm/` のBGMを小さめに添える処理に対応
- FOCUS MODE の Export step を Horizontal Video → Smart Horizontal Edit → Shorts の順に更新

## Phase 3.9

- Smart Horizontal Edit の区間選定を強化
- 3秒fallbackを廃止し、20〜90秒の区間へ拡張
- `shorts_candidates.json` を優先し、短すぎる候補はsource durationから横編集向けに補正
- 候補がない場合はsource durationから3〜5区間を自動生成
- `smart_horizontal_sequence.json` にdurationを保存
- LIVE STATUSに区間数と合計尺を表示
