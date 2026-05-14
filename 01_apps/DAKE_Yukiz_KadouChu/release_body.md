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
