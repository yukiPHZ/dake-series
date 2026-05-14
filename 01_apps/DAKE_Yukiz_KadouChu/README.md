# Dakeユキズ稼働中

Dakeユキズ稼働中は、LIVEアーカイブ動画・撮影動画・画面録画を取り込み、YouTube投稿前の素材整理をローカルPCで進める制作運用コンソールです。動画編集ソフトではなく、制作が止まらないためのGPU制作補助脳です。

基本UIは英語、補助脳ログと完了文言だけ日本語に寄せています。

## Phase 1でできること

- 動画ファイルの取り込み（mp4 / mov / mkv / webm）
- FFmpeg / FFprobe / yt-dlp / Ollama / GitHub CLI / Wrangler の環境チェック
- FFmpegの `h264_nvenc` 検出
- FFprobeによる動画情報取得
- ETA / Expected Finish の簡易表示
- faster-whisper が使える環境での文字起こし
- 文字起こし不可時の `transcript_unavailable.txt` 作成
- Shorts候補JSONの作成
- 候補1本目のプレビュークリップ作成
- タイトル案・説明欄・タグ・アップロードメモの雛形出力
- 出力フォルダを開く
- 処理ログ保存

## Phase 1でまだやらないこと

- YouTubeへの自動投稿
- 自動公開
- 完全自動編集
- PSD編集
- BGM自動ミックス
- Shorts 9:16の高度な自動クロップ
- YouTube LIVE URLからの本ダウンロード

## 必要な外部ツール

- Python 3.10+
- FFmpeg / FFprobe
- yt-dlp（Phase 1ではメタデータ取得のみ）
- faster-whisper（任意。未導入でもアプリは継続）
- Ollama（任意。localhost APIが起動していればメタデータ案に使用）
- GitHub CLI `gh`（確認のみ）
- Wrangler（確認のみ）

`requirements.txt` には `customtkinter`, `pillow`, `opencv-python`, `faster-whisper`, `requests`, `psutil`, `pyinstaller` を記載しています。`faster-whisper` は環境構築が重い場合があります。Importできない場合でも、アプリは `Transcription unavailable` として処理を続けます。

## 使い方

1. `python main.py` で起動します。
2. `Select Video File` から動画素材を選択します。
3. CLI / GPU / メディア情報の状態を確認します。
4. `Start Production Run` を押します。
5. 完了後、`Open Output Folder` で生成物を確認します。

YouTube LIVE URL欄はPhase 1ではメタデータ確認用です。本ダウンロードは行いません。

## フォルダ構成

```text
DAKE_Yukiz_KadouChu/
  main.py
  build.bat
  README.md
  release_body.md
  requirements.txt
  assets/
    peakheadz_logo.png  # 任意。置けば起動時に小さく表示
  data/
    inbox/
    bgm/
    templates/
      thumbnails/
      logos/
    outputs/
    logs/
  core/
  ui/
```

BGMファイルは `data/bgm` に置きます。過去サムネPNGやPSDは `data/templates/thumbnails` に置きます。PEAKHEADZロゴは `assets/peakheadz_logo.png` に置きます。Phase 2以降でBGM選定・サムネ生成に使います。PSD自体の編集はPhase 1では行いません。

## 出力されるファイル

出力は `data/outputs/{project_name}/` に作成されます。

- `source_manifest.json`
- `media_info.json`
- `transcript.txt`
- `transcript.srt`
- `transcript_unavailable.txt`
- `shorts_candidates.json`
- `shorts/short_01_preview.mp4`
- `shorts/preview_unavailable.txt`
- `metadata/title_ideas.txt`
- `metadata/description_draft.txt`
- `metadata/tags.txt`
- `metadata/upload_notes.txt`
- `logs/process.log`

## 補助脳の位置づけ

補助脳は、素材を見て、止まらず次の作業へ進むためのローカル支援です。Ollamaが起動している場合は localhost API でタイトル案などを試作します。起動していない場合は固定テンプレートで代替します。

キーフレーズは次の通りです。

- 稼働中。
- 整っています。
- 側に。
- 止まりません。

## 自動公開しない方針

Phase 1ではYouTube投稿、GitHub Release、Cloudflare deployを自動実行しません。APIキーやトークンも保存しません。CLIは存在確認とローカル処理補助に限定します。

## GPU / NVENCについて

起動時に `ffmpeg -encoders` を確認し、`h264_nvenc` が見つかれば `NVENC ONLINE` と表示します。NVIDIA GPU環境ではFFmpegのGPUエンコードが使える可能性がありますが、ドライバ・FFmpegビルド・GPU世代に依存します。使えない場合はCPUエンコードへフォールバックします。

Phase 1.5では `nvidia-smi` が使える場合のみGPU名とVRAMを表示します。`nvidia-smi` がない場合は `GPU SKIPPED` として扱い、必須にはしません。

## build

`build.bat` はPyInstallerで `dist/DakeYukiz_KadouChu.exe` を作成します。共通アイコン `..\..\02_assets\dake_icon.ico` を使用します。`build/`, `dist/`, `*.spec`, `*.exe` はGit管理しません。

## Phase 1.5 System Check

制作基地の起動感を強めるため、`Run System Check` で以下を確認します。

- FFmpeg / FFprobe / yt-dlp の導入状態
- GitHub CLI `gh auth status` による認証確認
- Wrangler `wrangler whoami` による認証確認
- Ollama localhost API 接続確認とモデル名の簡易表示
- FFmpeg encoders から `h264_nvenc` / `hevc_nvenc` を確認
- `nvidia-smi` が使える場合のみGPU名とVRAMを表示
- 不足がある場合の `data/outputs/system_check/install_guide.txt` 生成

表示は `ONLINE`, `MISSING`, `UNAUTHORIZED`, `READY`, `UNAVAILABLE` を中心にしています。GitHub CLI / Wrangler は認証済みの場合 `AUTHORIZED`、未認証の場合 `UNAUTHORIZED` と表示します。未導入の道具があってもアプリは落ちません。

## Phase 1.6 CLI導入補助モード

System欄に `Open Install Guide`, `Copy Install Commands`, `Recheck System` を追加しました。

- `Open Install Guide`: `data/outputs/system_check/install_guide.txt` を開きます。
- `Copy Install Commands`: 不足しているCLIだけの導入候補コマンドをクリップボードへコピーします。
- `Recheck System`: CLI / GPU / Ollama / NVENC / npm を再確認します。

このアプリは自動インストールしません。`winget`, `npm`, `pip`, exe download などを勝手に実行しません。PATH変更、管理者権限、利用規約確認が絡むため、コマンドは必ず内容を確認してからユーザーが実行してください。

WranglerにはNode.js / npm が必要です。npmがない場合は先にNode.jsを入れてから `npm install -g wrangler` を実行します。OllamaとGPUは、FFmpegなどのCLIより先に `READY` として認識される場合があります。

## Phase 1.7 初回動画テスト

`Select Test Video`, `Run First Video Test`, `Open Test Output` を追加しました。実動画を1本選び、`ffprobe` で動画情報を取得し、`ffmpeg` で先頭10秒の `test_clip.mp4` を作成します。

- `data/outputs/first_video_test/media_info.json` に動画情報を保存
- `data/outputs/first_video_test/test_clip.mp4` に10秒テストクリップを保存
- `h264_nvenc` が使える場合はNVENCを優先
- NVENC失敗時はCPU `libx264` へフォールバック
- `data/outputs/first_video_test/logs/test_log.txt` に結果ログを保存

元動画は絶対に変更しません。元動画と同じフォルダにも出力しません。すべて `data/outputs/first_video_test` 配下に保存します。FFmpeg / FFprobe が未導入でもアプリは落ちず、CLI導入補助へ進めます。

## Phase 1.8 投稿パッケージ生成

`Generate Posting Package`, `Open Package Folder` を追加しました。動画を1本選び、YouTube投稿前に人間が確認する素材一式を `data/outputs/packages/{YYYYMMDD_HHMM}_{safe_video_name}/` に作成します。

出力フォルダ構成:

```text
data/outputs/packages/{YYYYMMDD_HHMM}_{safe_video_name}/
├ media_info.json または media_info_unavailable.json
├ transcript.txt / transcript.srt または transcript_unavailable.txt
├ shorts_candidates.json
├ metadata/
│  ├ title_ideas.txt
│  ├ description_draft.txt
│  ├ tags.txt
│  └ upload_notes.txt
└ logs/
   └ package_log.txt
```

`ffprobe` が使える場合は `media_info.json` を保存します。使えない場合は `media_info_unavailable.json` を作成し、処理は続行します。`faster-whisper` が使える場合は `transcript.txt` と `transcript.srt` を作成し、使えない場合は `transcript_unavailable.txt` を作成します。

Ollamaが `READY` の場合は localhost のローカルLLMで、短い日本語タイトル案の生成を試します。Ollamaが応答しない場合や生成に失敗した場合は、固定テンプレートで `title_ideas.txt`, `description_draft.txt`, `tags.txt`, `upload_notes.txt` を整えます。

この機能はYouTubeへ自動投稿しません。自動公開もしません。OpenAI APIや外部APIへ動画・文字起こしを送りません。元動画は変更せず、出力は `data/outputs/packages` 配下だけに作成します。

## Phase 2候補

- YouTube LIVE URLからyt-dlpで本取得
- YouTube APIで下書きアップロード
- BGM自動選定
- サムネ自動生成
- 過去サムネ/PSD/PNGテンプレ参照
- Shorts 9:16自動クロップ
- 顔/手元/画面中心の自動追跡
- Ollamaによるより高度な補助脳コメント
- GitHub CLI / Wrangler の操作補助
- DAKE Launcher連携
- 処理完了通知
- ETA精度向上
- GPU使用率表示
- NVENC/CPU切替設定

## DAKE_META

```json
{
  "app_key": "DAKE_Yukiz_KadouChu",
  "display_name": "Dakeユキズ稼働中",
  "launcher_title": "ユキズ稼働中",
  "launcher_description": "動画素材を取り込み、補助脳で投稿準備を整える制作コンソールです。",
  "site_title": "Dakeユキズ稼働中",
  "site_description": "LIVEや撮影動画を、YouTubeに出せる形へ静かに整えるGPU制作補助脳です。",
  "update_summary": "Dakeユキズ稼働中 Phase 1 を追加しました。",
  "folder_name": "DAKE_Yukiz_KadouChu",
  "exe_name": "DakeYukiz_KadouChu.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "draft",
  "show_in_launcher": false,
  "show_on_site": false
}
```

## RELEASE_BODY

Dakeユキズ稼働中 v0.1.0。動画素材を取り込み、補助脳で投稿準備を整える制作コンソールです。
