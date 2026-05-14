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

## Phase 2.0 補助脳レビュー

`Run Assistant Review`, `Open Review File` を追加しました。直近で生成した投稿パッケージ、または `Select Package Folder` で選んだ `data/outputs/packages` 配下のpackageを読み取り、package直下に `assistant_review.md` を作成します。

レビューは以下のファイルから、存在するものだけを読みます。欠けているファイルがあっても処理は止まりません。

- `transcript.txt`
- `transcript_unavailable.txt`
- `shorts_candidates.json`
- `metadata/title_ideas.txt`
- `metadata/description_draft.txt`
- `metadata/tags.txt`
- `metadata/upload_notes.txt`
- `media_info.json`
- `media_info_unavailable.json`

`assistant_review.md` には、Summary、Atmosphere、Recommended Shorts、Title Direction、Description Notes、Before Publish、Assistant Note を出力します。Ollamaが利用できる場合は、長いtranscriptを先頭・中盤・末尾の抜粋に抑えてローカルLLMへ渡し、レビュー生成を試します。Ollamaが利用できない場合や応答に失敗した場合は、既存ファイルの有無、Shorts候補数、タイトル案、media_infoの有無を元に固定テンプレートでレビューを作成します。

補助脳レビューは提案だけを行います。YouTubeへ自動投稿せず、自動公開もせず、最終判断はユーザーが行います。OpenAI APIや外部APIへ投稿パッケージ内容を送りません。

## Phase 2.1 採用候補の整理

`SELECTED OUTPUTS` セクションを追加しました。直近の投稿パッケージ、または `Select Package Folder` で選択したpackageから候補を読み取り、Shorts候補とタイトル候補を選んで `selected/` に選択ドラフトを書き出します。

読み取るファイル:

- `shorts_candidates.json`
- `metadata/title_ideas.txt`
- `metadata/description_draft.txt`
- `metadata/tags.txt`
- `metadata/upload_notes.txt`
- `assistant_review.md`

`Refresh Candidates` で候補を再読み込みします。Shorts候補は最大5件、タイトル候補は最大7件まで選択できます。`Export Selected Draft` を押すと以下を作成します。

```text
selected/
├ selected_short.json
├ selected_title.txt
├ selected_description.txt
├ selected_tags.txt
├ selected_upload_notes.txt
└ selected_summary.md
```

まだ何も選択していない場合は、Shorts #1 と Title #1 を仮採用します。`selected_summary.md` には選択したShorts、タイトル、説明欄、タグ、公開前メモ、人間判断の注意をまとめます。YouTubeへ自動投稿せず、自動公開もしません。元動画や `assistant_review.md` は変更せず、上書き対象はpackage内の `selected/` 配下だけです。

## Phase 2.2 Selected Shorts Preview生成

`Generate Selected Short Preview`, `Open Short Preview` を追加しました。`selected/selected_short.json` から開始・終了・durationを読み取り、元動画から元比率のまま短いプレビュー動画を切り出します。

出力:

```text
selected/
├ short_preview.mp4
└ short_preview_log.txt
```

`selected_short.json` がない場合は、`shorts_candidates.json` の #1 を仮採用して `selected/selected_short.json` を作成してから処理します。今後生成する投稿パッケージには `package_meta.json` と `logs/package_log.txt` に `source_video_path` を保存し、Shortsプレビュー生成で参照します。古いpackageなどで元動画パスが見つからない場合は、動画選択ダイアログから元動画を指定できます。

FFmpegが利用できる場合は `short_preview.mp4` を生成します。NVENCがONLINEなら `h264_nvenc` を優先し、失敗時はCPU `libx264` にフォールバックします。FFmpegが未導入の場合でもアプリは落ちず、`short_preview_log.txt` に理由を残します。Phase 2.2では9:16クロップは行わず、元比率のまま切り出します。9:16自動クロップはPhase 2以降の候補です。

この機能もYouTubeへ自動投稿せず、自動公開もしません。元動画は変更せず、出力と上書きはpackage内の `selected/` 配下だけに限定します。

## Phase 2.3 9:16 Shorts Export

`Generate 9:16 Short`, `Open 9:16 Short` を追加しました。`selected/selected_short.json` から選択区間を読み取り、YouTube Shorts向けの縦動画 `selected/short_vertical_1080x1920.mp4` を作成します。

出力:

```text
selected/
├ short_vertical_1080x1920.mp4
└ short_vertical_log.txt
```

Phase 2.3では高度な人物追跡や自動中心検出は行いません。安定性を優先し、FFmpegの `filter_complex` で「背景ぼかし + 前景中央配置」を行います。

- 背景: 元動画を1080x1920に拡大、中央crop、blur
- 前景: 元動画を1080x1920内に収めて中央配置
- 音声: AAC 192kbpsへ変換

NVENCがONLINEの場合は `h264_nvenc` を優先します。失敗した場合はCPU `libx264` へフォールバックします。9:16出力でもYouTubeへ自動投稿せず、自動公開もしません。元動画は変更せず、出力と上書きはpackage内の `selected/` 配下だけに限定します。

## Phase 2.4 音を置く 連携

DAKE_Music_Otooku（音を置く）で生成した Project Box を読み込み、BGM素材や補助脳提案を、現在の動画制作へ橋渡しできます。これは編集機能ではなく、制作箱を整えるための接続機能です。

読み取り対象:

```text
DAKE_Music_Otooku/data/outputs/projects/
```

`PROJECT BRIDGE` では、Project Box一覧、Selected Project、Preset、Suggested Use、BGM一覧を表示します。`notes/project_notes.txt` がある場合は `Selected Preset` と `Suggested Use` を読み込みます。存在しない場合もアプリは落ちず、`Project notes unavailable.` と表示します。

Previewは必要最小限です。優先は `pygame`、fallbackは `winsound`、最後に `os.startfile` です。複数同時再生はせず、`Preview Start` の前に既存プレビューを停止します。高機能プレイヤー化、波形編集、タイムライン編集は行いません。

`Add to Current Video Box` は、選択BGMを現在の投稿package内へコピーします。

```text
package/
└ selected/
  └ bgm/
    └ selected_bgm.mp3
```

同名ファイルがある場合は上書きせず、`filename_01.mp3` のように連番で回避します。元音源は変更しません。

`Generate Upload Metadata` は以下を作成します。

```text
package/
└ selected/
  └ upload/
    └ metadata_draft.txt
```

OllamaがREADYの場合は、Project BoxのPreset、BGM、Suggested Use、Shorts Directionをローカル補助脳へ渡し、`editing_mood`、`suggested_scene`、`shorts_direction`、`title_direction` の提案を追記します。Ollamaが使えない場合や失敗した場合は固定テンプレートで継続します。YouTube自動投稿、自動公開、元動画や元音源の変更は行いません。

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
