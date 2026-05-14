# 音を置く

言葉から、動画や配信用の小さな音素材を作る補助アプリです。

「空気を、音として置く」ためのDAKEアプリであり、作曲ソフトやDAWではありません。入力した言葉をもとに、補助脳がBPM・雰囲気・構成案を整え、環境があればローカルのAudioCraft / MusicGenで短い音素材を生成します。生成できない環境でも、プロンプト・設計メモ・出力フォルダを作ります。

## できること

- CLI / ローカル環境チェック
- 言葉から音の方向性を整理
- Ollama利用時の補助脳プロンプト生成
- AudioCraft / MusicGen 利用可能時の短いwav生成
- FFmpegによるwav/mp3変換
- ループプレビュー作成
- 出力フォルダ作成
- 処理ログ保存

## Phase 1でまだやらないこと

- 完成曲制作
- ボーカル曲制作
- DAW機能
- MIDI編集
- ピアノロール
- 自動公開
- 有料API連携
- 特定アーティストや既存曲の再現

## 必要な外部ツール

必須ではありません。未導入でもアプリは落ちず、設計メモとプロンプトを保存します。

- ffmpeg: 音量調整、wav/mp3変換、ループプレビュー作成
- ffprobe: 音源の尺確認
- ollama: ローカル補助脳
- audiocraft / MusicGen: ローカル音生成
- torch / CUDA: MusicGen利用時のGPU確認

## 次に入れるもの

1. requests
   - Ollama補助脳接続に使用

2. FFmpeg
   - wav / mp3変換、ループ整形に使用

3. Ollama
   - ローカル補助脳として、音の方向性を生成

## 使い方

1. `python main.py` で起動します。
2. PROMPTに言葉や空気感を入力します。
3. 必要なら「参照音源を選ぶ」で既存音源を指定します。
4. 「音を置く」を押します。
5. 完了後、「出力フォルダを開く」で素材を確認します。

## 出力フォルダ構成

`~/Music/Otooku/{project_name}/` に以下を作成します。

```text
audio/
  generated_preview.wav
  generated_preview.mp3
  generated.wav
  generated.mp3
  source_converted.wav
  source_converted.mp3
  loop_preview.mp3
  loop_pack/
    quiet_loop_30s.mp3
    quiet_loop_30s.wav
    quiet_loop_60s.mp3
    quiet_loop_60s.wav
    quiet_loop_180s.mp3
    quiet_loop_180s.wav
video_bgm_pack/
  bgm/
    shorts/
    long/
    ambient/
    work/
  notes/
    usage_note.txt
    shorts_ideas.txt
    long_video_ideas.txt
  export_log.txt
prompts/
  music_direction.txt
  musicgen_prompt.txt
notes/
  usage_note.txt
  loop_notes.txt
logs/
  process_log.txt
setup_needed.txt
```

`generated_preview.wav` は、MusicGen未導入でもFFmpegの簡易ambient生成で作成します。`generated.wav`、`generated.mp3`、`loop_preview.mp3` は、MusicGen生成・参照音源・Tiny Ambient生成をFFmpeg整形まで進めた場合に作成されます。

プリセットを選択した場合は、`music_direction.txt`、`musicgen_prompt.txt`、`usage_note.txt`、`loop_notes.txt`、`video_bgm_pack/notes/usage_note.txt` にプリセット名とタグを追記します。

## 参照音源の整形

FFmpeg が導入されている場合、既存の音源ファイルを選択して、wav / mp3 / loop_preview.mp3 を出力できます。

MusicGen未導入でも、手持ちのBGMや効果音を素材として整えることができます。

## Loop Pack

音を置く は、既存音源から複数長さのループ素材を生成できます。

動画編集前に、“置ける音素材棚”を作る用途を想定しています。

## Video BGM Pack

Loop Pack を、動画制作向けの用途別素材棚として整理できます。

Shorts / Long / Ambient / Work 用に分類し、補助脳が使用イメージを生成します。

## Audio Preview

出力された mp3 / wav を、アプリ内または既定プレイヤーで簡易確認できます。

これは編集機能ではなく、生成された音素材をその場で確認するためのPreviewです。

生成後は `generated_preview.wav` を優先して再生対象にします。

`pygame` が利用できる場合は mp3 / wav をアプリ内で再生できます。`pygame` がない場合は wav をWindows標準の `winsound` で優先再生し、mp3は同名wavがある場合はそちらを再生します。既定プレイヤーで開くのは最後のfallbackです。

FFmpeg / FFprobe の実行時は、Windowsで黒いコンソールウィンドウが出にくいよう `CREATE_NO_WINDOW` を指定しています。

## 補助脳シリーズUI

音を置く は、BRAINZ / Dakeユキズ稼働中 と同じ補助脳シリーズとして、ブラックトーンのUIへ寄せています。

dark background / thin border / blue-purple accent を基調にし、ログは小さく、音素材生成の状態だけを確認できるようにしています。

## Tiny Ambient Generator

音を置く は、文章から小さなambient preview音を生成できます。

これは完成曲制作ではなく、「空気を置く」ための最初の音生成です。

MusicGen未導入でも、FFmpegだけでPreview音を生成します。生成後は `generated_preview.wav` と、可能なら `generated_preview.mp3` を作成し、Preview一覧へ自動反映します。

## Favorite

Previewで確認した音源を、あとで使うための Favorite 棚へ保存できます。

Favorite は `data/favorites` に保存されます。これは音楽管理ソフトではなく、動画やサイト制作で使えそうな音を一時的に置くための棚です。

保存時は音源を `data/favorites/audio/` にコピーし、`favorite_index.json` と `notes/favorite_note.txt` に出どころ、日時、プリセット、タグ、尺を残します。同名ファイルがある場合は上書きせず、日時付きの名前で保存します。

## Project Bridge

Favorite音源を、動画制作向けフォルダへ橋渡しできます。

これは動画編集ではなく、「制作箱」を準備するための機能です。

制作箱は `~/Music/Otooku/projects/{project_name}/` に作成し、`raw` / `bgm` / `shorts` / `notes` / `thumbnails` / `export` / `upload` を用意します。選択したFavorite音源は `bgm/` へコピーし、`notes/project_notes.txt` と `export_log.txt` を保存します。

## Preset System

音を置く は、入力文だけでなく、BORINEF / holiday-jinja / YUKIZ稼働中 などのプリセットを使って、音の方向性を揃えることができます。

プリセットは固定命令ではなく、補助脳に渡す空気のメモとして扱います。

プリセット定義は `data/presets/music_presets.json` に置き、Phase 7では以下の5つのみを初期登録しています。

- BORINEF
- holiday-jinja
- YUKIZ稼働中
- quiet work
- blue memory

## AudioCraft / MusicGen について

MetaのAudioCraftは、AudioGen / MusicGenなどの音声生成モデルの推論・学習コードを含むPyTorchライブラリです。環境構築が重いため、Phase 1では任意です。

AudioCraft / MusicGen がimportできる場合のみ、10〜20秒程度の短い素材生成を試します。未導入、モデル未取得、GPU不足、推論エラーなどの場合は、生成を止めて `setup_needed.txt` を保存します。

## Ollama について

Ollamaは `http://localhost:11434` のローカルAPIを確認します。利用可能な場合、入力文から以下を生成します。

- mood
- BPM
- key候補
- instrumentation
- loop length
- musicgen prompt
- negative notes
- usage idea

Ollamaが未起動または未導入の場合は、固定テンプレートで代替します。

## ローカル補助脳

Ollama + qwen2.5:7b を使うことで、ローカルGPU上で音の方向性を生成できます。

外部APIではなく、自分のPC内で補助脳が動作します。

## FFmpegについて

FFmpegが利用可能な場合、生成音源または参照音源に対して以下を行います。

- wav変換
- mp3変換
- 簡易音量ノーマライズ
- fade in / fade out
- loop_preview.mp3 作成

FFmpegがない場合は、`FFmpeg is required for audio export` を表示し、アプリは落としません。

FFmpegを入れると wav/mp3変換やループ整形が使えます。未導入でもプロンプト生成は使えます。Phase 2ではFFmpeg / FFprobe未導入でも落ちず、画面上に `FFmpeg is offline. Prompt output is still available.` と表示します。

## 著作権・利用上の注意

このアプリは、動画やサイト用のオリジナル音素材を作る補助ツールです。既存楽曲や特定アーティストの権利を侵害する目的で使用しないでください。

生成物の利用可否は、使用したモデル・素材・配布条件を確認してください。外部サービスへ勝手にアップロードせず、生成音源を自動公開する機能も持ちません。

## DAKE_META

```json
{
  "app_key": "DAKE_Music_Otooku",
  "display_name": "音を置く",
  "launcher_title": "音を置く",
  "launcher_description": "言葉から、動画や配信用の小さな音素材を作る補助アプリです。",
  "site_title": "音を置く",
  "site_description": "言葉を入力して、BGMやループ素材の方向性を整えるDAKEアプリです。",
  "update_summary": "音を置く Phase 1 を追加しました。",
  "folder_name": "DAKE_Music_Otooku",
  "exe_name": "DakeMusic_Otooku.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "draft",
  "show_in_launcher": false,
  "show_on_site": false
}
```

## RELEASE_BODY

# 音を置く v0.1.0

言葉から、動画や配信用の小さな音素材を作る補助アプリです。

## できること

- 言葉から音の方向性を整理
- Ollama利用時の補助脳プロンプト生成
- AudioCraft / MusicGen 利用可能時の短い音源生成
- FFmpegによるwav/mp3変換
- ループプレビュー作成
- 出力フォルダ作成
- 処理ログ保存

## まだやらないこと

- 完成曲制作
- ボーカル曲制作
- DAW機能
- MIDI編集
- 自動公開
- 有料API連携
