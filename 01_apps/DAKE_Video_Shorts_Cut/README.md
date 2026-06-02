# Dakeショート切り出し

Dakeショート切り出しは、ローカルMP4からショート動画候補を3本作成し、サムネ画像とタイトル案を同時に出力するDAKEアプリです。

YouTube URL取得や投稿機能は持たせず、MP4を入れてショート候補を出すことだけに絞っています。

## 概要

- MP4を1本選ぶ
- ffprobeで動画情報を確認
- 動画全体を3分割し、それぞれからショート候補を作成
- 9:16 / 1080x1920 のMP4を出力
- 標準では横動画全体を見せ、上下にぼかし背景を入れて出力
- サムネ画像とタイトル案テキストを出力
- `ショートごとに文字起こし` をONにすると、生成したショート候補ごとの文字起こしを出力
- ショート候補ごとに評価を行い、使えそう度・タイトル案・サムネ文言案を出力
- `Ollamaで候補評価` をONにすると、Ollama起動時だけローカルAI評価を試行
- おすすめ候補を上部に表示し、タイトル案・サムネ文言・理由を確認
- QRコードで同じWi-Fi内のスマホから保存できるページを表示

## 使い方

1. `python main.py` で起動
2. MP4をドラッグ＆ドロップ、または `MP4を選ぶ` から選択
3. 必要に応じて `全体表示（おすすめ）` をON/OFF
4. 必要に応じて `ショートごとに文字起こし` をON/OFF
5. 必要に応じて `Ollamaで候補評価` をON
6. `ショートを作成` を押す
7. 生成後、保存先フォルダとQRコードを確認
8. スマホでQRコードを読み取り、ブラウザから必要なファイルを保存

出力フォルダは入力動画と同じ階層に作成されます。入力動画が `dake_shorts_output_...` フォルダ内にある場合は、出力フォルダの入れ子を避けるため一段上に作成します。

```text
dake_shorts_output_YYYYMMDD_HHMMSS/
  short_01.mp4
  short_02.mp4
  short_03.mp4
  thumb_01.jpg
  thumb_02.jpg
  thumb_03.jpg
  title_01.txt
  title_02.txt
  title_03.txt
  thumb_text_01.txt
  thumb_text_02.txt
  thumb_text_03.txt
  short_01_transcript.txt
  short_01_segments.json
  short_02_transcript.txt
  short_02_segments.json
  short_03_transcript.txt
  short_03_segments.json
  shorts_review.txt
  shorts_review.json
```

## 文字起こし

v0.3では、長尺元動画全体ではなく、生成したショート動画だけを文字起こしします。

- `ショートごとに文字起こし` をONにすると `short_01_transcript.txt` などが生成されます。
- 長尺元動画を丸ごと読まないため、処理の長時間化を避けやすくしています。
- 初回実行時や音声の長いショートでは時間がかかります。
- 文字起こしには `faster-whisper` が必要です。
- 環境により `faster-whisper` の導入やモデル取得ができない場合があります。その場合でもショート動画・サムネ・タイトル案の生成は継続します。
- 候補ごとの文字起こしは、簡易評価と今後のOllama/OpenAI評価に使います。

## 候補評価

v0.6では、通常のルールベース評価に加えて、Ollamaが起動している場合だけローカルAI評価を使えます。

- `short_01_transcript.txt` などを読み、0〜100点の使えそう度を付けます。
- `shorts_review.txt` と `shorts_review.json` を出力します。
- `title_01.txt` 〜 `title_03.txt` は評価結果のタイトル案で更新されます。
- `thumb_text_01.txt` 〜 `thumb_text_03.txt` は評価結果のサムネ文言で作成されます。
- QR転送ページにも候補ごとの score / title / thumbnail_text / reason を表示します。
- `Ollamaで候補評価` は初期OFFです。
- Ollamaの既定モデルは `qwen2.5:7b` です。
- Ollama未導入・未起動・モデル未取得の場合でも、通常評価へ戻って動作します。
- OpenAI APIはまだ未使用です。
- 熱量判定、OpenAI API判定は未実装です。

設定は `DakeVideo_Shorts_Cut_config.json` に保存されます。このファイルはGit管理しません。

```json
{
  "display_full_mode": true,
  "use_ollama_review": false,
  "ollama_model": "qwen2.5:7b",
  "ollama_url": "http://localhost:11434/api/generate"
}
```

## v0.6 実用表示

- 生成結果欄の上部におすすめ候補カードを表示します。
- おすすめカードには score / title / thumbnail_text / reason をまとめて表示します。
- QR転送ページの上部におすすめ候補を表示します。
- QR転送ページでは、スマホ投稿素材として `short_01.mp4` / `thumb_01.jpg` / `title_01.txt` / `thumb_text_01.txt` / `short_01_transcript.txt` にすぐ移動できます。
- 候補一覧にも `thumb_text_01.txt` などを表示します。
- JSON系ファイルは下部の開発用ファイルにまとめています。
- OpenAI APIは未使用です。
- Ollamaは任意機能です。

## v0.7 縦動画見栄え改善

- ショート動画の標準出力を、横動画全体表示＋上下ぼかし背景に変更しました。
- 横動画全体が見えるため、制作配信や作業配信の画面情報を残しやすくなります。
- 背景には拡大コピーした動画を強めにぼかして配置します。
- 中央には元動画全体を1080px幅に収めて表示します。
- `全体表示（おすすめ）` をOFFにすると、従来の中央クロップで出力できます。

## 必要なもの

- ffmpeg
- ffprobe
- Pythonライブラリ
  - qrcode
  - Pillow
  - tkinterdnd2
  - faster-whisper

Pythonライブラリは以下でインストールできます。

```powershell
pip install -r requirements.txt
```

ffmpeg / ffprobe はPATHから実行できる状態にしてください。

## QR転送の注意

- PCとスマホが同じWi-Fi内にあるときに使います。
- iPhoneの写真アプリへ完全自動保存はしません。
- iPhoneではQRで転送ページを開きます。
- 動画を開き、共有メニューから「ビデオを保存」を選んでください。
- 保存先や表示挙動はiOS/Safari側の仕様に依存します。
- 画像もブラウザ表示後、共有や保存操作で手動保存してください。
- PC側アプリを閉じると転送ページも停止します。

## v1でやらないこと

- YouTube URL取得
- 自動投稿
- BGM追加
- テロップ編集
- AI高精度判定
- 熱量判定
- SNS投稿予約

## ビルド方法

`build.bat` を実行すると、`dist/DakeVideo_Shorts_Cut.exe` を生成します。

- 出力exe名: `DakeVideo_Shorts_Cut.exe`

DAKE共通アイコン `..\..\02_assets\dake_icon.ico` を使用します。

## 検証

```powershell
python -m py_compile main.py
python main.py --launch-check
```

実動画で確認する場合:

```powershell
python main.py --process-check "C:\path\to\sample.mp4"
```

Ollama評価をCLIで確認する場合:

```powershell
python main.py --process-check "C:\path\to\sample.mp4" --ollama-review
```

従来の中央クロップで確認する場合:

```powershell
python main.py --process-check "C:\path\to\sample.mp4" --center-crop
```

## DAKE_META

```json
{
  "app_key": "video_shorts_cut",
  "display_name": "ショート切り出し",
  "launcher_title": "ショート切り出し",
  "launcher_description": "MP4からショート候補とサムネを作成します。",
  "site_title": "Dakeショート切り出し",
  "site_description": "MP4を入れるだけで、ショート動画候補・サムネ・タイトル案を作成します。",
  "update_summary": "横動画全体表示＋上下ぼかし背景によるショート動画出力に対応",
  "folder_name": "DAKE_Video_Shorts_Cut",
  "exe_name": "DakeVideo_Shorts_Cut.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

```text
MP4からショート動画候補を作成
サムネ画像とタイトル案も同時出力
QRコードでスマホへ転送
Windows向けexe
```
