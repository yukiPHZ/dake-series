# Dakeショート切り出し

Dakeショート切り出しは、ローカルMP4からショート動画候補を3本作成し、サムネ画像とタイトル案を同時に出力するDAKEアプリです。

YouTube URL取得や投稿機能は持たせず、MP4を入れてショート候補を出すことだけに絞っています。

## 概要

- MP4を1本選ぶ
- ffprobeで動画情報を確認
- 動画全体を3分割し、それぞれからショート候補を作成
- 9:16 / 1080x1920 のMP4を出力
- サムネ画像とタイトル案テキストを出力
- `ショートごとに文字起こし` をONにすると、生成したショート候補ごとの文字起こしを出力
- QRコードで同じWi-Fi内のスマホから保存できるページを表示

## 使い方

1. `python main.py` で起動
2. MP4をドラッグ＆ドロップ、または `MP4を選ぶ` から選択
3. 必要に応じて `ショートごとに文字起こし` をON/OFF
4. `ショートを作成` を押す
5. 生成後、保存先フォルダとQRコードを確認
6. スマホでQRコードを読み取り、ブラウザから必要なファイルを保存

出力フォルダは入力動画と同じ階層に作成されます。

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
  short_01_transcript.txt
  short_01_segments.json
  short_02_transcript.txt
  short_02_segments.json
  short_03_transcript.txt
  short_03_segments.json
```

## 文字起こし

v0.3では、長尺元動画全体ではなく、生成したショート動画だけを文字起こしします。

- `ショートごとに文字起こし` をONにすると `short_01_transcript.txt` などが生成されます。
- 長尺元動画を丸ごと読まないため、処理の長時間化を避けやすくしています。
- 初回実行時や音声の長いショートでは時間がかかります。
- 文字起こしには `faster-whisper` が必要です。
- 環境により `faster-whisper` の導入やモデル取得ができない場合があります。その場合でもショート動画・サムネ・タイトル案の生成は継続します。
- 候補ごとの文字起こしは、今後Ollama/OpenAIで候補評価するための前段階です。
- v0.3では熱量判定、OpenAI API判定、Ollama判定は未実装です。

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

## DAKE_META

```json
{
  "app_key": "video_shorts_cut",
  "display_name": "ショート切り出し",
  "launcher_title": "ショート切り出し",
  "launcher_description": "MP4からショート候補とサムネを作成します。",
  "site_title": "Dakeショート切り出し",
  "site_description": "MP4を入れるだけで、ショート動画候補・サムネ・タイトル案を作成します。",
  "update_summary": "生成したショート候補ごとに文字起こしを出力できるように更新",
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
