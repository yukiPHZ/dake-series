# Dake潜って獲る

## アプリ概要

`Dake潜って獲る` は、80年代〜初期90年代のLCD携帯ゲーム風に遊べるDAKE番外ミニゲームです。船から海へ潜り、銛で小魚・大魚・宝箱を獲って、サメを避けながら船へ戻ると得点になります。

単機能・軽量・即起動を優先し、外部画像素材や追加依存を使わない Python / Tkinter 製のCanvasゲームです。

## 操作方法

- `Enter`: 開始 / 再開
- `R`: リスタート
- `↑` / `W`: 上へ移動
- `↓` / `S`: 下へ移動
- `←` / `A`: 左へ移動
- `→` / `D`: 右へ移動
- `Space`: 銛を撃つ
- `P`: 一時停止 / 再開

## ゲームルール

- ダイバーは船の位置から開始します。
- 海中は横5列 × 縦6段のマス移動です。
- 海中で `Space` を押すと、向いている方向へ銛を撃ちます。
- 銛が魚または宝箱に当たると、獲物保持状態になります。
- 獲物を取った時点では得点せず、船へ戻ると得点になります。
- 得点後、`HOLD` は空に戻ります。
- サメに触れるとライフが1減ります。
- 獲物保持中にサメに触れると、保持中の獲物も失います。
- ライフ0でゲームオーバーです。
- ゲームオーバー時は `Enter` または `R` で再開できます。
- BESTスコアは `diver_catch_config.json` に保存されます。

## 得点ルール

- 小魚: 10点
- 大魚: 30点
- 宝箱: 100点

## DAKE_META

```json
{
  "app_key": "game_diver_catch",
  "display_name": "Dake潜って獲る",
  "launcher_title": "Dake潜って獲る",
  "launcher_description": "船から潜って、銛で獲って、戻って得点。LCD携帯ゲーム風ミニゲーム。",
  "site_title": "Dake潜って獲る",
  "site_description": "80年代〜初期90年代のLCD携帯ゲーム風に、ダイバーが海へ潜って獲物を持ち帰るミニゲームです。",
  "update_summary": "LCD携帯ゲーム風のダイバーゲームを追加。",
  "folder_name": "DAKE_Game_Diver_Catch",
  "exe_name": "DakeGame_Diver_Catch.exe",
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true
}
```

## RELEASE_BODY

Dake潜って獲る
船から潜って、銛で獲って、戻って得点
LCD携帯ゲーム風ミニゲーム
Windows向けexe
