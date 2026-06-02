# DAKE Alien Road

## アプリ概要

DAKE Alien Road（表示名: エイリアンロード）は、子供のころの1台1ゲーム携帯機風レースゲームをDAKE番外ミニゲームとして復刻した軽量アプリです。

左か右か、それDAKE。扇状に迫る道で障害物とエイリアン機体を避け続ける、反射神経型の単機能ゲームです。

## 操作方法

- `SPACE` / `Enter`: 開始・再開
- `←` / `A`: 左へ移動
- `→` / `D`: 右へ移動
- 画面内の「左」「右」ボタン: 左右移動

操作は左/右だけです。起動後すぐゲーム画面が表示され、余計な設定はありません。

## ゲームルール

- 自機は画面下部に固定されます。
- 道は奥から手前に扇状に広がるパース表示です。
- レーンは5レーンです。
- 障害物は奥から手前へ流れてきます。
- 一部の敵機は横方向にレーン移動します。
- 時間経過で速度が上がります。
- 一定時間ごとに `STAGE` が上がります。
- 衝突したらゲームオーバーです。
- スコア、ステージ、速度、ベストスコアを表示します。
- ベストスコアはローカルconfigに保存されます。

## SNS投稿補助について

ゲームオーバー後に、以下の補助ボタンを表示します。

- 投稿文をコピー
- X投稿画面を開く
- もう一度遊ぶ

投稿文には `DAKE Alien Road`、スコア、`STAGE`、`SPEED`、「左か右か、それDAKE。」、`#DAKEAlienRoad` を含めます。

SNSへの自動投稿はしません。OAuth認証は実装しません。アカウント連携もしません。投稿画面を開く、またはクリップボードへコピーするだけです。

## DAKE_META

```json
{
  "app_key": "game_alien_road",
  "display_name": "エイリアンロード",
  "folder_name": "DAKE_Game_Alien_Road",
  "exe_name": "DakeGame_AlienRoad.exe",
  "status": "internal",
  "show_in_launcher": true,
  "show_on_site": true,
  "release_url": "",
  "screenshot_path": "assets/screenshot.webp",
  "site_title": "エイリアンロード",
  "site_description": "左か右か、それDAKE。扇状に迫る道を避け続ける反射神経ミニゲーム。",
  "update_summary": "左/右だけで遊べるLCD携帯ゲーム風ミニゲームを追加。"
}
```

## RELEASE_BODY

DAKE番外ゲームとして、LCD携帯ゲーム風の反射神経ミニゲーム「DAKE Alien Road」を追加しました。

- 左/右だけで遊べる単機能ミニゲーム
- 扇状に広がる道路と5レーンの回避プレイ
- 固定障害物と、横方向に移動するエイリアン機体
- 時間経過による速度上昇とSTAGE上昇
- スコア、STAGE、SPEED、ベストスコア表示
- ベストスコアをローカルconfigに保存
- ゲームオーバー後の投稿文コピーとX投稿画面オープン
- 自動投稿、OAuth認証、アカウント連携なし
