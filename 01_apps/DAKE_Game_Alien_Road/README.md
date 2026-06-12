# DAKE Alien Road

## アプリ概要

DAKE Alien Road（表示名: エイリアンロード）は、子供のころの1台1ゲーム携帯機風レースゲームをDAKE番外ミニゲームとして復刻した軽量アプリです。

左か右か、それDAKE。扇子状に広がって迫る道で障害物とエイリアン機体を避け続ける、反射神経型の単機能ゲームです。

## 操作方法

- `SPACE` / `Enter`: 開始・再開
- `←` / `A`: 左へ移動
- `→` / `D`: 右へ移動
- 画面内の「左」「右」ボタン: 左右移動

操作は左/右だけです。起動後すぐゲーム画面が表示され、余計な設定はありません。

## ゲームルール

- 自機は画面下部に固定されます。
- 道は奥から手前に扇子状に広がるパース表示です。
- レーンは5レーンです。
- 障害物は奥から手前へ流れてきます。
- 敵機は横方向にレーン移動することがあります。
- 当たり判定は見た目より少し甘めです。
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
  "display_name": "DakeAlien Road",
  "folder_name": "DAKE_Game_Alien_Road",
  "exe_name": "DakeGame_AlienRoad.exe",
  "status": "available",
  "show_in_launcher": true,
  "show_on_site": true,
  "release_url": "https://github.com/yukiPHZ/dake-series/releases/tag/DAKE_Game_Alien_Road_v1.0.0",
  "screenshot_path": "assets/screenshot.webp",
  "site_title": "DakeAlien Road",
  "site_description": "左右に動いて障害物とエイリアンを避ける、軽量なレトロ風ミニゲームです。",
  "update_summary": "市場向けミニゲームとして正式出荷準備しました。",
  "launcher_title": "DakeAlien Road",
  "launcher_description": "左右移動だけで遊べる、LCD携帯ゲーム風ミニゲーム。",
  "app_type": "market",
  "completion_goal": "formal_release",
  "demo_video_path": "release_artifacts/demo.mp4",
  "demo_video_url": "",
  "social_release_path": "release_artifacts/social_release.json"
}
```

## RELEASE_BODY

- 左右移動だけで遊べるレトロ風ミニゲーム
- 5レーンの回避アクション
- ベストスコア保存
- Windows向けexe
