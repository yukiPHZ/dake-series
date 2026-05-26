# DAKE_SCREENSHOT_RULE

## 正式出荷画像の役割

### assets/screenshot.webp

役割:

```text
実際に動く感
```

用途:

- dakeapp.com
- README
- GitHub Release
- Launcher

ルール:

- 実際の起動画面を撮る。
- アプリウインドウのみを撮る。
- 実務中の自然状態を優先する。
- 横幅1200px超のみ縮小する。
- 引き延ばしは禁止する。
- DAKE正式出荷物として必ず保持する。

### assets/booth_thumbnail.jpg

役割:

```text
一覧で止まる感
```

用途:

- BOOTH一覧
- 商品一覧
- 視認性
- 商品理解

ルール:

- 1200x1200で生成する。
- 正方形にする。
- BOOTH一覧で読めることを優先する。
- 小サイズでも文字やUIが潰れないようにする。
- BOOTH向けに最適化する。
- `assets/screenshot.webp` と `assets/screenshot.jpg` は削除しない。

アプリ画面スクリーンショットの作成ルールです。

## 正本

- 保存先: `01_apps/<app>/assets/screenshot.webp`
- READMEの `DAKE_META.screenshot_path` は `assets/screenshot.webp` にする。
- dakeapp.com 掲載用画像もこのファイルを使う。

## 撮影手順

1. 各アプリの `dist` 内exeを起動する。
2. 起動後、画面が表示されるまで少し待つ。
3. アプリウインドウだけを撮影する。
4. WebP形式で保存する。
5. exeを終了する。

## 品質ルール

- 起動直後の画面を撮る。
- 最低限の初期表示でよい。
- 余計なデスクトップ背景を入れない。
- アプリウインドウだけを撮る。
- WebP形式にする。
- 縦横比は必ず維持する。
- 元画像が1200px以下なら原寸維持。
- 元画像が1200pxより大きい場合のみ、横幅1200px以内に縮小する。
- 横幅1200pxへの引き延ばしは禁止。
- トリミングでウインドウ枠や重要なUIを欠けさせない。

## 失敗時

exeがない場合:

```json
"status": "missing_exe"
```

起動または撮影に失敗した場合:

```json
"status": "build_failed"
```

READMEに短く原因を残す。

## 撮影前のUI確認

- 空状態が不自然でない。
- ボタンや入力欄が重なっていない。
- フッターが見切れていない。
- 文字化けがない。
- アプリ名や用途が分かる。

## 画像確認

- `assets/screenshot.webp` が存在する。
- WebPとして開ける。
- 横幅が1200px以内。
- 明らかなデスクトップ背景が写っていない。
- 余白がありすぎず、アプリ画面としてそのまま使える。
