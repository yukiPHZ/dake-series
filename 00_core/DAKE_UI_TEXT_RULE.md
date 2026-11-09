# DAKE_UI_TEXT_RULE

UI文言または共通フッターを変更する時だけ参照します。日本語・英語を含む画面文言は `UI_TEXT` へ集約し、UTF-8で管理します。

## 基本定義

```python
APP_NAME = "Dakeアプリ名"
WINDOW_TITLE = "短いタイトル"
KAITORI_APP_ID = "pdf_merge"
KAITORI_URL = f"https://dakeapp.com/go/kaitori?app={KAITORI_APP_ID}"

UI_TEXT = {
    "main_title": "主見出し",
    "main_description": "短い説明",
    "footer_kaitori": "戸建買取査定",
    "footer_author": "by Yukihiko Kikuta",
}
```

`KAITORI_APP_ID` はアプリごとの安定した英小文字・数字・アンダースコア識別子にします。実際の査定ページURLをexeへ固定せず、DAKE側の恒久入口を使います。

## 共通フッター

表示は次を基本とします。

```text
戸建買取査定 ｜ by Yukihiko Kikuta
```

- `戸建買取査定` だけを `KAITORI_URL` へのリンクにする。
- 区切りの `｜` と `by Yukihiko Kikuta` はリンクにしない。
- Instagram、長いコピーライト、Vibe-Coded、しまりす不動産の長いブランド表記、不要な外部リンクを新仕様へ追加しない。
- フッターを主操作より目立たせない。
- 既存全アプリを一括変更せず、次回更新対象から移行する。

戸建買取査定は株式会社さくら都市の業務であり、同社の基準・手続きに基づいて対応します。「しまりす不動産が戸建を買い取る」と読める表現を使いません。

## UI_TEXTに入れるもの

- 見出し、説明、ボタン、ラベル
- プレースホルダー、状態、エラー
- ダイアログ、フッター
- 出力名等の固定文言

画面部品へ同じ文字列を複数回直書きしません。

## レビュー

- 日本語が短く自然で、文字化けがない。
- ボタンは動詞中心。
- エラーは原因と次の行動が短く分かる。
- `by Yukihiko Kikuta` が非リンク。
- アプリ本体から利用状況やファイル情報を送信していない。

WindowsファイルプロパティはUI文言と別管理です。定義は `DAKE_COMMON_SPEC.md`、生成方法は必要時に `DAKE_BUILD_RULE.md` を参照します。
