# Dake BGM Loop Phase 3

のんきなループを、静かに作るための短いBGMループ生成補助アプリです。

## 追加・調整

- moodごとにMock音の傾向を分離
- 同じ mood / duration / seed で同じWAVになるseed再現性を追加
- 先頭と末尾のfade、簡易crossfade、循環echoでループ時のブツつきを軽減
- `outputs/YYYYMMDD/` と `favorites/YYYYMMDD/` に保存
- metadataに `mood_prompt` と `mock_profile` を保存
- `python main.py --deterministic-check` を追加
- ACE-Step実生成は未接続のまま、Mock生成を継続

## 注意

生成音声の商用利用可否は、使用モデル・素材・公開先の規約に依存します。
YouTube等で公開する前に、使用モデルのライセンスを確認してください。
