# フォント監査と再現記録

> この監査は配布ZIPにも同梱します。本文の scripts/・evidence/ への相対リンクは開発フォルダで参照する検証根拠です。ZIPにはユーザー設定・試験作品・検証スクリプトを同梱しません。配布版の結論と識別は ../TEST_REPORT.md を参照してください。

2026-09-30。0.3.0 の合格報告を根拠にせず、現在の実体から再検証した。

## 修正前の実体

- 配布 EXE SHA256: ebff99d49ecdc2b2f4c53d224ec2f89c39342a717286410ef64202ea8d530014
- app.asar SHA256: 990fb93832a58899fd8e6163367dcdc8d1afd3d276de620064ce25860a6d847d
- app.asar 内 build/app.js と desktop/preload.cjs、fonts.cjs、main.cjs は修正前ソース/ビルドと SHA256 一致。差替え漏れを原因とはしない。
- 保存先の解決順は STADIO_USER_DATA → EXE横の .stadio-portable があれば user-data → %APPDATA%/DAKE STADIO。既存配布直下に portable マーカーはなかった。利用設定には Google フォント9種と復旧データが存在した。元設定は消さず、専用 profile に複製した。

## 同一操作で再現した不具合

[evidence/font-reliability-baseline/results.json](../evidence/font-reliability-baseline/results.json) とその fresh/copied 画像を参照。

| 操作 | 0.3.0 の実際 | 原因と対応 |
| --- | --- | --- |
| 見本文字を入力 | fresh/copied 両方で TypeError、見本が更新されない | querySelector の単一要素を for-of した。登録済み行だけの字形更新に再構成 |
| Windows 書体を選び次の文字を作る | Georgia 適用後も新しい文字は Yu Gothic | addText の固定書体を廃止し、選択既定書体を文書依存情報とともに引き継ぐ |
| Google 書体を選び次の文字を作る | 既存 Zen Antique Soft の登録後も次の文字は Yu Gothic | 同上。設定にはバイトを重複保存せず、取得済みキャッシュから復元 |
| 未取得 Google 書体を見る | Abel の見本は Yu Gothic UI | 未取得時は未検証表示。公開書体の取得→実バイナリで見本→使うに分離 |
| Windows の Bold 等の見本 | family だけ設定され、太さ/斜体は見本へ伝達されない | 選んだ PostScript face を抽出・登録してその weight/style で描画 |

Windows 301 faces の列挙と既存 Google キャッシュの取得・登録は修正前の今回試験でも成功した。すべての取得経路が故障していたとは断定しない。元ユーザー操作の全障害条件が完全に再現できたとも主張しない。

## 実装境界

- desktop/fonts.cjs: 元ファイル・許諾・ハッシュを照合する既存取得を継承。旧キャッシュの原本を上書きせず cmap を読み取る。STADIO_FONT_OFFLINE=1 では cache miss 後のネイティブ HTTP を禁止する。
- desktop/font-binary.cjs: TTC face 抽出に Unicode cmap 4/12 の収録検査を追加。glyph 0 と範囲外を収録扱いにしない。
- production-ui.js: Windows も binary→FontFace.load を通す。ロード失敗・変更済み対象への遅延適用を成功表示しない。取得ライブラリ、実行中登録、作品内依存、次の文字の既定設定を分ける。
- font-coverage.js: 旧作品に埋め込まれた書体の cmap もローカルだけで照合する。書体にない文字はフォールバックを明示する。Google へ作品・見本文字を送らない。
- UI_TEXT.fontReliability に文言を統合。Windows のバイトを .dake へ埋め込まない。Google は既存の許諾・同一バイト埋込みを維持する。

## 修正後の検証

[scripts/check-font-reliability.cjs](../scripts/check-font-reliability.cjs) を同じ操作の検証として再利用する。EXEのパスを引数に指定でき、STADIO_QA_PROFILE_SOURCE で安全な全設定コピーを seed として指定できる。

1. Windows Georgia Regular / Bold の列挙、実見本、文字列変更、適用、新規文字。
2. Google Zen Maru Gothic の公開データ新規取得、実見本、適用、新規文字。
3. .dake を実保存し、別プロセスで再起動。ネイティブ HTTP を禁止して既定書体・Windows既存文字・Google文字を復元。
4. オフラインの取得済み一覧から適用し、別の作品タブでも新規文字に同じ書体を使用。
5. 実際の Fabric 文字オブジェクトを、同一書体バイナリから独立に登録した参照 FontFace と画像として比較し、完全一致、一般代替書体との不一致を記録。名前と fonts.check だけで判定しない。
6. FontFace.load の失敗を明示的に注入し、失敗表示・レイヤー保持を確認。この項目は現実の OS 障害再現とは区別する。

操作は renderer のマウス/キー入力、保存は起動引数で許可された試験 .dake へのネイティブ書込み。Windowsファイルダイアログの手動操作試験とは区別する。試験出力には画像、保存文書、画素ハッシュ、例外一覧を残す。最終 ZIP 展開版を検証するまで、ソース EXE の合格を最終配布版の合格とは呼ばない。

ソース検証結果: [evidence/font-reliability-source-results.json](../evidence/font-reliability-source-results.json)。最終EXEの結果は同じフォルダの font-reliability-packaged-results.json に記録する。

## 同名書体の版に関する境界

複数の作品を切り替えたときは、その作品が参照するバイナリ版を有効な FontFace として優先し、必要なら文字計測キャッシュを更新する。取得ライブラリのバイト自体を削除しない。一つの作品内で同じ family/style/重なるweightに異なるバイナリ版を同時に混ぜる機能は、この版では扱わない。文字ごとの独立alias設計は後続課題とし、判明した競合を黙って別字形へ置き換えない。旧作品の sha 不明なOSフォント参照について、過去に使われたバイナリ版の同定を保証しない。
