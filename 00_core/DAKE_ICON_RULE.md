# DAKE_ICON_RULE

DAKEシリーズの統一アイコンルールです。

## 正本

- 共通アイコン: `02_assets/dake_icon.ico`
- 全アプリはこのicoを参照する。
- pngやsvgは補助素材として扱い、Windows exeの正本はicoにする。

## アプリ実装

起動時は候補パスから共通アイコンを探す。

候補例:

```python
base / ".." / ".." / "02_assets" / "dake_icon.ico"
base / ".." / ".." / ".." / "02_assets" / "dake_icon.ico"
Path(__file__).resolve().parent / ".." / ".." / "02_assets" / "dake_icon.ico"
```

ルール:

- アイコンが見つからなくてもアプリ起動は止めない。
- `iconbitmap` の失敗は握りつぶしてよい。
- アイコンは機能説明の代わりにしない。

## build.bat

PyInstallerでは必ず共通アイコンを指定する。

```bat
--icon=..\..\02_assets\dake_icon.ico ^
```

## 禁止

- アプリごとの個別アイコンを作らない。
- PDF、画像、メールなど機能別アイコンで識別しない。
- Releaseごとにアイコンを変えない。
- アイコン未設定のまま配布しない。

## 理由

- DAKEシリーズとして一目でまとまる。
- ランチャーやdakeapp.comで統一感を保てる。
- 個別アイコン管理の手間を増やさない。
- 小さなアプリ群としての軽さを守る。
