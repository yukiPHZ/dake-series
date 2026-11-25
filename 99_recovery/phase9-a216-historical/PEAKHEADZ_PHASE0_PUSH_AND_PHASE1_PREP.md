# PEAKHEADZ Phase0 Push / Phase1 Prep

作成日: 2026-06-02

## 目的

Phase0 audit commit の push 失敗理由を整理し、Phase1 の凍結対象を確認する。

この準備では凍結変更、削除、移動、remote追加、pushは行わない。

## Git状態確認

対象cwd:

`C:\Users\yukiz\Documents\New project`

確認結果:

- 現在ブランチ: `main`
- `git status`: 既存の未stage変更と多数の未追跡ファイルあり
- 代表的な既存未stage変更:
  - `DAKE_QPSC_Dashboard/README.md`
  - `DAKE_QPSC_Dashboard/main.py`
  - `_codex_work/DAKE_series/01_apps/DAKE_BOOTH_Assist/main.py`
- Phase0 audit commit: `a6fe3bb Add PEAKHEADZ_ROOT phase0 audit`
- 直近HEAD: `be66d81 Freeze web dashboard and add web index`
- branch一覧:
  - `main`: `be66d81 Freeze web dashboard and add web index`
  - `master`: `a6fe3bb Add PEAKHEADZ_ROOT phase0 audit`
- remote: 未設定

直近log:

```text
be66d81 Freeze web dashboard and add web index
a6fe3bb Add PEAKHEADZ_ROOT phase0 audit
04e5b50 Add QPSC Dashboard
4431455 Add DAKE_Mansion_Schedule
8405948 Add DakeBOOTH Assist
```

## Push可否

`C:\Users\yukiz\Documents\New project` は remote が未設定のため push 不可。

これは upstream 未設定だけの状態ではなく、`git remote -v` が空の状態。勝手に remote は追加しない。

push可能にするには、人間が送信先remoteを決めた後に `git remote add` または既存remote設定を行う必要がある。

## Phase1凍結対象候補

| 対象 | 現状 | Phase1判断 |
| --- | --- | --- |
| `DAKE_Brainz_Search` | `status: draft`、launcher/site非表示 | 凍結対象 |
| `DAKE_Brainz_OIKAWA` | `status: experimental`、launcher/site非表示 | 凍結対象 |
| `DAKE_Wake_Brainz` | `status: draft`、遠隔起動・状態確認 | 今回は凍結対象外、後続判断 |
| `DAKE_Approve_Brainz` | `status: frozen` | 追加対応不要 |
| QPSC関連README | BRAINZ/OIKAWA/Wakeの旧役割記述あり | Phase1では変更しない |
| Launcher/Dashboard表示参照 | READMEの `DAKE_META` を読む設計 | BRAINZ/OIKAWAの `show_*: false` 維持で対応 |

## DAKE_Wake_Brainz 判断材料

DAKE_Wake_Brainz は Wake on LAN と状態確認を担うため、今回の「遠隔操作は作らない」方針と衝突する可能性がある。

一方で、Phase1の明示対象は `DAKE_Brainz_Search` と `DAKE_Brainz_OIKAWA` のみ。Wakeを即凍結すると、既存の起床確認用途まで止める判断になる。

扱い案:

- 推奨: Phase1では変更せず、凍結対象外・後続判断とする
- 後続で確認すること: Wake on LANを完全停止するのか、既存メモとして残すだけにするのか
- 禁止: 今回のPhase1で削除、README変更、status変更を行うこと

## Phase1準備結論

Phase1では `DAKE_Brainz_Search` と `DAKE_Brainz_OIKAWA` のREADMEだけを凍結対象にする。

`DAKE_Wake_Brainz` と `DAKE_Approve_Brainz` は変更しない。

`New project` 側は remote 未設定のため、成果物commit後も push は不可。
