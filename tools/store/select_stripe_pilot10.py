from __future__ import annotations

import argparse
import csv
from collections import Counter
from datetime import datetime, timedelta, timezone
from pathlib import Path


REPORTS_DIR = Path(__file__).resolve().parents[1] / "reports"
DEFAULT_REVIEW_CSV = REPORTS_DIR / "stripe_payment_link_rollout_review.csv"
DEFAULT_CANDIDATE_CSV = REPORTS_DIR / "stripe_payment_link_candidates.csv"
DEFAULT_OUTPUT_CSV = REPORTS_DIR / "stripe_payment_link_pilot10_selection.csv"
DEFAULT_OUTPUT_MD = REPORTS_DIR / "stripe_payment_link_pilot10_selection.md"
JST = timezone(timedelta(hours=9))

PILOT_SELECTION = [
    (
        "dake_pdf_viewer",
        "PDF系",
        "PDF閲覧で内容が分かりやすく、既存Stripe対応PDF系と近い通常アプリ。",
        "PDF viewer baseline for the pilot.",
    ),
    (
        "dake_pdf_reorder",
        "PDF系",
        "PDFページ並び替えで用途が明確。PDF系の別操作パターンを確認できる。",
        "PDF page reordering workflow.",
    ),
    (
        "dake_pdf_splitone",
        "PDF系",
        "PDF分割系の代表として選定。既存PDF結合/圧縮とは異なる操作カテゴリ。",
        "PDF split workflow.",
    ),
    (
        "dake_image_heictojpg",
        "画像系",
        "HEICからJPGへの変換で商品内容が直感的。画像変換系の代表。",
        "Image conversion workflow.",
    ),
    (
        "dake_image_topdf",
        "画像系",
        "画像PDF化でPDF/画像の横断カテゴリを確認できる。",
        "Image to PDF workflow.",
    ),
    (
        "DAKE_Sticky_Memo",
        "メモ系",
        "付箋メモは軽量で用途が明確。メモ系の代表として適している。",
        "Memo app baseline.",
    ),
    (
        "DAKE_Mail_Draft",
        "メール系",
        "メール下書きは自動送信しない補助ツールで、説明・免責が比較的整理しやすい。",
        "Mail draft helper; no automatic send.",
    ),
    (
        "DAKE_Backup",
        "作業補助系",
        "バックアップ補助は実務用途が明確。購入前注意書きは維持する。",
        "Backup helper; keep backup/disclaimer wording.",
    ),
    (
        "dake_folder_list",
        "ファイル整理系",
        "フォルダ一覧化で配布導線・商品説明が分かりやすい。",
        "File listing workflow.",
    ),
    (
        "dake_year_age",
        "年数/計算系",
        "年数計算系を1件だけ入れる。高度な税務判断ではなく、比較的低リスクな計算枠。",
        "Limited calculation/date pilot item.",
    ),
]

EXPLICIT_EXCLUSIONS = {
    "DAKE_Pack_Document": "Pack商品は手動作成候補。github_release_urlがなく、Pack ZIP/pack_ready/BOOTH導線/購入後案内を個別確認する。",
    "DAKE_Pack_Memo": "Pack商品は手動作成候補。github_release_urlがなく、Pack ZIP/pack_ready/BOOTH導線/購入後案内を個別確認する。",
    "video_shorts_cut": "payment_status=preparing。BOOTH URLなし、販売導線未確定のため保留。",
    "game_alien_road": "ゲーム系はtax_code要確認のため後続確認へ回す。",
    "game_diver_catch": "ゲーム系はtax_code要確認のため後続確認へ回す。",
}

OUTPUT_FIELDS = [
    "id",
    "type",
    "title",
    "price",
    "currency",
    "category",
    "source_original",
    "booth_url",
    "github_release_url",
    "stripe_product_name",
    "tax_code_candidate",
    "metadata_ready",
    "selection_reason",
    "notes",
    "review_result",
    "creation_method",
    "price_check",
]


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        return list(csv.DictReader(file))


def write_csv(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=OUTPUT_FIELDS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def by_id(rows: list[dict[str, str]], id_key: str) -> dict[str, dict[str, str]]:
    return {row[id_key]: row for row in rows if row.get(id_key)}


def is_game(row: dict[str, str]) -> bool:
    text = " ".join(
        str(row.get(key, ""))
        for key in ("id", "type", "title", "category")
    ).lower()
    return "game" in text or "ゲーム" in text


def build_selection(review_rows: list[dict[str, str]], candidate_rows: list[dict[str, str]]) -> list[dict[str, str]]:
    review_by_id = by_id(review_rows, "id")
    candidate_by_id = by_id(candidate_rows, "store_id")
    output: list[dict[str, str]] = []

    for item_id, reason, selection_reason, notes in PILOT_SELECTION:
        if item_id not in review_by_id:
            raise ValueError(f"pilot item not found in rollout review: {item_id}")
        review = review_by_id[item_id]
        candidate = candidate_by_id.get(item_id, {})

        if review.get("review_result") != "create":
            raise ValueError(f"pilot item must be review_result=create: {item_id}")
        if review.get("creation_method") != "api_candidate":
            raise ValueError(f"pilot item must be creation_method=api_candidate: {item_id}")
        if review.get("price_check") != "price_ok":
            raise ValueError(f"pilot item must have price_ok: {item_id}")
        if review.get("type") == "pack":
            raise ValueError(f"pilot item must not be a pack: {item_id}")
        if is_game({**candidate, **review}):
            raise ValueError(f"pilot item must not be a game item: {item_id}")
        if review.get("tax_code_review_required") == "yes":
            raise ValueError(f"pilot item must not require tax_code review: {item_id}")
        if not review.get("booth_url"):
            raise ValueError(f"pilot item must have booth_url: {item_id}")

        output.append(
            {
                "id": item_id,
                "type": review.get("type", ""),
                "title": review.get("title", ""),
                "price": review.get("price", ""),
                "currency": review.get("currency", ""),
                "category": candidate.get("category", ""),
                "source_original": review.get("source_original", ""),
                "booth_url": review.get("booth_url", ""),
                "github_release_url": review.get("github_release_url", ""),
                "stripe_product_name": review.get("stripe_product_name", ""),
                "tax_code_candidate": review.get("tax_code_candidate", ""),
                "metadata_ready": review.get("metadata_ready", ""),
                "selection_reason": selection_reason,
                "notes": f"{reason}; {notes}",
                "review_result": review.get("review_result", ""),
                "creation_method": review.get("creation_method", ""),
                "price_check": review.get("price_check", ""),
            }
        )

    if len(output) != 10:
        raise ValueError(f"pilot selection must contain 10 rows, got {len(output)}")
    return output


def md(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("|", "\\|")
    text = "<br>".join(part.strip() for part in text.split("\n") if part.strip())
    return text or "-"


def yen(value: str) -> str:
    try:
        return f"{int(str(value).replace(',', '')):,} JPY"
    except ValueError:
        return "-"


def table(headers: list[str], rows: list[list[object]]) -> str:
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join("---" for _ in headers) + " |",
    ]
    if not rows:
        rows = [["-" for _ in headers]]
    for row in rows:
        lines.append("| " + " | ".join(md(cell) for cell in row) + " |")
    return "\n".join(lines)


def write_report(
    path: Path,
    selected: list[dict[str, str]],
    review_rows: list[dict[str, str]],
    review_csv: Path,
    candidates_csv: Path,
    output_csv: Path,
) -> None:
    review_by_id = by_id(review_rows, "id")
    genre_counts = Counter(row["notes"].split(";", 1)[0] for row in selected)
    excluded_rows = [
        [
            item_id,
            review_by_id.get(item_id, {}).get("type", "-"),
            review_by_id.get(item_id, {}).get("title", "-"),
            reason,
        ]
        for item_id, reason in EXPLICIT_EXCLUSIONS.items()
    ]
    tax_excluded = [
        [
            item_id,
            review_by_id.get(item_id, {}).get("type", "-"),
            review_by_id.get(item_id, {}).get("title", "-"),
            review_by_id.get(item_id, {}).get("tax_code_candidate", "-"),
            reason,
        ]
        for item_id, reason in EXPLICIT_EXCLUSIONS.items()
        if item_id.startswith("game_")
    ]

    content = f"""# Stripe Payment Link Pilot 10 Selection

Generated: {datetime.now(JST).strftime('%Y-%m-%d %H:%M:%S %z')}

## 目的

Phase 12で分類したStripe Payment Link API作成候補45件から、次Phaseのtest mode / dry-run検証に使う先行10件を選定する。今回はStripe API、Stripe Secret Key、Product / Price / Payment Link作成は行わない。

## 入力ファイル

- `{review_csv.relative_to(review_csv.parents[2]).as_posix()}`
- `{candidates_csv.relative_to(candidates_csv.parents[2]).as_posix()}`
- `tools/reports/stripe_payment_links_full_rollout_plan.md`
- `tools/generated/store_products.generated.json`

## 選定方針

- `review_result=create`
- `creation_method=api_candidate`
- `price_check=price_ok`
- `booth_url` あり
- `metadata_ready=yes`
- Pack、preparing、ゲーム系、tax_codeレビュー必須の商品は除外
- PDF、画像、メモ、メール、作業補助、ファイル整理、年数/計算系に分散
- 計算系は先行10件では1件に留める

## 除外したもの

{table(['id', 'type', 'title', 'reason'], excluded_rows)}

## 先行10件

{table(['id', 'type', 'title', 'price', 'category', 'reason'], [
    [row['id'], row['type'], row['title'], yen(row['price']), row['category'], row['selection_reason']]
    for row in selected
])}

## ジャンル分散

{table(['genre', 'count'], [[genre, count] for genre, count in genre_counts.items()])}

## tax_code確認

選定10件はすべて `tax_code_review_required=no` の通常候補。tax_code候補は全件 `txcd_10202003`。ゲーム系2件は今回除外し、後続のtax_code確認へ回す。

{table(['id', 'type', 'title', 'candidate', 'reason'], tax_excluded)}

## metadata確認

選定10件はすべて `metadata_ready=yes`。維持するmetadataキーは `dake_item_id`, `dake_type`, `source_repo`, `source_original`, `store_url`, `booth_url`, `github_release_url`。

## 次Phaseで使う想定

- Stripe test mode用dry-runスクリプトの入力リストとして `tools/reports/stripe_payment_link_pilot10_selection.csv` を使う。
- dry-runではProduct / Price / Payment Link作成予定だけを表示し、Stripe APIは本実行しない。
- metadataで `dake_item_id` を紐付け、既存Product重複を避ける。
- 作成後に保存するのはPayment Link URLと必要最小限のStripe IDだけにする。

## 今回やらなかったこと

- Stripe API実行
- Stripe Secret Key使用
- Product作成
- Price作成
- Payment Link作成
- `ORIGINAL.md`更新
- generated JSON更新
- dake-store-site同期
- Store本番反映
- Pack商品のStripe作成
- ゲーム系商品のStripe作成

## 次Phase提案

1. `stripe_payment_link_pilot10_selection.csv` を入力に、test mode + dry-run専用スクリプトを設計する。
2. dry-run出力でProduct名、価格、currency、tax_code候補、metadataを確認する。
3. Stripe API本実行はさらに次のPhaseで、人間レビュー後に限定して行う。
4. Pack 2件はDashboard手動作成で先に導線を確認する。
5. ゲーム系2件はtax_code候補を確認してから別枠で進める。
"""
    path.write_text(content, encoding="utf-8", newline="\n")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Select 10 pilot Stripe Payment Link candidates without calling Stripe APIs.")
    parser.add_argument("--review-csv", type=Path, default=DEFAULT_REVIEW_CSV)
    parser.add_argument("--candidate-csv", type=Path, default=DEFAULT_CANDIDATE_CSV)
    parser.add_argument("--output-csv", type=Path, default=DEFAULT_OUTPUT_CSV)
    parser.add_argument("--output-md", type=Path, default=DEFAULT_OUTPUT_MD)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    review_rows = read_csv(args.review_csv)
    candidate_rows = read_csv(args.candidate_csv)
    selected = build_selection(review_rows, candidate_rows)
    write_csv(args.output_csv, selected)
    write_report(args.output_md, selected, review_rows, args.review_csv, args.candidate_csv, args.output_csv)
    print(f"selected={len(selected)}")
    print("ids=" + ", ".join(row["id"] for row in selected))
    print(f"wrote {args.output_csv}")
    print(f"wrote {args.output_md}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
