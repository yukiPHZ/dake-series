from __future__ import annotations

import argparse
import csv
from collections import Counter
from datetime import datetime, timedelta, timezone
from pathlib import Path


DEFAULT_INPUT = Path(__file__).resolve().parents[1] / "reports" / "stripe_payment_link_candidates.csv"
DEFAULT_OUTPUT_CSV = Path(__file__).resolve().parents[1] / "reports" / "stripe_payment_link_rollout_review.csv"
DEFAULT_OUTPUT_MD = Path(__file__).resolve().parents[1] / "reports" / "stripe_payment_link_rollout_review.md"
JST = timezone(timedelta(hours=9))

OUTPUT_FIELDS = [
    "id",
    "type",
    "title",
    "price",
    "currency",
    "current_payment_status",
    "booth_url",
    "github_release_url",
    "source_original",
    "review_result",
    "creation_method",
    "price_check",
    "tax_code_candidate",
    "tax_code_review_required",
    "stripe_product_name",
    "metadata_ready",
    "memo",
]


def read_rows(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as file:
        return list(csv.DictReader(file))


def write_csv(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=OUTPUT_FIELDS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def parse_price(value: str) -> int | None:
    text = str(value or "").replace(",", "").strip()
    if not text:
        return None
    try:
        return int(text)
    except ValueError:
        return None


def price_check(row: dict[str, str]) -> str:
    price = parse_price(row.get("price", ""))
    if price is None:
        return "price_missing"
    if price <= 0:
        return "price_review"
    return "price_ok"


def is_game(row: dict[str, str]) -> bool:
    text = " ".join(
        str(row.get(key, ""))
        for key in ("store_id", "id", "type", "title", "category")
    ).lower()
    return "game" in text or "ゲーム" in text


def metadata_ready(row: dict[str, str]) -> str:
    required = ["store_id", "type", "source_repo", "source_original", "store_url"]
    if any(not str(row.get(key, "")).strip() for key in required):
        return "no"
    optional = ["booth_url", "github_release_url"]
    if any(not str(row.get(key, "")).strip() for key in optional):
        return "partial"
    return "yes"


def review_row(row: dict[str, str]) -> dict[str, str]:
    item_id = row.get("store_id") or row.get("id") or ""
    item_type = row.get("type", "")
    current_payment_status = row.get("payment_status", "")
    price_state = price_check(row)
    game = is_game(row)
    tax_review = "yes" if game else "no"
    product_name = row.get("stripe_product_name") or row.get("title") or item_id

    if current_payment_status == "preparing":
        review_result = "hold"
        creation_method = "hold"
        memo = "payment_status=preparing; BOOTH URL and sales flow are not confirmed"
    elif price_state != "price_ok":
        review_result = "hold"
        creation_method = "hold"
        memo = f"{price_state}; do not create Stripe Payment Link before price review"
    elif item_type == "pack":
        review_result = "manual"
        creation_method = "manual_dashboard"
        memo = (
            "Pack item; create manually first because github_release_url is empty "
            "and pack_ready, BOOTH flow, and post-purchase guidance need individual review"
        )
    elif item_type == "app" and row.get("booth_url"):
        review_result = "create"
        creation_method = "api_candidate"
        if game:
            memo = "single DAKE game app; API candidate, but tax_code review is required"
        else:
            memo = "single DAKE app; API candidate with price, BOOTH URL, and source_original present"
    elif item_type == "app":
        review_result = "hold"
        creation_method = "hold"
        memo = "single DAKE app but BOOTH URL is missing; sales flow must be confirmed"
    else:
        review_result = "exclude"
        creation_method = "none"
        memo = "not covered by Phase 12 rollout rules"

    source_memo = row.get("memo", "").strip()
    if source_memo and source_memo not in memo:
        memo = f"{memo}; source memo: {source_memo}"

    return {
        "id": item_id,
        "type": item_type,
        "title": row.get("title", ""),
        "price": row.get("price", ""),
        "currency": row.get("currency", ""),
        "current_payment_status": current_payment_status,
        "booth_url": row.get("booth_url", ""),
        "github_release_url": row.get("github_release_url", ""),
        "source_original": row.get("source_original", ""),
        "review_result": review_result,
        "creation_method": creation_method,
        "price_check": price_state,
        "tax_code_candidate": row.get("tax_code_candidate", ""),
        "tax_code_review_required": tax_review,
        "stripe_product_name": product_name,
        "metadata_ready": metadata_ready(row),
        "memo": memo,
    }


def md(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("|", "\\|")
    text = "<br>".join(part.strip() for part in text.split("\n") if part.strip())
    return text or "-"


def yen(value: str) -> str:
    price = parse_price(value)
    return f"{price:,} JPY" if price is not None else "-"


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


def write_report(path: Path, rows: list[dict[str, str]], input_path: Path, output_csv: Path) -> None:
    counts = Counter(row["review_result"] for row in rows)
    methods = Counter(row["creation_method"] for row in rows)
    price_counts = Counter(row["price_check"] for row in rows)
    tax_review_rows = [row for row in rows if row["tax_code_review_required"] == "yes"]
    api_rows = [row for row in rows if row["review_result"] == "create"]
    manual_rows = [row for row in rows if row["review_result"] == "manual"]
    hold_rows = [row for row in rows if row["review_result"] == "hold"]
    exclude_rows = [row for row in rows if row["review_result"] == "exclude"]

    content = f"""# Stripe Payment Link Rollout Review

Generated: {datetime.now(JST).strftime('%Y-%m-%d %H:%M:%S %z')}

## 目的

Phase 11 の `stripe_payment_link_candidates.csv` をレビューし、Stripe Payment Link作成対象を `create` / `manual` / `hold` / `exclude` に分類する。今回はStripe API、Stripe Secret Key、Payment Link / Product / Price作成は行わない。

## 現状

- Stripe未対応候補: {len(rows)}
- API作成候補: {counts.get('create', 0)}
- 手動作成候補: {counts.get('manual', 0)}
- 保留: {counts.get('hold', 0)}
- 対象外: {counts.get('exclude', 0)}
- price_missing: {price_counts.get('price_missing', 0)}
- tax_code_review_required: {len(tax_review_rows)}

## 入力ファイル

- `{input_path.relative_to(input_path.parents[2]).as_posix()}`
- `tools/generated/store_products.generated.json`
- `tools/reports/stripe_payment_links_full_rollout_plan.md`
- `00_core/DAKE_STORE_GENERATED_SPEC.md`
- `00_core/DAKE_STORE_OPERATION_RULE.md`

## 出力ファイル

- `{output_csv.relative_to(output_csv.parents[2]).as_posix()}`
- `{path.relative_to(path.parents[2]).as_posix()}`

## 集計

{table(['key', 'count'], [
    ['Stripe未対応', len(rows)],
    ['API作成候補', counts.get('create', 0)],
    ['手動作成候補', counts.get('manual', 0)],
    ['保留', counts.get('hold', 0)],
    ['対象外', counts.get('exclude', 0)],
    ['creation_method.api_candidate', methods.get('api_candidate', 0)],
    ['creation_method.manual_dashboard', methods.get('manual_dashboard', 0)],
    ['creation_method.hold', methods.get('hold', 0)],
    ['price_ok', price_counts.get('price_ok', 0)],
    ['price_missing', price_counts.get('price_missing', 0)],
    ['price_review', price_counts.get('price_review', 0)],
])}

## API作成候補

{table(['id', 'type', 'title', 'price', 'tax_code', 'memo'], [
    [row['id'], row['type'], row['title'], yen(row['price']), row['tax_code_candidate'], row['memo']]
    for row in api_rows
])}

## 手動作成候補

{table(['id', 'type', 'title', 'reason'], [
    [row['id'], row['type'], row['title'], row['memo']]
    for row in manual_rows
])}

## 保留候補

{table(['id', 'type', 'title', 'reason'], [
    [row['id'], row['type'], row['title'], row['memo']]
    for row in hold_rows
])}

## 対象外候補

{table(['id', 'type', 'title', 'reason'], [
    [row['id'], row['type'], row['title'], row['memo']]
    for row in exclude_rows
])}

## tax_codeレビュー対象

{table(['id', 'type', 'title', 'candidate', 'reason'], [
    [row['id'], row['type'], row['title'], row['tax_code_candidate'], row['memo']]
    for row in tax_review_rows
])}

## Pack商品の扱い

`DAKE_Pack_Document` と `DAKE_Pack_Memo` は手動作成候補にする。理由は、`github_release_url` がなく、Pack ZIP、`pack_ready/`、BOOTH導線、購入後案内を個別確認した方が安全なため。最初のPack 2件はStripe Dashboardで手動作成し、動線が固まった後にAPI化を検討する。

## metadata方針

維持するmetadataキー:

- `dake_item_id`
- `dake_type`
- `source_repo`
- `source_original`
- `store_url`
- `booth_url`
- `github_release_url`

空欄の値は空欄またはnull相当でよい。ただしSecret、個人情報、購入者情報、内部トークンはmetadataに入れない。

## API作成前の安全ルール

- Stripe Secret Keyは環境変数のみで扱う。
- Secretをpublic JS、generated JSON、GitHub repo、Store静的ファイル、`ORIGINAL.md` に入れない。
- Store側にカード情報、購入者DB、Stripe Secretを置かない。
- 最初は必ずStripe test modeで実行する。
- dry-runで作成予定Product / Price / Payment Linkを出し、人間レビュー後に本実行する。
- `metadata.dake_item_id` で既存Product重複を避ける。
- 作成後はPayment Link URLを対象商品の `ORIGINAL.md` に戻し、generated JSONを再生成する。

## 次Phase提案

1. API作成候補45件から先行10件を選ぶ。
2. 先行10件についてStripe test mode + dry-run用スクリプトを設計する。
3. Pack 2件はStripe Dashboardで手動作成し、購入後案内と配布導線を確認する。
4. `video_shorts_cut` はBOOTH URLと販売導線が確定するまで保留する。
5. Payment Link作成後、URLだけを各 `ORIGINAL.md` へ戻し、Store生成・同期は別Phaseで行う。
"""
    path.write_text(content, encoding="utf-8", newline="\n")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Review Stripe Payment Link rollout candidates without calling Stripe APIs.")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--output-csv", type=Path, default=DEFAULT_OUTPUT_CSV)
    parser.add_argument("--output-md", type=Path, default=DEFAULT_OUTPUT_MD)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_rows = read_rows(args.input)
    reviewed = [review_row(row) for row in input_rows]
    write_csv(args.output_csv, reviewed)
    write_report(args.output_md, reviewed, args.input, args.output_csv)
    counts = Counter(row["review_result"] for row in reviewed)
    print(f"input_rows={len(input_rows)}")
    print(f"output_rows={len(reviewed)}")
    print(
        "review_result: "
        + ", ".join(f"{key}={counts.get(key, 0)}" for key in ("create", "manual", "hold", "exclude"))
    )
    print(f"wrote {args.output_csv}")
    print(f"wrote {args.output_md}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
