from __future__ import annotations

import json
import re
from collections import Counter
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any


SOURCE_POLICY = (
    "ORIGINAL.md is the source of truth. "
    "This file is generated and must not be edited manually."
)

NULL_TOKENS = {
    "",
    "-",
    "なし",
    "無し",
    "該当なし",
    "未確定",
    "未設定",
    "要確認",
    "既存ファイルに記載なし",
}

EXCLUDED_STATUS = {"draft", "frozen", "prototype", "internal", "experimental"}


def repo_root() -> Path:
    return Path(__file__).resolve().parents[2]


ROOT = repo_root()
GENERATED_DIR = ROOT / "tools" / "generated"
REPORT_DIR = ROOT / "tools" / "reports"
OUTPUT_JSON = GENERATED_DIR / "store_products.generated.json"
OUTPUT_REPORT = REPORT_DIR / "original_phase52_stripe_payment_link_ready.md"
SHIMARISU_ORIGINAL = Path(r"C:\Users\yukiz\devlop\SHIMARISU\ORIGINAL.md")


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def source_label(path: Path, source_repo: str) -> str:
    if source_repo == "DAKE_series":
        return path.relative_to(ROOT).as_posix()
    return path.as_posix()


def normalize_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, list):
        cleaned = [normalize_value(item) for item in value]
        return [item for item in cleaned if item is not None]
    text = str(value).strip()
    if len(text) >= 2 and text.startswith("`") and text.endswith("`"):
        text = text[1:-1].strip()
    text = re.sub(r"^```(?:text|json)?\s*", "", text.strip(), flags=re.MULTILINE)
    text = re.sub(r"\s*```$", "", text.strip(), flags=re.MULTILINE)
    text = text.strip()
    compact = re.sub(r"\s+", "", text)
    if compact in NULL_TOKENS:
        return None
    if any(token in text for token in ["未確定", "未設定", "要確認", "既存ファイルに記載なし"]):
        return None
    return text


def normalize_price(value: Any) -> int | None:
    text = normalize_value(value)
    if text is None:
        return None
    match = re.search(r"([0-9][0-9,]*)\s*円", str(text))
    if not match:
        match = re.search(r"([0-9][0-9,]*)", str(text))
    if not match:
        return None
    return int(match.group(1).replace(",", ""))


def clean_text_block(lines: list[str]) -> str | None:
    text = "\n".join(lines).strip()
    text = re.sub(r"^```(?:text|json)?\s*", "", text, flags=re.MULTILINE).strip()
    text = re.sub(r"\s*```$", "", text, flags=re.MULTILINE).strip()
    return normalize_value(text)


def split_sections(markdown: str) -> dict[str, list[str]]:
    sections: dict[str, list[str]] = {}
    current: str | None = None
    for line in markdown.splitlines():
        match = re.match(r"^##\s+(.+?)\s*$", line)
        if match and not line.startswith("###"):
            current = match.group(1).strip()
            sections[current] = []
            continue
        if current is not None:
            sections[current].append(line)
    return sections


def parse_bullets(lines: list[str]) -> dict[str, Any]:
    values: dict[str, Any] = {}
    current_key: str | None = None
    current_lines: list[str] = []

    def flush() -> None:
        nonlocal current_key, current_lines
        if current_key is None:
            return
        if current_lines and all(re.match(r"^\s*-\s+", line) for line in current_lines if line.strip()):
            values[current_key] = [
                normalize_value(re.sub(r"^\s*-\s+", "", line))
                for line in current_lines
                if normalize_value(re.sub(r"^\s*-\s+", "", line)) is not None
            ]
        else:
            values[current_key] = clean_text_block(current_lines)
        current_key = None
        current_lines = []

    for line in lines:
        match = re.match(r"^-\s+([^:：]+)[:：]\s*(.*)$", line)
        if match:
            flush()
            current_key = match.group(1).strip()
            first = match.group(2).strip()
            current_lines = [first] if first else []
            continue
        if current_key is not None:
            current_lines.append(line)
    flush()
    return values


def section_text(sections: dict[str, list[str]], name: str) -> str | None:
    return clean_text_block(sections.get(name, []))


def first_non_null(*values: Any) -> Any:
    for value in values:
        normalized = normalize_value(value)
        if normalized is not None:
            return normalized
    return None


def first_path(value: Any) -> str | None:
    normalized = normalize_value(value)
    if normalized is None:
        return None
    text = str(normalized)
    if " / " in text:
        text = text.split(" / ", 1)[0].strip()
    if "," in text:
        text = text.split(",", 1)[0].strip()
    text = re.sub(r"[（(].*?[）)]", "", text).strip()
    return normalize_value(text)


def url_from_text(value: Any) -> str | None:
    normalized = normalize_value(value)
    if normalized is None:
        return None
    match = re.search(r"https?://\S+", str(normalized))
    if not match:
        return None
    return match.group(0).rstrip("`),。")


def metadata_line(markdown: str, key: str) -> str | None:
    match = re.search(rf"^\s*[-*]\s*{re.escape(key)}\s*:\s*(.+?)\s*$", markdown, re.MULTILINE)
    if not match:
        return None
    return normalize_value(match.group(1))


def parse_table(lines: list[str]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    table_lines = [line.strip() for line in lines if line.strip().startswith("|")]
    if len(table_lines) < 2:
        return rows
    headers = [cell.strip() for cell in table_lines[0].strip("|").split("|")]
    for line in table_lines[2:]:
        cells = [cell.strip() for cell in line.strip("|").split("|")]
        if len(cells) != len(headers):
            continue
        row = {headers[index]: normalize_value(cell) or "" for index, cell in enumerate(cells)}
        if any(row.values()):
            rows.append(row)
    return rows


def classify_source(path: Path) -> tuple[str, str]:
    try:
        relative = path.relative_to(ROOT)
    except ValueError:
        return "shimarisu_pack", "SHIMARISU"
    parts = relative.parts
    if parts[0] == "01_apps":
        return "app", "DAKE_series"
    if parts[0] == "04_packs":
        return "pack", "DAKE_series"
    return "unknown", "DAKE_series"



def payment_status(stripe_payment_link: str | None, booth_url: str | None) -> str:
    if stripe_payment_link:
        return "stripe_ready"
    if booth_url:
        return "booth_only"
    return "preparing"


def stripe_payment_link_from(store: dict[str, Any], booth: dict[str, Any], markdown: str) -> str | None:
    return first_non_null(
        url_from_text(store.get("Stripe Payment Link")),
        url_from_text(store.get("stripe_payment_link")),
        url_from_text(store.get("Payment Link")),
        url_from_text(store.get("Stripe URL")),
        url_from_text(booth.get("Stripe Payment Link")),
        url_from_text(booth.get("stripe_payment_link")),
        url_from_text(metadata_line(markdown, "stripe_payment_link")),
    )


def extract_item(path: Path) -> tuple[dict[str, Any] | None, dict[str, str] | None, dict[str, Any]]:
    item_type, source_repo = classify_source(path)
    markdown = read_text(path)
    sections = split_sections(markdown)
    basic = parse_bullets(sections.get("基本情報", []))
    store = parse_bullets(sections.get("Store表示用情報", []))
    booth = parse_bullets(sections.get("booth_product生成用情報", []))
    distribution = parse_bullets(sections.get("配布・ダウンロード方針", []))

    item_id = first_non_null(basic.get("app_id"), basic.get("pack_id"), path.parent.name)
    status = first_non_null(basic.get("status"))
    if status is None:
        return None, {
            "source_original": source_label(path, source_repo),
            "reason": "status missing",
        }, {"status": status, "type": item_type}
    if str(status).lower() in EXCLUDED_STATUS or str(status).lower() != "available":
        return None, {
            "source_original": source_label(path, source_repo),
            "reason": f"status={status}",
        }, {"status": status, "type": item_type}

    title = first_non_null(store.get("商品名"), basic.get("title"), booth.get("商品名"), item_id)
    short_title = first_non_null(basic.get("short_title"), title)
    catch = first_non_null(store.get("キャッチ"), store.get("キャッチ補足"), section_text(sections, "公開用説明の元情報"))
    description = first_non_null(store.get("説明"), section_text(sections, "公開用説明の元情報"), booth.get("商品紹介文"))
    price = normalize_price(first_non_null(store.get("価格"), basic.get("price"), booth.get("価格案")))
    category = first_non_null(basic.get("category"))
    tags = normalize_value(booth.get("タグ")) or []
    if isinstance(tags, str):
        tags = [line.strip("- ・") for line in tags.splitlines() if line.strip("- ・")]

    store_image = first_path(store.get("画像"))
    booth_image = first_path(booth.get("商品画像"))
    booth_support = first_path(booth.get("補助画像"))
    image = first_non_null(store_image, booth_image, booth_support)
    thumbnail = first_non_null(booth_image, store_image, booth_support)

    store_download = first_non_null(store.get("ダウンロード導線"))
    download_url = url_from_text(store_download)
    download_type = None
    if download_url:
        if "booth.pm" in download_url:
            download_type = "booth"
        elif "github.com" in download_url:
            download_type = "github_release"
        else:
            download_type = "external"

    booth_url = first_non_null(
        url_from_text(booth.get("BOOTH URL")),
        url_from_text(distribution.get("BOOTH")),
        url_from_text(distribution.get("BOOTH URL")),
        url_from_text(metadata_line(markdown, "booth_url")),
    )
    github_release_url = first_non_null(
        url_from_text(booth.get("GitHub Release")),
        url_from_text(distribution.get("GitHub Release")),
    )
    support_policy = first_non_null(store.get("サポート方針"))
    disclaimer = section_text(sections, "免責・注意事項")
    stripe_payment_link = stripe_payment_link_from(store, booth, markdown)

    included_items: list[dict[str, str]] = []
    if item_type in {"pack", "shimarisu_pack"}:
        included_items = parse_table(sections.get("同梱アプリ・構成物", []))

    item = {
        "id": item_id,
        "type": item_type,
        "source_repo": source_repo,
        "source_original": source_label(path, source_repo),
        "title": title,
        "short_title": short_title,
        "catch": catch,
        "description": description,
        "price": price,
        "currency": "JPY",
        "status": status,
        "category": category,
        "tags": tags,
        "image": image,
        "thumbnail": thumbnail,
        "download_type": download_type,
        "download_url": download_url,
        "booth_url": booth_url,
        "github_release_url": github_release_url,
        "store_url": None,
        "support_policy": support_policy,
        "disclaimer": disclaimer,
        "included_items": included_items,
        "stripe_payment_link": stripe_payment_link,
        "payment_status": payment_status(stripe_payment_link, booth_url),
        "source_kind": "original",
        "is_generated": True,
        "stripe_price_id": None,
    }
    for key, value in list(item.items()):
        if isinstance(value, str):
            item[key] = normalize_value(value)
    return item, None, {"status": status, "type": item_type}


def discover_originals() -> tuple[list[Path], list[dict[str, str]]]:
    originals: list[Path] = []
    skipped: list[dict[str, str]] = []
    originals.extend(sorted((ROOT / "01_apps").glob("*/ORIGINAL.md")))
    originals.extend(sorted((ROOT / "04_packs").glob("*/ORIGINAL.md")))
    if SHIMARISU_ORIGINAL.exists():
        originals.append(SHIMARISU_ORIGINAL)
    else:
        skipped.append(
            {
                "source_original": SHIMARISU_ORIGINAL.as_posix(),
                "reason": "missing SHIMARISU ORIGINAL",
            }
        )
    return originals, skipped


def build_report(
    originals: list[Path],
    items: list[dict[str, Any]],
    skipped: list[dict[str, str]],
    generated_at: str,
) -> str:
    type_counts = Counter(item["type"] for item in items)
    unresolved = Counter()
    for item in items:
        for key in ["download_url", "store_url", "support_policy", "image", "thumbnail", "stripe_payment_link", "stripe_price_id"]:
            if item.get(key) is None:
                unresolved[key] += 1

    payment_counts = Counter(item.get("payment_status") for item in items)
    shimarisu_result = "included" if any(item["type"] == "shimarisu_pack" for item in items) else "skipped"

    lines = [
        "# ORIGINAL Phase5-2 Stripe Payment Link Ready",
        "",
        "## 目的",
        "",
        "`ORIGINAL.md` 由来の Store 表示用 generated JSON に Stripe Payment Link 用フィールドを追加し、Store UI の購入導線出し分け準備を確認する。",
        "",
        "## 生成スクリプト",
        "",
        "- `tools/store/generate_store_products.py`",
        "",
        "## 生成ファイル",
        "",
        "- `tools/generated/store_products.generated.json`",
        "- `tools/generated/README.md`",
        "",
        "## 参照したORIGINAL",
        "",
        f"- discovered: {len(originals)}",
        "",
        "## 生成対象件数",
        "",
        f"- generated_at: {generated_at}",
        f"- items: {len(items)}",
        "",
        "## type別件数",
        "",
    ]
    for item_type in ["app", "pack", "shimarisu_pack"]:
        lines.append(f"- {item_type}: {type_counts.get(item_type, 0)}")
    lines.extend(["", "## skipped件数", "", f"- skipped: {len(skipped)}", ""])
    if skipped:
        lines.extend(["| source_original | reason |", "|---|---|"])
        for row in skipped:
            lines.append(f"| `{row['source_original']}` | {row['reason']} |")
        lines.append("")
    lines.extend(["## payment_status別件数", ""])
    for status_name in ["stripe_ready", "booth_only", "preparing", "free_download", "not_for_sale"]:
        lines.append(f"- {status_name}: {payment_counts.get(status_name, 0)}")
    lines.append("")
    lines.extend(["## 主な未確定項目", ""])
    for key, count in unresolved.items():
        lines.append(f"- {key}: {count}")
    lines.extend(
        [
            "",
            "## SHIMARISU Pack参照結果",
            "",
            f"- result: {shimarisu_result}",
            f"- source: `{SHIMARISU_ORIGINAL.as_posix()}`",
            "",
            "## 注意点",
            "",
            "- `store_products.generated.json` は正本ではなく、手編集禁止の生成物。",
            "- Store表示を変更する場合は、各商品の `ORIGINAL.md` を修正する。",
            "- 未確定値はJSON上では `null` に正規化した。",
            "- Stripe Payment Linkが未設定の商品には、Stripe購入ボタンを出さない。",
            "- Stripe Checkout API、Pages Functions、Webhook、R2、download_url確定は今回未実施。",
            "",
            "## 次Phase提案",
            "",
            "1. Stripe Payment Linkを付ける商品を絞る。",
            "2. Payment Linkを `ORIGINAL.md` へ戻す運用を決める。",
            "3. Store本番反映前にPayment Linkあり商品のみ `Stripeで購入` を目視確認する。",
            "4. Checkout API / Webhook / R2 は別Phaseで検討する。",
        ]
    )
    return "\n".join(lines) + "\n"


def main() -> int:
    GENERATED_DIR.mkdir(parents=True, exist_ok=True)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)

    originals, skipped = discover_originals()
    items: list[dict[str, Any]] = []
    for path in originals:
        try:
            item, skip, _meta = extract_item(path)
        except Exception as exc:  # noqa: BLE001 - report and continue for audit safety
            skipped.append({"source_original": path.as_posix(), "reason": f"parse_error: {exc}"})
            continue
        if item is not None:
            items.append(item)
        if skip is not None:
            skipped.append(skip)

    items.sort(key=lambda row: (row["type"], row["id"] or "", row["source_original"]))
    generated_at = datetime.now(timezone(timedelta(hours=9))).isoformat(timespec="seconds")
    data = {
        "generated_at": generated_at,
        "source_policy": SOURCE_POLICY,
        "schema_version": "1.0.0",
        "do_not_edit": True,
        "items": items,
        "skipped": skipped,
    }
    OUTPUT_JSON.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (GENERATED_DIR / "README.md").write_text(
        "# Generated Store Data\n\n"
        "`store_products.generated.json` is generated from `ORIGINAL.md` files.\n\n"
        "Do not edit generated JSON manually. Update the source `ORIGINAL.md` and run the generator again.\n",
        encoding="utf-8",
    )
    OUTPUT_REPORT.write_text(build_report(originals, items, skipped, generated_at), encoding="utf-8")
    print(f"generated items: {len(items)}")
    print(f"skipped: {len(skipped)}")
    print(f"output: {OUTPUT_JSON.relative_to(ROOT).as_posix()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
