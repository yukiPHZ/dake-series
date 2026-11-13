"""Check only the assets needed for human BOOTH registration.

``BOOTH_PREP = READY`` means DakeBOOTH Assist and a human have the product
text, thumbnail, ZIP, and support text needed to register the product. GitHub
Release state, BOOTH URL, release capture, video, SNS, Buffer, dakeapp.com,
Store, Cloudflare, and Stripe are intentionally outside this result.
"""

from __future__ import annotations

import argparse
import csv
import re
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from release_source_policy import read_app_source


ROOT = Path(__file__).resolve().parents[1]
APPS_DIR = ROOT / "01_apps"
REPORT_DIR = ROOT / "tools" / "reports"
BOOTH_STATUS = "available"
STATUS_DISPLAY_ORDER = ["available", "frozen", "draft", "experimental", "private", "internal", "unknown"]


@dataclass
class AppCheck:
    app: str
    status: str = "unknown"
    source_kind: str = ""
    source_path: str = ""
    source_ready: bool = False
    original_missing: bool = False
    booth_ready_dir: bool = False
    booth_product: bool = False
    booth_product_path: str = ""
    thumbnail: bool = False
    thumbnail_path: str = ""
    zip_file: bool = False
    zip_path: str = ""
    support_text: bool = False
    support_text_path: str = ""
    screenshot: bool = False
    screenshot_path: str = ""
    booth_url_state: str = "booth_product_missing"
    booth_url: str = ""
    actions: list[str] = field(default_factory=list)

    @property
    def distribution_candidate(self) -> bool:
        return self.status == BOOTH_STATUS

    @property
    def booth_prep_ready(self) -> bool:
        return self.distribution_candidate and all(
            [
                self.source_ready,
                self.booth_ready_dir,
                self.booth_product,
                self.thumbnail,
                self.zip_file,
                self.support_text,
            ]
        )


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


def first_existing(paths: list[Path]) -> Path | None:
    return next((path for path in paths if path.exists()), None)


def relative(path: Path | None, base: Path) -> str:
    return path.relative_to(base).as_posix() if path else ""


def booth_url_from_product(path: Path | None) -> tuple[str, str]:
    if path is None:
        return "booth_product_missing", ""
    match = re.search(r"(?ms)^# URL\s*\n(.*?)(?=\n# |\Z)", read_text(path))
    if not match:
        return "url_section_missing", ""
    url = match.group(1).strip()
    return ("set", url) if url else ("empty", "")


def valid_zip(path: Path | None) -> bool:
    return bool(path and path.is_file() and path.stat().st_size > 0 and zipfile.is_zipfile(path))


def check_app(app_dir: Path) -> AppCheck:
    result = AppCheck(app=app_dir.name)
    source = read_app_source(app_dir, ROOT)
    result.status = source.status
    result.source_kind = source.source_kind
    result.source_path = source.source_label
    result.source_ready = bool(source.meta)
    result.original_missing = source.original_missing

    ready_dir = app_dir / "booth_ready"
    result.booth_ready_dir = ready_dir.is_dir()

    product_path = first_existing([ready_dir / "booth_product.txt", app_dir / "booth_product.txt"])
    result.booth_product = product_path is not None
    result.booth_product_path = relative(product_path, app_dir)
    result.booth_url_state, result.booth_url = booth_url_from_product(product_path)

    thumbnail_path = first_existing(
        [ready_dir / "booth_thumbnail.jpg", app_dir / "assets" / "booth_thumbnail.jpg"]
    )
    result.thumbnail = thumbnail_path is not None
    result.thumbnail_path = relative(thumbnail_path, app_dir)

    zip_paths = sorted(ready_dir.glob("*.zip")) if ready_dir.is_dir() else []
    zip_path = next((path for path in zip_paths if valid_zip(path)), None)
    result.zip_file = zip_path is not None
    result.zip_path = relative(zip_path, app_dir)

    support_path = first_existing([ready_dir / "README.txt", ready_dir / "注意事項.txt"])
    result.support_text = support_path is not None
    result.support_text_path = relative(support_path, app_dir)

    screenshot_path = first_existing(
        [
            ready_dir / "screenshot.jpg",
            app_dir / "assets" / "screenshot.webp",
            app_dir / "assets" / "screenshot.jpg",
        ]
    )
    result.screenshot = screenshot_path is not None
    result.screenshot_path = relative(screenshot_path, app_dir)

    if not result.distribution_candidate:
        result.actions.append("excluded from normal BOOTH preparation by status")
        return result
    if not result.source_ready:
        result.actions.append("add/fix ORIGINAL.md metadata; Legacy README fallback is allowed only during migration")
    elif result.original_missing:
        result.actions.append("Legacy: ORIGINAL.md missing; migrate when this app is next updated")
    if not result.booth_ready_dir:
        result.actions.append("generate booth_ready/ with tools/make_booth_ready.py")
    if not result.booth_product:
        result.actions.append("generate booth_product.txt")
    elif not result.booth_product_path.startswith("booth_ready/"):
        result.actions.append("Legacy: booth_product.txt is outside booth_ready/")
    if not result.thumbnail:
        result.actions.append("generate booth_thumbnail.jpg")
    elif not result.thumbnail_path.startswith("booth_ready/"):
        result.actions.append("Legacy: booth_thumbnail.jpg is outside booth_ready/")
    if not result.zip_file:
        result.actions.append("prepare a valid distribution ZIP under booth_ready/")
    if not result.support_text:
        result.actions.append("add booth_ready/README.txt or booth_ready/注意事項.txt")
    if not result.screenshot:
        result.actions.append("optional: add a screenshot when human registration needs one")
    return result


def status_counts(results: list[AppCheck]) -> Counter[str]:
    return Counter(result.status for result in results)


def candidates(results: list[AppCheck]) -> list[AppCheck]:
    return [result for result in results if result.distribution_candidate]


def yn(value: bool) -> str:
    return "OK" if value else "NG"


def write_markdown(results: list[AppCheck], path: Path) -> None:
    selected = candidates(results)
    ready = [result for result in selected if result.booth_prep_ready]
    pending = [result for result in selected if not result.booth_prep_ready]
    counts = status_counts(results)
    url_counts = Counter(result.booth_url_state for result in selected)

    lines = [
        "# DAKE BOOTH Preparation Check",
        "",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "Codex checks registration materials only. BOOTH registration and publication are human operations.",
        "",
        "## Summary",
        "",
        f"- checked: {len(results)}",
        f"- distribution candidates: {len(selected)}",
        f"- BOOTH_PREP READY: {len(ready)}",
        f"- BOOTH_PREP PENDING: {len(pending)}",
        f"- excluded by status: {len(results) - len(selected)}",
        f"- original_missing: {sum(result.original_missing for result in selected)}",
        "",
        "## Status Counts",
        "",
    ]
    for status in STATUS_DISPLAY_ORDER:
        lines.append(f"- {status}: {counts.get(status, 0)}")
    for status, count in sorted(counts.items()):
        if status not in STATUS_DISPLAY_ORDER:
            lines.append(f"- {status}: {count}")

    lines.extend(
        [
            "",
            "## Required Materials",
            "",
            f"- booth_product.txt missing: {sum(not item.booth_product for item in selected)}",
            f"- booth_thumbnail.jpg missing: {sum(not item.thumbnail for item in selected)}",
            f"- valid ZIP missing: {sum(not item.zip_file for item in selected)}",
            f"- README/注意事項 missing: {sum(not item.support_text for item in selected)}",
            "",
            "## BOOTH URL (informational only)",
            "",
            f"- set: {url_counts.get('set', 0)}",
            f"- empty: {url_counts.get('empty', 0)}",
            f"- section missing: {url_counts.get('url_section_missing', 0)}",
            f"- product missing: {url_counts.get('booth_product_missing', 0)}",
            "",
            "BOOTH URL is not a BOOTH_PREP requirement. Record it only after human registration.",
            "",
            "## Applications",
            "",
            "| app | BOOTH_PREP | source | product | thumbnail | zip | support | screenshot | booth_url | next_action |",
            "| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |",
        ]
    )
    for result in results:
        state = "READY" if result.booth_prep_ready else ("PENDING" if result.distribution_candidate else "EXCLUDED")
        action = "<br>".join(result.actions) or "-"
        lines.append(
            f"| {result.app} | {state} | {result.source_kind or '-'} | {yn(result.booth_product)} | "
            f"{yn(result.thumbnail)} | {yn(result.zip_file)} | {yn(result.support_text)} | "
            f"{yn(result.screenshot)} | {result.booth_url_state} | {action} |"
        )

    lines.extend(
        [
            "",
            "## Legacy compatibility",
            "",
            "Root-level booth_product.txt and assets/booth_thumbnail.jpg remain readable during migration.",
            "Legacy closed/v2 closed, Release capture, SNS, Buffer, video, Store, Cloudflare, and Stripe are not normal BOOTH preparation results.",
        ]
    )
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_csv(results: list[AppCheck], path: Path) -> None:
    fields = [
        "app",
        "status",
        "distribution_candidate",
        "booth_prep_ready",
        "source_kind",
        "source_path",
        "original_missing",
        "booth_ready_dir",
        "booth_product",
        "booth_product_path",
        "thumbnail",
        "thumbnail_path",
        "zip_file",
        "zip_path",
        "support_text",
        "support_text_path",
        "screenshot",
        "screenshot_path",
        "booth_url_state",
        "booth_url",
        "actions",
    ]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for result in results:
            writer.writerow(
                {
                    **{field: getattr(result, field) for field in fields if hasattr(result, field)},
                    "distribution_candidate": result.distribution_candidate,
                    "booth_prep_ready": result.booth_prep_ready,
                    "actions": "; ".join(result.actions),
                }
            )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Check only BOOTH registration materials; does not open, register, or publish BOOTH."
    )
    parser.add_argument("--report-dir", type=Path, default=REPORT_DIR)
    parser.add_argument("--app", help="Check one app folder such as DAKE_Mail_Draft.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    app_dirs = sorted(path for path in APPS_DIR.iterdir() if path.is_dir() and path.name.startswith("DAKE_"))
    if args.app:
        app_dirs = [path for path in app_dirs if path.name == args.app]
        if not app_dirs:
            print(f"app not found: {args.app}")
            return 1
    results = [check_app(app_dir) for app_dir in app_dirs]

    report_dir = args.report_dir
    report_dir.mkdir(parents=True, exist_ok=True)
    md_path = report_dir / "booth_ready_check.md"
    csv_path = report_dir / "booth_ready_check.csv"
    write_markdown(results, md_path)
    write_csv(results, csv_path)

    selected = candidates(results)
    ready = [result for result in selected if result.booth_prep_ready]
    pending = [result for result in selected if not result.booth_prep_ready]
    print("DAKE BOOTH Preparation Check")
    print("boundary: assets only; BOOTH registration and publication are human operations")
    print(f"checked: {len(results)}")
    print(f"distribution candidates: {len(selected)}")
    print(f"BOOTH_PREP READY: {len(ready)}")
    print(f"BOOTH_PREP PENDING: {len(pending)}")
    for result in pending:
        print(f"- {result.app}: {', '.join(result.actions)}")
    print(f"report: {md_path}")
    print(f"csv: {csv_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
