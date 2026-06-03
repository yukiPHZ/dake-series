"""Check DAKE pack product readiness and stale state."""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
APPS_DIR = ROOT / "01_apps"
PACKS_DIR = ROOT / "04_packs"
DEFAULT_REPORT_DIR = ROOT / "tools" / "reports"
PRODUCT_FILE_NAME = "booth_product.txt"
READY_DIR_NAME = "pack_ready"
MANIFEST_NAME = "pack_manifest.json"


@dataclass
class PackCheck:
    folder: str
    display_name: str = ""
    status: str = "unknown"
    price: int = 0
    included_count: int = 0
    booth_url: str = ""
    has_thumbnail: bool = False
    has_ready_thumbnail: bool = False
    thumbnail_size: str = ""
    ready_thumbnail_size: str = ""
    thumbnail_ok: bool = False
    ready_thumbnail_ok: bool = False
    product_references_thumbnail: bool = False
    has_product: bool = False
    has_ready_dir: bool = False
    has_readme_txt: bool = False
    has_notice: bool = False
    has_pack_zip: bool = False
    pack_zip: str = ""
    has_manifest: bool = False
    source_zip_count: int = 0
    source_zip_stale_count: int = 0
    missing: list[str] = field(default_factory=list)
    stale: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def materials_score(self) -> int:
        return sum((self.thumbnail_ok, self.has_product and self.product_references_thumbnail, self.has_ready_dir and self.ready_thumbnail_ok))

    @property
    def ready(self) -> bool:
        return (
            self.status == "available"
            and self.materials_score == 3
            and self.thumbnail_ok
            and self.ready_thumbnail_ok
            and self.product_references_thumbnail
            and self.has_readme_txt
            and self.has_notice
            and self.has_pack_zip
            and self.has_manifest
            and not self.missing
            and not self.stale
        )

    @property
    def next_step(self) -> str:
        if self.missing:
            return "make_pack_ready"
        if self.stale:
            return "regenerate_pack"
        if not self.booth_url:
            return "booth_register"
        return "ready"


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


def extract_json_section(text: str, heading: str) -> dict[str, Any] | None:
    match = re.search(rf"(?s)^##\s+{re.escape(heading)}\s*```json\s*(.*?)\s*```", text, re.MULTILINE)
    if not match:
        return None
    loaded = json.loads(match.group(1))
    return loaded if isinstance(loaded, dict) else None


def extract_meta(readme_path: Path, heading: str) -> dict[str, Any]:
    if not readme_path.exists():
        return {}
    try:
        return extract_json_section(read_text(readme_path), heading) or {}
    except (OSError, json.JSONDecodeError):
        return {}


def extract_booth_url_from_text(text: str) -> str:
    lines = text.splitlines()
    for index, line in enumerate(lines):
        stripped = line.strip()
        lowered = stripped.lower()
        if lowered.startswith("booth url:"):
            return stripped.split(":", 1)[1].strip()
        if lowered.startswith("booth_url=") or lowered.startswith("url="):
            return stripped.split("=", 1)[1].strip()
        if lowered in {"# url", "# booth url"}:
            for value_line in lines[index + 1 :]:
                value = value_line.strip()
                if not value:
                    continue
                if value.startswith("#"):
                    break
                return value
    return ""


def source_booth_url(app_dir: Path) -> str:
    for candidate in (app_dir / "booth_ready" / PRODUCT_FILE_NAME, app_dir / PRODUCT_FILE_NAME):
        if candidate.exists():
            url = extract_booth_url_from_text(read_text(candidate))
            if url:
                return url
    return ""


def pack_booth_url(pack_dir: Path, meta: dict[str, Any]) -> str:
    meta_url = str(meta.get("booth_url") or "").strip()
    if meta_url:
        return meta_url
    product = pack_dir / READY_DIR_NAME / PRODUCT_FILE_NAME
    if product.exists():
        return extract_booth_url_from_text(read_text(product))
    return ""


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def choose_source_zip(app_dir: Path) -> Path | None:
    ready_dir = app_dir / "booth_ready"
    if not ready_dir.exists():
        return None
    zips = sorted((path for path in ready_dir.iterdir() if path.is_file() and path.suffix.lower() == ".zip"), key=lambda item: item.name.lower())
    return zips[0] if zips else None


def load_manifest(pack_dir: Path) -> dict[str, Any]:
    path = pack_dir / MANIFEST_NAME
    if not path.exists():
        return {}
    try:
        loaded = json.loads(read_text(path))
    except json.JSONDecodeError:
        return {}
    return loaded if isinstance(loaded, dict) else {}


def image_size_text(path: Path) -> tuple[str, bool]:
    if not path.exists():
        return "", False
    try:
        with Image.open(path) as image:
            width, height = image.size
    except Exception:
        return "unreadable", False
    return f"{width}x{height}", width == 1200 and height == 1200


def check_pack(pack_dir: Path) -> PackCheck:
    meta = extract_meta(pack_dir / "README.md", "PACK_META")
    result = PackCheck(
        folder=pack_dir.name,
        display_name=str(meta.get("display_name") or pack_dir.name),
        status=str(meta.get("status") or "unknown").strip().lower() or "unknown",
        price=int(meta.get("price") or 0),
        booth_url=pack_booth_url(pack_dir, meta),
    )
    included = meta.get("included_apps")
    if not isinstance(included, list):
        included = []
    result.included_count = len(included)
    ready_dir = pack_dir / READY_DIR_NAME
    result.has_ready_dir = ready_dir.is_dir()
    thumbnail_path = pack_dir / "assets" / "booth_thumbnail.jpg"
    ready_thumbnail_path = ready_dir / "booth_thumbnail.jpg"
    result.has_thumbnail = thumbnail_path.exists()
    result.has_ready_thumbnail = ready_thumbnail_path.exists()
    result.thumbnail_size, result.thumbnail_ok = image_size_text(thumbnail_path)
    result.ready_thumbnail_size, result.ready_thumbnail_ok = image_size_text(ready_thumbnail_path)
    result.has_product = (ready_dir / PRODUCT_FILE_NAME).exists()
    if result.has_product:
        result.product_references_thumbnail = "assets/booth_thumbnail.jpg" in read_text(ready_dir / PRODUCT_FILE_NAME)
    result.has_readme_txt = (ready_dir / "README.txt").exists()
    result.has_notice = (ready_dir / "注意事項.txt").exists()
    pack_zips = sorted((path for path in ready_dir.glob("*.zip") if path.is_file()), key=lambda item: item.name.lower()) if ready_dir.exists() else []
    result.has_pack_zip = bool(pack_zips)
    result.pack_zip = pack_zips[0].name if pack_zips else ""
    result.has_manifest = (pack_dir / MANIFEST_NAME).exists()
    manifest = load_manifest(pack_dir)

    for key, present in (
        ("PACK_META", bool(meta)),
        ("included_apps", bool(included)),
        ("assets/booth_thumbnail.jpg", result.has_thumbnail),
        ("assets/booth_thumbnail.jpg:1200x1200", result.thumbnail_ok),
        ("pack_ready/booth_thumbnail.jpg", result.has_ready_thumbnail),
        ("pack_ready/booth_thumbnail.jpg:1200x1200", result.ready_thumbnail_ok),
        ("pack_ready/booth_product.txt", result.has_product),
        ("booth_product:assets/booth_thumbnail.jpg", result.product_references_thumbnail),
        ("pack_ready/README.txt", result.has_readme_txt),
        ("pack_ready/注意事項.txt", result.has_notice),
        ("pack_ready/*.zip", result.has_pack_zip),
        ("pack_manifest.json", result.has_manifest),
    ):
        if not present:
            result.missing.append(key)

    manifest_sources = {
        str(item.get("folder")): item
        for item in manifest.get("included_apps", [])
        if isinstance(item, dict)
    } if manifest else {}
    for folder in included:
        app_dir = APPS_DIR / str(folder)
        app_meta = extract_meta(app_dir / "README.md", "DAKE_META")
        if not app_dir.exists():
            result.missing.append(f"{folder}:app_folder")
            continue
        if str(app_meta.get("status") or "").strip().lower() != "available":
            result.missing.append(f"{folder}:status_available")
        if not str(app_meta.get("release_url") or "").strip():
            result.missing.append(f"{folder}:release_url")
        booth_url = source_booth_url(app_dir)
        if not booth_url:
            result.missing.append(f"{folder}:booth_url")
        zip_path = choose_source_zip(app_dir)
        if zip_path is None:
            result.missing.append(f"{folder}:source_zip")
            continue
        result.source_zip_count += 1
        current_sha = sha256(zip_path)
        manifest_item = manifest_sources.get(str(folder), {})
        if manifest_item:
            if current_sha != str(manifest_item.get("source_zip_sha256") or ""):
                result.stale.append(f"{folder}:source_zip_sha256")
                result.source_zip_stale_count += 1
            if str(app_meta.get("release_url") or "").strip() != str(manifest_item.get("release_url") or "").strip():
                result.stale.append(f"{folder}:release_url")
            if booth_url != str(manifest_item.get("booth_url") or "").strip():
                result.stale.append(f"{folder}:booth_url")
        elif manifest:
            result.stale.append(f"{folder}:manifest_missing_source")

    if manifest:
        manifest_pack_zip = str(manifest.get("pack_zip") or "")
        if manifest_pack_zip:
            manifest_pack_path = ROOT / manifest_pack_zip
            if manifest_pack_path.exists() and str(manifest.get("pack_zip_sha256") or "") != sha256(manifest_pack_path):
                result.stale.append("pack_zip_sha256")
        if str(manifest.get("booth_url") or "").strip() != result.booth_url:
            result.stale.append("pack_booth_url")
    return result


def discover_packs(only_available: bool) -> list[Path]:
    if not PACKS_DIR.exists():
        return []
    packs: list[Path] = []
    for path in sorted((child for child in PACKS_DIR.iterdir() if child.is_dir()), key=lambda item: item.name.lower()):
        meta = extract_meta(path / "README.md", "PACK_META")
        if only_available and str(meta.get("status") or "").strip().lower() != "available":
            continue
        packs.append(path)
    return packs


def write_reports(results: list[PackCheck], report_dir: Path) -> None:
    report_dir.mkdir(parents=True, exist_ok=True)
    csv_path = report_dir / "pack_ready_check.csv"
    md_path = report_dir / "pack_ready_check.md"
    with csv_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "folder",
                "display_name",
                "status",
                "price",
                "included_count",
                "booth_url",
                "thumbnail",
                "ready_thumbnail",
                "thumbnail_size",
                "ready_thumbnail_size",
                "thumbnail_ok",
                "ready_thumbnail_ok",
                "product_references_thumbnail",
                "booth_product",
                "pack_ready",
                "readme_txt",
                "notice",
                "pack_zip",
                "manifest",
                "source_zip_count",
                "source_zip_stale_count",
                "booth_score",
                "missing",
                "stale",
                "next_step",
            ],
        )
        writer.writeheader()
        for result in results:
            writer.writerow(
                {
                    "folder": result.folder,
                    "display_name": result.display_name,
                    "status": result.status,
                    "price": result.price,
                    "included_count": result.included_count,
                    "booth_url": result.booth_url,
                    "thumbnail": result.has_thumbnail,
                    "ready_thumbnail": result.has_ready_thumbnail,
                    "thumbnail_size": result.thumbnail_size,
                    "ready_thumbnail_size": result.ready_thumbnail_size,
                    "thumbnail_ok": result.thumbnail_ok,
                    "ready_thumbnail_ok": result.ready_thumbnail_ok,
                    "product_references_thumbnail": result.product_references_thumbnail,
                    "booth_product": result.has_product,
                    "pack_ready": result.has_ready_dir,
                    "readme_txt": result.has_readme_txt,
                    "notice": result.has_notice,
                    "pack_zip": result.pack_zip,
                    "manifest": result.has_manifest,
                    "source_zip_count": result.source_zip_count,
                    "source_zip_stale_count": result.source_zip_stale_count,
                    "booth_score": f"{result.materials_score}/3",
                    "missing": "; ".join(result.missing),
                    "stale": "; ".join(result.stale),
                    "next_step": result.next_step,
                }
            )

    lines = [
        "# DAKE Pack Ready Check",
        "",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        f"- Pack total: {len(results)}",
        f"- Ready: {sum(1 for result in results if result.ready)}",
        f"- BOOTH URL set: {sum(1 for result in results if result.booth_url)}",
        f"- BOOTH URL unset: {sum(1 for result in results if not result.booth_url)}",
        f"- Pack thumbnail ok: {sum(1 for result in results if result.thumbnail_ok and result.ready_thumbnail_ok)}",
        f"- Pack thumbnail missing: {sum(1 for result in results if not result.has_thumbnail or not result.has_ready_thumbnail)}",
        f"- Pack thumbnail wrong size: {sum(1 for result in results if (result.has_thumbnail and not result.thumbnail_ok) or (result.has_ready_thumbnail and not result.ready_thumbnail_ok))}",
        f"- Stale: {sum(1 for result in results if result.stale)}",
        "",
        "## Packs",
        "",
        "| Folder | Display | Status | BOOTH | Thumb | Score | Zip | Missing | Stale | Next |",
        "|---|---|---|---|---|---|---|---|---|---|",
    ]
    for result in results:
        lines.append(
            "| {folder} | {display} | {status} | {booth} | {thumb} | {score} | {zip} | {missing} | {stale} | {next} |".format(
                folder=result.folder,
                display=result.display_name,
                status=result.status,
                booth="set" if result.booth_url else "unset",
                thumb="ok" if result.thumbnail_ok and result.ready_thumbnail_ok else "check",
                score=f"{result.materials_score}/3",
                zip=result.pack_zip or "-",
                missing=", ".join(result.missing) or "-",
                stale=", ".join(result.stale) or "-",
                next=result.next_step,
            )
        )
    md_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8", newline="\n")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--report-dir", default=str(DEFAULT_REPORT_DIR))
    parser.add_argument("--only-available", action="store_true")
    args = parser.parse_args()
    results = [check_pack(path) for path in discover_packs(args.only_available)]
    write_reports(results, Path(args.report_dir))
    for result in results:
        state = "OK" if result.ready else "CHECK"
        print(f"{state}: {result.folder} booth={result.materials_score}/3 thumb={result.thumbnail_size or '-'}/{result.ready_thumbnail_size or '-'} url={'set' if result.booth_url else 'empty'} zip={result.pack_zip or '-'} next={result.next_step}")
        if result.missing:
            print(f"  missing: {', '.join(result.missing)}")
        if result.stale:
            print(f"  stale: {', '.join(result.stale)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
