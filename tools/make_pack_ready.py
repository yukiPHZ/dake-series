"""Generate BOOTH-ready DAKE pack products.

Packs live under 04_packs and are managed as pseudo products. A pack zip
bundles the source apps' existing booth_ready/*.zip files plus pack-level
README and notice files. It must not include source code, build folders,
dist folders, specs, pycache, or direct exe files.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
APPS_DIR = ROOT / "01_apps"
PACKS_DIR = ROOT / "04_packs"
READY_DIR_NAME = "pack_ready"
PRODUCT_FILE_NAME = "booth_product.txt"
MANIFEST_NAME = "pack_manifest.json"
THUMBNAIL_SIZE = (1200, 1200)
JST = timezone(timedelta(hours=9))
WINDOWS_FONT_DIR = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
FONT_CANDIDATES = {
    "bold": ["YuGothB.ttc", "YuGothM.ttc", "meiryob.ttc", "meiryo.ttc", "msgothic.ttc"],
    "regular": ["YuGothM.ttc", "YuGothR.ttc", "meiryo.ttc", "msgothic.ttc"],
}
NOTICE_TEXT = """【注意事項】

・このZIPはDAKEシリーズのパック商品です
・収録アプリは、それぞれのZIPを解凍してから起動してください
・Windows向けアプリです
・大切なファイルは事前にバックアップしてください
・本パックおよび収録アプリの無断転載・再配布を禁止します
・環境によっては起動時にWindowsの警告が表示される場合があります

PEAKHEADZ
https://peakheadz.com
"""


@dataclass
class SourceApp:
    folder: str
    display_name: str = ""
    status: str = "unknown"
    release_url: str = ""
    booth_url: str = ""
    zip_path: Path | None = None
    zip_sha256: str = ""
    zip_size: int = 0
    missing: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.missing


@dataclass
class PackResult:
    folder: str
    display_name: str
    status: str
    skipped: bool = False
    ok: bool = True
    missing: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    generated: list[str] = field(default_factory=list)
    zip_name: str = ""
    zip_size: int = 0
    booth_url: str = ""


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


def write_text(path: Path, text: str, dry_run: bool, generated: list[str]) -> None:
    generated.append(str(path.relative_to(ROOT)))
    if dry_run:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text.replace("\r\n", "\n").rstrip() + "\n", encoding="utf-8", newline="\n")


def extract_json_section(text: str, heading: str) -> dict[str, Any] | None:
    match = re.search(rf"(?s)^##\s+{re.escape(heading)}\s*```json\s*(.*?)\s*```", text, re.MULTILINE)
    if not match:
        return None
    loaded = json.loads(match.group(1))
    return loaded if isinstance(loaded, dict) else None


def extract_pack_body(text: str) -> str:
    match = re.search(r"(?s)^##\s+PACK_BODY\s*(.*?)(?:\n##\s+|\Z)", text, re.MULTILINE)
    return match.group(1).strip() if match else ""


def read_pack_meta(pack_dir: Path) -> tuple[dict[str, Any], str]:
    readme_path = pack_dir / "README.md"
    if not readme_path.exists():
        return {}, ""
    text = read_text(readme_path)
    return extract_json_section(text, "PACK_META") or {}, extract_pack_body(text)


def read_dake_meta(app_dir: Path) -> dict[str, Any]:
    readme_path = app_dir / "README.md"
    if not readme_path.exists():
        return {}
    text = read_text(readme_path)
    return extract_json_section(text, "DAKE_META") or {}


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


def preserved_pack_booth_url(pack_dir: Path, meta: dict[str, Any]) -> str:
    meta_url = str(meta.get("booth_url") or "").strip()
    if meta_url:
        return meta_url
    product_path = pack_dir / READY_DIR_NAME / PRODUCT_FILE_NAME
    if product_path.exists():
        return extract_booth_url_from_text(read_text(product_path))
    return ""


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def choose_source_zip(app_dir: Path) -> tuple[Path | None, list[str]]:
    ready_dir = app_dir / "booth_ready"
    if not ready_dir.exists():
        return None, []
    zip_files = sorted((path for path in ready_dir.iterdir() if path.is_file() and path.suffix.lower() == ".zip"), key=lambda item: item.name.lower())
    warnings: list[str] = []
    if len(zip_files) > 1:
        warnings.append("source has multiple booth_ready zip files; first one was selected")
    return (zip_files[0] if zip_files else None), warnings


def source_app(folder: str) -> SourceApp:
    app_dir = APPS_DIR / folder
    result = SourceApp(folder=folder)
    if not app_dir.exists():
        result.missing.append("app_folder")
        return result
    meta = read_dake_meta(app_dir)
    result.display_name = str(meta.get("display_name") or meta.get("site_title") or meta.get("launcher_title") or folder)
    result.status = str(meta.get("status") or "unknown").strip().lower() or "unknown"
    result.release_url = str(meta.get("release_url") or "").strip()
    result.booth_url = source_booth_url(app_dir)
    if result.status != "available":
        result.missing.append("status_available")
    if not result.release_url:
        result.missing.append("release_url")
    if not result.booth_url:
        result.missing.append("booth_url")
    zip_path, warnings = choose_source_zip(app_dir)
    result.warnings.extend(warnings)
    if zip_path is None:
        result.missing.append("source_zip")
    else:
        result.zip_path = zip_path
        result.zip_size = zip_path.stat().st_size
        result.zip_sha256 = sha256(zip_path)
    return result


def as_int(value: object, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def normalize_items(meta: dict[str, Any], pack_dir: Path) -> list[str]:
    raw_items = meta.get("included_apps")
    if isinstance(raw_items, list):
        items = [str(item).strip() for item in raw_items if str(item).strip()]
    else:
        items = []
    pack_items_path = pack_dir / "pack_items.txt"
    if pack_items_path.exists():
        file_items = [
            line.strip()
            for line in read_text(pack_items_path).splitlines()
            if line.strip() and not line.strip().startswith("#")
        ]
        if file_items and file_items != items:
            items = file_items
    return items


def tag_text(meta: dict[str, Any]) -> str:
    tags = meta.get("tags")
    if isinstance(tags, list):
        return "\n".join(str(tag).strip() for tag in tags if str(tag).strip())
    return str(tags or "").strip()


def wrap_lines(text: str, font: ImageFont.ImageFont, max_width: int, draw: ImageDraw.ImageDraw) -> list[str]:
    lines: list[str] = []
    for paragraph in text.splitlines() or [text]:
        current = ""
        for char in paragraph:
            trial = current + char
            if draw.textbbox((0, 0), trial, font=font)[2] <= max_width or not current:
                current = trial
            else:
                lines.append(current)
                current = char
        if current:
            lines.append(current)
    return lines


def font(kind: str, size: int) -> ImageFont.ImageFont:
    for name in FONT_CANDIDATES[kind]:
        path = WINDOWS_FONT_DIR / name
        if path.exists():
            return ImageFont.truetype(str(path), size)
    return ImageFont.load_default()


def make_thumbnail(pack_dir: Path, meta: dict[str, Any], sources: list[SourceApp], dry_run: bool, generated: list[str]) -> None:
    path = pack_dir / "assets" / "booth_thumbnail.jpg"
    generated.append(str(path.relative_to(ROOT)))
    if dry_run:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    image = Image.new("RGB", THUMBNAIL_SIZE, "#F7F9FC")
    draw = ImageDraw.Draw(image)
    width, height = THUMBNAIL_SIZE
    draw.rectangle((0, 0, width, 16), fill="#2457C5")
    draw.rectangle((78, 96, width - 78, height - 96), fill="#FFFFFF", outline="#D7DCE7", width=3)
    draw.text((116, 128), "PEAKHEADZ / DAKE PACK", fill="#667085", font=font("bold", 34))
    title = str(meta.get("display_name") or pack_dir.name)
    title_font = font("bold", 78)
    for line in wrap_lines(title, title_font, 960, draw)[:2]:
        draw.text((116, 210), line, fill="#182230", font=title_font)
        break
    summary = str(meta.get("summary") or "")
    summary_font = font("regular", 34)
    y = 410
    for line in wrap_lines(summary, summary_font, 900, draw)[:3]:
        draw.text((116, y), line, fill="#344054", font=summary_font)
        y += 48
    draw.rectangle((116, 620, width - 116, 626), fill="#E4E7EC")
    item_font = font("regular", 31)
    y = 675
    for source in sources[:5]:
        label = source.display_name or source.folder
        draw.text((132, y), f"+ {label}", fill="#1D4CB4", font=item_font)
        y += 48
    price = as_int(meta.get("price"))
    draw.text((116, height - 185), f"{price:,} yen", fill="#182230", font=font("bold", 50))
    draw.text((116, height - 116), "Windows向け / 実務用 / 軽量", fill="#667085", font=font("regular", 28))
    image.save(path, "JPEG", quality=92, optimize=True)


def make_readme_txt(meta: dict[str, Any], sources: list[SourceApp]) -> str:
    lines = "\n".join(f"・{source.display_name or source.folder} ({source.folder})" for source in sources)
    return f"""{meta.get('display_name')}

{meta.get('summary')}

収録アプリ:
{lines}

使い方:
1. このパックZIPを解凍します。
2. appsフォルダ内の各アプリZIPを解凍します。
3. 使いたいアプリのexeを起動します。

PEAKHEADZ
https://peakheadz.com

Vibe-Coded by Yukihiko Kikuta
"""


def make_product_text(meta: dict[str, Any], body: str, sources: list[SourceApp], booth_url: str) -> str:
    display_name = str(meta.get("display_name") or meta.get("folder_name") or "DAKE Pack")
    price = as_int(meta.get("price"))
    summary = str(meta.get("summary") or "").strip()
    source_lines = "\n".join(f"・{source.display_name or source.folder} / {source.booth_url or 'BOOTH URL未設定'}" for source in sources)
    distribution_lines = "\n".join(f"・apps/{source.folder}/{source.zip_path.name if source.zip_path else source.folder + '.zip'}" for source in sources)
    return f"""# 商品名
{display_name}

# 価格案
{price}円

# 商品紹介文
{summary}

{body}

# 収録アプリ
{source_lines}

# 使い方
1. パックZIPを解凍します。
2. appsフォルダ内の各アプリZIPを解凍します。
3. 使いたいアプリのexeを起動します。

# 注意事項
Windows向けです。大切なファイルは事前にバックアップしてください。収録アプリの詳しい使い方は各アプリのREADMEを確認してください。

# 対象
日常業務で小さな作業をまとめて処理したい方。

# 配布物
{distribution_lines}
・README.txt
・注意事項.txt

# 免責
本ソフトウェアの利用によって発生した損害について、作者は責任を負いません。

# コピーライト
PEAKHEADZ / Vibe-Coded by Yukihiko Kikuta

# タグ
{tag_text(meta)}

# URL
{booth_url}
"""


def build_manifest(pack_dir: Path, meta: dict[str, Any], sources: list[SourceApp], booth_url: str, zip_path: Path | None) -> dict[str, Any]:
    source_items: list[dict[str, Any]] = []
    for source in sources:
        source_items.append(
            {
                "folder": source.folder,
                "display_name": source.display_name,
                "status": source.status,
                "release_url": source.release_url,
                "booth_url": source.booth_url,
                "source_zip": source.zip_path.relative_to(ROOT).as_posix() if source.zip_path else "",
                "source_zip_name": source.zip_path.name if source.zip_path else "",
                "source_zip_size": source.zip_size,
                "source_zip_sha256": source.zip_sha256,
            }
        )
    manifest: dict[str, Any] = {
        "schema": "dake_pack_manifest_v1",
        "generated_at": datetime.now(JST).isoformat(timespec="seconds"),
        "folder_name": pack_dir.name,
        "display_name": meta.get("display_name", pack_dir.name),
        "status": meta.get("status", "unknown"),
        "price": as_int(meta.get("price")),
        "booth_url": booth_url,
        "included_apps": source_items,
        "pack_zip": zip_path.relative_to(ROOT).as_posix() if zip_path else "",
        "pack_zip_size": zip_path.stat().st_size if zip_path and zip_path.exists() else 0,
        "pack_zip_sha256": sha256(zip_path) if zip_path and zip_path.exists() else "",
    }
    return manifest


def make_pack_zip(pack_dir: Path, sources: list[SourceApp], dry_run: bool, generated: list[str]) -> Path:
    ready_dir = pack_dir / READY_DIR_NAME
    zip_path = ready_dir / f"{pack_dir.name}.zip"
    generated.append(str(zip_path.relative_to(ROOT)))
    if dry_run:
        return zip_path
    ready_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.write(ready_dir / "README.txt", "README.txt")
        archive.write(ready_dir / "注意事項.txt", "注意事項.txt")
        for source in sources:
            if source.zip_path is None:
                continue
            archive.write(source.zip_path, f"apps/{source.folder}/{source.zip_path.name}")
    return zip_path


def process_pack(pack_dir: Path, only_available: bool, dry_run: bool) -> PackResult:
    meta, body = read_pack_meta(pack_dir)
    display_name = str(meta.get("display_name") or pack_dir.name)
    status = str(meta.get("status") or "unknown").strip().lower() or "unknown"
    result = PackResult(folder=pack_dir.name, display_name=display_name, status=status)
    if only_available and status != "available":
        result.skipped = True
        return result
    if not meta:
        result.ok = False
        result.missing.append("PACK_META")
        return result
    items = normalize_items(meta, pack_dir)
    if not items:
        result.ok = False
        result.missing.append("included_apps")
        return result
    sources = [source_app(folder) for folder in items]
    for source in sources:
        result.warnings.extend(f"{source.folder}: {warning}" for warning in source.warnings)
        result.missing.extend(f"{source.folder}:{missing}" for missing in source.missing)
    result.ok = not result.missing
    booth_url = preserved_pack_booth_url(pack_dir, meta)
    result.booth_url = booth_url
    ready_dir = pack_dir / READY_DIR_NAME
    write_text(ready_dir / "README.txt", make_readme_txt(meta, sources), dry_run, result.generated)
    write_text(ready_dir / "注意事項.txt", NOTICE_TEXT, dry_run, result.generated)
    write_text(ready_dir / PRODUCT_FILE_NAME, make_product_text(meta, body, sources, booth_url), dry_run, result.generated)
    make_thumbnail(pack_dir, meta, sources, dry_run, result.generated)
    zip_path = make_pack_zip(pack_dir, sources, dry_run, result.generated)
    if not dry_run and zip_path.exists():
        result.zip_name = zip_path.name
        result.zip_size = zip_path.stat().st_size
    manifest = build_manifest(pack_dir, meta, sources, booth_url, zip_path if not dry_run else None)
    write_text(pack_dir / MANIFEST_NAME, json.dumps(manifest, ensure_ascii=False, indent=2), dry_run, result.generated)
    return result


def discover_packs(selected: str | None) -> list[Path]:
    if selected:
        return [PACKS_DIR / selected]
    if not PACKS_DIR.exists():
        return []
    return sorted((path for path in PACKS_DIR.iterdir() if path.is_dir() and not path.name.startswith(".")), key=lambda item: item.name.lower())


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pack", help="process one pack folder")
    parser.add_argument("--only-available", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    results = [process_pack(pack, args.only_available, args.dry_run) for pack in discover_packs(args.pack)]
    for result in results:
        state = "SKIP" if result.skipped else ("OK" if result.ok else "CHECK")
        print(f"{state}: {result.folder} ({result.display_name})")
        if result.zip_name:
            print(f"  zip: {result.zip_name} ({result.zip_size:,} bytes)")
        if result.missing:
            print(f"  missing: {', '.join(result.missing)}")
        if result.warnings:
            print(f"  warnings: {', '.join(result.warnings)}")
        if args.dry_run:
            print(f"  dry-run targets: {len(result.generated)}")
    failed = [result for result in results if not result.skipped and not result.ok]
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
