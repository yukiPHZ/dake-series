"""Generate BOOTH-ready assets for the DAKE formal shipping line.

DAKE formal shipping includes README metadata, release_body.md,
assets/screenshot.webp, assets/booth_thumbnail.jpg, booth_product.txt,
booth_ready assets, GitHub Release, BOOTH, dakeapp.com, and Cloudflare
verification. This tool owns the BOOTH-ready asset generation part of that
line and must preserve manually entered BOOTH URLs in booth_product.txt.

booth_thumbnail.jpg / booth_product.txt / booth_ready/ are formal shipping
assets. They are not optional cleanup after GitHub Release.
The BOOTH registration source file used in practice is
booth_ready/booth_product.txt.

When --only-available is used, only DAKE_META.status == available is processed.
Frozen, draft, experimental, private, and internal apps are skipped as
non-shipping candidates.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont

from release_source_policy import metadata_line


ROOT = Path(__file__).resolve().parents[1]
APPS_DIR = ROOT / "01_apps"
NOTICE_TEXT = """【注意事項】

・Windows向けアプリです
・ご利用は自己責任でお願いいたします
・大切なファイルは事前にバックアップを推奨します
・本ソフトウェアの無断転載・再配布を禁止します
・環境によっては起動時にWindowsの警告が表示される場合があります

PEAKHEADZ
https://peakheadz.com
"""
THUMBNAIL_SIZE = (1200, 1200)
WINDOWS_FONT_DIR = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
FONT_CANDIDATES = {
    "bold": [
        "YuGothB.ttc",
        "YuGothM.ttc",
        "meiryob.ttc",
        "meiryo.ttc",
        "msgothic.ttc",
    ],
    "regular": [
        "YuGothM.ttc",
        "YuGothR.ttc",
        "meiryo.ttc",
        "msgothic.ttc",
    ],
}


@dataclass
class AppResult:
    folder: str
    ok: bool = True
    status: str = "unknown"
    skipped: bool = False
    missing: list[str] = field(default_factory=list)
    generated: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    zip_name: str = ""
    product_title: str = ""
    price: int = 300
    price_source: str = ""
    thumbnail_error: str = ""
    product_updated: bool = False
    booth_url: str = ""


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")


def extract_meta(readme_text: str) -> dict | None:
    match = re.search(r"(?s)## DAKE_META\s*```json\s*(\{.*?\})\s*```", readme_text)
    if not match:
        match = re.search(r"(?s)## DAKE_META.*?(\{.*?\})", readme_text)
    if not match:
        return None
    return json.loads(match.group(1))


def extract_release_body(readme_text: str) -> str:
    match = re.search(r"(?s)## RELEASE_BODY\s*(.*?)(?:\n## |\Z)", readme_text)
    if not match:
        return ""
    body = match.group(1).strip()
    body = re.sub(r"^```[a-zA-Z0-9_-]*\s*", "", body)
    body = re.sub(r"\s*```$", "", body).strip()
    return body


def normalize_features(body: str, fallback: str) -> list[str]:
    lines: list[str] = []
    for raw_line in body.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or line.startswith("```"):
            continue
        if line.startswith("- "):
            line = line[2:].strip()
        if line.startswith("・"):
            line = line[1:].strip()
        lines.append(line)
    if not lines and fallback:
        lines.append(fallback)
    cleaned: list[str] = []
    for line in lines:
        if line not in cleaned:
            cleaned.append(line)
    return cleaned[:5]


def product_title(meta: dict) -> str:
    return meta.get("site_title") or meta.get("display_name") or meta.get("folder_name", "DAKE App")


def thumbnail_title(meta: dict) -> str:
    return (
        meta.get("display_name")
        or meta.get("site_title")
        or meta.get("launcher_title")
        or meta.get("folder_name", "DAKE App")
    )


def short_description(meta: dict) -> str:
    return (
        meta.get("launcher_description")
        or meta.get("site_description")
        or meta.get("update_summary")
        or product_title(meta)
    )


def thumbnail_description(meta: dict) -> str:
    raw = (
        meta.get("site_description")
        or meta.get("launcher_description")
        or meta.get("update_summary")
        or ""
    )
    text = re.sub(r"\s+", " ", raw).strip()
    if len(text) > 76:
        sentence = re.split(r"(?<=。)", text, maxsplit=1)[0].strip()
        if 0 < len(sentence) <= 76:
            text = sentence
        else:
            text = text[:75].rstrip("、。 ") + "…"
    return text


def price_for(folder: str, title: str, features: list[str]) -> int:
    blob = f"{folder} {title} {' '.join(features)}"
    if "PDF" in blob:
        return 500
    if any(token in blob for token in ["Git", "Brainz", "BRAINZ", "OIKAWA", "Wake", "Backup", "バックアップ"]):
        return 500
    if "Memo" in blob or "メモ" in blob:
        return 300
    if "Image" in blob or "画像" in blob or "HEIC" in blob:
        return 500
    return 300


def canonical_price_for(app_dir: Path) -> int | None:
    original_path = app_dir / "ORIGINAL.md"
    if not original_path.exists():
        return None
    raw = metadata_line(read_text(original_path), "price")
    match = re.search(r"[0-9][0-9,]*", raw)
    if not match:
        return None
    return int(match.group(0).replace(",", ""))


def tags_for(folder: str, title: str, features: list[str]) -> list[str]:
    blob = f"{folder} {title} {' '.join(features)}"
    tags: list[str] = []
    if "PDF" in blob:
        tags.append("PDF")
    if "Memo" in blob or "メモ" in blob:
        tags.append("メモ")
    if "Image" in blob or "画像" in blob or "HEIC" in blob or "写真" in blob:
        tags.append("画像")
    if "Mail" in blob or "メール" in blob:
        tags.append("メール")
    if any(token in blob for token in ["Git", "Brainz", "BRAINZ", "OIKAWA", "Wake"]):
        tags.append("開発")
    for tag in ["Windows", "実務", "ツール", "仕事効率化", "軽量", "シンプル"]:
        tags.append(tag)
    unique: list[str] = []
    for tag in tags:
        if tag not in unique:
            unique.append(tag)
    return unique


def make_readme_txt(title: str, description: str, features: list[str]) -> str:
    feature_lines = "\n".join(f"・{feature}" for feature in features[:5])
    return f"""{title}

{description}

主な特徴：
{feature_lines}

PEAKHEADZ
https://peakheadz.com

Vibe-Coded by Yukihiko Kikuta
"""


def extract_booth_url(product_text: str) -> str:
    match = re.search(r"(?ms)^# URL\s*\n(.*?)(?=\n# |\Z)", product_text)
    if not match:
        return ""
    return match.group(1).strip()


def app_meta_for_filter(app_dir: Path) -> tuple[dict | None, str]:
    readme_path = app_dir / "README.md"
    if not readme_path.exists():
        return None, "missing README.md"
    try:
        meta = extract_meta(read_text(readme_path))
    except json.JSONDecodeError as exc:
        return None, f"invalid DAKE_META JSON: {exc}"
    if meta is None:
        return None, "missing DAKE_META"
    return meta, str(meta.get("status") or "unknown")


def make_product_txt(
    title: str,
    price: int,
    description: str,
    features: list[str],
    tags: list[str],
    zip_name: str,
    release_url: str = "",
    booth_url: str = "",
) -> str:
    feature_lines = "\n".join(f"・{feature}" for feature in features[:5])
    tag_lines = "\n".join(tags)
    intro = f"""{description}

{feature_lines}

実務の流れを、
少し静かにするための道具です。"""
    return f"""# 商品名
{title}

# 価格案
{price}円

# 商品紹介文
{intro}

# タグ
{tag_lines}

# BOOTH商品画像
assets/booth_thumbnail.jpg

# 補助画像
assets/screenshot.jpg

# 作品ファイル
booth_ready/{zip_name}

# GitHub Release
{release_url}

# 注意事項
Windows向けアプリです。
ご利用は自己責任でお願いいたします。
大切なファイルは事前にバックアップを推奨します。
本ソフトウェアの無断転載・再配布を禁止します。

# URL
{booth_url}
""".rstrip("\n") + "\n"


def convert_webp_to_jpg(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(src) as image:
        if image.mode in {"RGBA", "LA"}:
            background = Image.new("RGB", image.size, (255, 255, 255))
            alpha = image.getchannel("A") if image.mode == "RGBA" else image.getchannel(1)
            background.paste(image.convert("RGBA"), mask=alpha)
            image = background
        else:
            image = image.convert("RGB")
        image.save(dst, "JPEG", quality=95, optimize=True)


def load_font(size: int, *, bold: bool = False) -> ImageFont.ImageFont:
    names = FONT_CANDIDATES["bold" if bold else "regular"]
    for name in names:
        path = WINDOWS_FONT_DIR / name
        if path.exists():
            try:
                return ImageFont.truetype(str(path), size=size)
            except OSError:
                continue
    return ImageFont.load_default()


def text_width(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont) -> int:
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0]


def fit_font(draw: ImageDraw.ImageDraw, text: str, max_width: int, start_size: int, min_size: int, *, bold: bool) -> ImageFont.ImageFont:
    for size in range(start_size, min_size - 1, -2):
        font = load_font(size, bold=bold)
        if text_width(draw, text, font) <= max_width:
            return font
    return load_font(min_size, bold=bold)


def trim_to_width(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_width: int) -> str:
    if text_width(draw, text, font) <= max_width:
        return text
    ellipsis = "…"
    trimmed = text
    while trimmed and text_width(draw, trimmed + ellipsis, font) > max_width:
        trimmed = trimmed[:-1]
    return trimmed + ellipsis if trimmed else ellipsis


def wrap_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_width: int, max_lines: int) -> list[str]:
    if not text:
        return []
    lines: list[str] = []
    current = ""
    for char in text:
        candidate = current + char
        if current and text_width(draw, candidate, font) > max_width:
            lines.append(current)
            current = char
            if len(lines) == max_lines:
                break
        else:
            current = candidate
    if len(lines) < max_lines and current:
        lines.append(current)
    if len(lines) > max_lines:
        lines = lines[:max_lines]
    if len(lines) == max_lines:
        consumed = "".join(lines)
        if len(consumed) < len(text):
            lines[-1] = trim_to_width(draw, lines[-1] + "…", font, max_width)
    return lines


def make_background(size: tuple[int, int]) -> Image.Image:
    width, height = size
    image = Image.new("RGB", size, "#f7f8fa")
    draw = ImageDraw.Draw(image)
    for y in range(height):
        ratio = y / max(height - 1, 1)
        value = int(255 - ratio * 9)
        draw.line([(0, y), (width, y)], fill=(value, value, min(value + 1, 255)))
    return image


def create_booth_thumbnail(src: Path, dst: Path, title: str, description: str) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    canvas = make_background(THUMBNAIL_SIZE)
    draw = ImageDraw.Draw(canvas)
    width, height = THUMBNAIL_SIZE
    margin_x = 120
    max_text_width = width - margin_x * 2

    title_font = fit_font(draw, title, max_text_width, 66, 42, bold=True)
    desc_font = load_font(30)
    title = trim_to_width(draw, title, title_font, max_text_width)
    desc_lines = wrap_text(draw, description, desc_font, max_text_width, 2)

    title_bbox = draw.textbbox((0, 0), title, font=title_font)
    title_x = (width - (title_bbox[2] - title_bbox[0])) // 2
    draw.text((title_x, 112), title, fill="#1f2933", font=title_font)

    desc_y = 205
    for line in desc_lines:
        line_bbox = draw.textbbox((0, 0), line, font=desc_font)
        line_x = (width - (line_bbox[2] - line_bbox[0])) // 2
        draw.text((line_x, desc_y), line, fill="#4b5563", font=desc_font)
        desc_y += 42
    draw.rounded_rectangle((width // 2 - 54, 305, width // 2 + 54, 311), radius=3, fill="#8fa1b4")

    with Image.open(src) as screenshot:
        screenshot = screenshot.convert("RGB")
        max_w = 900
        max_h = 650
        scale = min(max_w / screenshot.width, max_h / screenshot.height, 1.0)
        shot_size = (max(1, int(screenshot.width * scale)), max(1, int(screenshot.height * scale)))
        screenshot = screenshot.resize(shot_size, Image.Resampling.LANCZOS)

    shot_x = (width - screenshot.width) // 2
    shot_y = 390 + (650 - screenshot.height) // 2
    panel_pad = 24
    panel = (
        shot_x - panel_pad,
        shot_y - panel_pad,
        shot_x + screenshot.width + panel_pad,
        shot_y + screenshot.height + panel_pad,
    )

    shadow = Image.new("RGBA", THUMBNAIL_SIZE, (0, 0, 0, 0))
    shadow_draw = ImageDraw.Draw(shadow)
    shadow_draw.rounded_rectangle(
        (panel[0] + 6, panel[1] + 10, panel[2] + 6, panel[3] + 10),
        radius=24,
        fill=(31, 41, 51, 28),
    )
    shadow = shadow.filter(ImageFilter.GaussianBlur(10))
    canvas = Image.alpha_composite(canvas.convert("RGBA"), shadow).convert("RGB")
    draw = ImageDraw.Draw(canvas)
    draw.rounded_rectangle(panel, radius=24, fill="#ffffff", outline="#dce2e8", width=1)
    canvas.paste(screenshot, (shot_x, shot_y))
    canvas.save(dst, "JPEG", quality=94, optimize=True)


def create_zip(zip_path: Path, exe_path: Path, readme_path: Path, notice_path: Path) -> None:
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.write(exe_path, arcname=exe_path.name)
        archive.write(readme_path, arcname="README.txt")
        archive.write(notice_path, arcname="注意事項.txt")


def find_exe(app_dir: Path, meta: dict | None) -> Path | None:
    if meta and meta.get("exe_name"):
        expected = app_dir / "dist" / meta["exe_name"]
        if expected.exists():
            return expected
    exes = sorted((app_dir / "dist").glob("*.exe"))
    return exes[0] if exes else None


def process_app(app_dir: Path, *, dry_run: bool = False) -> AppResult:
    result = AppResult(folder=app_dir.name)
    readme_path = app_dir / "README.md"
    if not readme_path.exists():
        result.missing.append("README.md")
        result.ok = False
        return result

    readme_text = read_text(readme_path)
    try:
        meta = extract_meta(readme_text)
    except json.JSONDecodeError as exc:
        result.missing.append(f"DAKE_META invalid JSON: {exc}")
        result.ok = False
        meta = None

    if meta is None:
        result.missing.append("DAKE_META")
        result.ok = False
        return result

    result.status = str(meta.get("status") or "unknown")
    release_body = extract_release_body(readme_text)
    if not release_body:
        result.missing.append("RELEASE_BODY")

    release_body_path = app_dir / "release_body.md"
    if not release_body_path.exists() and release_body:
        if dry_run:
            result.generated.append("release_body.md (dry-run)")
        else:
            write_text(release_body_path, release_body + "\n")
            result.generated.append("release_body.md")
    elif not release_body_path.exists():
        result.missing.append("release_body.md")
    else:
        existing_body = read_text(release_body_path).strip()
        if existing_body and not release_body:
            release_body = existing_body

    title = product_title(meta)
    result.product_title = title
    description = short_description(meta)
    features = normalize_features(release_body, description)
    if not any("Windows" in feature for feature in features):
        features.append("Windows向け")
    features = features[:5]
    canonical_price = canonical_price_for(app_dir)
    if canonical_price is not None:
        result.price = canonical_price
        result.price_source = "ORIGINAL.md"
    else:
        result.price = price_for(app_dir.name, title, features)
        result.price_source = "inferred"
        result.warnings.append("price inferred because ORIGINAL.md price is missing")

    screenshot_webp = app_dir / meta.get("screenshot_path", "assets/screenshot.webp")
    screenshot_jpg = app_dir / "assets" / "screenshot.jpg"
    thumbnail_jpg = app_dir / "assets" / "booth_thumbnail.jpg"
    booth_dir = app_dir / "booth_ready"
    booth_jpg = booth_dir / "screenshot.jpg"
    booth_thumbnail_jpg = booth_dir / "booth_thumbnail.jpg"
    if not dry_run:
        booth_dir.mkdir(parents=True, exist_ok=True)

    if screenshot_webp.exists():
        if dry_run:
            result.generated.extend(["assets/screenshot.jpg (dry-run)", "booth_ready/screenshot.jpg (dry-run)"])
        else:
            try:
                convert_webp_to_jpg(screenshot_webp, screenshot_jpg)
                shutil.copy2(screenshot_jpg, booth_jpg)
                result.generated.extend(["assets/screenshot.jpg", "booth_ready/screenshot.jpg"])
            except Exception as exc:  # noqa: BLE001
                result.missing.append(f"screenshot.jpg conversion failed: {exc}")
                result.ok = False
    else:
        result.missing.append("assets/screenshot.webp")
        result.warnings.append("screenshot.jpg not generated")

    thumbnail_source_ready = screenshot_jpg.exists() or (dry_run and screenshot_webp.exists())
    if thumbnail_source_ready:
        if dry_run:
            result.generated.extend(
                ["assets/booth_thumbnail.jpg (dry-run)", "booth_ready/booth_thumbnail.jpg (dry-run)"]
            )
        else:
            try:
                create_booth_thumbnail(
                    screenshot_jpg,
                    thumbnail_jpg,
                    thumbnail_title(meta),
                    thumbnail_description(meta),
                )
                shutil.copy2(thumbnail_jpg, booth_thumbnail_jpg)
                result.generated.extend(["assets/booth_thumbnail.jpg", "booth_ready/booth_thumbnail.jpg"])
            except Exception as exc:  # noqa: BLE001
                result.thumbnail_error = str(exc)
                result.missing.append(f"booth_thumbnail.jpg generation failed: {exc}")
                result.ok = False
    else:
        result.thumbnail_error = "assets/screenshot.jpg missing"
        result.missing.append("assets/screenshot.jpg")

    readme_txt = booth_dir / "README.txt"
    notice_txt = booth_dir / "注意事項.txt"
    product_txt = booth_dir / "booth_product.txt"
    if product_txt.exists():
        result.booth_url = extract_booth_url(read_text(product_txt))

    exe_path = find_exe(app_dir, meta)
    if exe_path is None:
        result.missing.append("dist/*.exe")
        zip_name = f"{meta.get('exe_name') or app_dir.name}.zip".replace(".exe.zip", ".zip")
    else:
        zip_name = f"{exe_path.stem}.zip"
    result.zip_name = zip_name

    product = make_product_txt(
        title=title,
        price=result.price,
        description=description,
        features=features,
        tags=tags_for(app_dir.name, title, features),
        zip_name=zip_name,
        release_url=str(meta.get("release_url") or ""),
        booth_url=result.booth_url,
    )
    if not dry_run:
        write_text(readme_txt, make_readme_txt(title, description, features))
        write_text(notice_txt, NOTICE_TEXT)
        write_text(product_txt, product)
    result.product_updated = True
    if dry_run:
        result.generated.extend(
            [
                "booth_ready/README.txt (dry-run)",
                "booth_ready/注意事項.txt (dry-run)",
                "booth_ready/booth_product.txt (dry-run)",
            ]
        )
    else:
        result.generated.extend(["booth_ready/README.txt", "booth_ready/注意事項.txt", "booth_ready/booth_product.txt"])

    if exe_path is not None:
        if dry_run:
            result.generated.append(f"booth_ready/{zip_name} (dry-run)")
        else:
            create_zip(booth_dir / zip_name, exe_path, readme_txt, notice_txt)
            result.generated.append(f"booth_ready/{zip_name}")

    for required in ["build.bat", ".gitignore"]:
        if not (app_dir / required).exists():
            result.missing.append(required)

    if meta.get("release_url") == "":
        result.warnings.append("release_url empty")

    if result.missing:
        result.ok = False
    return result


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate DAKE BOOTH-ready assets.")
    parser.add_argument("--only-available", action="store_true", help="Process only status: available apps; skips frozen/draft/experimental/private/internal.")
    parser.add_argument("--app", help="Process a single app folder such as DAKE_Mail_Draft.")
    parser.add_argument("--dry-run", action="store_true", help="Show targets and actions without writing files.")
    return parser.parse_args()


def select_app_dirs(args: argparse.Namespace) -> tuple[list[Path], list[tuple[str, str]]]:
    app_dirs = sorted(path for path in APPS_DIR.iterdir() if path.is_dir() and path.name.startswith("DAKE_"))
    if args.app:
        app_dirs = [path for path in app_dirs if path.name == args.app]
    selected: list[Path] = []
    skipped: list[tuple[str, str]] = []
    for app_dir in app_dirs:
        _meta, status = app_meta_for_filter(app_dir)
        if args.only_available and status != "available":
            skipped.append((app_dir.name, status))
            continue
        selected.append(app_dir)
    return selected, skipped


def main() -> int:
    args = parse_args()
    app_dirs, skipped = select_app_dirs(args)
    if args.app and not app_dirs and not skipped:
        print(f"app not found: {args.app}")
        return 1
    results = [process_app(app_dir, dry_run=args.dry_run) for app_dir in app_dirs]

    mode = "dry-run" if args.dry_run else "write"
    print(f"mode: {mode}")
    print(f"only_available: {args.only_available}")
    if args.app:
        print(f"app: {args.app}")
    print(f"target apps: {len(results)}")
    print(f"skipped by status: {len(skipped)}")
    for name, status in skipped:
        print(f"- skipped {name}: status={status}")

    print(f"screenshot.jpg actions: {sum(any(item.startswith('assets/screenshot.jpg') for item in r.generated) for r in results)}")
    print(f"booth_thumbnail.jpg actions: {sum(any(item.startswith('assets/booth_thumbnail.jpg') for item in r.generated) for r in results)}")
    print(f"booth_ready exists: {sum((APPS_DIR / r.folder / 'booth_ready').exists() for r in results)}")
    print(f"zip actions: {sum(any('.zip' in item for item in r.generated) for r in results)}")
    print(f"booth_product.txt actions: {sum(r.product_updated for r in results)}")
    print(f"BOOTH URL preserved: {sum(1 for r in results if r.booth_url)}")
    print("")
    print("不足/警告:")
    for result in results:
        if result.missing or result.warnings:
            parts = []
            if result.missing:
                parts.append("missing=" + ", ".join(result.missing))
            if result.warnings:
                parts.append("warnings=" + ", ".join(result.warnings))
            print(f"- {result.folder}: {'; '.join(parts)}")

    print("")
    print("BOOTH初期候補:")
    if args.app:
        for result in results:
            print(
                f"- {result.folder}: {result.product_title} / {result.price}円 / "
                f"{result.zip_name} / price_source={result.price_source}"
            )
        return 0
    for name in ["DAKE_PDF_Merge", "DAKE_Sticky_Memo", "DAKE_Maji_Memo", "DAKE_Git_Memo", "DAKE_Yesterday_Task_Memo"]:
        match = next((r for r in results if r.folder == name), None)
        if match:
            print(f"- {match.folder}: {match.product_title} / {match.price}円 / {match.zip_name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
