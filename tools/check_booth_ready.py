"""Check DAKE app BOOTH-ready assets and icon wiring.

The script audits every 01_apps/DAKE_* app and writes a Markdown and CSV
report under tools/reports/. It does not launch apps, capture screenshots, or
publish anything; missing items are reported with the next safe action.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
APPS_DIR = ROOT / "01_apps"
REPORT_DIR = ROOT / "tools" / "reports"
COMMON_ICON_PARTS = ("02_assets", "dake_icon.ico")
ROOT_GITIGNORE_REQUIRED = [
    "build/",
    "dist/",
    "*.spec",
    "*_config.json",
    "__pycache__/",
    "*.pyc",
    "playwright_profile/",
]


@dataclass
class AppCheck:
    app: str
    readme: bool = False
    dake_meta: bool = False
    release_body: bool = False
    screenshot: bool = False
    thumbnail: bool = False
    booth_product: bool = False
    booth_ready: bool = False
    zip_file: bool = False
    build_bat: bool = False
    main_py: bool = False
    dist_exe: bool = False
    icon_build: bool = False
    icon_main: bool = False
    actions: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return all(
            [
                self.readme,
                self.dake_meta,
                self.release_body,
                self.screenshot,
                self.thumbnail,
                self.booth_product,
                self.booth_ready,
                self.zip_file,
                self.build_bat,
                self.main_py,
                self.dist_exe,
                self.icon_build,
                self.icon_main,
            ]
        )


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")


def normalize_path_text(text: str) -> str:
    return text.replace("/", "\\").lower()


def has_common_icon(text: str) -> bool:
    normalized = normalize_path_text(text)
    return all(part.lower() in normalized for part in COMMON_ICON_PARTS)


def has_release_body(readme_text: str) -> bool:
    return "## RELEASE_BODY" in readme_text and bool(
        readme_text.split("## RELEASE_BODY", 1)[1].strip()
    )


def has_valid_dake_meta(readme_text: str) -> bool:
    if "## DAKE_META" not in readme_text or "```json" not in readme_text:
        return False
    try:
        block = readme_text.split("## DAKE_META", 1)[1].split("```json", 1)[1].split("```", 1)[0]
        json.loads(block)
    except (IndexError, json.JSONDecodeError):
        return False
    return True


def find_booth_product(app_dir: Path) -> Path | None:
    candidates = [
        app_dir / "booth_product.txt",
        app_dir / "booth_ready" / "booth_product.txt",
    ]
    return next((path for path in candidates if path.exists()), None)


def check_icon_build(build_text: str) -> bool:
    return "--icon" in build_text and has_common_icon(build_text)


def check_icon_main(main_text: str) -> bool:
    is_tk_app = "tkinter" in main_text or "tk.Tk" in main_text or "tk." in main_text
    is_win32_app = "CreateWindowEx" in main_text or "WNDCLASSEX" in main_text
    if is_win32_app:
        return has_common_icon(main_text) and "LoadImage" in main_text and "SendMessage" in main_text
    if not is_tk_app:
        return True
    if "iconbitmap" not in main_text or not has_common_icon(main_text):
        return False
    icon_positions = [match.start() for match in re.finditer(r"iconbitmap", main_text)]
    for position in icon_positions:
        window = main_text[max(0, position - 700) : min(len(main_text), position + 700)]
        if "try:" in window and "except" in window:
            return True
    return False


def add_action(result: AppCheck, condition: bool, action: str) -> None:
    if not condition:
        result.actions.append(action)


def check_app(app_dir: Path) -> AppCheck:
    result = AppCheck(app=app_dir.name)

    readme_path = app_dir / "README.md"
    release_body_path = app_dir / "release_body.md"
    screenshot_path = app_dir / "assets" / "screenshot.webp"
    thumbnail_path = app_dir / "assets" / "booth_thumbnail.jpg"
    booth_ready_dir = app_dir / "booth_ready"
    build_path = app_dir / "build.bat"
    main_path = app_dir / "main.py"
    dist_dir = app_dir / "dist"

    result.readme = readme_path.exists()
    if result.readme:
        readme_text = read_text(readme_path)
        result.dake_meta = has_valid_dake_meta(readme_text)
        readme_has_release = has_release_body(readme_text)
    else:
        readme_has_release = False

    result.release_body = release_body_path.exists()
    result.screenshot = screenshot_path.exists()
    result.thumbnail = thumbnail_path.exists()
    result.booth_product = find_booth_product(app_dir) is not None
    result.booth_ready = booth_ready_dir.exists()
    result.zip_file = booth_ready_dir.exists() and any(booth_ready_dir.glob("*.zip"))
    result.build_bat = build_path.exists()
    result.main_py = main_path.exists()
    result.dist_exe = dist_dir.exists() and any(dist_dir.glob("*.exe"))

    if result.build_bat:
        result.icon_build = check_icon_build(read_text(build_path))

    if result.main_py:
        result.icon_main = check_icon_main(read_text(main_path))

    add_action(result, result.readme, "create README.md")
    add_action(result, result.dake_meta, "add or fix DAKE_META JSON")
    if not result.release_body:
        if result.readme and readme_has_release:
            result.actions.append("generate release_body.md from README RELEASE_BODY")
        else:
            result.actions.append("add RELEASE_BODY to README and generate release_body.md")
    add_action(result, result.screenshot, "need screenshot: create assets/screenshot.webp")
    add_action(result, result.thumbnail, "generate assets/booth_thumbnail.jpg with tools/make_booth_ready.py")
    add_action(result, result.booth_product, "generate booth_product.txt with tools/make_booth_ready.py")
    add_action(result, result.booth_ready, "generate booth_ready/ with tools/make_booth_ready.py")
    if not result.zip_file:
        if result.dist_exe:
            result.actions.append("generate BOOTH ready zip from dist/*.exe")
        else:
            result.actions.append("need build: create dist/*.exe before zip")
    add_action(result, result.build_bat, "create build.bat")
    add_action(result, result.main_py, "create main.py")
    add_action(result, result.dist_exe, "need build: create dist/*.exe")
    add_action(result, result.icon_build, "add common --icon setting to build.bat")
    add_action(result, result.icon_main, "add safe common icon setting to main.py")

    return result


def root_gitignore_missing() -> list[str]:
    path = ROOT / ".gitignore"
    if not path.exists():
        return ROOT_GITIGNORE_REQUIRED
    text = read_text(path)
    return [pattern for pattern in ROOT_GITIGNORE_REQUIRED if pattern not in text]


def yn(value: bool) -> str:
    return "OK" if value else "NG"


def write_markdown(results: list[AppCheck], missing_gitignore: list[str], path: Path) -> None:
    checked = len(results)
    ok_results = [result for result in results if result.ok]
    missing_results = [result for result in results if not result.ok]
    lines: list[str] = [
        "# DAKE BOOTH Ready Check",
        "",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
        "## Summary",
        "",
        f"- checked: {checked}",
        f"- ok: {len(ok_results)}",
        f"- missing: {len(missing_results)}",
        f"- screenshot.webp missing: {sum(not r.screenshot for r in results)}",
        f"- booth_thumbnail.jpg missing: {sum(not r.thumbnail for r in results)}",
        f"- booth_product.txt missing: {sum(not r.booth_product for r in results)}",
        f"- booth_ready/ missing: {sum(not r.booth_ready for r in results)}",
        f"- zip missing: {sum(not r.zip_file for r in results)}",
        f"- icon build missing: {sum(not r.icon_build for r in results)}",
        f"- icon main missing: {sum(not r.icon_main for r in results)}",
        "",
        "## Git Ignore",
        "",
    ]
    if missing_gitignore:
        lines.extend(f"- missing: `{pattern}`" for pattern in missing_gitignore)
    else:
        lines.append("- OK")

    lines.extend(
        [
            "",
            "## Missing",
            "",
            "| app | README | DAKE_META | release_body | screenshot | thumbnail | booth_product | booth_ready | zip | dist_exe | icon_build | icon_main | next_action |",
            "| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |",
        ]
    )
    for result in missing_results:
        action = "<br>".join(result.actions) if result.actions else "-"
        lines.append(
            "| "
            + " | ".join(
                [
                    result.app,
                    yn(result.readme),
                    yn(result.dake_meta),
                    yn(result.release_body),
                    yn(result.screenshot),
                    yn(result.thumbnail),
                    yn(result.booth_product),
                    yn(result.booth_ready),
                    yn(result.zip_file),
                    yn(result.dist_exe),
                    yn(result.icon_build),
                    yn(result.icon_main),
                    action,
                ]
            )
            + " |"
        )

    lines.extend(["", "## OK", "", "| app | status |", "| --- | --- |"])
    for result in ok_results:
        lines.append(f"| {result.app} | OK |")

    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_csv(results: list[AppCheck], path: Path) -> None:
    fields = [
        "app",
        "ok",
        "readme",
        "dake_meta",
        "release_body",
        "screenshot",
        "thumbnail",
        "booth_product",
        "booth_ready",
        "zip",
        "build_bat",
        "main_py",
        "dist_exe",
        "icon_build",
        "icon_main",
        "actions",
    ]
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for result in results:
            writer.writerow(
                {
                    "app": result.app,
                    "ok": result.ok,
                    "readme": result.readme,
                    "dake_meta": result.dake_meta,
                    "release_body": result.release_body,
                    "screenshot": result.screenshot,
                    "thumbnail": result.thumbnail,
                    "booth_product": result.booth_product,
                    "booth_ready": result.booth_ready,
                    "zip": result.zip_file,
                    "build_bat": result.build_bat,
                    "main_py": result.main_py,
                    "dist_exe": result.dist_exe,
                    "icon_build": result.icon_build,
                    "icon_main": result.icon_main,
                    "actions": "; ".join(result.actions),
                }
            )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Check DAKE BOOTH-ready assets.")
    parser.add_argument("--report-dir", type=Path, default=REPORT_DIR)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    report_dir = args.report_dir
    report_dir.mkdir(parents=True, exist_ok=True)

    app_dirs = sorted(path for path in APPS_DIR.iterdir() if path.is_dir() and path.name.startswith("DAKE_"))
    results = [check_app(app_dir) for app_dir in app_dirs]
    missing_gitignore = root_gitignore_missing()

    md_path = report_dir / "booth_ready_check.md"
    csv_path = report_dir / "booth_ready_check.csv"
    write_markdown(results, missing_gitignore, md_path)
    write_csv(results, csv_path)

    checked = len(results)
    ok_count = sum(result.ok for result in results)
    missing_count = checked - ok_count
    print("DAKE BOOTH Ready Check")
    print(f"checked: {checked}")
    print(f"ok: {ok_count}")
    print(f"missing: {missing_count}")
    print(f"screenshot.webp missing: {sum(not r.screenshot for r in results)}")
    print(f"booth_thumbnail.jpg missing: {sum(not r.thumbnail for r in results)}")
    print(f"booth_product.txt missing: {sum(not r.booth_product for r in results)}")
    print(f"booth_ready/ missing: {sum(not r.booth_ready for r in results)}")
    print(f"zip missing: {sum(not r.zip_file for r in results)}")
    print(f"icon build missing: {sum(not r.icon_build for r in results)}")
    print(f"icon main missing: {sum(not r.icon_main for r in results)}")
    if missing_gitignore:
        print("gitignore missing: " + ", ".join(missing_gitignore))
    else:
        print("gitignore: OK")
    print(f"report: {md_path}")
    print(f"csv: {csv_path}")
    if missing_count:
        print("missing apps:")
        for result in results:
            if not result.ok:
                print(f"- {result.app}: {', '.join(result.actions)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
