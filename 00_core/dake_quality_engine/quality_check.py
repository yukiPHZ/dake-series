from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

try:
    from .ui_guard import check_file
except ImportError:  # pragma: no cover - direct script fallback.
    from ui_guard import check_file  # type: ignore


@dataclass(frozen=True)
class CheckResult:
    level: str
    item: str
    detail: str


def _read_readme(app_dir: Path) -> str:
    return (app_dir / "README.md").read_text(encoding="utf-8")


def _extract_dake_meta(readme_text: str) -> dict[str, Any] | None:
    match = re.search(r"##\s+DAKE_META\s*```json\s*(\{.*?\})\s*```", readme_text, re.DOTALL)
    if not match:
        return None
    try:
        value = json.loads(match.group(1))
    except json.JSONDecodeError:
        return None
    return value if isinstance(value, dict) else None


def _exists(path: Path, item: str, *, warn: bool = False) -> CheckResult:
    if path.exists():
        return CheckResult("OK", item, str(path))
    return CheckResult("WARN" if warn else "NG", item, f"missing: {path}")


def check_app(app_dir: Path) -> list[CheckResult]:
    app_dir = app_dir.resolve()
    results: list[CheckResult] = []

    readme_path = app_dir / "README.md"
    results.append(_exists(readme_path, "README.md"))
    readme_text = ""
    meta: dict[str, Any] | None = None
    if readme_path.exists():
        try:
            readme_text = _read_readme(app_dir)
            results.append(
                CheckResult("OK" if "## RELEASE_BODY" in readme_text else "NG", "RELEASE_BODY", "present" if "## RELEASE_BODY" in readme_text else "missing")
            )
            meta = _extract_dake_meta(readme_text)
            results.append(CheckResult("OK" if meta else "NG", "DAKE_META", "valid JSON" if meta else "missing or invalid"))
        except Exception as exc:
            results.append(CheckResult("NG", "README parse", str(exc)))
    else:
        results.append(CheckResult("NG", "DAKE_META", "README missing"))
        results.append(CheckResult("NG", "RELEASE_BODY", "README missing"))

    results.extend(
        [
            _exists(app_dir / "release_body.md", "release_body.md"),
            _exists(app_dir / "assets" / "screenshot.webp", "assets/screenshot.webp"),
            _exists(app_dir / "assets" / "booth_thumbnail.jpg", "assets/booth_thumbnail.jpg"),
            _exists(app_dir / "booth_ready", "booth_ready/"),
            _exists(app_dir / "build.bat", "build.bat"),
            _exists(app_dir / "main.py", "main.py"),
        ]
    )

    booth_product_candidates = [app_dir / "booth_product.txt", app_dir / "booth_ready" / "booth_product.txt"]
    results.append(
        CheckResult(
            "OK" if any(path.exists() for path in booth_product_candidates) else "NG",
            "booth_product.txt",
            "present" if any(path.exists() for path in booth_product_candidates) else "missing",
        )
    )

    dist_exes = list((app_dir / "dist").glob("*.exe"))
    results.append(CheckResult("OK" if dist_exes else "WARN", "dist/*.exe", ", ".join(path.name for path in dist_exes) if dist_exes else "missing"))

    if meta is not None:
        release_url = str(meta.get("release_url", "")).strip()
        results.append(CheckResult("OK" if release_url else "WARN", "DAKE_META.release_url", release_url or "empty"))

    for issue in check_file(app_dir / "main.py"):
        results.append(CheckResult(issue.level, "UI Guard", f"line {issue.line}: {issue.message}" if issue.line else issue.message))

    return results


def print_results(results: list[CheckResult]) -> int:
    has_ng = False
    for result in results:
        if result.level == "NG":
            has_ng = True
        print(f"[{result.level}] {result.item}: {result.detail}")
    print("RESULT:", "NG" if has_ng else "OK")
    return 1 if has_ng else 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="DAKE shipping quality checker")
    parser.add_argument("--app", required=True, help="Path to a DAKE app directory")
    args = parser.parse_args(argv)
    return print_results(check_app(Path(args.app)))


if __name__ == "__main__":
    raise SystemExit(main())
