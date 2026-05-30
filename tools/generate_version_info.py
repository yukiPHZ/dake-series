from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

COMPANY_NAME = "\u3057\u307e\u308a\u3059\u4e0d\u52d5\u7523"
COPYRIGHT = "\u00a9 2026 \u3057\u307e\u308a\u3059\u4e0d\u52d5\u7523 \u2014 Vibe-Coded by Yukihiko Kikuta"
DEFAULT_VERSION = "1.0.0"


def read_meta(app_dir: Path) -> dict[str, Any]:
    readme = app_dir / "README.md"
    text = readme.read_text(encoding="utf-8")
    match = re.search(r"##\s+DAKE_META\s*```json\s*(\{.*?\})\s*```", text, re.DOTALL)
    if not match:
        raise ValueError(f"DAKE_META not found: {readme}")
    return json.loads(match.group(1))


def version_from_meta(meta: dict[str, Any]) -> str:
    value = str(meta.get("version", "")).strip()
    if value:
        return value.lstrip("v")
    release_url = str(meta.get("release_url", "")).strip()
    match = re.search(r"[_-]v(\d+\.\d+\.\d+)", release_url)
    return match.group(1) if match else DEFAULT_VERSION


def version_tuple(version: str) -> tuple[int, int, int, int]:
    parts = [int(part) for part in re.findall(r"\d+", version)[:4]]
    while len(parts) < 4:
        parts.append(0)
    return tuple(parts[:4])  # type: ignore[return-value]


def escape(value: object) -> str:
    return str(value).replace("\\", "\\\\").replace("'", "\\'")


def build_version_info(meta: dict[str, Any]) -> str:
    display_name = str(meta.get("display_name") or meta.get("site_title") or meta.get("launcher_title") or "DAKE")
    file_description = str(meta.get("site_title") or meta.get("launcher_title") or display_name)
    exe_name = str(meta.get("exe_name") or f"{display_name}.exe")
    version = version_from_meta(meta)
    filevers = version_tuple(version)
    productvers = filevers
    strings = {
        "CompanyName": COMPANY_NAME,
        "FileDescription": file_description,
        "FileVersion": version,
        "InternalName": Path(exe_name).stem,
        "LegalCopyright": COPYRIGHT,
        "OriginalFilename": exe_name,
        "ProductName": display_name,
        "ProductVersion": version,
    }
    string_lines = "\n".join(
        f"        StringStruct('{escape(key)}', '{escape(value)}')," for key, value in strings.items()
    )
    return f"""# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={filevers},
    prodvers={productvers},
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable(
        '041104b0',
[
{string_lines}
      ])
    ]),
    VarFileInfo([VarStruct('Translation', [1041, 1200])])
  ]
)
"""


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate PyInstaller VersionInfo from README DAKE_META")
    parser.add_argument("--app", "--app-dir", dest="app", required=True, help="DAKE app directory")
    parser.add_argument("--out", default="version_info.txt", help="output version info path")
    parser.add_argument("--print-summary", action="store_true")
    args = parser.parse_args()

    app_dir = Path(args.app).resolve()
    out_path = Path(args.out)
    if not out_path.is_absolute():
        out_path = app_dir / out_path
    meta = read_meta(app_dir)
    out_path.write_text(build_version_info(meta), encoding="utf-8")
    if args.print_summary:
        print(f"generated: {out_path}")
        print(f"product: {meta.get('display_name', '')}")
        print(f"version: {version_from_meta(meta)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
