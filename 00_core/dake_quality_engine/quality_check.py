from __future__ import annotations

import argparse
import json
import re
import struct
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

try:
    from .ui_guard import check_file
except ImportError:  # pragma: no cover - direct script fallback.
    from ui_guard import check_file  # type: ignore


REQUIRED_ICO_SIZES = frozenset({16, 32, 48, 256})
WINDOW_ICON_CALL_RE = re.compile(r"\b(?:wm_iconbitmap|iconbitmap|iconphoto|setWindowIcon)\s*\(")
BAT_ICON_RE = re.compile(r"""(?:--icon|-i)(?:=|\s+)(?:"([^"]+)"|'([^']+)'|([^\s^]+))""", re.IGNORECASE)
SPEC_ICON_RE = re.compile(r"""\bicon\s*=\s*(?:\[\s*)?r?["']([^"']+\.ico)["']""", re.IGNORECASE)

# Classic UTF-8-decoded-as-CP932 mojibake signatures and suspicious markers.
KNOWN_MOJIBAKE_RE = re.compile(
    r"(?:雋ｩ螢ｲ|迥ｶ諷|譁・ｭ|繝[ァ-ヶー]|繧[ァ-ヶー]|[\ufffd\u0080-\u009f])"
)
SUSPICIOUS_MOJIBAKE_SPAN_RE = re.compile(r"[^\x00-\x7f\s]{3,}")
# Rare CP932-decoded UTF-8 lead characters frequently appearing in Japanese mojibake.
MOJIBAKE_HINT_CHARS = frozenset("雋ｩ螢ｲ迥ｶ諷譁繝繧驛ｼ螟荳蜿蟷隴蛻鬆ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ")
READABLE_JA_RE = re.compile(r"^[ぁ-んァ-ヶ一-龯々ー、。・：（）「」『』！？0-9A-Za-z_\-\s]+$")


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
    return CheckResult("WARN" if warn else "FAIL", item, f"missing: {path}")


def detect_mojibake_lines(text: str) -> list[tuple[int, str]]:
    """Detect replacement characters and UTF-8/CP932 mojibake lines."""
    hits: list[tuple[int, str]] = []
    for line_no, line in enumerate(text.splitlines(), start=1):
        match = KNOWN_MOJIBAKE_RE.search(line)
        if match:
            hits.append((line_no, match.group(0)))
            continue
        # Round-trip heuristic for spans containing mojibake hint chars.
        if any(ch in MOJIBAKE_HINT_CHARS for ch in line):
            for span_match in SUSPICIOUS_MOJIBAKE_SPAN_RE.finditer(line):
                span = span_match.group(0)
                if not any(ch in MOJIBAKE_HINT_CHARS for ch in span):
                    continue
                try:
                    raw = span.encode("cp932")
                    decoded = raw.decode("utf-8")
                except (UnicodeEncodeError, UnicodeDecodeError):
                    continue
                if decoded != span and len(decoded) >= 2 and READABLE_JA_RE.match(decoded):
                    hits.append((line_no, f"{span} (-> {decoded})"))
                    break
    return hits


def check_original_md(app_dir: Path) -> list[CheckResult]:
    original_path = app_dir / "ORIGINAL.md"
    if not original_path.is_file():
        return [CheckResult("FAIL", "ORIGINAL.md", f"missing: {original_path}")]
    try:
        text = original_path.read_bytes().decode("utf-8")
    except UnicodeDecodeError as exc:
        return [CheckResult("FAIL", "ORIGINAL.md", f"invalid UTF-8: {exc}")]

    mojibake = detect_mojibake_lines(text)
    if mojibake:
        sample = ", ".join(f"line {ln}: {token}" for ln, token in mojibake[:5])
        return [CheckResult("FAIL", "ORIGINAL.md (mojibake)", f"detected: {sample}")]
    return [CheckResult("OK", "ORIGINAL.md", str(original_path))]


def parse_ico_resolutions(ico_path: Path) -> tuple[list[tuple[int, int, int]], str | None]:
    """Parse Windows .ico header and return [(width, height, bpp), ...] or an error."""
    try:
        data = ico_path.read_bytes()
    except OSError as exc:
        return [], f"cannot read ico: {exc}"
    if len(data) < 6:
        return [], "ico file too small"
    reserved, img_type, count = struct.unpack_from("<HHH", data, 0)
    if reserved != 0 or img_type != 1 or count < 1:
        return [], f"invalid ico header (reserved={reserved}, type={img_type}, count={count})"
    if len(data) < 6 + count * 16:
        return [], "truncated ico directory"
    entries: list[tuple[int, int, int]] = []
    for i in range(count):
        w, h, _cc, _res, _planes, bpp, _size, _offset = struct.unpack_from("<BBBBHHII", data, 6 + i * 16)
        entries.append((w or 256, h or 256, bpp))
    return entries, None


def pe_has_icon_resource(exe_path: Path) -> tuple[bool, str]:
    """Check whether a Windows PE .exe contains RT_ICON(3) and RT_GROUP_ICON(14) resources."""
    try:
        data = exe_path.read_bytes()
    except OSError as exc:
        return False, f"cannot read exe: {exc}"
    if len(data) < 0x40 or data[:2] != b"MZ":
        return False, "not a valid MZ/PE file"
    pe_offset = struct.unpack_from("<I", data, 0x3C)[0]
    if pe_offset + 24 > len(data) or data[pe_offset : pe_offset + 4] != b"PE\x00\x00":
        return False, "missing PE header"
    num_sections = struct.unpack_from("<H", data, pe_offset + 6)[0]
    opt_size = struct.unpack_from("<H", data, pe_offset + 20)[0]
    opt_offset = pe_offset + 24
    if opt_offset + 2 > len(data):
        return False, "truncated optional header"
    magic = struct.unpack_from("<H", data, opt_offset)[0]
    if magic == 0x10B:  # PE32
        dd_offset = opt_offset + 96
    elif magic == 0x20B:  # PE32+
        dd_offset = opt_offset + 112
    else:
        return False, f"unknown PE magic: {hex(magic)}"
    if dd_offset + 3 * 8 > len(data):
        return False, "truncated data directories"
    rsrc_rva, rsrc_size = struct.unpack_from("<II", data, dd_offset + 2 * 8)
    if rsrc_rva == 0 or rsrc_size == 0:
        return False, "no PE resource table"

    sec_offset = opt_offset + opt_size
    rsrc_raw: int | None = None
    for i in range(num_sections):
        s_off = sec_offset + i * 40
        if s_off + 40 > len(data):
            break
        v_size, v_addr, raw_size, raw_ptr = struct.unpack_from("<IIII", data, s_off + 8)
        span = max(v_size, raw_size)
        if v_addr <= rsrc_rva < v_addr + span:
            rsrc_raw = raw_ptr + (rsrc_rva - v_addr)
            break
    if rsrc_raw is None or rsrc_raw + 16 > len(data):
        return False, "resource section not mapped"

    num_named, num_id = struct.unpack_from("<HH", data, rsrc_raw + 12)
    entry_offset = rsrc_raw + 16
    res_ids: set[int] = set()
    for i in range(num_named + num_id):
        if entry_offset + i * 8 + 8 > len(data):
            break
        name_or_id, _ = struct.unpack_from("<II", data, entry_offset + i * 8)
        if not (name_or_id & 0x80000000):
            res_ids.add(name_or_id)

    has_icon = 3 in res_ids
    has_group = 14 in res_ids
    if has_icon and has_group:
        return True, f"RT_ICON(3) & RT_GROUP_ICON(14) embedded (resource types={sorted(res_ids)})"
    return False, f"icon resource missing (found resource types={sorted(res_ids)})"


def _extract_build_icon_refs(app_dir: Path) -> tuple[list[tuple[str, str]], list[str]]:
    """Return ([(source_label, raw_icon_path), ...], checked_files)."""
    refs: list[tuple[str, str]] = []
    checked: list[str] = []

    bat_path = app_dir / "build.bat"
    if bat_path.is_file():
        checked.append("build.bat")
        bat_text = bat_path.read_text(encoding="utf-8", errors="replace")
        for match in BAT_ICON_RE.finditer(bat_text):
            raw = next((g for g in match.groups() if g), "")
            if raw:
                refs.append(("build.bat", raw))

    for spec_path in sorted(app_dir.glob("*.spec")):
        checked.append(spec_path.name)
        spec_text = spec_path.read_text(encoding="utf-8", errors="replace")
        for match in SPEC_ICON_RE.finditer(spec_text):
            refs.append((spec_path.name, match.group(1)))

    return refs, checked


def check_icons(app_dir: Path, dist_exes: list[Path]) -> list[CheckResult]:
    results: list[CheckResult] = []

    # 1. Build --icon / .spec icon= check
    build_refs, checked_files = _extract_build_icon_refs(app_dir)
    resolved_ref_icos: list[Path] = []
    if not build_refs:
        checked_desc = ", ".join(checked_files) if checked_files else "build.bat / *.spec"
        results.append(CheckResult("FAIL", "Icon (build --icon)", f"no --icon / icon= specified in {checked_desc}"))
    else:
        missing_refs: list[str] = []
        for source_label, raw_ref in build_refs:
            normalized = raw_ref.replace("\\", "/")
            candidate = (app_dir / normalized).resolve()
            if candidate.is_file():
                resolved_ref_icos.append(candidate)
            else:
                missing_refs.append(f"{source_label}: {raw_ref}")
        if missing_refs:
            results.append(
                CheckResult("FAIL", "Icon (build --icon)", f"referenced icon not found: {', '.join(missing_refs)}")
            )
        else:
            ref_desc = ", ".join(f"{src} -> {raw}" for src, raw in build_refs)
            results.append(CheckResult("OK", "Icon (build --icon)", ref_desc))

    # 2. .ico existence check (local assets/*.ico or valid shared icon referenced by build)
    local_icos = sorted((app_dir / "assets").glob("*.ico")) if (app_dir / "assets").is_dir() else []
    target_ico: Path | None = None
    if local_icos:
        target_ico = local_icos[0]
        results.append(CheckResult("OK", "Icon (.ico file)", f"local: {target_ico}"))
    elif resolved_ref_icos:
        target_ico = resolved_ref_icos[0]
        results.append(CheckResult("OK", "Icon (.ico file)", f"shared: {target_ico}"))
    else:
        results.append(
            CheckResult(
                "FAIL",
                "Icon (.ico file)",
                f"missing assets/*.ico and no valid --icon target in {app_dir}",
            )
        )

    # 3. .ico embedded resolutions check
    if target_ico is not None:
        entries, err = parse_ico_resolutions(target_ico)
        if err:
            results.append(CheckResult("FAIL", "Icon (.ico resolutions)", f"{target_ico.name}: {err}"))
        else:
            widths = {w for w, h, _bpp in entries if w == h}
            missing_sizes = sorted(REQUIRED_ICO_SIZES - widths)
            size_str = ", ".join(f"{w}x{h}" for w, h, _bpp in entries)
            if missing_sizes:
                results.append(
                    CheckResult(
                        "FAIL",
                        "Icon (.ico resolutions)",
                        f"{target_ico.name} missing required sizes {missing_sizes} (found: {size_str})",
                    )
                )
            else:
                results.append(
                    CheckResult("OK", "Icon (.ico resolutions)", f"{target_ico.name} ({len(entries)} sizes: {size_str})")
                )

    # 4. main.py window icon setup check
    main_path = app_dir / "main.py"
    if main_path.is_file():
        main_text = main_path.read_text(encoding="utf-8", errors="replace")
        if WINDOW_ICON_CALL_RE.search(main_text):
            results.append(CheckResult("OK", "Icon (main.py window)", "window icon call detected"))
        else:
            results.append(
                CheckResult(
                    "FAIL",
                    "Icon (main.py window)",
                    "missing window icon setup (iconbitmap / wm_iconbitmap / iconphoto / setWindowIcon)",
                )
            )
    else:
        results.append(CheckResult("FAIL", "Icon (main.py window)", "main.py missing"))

    # 5. dist/*.exe embedded PE icon resource check
    if dist_exes:
        for exe_path in dist_exes:
            ok, detail = pe_has_icon_resource(exe_path)
            results.append(CheckResult("OK" if ok else "FAIL", f"Icon (EXE resource: {exe_path.name})", detail))

    return results


def check_app(app_dir: Path) -> list[CheckResult]:
    app_dir = app_dir.resolve()
    results: list[CheckResult] = []

    results.extend(check_original_md(app_dir))

    readme_path = app_dir / "README.md"
    results.append(_exists(readme_path, "README.md"))
    readme_text = ""
    meta: dict[str, Any] | None = None
    if readme_path.exists():
        try:
            readme_text = _read_readme(app_dir)
            results.append(
                CheckResult(
                    "OK" if "## RELEASE_BODY" in readme_text else "FAIL",
                    "RELEASE_BODY",
                    "present" if "## RELEASE_BODY" in readme_text else "missing",
                )
            )
            meta = _extract_dake_meta(readme_text)
            results.append(
                CheckResult("OK" if meta else "FAIL", "DAKE_META", "valid JSON" if meta else "missing or invalid")
            )
        except Exception as exc:
            results.append(CheckResult("FAIL", "README parse", str(exc)))
    else:
        results.append(CheckResult("FAIL", "DAKE_META", "README missing"))
        results.append(CheckResult("FAIL", "RELEASE_BODY", "README missing"))

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
            "OK" if any(path.exists() for path in booth_product_candidates) else "FAIL",
            "booth_product.txt",
            "present" if any(path.exists() for path in booth_product_candidates) else "missing",
        )
    )

    dist_exes = sorted((app_dir / "dist").glob("*.exe")) if (app_dir / "dist").is_dir() else []
    results.append(
        CheckResult(
            "OK" if dist_exes else "WARN",
            "dist/*.exe",
            ", ".join(path.name for path in dist_exes) if dist_exes else "missing",
        )
    )

    results.extend(check_icons(app_dir, dist_exes))

    if meta is not None:
        release_url = str(meta.get("release_url", "")).strip()
        results.append(CheckResult("OK" if release_url else "WARN", "DAKE_META.release_url", release_url or "empty"))

    for issue in check_file(app_dir / "main.py"):
        results.append(
            CheckResult(
                issue.level,
                "UI Guard",
                f"line {issue.line}: {issue.message}" if issue.line else issue.message,
            )
        )

    return results


def print_results(results: list[CheckResult]) -> int:
    has_fail = False
    for result in results:
        if result.level in ("FAIL", "NG"):
            has_fail = True
        print(f"[{result.level}] {result.item}: {result.detail}")
    print("RESULT:", "FAIL" if has_fail else "OK")
    return 1 if has_fail else 0


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="DAKE shipping quality checker")
    parser.add_argument("--app", required=True, help="Path to a DAKE app directory")
    args = parser.parse_args(argv)
    return print_results(check_app(Path(args.app)))


if __name__ == "__main__":
    raise SystemExit(main())
