from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

JAPANESE_RE = r"ぁ-んァ-ン一-龥ー"
DIRECT_TEXT_RE = re.compile(r"\btext\s*=\s*([\"'])(?P<text>[^\"']*[" + JAPANESE_RE + r"][^\"']*)\1")
CONSTANTS = ("APP_NAME", "WINDOW_TITLE", "UI_TEXT")
LEGACY_FOOTER_KEYS = (
    "footer_left",
    "footer_caption",
    "footer_link_1",
    "footer_link_2",
    "footer_separator",
    "footer_copyright",
)
CURRENT_FOOTER_KEYS = ("footer_kaitori", "footer_author")

CANONICAL_FOOTER_LINK_TEXT = "戸建買取査定"
CANONICAL_FOOTER_AUTHOR_TEXT = "by Yukihiko Kikuta"
CANONICAL_KAITORI_URL_RE = re.compile(r"https://dakeapp\.com/go/kaitori\?app=[a-z0-9_]+")

FORBIDDEN_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("Vibe-Coded", re.compile(r"\bvibe[\s-]*cod(?:ing|ed)\b", re.IGNORECASE)),
    ("Instagram", re.compile(r"instagram", re.IGNORECASE)),
    ("sakurayk.notion.site", re.compile(r"sakurayk\.notion\.site", re.IGNORECASE)),
)
FOOTER_CONTEXT_RE = re.compile(r"(?:footer|copyright|credit|links|©|by\s+yukihiko)", re.IGNORECASE)


@dataclass(frozen=True)
class UiGuardIssue:
    level: str
    message: str
    line: int | None = None


def _find_footer_shimarisu_lines(lines: list[str]) -> list[int]:
    """Return 1-based line numbers where 'しまりす不動産' appears in footer/copyright/UI_TEXT context."""
    hits: list[int] = []
    in_footer_func = False
    footer_indent = 0
    in_ui_text = False
    ui_brace_depth = 0

    for index, line in enumerate(lines, start=1):
        stripped = line.lstrip()
        indent = len(line) - len(stripped)

        if stripped.startswith("def "):
            if "footer" in stripped.lower():
                in_footer_func = True
                footer_indent = indent
            elif in_footer_func and indent <= footer_indent:
                in_footer_func = False
        elif in_footer_func and stripped and not stripped.startswith("#") and indent <= footer_indent:
            in_footer_func = False

        if re.match(r"^\s*UI_TEXT\s*=\s*\{", line):
            in_ui_text = True
            ui_brace_depth = line.count("{") - line.count("}")
            if ui_brace_depth <= 0:
                in_ui_text = False
        elif in_ui_text:
            ui_brace_depth += line.count("{") - line.count("}")

        if "しまりす不動産" in line:
            if in_footer_func or in_ui_text or FOOTER_CONTEXT_RE.search(line):
                hits.append(index)

        if in_ui_text and ui_brace_depth <= 0:
            in_ui_text = False

    return hits


def check_source(source: str) -> list[UiGuardIssue]:
    issues: list[UiGuardIssue] = []
    lines = source.splitlines()

    for name in CONSTANTS:
        if not re.search(rf"^\s*{re.escape(name)}\s*=", source, re.MULTILINE):
            issues.append(UiGuardIssue("WARN", f"{name} is not defined"))

    # 1. Legacy footer keys must FAIL
    for key in LEGACY_FOOTER_KEYS:
        key_pattern = re.compile(rf"""["']{re.escape(key)}["']""")
        for index, line in enumerate(lines, start=1):
            if key_pattern.search(line):
                issues.append(UiGuardIssue("FAIL", f"legacy footer key detected: {key}", index))
                break

    # 2. Forbidden strings must FAIL
    for label, pattern in FORBIDDEN_PATTERNS:
        for index, line in enumerate(lines, start=1):
            if pattern.search(line):
                issues.append(UiGuardIssue("FAIL", f"forbidden footer/credit string detected: {label}", index))
                break

    for line_no in _find_footer_shimarisu_lines(lines):
        issues.append(UiGuardIssue("FAIL", "forbidden footer string detected: しまりす不動産", line_no))
        break

    # 3. Canonical new footer elements must ALL be present, otherwise FAIL
    if CANONICAL_FOOTER_LINK_TEXT not in source:
        issues.append(
            UiGuardIssue("FAIL", f"canonical footer link text missing: {CANONICAL_FOOTER_LINK_TEXT}")
        )
    if CANONICAL_FOOTER_AUTHOR_TEXT not in source:
        issues.append(
            UiGuardIssue("FAIL", f"canonical footer author text missing: {CANONICAL_FOOTER_AUTHOR_TEXT}")
        )
    if not CANONICAL_KAITORI_URL_RE.search(source):
        issues.append(
            UiGuardIssue(
                "FAIL",
                "canonical footer kaitori URL missing: https://dakeapp.com/go/kaitori?app=<app_id>",
            )
        )

    # 4. Direct Japanese text= hardcoded outside UI_TEXT must FAIL
    for index, line in enumerate(lines, start=1):
        match = DIRECT_TEXT_RE.search(line)
        if match:
            issues.append(
                UiGuardIssue(
                    "FAIL",
                    f"direct Japanese text= detected: {match.group('text')}",
                    index,
                )
            )

    return issues


def check_file(path: str | Path) -> list[UiGuardIssue]:
    file_path = Path(path)
    if not file_path.exists():
        return [UiGuardIssue("FAIL", f"main.py not found: {file_path}")]
    return check_source(file_path.read_text(encoding="utf-8"))
