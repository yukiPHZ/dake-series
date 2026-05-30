from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

JAPANESE_RE = r"ぁ-んァ-ン一-龥ー"
DIRECT_TEXT_RE = re.compile(r"\btext\s*=\s*([\"'])(?P<text>[^\"']*[" + JAPANESE_RE + r"][^\"']*)\1")
CONSTANTS = ("APP_NAME", "WINDOW_TITLE", "COPYRIGHT", "UI_TEXT")
FOOTER_KEYS = (
    "footer_left",
    "footer_caption",
    "footer_link_1",
    "footer_link_2",
    "footer_separator",
    "footer_copyright",
)


@dataclass(frozen=True)
class UiGuardIssue:
    level: str
    message: str
    line: int | None = None


def check_source(source: str) -> list[UiGuardIssue]:
    issues: list[UiGuardIssue] = []
    for name in CONSTANTS:
        if not re.search(rf"^\s*{re.escape(name)}\s*=", source, re.MULTILINE):
            issues.append(UiGuardIssue("WARN", f"{name} is not defined"))

    for key in FOOTER_KEYS:
        if key not in source:
            issues.append(UiGuardIssue("WARN", f"UI_TEXT footer key missing: {key}"))

    lines = source.splitlines()
    for index, line in enumerate(lines, start=1):
        match = DIRECT_TEXT_RE.search(line)
        if match:
            issues.append(
                UiGuardIssue(
                    "NG",
                    f"direct Japanese text= detected: {match.group('text')}",
                    index,
                )
            )
    return issues


def check_file(path: str | Path) -> list[UiGuardIssue]:
    file_path = Path(path)
    if not file_path.exists():
        return [UiGuardIssue("NG", f"main.py not found: {file_path}")]
    return check_source(file_path.read_text(encoding="utf-8"))
