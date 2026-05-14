from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass
from pathlib import Path

from core.app_config import logs_dir, now_iso, read_text_safe
from core.db import BrainzDatabase, DocumentRecord
from core.ollama_embeddings import generate_embeddings_for_document
from core.text_splitter import split_text


SOURCE_TYPE_CODEX = "codex_result"


@dataclass(frozen=True)
class CodexParsedResult:
    title: str
    summary: str
    changed_files: list[str]
    created_files: list[str]
    test_results: str
    build_results: str
    commit_hash: str
    push_result: str
    git_status: str
    phase_notes: str
    raw_text: str
    content_hash: str
    imported_at: str


@dataclass(frozen=True)
class CodexImportResult:
    title: str
    changed: bool
    skipped_duplicate: bool
    changed_files_count: int
    commit_hash: str
    content_hash: str
    log_path: str


def import_codex_text(raw_text: str, database: BrainzDatabase, source_label: str = "") -> CodexImportResult:
    parsed = parse_codex_result(raw_text)
    if not parsed.raw_text.strip():
        raise ValueError("empty codex result")

    document = build_document_record(parsed, source_label=source_label)
    chunks = split_text(document.content)
    document_id, changed = database.upsert_document(document, chunks)
    if changed:
        try:
            generate_embeddings_for_document(database, document_id)
        except Exception:
            pass
    database.upsert_codex_result(parsed, document.path)
    result_without_log = CodexImportResult(
        title=parsed.title,
        changed=changed,
        skipped_duplicate=not changed,
        changed_files_count=len(parsed.changed_files),
        commit_hash=parsed.commit_hash,
        content_hash=parsed.content_hash,
        log_path="",
    )
    log_path = write_import_log(result_without_log)
    return CodexImportResult(
        title=result_without_log.title,
        changed=result_without_log.changed,
        skipped_duplicate=result_without_log.skipped_duplicate,
        changed_files_count=result_without_log.changed_files_count,
        commit_hash=result_without_log.commit_hash,
        content_hash=result_without_log.content_hash,
        log_path=str(log_path),
    )


def import_codex_file(path: Path, database: BrainzDatabase) -> CodexImportResult:
    text = read_text_safe(path)
    return import_codex_text(text, database, source_label=str(path.resolve()))


def parse_codex_result(raw_text: str) -> CodexParsedResult:
    text = (raw_text or "").strip()
    content_hash = hashlib.sha256(text.encode("utf-8", errors="replace")).hexdigest()
    lines = [line.strip() for line in text.splitlines()]
    nonempty = [line for line in lines if line]
    commit_hash = extract_commit_hash(text)
    changed_files = extract_changed_files(lines, text)
    created_files = extract_created_files(lines, changed_files)
    test_results = extract_matching_lines(lines, TEST_PATTERNS)
    build_results = extract_matching_lines(lines, BUILD_PATTERNS)
    push_result = extract_first_matching_line(lines, PUSH_PATTERNS)
    git_status = extract_first_matching_line(lines, GIT_STATUS_PATTERNS)
    phase_notes = extract_matching_lines(lines, PHASE_PATTERNS)
    title = extract_title(nonempty, commit_hash)
    summary = extract_summary(nonempty)
    imported_at = now_iso()
    return CodexParsedResult(
        title=title,
        summary=summary,
        changed_files=changed_files,
        created_files=created_files,
        test_results=test_results,
        build_results=build_results,
        commit_hash=commit_hash,
        push_result=push_result,
        git_status=git_status,
        phase_notes=phase_notes,
        raw_text=text,
        content_hash=content_hash,
        imported_at=imported_at,
    )


COMMIT_PATTERNS = (
    re.compile(r"(?:commit|commit hash|コミット)\s*:?\s*`?([0-9a-f]{7,40})`?", re.IGNORECASE),
    re.compile(r"\b([0-9a-f]{7,40})\b"),
)
FILE_PATTERN = re.compile(
    r"(?:(?:[A-Za-z]:)?[\\/])?[\w .()@+\-\\/]+?\.(?:py|md|bat|txt|json|yml|yaml|toml|ini|csv|tsv|html|css|js|ts|tsx|jsx|webp|png|ico)",
    re.IGNORECASE,
)
TEST_PATTERNS = (
    re.compile(r"test|smoke|launch-check|確認|検証|OK|成功", re.IGNORECASE),
)
BUILD_PATTERNS = (
    re.compile(r"build|build\.bat|pyinstaller|dist|exe", re.IGNORECASE),
)
PUSH_PATTERNS = (
    re.compile(r"push|origin/main|origin main", re.IGNORECASE),
)
GIT_STATUS_PATTERNS = (
    re.compile(r"git status|status clean|clean", re.IGNORECASE),
)
PHASE_PATTERNS = (
    re.compile(r"phase|候補|未実装|次", re.IGNORECASE),
)


def extract_commit_hash(text: str) -> str:
    for pattern in COMMIT_PATTERNS:
        match = pattern.search(text)
        if match:
            return match.group(1)
    return ""


def extract_file_paths(text: str) -> list[str]:
    matches = []
    for match in FILE_PATTERN.finditer(text):
        value = match.group(0).strip("`'\"<>[]()").strip().lstrip("- ").strip()
        value = value.replace("\\", "/")
        if value.lower().startswith(("python ", "git ", "cmd ")):
            continue
        if value and value not in matches:
            matches.append(value)
    return matches


def extract_changed_files(lines: list[str], raw_text: str) -> list[str]:
    changed = []
    markers = ("追加", "更新", "変更", "修正", "作成", "changed", "updated", "modified", "added", "created")
    for line in lines:
        if not any(marker in line.lower() for marker in markers):
            continue
        for path in extract_file_paths(line):
            if path not in changed:
                changed.append(path)
    return changed or extract_file_paths(raw_text)


def extract_created_files(lines: list[str], fallback_files: list[str]) -> list[str]:
    created = []
    for line in lines:
        lower = line.lower()
        if not any(token in lower for token in ("add", "added", "create", "created", "追加", "作成")):
            continue
        for path in extract_file_paths(line):
            if path not in created:
                created.append(path)
    return created or fallback_files[:1]


def extract_matching_lines(lines: list[str], patterns: tuple[re.Pattern[str], ...], limit: int = 12) -> str:
    matched = []
    for line in lines:
        if not line:
            continue
        if any(pattern.search(line) for pattern in patterns):
            matched.append(line)
        if len(matched) >= limit:
            break
    return "\n".join(matched)


def extract_first_matching_line(lines: list[str], patterns: tuple[re.Pattern[str], ...]) -> str:
    for line in lines:
        if line and any(pattern.search(line) for pattern in patterns):
            return line
    return ""


def extract_title(lines: list[str], commit_hash: str) -> str:
    for line in lines:
        if line.startswith("#"):
            clean = line.strip("# ").strip()
            if clean:
                return clean[:90]
    if commit_hash:
        commit_line_pattern = re.compile(rf"{re.escape(commit_hash)}[`']?\s+(.+)$", re.IGNORECASE)
        for line in lines:
            match = commit_line_pattern.search(line)
            if match:
                clean = match.group(1).strip(" -`")
                if clean:
                    return clean[:90]
    for line in lines:
        clean = line.strip("# -*")
        if clean and len(clean) <= 90 and not clean.lower().startswith(("commit", "push", "git status")):
            return clean
    if commit_hash:
        return f"Codex result {commit_hash[:7]}"
    return "Codex result"


def extract_summary(lines: list[str]) -> str:
    meaningful = []
    for line in lines:
        clean = line.strip("- ")
        if not clean or clean.startswith("::"):
            continue
        meaningful.append(clean)
        if len(meaningful) >= 5:
            break
    return "\n".join(meaningful)


def build_document_record(parsed: CodexParsedResult, source_label: str = "") -> DocumentRecord:
    key = parsed.commit_hash or parsed.content_hash
    content = "\n".join(
        [
            f"Title: {parsed.title}",
            f"Summary: {parsed.summary}",
            f"Source Type: {SOURCE_TYPE_CODEX}",
            f"Commit Hash: {parsed.commit_hash}",
            f"Changed Files: {json.dumps(parsed.changed_files, ensure_ascii=False)}",
            f"Created Files: {json.dumps(parsed.created_files, ensure_ascii=False)}",
            f"Test Results: {parsed.test_results}",
            f"Build Results: {parsed.build_results}",
            f"Push Result: {parsed.push_result}",
            f"Git Status: {parsed.git_status}",
            f"Phase Notes: {parsed.phase_notes}",
            "",
            parsed.raw_text,
        ]
    )
    label = source_label or "Codex result paste"
    return DocumentRecord(
        path=f"codex_result://{key}",
        title=parsed.title,
        source_type=SOURCE_TYPE_CODEX,
        source_label=label,
        conversation_id=parsed.commit_hash,
        conversation_title=parsed.title,
        role="codex",
        message_index=0,
        source_created_at=parsed.imported_at,
        source_updated_at=parsed.imported_at,
        codex_summary=parsed.summary,
        changed_files_json=json.dumps(parsed.changed_files, ensure_ascii=False),
        created_files_json=json.dumps(parsed.created_files, ensure_ascii=False),
        test_results=parsed.test_results,
        build_results=parsed.build_results,
        commit_hash=parsed.commit_hash,
        push_result=parsed.push_result,
        git_status=parsed.git_status,
        phase_notes=parsed.phase_notes,
        created_at=parsed.imported_at,
        modified_at=parsed.imported_at,
        indexed_at=now_iso(),
        hash=parsed.content_hash,
        content=content,
    )


def write_import_log(result: CodexImportResult) -> Path:
    logs_dir().mkdir(parents=True, exist_ok=True)
    path = logs_dir() / f"codex_import_{now_iso().replace(':', '').replace('-', '').replace('T', '_')}.log"
    lines = [
        "Codex result detected.",
        f"Commit hash found: {result.commit_hash}",
        f"Changed files: {result.changed_files_count}",
        f"Skipped duplicate: {result.skipped_duplicate}",
        f"Content hash: {result.content_hash}",
    ]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return path
