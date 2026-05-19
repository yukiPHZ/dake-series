# -*- coding: utf-8 -*-
from __future__ import annotations

import hashlib
import shutil
from dataclasses import dataclass, field, replace
from pathlib import Path

from .logger import BackupLogger, make_timestamp


ARCHIVE_DIR_NAME = "backup_archive"
CHUNK_SIZE = 1024 * 1024


class BackupError(RuntimeError):
    pass


@dataclass(frozen=True)
class BackupItem:
    relative_path: str
    source_path: str
    destination_path: str
    action: str
    size: int = 0
    archive_path: str = ""


@dataclass
class DiffResult:
    source_folder: str
    destination_folder: str
    added: list[BackupItem] = field(default_factory=list)
    updated: list[BackupItem] = field(default_factory=list)
    unchanged: int = 0
    preserved_destination_only: int = 0

    @property
    def archive_planned(self) -> int:
        return len(self.updated)

    @property
    def delete_planned(self) -> int:
        return 0

    @property
    def copy_planned(self) -> int:
        return len(self.added) + len(self.updated)

    def summary(self) -> dict[str, int]:
        return {
            "added": len(self.added),
            "updated": len(self.updated),
            "archive": self.archive_planned,
            "delete": self.delete_planned,
            "unchanged": self.unchanged,
            "preserved_destination_only": self.preserved_destination_only,
        }


@dataclass
class BackupResult:
    diff: DiffResult
    copied: list[BackupItem]
    archived: list[BackupItem]
    timestamp: str
    log_path: str


def _is_relative_to(child: Path, parent: Path) -> bool:
    try:
        child.relative_to(parent)
    except ValueError:
        return False
    return True


def _normalize_folder(path_text: str, label: str) -> Path:
    text = str(path_text or "").strip()
    if not text:
        raise BackupError(f"{label}を選択してください。")
    return Path(text).expanduser()


def validate_folders(source_folder: str, destination_folder: str) -> tuple[Path, Path]:
    source = _normalize_folder(source_folder, "正本フォルダ")
    destination = _normalize_folder(destination_folder, "避難先フォルダ")
    if not source.exists() or not source.is_dir():
        raise BackupError("正本フォルダが見つかりません。")

    source_resolved = source.resolve(strict=True)
    destination_resolved = destination.resolve(strict=False)
    if source_resolved == destination_resolved:
        raise BackupError("正本フォルダと避難先フォルダは同じ場所にできません。")
    if _is_relative_to(destination_resolved, source_resolved) or _is_relative_to(
        source_resolved, destination_resolved
    ):
        raise BackupError("正本フォルダと避難先フォルダは親子関係にしないでください。")
    return source, destination


def hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(CHUNK_SIZE), b""):
            digest.update(chunk)
    return digest.hexdigest()


def files_match(source: Path, destination: Path) -> bool:
    if not destination.exists() or not destination.is_file():
        return False
    if source.stat().st_size != destination.stat().st_size:
        return False
    return hash_file(source) == hash_file(destination)


def iter_files(root: Path) -> list[Path]:
    return sorted(
        (path for path in root.rglob("*") if path.is_file()),
        key=lambda path: path.relative_to(root).as_posix().lower(),
    )


def count_destination_only(source: Path, destination: Path) -> int:
    if not destination.exists():
        return 0
    count = 0
    for path in iter_files(destination):
        try:
            rel = path.relative_to(destination)
        except ValueError:
            continue
        if rel.parts and rel.parts[0] == ARCHIVE_DIR_NAME:
            continue
        if not (source / rel).exists():
            count += 1
    return count


def scan_diff(source_folder: str, destination_folder: str) -> DiffResult:
    source, destination = validate_folders(source_folder, destination_folder)
    diff = DiffResult(str(source), str(destination))

    for source_path in iter_files(source):
        rel_path = source_path.relative_to(source)
        rel_text = rel_path.as_posix()
        destination_path = destination / rel_path
        if destination_path.exists() and destination_path.is_dir():
            raise BackupError(f"避難先に同名フォルダがあります: {rel_text}")
        item = BackupItem(
            relative_path=rel_text,
            source_path=str(source_path),
            destination_path=str(destination_path),
            action="copy",
            size=source_path.stat().st_size,
        )
        if not destination_path.exists():
            diff.added.append(replace(item, action="add"))
        elif files_match(source_path, destination_path):
            diff.unchanged += 1
        else:
            diff.updated.append(replace(item, action="update"))

    diff.preserved_destination_only = count_destination_only(source, destination)
    return diff


def _copy_file(source_path: Path, destination_path: Path) -> None:
    if destination_path.parent.exists() and not destination_path.parent.is_dir():
        raise BackupError(f"避難先の親パスがフォルダではありません: {destination_path.parent}")
    destination_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source_path, destination_path)


def execute_backup(
    source_folder: str,
    destination_folder: str,
    diff: DiffResult | None = None,
    logger: BackupLogger | None = None,
    timestamp: str | None = None,
) -> BackupResult:
    source, destination = validate_folders(source_folder, destination_folder)
    destination.mkdir(parents=True, exist_ok=True)
    run_timestamp = timestamp or make_timestamp()
    backup_logger = logger or BackupLogger(run_timestamp)
    current_diff = diff or scan_diff(str(source), str(destination))
    archive_root = destination / ARCHIVE_DIR_NAME / run_timestamp
    copied: list[BackupItem] = []
    archived: list[BackupItem] = []

    backup_logger.write("DAKE Backup START")
    backup_logger.write(f"SOURCE: {source}")
    backup_logger.write(f"DESTINATION: {destination}")
    backup_logger.write(
        "DIFF: "
        f"added={len(current_diff.added)} "
        f"updated={len(current_diff.updated)} "
        f"archive={current_diff.archive_planned} "
        "delete_planned=0"
    )

    for item in current_diff.updated:
        destination_path = Path(item.destination_path)
        archive_path = archive_root / Path(item.relative_path)
        archive_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(destination_path, archive_path)
        archived_item = replace(item, archive_path=str(archive_path))
        archived.append(archived_item)
        backup_logger.write(f"ARCHIVE: {item.relative_path} -> {archive_path}")

    for item in [*current_diff.added, *current_diff.updated]:
        _copy_file(Path(item.source_path), Path(item.destination_path))
        copied.append(item)
        backup_logger.write(f"COPY: {item.relative_path}")

    backup_logger.write(
        f"DAKE Backup END copied={len(copied)} archived={len(archived)} delete_planned=0"
    )
    return BackupResult(
        diff=current_diff,
        copied=copied,
        archived=archived,
        timestamp=run_timestamp,
        log_path=str(backup_logger.path),
    )

