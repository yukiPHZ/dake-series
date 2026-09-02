# -*- coding: utf-8 -*-
"""Safe, GUI-independent file-name validation and transactional renaming."""

from __future__ import annotations

import os
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable


PDF_SUFFIX = ".pdf"
WINDOWS_FORBIDDEN = frozenset('<>:"/\\|?*')
WINDOWS_RESERVED = frozenset(
    {"CON", "PRN", "AUX", "NUL"}
    | {f"COM{number}" for number in range(1, 10)}
    | {f"LPT{number}" for number in range(1, 10)}
)
MAX_COMPONENT_LENGTH = 255


@dataclass(frozen=True)
class FileSnapshot:
    path: Path
    size: int
    mtime_ns: int
    ctime_ns: int
    device: int
    inode: int

    @classmethod
    def capture(cls, path: Path | str) -> "FileSnapshot":
        resolved = Path(path).resolve()
        stat = resolved.stat()
        return cls(
            path=resolved,
            size=stat.st_size,
            mtime_ns=stat.st_mtime_ns,
            ctime_ns=stat.st_ctime_ns,
            device=stat.st_dev,
            inode=stat.st_ino,
        )

    def matches_disk(self) -> bool:
        try:
            current = self.path.stat()
        except OSError:
            return False
        return (
            current.st_size == self.size
            and current.st_mtime_ns == self.mtime_ns
            and current.st_ctime_ns == self.ctime_ns
            and current.st_dev == self.device
            and current.st_ino == self.inode
        )


@dataclass(frozen=True)
class RenameRequest:
    snapshot: FileSnapshot
    requested_stem: str


@dataclass(frozen=True)
class ValidationIssue:
    code: str
    source_name: str
    detail: str = ""


class RenameValidationError(ValueError):
    def __init__(self, issues: Iterable[ValidationIssue]):
        self.issues = tuple(issues)
        summary = "; ".join(
            f"{issue.source_name}: {issue.code}{(': ' + issue.detail) if issue.detail else ''}"
            for issue in self.issues
        )
        super().__init__(summary)


class RenameTransactionError(RuntimeError):
    def __init__(self, cause: BaseException, rollback_errors: Iterable[str] = ()):
        self.cause = cause
        self.rollback_errors = tuple(rollback_errors)
        message = f"rename failed: {cause}"
        if self.rollback_errors:
            message += "; rollback errors: " + "; ".join(self.rollback_errors)
        super().__init__(message)


@dataclass(frozen=True)
class RenameEntry:
    snapshot: FileSnapshot
    destination: Path

    @property
    def source(self) -> Path:
        return self.snapshot.path


@dataclass(frozen=True)
class RenamePlan:
    folder: Path
    entries: tuple[RenameEntry, ...]


@dataclass(frozen=True)
class UndoEntry:
    original_path: Path
    renamed_snapshot: FileSnapshot


@dataclass(frozen=True)
class UndoRecord:
    folder: Path
    entries: tuple[UndoEntry, ...]


MoveFunction = Callable[[Path, Path], None]


def normalize_requested_stem(value: str) -> str:
    """Remove exactly one explicitly entered .pdf suffix; never trim whitespace/dots."""
    if value.casefold().endswith(PDF_SUFFIX):
        return value[: -len(PDF_SUFFIX)]
    return value


def validate_windows_stem(stem: str) -> tuple[str, ...]:
    errors: list[str] = []
    if stem == "":
        errors.append("empty")
        return tuple(errors)
    if any(character in WINDOWS_FORBIDDEN for character in stem):
        errors.append("forbidden_character")
    if any(ord(character) < 32 or ord(character) == 127 for character in stem):
        errors.append("control_character")
    if stem.endswith("."):
        errors.append("trailing_dot")
    if stem.endswith(" "):
        errors.append("trailing_space")
    reserved_base = stem.split(".", 1)[0].upper()
    if reserved_base in WINDOWS_RESERVED:
        errors.append("reserved_name")
    utf16_units = len((stem + PDF_SUFFIX).encode("utf-16-le")) // 2
    if utf16_units > MAX_COMPONENT_LENGTH:
        errors.append("too_long")
    return tuple(errors)


def _folder_entries_by_casefold(folder: Path) -> dict[str, list[Path]]:
    entries: dict[str, list[Path]] = {}
    try:
        children = tuple(folder.iterdir())
    except OSError as exc:
        raise RenameValidationError((ValidationIssue("folder_unreadable", folder.name, str(exc)),)) from exc
    for child in children:
        entries.setdefault(child.name.casefold(), []).append(child.resolve())
    return entries


def build_rename_plan(requests: Iterable[RenameRequest]) -> RenamePlan:
    request_list = tuple(requests)
    if not request_list:
        raise RenameValidationError((ValidationIssue("no_files", ""),))

    folder = request_list[0].snapshot.path.parent.resolve()
    issues: list[ValidationIssue] = []
    destinations: list[Path] = []
    destination_owners: dict[str, list[str]] = {}

    for request in request_list:
        source = request.snapshot.path.resolve()
        if source.parent != folder or source.suffix.casefold() != PDF_SUFFIX:
            issues.append(ValidationIssue("invalid_source", source.name))
        if not source.exists():
            issues.append(ValidationIssue("source_missing", source.name))
        elif not request.snapshot.matches_disk():
            issues.append(ValidationIssue("source_changed", source.name))

        stem = normalize_requested_stem(request.requested_stem)
        for code in validate_windows_stem(stem):
            issues.append(ValidationIssue(code, source.name, request.requested_stem))
        destination = folder / f"{stem}{PDF_SUFFIX}"
        destinations.append(destination)
        destination_owners.setdefault(destination.name.casefold(), []).append(source.name)

    for folded_name, owners in destination_owners.items():
        if len(owners) > 1:
            detail = ", ".join(owners)
            for owner in owners:
                issues.append(ValidationIssue("duplicate_destination", owner, detail))

    changed_source_names = {
        request.snapshot.path.name.casefold()
        for request, destination in zip(request_list, destinations)
        if request.snapshot.path.name != destination.name
    }
    disk_entries = _folder_entries_by_casefold(folder)
    for request, destination in zip(request_list, destinations):
        source = request.snapshot.path
        if source.name == destination.name:
            continue
        for existing in disk_entries.get(destination.name.casefold(), ()):  # case-insensitive Windows view
            if existing.name.casefold() not in changed_source_names:
                issues.append(ValidationIssue("destination_exists", source.name, existing.name))

    if issues:
        raise RenameValidationError(issues)

    entries = tuple(
        # Do not resolve the not-yet-created destination on Windows: resolve()
        # canonicalizes it to an existing differently-cased source path and
        # would erase a case-only rename such as report.pdf -> REPORT.pdf.
        RenameEntry(request.snapshot, destination)
        for request, destination in zip(request_list, destinations)
        if request.snapshot.path.name != destination.name
    )
    return RenamePlan(folder=folder, entries=entries)


def _unique_temp_path(folder: Path, occupied: set[str], marker: str) -> Path:
    while True:
        candidate = folder / f".__dake_overview_{marker}_{uuid.uuid4().hex}.tmp"
        folded = candidate.name.casefold()
        if folded not in occupied and not candidate.exists():
            occupied.add(folded)
            return candidate


def _default_move(source: Path, destination: Path) -> None:
    os.rename(source, destination)


def _casefold_path_exists(path: Path) -> bool:
    folded = path.name.casefold()
    return any(child.name.casefold() == folded for child in path.parent.iterdir())


def _preflight_plan(plan: RenamePlan) -> None:
    requests = tuple(
        RenameRequest(entry.snapshot, entry.destination.stem)
        for entry in plan.entries
    )
    if not requests:
        return
    refreshed = build_rename_plan(requests)
    if tuple((entry.source, entry.destination) for entry in refreshed.entries) != tuple(
        (entry.source, entry.destination) for entry in plan.entries
    ):
        raise RenameValidationError((ValidationIssue("plan_changed", plan.folder.name),))


def execute_rename_plan(plan: RenamePlan, move: MoveFunction | None = None) -> UndoRecord:
    if not plan.entries:
        return UndoRecord(plan.folder, ())
    _preflight_plan(plan)
    mover = move or _default_move
    occupied = {path.name.casefold() for path in plan.folder.iterdir()}
    temporary = {
        entry.source: _unique_temp_path(plan.folder, occupied, "stage")
        for entry in plan.entries
    }
    current = {entry.source: entry.source for entry in plan.entries}

    try:
        for entry in plan.entries:
            temp = temporary[entry.source]
            mover(entry.source, temp)
            current[entry.source] = temp
        for entry in plan.entries:
            temp = current[entry.source]
            if _casefold_path_exists(entry.destination):
                raise FileExistsError(entry.destination)
            mover(temp, entry.destination)
            current[entry.source] = entry.destination
    except BaseException as exc:
        rollback_errors: list[str] = []
        recovery: dict[Path, Path] = {}
        for entry in plan.entries:
            location = current[entry.source]
            if location == entry.source:
                continue
            recovery_path = _unique_temp_path(plan.folder, occupied, "rollback")
            try:
                mover(location, recovery_path)
                recovery[entry.source] = recovery_path
            except BaseException as rollback_exc:
                rollback_errors.append(f"{location.name}: {rollback_exc}")
        for entry in plan.entries:
            recovery_path = recovery.get(entry.source)
            if recovery_path is None:
                continue
            try:
                if entry.source.exists():
                    raise FileExistsError(entry.source)
                mover(recovery_path, entry.source)
            except BaseException as rollback_exc:
                rollback_errors.append(f"{entry.source.name}: {rollback_exc}")
        raise RenameTransactionError(exc, rollback_errors) from exc

    undo_entries = tuple(
        UndoEntry(entry.source, FileSnapshot.capture(entry.destination))
        for entry in plan.entries
    )
    return UndoRecord(plan.folder, undo_entries)


def rename_batch(
    requests: Iterable[RenameRequest], move: MoveFunction | None = None
) -> tuple[RenamePlan, UndoRecord]:
    plan = build_rename_plan(requests)
    return plan, execute_rename_plan(plan, move=move)


def build_undo_plan(record: UndoRecord) -> RenamePlan:
    if not record.entries:
        raise RenameValidationError((ValidationIssue("nothing_to_undo", ""),))
    requests = tuple(
        RenameRequest(entry.renamed_snapshot, entry.original_path.stem)
        for entry in record.entries
    )
    return build_rename_plan(requests)


def undo_rename(record: UndoRecord, move: MoveFunction | None = None) -> RenamePlan:
    plan = build_undo_plan(record)
    execute_rename_plan(plan, move=move)
    return plan
