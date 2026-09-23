"""Locate and package onefile and onedir Windows distribution artifacts."""

from __future__ import annotations

import hashlib
import os
import tempfile
import zipfile
from pathlib import Path


def find_exe(app_dir: Path, exe_name: str = "") -> Path | None:
    dist = app_dir / "dist"
    if not dist.is_dir():
        return None
    if exe_name:
        direct = dist / exe_name
        if direct.is_file():
            return direct
        nested = sorted(path for path in dist.glob(f"*/{exe_name}") if path.is_file())
        if nested:
            preferred = next((path for path in nested if path.parent.name == path.stem), None)
            return preferred or nested[0]
        return None
    candidates = sorted(path for path in dist.glob("*.exe") if path.is_file())
    candidates.extend(sorted(path for path in dist.glob("*/*.exe") if path.is_file()))
    return candidates[0] if candidates else None


def is_onedir(exe_path: Path) -> bool:
    return exe_path.parent.parent.name.lower() == "dist"


def runtime_files(exe_path: Path) -> list[tuple[Path, str]]:
    if not exe_path.is_file():
        raise FileNotFoundError(exe_path)
    if not is_onedir(exe_path):
        return [(exe_path, exe_path.name)]
    root = exe_path.parent
    files = sorted(path for path in root.rglob("*") if path.is_file())
    if any(path.is_symlink() for path in root.rglob("*")):
        raise ValueError(f"Runtime contains a symlink: {root}")
    return [(path, (Path(root.name) / path.relative_to(root)).as_posix()) for path in files]


def write_verified_zip(zip_path: Path, exe_path: Path, readme: Path, notice: Path) -> None:
    payload = runtime_files(exe_path)
    expected = payload + [(readme, "README.txt"), (notice, "注意事項.txt")]
    names = [name for _, name in expected]
    if len(names) != len(set(names)):
        raise ValueError("Duplicate archive paths")
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(prefix=zip_path.stem + "-", suffix=".zip", dir=zip_path.parent, delete=False) as handle:
        temporary = Path(handle.name)
    try:
        with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for source, name in expected:
                archive.write(source, arcname=name)
        verify_zip(temporary, expected)
        os.replace(temporary, zip_path)
    finally:
        temporary.unlink(missing_ok=True)


def verify_zip(zip_path: Path, expected: list[tuple[Path, str]]) -> None:
    expected_names = {name for _, name in expected}
    with zipfile.ZipFile(zip_path) as archive:
        if set(archive.namelist()) != expected_names or len(archive.namelist()) != len(expected_names):
            raise ValueError("Archive file list differs from distribution payload")
        for source, name in expected:
            with source.open("rb") as original:
                source_hash = hashlib.file_digest(original, "sha256").digest()
            with archive.open(name) as member:
                archive_hash = hashlib.file_digest(member, "sha256").digest()
            if source_hash != archive_hash:
                raise ValueError(f"Archive content mismatch: {name}")


def verify_runtime_zip(zip_path: Path, exe_path: Path) -> None:
    payload = runtime_files(exe_path)
    with zipfile.ZipFile(zip_path) as archive:
        names = set(archive.namelist())
        for source, name in payload:
            if name not in names:
                raise ValueError(f"Runtime file missing from archive: {name}")
            with source.open("rb") as original, archive.open(name) as member:
                if hashlib.file_digest(original, "sha256").digest() != hashlib.file_digest(member, "sha256").digest():
                    raise ValueError(f"Runtime file differs from archive: {name}")
