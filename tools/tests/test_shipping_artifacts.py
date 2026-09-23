from __future__ import annotations

import sys
import zipfile
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from shipping_artifacts import find_exe, is_onedir, verify_runtime_zip, write_verified_zip


@pytest.mark.parametrize("onedir", [False, True])
def test_distribution_zip_roundtrip(tmp_path: Path, onedir: bool) -> None:
    app = tmp_path / "app"
    dist = app / "dist"
    exe = dist / "Sample" / "Sample.exe" if onedir else dist / "Sample.exe"
    exe.parent.mkdir(parents=True)
    exe.write_bytes(b"synthetic executable")
    if onedir:
        dependency = exe.parent / "_internal" / "dependency.dll"
        dependency.parent.mkdir()
        dependency.write_bytes(b"synthetic dependency")
    readme = app / "README.txt"
    notice = app / "notice.txt"
    readme.write_text("README", encoding="utf-8")
    notice.write_text("NOTICE", encoding="utf-8")

    assert find_exe(app, "Sample.exe") == exe
    assert is_onedir(exe) is onedir
    archive_path = app / "booth_ready" / "Sample.zip"
    write_verified_zip(archive_path, exe, readme, notice)
    verify_runtime_zip(archive_path, exe)

    extracted = tmp_path / "extracted"
    with zipfile.ZipFile(archive_path) as archive:
        archive.extractall(extracted)
    runtime_exe = extracted / "Sample" / "Sample.exe" if onedir else extracted / "Sample.exe"
    assert runtime_exe.read_bytes() == exe.read_bytes()
    if onedir:
        assert (extracted / "Sample" / "_internal" / "dependency.dll").read_bytes() == b"synthetic dependency"


def test_onedir_zip_rejects_missing_runtime_file(tmp_path: Path) -> None:
    exe = tmp_path / "dist" / "Sample" / "Sample.exe"
    exe.parent.mkdir(parents=True)
    exe.write_bytes(b"exe")
    dependency = exe.parent / "_internal" / "dependency.dll"
    dependency.parent.mkdir()
    dependency.write_bytes(b"dll")
    archive_path = tmp_path / "incomplete.zip"
    with zipfile.ZipFile(archive_path, "w") as archive:
        archive.write(exe, "Sample/Sample.exe")
    with pytest.raises(ValueError, match="dependency.dll"):
        verify_runtime_zip(archive_path, exe)
