# -*- coding: utf-8 -*-
from __future__ import annotations

import os
from pathlib import Path

import pytest

from rename_core import (
    FileSnapshot,
    RenameRequest,
    RenameTransactionError,
    RenameValidationError,
    build_rename_plan,
    build_undo_plan,
    normalize_requested_stem,
    rename_batch,
    undo_rename,
    validate_windows_stem,
)


def make_pdf(folder: Path, name: str, payload: bytes | None = None) -> Path:
    path = folder / name
    path.write_bytes(payload or b"%PDF-1.4\nsynthetic\n%%EOF\n")
    return path


def requests_for(folder: Path, names: list[str], stems: list[str]) -> list[RenameRequest]:
    return [
        RenameRequest(FileSnapshot.capture(folder / name), stem)
        for name, stem in zip(names, stems)
    ]


def issue_codes(exc: RenameValidationError) -> set[str]:
    return {issue.code for issue in exc.issues}


def test_normal_multiple_and_japanese_names(tmp_path: Path) -> None:
    names = ["001.pdf", "002.pdf", "003.pdf"]
    for name in names:
        make_pdf(tmp_path, name)
    _, undo = rename_batch(requests_for(tmp_path, names, ["契約書（原本）", "重要事項説明書", "間取り＿１階"]))
    assert {path.name for path in tmp_path.iterdir()} == {
        "契約書（原本）.pdf", "重要事項説明書.pdf", "間取り＿１階.pdf"
    }
    assert len(undo.entries) == 3


def test_two_file_swap(tmp_path: Path) -> None:
    make_pdf(tmp_path, "A.pdf", b"A")
    make_pdf(tmp_path, "B.pdf", b"B")
    rename_batch(requests_for(tmp_path, ["A.pdf", "B.pdf"], ["B", "A"]))
    assert (tmp_path / "A.pdf").read_bytes() == b"B"
    assert (tmp_path / "B.pdf").read_bytes() == b"A"


def test_three_file_cycle(tmp_path: Path) -> None:
    for name in ("A", "B", "C"):
        make_pdf(tmp_path, f"{name}.pdf", name.encode())
    rename_batch(requests_for(tmp_path, ["A.pdf", "B.pdf", "C.pdf"], ["B", "C", "A"]))
    assert [(tmp_path / f"{name}.pdf").read_bytes() for name in ("A", "B", "C")] == [b"C", b"A", b"B"]


def test_case_only_change(tmp_path: Path) -> None:
    make_pdf(tmp_path, "report.pdf")
    rename_batch(requests_for(tmp_path, ["report.pdf"], ["REPORT"]))
    assert (tmp_path / "REPORT.pdf").exists()


@pytest.mark.parametrize(
    ("value", "code"),
    [
        ("", "empty"),
        ("bad:name", "forbidden_character"),
        ("bad\x01name", "control_character"),
        ("bad.", "trailing_dot"),
        ("bad ", "trailing_space"),
        ("CON", "reserved_name"),
        ("CON.txt", "reserved_name"),
        ("LPT1.memo", "reserved_name"),
        ("a" * 252, "too_long"),
        ("😀" * 126, "too_long"),
    ],
)
def test_invalid_windows_names(value: str, code: str) -> None:
    assert code in validate_windows_stem(value)


def test_explicit_pdf_suffix_is_normalized_once() -> None:
    assert normalize_requested_stem("契約書.PDF") == "契約書"
    assert normalize_requested_stem("契約書.pdf.pdf") == "契約書.pdf"
    assert normalize_requested_stem("契約書.pdf ") == "契約書.pdf "


def test_case_insensitive_duplicate(tmp_path: Path) -> None:
    make_pdf(tmp_path, "A.pdf")
    make_pdf(tmp_path, "B.pdf")
    with pytest.raises(RenameValidationError) as caught:
        build_rename_plan(requests_for(tmp_path, ["A.pdf", "B.pdf"], ["same", "SAME"]))
    assert "duplicate_destination" in issue_codes(caught.value)


def test_collision_with_non_target_file(tmp_path: Path) -> None:
    make_pdf(tmp_path, "A.pdf")
    make_pdf(tmp_path, "occupied.pdf")
    with pytest.raises(RenameValidationError) as caught:
        build_rename_plan(requests_for(tmp_path, ["A.pdf"], ["OCCUPIED"]))
    assert "destination_exists" in issue_codes(caught.value)


def test_missing_source(tmp_path: Path) -> None:
    path = make_pdf(tmp_path, "A.pdf")
    snapshot = FileSnapshot.capture(path)
    path.unlink()
    with pytest.raises(RenameValidationError) as caught:
        build_rename_plan([RenameRequest(snapshot, "B")])
    assert "source_missing" in issue_codes(caught.value)


def test_externally_changed_source(tmp_path: Path) -> None:
    path = make_pdf(tmp_path, "A.pdf")
    snapshot = FileSnapshot.capture(path)
    path.write_bytes(b"changed payload with another size")
    with pytest.raises(RenameValidationError) as caught:
        build_rename_plan([RenameRequest(snapshot, "B")])
    assert "source_changed" in issue_codes(caught.value)


def test_injected_failure_rolls_back_every_file(tmp_path: Path) -> None:
    names = ["A.pdf", "B.pdf", "C.pdf"]
    for name in names:
        make_pdf(tmp_path, name, name.encode())
    calls = 0

    def fail_once(source: Path, destination: Path) -> None:
        nonlocal calls
        calls += 1
        if calls == 5:
            raise OSError("injected failure")
        os.rename(source, destination)

    with pytest.raises(RenameTransactionError) as caught:
        rename_batch(requests_for(tmp_path, names, ["B", "C", "A"]), move=fail_once)
    assert not caught.value.rollback_errors
    assert {path.name for path in tmp_path.iterdir()} == set(names)
    assert all((tmp_path / name).read_bytes() == name.encode() for name in names)
    assert not any(path.name.startswith(".__dake_overview_") for path in tmp_path.iterdir())


def test_undo_success(tmp_path: Path) -> None:
    make_pdf(tmp_path, "A.pdf", b"A")
    make_pdf(tmp_path, "B.pdf", b"B")
    _, undo = rename_batch(requests_for(tmp_path, ["A.pdf", "B.pdf"], ["B", "A"]))
    undo_rename(undo)
    assert (tmp_path / "A.pdf").read_bytes() == b"A"
    assert (tmp_path / "B.pdf").read_bytes() == b"B"


def test_undo_collision_aborts_before_writes(tmp_path: Path) -> None:
    make_pdf(tmp_path, "A.pdf", b"A")
    _, undo = rename_batch(requests_for(tmp_path, ["A.pdf"], ["B"]))
    make_pdf(tmp_path, "A.pdf", b"external")
    before = {path.name: path.read_bytes() for path in tmp_path.iterdir()}
    with pytest.raises(RenameValidationError) as caught:
        build_undo_plan(undo)
    assert "destination_exists" in issue_codes(caught.value)
    assert {path.name: path.read_bytes() for path in tmp_path.iterdir()} == before
