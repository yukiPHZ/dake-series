# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import os
import queue
import struct
import tempfile
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pythoncom
import win32api
import win32clipboard
import win32con
import win32com.server.util
from win32com import storagecon
from win32com.shell import shellcon


FILE_GROUP_DESCRIPTOR_W = win32clipboard.RegisterClipboardFormat("FileGroupDescriptorW")
FILE_CONTENTS = win32clipboard.RegisterClipboardFormat("FileContents")
CONTENT_TYMEDS = (
    pythoncom.TYMED_HGLOBAL,
    pythoncom.TYMED_ISTREAM,
    pythoncom.TYMED_ISTORAGE,
)
CONTENT_TYMED_MASK = (
    pythoncom.TYMED_HGLOBAL
    | pythoncom.TYMED_ISTREAM
    | pythoncom.TYMED_ISTORAGE
)
STREAM_CHUNK_SIZE = 1024 * 1024
MAX_DESCRIPTOR_COUNT = 1000
APP_DATA_DIR = Path(
    os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")
) / "PEAKHEADZ" / "DakeMailList"
DIAGNOSTIC_LOG_PATH = APP_DATA_DIR / "logs" / "ole_dnd.log"


@dataclass(frozen=True)
class OleDropPayload:
    paths: list[str]
    cleanup: object | None = None
    error: bool = False


class _GUID(ctypes.Structure):
    _pack_ = 4
    _fields_ = [
        ("Data1", ctypes.c_uint32),
        ("Data2", ctypes.c_uint16),
        ("Data3", ctypes.c_uint16),
        ("Data4", ctypes.c_ubyte * 8),
    ]


class _SIZEL(ctypes.Structure):
    _pack_ = 4
    _fields_ = [("cx", ctypes.c_int32), ("cy", ctypes.c_int32)]


class _POINTL(ctypes.Structure):
    _pack_ = 4
    _fields_ = [("x", ctypes.c_int32), ("y", ctypes.c_int32)]


class _FILETIME(ctypes.Structure):
    _pack_ = 4
    _fields_ = [("dwLowDateTime", ctypes.c_uint32), ("dwHighDateTime", ctypes.c_uint32)]


class _FILEDESCRIPTORW(ctypes.Structure):
    _pack_ = 4
    _fields_ = [
        ("dwFlags", ctypes.c_uint32),
        ("clsid", _GUID),
        ("sizel", _SIZEL),
        ("pointl", _POINTL),
        ("dwFileAttributes", ctypes.c_uint32),
        ("ftCreationTime", _FILETIME),
        ("ftLastAccessTime", _FILETIME),
        ("ftLastWriteTime", _FILETIME),
        ("nFileSizeHigh", ctypes.c_uint32),
        ("nFileSizeLow", ctypes.c_uint32),
        ("cFileName", ctypes.c_wchar * 260),
    ]


class OleDndDiagnostics:
    def __init__(self, path: Path = DIAGNOSTIC_LOG_PATH) -> None:
        self.path = path
        self._lock = threading.Lock()
        self._reset()

    def _reset(self) -> None:
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            self.path.write_text("", encoding="utf-8")
        except OSError:
            return
        self.record("ole_init")

    def record(self, event: str, **values: object) -> None:
        timestamp = datetime.now().isoformat(timespec="seconds")
        details = " ".join(f"{key}={value}" for key, value in values.items())
        line = f"{timestamp} {event}"
        if details:
            line += f" {details}"
        try:
            with self._lock:
                with self.path.open("a", encoding="utf-8", newline="\n") as file:
                    file.write(line + "\n")
        except OSError:
            return


def _hresult(error: BaseException) -> str:
    value = getattr(error, "hresult", None)
    if not isinstance(value, int):
        return "unknown"
    return f"0x{value & 0xFFFFFFFF:08X}"


def _exception_type(error: BaseException) -> str:
    return type(error).__name__


def _tymed_name(tymed: int) -> str:
    names = {
        pythoncom.TYMED_ISTREAM: "TYMED_ISTREAM",
        pythoncom.TYMED_HGLOBAL: "TYMED_HGLOBAL",
        pythoncom.TYMED_ISTORAGE: "TYMED_ISTORAGE",
    }
    if tymed in names:
        return names[tymed]
    selected = [name for value, name in names.items() if tymed & value]
    return "|".join(selected) if selected else f"TYMED_{tymed}"


def _clipboard_format_name(format_id: int) -> str:
    if format_id == win32con.CF_HDROP:
        return "CF_HDROP"
    if format_id == FILE_GROUP_DESCRIPTOR_W:
        return "FileGroupDescriptorW"
    if format_id == FILE_CONTENTS:
        return "FileContents"
    try:
        return win32clipboard.GetClipboardFormatName(format_id)
    except Exception:
        return f"format_{format_id}"


def _enum_formats(data_object: object) -> list[tuple[str, int, int, int]]:
    formats: list[tuple[str, int, int, int]] = []
    try:
        enumerator = data_object.EnumFormatEtc(pythoncom.DATADIR_GET)
        while True:
            batch = enumerator.Next(1)
            if not batch:
                break
            formatetc = batch[0]
            formats.append(
                (
                    _clipboard_format_name(int(formatetc[0])),
                    int(formatetc[2]),
                    int(formatetc[3]),
                    int(formatetc[4]),
                )
            )
    except Exception:
        return formats
    return formats


def _formatetc(format_id: int, index: int, tymed: int) -> tuple[int, None, int, int, int]:
    return (format_id, None, pythoncom.DVASPECT_CONTENT, index, tymed)


def _query_format(
    data_object: object,
    format_id: int,
    index: int,
    tymed: int,
    diagnostics: OleDndDiagnostics | None = None,
) -> bool:
    if diagnostics is not None:
        diagnostics.record(
            "query_get_data_start",
            format=_clipboard_format_name(format_id),
            lIndex=index,
            tymed=_tymed_name(tymed),
        )
    try:
        data_object.QueryGetData(_formatetc(format_id, index, tymed))
    except Exception as error:
        if diagnostics is not None:
            diagnostics.record(
                "query_get_data_result",
                format=_clipboard_format_name(format_id),
                lIndex=index,
                tymed=_tymed_name(tymed),
                supported=0,
                hresult=_hresult(error),
                exception=_exception_type(error),
            )
        return False
    if diagnostics is not None:
        diagnostics.record(
            "query_get_data_result",
            format=_clipboard_format_name(format_id),
            lIndex=index,
            tymed=_tymed_name(tymed),
            supported=1,
        )
    return True


def _supports_any_content_medium(
    data_object: object,
    index: int,
    diagnostics: OleDndDiagnostics | None = None,
) -> bool:
    return _query_format(
        data_object,
        FILE_CONTENTS,
        index,
        CONTENT_TYMED_MASK,
        diagnostics,
    )


def _extract_hdrop_paths(data_object: object) -> list[str]:
    medium = data_object.GetData(
        _formatetc(win32con.CF_HDROP, -1, pythoncom.TYMED_HGLOBAL)
    )
    if medium.tymed != pythoncom.TYMED_HGLOBAL:
        raise RuntimeError("CF_HDROP did not return TYMED_HGLOBAL")
    handle = medium.data_handle
    count = int(win32api.DragQueryFile(handle))
    return [str(win32api.DragQueryFile(handle, index)) for index in range(count)]


def _parse_file_group_descriptor_w(blob: bytes) -> list[str]:
    if len(blob) < 4:
        raise ValueError("FileGroupDescriptorW is truncated")
    count = struct.unpack_from("<I", blob, 0)[0]
    if count > MAX_DESCRIPTOR_COUNT:
        raise ValueError("FileGroupDescriptorW item count is unreasonable")

    descriptor_size = ctypes.sizeof(_FILEDESCRIPTORW)
    required_size = 4 + count * descriptor_size
    if len(blob) < required_size:
        raise ValueError("FileGroupDescriptorW descriptor array is truncated")

    filenames: list[str] = []
    for index in range(count):
        offset = 4 + index * descriptor_size
        descriptor = _FILEDESCRIPTORW.from_buffer_copy(blob, offset)
        filenames.append(str(descriptor.cFileName).split("\x00", 1)[0])
    return filenames


def _read_descriptor_names(
    data_object: object,
    diagnostics: OleDndDiagnostics,
) -> list[str]:
    if not _query_format(
        data_object,
        FILE_GROUP_DESCRIPTOR_W,
        -1,
        pythoncom.TYMED_HGLOBAL,
        diagnostics,
    ):
        raise RuntimeError("FileGroupDescriptorW is unavailable")
    diagnostics.record("descriptor_get_data_start")
    try:
        medium = data_object.GetData(
            _formatetc(FILE_GROUP_DESCRIPTOR_W, -1, pythoncom.TYMED_HGLOBAL)
        )
    except Exception as error:
        diagnostics.record(
            "descriptor_get_data_failure",
            hresult=_hresult(error),
            exception=_exception_type(error),
        )
        raise
    diagnostics.record("descriptor_get_data_success", tymed=_tymed_name(int(medium.tymed)))
    if medium.tymed != pythoncom.TYMED_HGLOBAL:
        raise RuntimeError("FileGroupDescriptorW did not return TYMED_HGLOBAL")
    names = _parse_file_group_descriptor_w(bytes(medium.data))
    diagnostics.record("descriptor_parse_success", count=len(names))
    return names


def _get_file_contents_medium(
    data_object: object,
    index: int,
    diagnostics: OleDndDiagnostics,
):
    if not _query_format(
        data_object,
        FILE_CONTENTS,
        index,
        CONTENT_TYMED_MASK,
        diagnostics,
    ):
        raise RuntimeError("FileContents is unavailable")

    diagnostics.record(
        "file_contents_get_data_start",
        lIndex=index,
        requested_tymed=_tymed_name(CONTENT_TYMED_MASK),
    )
    try:
        medium = data_object.GetData(
            _formatetc(FILE_CONTENTS, index, CONTENT_TYMED_MASK)
        )
    except Exception as error:
        diagnostics.record(
            "file_contents_get_data_failure",
            lIndex=index,
            requested_tymed=_tymed_name(CONTENT_TYMED_MASK),
            hresult=_hresult(error),
            exception=_exception_type(error),
        )
        raise
    actual_tymed = int(medium.tymed)
    diagnostics.record(
        "file_contents_get_data_success",
        lIndex=index,
        requested_tymed=_tymed_name(CONTENT_TYMED_MASK),
        actual_tymed=_tymed_name(actual_tymed),
    )
    if actual_tymed not in CONTENT_TYMEDS:
        raise RuntimeError(f"Unexpected FileContents tymed={actual_tymed}")
    return medium


def _write_stream(
    stream: object,
    destination: Path,
    diagnostics: OleDndDiagnostics,
    index: int,
) -> None:
    diagnostics.record("stream_write_start", lIndex=index)
    diagnostics.record("stream_seek_start", lIndex=index)
    try:
        stream.Seek(0, pythoncom.STREAM_SEEK_SET)
    except Exception as error:
        diagnostics.record(
            "stream_seek_failure",
            lIndex=index,
            hresult=_hresult(error),
            exception=_exception_type(error),
        )
    else:
        diagnostics.record("stream_seek_success", lIndex=index)
    written = 0
    with destination.open("wb") as file:
        while True:
            chunk = stream.Read(STREAM_CHUNK_SIZE)
            if not chunk:
                break
            file.write(chunk)
            written += len(chunk)
    if written == 0:
        raise RuntimeError("FileContents stream was empty")
    diagnostics.record("stream_write_complete", lIndex=index, bytes=written)


def _write_storage(
    storage: object,
    destination: Path,
    diagnostics: OleDndDiagnostics,
    index: int,
) -> None:
    diagnostics.record("storage_write_start", lIndex=index)
    mode = (
        storagecon.STGM_CREATE
        | storagecon.STGM_READWRITE
        | storagecon.STGM_SHARE_EXCLUSIVE
    )
    target = pythoncom.StgCreateDocfile(str(destination), mode, 0)
    try:
        storage.CopyTo(None, None, target)
        target.Commit(storagecon.STGC_DEFAULT)
    finally:
        target = None
    diagnostics.record("storage_write_complete", lIndex=index)


def _write_medium(
    medium: object,
    destination: Path,
    diagnostics: OleDndDiagnostics,
    index: int,
) -> int:
    tymed = int(medium.tymed)
    if tymed == pythoncom.TYMED_ISTREAM:
        _write_stream(medium.data, destination, diagnostics, index)
    elif tymed == pythoncom.TYMED_HGLOBAL:
        diagnostics.record("hglobal_write_start", lIndex=index)
        data = bytes(medium.data)
        destination.write_bytes(data)
        diagnostics.record("hglobal_write_complete", lIndex=index, bytes=len(data))
    elif tymed == pythoncom.TYMED_ISTORAGE:
        _write_storage(medium.data, destination, diagnostics, index)
    else:
        raise RuntimeError(f"Unsupported STGMEDIUM tymed={tymed}")
    return tymed


def _extract_virtual_msg_files(
    data_object: object,
    diagnostics: OleDndDiagnostics,
) -> tuple[list[str], tempfile.TemporaryDirectory[str]]:
    filenames = _read_descriptor_names(data_object, diagnostics)
    diagnostics.record("descriptor", count=len(filenames))
    temporary_directory = tempfile.TemporaryDirectory(prefix="DakeMail_List_Outlook_")
    paths: list[str] = []

    for index, filename in enumerate(filenames):
        if Path(filename).suffix.lower() != ".msg":
            diagnostics.record("item_skipped", lIndex=index, reason="not_msg")
            continue
        destination = Path(temporary_directory.name) / f"outlook_mail_{index + 1:04d}.msg"
        medium = None
        try:
            diagnostics.record("file_contents_begin", lIndex=index)
            medium = _get_file_contents_medium(data_object, index, diagnostics)
            tymed = _write_medium(medium, destination, diagnostics, index)
            paths.append(str(destination))
            diagnostics.record("item_success", lIndex=index, tymed=tymed)
        except Exception as error:
            diagnostics.record(
                "item_failure",
                lIndex=index,
                hresult=_hresult(error),
                exception=_exception_type(error),
            )
        finally:
            medium = None

    if not paths:
        temporary_directory.cleanup()
        raise RuntimeError("No Outlook .msg data was extracted")
    return paths, temporary_directory


class NativeOleDropTarget:
    _public_methods_ = ["DragEnter", "DragOver", "DragLeave", "Drop"]
    _com_interfaces_ = [pythoncom.IID_IDropTarget]

    def __init__(
        self,
        drop_queue: queue.Queue[OleDropPayload],
        diagnostics: OleDndDiagnostics,
    ) -> None:
        self._drop_queue = drop_queue
        self._diagnostics = diagnostics
        self._can_accept = False

    def _supports_drop(self, data_object: object) -> bool:
        if _query_format(data_object, win32con.CF_HDROP, -1, pythoncom.TYMED_HGLOBAL):
            return True
        return _query_format(
            data_object,
            FILE_GROUP_DESCRIPTOR_W,
            -1,
            pythoncom.TYMED_HGLOBAL,
        ) and _supports_any_content_medium(data_object, 0, self._diagnostics)

    def DragEnter(self, data_object, key_state, point, effect):
        formats = _enum_formats(data_object)
        format_names = [item[0] for item in formats]
        for format_name, aspect, index, tymed in formats:
            if format_name in ("FileGroupDescriptorW", "FileContents"):
                self._diagnostics.record(
                    "advertised_format",
                    format=format_name,
                    dwAspect=aspect,
                    lIndex=index,
                    tymed=_tymed_name(tymed),
                )
        self._can_accept = self._supports_drop(data_object)
        self._diagnostics.record(
            "drag_enter",
            formats=",".join(format_names),
            descriptor=int("FileGroupDescriptorW" in format_names),
            contents=int("FileContents" in format_names),
            accepted=int(self._can_accept),
        )
        return shellcon.DROPEFFECT_COPY if self._can_accept else shellcon.DROPEFFECT_NONE

    def DragOver(self, key_state, point, effect):
        return shellcon.DROPEFFECT_COPY if self._can_accept else shellcon.DROPEFFECT_NONE

    def DragLeave(self):
        self._can_accept = False
        self._diagnostics.record("drag_leave")

    def Drop(self, data_object, key_state, point, effect):
        self._diagnostics.record("drop_called")
        self._can_accept = False

        cleanup: object | None = None
        try:
            if _query_format(data_object, win32con.CF_HDROP, -1, pythoncom.TYMED_HGLOBAL):
                paths = _extract_hdrop_paths(data_object)
                self._diagnostics.record("hdrop", count=len(paths))
            else:
                paths, cleanup = _extract_virtual_msg_files(data_object, self._diagnostics)
            self._drop_queue.put(OleDropPayload(paths=paths, cleanup=cleanup))
            self._diagnostics.record("queue_put", count=len(paths), error=0)
        except Exception as error:
            if cleanup is not None:
                cleanup.cleanup()
            self._diagnostics.record(
                "drop_failure",
                hresult=_hresult(error),
                exception=_exception_type(error),
            )
            self._drop_queue.put(OleDropPayload(paths=[], error=True))
            self._diagnostics.record("queue_put", count=0, error=1)
            self._diagnostics.record("drop_end", effect="none")
            return shellcon.DROPEFFECT_NONE

        self._diagnostics.record("drop_success", count=len(paths))
        self._diagnostics.record("drop_end", effect="copy")
        return shellcon.DROPEFFECT_COPY


class NativeOleDropRegistration:
    def __init__(
        self,
        widgets: Iterable[object],
        drop_queue: queue.Queue[OleDropPayload],
    ) -> None:
        self.diagnostics = OleDndDiagnostics()
        self._target = NativeOleDropTarget(drop_queue, self.diagnostics)
        self._com_target = win32com.server.util.wrap(
            self._target,
            pythoncom.IID_IDropTarget,
        )
        self._hwnds: list[int] = []
        try:
            for widget in widgets:
                hwnd = int(widget.winfo_id())
                if hwnd in self._hwnds:
                    continue
                pythoncom.RegisterDragDrop(hwnd, self._com_target)
                self._hwnds.append(hwnd)
        except Exception:
            self.revoke()
            raise
        self.diagnostics.record("register", window_count=len(self._hwnds))

    def revoke(self) -> None:
        for hwnd in reversed(self._hwnds):
            try:
                pythoncom.RevokeDragDrop(hwnd)
            except pythoncom.com_error:
                pass
        if self._hwnds:
            self.diagnostics.record("revoke", window_count=len(self._hwnds))
        self._hwnds.clear()
