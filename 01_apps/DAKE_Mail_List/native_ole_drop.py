# -*- coding: utf-8 -*-
from __future__ import annotations

import ctypes
import struct
import tempfile
import threading
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterable

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
    pythoncom.TYMED_ISTREAM,
    pythoncom.TYMED_HGLOBAL,
    pythoncom.TYMED_ISTORAGE,
)
STREAM_CHUNK_SIZE = 1024 * 1024
MAX_DESCRIPTOR_COUNT = 1000
DIAGNOSTIC_LOG_PATH = Path(tempfile.gettempdir()) / "DakeMail_List_ole_dnd.log"


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


def _enum_format_names(data_object: object) -> list[str]:
    names: list[str] = []
    try:
        enumerator = data_object.EnumFormatEtc(pythoncom.DATADIR_GET)
        while True:
            batch = enumerator.Next(1)
            if not batch:
                break
            names.append(_clipboard_format_name(int(batch[0][0])))
    except Exception:
        return names
    return names


def _formatetc(format_id: int, index: int, tymed: int) -> tuple[int, None, int, int, int]:
    return (format_id, None, pythoncom.DVASPECT_CONTENT, index, tymed)


def _query_format(data_object: object, format_id: int, index: int, tymed: int) -> bool:
    try:
        data_object.QueryGetData(_formatetc(format_id, index, tymed))
    except Exception:
        return False
    return True


def _supports_any_content_medium(data_object: object, index: int) -> bool:
    return any(_query_format(data_object, FILE_CONTENTS, index, tymed) for tymed in CONTENT_TYMEDS)


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


def _read_descriptor_names(data_object: object) -> list[str]:
    medium = data_object.GetData(
        _formatetc(FILE_GROUP_DESCRIPTOR_W, -1, pythoncom.TYMED_HGLOBAL)
    )
    if medium.tymed != pythoncom.TYMED_HGLOBAL:
        raise RuntimeError("FileGroupDescriptorW did not return TYMED_HGLOBAL")
    return _parse_file_group_descriptor_w(bytes(medium.data))


def _get_file_contents_medium(data_object: object, index: int):
    last_error: BaseException | None = None
    for tymed in CONTENT_TYMEDS:
        try:
            return data_object.GetData(_formatetc(FILE_CONTENTS, index, tymed))
        except Exception as error:
            last_error = error
    if last_error is not None:
        raise last_error
    raise RuntimeError("FileContents is unavailable")


def _write_stream(stream: object, destination: Path) -> None:
    try:
        stream.Seek(0, pythoncom.STREAM_SEEK_SET)
    except Exception:
        pass
    with destination.open("wb") as file:
        while True:
            chunk = stream.Read(STREAM_CHUNK_SIZE)
            if not chunk:
                break
            file.write(chunk)


def _write_storage(storage: object, destination: Path) -> None:
    mode = storagecon.STGM_CREATE | storagecon.STGM_READWRITE | storagecon.STGM_SHARE_EXCLUSIVE
    target = pythoncom.StgCreateDocfile(str(destination), mode, 0)
    storage.CopyTo(None, None, target)
    target.Commit(storagecon.STGC_DEFAULT)


def _write_medium(medium: object, destination: Path) -> int:
    tymed = int(medium.tymed)
    if tymed == pythoncom.TYMED_ISTREAM:
        _write_stream(medium.data, destination)
    elif tymed == pythoncom.TYMED_HGLOBAL:
        destination.write_bytes(bytes(medium.data))
    elif tymed == pythoncom.TYMED_ISTORAGE:
        _write_storage(medium.data, destination)
    else:
        raise RuntimeError(f"Unsupported STGMEDIUM tymed={tymed}")
    return tymed


def _extract_virtual_msg_files(
    data_object: object,
    diagnostics: OleDndDiagnostics,
) -> tuple[list[str], tempfile.TemporaryDirectory[str]]:
    filenames = _read_descriptor_names(data_object)
    diagnostics.record("descriptor", count=len(filenames))
    temporary_directory = tempfile.TemporaryDirectory(prefix="DakeMail_List_Outlook_")
    paths: list[str] = []

    for index, filename in enumerate(filenames):
        if Path(filename).suffix.lower() != ".msg":
            diagnostics.record("item_skipped", lIndex=index, reason="not_msg")
            continue
        destination = Path(temporary_directory.name) / f"outlook_mail_{index + 1:04d}.msg"
        try:
            medium = _get_file_contents_medium(data_object, index)
            tymed = _write_medium(medium, destination)
            paths.append(str(destination))
            diagnostics.record("item_success", lIndex=index, tymed=tymed)
        except Exception as error:
            diagnostics.record("item_failure", lIndex=index, hresult=_hresult(error))

    if not paths:
        temporary_directory.cleanup()
        raise RuntimeError("No Outlook .msg data was extracted")
    return paths, temporary_directory


class NativeOleDropTarget:
    _public_methods_ = ["DragEnter", "DragOver", "DragLeave", "Drop"]
    _com_interfaces_ = [pythoncom.IID_IDropTarget]

    def __init__(
        self,
        on_drop: Callable[[list[str], object | None], bool],
        is_busy: Callable[[], bool],
        on_error: Callable[[], None],
        diagnostics: OleDndDiagnostics,
    ) -> None:
        self._on_drop = on_drop
        self._is_busy = is_busy
        self._on_error = on_error
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
        ) and _supports_any_content_medium(data_object, 0)

    def DragEnter(self, data_object, key_state, point, effect):
        format_names = _enum_format_names(data_object)
        self._can_accept = not self._is_busy() and self._supports_drop(data_object)
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
        if self._is_busy():
            self._diagnostics.record("drop_rejected", reason="busy")
            return shellcon.DROPEFFECT_NONE

        cleanup: object | None = None
        try:
            if _query_format(data_object, win32con.CF_HDROP, -1, pythoncom.TYMED_HGLOBAL):
                paths = _extract_hdrop_paths(data_object)
                self._diagnostics.record("hdrop", count=len(paths))
            else:
                paths, cleanup = _extract_virtual_msg_files(data_object, self._diagnostics)
            accepted = self._on_drop(paths, cleanup)
        except Exception as error:
            if cleanup is not None:
                cleanup.cleanup()
            self._diagnostics.record("drop_failure", hresult=_hresult(error))
            self._on_error()
            return shellcon.DROPEFFECT_NONE

        if not accepted:
            if cleanup is not None:
                cleanup.cleanup()
            self._diagnostics.record("drop_rejected", reason="no_msg")
            return shellcon.DROPEFFECT_NONE

        self._diagnostics.record("drop_success", count=len(paths))
        return shellcon.DROPEFFECT_COPY


class NativeOleDropRegistration:
    def __init__(
        self,
        widgets: Iterable[object],
        on_drop: Callable[[list[str], object | None], bool],
        is_busy: Callable[[], bool],
        on_error: Callable[[], None],
    ) -> None:
        self.diagnostics = OleDndDiagnostics()
        self._target = NativeOleDropTarget(on_drop, is_busy, on_error, self.diagnostics)
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
