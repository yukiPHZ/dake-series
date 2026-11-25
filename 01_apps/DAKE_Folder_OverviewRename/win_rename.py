"""Windows no-replace rename with source and ancestor delete-sharing locks."""
import ctypes
from ctypes import wintypes as w
from contextlib import ExitStack, contextmanager
import os

@contextmanager
def locked_directory(path, delete=False):
    kernel=ctypes.WinDLL('kernel32',use_last_error=True)
    create=kernel.CreateFileW
    create.argtypes=[w.LPCWSTR,w.DWORD,w.DWORD,w.LPVOID,w.DWORD,w.DWORD,w.HANDLE]
    create.restype=w.HANDLE
    handle=create(str(path),0x80000000 | (0x10000 if delete else 0),3,None,3,0x02000000 | 0x00200000,None)
    if handle==w.HANDLE(-1).value:raise ctypes.WinError(ctypes.get_last_error())
    close=kernel.CloseHandle;close.argtypes=[w.HANDLE];close.restype=w.BOOL
    try:yield handle
    finally:close(handle)

def rename_checked(change,check):
    if os.name!='nt':raise OSError('windows_only')
    source=change.snapshot.path
    with ExitStack() as locks:
        # Deny deletion/rename of every ancestor while the final path is used.
        for parent in reversed(source.parents):locks.enter_context(locked_directory(parent))
        handle=locks.enter_context(locked_directory(source,True))
        check([change])
        class RenameInfo(ctypes.Structure):
            _fields_=[('ReplaceIfExists',w.BOOLEAN),('RootDirectory',w.HANDLE),('FileNameLength',w.DWORD),('FileName',w.WCHAR*1)]
        encoded=str(change.destination).encode('utf-16-le')
        buffer=ctypes.create_string_buffer(RenameInfo.FileName.offset+len(encoded)+2)
        info=RenameInfo.from_buffer(buffer)
        info.ReplaceIfExists=0;info.RootDirectory=None;info.FileNameLength=len(encoded)
        ctypes.memmove(ctypes.addressof(buffer)+RenameInfo.FileName.offset,encoded,len(encoded))
        api=ctypes.WinDLL('kernel32',use_last_error=True).SetFileInformationByHandle
        api.argtypes=[w.HANDLE,ctypes.c_int,w.LPVOID,w.DWORD];api.restype=w.BOOL
        if not api(handle,3,buffer,len(buffer)):raise ctypes.WinError(ctypes.get_last_error())
