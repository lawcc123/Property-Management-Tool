"""Keep Windows workbook visibility when atomically publishing a staged file."""
import os


VISIBILITY = 0x2 | 0x4  # FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM
NORMAL = 0x80


def preserve_visibility(source, staged):
    """Set visibility before replacement so an attribute failure leaves input intact.

    Do not copy all file attributes: archive, compression and other filesystem
    state belong to the newly written file. Intentionally hidden inputs stay hidden.
    """
    if os.name != "nt":
        return
    wanted = os.stat(source).st_file_attributes & VISIBILITY
    _set_visibility(staged, wanted)


def prepare_staged_file(staged):
    """Some shares hide new files, which also prevents truncating them on Windows."""
    if os.name == "nt":
        _set_visibility(staged, 0)


def _set_visibility(staged, wanted):
    current = os.stat(staged).st_file_attributes
    if current & VISIBILITY == wanted:
        return
    attributes = (current & ~VISIBILITY) | wanted
    attributes = (attributes & ~NORMAL) or NORMAL
    import ctypes
    from ctypes import wintypes
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    set_attributes = kernel.SetFileAttributesW
    set_attributes.argtypes = [wintypes.LPCWSTR, wintypes.DWORD]
    set_attributes.restype = wintypes.BOOL
    if not set_attributes(os.fspath(staged), attributes):
        raise ctypes.WinError(ctypes.get_last_error())
