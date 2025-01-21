import os
import msvcrt
import _winapi

CREATE_NEW                  = 1
CREATE_ALWAYS               = 2
OPEN_EXISTING               = 3
OPEN_ALWAYS                 = 4
TRUNCATE_EXISTING           = 5
FILE_SHARE_READ             = 0x00000001
FILE_SHARE_WRITE            = 0x00000002
FILE_SHARE_DELETE           = 0x00000004
FILE_SHARE_VALID_FLAGS      = 0x00000007
FILE_ATTRIBUTE_READONLY     = 0x00000001
FILE_ATTRIBUTE_NORMAL       = 0x00000080
FILE_ATTRIBUTE_TEMPORARY    = 0x00000100
FILE_FLAG_DELETE_ON_CLOSE   = 0x04000000
FILE_FLAG_SEQUENTIAL_SCAN   = 0x08000000
FILE_FLAG_RANDOM_ACCESS     = 0x10000000
GENERIC_READ                = 0x80000000
GENERIC_WRITE               = 0x40000000
DELETE                      = 0x00010000
NULL                        = 0

_ACCESS_MASK = os.O_RDONLY | os.O_WRONLY | os.O_RDWR
_ACCESS_MAP  = {os.O_RDONLY : GENERIC_READ,
                os.O_WRONLY : GENERIC_WRITE,
                os.O_RDWR   : GENERIC_READ | GENERIC_WRITE}

_CREATE_MASK = os.O_CREAT | os.O_EXCL | os.O_TRUNC
_CREATE_MAP  = {0                                   : OPEN_EXISTING,
                os.O_EXCL                           : OPEN_EXISTING,
                os.O_CREAT                          : OPEN_ALWAYS,
                os.O_CREAT | os.O_EXCL              : CREATE_NEW,
                os.O_CREAT | os.O_TRUNC | os.O_EXCL : CREATE_NEW,
                os.O_TRUNC                          : TRUNCATE_EXISTING,
                os.O_TRUNC | os.O_EXCL              : TRUNCATE_EXISTING,
                os.O_CREAT | os.O_TRUNC             : CREATE_ALWAYS}


def os_open(file, flags, mode=0o777,
            *, share_flags=FILE_SHARE_VALID_FLAGS):
    '''
    Replacement for os.open() allowing moving or unlinking before closing
    '''
    if not isinstance(flags, int) and mode >= 0:
        raise ValueError('bad flags: %r' % flags)

    if not isinstance(mode, int) and mode >= 0:
        raise ValueError('bad mode: %r' % mode)

    if share_flags & ~FILE_SHARE_VALID_FLAGS:
        raise ValueError('bad share_flags: %r' % share_flags)

    access_flags = _ACCESS_MAP[flags & _ACCESS_MASK]
    create_flags = _CREATE_MAP[flags & _CREATE_MASK]
    attrib_flags = FILE_ATTRIBUTE_NORMAL

    if flags & os.O_CREAT and mode & ~0o444 == 0:
        attrib_flags = FILE_ATTRIBUTE_READONLY

    if flags & os.O_TEMPORARY:
        share_flags |= FILE_SHARE_DELETE
        attrib_flags |= FILE_FLAG_DELETE_ON_CLOSE
        access_flags |= DELETE

    if flags & os.O_SHORT_LIVED:
        attrib_flags |= FILE_ATTRIBUTE_TEMPORARY

    if flags & os.O_SEQUENTIAL:
        attrib_flags |= FILE_FLAG_SEQUENTIAL_SCAN

    if flags & os.O_RANDOM:
        attrib_flags |= FILE_FLAG_RANDOM_ACCESS

    h = _winapi.CreateFile(file, access_flags, share_flags, NULL,
                           create_flags, attrib_flags, NULL)
    return msvcrt.open_osfhandle(h, flags | os.O_NOINHERIT)


import psutil
from ctypes.wintypes import HANDLE, DWORD, LPCWSTR
import ctypes

def write_to_file(handle, data:str):
    """Writes data to the file associated with the given handle."""
    # Prepare data for writing
    buffer = data.encode('utf-8')
    # buffer = ctypes.create_string_buffer((data).encode('utf-8'))  # Convert string to bytes
    # bytes_written = DWORD(0)  # To store the number of bytes written

    # Call WriteFile
    rc, bwr = _winapi.WriteFile(
        handle,
        buffer,
    )
    if not rc:
        raise OSError("Failed to write to the file.")
    return rc
    # return bytes_written.value

def acquire_file_and_write_pid(file_path):
    """Try to acquire an exclusive lock on the file."""
    handle = _winapi.CreateFile(
        file_path,
        GENERIC_WRITE,
        FILE_SHARE_READ,
        0,
        CREATE_ALWAYS,
        FILE_ATTRIBUTE_NORMAL,
        0,
    )
    write_to_file(handle, str(os.getpid()))
    return handle

def close_file(handle):
    pass
    _winapi.CloseHandle(handle)

def check_existing_lock_and_pid(lock_file):
    """Check if a process holding the lock is still running."""
    try:
        with open(lock_file, 'r') as f:
            pid = int(f.read().strip())
            if psutil.pid_exists(pid):
                return pid  # Process is still running
    except (ValueError, FileNotFoundError):
        pass
    return False


if __name__ == "__main__":
    import config
    config = config.load()

    p = config.xlpro_lock_path

    try:
        handle = acquire_file_and_write_pid(p)
    except Exception as e:
        b = check_existing_lock_and_pid(p)
        pass
    close_file(handle)

    


