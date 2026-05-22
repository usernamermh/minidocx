from __future__ import annotations

import ctypes
import os
import sys
import threading
import time
from ctypes import wintypes


TOKEN_ADJUST_PRIVILEGES = 0x0020
TOKEN_QUERY = 0x0008
SE_PRIVILEGE_ENABLED = 0x00000002

PROCESS_QUERY_INFORMATION = 0x0400
PROCESS_SET_QUOTA = 0x0100


class FILETIME(ctypes.Structure):
    _fields_ = [
        ("dwLowDateTime", wintypes.DWORD),
        ("dwHighDateTime", wintypes.DWORD),
    ]


class MEMORYSTATUSEX(ctypes.Structure):
    _fields_ = [
        ("dwLength", wintypes.DWORD),
        ("dwMemoryLoad", wintypes.DWORD),
        ("ullTotalPhys", ctypes.c_ulonglong),
        ("ullAvailPhys", ctypes.c_ulonglong),
        ("ullTotalPageFile", ctypes.c_ulonglong),
        ("ullAvailPageFile", ctypes.c_ulonglong),
        ("ullTotalVirtual", ctypes.c_ulonglong),
        ("ullAvailVirtual", ctypes.c_ulonglong),
        ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
    ]


_cpu_lock = threading.Lock()
_last_cpu_times: tuple[int, int, int] | None = None


def _filetime_to_int(value: FILETIME) -> int:
    return (int(value.dwHighDateTime) << 32) + int(value.dwLowDateTime)


def _kernel32():
    if sys.platform != "win32":
        raise RuntimeError("Resource monitoring is only supported on Windows.")
    return ctypes.WinDLL("kernel32", use_last_error=True)


def _psapi():
    if sys.platform != "win32":
        raise RuntimeError("Resource cleanup is only supported on Windows.")
    return ctypes.WinDLL("psapi", use_last_error=True)


def get_cpu_percent() -> float:
    kernel32 = _kernel32()
    idle = FILETIME()
    kernel = FILETIME()
    user = FILETIME()
    if not kernel32.GetSystemTimes(ctypes.byref(idle), ctypes.byref(kernel), ctypes.byref(user)):
        raise ctypes.WinError(ctypes.get_last_error())

    current = (_filetime_to_int(idle), _filetime_to_int(kernel), _filetime_to_int(user))
    with _cpu_lock:
        global _last_cpu_times
        previous = _last_cpu_times
        _last_cpu_times = current

    if previous is None:
        return 0.0

    idle_delta = current[0] - previous[0]
    kernel_delta = current[1] - previous[1]
    user_delta = current[2] - previous[2]
    total_delta = kernel_delta + user_delta
    if total_delta <= 0:
        return 0.0
    busy_delta = max(total_delta - idle_delta, 0)
    return round(min(max(busy_delta * 100 / total_delta, 0.0), 100.0), 1)


def get_memory_info() -> dict:
    kernel32 = _kernel32()
    memory = MEMORYSTATUSEX()
    memory.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
    if not kernel32.GlobalMemoryStatusEx(ctypes.byref(memory)):
        raise ctypes.WinError(ctypes.get_last_error())

    total = int(memory.ullTotalPhys)
    available = int(memory.ullAvailPhys)
    used = max(total - available, 0)
    mb = 1024 * 1024
    return {
        "percent": int(memory.dwMemoryLoad),
        "total_mb": round(total / mb),
        "available_mb": round(available / mb),
        "used_mb": round(used / mb),
    }


def get_resource_stats() -> dict:
    return {
        "ok": True,
        "ts": time.time(),
        "cpu_percent": get_cpu_percent(),
        "memory": get_memory_info(),
    }


def enable_increase_quota_privilege() -> bool:
    if sys.platform != "win32":
        return False
    advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)
    kernel32 = _kernel32()

    token = wintypes.HANDLE()
    if not advapi32.OpenProcessToken(
        kernel32.GetCurrentProcess(),
        TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY,
        ctypes.byref(token),
    ):
        return False

    class LUID(ctypes.Structure):
        _fields_ = [("LowPart", wintypes.DWORD), ("HighPart", wintypes.LONG)]

    class LUID_AND_ATTRIBUTES(ctypes.Structure):
        _fields_ = [("Luid", LUID), ("Attributes", wintypes.DWORD)]

    class TOKEN_PRIVILEGES(ctypes.Structure):
        _fields_ = [
            ("PrivilegeCount", wintypes.DWORD),
            ("Privileges", LUID_AND_ATTRIBUTES * 1),
        ]

    try:
        luid = LUID()
        if not advapi32.LookupPrivilegeValueW(None, "SeIncreaseQuotaPrivilege", ctypes.byref(luid)):
            return False

        privileges = TOKEN_PRIVILEGES()
        privileges.PrivilegeCount = 1
        privileges.Privileges[0].Luid = luid
        privileges.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED

        return bool(
            advapi32.AdjustTokenPrivileges(
                token,
                False,
                ctypes.byref(privileges),
                ctypes.sizeof(privileges),
                None,
                None,
            )
        )
    finally:
        kernel32.CloseHandle(token)


def is_admin() -> bool:
    if sys.platform != "win32":
        return False
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def empty_all_working_sets() -> dict:
    psapi = _psapi()
    kernel32 = _kernel32()
    privilege_enabled = enable_increase_quota_privilege()

    enum_processes = psapi.EnumProcesses
    enum_processes.argtypes = [wintypes.LPDWORD, wintypes.DWORD, wintypes.LPDWORD]
    enum_processes.restype = wintypes.BOOL

    open_process = kernel32.OpenProcess
    open_process.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    open_process.restype = wintypes.HANDLE

    empty_working_set = psapi.EmptyWorkingSet
    empty_working_set.argtypes = [wintypes.HANDLE]
    empty_working_set.restype = wintypes.BOOL

    close_handle = kernel32.CloseHandle
    close_handle.argtypes = [wintypes.HANDLE]
    close_handle.restype = wintypes.BOOL

    max_processes = 8192
    pid_array_type = wintypes.DWORD * max_processes
    pids = pid_array_type()
    bytes_returned = wintypes.DWORD()
    if not enum_processes(pids, ctypes.sizeof(pids), ctypes.byref(bytes_returned)):
        raise ctypes.WinError(ctypes.get_last_error())

    pid_count = bytes_returned.value // ctypes.sizeof(wintypes.DWORD)
    cleaned_count = 0
    attempted_count = 0
    for pid in pids[:pid_count]:
        if pid in (0, 4, os.getpid()):
            continue
        attempted_count += 1
        handle = open_process(PROCESS_QUERY_INFORMATION | PROCESS_SET_QUOTA, False, int(pid))
        if not handle:
            continue
        try:
            if empty_working_set(handle):
                cleaned_count += 1
        finally:
            close_handle(handle)

    return {
        "ok": True,
        "attempted_count": attempted_count,
        "cleaned_count": cleaned_count,
        "privilege_enabled": privilege_enabled,
        "is_admin": is_admin(),
        "stats": get_resource_stats(),
    }
