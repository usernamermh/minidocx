import ctypes
from ctypes import wintypes
import time
import sys

# 加载 DLL
advapi32 = ctypes.WinDLL('advapi32.dll')
kernel32 = ctypes.WinDLL('kernel32.dll')
psapi = ctypes.WinDLL('psapi.dll')

# === 权限常量 ===
TOKEN_ADJUST_PRIVILEGES = 0x0020
TOKEN_QUERY = 0x0008
SE_PRIVILEGE_ENABLED = 0x00000002

PROCESS_QUERY_INFORMATION = 0x0400
PROCESS_SET_QUOTA = 0x0100  # 必需！

# === 启用 SeIncreaseQuotaPrivilege ===
def enable_increase_quota_privilege():
    """启用 'SeIncreaseQuotaPrivilege'，这是 EmptyWorkingSet 所需的"""
    hToken = wintypes.HANDLE()
    if not advapi32.OpenProcessToken(
        kernel32.GetCurrentProcess(),
        TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY,
        ctypes.byref(hToken)
    ):
        return False

    class LUID(ctypes.Structure):
        _fields_ = [("LowPart", wintypes.DWORD), ("HighPart", wintypes.LONG)]

    class LUID_AND_ATTRIBUTES(ctypes.Structure):
        _fields_ = [("Luid", LUID), ("Attributes", wintypes.DWORD)]

    class TOKEN_PRIVILEGES(ctypes.Structure):
        _fields_ = [
            ("PrivilegeCount", wintypes.DWORD),
            ("Privileges", LUID_AND_ATTRIBUTES * 1)
        ]

    luid = LUID()
    if not advapi32.LookupPrivilegeValueW(None, "SeIncreaseQuotaPrivilege", ctypes.byref(luid)):
        kernel32.CloseHandle(hToken)
        return False

    tp = TOKEN_PRIVILEGES()
    tp.PrivilegeCount = 1
    tp.Privileges[0].Luid = luid
    tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED

    result = advapi32.AdjustTokenPrivileges(
        hToken,
        False,
        ctypes.byref(tp),
        ctypes.sizeof(tp),
        None,
        None
    )
    kernel32.CloseHandle(hToken)
    return result != 0

# === API 声明 ===
EnumProcesses = psapi.EnumProcesses
EnumProcesses.argtypes = [wintypes.LPDWORD, wintypes.DWORD, wintypes.LPDWORD]
EnumProcesses.restype = wintypes.BOOL

OpenProcess = kernel32.OpenProcess
OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
OpenProcess.restype = wintypes.HANDLE

EmptyWorkingSet = psapi.EmptyWorkingSet
EmptyWorkingSet.argtypes = [wintypes.HANDLE]
EmptyWorkingSet.restype = wintypes.BOOL

CloseHandle = kernel32.CloseHandle
CloseHandle.argtypes = [wintypes.HANDLE]
CloseHandle.restype = wintypes.BOOL

def empty_all_working_sets():
    max_processes = 2048
    arr_type = wintypes.DWORD * max_processes
    pids = arr_type()
    bytes_returned = wintypes.DWORD()

    if not EnumProcesses(pids, ctypes.sizeof(pids), ctypes.byref(bytes_returned)):
        return

    num_pids = bytes_returned.value // ctypes.sizeof(wintypes.DWORD)
    pid_list = pids[:num_pids]

    emptied_count = 0
    for pid in pid_list:
        if pid in (0, 4):  # Skip Idle and System
            continue
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_SET_QUOTA, False, pid)
        if hProcess:
            if EmptyWorkingSet(hProcess):
                emptied_count += 1
            CloseHandle(hProcess)
    print(f"[{time.strftime('%H:%M:%S')}] 清空了 {emptied_count} 个进程的工作集")

def main():
    print("🔄 正在初始化...")
    
    # 检查并启用特权
    if not enable_increase_quota_privilege():
        print("⚠️  无法启用 SeIncreaseQuotaPrivilege（需要管理员权限）")
    else:
        print("✅ 已启用 SeIncreaseQuotaPrivilege")

    # 检查管理员
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
    except:
        is_admin = False
    if not is_admin:
        print("⚠️  警告：未以管理员身份运行！效果将受限。\n")
    else:
        print("✅ 当前为管理员模式\n")

    print("开始自动清空工作集（每5秒一次）...")
    print("按 Ctrl+C 停止\n")
    
    try:
        while True:
            empty_all_working_sets()
            time.sleep(60 * 60)
    except KeyboardInterrupt:
        print("\n🛑 已停止")

if __name__ == "__main__":
    if sys.platform != "win32":
        print("仅支持 Windows！")
        exit(1)
    main()