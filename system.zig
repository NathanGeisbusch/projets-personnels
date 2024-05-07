const std = @import("std");
const windows = std.os.windows;

const LUID = extern struct {
    LowPart: windows.DWORD,
    HighPart: windows.LONG,
};
const LUID_AND_ATTRIBUTES = extern struct {
    Luid: LUID,
    Attributes: windows.DWORD,
};
const TOKEN_PRIVILEGES = extern struct {
    PrivilegeCount: windows.DWORD,
    Privileges: [1]LUID_AND_ATTRIBUTES,
};
const TOKEN_ADJUST_PRIVILEGES = 0x20;
const TOKEN_QUERY = 0x8;
const SE_PRIVILEGE_ENABLED = 0x2;
const SE_SHUTDOWN_NAME = "SeShutdownPrivilege";
const EWX_SHUTDOWN = 0x1;
const SHTDN_REASON_FLAG_USER_DEFINED = 0x40000000;

extern "advapi32" fn OpenProcessToken(ProcessHandle: windows.HANDLE, DesiredAccess: windows.DWORD, TokenHandle: *windows.HANDLE) callconv(windows.WINAPI) windows.BOOL;

extern "advapi32" fn LookupPrivilegeValueA(lpSystemName: ?windows.LPCSTR, lpName: ?windows.LPCSTR, lpLuid: *LUID) callconv(windows.WINAPI) windows.BOOL;

extern "advapi32" fn AdjustTokenPrivileges(
    TokenHandle: ?windows.HANDLE,
    DisableAllPrivileges: windows.BOOL,
    NewState: ?*TOKEN_PRIVILEGES,
    BufferLength: windows.UINT,
    PreviousState: ?*TOKEN_PRIVILEGES,
    ReturnLength: ?*windows.UINT,
) callconv(windows.WINAPI) windows.BOOL;

extern "user32" fn ExitWindowsEx(uFlags: windows.UINT, dwReason: windows.DWORD) callconv(windows.WINAPI) windows.BOOL;

pub fn shutdown() void {
    var hToken: windows.HANDLE = undefined;
    var tkp: TOKEN_PRIVILEGES = undefined;
    if (OpenProcessToken(windows.GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken) == 0) {
        return;
    }
    if (LookupPrivilegeValueA(null, SE_SHUTDOWN_NAME, &tkp.Privileges[0].Luid) == 0) {
        return;
    }
    tkp.PrivilegeCount = 1;
    tkp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;
    if (AdjustTokenPrivileges(hToken, 0, &tkp, 0, null, null) == 0) {
        return;
    }
    if (ExitWindowsEx(EWX_SHUTDOWN, SHTDN_REASON_FLAG_USER_DEFINED) == 0) {
        return;
    }
}
