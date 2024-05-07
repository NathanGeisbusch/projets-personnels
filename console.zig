const std = @import("std");
const windows = std.os.windows;

const SW_HIDE = 0;

extern "kernel32" fn GetConsoleWindow() callconv(windows.WINAPI) windows.HWND;
extern "kernel32" fn ShowWindow(hWnd: windows.HWND, nCmdShow: windows.INT) callconv(windows.WINAPI) windows.BOOL;

pub fn hide() void {
    _ = ShowWindow(GetConsoleWindow(), SW_HIDE);
}
