const windows = @import("std").os.windows;

const WH_KEYBOARD_LL = 13;
const WH_MOUSE_LL = 14;
const HHOOK = *opaque {};
const HOOKPROC = *const fn (windows.INT, windows.WPARAM, windows.LPARAM) callconv(.C) windows.LRESULT;
const MSG = extern struct {
    hwnd: ?windows.HWND,
    message: windows.UINT,
    wParam: windows.WPARAM,
    lParam: windows.LPARAM,
    time: windows.UINT,
    pt: windows.POINT,
};

extern "user32" fn SetWindowsHookExA(idHook: windows.INT, lpfn: HOOKPROC, hmod: ?windows.HINSTANCE, dwThreadId: windows.DWORD) callconv(windows.WINAPI) HHOOK;

extern "user32" fn CallNextHookEx(hhk: ?HHOOK, nCode: windows.INT, wParam: windows.WPARAM, lParam: windows.LPARAM) callconv(windows.WINAPI) windows.LRESULT;

extern "user32" fn UnhookWindowsHookEx(hhk: ?HHOOK) callconv(windows.WINAPI) windows.BOOL;

extern "user32" fn GetMessageA(msg: ?*MSG, hwnd: ?windows.HWND, uMsgFilterMin: windows.UINT, uMsgFilterMax: windows.UINT) callconv(windows.WINAPI) windows.BOOL;

pub const Keylogger = struct {
    hookKeyboard: ?HHOOK = null,
    hookMouse: ?HHOOK = null,

    pub fn run(self: *Keylogger, comptime callback: fn () void) void {
        const hookCallback = struct {
            fn closure(nCode: windows.INT, wParam: windows.WPARAM, lParam: windows.LPARAM) callconv(.C) windows.LRESULT {
                callback();
                return CallNextHookEx(null, nCode, wParam, lParam);
            }
        }.closure;
        var msg = MSG{ .hwnd = null, .message = 0, .wParam = 0, .lParam = 0, .time = 0, .pt = windows.POINT{ .x = 0, .y = 0 } };
        self.hookKeyboard = SetWindowsHookExA(WH_KEYBOARD_LL, &hookCallback, null, 0);
        self.hookMouse = SetWindowsHookExA(WH_MOUSE_LL, &hookCallback, null, 0);
        _ = GetMessageA(&msg, null, 0, 0);
    }

    pub fn stop(self: *Keylogger) void {
        if (self.hookKeyboard != null) {
            _ = UnhookWindowsHookEx(self.hookKeyboard);
            self.hookKeyboard = null;
        }
        if (self.hookMouse != null) {
            _ = UnhookWindowsHookEx(self.hookMouse);
            self.hookMouse = null;
        }
    }
};
