import ctypes
import threading
from ctypes import wintypes

from PySide6.QtCore import QObject, Signal


WH_KEYBOARD_LL = 13
HC_ACTION = 0

WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_SYSKEYDOWN = 0x0104
WM_SYSKEYUP = 0x0105
WM_QUIT = 0x0012

VK_R = 0x52
VK_ESCAPE = 0x1B


LRESULT = ctypes.c_ssize_t
HHOOK = wintypes.HANDLE
HINSTANCE = wintypes.HANDLE

HOOKPROC = ctypes.WINFUNCTYPE(
    LRESULT,
    ctypes.c_int,
    wintypes.WPARAM,
    wintypes.LPARAM,
)


class KBDLLHOOKSTRUCT(ctypes.Structure):
    _fields_ = [
        ("vkCode", wintypes.DWORD),
        ("scanCode", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class GlobalReplacementHotkeys(QObject):
    replace_requested = Signal()
    clear_requested = Signal()
    activation_failed = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)

        self.user32 = ctypes.WinDLL(
            "user32",
            use_last_error=True,
        )
        self.kernel32 = ctypes.WinDLL(
            "kernel32",
            use_last_error=True,
        )

        self._configure_windows_api()

        self.active = False
        self._hook = None
        self._hook_proc = None
        self._thread = None
        self._thread_id = 0
        self._ready = threading.Event()
        self._install_error = ""

    def _configure_windows_api(self):
        self.kernel32.GetCurrentThreadId.argtypes = []
        self.kernel32.GetCurrentThreadId.restype = wintypes.DWORD

        self.kernel32.GetModuleHandleW.argtypes = [
            wintypes.LPCWSTR,
        ]
        self.kernel32.GetModuleHandleW.restype = HINSTANCE

        self.user32.SetWindowsHookExW.argtypes = [
            ctypes.c_int,
            HOOKPROC,
            HINSTANCE,
            wintypes.DWORD,
        ]
        self.user32.SetWindowsHookExW.restype = HHOOK

        self.user32.CallNextHookEx.argtypes = [
            HHOOK,
            ctypes.c_int,
            wintypes.WPARAM,
            wintypes.LPARAM,
        ]
        self.user32.CallNextHookEx.restype = LRESULT

        self.user32.UnhookWindowsHookEx.argtypes = [
            HHOOK,
        ]
        self.user32.UnhookWindowsHookEx.restype = wintypes.BOOL

        self.user32.GetMessageW.argtypes = [
            ctypes.POINTER(wintypes.MSG),
            wintypes.HWND,
            wintypes.UINT,
            wintypes.UINT,
        ]
        self.user32.GetMessageW.restype = wintypes.BOOL

        self.user32.TranslateMessage.argtypes = [
            ctypes.POINTER(wintypes.MSG),
        ]
        self.user32.TranslateMessage.restype = wintypes.BOOL

        self.user32.DispatchMessageW.argtypes = [
            ctypes.POINTER(wintypes.MSG),
        ]
        self.user32.DispatchMessageW.restype = LRESULT

        self.user32.PostThreadMessageW.argtypes = [
            wintypes.DWORD,
            wintypes.UINT,
            wintypes.WPARAM,
            wintypes.LPARAM,
        ]
        self.user32.PostThreadMessageW.restype = wintypes.BOOL

    def activate(self):
        if self.active:
            return True

        self._ready.clear()
        self._install_error = ""

        self._thread = threading.Thread(
            target=self._keyboard_thread,
            daemon=True,
        )
        self._thread.start()

        if not self._ready.wait(timeout=3):
            self.activation_failed.emit(
                "ToolPy could not start the keyboard listener."
            )
            return False

        if self._install_error:
            self.activation_failed.emit(
                self._install_error
            )
            return False

        self.active = True
        return True

    def deactivate(self):
        if not self.active and self._thread is None:
            return

        self.active = False

        if self._thread_id:
            self.user32.PostThreadMessageW(
                self._thread_id,
                WM_QUIT,
                0,
                0,
            )

        if (
            self._thread is not None
            and self._thread.is_alive()
        ):
            self._thread.join(timeout=2)

        self._thread = None
        self._thread_id = 0
        self._hook = None
        self._hook_proc = None

    def _keyboard_thread(self):
        self._thread_id = (
            self.kernel32.GetCurrentThreadId()
        )

        @HOOKPROC
        def hook_callback(
            code,
            w_param,
            l_param,
        ):
            if code == HC_ACTION:
                keyboard = ctypes.cast(
                    l_param,
                    ctypes.POINTER(KBDLLHOOKSTRUCT),
                ).contents

                key = keyboard.vkCode

                if key in (VK_R, VK_ESCAPE):
                    if w_param in (
                        WM_KEYDOWN,
                        WM_SYSKEYDOWN,
                    ):
                        if key == VK_R:
                            self.replace_requested.emit()
                        else:
                            self.clear_requested.emit()

                    if w_param in (
                        WM_KEYDOWN,
                        WM_KEYUP,
                        WM_SYSKEYDOWN,
                        WM_SYSKEYUP,
                    ):
                        return 1

            return self.user32.CallNextHookEx(
                self._hook,
                code,
                w_param,
                l_param,
            )

        self._hook_proc = hook_callback

        module_handle = (
            self.kernel32.GetModuleHandleW(None)
        )

        if not module_handle:
            error_code = ctypes.get_last_error()
            self._install_error = (
                "ToolPy could not get its Windows module handle.\n\n"
                f"Windows error code: {error_code}"
            )
            self._ready.set()
            return

        ctypes.set_last_error(0)

        self._hook = self.user32.SetWindowsHookExW(
            WH_KEYBOARD_LL,
            self._hook_proc,
            module_handle,
            0,
        )

        if not self._hook:
            error_code = ctypes.get_last_error()
            self._install_error = (
                "ToolPy could not activate R and Esc.\n\n"
                f"Windows error code: {error_code}"
            )
            self._ready.set()
            return

        self._ready.set()

        message = wintypes.MSG()

        while self.user32.GetMessageW(
            ctypes.byref(message),
            None,
            0,
            0,
        ) > 0:
            self.user32.TranslateMessage(
                ctypes.byref(message)
            )
            self.user32.DispatchMessageW(
                ctypes.byref(message)
            )

        if self._hook:
            self.user32.UnhookWindowsHookEx(
                self._hook
            )

        self._hook = None
