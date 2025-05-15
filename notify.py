import time
from win32api import *
import win32con
import os
import random
from win32gui import *

NIIF_USER = 4

class NotifIcon:
    def __init__(self, name, icon_path=os.path.abspath("logo.ico"), timeout=200):
        self.timeout = timeout
        self.ID = random.randrange(1 << 30, 1 << 31)
        self.name = name
        message_map = {
                win32con.WM_DESTROY: self.OnDestroy,
        }

        wc = WNDCLASS()
        hinst = wc.hInstance = GetModuleHandle(None)
        wc.lpszClassName = f"PythonTaskbar_{self.name}"
        wc.lpfnWndProc = message_map
        classAtom = RegisterClass(wc)

        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = CreateWindow( classAtom, "Taskbar", style, \
                0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, \
                0, 0, hinst, None)
        UpdateWindow(self.hwnd)

        icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
        try:
            self.hicon = LoadImage(hinst, icon_path, win32con.IMAGE_ICON, 0, 0, icon_flags)
        except:
            self.hicon = LoadIcon(0, win32con.IDI_APPLICATION)

        self.flags = NIF_ICON | NIF_MESSAGE | NIF_TIP

        nid = (self.hwnd, self.ID, self.flags, win32con.WM_USER+20, self.hicon, self.name)
        Shell_NotifyIcon(NIM_ADD, nid)

        self.show_balloon(" ", f"{self.name} started", NIIF_INFO)

    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_traceback):
        if exc_type:
            self.close(f"An error occured: {exc_type.__qualname__}", 10)
        else:
            self.close()

    def OnDestroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, self.ID, NIF_INFO, win32con.WM_USER+20, self.hicon, self.name, "")
        Shell_NotifyIcon(NIM_MODIFY, nid)
        PostQuitMessage(0)
        return 0
    
    def show_balloon(self, msg, title=None, icon=NIIF_USER):
        if "an error occured" in msg.lower():
            icon = NIIF_ERROR
        if not title:
            title = self.name

        nid = (self.hwnd, self.ID, NIF_INFO, win32con.WM_USER+20, self.hicon, self.name, "")
        Shell_NotifyIcon(NIM_MODIFY, nid)

        nid = (self.hwnd, self.ID, NIF_INFO, win32con.WM_USER+20, self.hicon, self.name, msg, self.timeout, title, icon)
        Shell_NotifyIcon(NIM_MODIFY, nid)

    def close(self, exit_msg=" ", wait=3):
        self.show_balloon(exit_msg, "Exiting....", NIIF_INFO)
        time.sleep(wait)
        DestroyWindow(self.hwnd)