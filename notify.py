import time
from win32api import *
import win32con
import os
import random
from win32gui import *

NIIF_USER = 4

class NotifIcon:
    def __init__(self, name, icon_path=os.path.abspath("logo.ico")):
        self.checkMessages = PumpWaitingMessages

        message_map = {
                win32con.WM_DESTROY: self.OnDestroy,
                win32con.WM_USER+20: self.onMessage
        }

        hicon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE

        wc = WNDCLASS()
        hinst = wc.hInstance = GetModuleHandle(None)
        wc.lpszClassName = f"PythonTaskbar_{name}"
        wc.lpfnWndProc = message_map
        classAtom = RegisterClass(wc)

        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = CreateWindow( classAtom, "Taskbar", style, \
                0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, \
                0, 0, hinst, None)
        UpdateWindow(self.hwnd)

        try:
            hicon = LoadImage(hinst, icon_path, win32con.IMAGE_ICON, 0, 0, hicon_flags)
        except:
            hicon = LoadIcon(0, win32con.IDI_APPLICATION)

        self.nid = {
            "hwnd": self.hwnd,
            "ID": random.randrange(1 << 30, 1 << 31),
            "flags": NIF_ICON | NIF_MESSAGE | NIF_TIP,
            "uCallbackMessage": win32con.WM_USER+20,
            "hicon": hicon,
            "name": name,
            "msg": "",
            "timeout": 5000, # deprecated > XP
            "title": ""
        }

        Shell_NotifyIcon(NIM_ADD, tuple(self.nid.values()))
        self.nid["flags"] = NIF_INFO

        self.show_balloon(" ", f"{name} started", NIIF_INFO)

    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_traceback):
        if exc_type:
            self.close(f"An error occured: {exc_type.__qualname__}")
        else:
            self.close()
        return self

    def OnDestroy(self, hwnd, msg, wparam, lparam):
        Shell_NotifyIcon(NIM_DELETE, tuple(self.nid.values()))
        PostQuitMessage(0)
        return 0
    
    def onMessage(self, hwnd, msg, wparam, lparam):
        if lparam == win32con.WM_LBUTTONDBLCLK:
            self.onQuit()
        return True
    
    def onQuit(self):
        PostQuitMessage(0)

    def show_balloon(self, msg, title=None, icon=NIIF_USER):
        if "an error occured" in msg.lower():
            icon = NIIF_ERROR
        if not title:
            title = self.nid["name"]
        self.nid["title"] = title

        self.nid["msg"] = ""
        Shell_NotifyIcon(NIM_MODIFY, tuple(self.nid.values()))

        self.nid["msg"] = msg
        Shell_NotifyIcon(NIM_MODIFY, (*self.nid.values(), icon))

    def close(self, exit_msg=" "):
        self.show_balloon(exit_msg, "Exiting....", NIIF_INFO)
        time.sleep(0.1)
        DestroyWindow(self.hwnd)

if __name__ == '__main__':
    with NotifIcon("Test notifier") as notifier:
        time.sleep(3)
        notifier.show_balloon("Double-click the tray icon to close")
        while True:
            if notifier.checkMessages() != 0:
                break
            time.sleep(0.05)