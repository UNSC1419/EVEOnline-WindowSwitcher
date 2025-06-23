import win32con
import win32gui
import win32api
import time


def window_show(hwnd):
    """将窗口显示到最前面"""
    if not hwnd:
        return False

    try:
        # 先取消置顶
        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                              win32con.SWP_NOSIZE | win32con.SWP_NOMOVE)
        # 再设置为普通窗口
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, 0, 0,
                              win32con.SWP_SHOWWINDOW | win32con.SWP_NOSIZE | win32con.SWP_NOMOVE)
        return True
    except Exception:
        return False  # 移除错误输出


def window_maximize(hwnd):
    """将窗口最大化"""
    if not hwnd:
        return False

    try:
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        return True
    except:
        return False


def window_active(hwnd):
    """将窗口置顶激活"""
    if not hwnd:
        return False

    try:
        # 防止前台限制
        VK_MENU = 0x12
        win32api.keybd_event(VK_MENU, 0, 0, 0)

        # 激活窗口
        win32gui.SetForegroundWindow(hwnd)

        # 释放Alt键
        win32api.keybd_event(VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)
        return True
    except Exception:
        return False  # 移除错误输出


def window_is_maximized(hwnd):
    """判断窗口是否最大化"""
    if not hwnd:
        return False

    try:
        placement = win32gui.GetWindowPlacement(hwnd)
        return placement[1] == win32con.SW_SHOWMAXIMIZED
    except:
        return False


def window_is_active(hwnd):
    """判断窗口是否为当前激活窗口"""
    if not hwnd:
        return False

    try:
        return hwnd == win32gui.GetForegroundWindow()
    except:
        return False