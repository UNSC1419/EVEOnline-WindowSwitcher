import win32api
import win32con
import win32gui
import win32process
import ctypes
import sys
import psutil


def run_as_admin() -> bool:
    """请求以管理员身份运行脚本"""
    if ctypes.windll.shell32.IsUserAnAdmin():
        return True
    else:
        if ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1) == 42:
            return True
        else:
            raise RuntimeError('提升权限失败')


def get_game_processes_list() -> list:
    """获取以"EVE - "开头的EVE游戏客户端进程"""
    result = []

    def call_back(hwnd, context):
        text = win32gui.GetWindowText(hwnd)
        if text.startswith('EVE - '):
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            processes_name = text.replace('EVE - ', '')
            context.append({'title': processes_name, 'pid': pid, 'hwnd': hwnd})

    win32gui.EnumWindows(call_back, result)
    return result


def get_new_game_processes_list() -> list:
    """获取以"EVE"开头的EVE游戏客户端进程"""
    result = []

    def call_back(hwnd, context):
        text = win32gui.GetWindowText(hwnd)
        if text.startswith('EVE'):
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            try:
                proc = psutil.Process(pid)
                if proc.name().lower() == 'exefile.exe':
                    processes_name = text.replace('EVE - ', '')
                    context.append({'title': processes_name, 'pid': pid, 'hwnd': hwnd})
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass  # 移除日志输出

    win32gui.EnumWindows(call_back, result)
    return result


def find_character_data(data_list, character_name):
    """查找指定角色的数据"""
    result = []
    for item in data_list:
        if item.get('title') == character_name:
            result.append(item)

    if not result:
        return None
    elif len(result) > 1:
        return result[0]  # 直接返回第一个结果

    return result[0] if len(result) == 1 else result


def get_character_client_pid_hwnd(character_name: str):
    """获取角色客户端的PID和HWND"""
    pid_list = get_new_game_processes_list()
    char_data = find_character_data(pid_list, character_name)

    if not char_data:
        return None, None

    return char_data.get('pid'), char_data.get('hwnd')