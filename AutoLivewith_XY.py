import threading
import pyautogui
from win32api import GetSystemMetrics
import win32gui
import ctypes
from ctypes import windll
import time
import pygetwindow as gw

# 获取屏幕分辨率
screen_width = GetSystemMetrics(0)
screen_height = GetSystemMetrics(1)

def get_dpi():
    # Constants
    LOGPIXELSX = 88
    LOGPIXELSY = 90

    # Make the process DPI aware
    ctypes.windll.shcore.SetProcessDpiAwareness(1)

    # Get desktop DC
    dc = windll.user32.GetDC(0)

    # Get DPI settings
    dpi_x = windll.gdi32.GetDeviceCaps(dc, LOGPIXELSX)
    dpi_y = windll.gdi32.GetDeviceCaps(dc, LOGPIXELSY)

    # Release DC
    windll.user32.ReleaseDC(0, dc)

    return dpi_x, dpi_y

dpi_x, dpi_y = get_dpi()
current_thread = None
lock = threading.Lock()

def start_thread(func):
    global current_thread
    with lock:
        if current_thread and current_thread.is_alive():
            print('Stopping thread')
            current_thread.stop.set()  # 設置停止標記
            current_thread.join()  # 等待現有的線程完成

        stop = threading.Event()
        current_thread = threading.Thread(target=func, args=(stop,))  # 建立新線程，並傳入標記
        current_thread.stop = stop  # 為線程添加 stop 事件對象
        current_thread.start()

def enum_windows_proc(hwnd, resultList):
    if win32gui.IsWindowVisible(hwnd):
        title = win32gui.GetWindowText(hwnd)
        if "Discord" in title:
            title_parts = title.split('-')
            if len(title_parts) > 1 and "Discord" in title_parts[1].strip():
                resultList.append(hwnd)
    return True

def calculate_physical_position_and_click(logical_x, logical_y, dpi_current=96, dpi_target=dpi_x):
    physical_x, physical_y = calculate_physical_position(logical_x, logical_y, dpi_current, dpi_target)
    pyautogui.click(physical_x, physical_y)  # 在指定位置执行点击
    time.sleep(0.5)  # 等待1秒，让窗口有时间响应
    return physical_x, physical_y

windows = []
win32gui.EnumWindows(enum_windows_proc, windows)

# 获取 Discord 窗口的大小和位置
if not windows:  
    print("未找到 Discord 窗口。")
else:
    hwnd = windows[0]
    x0, y0, x1, y1 = win32gui.GetWindowRect(hwnd)
    discord_width = x1 - x0
    discord_height = y1 - y0
    print(f"Discord坐標：({x0}, {y0}), ({x1}, {y1})")
    print(f"Dpi：{dpi_x}")

    def calculate_physical_position(logical_x, logical_y, dpi_current=96, dpi_target=dpi_x):
        physical_x = logical_x * dpi_target / dpi_current
        physical_y = logical_y * dpi_target / dpi_current
        round_x = round(physical_x)
        round_y = round(physical_y)
        return round_x, round_y

    # Activating Discord window
    discord_windows = gw.getWindowsWithTitle('Discord')
    if discord_windows:
        for window in discord_windows:
            title_parts = window.title.split('-')
            if len(title_parts) > 1 and "Discord" in title_parts[1].strip():
                print(f"找到並激活窗口：{window.title}")
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(1)  # wait for the window to be activated

    share_x, share_y = calculate_physical_position_and_click(x0 + 162, y1 - 80)
    print(f"分享的坐標位置：({share_x}, {share_y})")
    time.sleep(0.5)

    widget_x = x0 + discord_width // 2
    widget_y = y0 + discord_height // 2

    picture_x, picture_y = calculate_physical_position_and_click(widget_x - 133, widget_y - 114)
    print(f"畫面的坐標位置：({picture_x}, {picture_y})")

    screen_x, screen_y = calculate_physical_position_and_click(widget_x - 123, widget_y - 29)
    print(f"螢幕1的坐標位置：({screen_x}, {screen_y})")

    go_live_x, go_live_y = calculate_physical_position_and_click(widget_x + 186, widget_y + 247)
    print(f"Go Live坐標位置：({go_live_x}, {go_live_y})")
