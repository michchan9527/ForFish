import os
import sys
import time
import cv2
import keyboard
import numpy as np
import pyautogui
import pygetwindow as gw
from collections import defaultdict
from pynput.mouse import Controller, Button
import win32gui
from PIL import ImageGrab
import configparser
import win32gui, win32com.client

def get_application_path():
    # 檢查是否在打包的應用程式中運行
    if getattr(sys, 'frozen', False):
        # 如果是，使用 _MEIPASS 路徑
        return sys._MEIPASS
    else:
        # 如果不是，使用常規的腳本路徑
        return os.path.dirname(os.path.abspath(__file__))
    
def load_config():
    config = configparser.ConfigParser()
    filename = os.path.join(get_application_path(), 'config.ini')
    config.read(filename, encoding='utf-8')
    return config

def get_value(section, key):
    config = load_config()
    try:
        value = config.get(section, key)
    except configparser.NoOptionError:
        print(f"找不到 {section} 區段中的 {key} 選項")
        value = None
    return value

def set_value(section, key, value):
    config = load_config()
    config.set(section, key, value)
    save_config(config)

def save_config(config, filename='config.ini'):
    with open(filename, 'w') as configfile:
        config.write(configfile)

def get_debug():
    return get_value('Setting', 'debug') == 'True'

def set_debug(value):
    set_value('Setting', 'debug', str(value))

def callback(hwnd, extra):
    if 'minecraft' in win32gui.GetWindowText(hwnd).lower() and '-' not in win32gui.GetWindowText(hwnd).lower():
        extra.append(hwnd)

def find_minecraft_window():
    windows = []
    win32gui.EnumWindows(callback, windows)
    return windows[0] if windows else None

def capture_minecraft_window():
    debug = get_value('Setting', 'debug')
    global width, height
    hwnd = find_minecraft_window()

    if hwnd is None:
        print('未找到指定的 Minecraft 視窗')
        return None, None, None  # 或者你可以返回其他的預設值

    shell = win32com.client.Dispatch("WScript.Shell")
    shell.SendKeys('%')
    # 將 Minecraft 視窗啟動（切換到前臺）
    win32gui.SetForegroundWindow(hwnd)

    # 獲取窗口的位置和大小
    left, top, right, bot = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bot - top

    # 捕獲 Minecraft 視窗的內容
    screenshot = pyautogui.screenshot(region=(left, top, width, height))
    screenshot_np = np.array(screenshot)
    screenshot_bgr = cv2.cvtColor(screenshot_np, cv2.COLOR_RGB2BGR)
    # img = ImageGrab.grab(bbox = (left, top, right, bot))
    # img_np = np.array(img)
    # screenshot_bgr = cv2.cvtColor(img_np, cv2.COLOR_BGR2RGB)

    # 保存截圖到檔，使用當前時間作為檔案名
    if debug == 'True':
        save_path = os.path.join(get_application_path(), f'{time.strftime("%Y%m%d%H%M%S")}.png')
        cv2.imwrite(save_path, screenshot_bgr)

    return screenshot_bgr, left, top

def find_centers(screenshot_bgr):
    # 定義我們要搜尋的顏色範圍，這裡是黑色
    lower = np.array([0, 0, 0])
    upper = np.array([1, 1, 1])

    # 建立一個遮罩，將在指定範圍內的圖元設為白色，其餘為黑色
    mask = cv2.inRange(screenshot_bgr, lower, upper)

    # 找到遮罩中的輪廓
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    centers_dict = {}
    centers_list = []
    for cnt in contours:
        # 使用邊界矩形獲取輪廓的 x, y, w, h
        x, y, w, h = cv2.boundingRect(cnt)

        # 檢查該輪廓是否形成一個長方形
        if w > 20 and h > 20:
            # 計算長方形的中心點並保存
            center_x = x + w // 2
            center_y = y + h // 2

            centers_dict[(center_x, center_y)] = (w, h)
            centers_list.append((center_x, center_y))

            # 在中心點位置添加一個紅色的圓圈
            cv2.circle(screenshot_bgr, (center_x, center_y), 5, (0, 0, 255), -1)

    # 按 x, y 座標排序
    centers_list.sort(key=lambda c: (c[0], c[1]))
    
    # 保存標記過的截圖到檔，使用當前時間作為檔案名
    # save_path = os.path.join(get_application_path(), f'{time.strftime("%Y%m%d%H%M%S")}.png')
    # cv2.imwrite(save_path, screenshot_bgr)

    return centers_dict, centers_list

def select_target_centers(centers_list):
    global width, height
    # 获取屏幕的中心坐标
    screen_center = (width // 2, height // 2)
    
    # 移除等於屏幕中心的点
    centers_list = [center for center in centers_list if center != screen_center]
    # 找出具有相同間距的中心
    same_distance_centers = []
    prev_distance = None
    prev_center = None
    for i in range(1, len(centers_list)):
        distance = centers_list[i][1] - centers_list[i - 1][1]
        if distance == prev_distance:
            if prev_center not in same_distance_centers:
                same_distance_centers.append(prev_center)
            if centers_list[i - 1] not in same_distance_centers:
                same_distance_centers.append(centers_list[i - 1])
            if centers_list[i] not in same_distance_centers:
                same_distance_centers.append(centers_list[i])
        prev_distance = distance
        prev_center = centers_list[i - 1]

    # 如果沒有相同間距的中心,則尋找具有3個和4個中心的組
    grouped_centers_3_4 = []
    if not same_distance_centers:
        # 按 y 座標分組
        group_centers = defaultdict(list)
        for center in centers_list:
            group_centers[center[1]].append(center)

        # 搜索具有 3 或 4 個中心點的組
        for group in group_centers.values():
            if len(group) in [3, 4]:
                grouped_centers_3_4.extend(group)

    # 如果還沒有目標中心點，尋找具有3個x座標相同的組
    grouped_centers_3_x = []
    if not same_distance_centers and not grouped_centers_3_4:
        # 按 x 座標分組
        group_centers = defaultdict(list)
        for center in centers_list:
            group_centers[center[0]].append(center)

        # 搜索具有 3 個中心點的組
        for group in group_centers.values():
            if len(group) == 3:
                grouped_centers_3_x.extend(group)

    # 如果還沒有目標中心點，尋找只有一組坐標的組
    single_centers = []
    if not same_distance_centers and not grouped_centers_3_4 and not grouped_centers_3_x:
        # 按 y 座標分組
        group_centers = defaultdict(list)
        for center in centers_list:
            group_centers[center[1]].append(center)

        # 搜索只有 1 個中心點的組
        for group in group_centers.values():
            if len(group) == 1:
                single_centers.extend(group)

    return same_distance_centers, grouped_centers_3_4, grouped_centers_3_x, single_centers

def click_target_centers(screenshot_bgr, same_distance_centers, grouped_centers_3_4, grouped_centers_3_x, single_centers, win_x, win_y):
    debug = get_value('Setting', 'debug')
    # 獲取中心點的寬度和高度以及排序後的中心點列表
    centers_dict, centers_list = find_centers(screenshot_bgr)
    mouse = Controller()

    def save_screenshot(target_centers, target_type):
        # Create a copy of the screenshot
        screenshot_copy = screenshot_bgr.copy()
        # Draw circles around target centers
        for center in target_centers:
            cv2.circle(screenshot_copy, (center[0], center[1]), 5, (0, 255, 0), -1)
        # Save the marked screenshot with the target type and current time in the filename
        if debug == 'True':
            save_path = os.path.join(get_application_path(), f'{target_type}_{time.strftime("%Y%m%d%H%M%S")}.png')
            cv2.imwrite(save_path, screenshot_copy)

    # 根據查找結果選擇行動
    if same_distance_centers or grouped_centers_3_4:
        # print("相同距離中心: ", same_distance_centers)
        # print("以中心分組_3_4: ", grouped_centers_3_4)
        target_centers = same_distance_centers if same_distance_centers else grouped_centers_3_4
        target_type = "same_distance" if same_distance_centers else "grouped_3_4"
        save_screenshot(target_centers, target_type)
        # 在點擊螢幕之前，我們需要將 Minecraft 視窗的位置添加到中心點的座標上
        if target_centers and len(target_centers) == 3 or target_centers and len(target_centers) < 7:
            print("單擊相同距離中心...")
            second_smallest_y_center = sorted(target_centers, key=lambda c: c[1])[1]  # 選擇第二個中心點（索引為1）
            time.sleep(0.3)
            pyautogui.click(win_x + second_smallest_y_center[0], win_y + second_smallest_y_center[1])
        elif target_centers and len(target_centers) == 7:
            print("單擊以中心點分組_3_4...")
            second_smallest_y_center = sorted(target_centers, key=lambda c: c[1])[1]  # 選擇第二個中心點（索引為1）
            x, y = win_x + second_smallest_y_center[0], win_y + second_smallest_y_center[1]
            w, h = centers_dict[second_smallest_y_center]
            time.sleep(0.3)
            pyautogui.moveTo(x, y)
            mouse.click(Button.left, 1)
            time.sleep(0.3)
            # 將滑鼠移動到新的位置
            pyautogui.moveTo(x, y + h + 5)
    elif grouped_centers_3_x:
        print("單擊以中心點分組_3_x...")
        # print("以中心分組_3_x: ", grouped_centers_3_x)
        target_centers = grouped_centers_3_x
        save_screenshot(target_centers, "grouped_3_x")
        smallest_y_center = sorted(target_centers, key=lambda c: c[1])[0]  # 選擇最小的中心點
        second_smallest_y_center = sorted(target_centers, key=lambda c: c[1])[1]  # 選擇第二個中心點（索引為1）
        time.sleep(0.3)
        pyautogui.click(win_x + smallest_y_center[0], win_y + smallest_y_center[1])
        time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'a')  # 模擬按下 Ctrl+A
        time.sleep(0.3)
        pyautogui.typewrite(get_value('Setting', 'host'))  # 類比鍵盤輸入
        time.sleep(0.3)  # 等待1秒
        pyautogui.moveTo(win_x + second_smallest_y_center[0], win_y + second_smallest_y_center[1])
        mouse.click(Button.left, 1)
    elif single_centers:
        print("單擊單個中心...")
        # print("單中心: ", single_centers)
        # 針對只有一組坐標的情況，你可能需要修改下面的代碼以適應你的具體需求
        # 下面的代碼只是一個例子，我假設你只想點擊這個中心點
        target_center = single_centers[0]
        target_centers = single_centers
        save_screenshot(target_centers, "single")
        # print(target_center)
        time.sleep(0.3)
        pyautogui.click(win_x + target_center[0], win_y + target_center[1])

def main():
    print('等待使用者按下 {} 鍵開始...'.format(get_value('Setting', 'hot_key')))
    keyboard.wait(get_value('Setting', 'hot_key'))
    # 初始化嘗試次數
    attempts = 0
    while attempts < 4:  # 設置最大嘗試次數為3
        # 獲取截圖和窗口位置
        screenshot_bgr, win_x, win_y = capture_minecraft_window()
        if screenshot_bgr is None:  # 檢查返回的截圖是否為 None
            print('未能獲取 Minecraft 截圖')
            break  # 如果截圖為 None，則跳過此次迴圈
        centers_dict, centers_list = find_centers(screenshot_bgr)
        same_distance_centers, grouped_centers_3_4, grouped_centers_3_x, single_centers = select_target_centers(centers_list)
        is_first_time_single_centers = True
        # print("相同距離中心: ", same_distance_centers)
        # print("以中心分組_3_4: ", grouped_centers_3_4)
        # print("以中心分組_3_x: ", grouped_centers_3_x)
        # print("單中心: ", single_centers)

        # 如果找到了same_distance_centers，則執行2次select_target_centers和click_target_centers
        if same_distance_centers:
            time.sleep(0.3)
            for _ in range(3):
                # 獲取截圖和窗口位置
                screenshot_bgr, win_x, win_y = capture_minecraft_window()
                centers_dict, centers_list = find_centers(screenshot_bgr)
                same_distance_centers, grouped_centers_3_4, grouped_centers_3_x, single_centers = select_target_centers(centers_list)
                click_target_centers(screenshot_bgr, same_distance_centers, grouped_centers_3_4, grouped_centers_3_x, single_centers, win_x, win_y)

        # 如果找到了grouped_centers_3_4，則執行1次select_target_centers和click_target_centers
        elif grouped_centers_3_4:
            time.sleep(0.3)
            for _ in range(2):
                # 獲取截圖和窗口位置
                screenshot_bgr, win_x, win_y = capture_minecraft_window()
                centers_dict, centers_list = find_centers(screenshot_bgr)
                same_distance_centers, grouped_centers_3_4, grouped_centers_3_x, single_centers = select_target_centers(centers_list)
                click_target_centers(screenshot_bgr, same_distance_centers, grouped_centers_3_4, grouped_centers_3_x, single_centers, win_x, win_y)

        elif grouped_centers_3_x:
            time.sleep(0.3)
            # 獲取截圖和窗口位置
            screenshot_bgr, win_x, win_y = capture_minecraft_window()
            centers_dict, centers_list = find_centers(screenshot_bgr)
            same_distance_centers, grouped_centers_3_4, grouped_centers_3_x, single_centers = select_target_centers(centers_list)
            click_target_centers(screenshot_bgr, same_distance_centers, grouped_centers_3_4, grouped_centers_3_x, single_centers, win_x, win_y)
        
        elif single_centers:
            time.sleep(0.3)
            for _ in range(3):
                # 獲取截圖和窗口位置
                screenshot_bgr, win_x, win_y = capture_minecraft_window()
                centers_dict, centers_list = find_centers(screenshot_bgr)
                
                # 根據是否是第一次出現single_centers來決定是否執行select_target_centers(centers)
                if not is_first_time_single_centers:
                    same_distance_centers, grouped_centers_3_4, grouped_centers_3_x, single_centers = select_target_centers(centers_list)
                    
                click_target_centers(screenshot_bgr, same_distance_centers, grouped_centers_3_4, grouped_centers_3_x, single_centers, win_x, win_y)

                # 如果是第一次出現single_centers，那麼在完成操作後，修改標誌變數以便下次執行select_target_centers(centers)
                if is_first_time_single_centers:
                    is_first_time_single_centers = False

        # 檢查是否找到了目標中心
        if same_distance_centers or grouped_centers_3_4 or grouped_centers_3_x or single_centers:
            break  # 如果找到了，則跳出迴圈

        # 如果沒有找到，嘗試次數加1，然後繼續迴圈
        attempts += 1
        time.sleep(1)  # 每次嘗試之間的間隔時間

main()