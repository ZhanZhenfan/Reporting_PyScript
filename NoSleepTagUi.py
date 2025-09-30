import ctypes
import time
import threading
import tkinter as tk
import sys

# 防止睡眠相关常量
ES_CONTINUOUS    = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

# 键盘相关常量
VK_SHIFT = 0x10
VK_CONTROL = 0x11
KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP       = 0x0002

# 鼠标相关常量
MOUSEEVENTF_MOVE = 0x0001

def prevent_sleep():
    """
    每60秒调用 SetThreadExecutionState 防止系统睡眠和显示器关闭。
    """
    while True:
        result = ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
        )
        if result == 0:
            print("设置防睡眠状态失败。")
        else:
            print("已启用防睡眠功能。")
        time.sleep(60)

def simulate_shift():
    """
    模拟一次 Shift 键按下和释放。
    """
    ctypes.windll.user32.keybd_event(VK_SHIFT, 0, KEYEVENTF_EXTENDEDKEY, 0)
    time.sleep(0.1)
    ctypes.windll.user32.keybd_event(VK_SHIFT, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0)
    print("已模拟 Shift 键操作。")
    time.sleep(0.1)

def simulate_ctrl():
    """
    模拟一次 Ctrl 键按下和释放。
    """
    ctypes.windll.user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_EXTENDEDKEY, 0)
    time.sleep(0.1)
    ctypes.windll.user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0)
    print("已模拟 Ctrl 键操作。")
    time.sleep(0.1)

def simulate_mouse_movement():
    """
    模拟一次鼠标微移动（先移动，再复位）。
    """
    # 将鼠标移动5个像素（右下方向）
    ctypes.windll.user32.mouse_event(MOUSEEVENTF_MOVE, 5, 5, 0, 0)
    time.sleep(0.1)
    # 恢复原位置（左上方向移动5个像素）
    ctypes.windll.user32.mouse_event(MOUSEEVENTF_MOVE, -5, -5, 0, 0)
    print("已模拟鼠标移动。")
    time.sleep(0.1)

def simulate_user_activity():
    """
    每4分钟模拟多次操作以保持 Teams 在线状态。
    包括 Shift 键、Ctrl 键、鼠标移动等操作，确保活动足够丰富。
    """
    while True:
        print("开始一次用户活动模拟...")
        simulate_mouse_movement()
        simulate_ctrl()
        simulate_mouse_movement()
        simulate_ctrl()
        simulate_mouse_movement()
        simulate_ctrl()
        print("一次用户活动模拟完成。") 
        time.sleep(240)  # 每4分钟执行一次

def on_exit():
    """\\
    退出按钮的回调函数，结束程序。
    """
    print("退出程序。")
    sys.exit(0)

def create_exit_button():
    """
    创建一个简单的 Tkinter 窗口，包含一个退出按钮。
    """
    root = tk.Tk()
    root.title("退出程序")
    root.geometry("200x100")
    exit_button = tk.Button(root, text="退出", command=on_exit, width=10, height=2)
    exit_button.pack(expand=True)
    return root

if __name__ == "__main__":
    # 分别启动防睡眠和用户活动模拟的线程（均为守护线程）
    t1 = threading.Thread(target=prevent_sleep, daemon=True)
    t1.start()

    t2 = threading.Thread(target=simulate_user_activity, daemon=True)
    t2.start()

    # 创建包含退出按钮的 GUI 窗口
    root = create_exit_button()
    root.mainloop()
