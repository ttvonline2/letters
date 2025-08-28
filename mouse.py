import pyautogui
import time

while True:
    x, y = pyautogui.position()  # Lấy tọa độ chuột
    print(f"Vị trí chuột: ({x}, {y})")
    time.sleep(0.1)  # Dừng 0.1 giây
