# -*- coding: utf-8 -*-
"""
Workflow:
1) Nhấn SPACE -> chụp (snip) 1 vùng cố định và đưa ảnh vào clipboard
2) Click chuột trái vào điểm cố định 1 (CLICK1)
3) Thực hiện paste (Ctrl+V)
4) Chờ 3s -> Click chuột trái vào điểm cố định 2 (CLICK2)
5) Lúc này clipboard chứa block text nhiều dòng
6) Gộp thành 1 dòng, kiểm tra định dạng:
   "Người nhận:" + <V1> + "Địa chỉ:" + <V2> + "Điện thoại:" + <V3> + <V4>
   (V3 = 10 chữ số, V4 = 6 chữ số)
7) Sai định dạng -> dừng, in thông báo lỗi
8) Đúng -> ghi ra out_text.txt:  V1||||V2||||V3||||V4
"""

import time
import re
import sys
import threading

import pyautogui as pag
import keyboard
import pyperclip
from PIL import Image

# ====== CÁC HẰNG SỐ CẦN CHỈNH SỬA ======
# Vùng snip: (left, top, width, height) – theo pixel màn hình
SNIP_REGION = (192, 1074, 1023-192, 1282-1074)   # TODO: sửa theo máy bạn

# Điểm click 1 (x, y) – nơi bạn muốn dán ảnh sau bước (1)
CLICK1 = (-600, 400)                 # TODO: sửa theo máy bạn
 
# Điểm click 2 (x, y) – nơi app sẽ để text vào clipboard sau 3 giây
CLICK2 = (-463, 14)                 # TODO: sửa theo máy bạn

# File output
OUT_FILE = "out_text.txt "

# Thời gian chờ trước khi click 2 (giây)
WAIT_BEFORE_CLICK2 = 3.0

# Có dán ảnh vào clipboard hay không (Windows + pywin32)
ENABLE_IMAGE_CLIPBOARD = True
# =======================================


# --- Hỗ trợ: đưa ảnh vào clipboard trên Windows (DIB) ---
def _set_clipboard_image_windows(image: Image.Image):
    """
    Đưa ảnh PIL vào clipboard dạng DIB (yêu cầu pywin32).
    """
    try:
        import win32clipboard
        import win32con
    except Exception as e:
        raise RuntimeError("Thiếu pywin32 (win32clipboard). Cài: pip install pywin32") from e

    # Chuyển ảnh sang BMP (BGRX) và bỏ header 14 byte (BMP header) -> còn DIB
    from io import BytesIO
    output = BytesIO()
    bmp = image.convert("RGB")
    bmp.save(output, "BMP")
    data = output.getvalue()[14:]  # bỏ BITMAPFILEHEADER
    output.close()

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_DIB, data)
    finally:
        win32clipboard.CloseClipboard()


def snip_region_to_clipboard():
    """
    Chụp vùng SNIP_REGION. Nếu ENABLE_IMAGE_CLIPBOARD=True (Windows),
    đưa ảnh vào clipboard để Ctrl+V có thể dán.
    """
    # Chụp màn hình vùng cố định
    img = pag.screenshot(region=SNIP_REGION)

    if ENABLE_IMAGE_CLIPBOARD:
        try:
            _set_clipboard_image_windows(img)
        except Exception as e:
            print(f"[CẢNH BÁO] Không thể đặt ảnh vào clipboard: {e}")
            # fallback: lưu file tạm (người dùng có thể tự paste file nếu cần)
            img.save("snip_fallback.png")
            print("Ảnh đã lưu: snip_fallback.png")
    else:
        # Nếu không dùng clipboard ảnh, vẫn lưu để tham khảo
        img.save("snip_fallback.png")
        print("Ảnh đã lưu: snip_fallback.png")


def safe_get_clipboard_text(max_tries=5, delay=0.2):
    """
    Đọc text từ clipboard, thử vài lần cho chắc.
    """
    for _ in range(max_tries):
        try:
            txt = pyperclip.paste()
            if isinstance(txt, str) and txt.strip():
                return txt
        except Exception:
            pass
        time.sleep(delay)
    return ""


def normalize_spaces(s: str) -> str:
    """
    - Thay mọi loại xuống dòng bằng khoảng trắng
    - Rút gọn khoảng trắng liên tiếp thành 1 khoảng trắng
    - Trim hai đầu
    """
    s = s.replace("\r", " ").replace("\n", " ")
    s = re.sub(r"\s+", " ", s, flags=re.UNICODE)
    return s.strip()


def parse_block_and_validate(s: str):
    """
    Kỳ vọng chuỗi 1 dòng dạng:
    Người nhận: <V1> Địa chỉ: <V2> Điện thoại: <V3> <V4>

    Trả về tuple (V1, V2, V3, V4) nếu hợp lệ, ngược lại None.
    """
    # Cho phép linh hoạt khoảng trắng
    pattern = r'^Người nhận:\s*(.+?)\s+Địa chỉ:\s*(.+?)\s+Điện thoại:\s*(\d{10})\s+(\d{6})\s*$'
    m = re.match(pattern, s, flags=re.UNICODE)
    if not m:
        return None
    V1, V2, V3, V4 = m.group(1), m.group(2), m.group(3), m.group(4)
    return V1, V2, V3, V4


def append_output(v1, v2, v3, v4, out_file=OUT_FILE):
    line = f"{v1}||||{v2}||||{v3}||||{v4}\n"
    with open(out_file, "a", encoding="utf-8") as f:
        f.write(line)
    print(f"[OK] Đã ghi 1 dòng vào {out_file}.")


def run_workflow_once():
    # (1) Snip vùng cố định -> clipboard
    snip_region_to_clipboard()
    time.sleep(0.05)

    # (2) Click chuột trái vào điểm CLICK1
    pag.click(x=CLICK1[0], y=CLICK1[1])
    time.sleep(0.05)

    # (3) Paste (Ctrl+V)
    pag.hotkey("ctrl", "v")

    # (4) Chờ 3s rồi click CLICK2
    time.sleep(WAIT_BEFORE_CLICK2)
    pag.click(x=CLICK2[0], y=CLICK2[1])

    # (5) Lúc này clipboard có text nhiều dòng -> đọc
    raw_text = safe_get_clipboard_text()
    if not raw_text:
        print("[LỖI] Không đọc được text từ clipboard. Dừng chương trình.")
        return

    # (6) Gộp thành 1 dòng & kiểm tra định dạng
    one_line = normalize_spaces(raw_text)
    parsed = parse_block_and_validate(one_line)
    if not parsed:
        print("[LỖI] Text không đúng định dạng yêu cầu.")
        print("Nội dung 1 dòng đã đọc:")
        print(one_line)
        print("Kỳ vọng: 'Người nhận: <V1> Địa chỉ: <V2> Điện thoại: <V3> <V4>' " 
              "(V3=10 chữ số, V4=6 chữ số).")
        return

    v1, v2, v3, v4 = parsed

    # (8) Ghi file
    append_output(v1, v2, v3, v4)


def main():
    print(">>> Script đã chạy. Nhấn SPACE để thực thi chuỗi bước. Nhấn ESC để thoát.")
    # Dùng hotkey toàn cục: SPACE = chạy 1 lần quy trình, ESC = thoát
    def on_space():
        # chạy trong thread để không block listener
        t = threading.Thread(target=run_workflow_once, daemon=True)
        t.start()

    keyboard.add_hotkey("space", on_space)
    keyboard.wait("esc")
    print(">>> Thoát.")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
