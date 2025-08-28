# -*- coding: utf-8 -*-
"""
Chuỗi thao tác khi nhấn SPACE:
1) Snip vùng cố định -> clipboard (ảnh)
2) Click CLICK1
3) Ctrl+V (paste)
4) Chờ 3s -> Click CLICK2
5) Đọc text nhiều dòng từ clipboard
6) Gộp thành 1 dòng & validate theo mẫu:
   "Người nhận:" <V1> "Địa chỉ:" <V2> "Điện thoại:" <V3> <V4> (V3=10 chữ số, V4=6 chữ số)
7) Nếu sai: dừng chu kỳ (in lỗi)
8) Nếu đúng: ghi 1 dòng vào out_text.txt: V1||||V2||||V3||||V4
9) In ra console đúng cụm text vừa ghi
10) Click CLICK3
11) Chờ 1s -> Click CLICK4
12) Phát âm thanh ready.mp3 báo sẵn sàng
-> Kết thúc 1 chu kỳ, tiếp tục chờ SPACE. ESC để thoát.
"""

import time
import re
import threading
import pyautogui as pag
import keyboard
import pyperclip
from PIL import Image
import os
import pygame  # dùng pygame để phát âm thanh

# ====== HẰNG SỐ CẦN CHỈNH ======
SNIP_REGION = (134, 1080, 1151-134, 1345-1080)
CLICK1 = (-500, 175)                  # Điểm click khi dán ảnh
CLICK2 = (-496, -14)                  # Điểm click sau 3s để lấy text vào clipboard
CLICK3 = (-975, -139)                  # Bước (10)
CLICK4 = (-903, -76)                  # Bước (11)

WAIT_BEFORE_CLICK2 = 3.0              # bước (4)
WAIT_BEFORE_CLICK4 = 1.0              # bước (11)

OUT_FILE = "out_text.txt"
READY_SOUND = "ready.mp3"             # bước (12)

ENABLE_IMAGE_CLIPBOARD = True         # dùng pywin32 để đặt ảnh vào clipboard (Windows)
# =================================

# --------- AUDIO (pygame) ----------
_audio_lock = threading.Lock()
_audio_ready = False
_ready_sound_obj = None  # cache Sound object để phát nhanh nhiều lần

def _init_audio_once():
    """
    Khởi tạo pygame.mixer đúng 1 lần (thread-safe).
    """
    global _audio_ready
    with _audio_lock:
        if _audio_ready:
            return
        try:
            # Có thể truyền thêm tham số nếu cần: freq=44100, size=-16, channels=2, buffer=512
            pygame.mixer.init()
            _audio_ready = True
        except Exception as e:
            print(f"[CẢNH BÁO] Không khởi tạo được âm thanh (pygame.mixer): {e}")
            _audio_ready = False

def _ensure_ready_sound_loaded():
    """
    Bảo đảm đã load sẵn file READY_SOUND vào _ready_sound_obj (nếu có).
    """
    global _ready_sound_obj
    if _ready_sound_obj is not None:
        return
    if not _audio_ready:
        return
    path = os.path.abspath(READY_SOUND)
    if not os.path.exists(path):
        print(f"[CẢNH BÁO] Không tìm thấy file âm thanh: {path}")
        return
    try:
        _ready_sound_obj = pygame.mixer.Sound(path)
    except Exception as e:
        print(f"[CẢNH BÁO] Không load được âm thanh: {e}")
        _ready_sound_obj = None

def play_ready_sound():
    """
    Phát READY_SOUND bằng pygame.mixer ở thread riêng để không chặn UI.
    """
    def _play():
        try:
            _init_audio_once()
            if not _audio_ready:
                return
            _ensure_ready_sound_loaded()
            if _ready_sound_obj is None:
                return
            _ready_sound_obj.play()  # không chặn
        except Exception as e:
            print(f"[CẢNH BÁO] Lỗi phát âm thanh (pygame): {e}")
    threading.Thread(target=_play, daemon=True).start()
# -----------------------------------

# --- Đưa ảnh vào clipboard (Windows) ---
def _set_clipboard_image_windows(image: Image.Image):
    try:
        import win32clipboard
        import win32con
    except Exception as e:
        raise RuntimeError("Thiếu pywin32. Cài: pip install pywin32") from e

    from io import BytesIO
    buf = BytesIO()
    image.convert("RGB").save(buf, "BMP")
    data = buf.getvalue()[14:]  # bỏ BITMAPFILEHEADER 14 bytes -> DIB
    buf.close()

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_DIB, data)
    finally:
        win32clipboard.CloseClipboard()

def snip_region_to_clipboard():
    img = pag.screenshot(region=SNIP_REGION)
    if ENABLE_IMAGE_CLIPBOARD:
        try:
            _set_clipboard_image_windows(img)
        except Exception as e:
            print(f"[CẢNH BÁO] Không thể đặt ảnh vào clipboard: {e}")
            img.save("snip_fallback.png")
            print("Đã lưu ảnh tạm: snip_fallback.png")
    else:
        img.save("snip_fallback.png")
        print("Đã lưu ảnh tạm: snip_fallback.png")

def safe_get_clipboard_text(max_tries=5, delay=0.2):
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
    s = s.replace("\r", " ").replace("\n", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def parse_block_and_validate(s: str):
    pat = r'^Người nhận:\s*(.+?)\s+Địa chỉ:\s*(.+?)\s+Điện thoại:\s*(\d{10})\s+(\d{6})\s*$'
    m = re.match(pat, s)
    if not m:
        return None
    return m.group(1), m.group(2), m.group(3), m.group(4)

def append_output(v1, v2, v3, v4, out_file=OUT_FILE):
    line = f"{v1}||||{v2}||||{v3}||||{v4}\n"
    with open(out_file, "a", encoding="utf-8") as f:
        f.write(line)
    return line.rstrip("\n")

def run_workflow_once():
    # (1)
    snip_region_to_clipboard()
    time.sleep(0.05)

    # (2)
    pag.click(*CLICK1)
    time.sleep(0.05)

    # (3)
    pag.hotkey("ctrl", "v")

    # (4)
    time.sleep(WAIT_BEFORE_CLICK2)
    pag.click(*CLICK2)

    # (5)
    raw_text = safe_get_clipboard_text()
    if not raw_text:
        print("[LỖI] Không đọc được text từ clipboard. Dừng chu kỳ.")
        return

    # (6)
    one_line = normalize_spaces(raw_text)
    parsed = parse_block_and_validate(one_line)
    if not parsed:
        print("[LỖI] Text không đúng định dạng yêu cầu.")
        print("→ Nội dung 1 dòng:", one_line)
        print("Kỳ vọng: 'Người nhận: <V1> Địa chỉ: <V2> Điện thoại: <V3> <V4>' (V3=10 số, V4=6 số)")
        return

    v1, v2, v3, v4 = parsed

    # (8) + (9)
    saved_line = append_output(v1, v2, v3, v4)
    print("[OK] Đã ghi:", saved_line)

    # (10)
    pag.click(*CLICK3)

    # (11)
    time.sleep(WAIT_BEFORE_CLICK4)
    pag.click(*CLICK4)

    # (12)
    play_ready_sound()

def main():
    print(">>> Sẵn sàng. Nhấn SPACE để chạy chu kỳ. Nhấn ESC để thoát.")
    def on_space():
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
    finally:
        # đóng mixer gọn gàng nếu đã khởi tạo
        try:
            if pygame.mixer.get_init():
                pygame.mixer.quit()
        except Exception:
            pass
