import os
import re
from glob import glob
from openpyxl import Workbook

TXT_PATH = "out_text.txt"
EXCEL_DIR = "excel"
BASENAME = "output"          # output_<n>.xlsx
DELIM = "||||"

def ensure_excel_dir(path: str):
    os.makedirs(path, exist_ok=True)

def get_next_index(excel_dir: str, basename: str) -> int:
    """
    Tìm số lớn nhất trong các file dạng basename_<n>.xlsx
    và trả về n+1 (chỉ số cho file tiếp theo).
    """
    pattern = os.path.join(excel_dir, f"{basename}_*.xlsx")
    indices = []
    for fp in glob(pattern):
        name = os.path.basename(fp)
        m = re.match(fr"^{re.escape(basename)}_(\d+)\.xlsx$", name)
        if m:
            indices.append(int(m.group(1)))
    return (max(indices) + 1) if indices else 1

def process_txt_to_excel(txt_path: str, excel_dir: str, basename: str, delim: str):
    # Đảm bảo thư mục tồn tại
    ensure_excel_dir(excel_dir)

    # Đọc toàn bộ dòng trong file txt
    try:
        with open(txt_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
    except FileNotFoundError:
        print(f"Không tìm thấy file: {txt_path}")
        return

    lines = [ln.strip() for ln in lines if ln.strip()]
    if not lines:
        print("Không có dữ liệu để xuất.")
        return

    # Xác định tên file Excel mới
    idx = get_next_index(excel_dir, basename)
    filename = f"{basename}_{idx}.xlsx"
    filepath = os.path.join(excel_dir, filename)

    # Tạo workbook và ghi dữ liệu
    wb = Workbook()
    ws = wb.active
    for line in lines:
        values = [part.strip() for part in line.split(delim)]
        ws.append(values)

    wb.save(filepath)

    # Xóa sạch nội dung file txt
    with open(txt_path, "w", encoding="utf-8") as f:
        f.truncate(0)

    print(f"Đã tạo file: {filepath}")
    print(f"Đã xóa nội dung trong {txt_path}")

if __name__ == "__main__":
    process_txt_to_excel(TXT_PATH, EXCEL_DIR, BASENAME, DELIM)
