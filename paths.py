"""
Xác định thư mục làm việc, chạy đúng ở cả 2 chế độ: chạy từ mã nguồn và chạy
từ file .exe đã đóng gói bằng PyInstaller.

Vì sao cần file này: khi đóng gói kiểu onefile, PyInstaller giải nén toàn bộ
chương trình vào một thư mục tạm rồi XÓA nó lúc thoát. Nếu vẫn lấy đường dẫn
theo __file__ như code thường, config.json và profile đăng nhập sẽ được ghi
vào thư mục tạm đó và bay sạch sau mỗi lần chạy — người dùng phải đăng nhập
Facebook lại từ đầu mỗi lần mở tool.

Nên tách làm 2 khái niệm:
  BUNDLE_DIR : nơi chứa tài nguyên đóng kèm, CHỈ ĐỌC, mất khi thoát
  DATA_DIR   : nơi ghi dữ liệu người dùng, nằm cạnh file .exe, còn mãi
"""
import os
import sys

if getattr(sys, "frozen", False):
    # Đang chạy từ file .exe do PyInstaller tạo
    BUNDLE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    DATA_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    # Đang chạy từ mã nguồn (python gui.py)
    BUNDLE_DIR = DATA_DIR = os.path.dirname(os.path.abspath(__file__))


def data_path(*parts):
    """Đường dẫn tới file/thư mục dữ liệu người dùng (ghi được, giữ lâu dài)."""
    return os.path.join(DATA_DIR, *parts)


def bundle_path(*parts):
    """Đường dẫn tới tài nguyên đóng kèm trong gói (chỉ đọc)."""
    return os.path.join(BUNDLE_DIR, *parts)


def setup_playwright_env():
    """Trỏ Playwright vào Chromium đóng kèm trong gói.

    Máy người dùng không có Chromium trong cache và có thể không có mạng để
    tải, nên phải chỉ rõ trình duyệt nằm ngay trong gói (thư mục ms-playwright,
    do build_windows.spec đóng kèm).
    """
    if not getattr(sys, "frozen", False):
        return          # chạy từ mã nguồn: dùng cache Playwright bình thường

    browsers = bundle_path("ms-playwright")
    if os.path.isdir(browsers):
        os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", browsers)
    # Thư mục gói là chỉ đọc, có tải cũng không ghi được vào đó
    os.environ.setdefault("PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD", "1")
