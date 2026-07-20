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

    Lúc build, Chromium được cài với PLAYWRIGHT_BROWSERS_PATH=0 nên nó nằm ngay
    trong thư mục package playwright thay vì cache riêng của máy. Đặt lại đúng
    biến môi trường đó khi chạy để Playwright tìm thấy trình duyệt đã đóng kèm,
    thay vì đi tải mới về (máy người dùng không có, và có thể không có mạng).
    """
    if getattr(sys, "frozen", False):
        os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", "0")
        # Bỏ qua kiểm tra phiên bản driver — thư mục gói là chỉ đọc
        os.environ.setdefault("PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD", "1")
