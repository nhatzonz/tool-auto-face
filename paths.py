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
import io
import os
import sys

TEN_UNG_DUNG = "ToolDangBaiFacebook"


def _thu_muc_du_lieu_mac():
    """Chỗ ghi dữ liệu cho bản .app trên macOS.

    Không dùng được thư mục chứa file thực thi như bản Windows: trong một gói
    .app thì đó là ToolDangBaiFacebook.app/Contents/MacOS/ — ghi config và
    profile Chrome vào ruột app thì cập nhật app là mất sạch dữ liệu, mà macOS
    còn chặn ghi vào app đã ký số. Chỗ đúng theo quy ước của macOS là
    ~/Library/Application Support/<tên app>/.
    """
    goc = os.path.expanduser("~/Library/Application Support")
    return os.path.join(goc, TEN_UNG_DUNG)


if getattr(sys, "frozen", False):
    BUNDLE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    thu_muc_exe = os.path.dirname(os.path.abspath(sys.executable))
    if sys.platform == "darwin" and ".app/Contents/MacOS" in thu_muc_exe:
        DATA_DIR = _thu_muc_du_lieu_mac()
        os.makedirs(DATA_DIR, exist_ok=True)
    else:
        # Bản Windows (và bản macOS không đóng thành .app): dữ liệu nằm cạnh
        # file chạy, đúng như hướng dẫn người dùng đang mô tả.
        DATA_DIR = thu_muc_exe
else:
    # Đang chạy từ mã nguồn (python gui.py)
    BUNDLE_DIR = DATA_DIR = os.path.dirname(os.path.abspath(__file__))


def data_path(*parts):
    """Đường dẫn tới file/thư mục dữ liệu người dùng (ghi được, giữ lâu dài)."""
    return os.path.join(DATA_DIR, *parts)


def bundle_path(*parts):
    """Đường dẫn tới tài nguyên đóng kèm trong gói (chỉ đọc)."""
    return os.path.join(BUNDLE_DIR, *parts)


def guard_missing_stdout():
    """Chặn lỗi khi code gọi print() trong bản .exe không có cửa sổ dòng lệnh.

    Build với console=False thì Windows không cấp stdout/stderr, PyInstaller
    đặt chúng thành None. Khi đó mỗi lệnh print() ném AttributeError
    ('NoneType' object has no attribute 'write') — đủ để làm hỏng thao tác
    đang chạy, dù nội dung in ra chẳng ai nhìn thấy.

    Code này vốn viết để chạy trong terminal nên rải print() ở nhiều nhánh
    (menu chọn tài khoản, tạo file mẫu...). Thay vì đi sửa từng chỗ và vẫn có
    thể sót, hứng luôn ở đây: nuốt mọi thứ ghi ra stdout/stderr.

    Log mà người dùng cần xem đi đường khác — qua LOG_FN vào ô log của giao
    diện — nên không mất gì.
    """
    if not getattr(sys, "frozen", False):
        return

    class _Bo(io.IOBase):
        def write(self, _):
            return 0

        def flush(self):
            pass

    if sys.stdout is None:
        sys.stdout = _Bo()
    if sys.stderr is None:
        sys.stderr = _Bo()


def setup_playwright_env():
    """Chuẩn bị môi trường Playwright cho bản đã đóng gói.

    Tool dùng Google Chrome cài sẵn trên máy (channel='chrome') chứ không dùng
    Chromium của Playwright, nên không cần trỏ tới thư mục trình duyệt nào.
    Chỉ cần chặn Playwright tự đi tải: thư mục gói là chỉ đọc, tải cũng không
    ghi được, mà chờ tải xong lại làm tool đứng hình vài phút.
    """
    if getattr(sys, "frozen", False):
        os.environ.setdefault("PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD", "1")
