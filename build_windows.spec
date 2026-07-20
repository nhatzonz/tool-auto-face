# -*- mode: python ; coding: utf-8 -*-
"""
Cấu hình đóng gói tool thành ứng dụng Windows bằng PyInstaller.

Cách build (chạy trên máy Windows):
    pip install -r requirements.txt
    set PLAYWRIGHT_BROWSERS_PATH=%CD%\\ms-playwright
    playwright install chromium
    pyinstaller build_windows.spec

Kết quả: thư mục dist\\ToolDangBaiFacebook\\ — đưa cả thư mục này cho người
dùng, họ bấm vào ToolDangBaiFacebook.exe là chạy, không cần cài Python hay
trình duyệt gì thêm.

Vì sao tải Chromium vào thư mục ms-playwright ngay trong project: mặc định
Playwright để trình duyệt ở cache riêng của máy
(%USERPROFILE%\\AppData\\Local\\ms-playwright), nằm ngoài project nên
PyInstaller không thấy để đóng kèm.

Từng thử PLAYWRIGHT_BROWSERS_PATH=0 (đưa Chromium vào trong package playwright)
rồi để collect_all() tự gom, nhưng KHÔNG ĂN: Chromium nằm trong thư mục con
'.local-browsers' bắt đầu bằng dấu chấm, mà collect_all() bỏ qua thư mục ẩn.
Build vẫn xanh nhưng gói ra thiếu trình duyệt, chỉ nặng 51MB. Nên giờ chỉ
định thẳng đường dẫn ở dòng datas bên dưới cho chắc chắn.
"""
import os
from PyInstaller.utils.hooks import collect_all

# Gom package playwright: driver Node.js và các file .js đi kèm
pw_datas, pw_binaries, pw_hiddenimports = collect_all("playwright")

# Đóng kèm Chromium. Dừng hẳn nếu thiếu, thay vì build ra gói hỏng mà vẫn báo
# thành công — lỗi kiểu đó chỉ lộ ra khi đã đưa tới tay người dùng.
BROWSERS_DIR = os.path.abspath("ms-playwright")
if not os.path.isdir(BROWSERS_DIR):
    raise SystemExit(
        f"Không tìm thấy '{BROWSERS_DIR}'.\n"
        f"Chạy 2 lệnh sau trước khi build:\n"
        f"    set PLAYWRIGHT_BROWSERS_PATH={BROWSERS_DIR}\n"
        f"    playwright install chromium"
    )
pw_datas += [(BROWSERS_DIR, "ms-playwright")]

a = Analysis(
    ["gui.py"],
    pathex=[],
    binaries=pw_binaries,
    datas=pw_datas,
    hiddenimports=pw_hiddenimports + [
        # openpyxl nạp các module này động, PyInstaller không tự dò ra
        "openpyxl.cell._writer",
    ],
    hookspath=[],
    runtime_hooks=[],
    # Bớt dung lượng: tool không dùng mấy thư viện khoa học nặng này
    excludes=["numpy", "pandas", "matplotlib", "scipy", "PIL", "pytest"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    exclude_binaries=True,
    name="ToolDangBaiFacebook",
    debug=False,
    strip=False,
    upx=False,
    # console=False: giấu cửa sổ dòng lệnh đen, chỉ hiện giao diện.
    # Muốn xem lỗi chi tiết lúc gỡ rối thì tạm đổi thành True rồi build lại.
    console=False,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="ToolDangBaiFacebook",
)
