# -*- mode: python ; coding: utf-8 -*-
"""
Cấu hình đóng gói tool thành ứng dụng Windows bằng PyInstaller.

Cách build (chạy trên máy Windows):
    pip install -r requirements.txt
    set PLAYWRIGHT_BROWSERS_PATH=0
    playwright install chromium
    pyinstaller build_windows.spec

Kết quả: thư mục dist\\ToolDangBaiFacebook\\ — đưa cả thư mục này cho người
dùng, họ bấm vào ToolDangBaiFacebook.exe là chạy, không cần cài Python hay
trình duyệt gì thêm.

Vì sao PLAYWRIGHT_BROWSERS_PATH=0: mặc định Playwright tải Chromium về thư mục
cache riêng của máy (%USERPROFILE%\\AppData\\Local\\ms-playwright), nằm ngoài
project nên PyInstaller không thấy để đóng kèm. Đặt biến này thành 0 buộc
Chromium nằm ngay trong package playwright, nhờ đó collect_all() gom được vào
gói. Lúc chạy, paths.setup_playwright_env() đặt lại đúng biến đó.
"""
from PyInstaller.utils.hooks import collect_all

# Gom toàn bộ package playwright: driver Node.js, các file .js, và Chromium
pw_datas, pw_binaries, pw_hiddenimports = collect_all("playwright")

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
