# -*- mode: python ; coding: utf-8 -*-
"""
Cấu hình đóng gói tool thành ứng dụng Windows bằng PyInstaller.

Cách build (chạy trên máy Windows):
    pip install -r requirements.txt
    pyinstaller build_windows.spec

Kết quả: thư mục dist\\ToolDangBaiFacebook\\ — đưa cả thư mục này cho người
dùng, họ bấm vào ToolDangBaiFacebook.exe là chạy, không cần cài Python.

KHÔNG đóng kèm Chromium của Playwright, và đây là chủ ý: tool mở Google Chrome
thật trên máy người dùng qua channel='chrome' (xem open_profile trong
post_phong_tro_fb.py). Chromium của Playwright bị Facebook nhận diện là
automation nên không đăng nhập được. Vì vậy máy người dùng cần có sẵn Chrome,
đổi lại gói nhẹ hơn khoảng 150MB.

Vẫn phải gom package playwright vì nó cần driver Node.js để điều khiển Chrome.
"""
import os

from PyInstaller.utils.hooks import collect_all

# Gom package playwright: driver Node.js và các file .js đi kèm
pw_datas, pw_binaries, pw_hiddenimports = collect_all("playwright")

a = Analysis(
    ["gui.py"],
    pathex=[],
    binaries=pw_binaries,
    # Nút "📖 Hướng dẫn sử dụng" đọc thẳng file .md này, nên nó phải nằm trong
    # gói — thiếu là cửa sổ hướng dẫn mở ra trống trơn.
    # Ảnh minh hoạ của mục 3: chỉ gói khi thư mục có ảnh thật, vì PyInstaller
    # báo lỗi nếu trỏ vào thư mục rỗng. Thiếu ảnh thì hướng dẫn vẫn mở được,
    # chỗ đó hiện dòng "[chưa có ảnh: ...]".
    datas=pw_datas + [("HUONGDAN_NGUOI_DUNG.md", ".")] + (
        [("anh_huong_dan", "anh_huong_dan")]
        if os.path.isdir("anh_huong_dan") and os.listdir("anh_huong_dan") else []),
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
