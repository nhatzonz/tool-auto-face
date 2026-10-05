# -*- mode: python ; coding: utf-8 -*-
"""Đóng gói tool thành ứng dụng macOS (.app).

Cách build (chạy trên máy Mac):
    .venv313/bin/pip install -r requirements.txt
    .venv313/bin/pyinstaller build_macos.spec

Kết quả: dist/ToolDangBaiFacebook.app — kéo vào thư mục Applications là xong,
bấm vào chạy như app bình thường, không cần Terminal, không cần cài Python.

Khác bản Windows ở 3 điểm:
  - Gói thành .app chứ không phải thư mục rời, nên phải có khối BUNDLE.
  - Dữ liệu người dùng KHÔNG nằm cạnh file chạy mà ở
    ~/Library/Application Support/ToolDangBaiFacebook/ — xem paths.py.
  - Build trên máy Apple Silicon ra app Apple Silicon; máy Intel phải build
    riêng trên máy Intel.

Giống bản Windows: KHÔNG đóng kèm Chromium của Playwright, tool mở Google
Chrome thật trên máy (channel='chrome'), nên máy phải có sẵn Chrome.
"""
import os

from PyInstaller.utils.hooks import collect_all

pw_datas, pw_binaries, pw_hiddenimports = collect_all("playwright")

a = Analysis(
    ["gui.py"],
    pathex=[],
    binaries=pw_binaries,
    datas=pw_datas + [("HUONGDAN_NGUOI_DUNG.md", ".")] + (
        [("anh_huong_dan", "anh_huong_dan")]
        if os.path.isdir("anh_huong_dan") and os.listdir("anh_huong_dan") else []),
    hiddenimports=pw_hiddenimports + ["openpyxl.cell._writer"],
    hookspath=[],
    runtime_hooks=[],
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
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="ToolDangBaiFacebook",
)

app = BUNDLE(
    coll,
    name="ToolDangBaiFacebook.app",
    icon=None,
    bundle_identifier="com.nhatnguyen.tooldangbaifacebook",
    info_plist={
        "CFBundleName": "Tool đăng bài Facebook",
        "CFBundleDisplayName": "Tool đăng bài Facebook",
        "CFBundleShortVersionString": "1.0.0",
        # Không có dòng này thì macOS coi app là ứng dụng đời cũ và chạy ở chế
        # độ phóng to mờ nhoè trên màn Retina.
        "NSHighResolutionCapable": True,
        # App không phải trình duyệt nhưng nó điều khiển Chrome mở trang
        # Facebook, cần cho phép kết nối mạng thường.
        "LSApplicationCategoryType": "public.app-category.productivity",
    },
)
