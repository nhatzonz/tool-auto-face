"""Chụp ảnh minh hoạ cho mục 3 của hướng dẫn sử dụng.

Chạy:  .venv313/bin/python chup_anh_huong_dan.py      (macOS)

Sinh ra 3 file trong anh_huong_dan/ đúng tên mà HUONGDAN_NGUOI_DUNG.md đang
trỏ tới. Sửa giao diện xong chạy lại script là ảnh trong hướng dẫn mới theo,
khỏi phải chụp tay rồi cắt cúp từng cái.

macOS đòi quyền Screen Recording cho ứng dụng chạy script (Terminal, VS Code,
iTerm...). Thiếu quyền thì screencapture vẫn tạo file nhưng chỉ có hình nền
desktop, không có cửa sổ — nên script tự kiểm tra và báo, chứ ảnh hỏng kiểu đó
nhìn qua rất dễ tưởng là xong.
"""
import os
import subprocess
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import gui

THU_MUC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "anh_huong_dan")

# (chỉ số tab, tên file, hàm chuẩn bị dữ liệu cho tab đó)
CAN_CHUP = [
    (0, "tai_khoan_va_chay.png", None),
    (1, "noi_dung_bai_dang.png", lambda a: a.load_excel()),
    (3, "nhom_theo_phan_loai.png", "chon_phan_loai"),
]


def _vung(cua_so):
    return (cua_so.winfo_rootx(), cua_so.winfo_rooty(),
            cua_so.winfo_width(), cua_so.winfo_height())


def _chup(cua_so, duong_dan):
    x, y, w, h = _vung(cua_so)
    subprocess.run(["screencapture", "-x", "-R", f"{x},{y},{w},{h}", duong_dan],
                   check=True)


def main():
    os.makedirs(THU_MUC, exist_ok=True)
    app = gui.App()
    app.update()
    app.geometry("1160x760")
    app.lift()
    app.attributes("-topmost", True)
    for _ in range(8):
        app.update()
        time.sleep(0.2)

    notebook = [w for w in app.winfo_children() if w.winfo_class() == "TNotebook"][0]

    # Có quyền chụp không: chụp 1 lần có cửa sổ, 1 lần giấu cửa sổ đi. Giống hệt
    # nhau nghĩa là cả hai lần đều chỉ chụp được hình nền.
    # Không đặt tên bắt đầu bằng dấu chấm: screencapture từ chối ghi file ẩn
    thu1 = os.path.join(THU_MUC, "kiem_tra_quyen_1.png")
    thu2 = os.path.join(THU_MUC, "kiem_tra_quyen_2.png")
    _chup(app, thu1)
    app.withdraw()
    app.update()
    time.sleep(0.6)
    _chup(app, thu2)
    app.deiconify()
    app.update()
    time.sleep(0.5)
    giong_nhau = open(thu1, "rb").read() == open(thu2, "rb").read()
    os.remove(thu1)
    os.remove(thu2)

    if giong_nhau:
        app.destroy()
        print("✗ Không có quyền chụp màn hình — ảnh chụp ra chỉ có hình nền.\n"
              "  Mở System Settings → Privacy & Security → Screen Recording,\n"
              "  bật cho ứng dụng đang chạy script này (Terminal / VS Code),\n"
              "  thoát hẳn ứng dụng đó rồi mở lại và chạy lệnh này lần nữa.")
        return 1

    for chi_so, ten_file, chuan_bi in CAN_CHUP:
        notebook.select(chi_so)
        app.update()
        if chuan_bi == "chon_phan_loai":
            app.refresh_categories()
            if app.cat_list.size():
                app.cat_list.selection_set(0)
                app.on_select_category()
        elif chuan_bi:
            chuan_bi(app)
        for _ in range(5):
            app.update()
            time.sleep(0.2)
        dd = os.path.join(THU_MUC, ten_file)
        _chup(app, dd)
        print(f"✓ {ten_file}  ({os.path.getsize(dd) // 1024} KB)")

    app.destroy()
    print(f"\nXong. Ảnh nằm trong {THU_MUC}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
