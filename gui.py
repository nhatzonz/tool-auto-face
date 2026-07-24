"""
Giao diện cho tool đăng bài Facebook hàng loạt.

Chạy:  source .venv/bin/activate && python gui.py

Làm việc theo "chiến dịch" — phòng trọ, tuyển dụng, seeding website... Mỗi
chiến dịch có file Excel, thư mục ảnh, file log và bản đồ nhóm riêng; nick
Facebook dùng chung cho mọi chiến dịch.

Các tab:
  - Tài khoản & Chạy   : chọn/thêm nick Facebook, bấm chạy, xem log trực tiếp
  - Nội dung bài đăng  : sửa trực tiếp file Excel của chiến dịch
  - Dữ liệu chiến dịch : file Excel, thư mục ảnh, thư mục profile, các delay
  - Nhóm theo phân loại: thêm/sửa/xóa phân loại và danh sách link group
  - Hẹn giờ đăng       : lịch tự chạy theo giờ Việt Nam (xem lich_hen.py)
  - Tham gia nhóm      : cho nick vào nhóm, chạy tách hẳn khỏi việc đăng bài
                         (xem auto_join.py) — danh sách nhóm và lịch sử riêng

Việc đăng bài chạy trong thread riêng để giao diện không bị treo.
"""
import json
import os
import queue
import shutil
import subprocess
import sys
import calendar
import webbrowser
import threading
import tkinter as tk
from datetime import date, datetime
from tkinter import ttk, filedialog, messagebox

import openpyxl

import paths

# Phải gọi trước khi import các module có print() ở cấp module
paths.guard_missing_stdout()

import auto_join as joiner
import config as cfg_module
import lich_hen
import post_phong_tro_fb as bot
import tao_file_mau


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tool đăng bài Facebook")
        self.geometry("960x700")

        self.cfg = cfg_module.load_config()
        # Cấu hình cũ có thể còn 2 chiến dịch trỏ chung một file chống đăng
        # trùng (trước đây không có gì chặn). Tách ra ngay lúc mở tool.
        self.log_da_tach = cfg_module.bao_dam_log_rieng(self.cfg)
        if self.log_da_tach:
            cfg_module.save_config(self.cfg)
        self.log_queue = queue.Queue()      # worker đăng bài → giao diện
        # Việc tham gia nhóm có ô log riêng ở tab riêng, nên có hàng đợi riêng:
        # trộn chung thì log của hai việc khác hẳn nhau đổ lẫn vào một chỗ.
        self.join_log_queue = queue.Queue()
        self.stop_event = threading.Event()
        self.worker = None
        self.current_cat = None             # phân loại đang chọn ở tab Nhóm
        self.dang_khoa = False              # True khi đang đăng → cấm sửa
        self.excel_loaded = False
        self.excel_snapshot = []            # ảnh chụp lần nạp/ghi Excel gần nhất
        self.join_da_tick = set()           # link nhóm đang tick ở tab Tham gia nhóm

        self._build_campaign_bar()

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.tab_run = ttk.Frame(notebook)
        self.tab_data = ttk.Frame(notebook)
        self.tab_paths = ttk.Frame(notebook)
        self.tab_groups = ttk.Frame(notebook)
        self.tab_lich = ttk.Frame(notebook)
        self.tab_join = ttk.Frame(notebook)
        notebook.add(self.tab_run, text="  Tài khoản & Chạy  ")
        notebook.add(self.tab_data, text="  Nội dung bài đăng  ")
        notebook.add(self.tab_paths, text="  Dữ liệu chiến dịch  ")
        notebook.add(self.tab_groups, text="  Nhóm theo phân loại  ")
        notebook.add(self.tab_lich, text="  ⏰ Hẹn giờ đăng  ")
        notebook.add(self.tab_join, text="  👥 Tham gia nhóm  ")

        self._build_run_tab()
        self._build_data_tab()
        self._build_paths_tab()
        self._build_groups_tab()
        self._build_lich_tab()
        self._build_join_tab()

        self._tao_du_lieu_mau_lan_dau()
        self.load_campaign_into_views()
        self.refresh_accounts()
        self.refresh_lich()
        self.refresh_join_nicks()
        self.after(100, self._drain_log_queue)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(50, self._hich_ve_lai)

        # Windows tự mở tool lúc khởi động máy thì thu nhỏ ngay, đừng nhảy ra
        # chắn màn hình người dùng khi họ vừa đăng nhập.
        if lich_hen.mo_thu_nho():
            self.iconify()

        for ten, cu, moi in self.log_da_tach:
            self.append_log(
                f"⚠ Chiến dịch '{ten}' đang dùng chung file chống đăng trùng với "
                f"chiến dịch khác ({cu or 'chưa đặt'}) — đã tách sang file riêng: "
                f"{os.path.basename(moi)}")

        self._nhip_dong_ho()
        self.after(3000, self._vong_kiem_tra_lich)

    def _tao_du_lieu_mau_lan_dau(self):
        """Lần chạy đầu trên máy mới thì chưa có file Excel nào — tạo sẵn file
        mẫu và thư mục ảnh cho chiến dịch đang chọn.

        Không có bước này, người dùng mở bản .exe lên sẽ chỉ thấy dòng "⚠ Chưa
        có file" mà không biết phải tạo file thế nào, đúng cột nào.
        """
        camp = self.campaign()
        excel = camp.get("excel_path")
        if excel and not os.path.exists(excel):
            try:
                tao_file_mau.tao_file(excel, camp.get("sheet_name") or "Sheet1")
            except Exception as e:
                self.log_queue.put(f"Không tạo được file Excel mẫu: {e}")

        for d in (camp.get("images_dir"), self.cfg.get("profiles_dir")):
            if d:
                os.makedirs(d, exist_ok=True)

    def _hich_ve_lai(self):
        """Tk 8.5.9 (bản Apple kèm sẵn) hay vẽ ra cửa sổ trắng trơn trên macOS
        đời mới; đổi kích thước 1 pixel ép nó vẽ lại toàn bộ widget."""
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"{w + 1}x{h + 1}")
        self.after(50, lambda: self.geometry(f"{w}x{h}"))

    # ==================== CHIẾN DỊCH ====================

    def campaign(self):
        """Dict cấu hình của chiến dịch đang chọn."""
        return cfg_module.get_campaign(self.cfg)

    def dang_chay(self):
        """True nếu đang đăng bài. Sửa cấu hình lúc này sẽ đổi biến toàn cục
        mà worker thread đang dùng — bài đang chạy sẽ ghi log sai chiến dịch
        và tra nhầm bản đồ nhóm."""
        if self.dang_khoa or (self.worker and self.worker.is_alive()):
            messagebox.showwarning(
                "Đang chạy",
                "Đang có tiến trình chạy. Bấm '■ Dừng' và chờ dừng hẳn rồi mới "
                "sửa được cấu hình — sửa giữa chừng sẽ làm việc đang chạy đọc "
                "nhầm cấu hình.")
            return True
        return False

    def _build_campaign_bar(self):
        bar = ttk.LabelFrame(self, text="Chiến dịch")
        bar.pack(fill="x", padx=8, pady=8)

        self.campaign_var = tk.StringVar(value=self.cfg["active_campaign"])
        self.campaign_box = ttk.Combobox(
            bar, textvariable=self.campaign_var, state="readonly", width=28,
            values=list(self.cfg["campaigns"]),
        )
        self.campaign_box.pack(side="left", padx=8, pady=8)
        self.campaign_box.bind("<<ComboboxSelected>>", self.on_switch_campaign)

        ttk.Button(bar, text="+ Thêm", command=self.add_campaign).pack(side="left", padx=2)
        ttk.Button(bar, text="Nhân bản", command=self.duplicate_campaign).pack(side="left", padx=2)
        ttk.Button(bar, text="Đổi tên", command=self.rename_campaign).pack(side="left", padx=2)
        ttk.Button(bar, text="Xóa", command=self.delete_campaign).pack(side="left", padx=2)

    def refresh_campaign_box(self):
        self.campaign_box["values"] = list(self.cfg["campaigns"])
        self.campaign_var.set(self.cfg["active_campaign"])

    def on_switch_campaign(self, _event=None):
        """Đổi chiến dịch đang chọn rồi nạp lại toàn bộ ô nhập."""
        if self.dang_chay():
            self.campaign_var.set(self.cfg["active_campaign"])   # trả về như cũ
            return
        self.cfg["active_campaign"] = self.campaign_var.get()
        cfg_module.save_config(self.cfg)
        self.load_campaign_into_views()
        self.append_log(f"→ Đã chuyển sang chiến dịch: {self.cfg['active_campaign']}")

    def load_campaign_into_views(self):
        """Đổ dữ liệu chiến dịch hiện tại vào tab Dữ liệu và tab Nhóm."""
        camp = self.campaign()
        for key, var in self.path_vars.items():
            var.set(camp[key])
        self.profiles_var.set(self.cfg["profiles_dir"])
        for key, var in self.num_vars.items():
            var.set(str(camp[key]))
        self.chong_trung_var.set(bool(camp.get("chong_trung", True)))

        self.current_cat = None
        self.url_box.delete("1.0", "end")
        self.refresh_categories()
        self.load_excel()

    def add_campaign(self):
        if self.dang_chay():
            return
        name = SimplePrompt(self, "Chiến dịch mới",
                            "Tên chiến dịch (vd: Tuyển dụng, Seeding web):").result
        if not name:
            return
        if name in self.cfg["campaigns"]:
            messagebox.showinfo("Đã có", f"Chiến dịch '{name}' đã tồn tại.")
            return
        dang_dung = [c["posted_log"] for c in self.cfg["campaigns"].values()]
        self.cfg["campaigns"][name] = cfg_module.new_campaign(name, dang_dung)
        self.cfg["active_campaign"] = name
        cfg_module.save_config(self.cfg)
        self.refresh_campaign_box()
        self.load_campaign_into_views()
        messagebox.showinfo(
            "Đã tạo",
            f"Đã tạo chiến dịch '{name}'.\n\n"
            "Bước tiếp theo: sang tab 'Dữ liệu chiến dịch' chọn file Excel và "
            "thư mục ảnh, rồi sang tab 'Nhóm theo phân loại' thêm các nhóm.",
        )

    def duplicate_campaign(self):
        if self.dang_chay():
            return
        src = self.cfg["active_campaign"]
        name = SimplePrompt(self, "Nhân bản", f"Tên bản sao của '{src}':", f"{src} 2").result
        if not name:
            return
        if name in self.cfg["campaigns"]:
            messagebox.showinfo("Đã có", f"Chiến dịch '{name}' đã tồn tại.")
            return
        copy = json.loads(json.dumps(self.campaign()))
        # Log riêng, nếu không 2 chiến dịch sẽ coi nhau là đã đăng
        dang_dung = [c["posted_log"] for c in self.cfg["campaigns"].values()]
        copy["posted_log"] = cfg_module.new_campaign(name, dang_dung)["posted_log"]
        self.cfg["campaigns"][name] = copy
        self.cfg["active_campaign"] = name
        cfg_module.save_config(self.cfg)
        self.refresh_campaign_box()
        self.load_campaign_into_views()

    def rename_campaign(self):
        if self.dang_chay():
            return
        old = self.cfg["active_campaign"]
        new = SimplePrompt(self, "Đổi tên chiến dịch", f"Tên mới cho '{old}':", old).result
        if not new or new == old:
            return
        if new in self.cfg["campaigns"]:
            messagebox.showinfo("Đã có", f"Chiến dịch '{new}' đã tồn tại.")
            return
        # Dựng lại dict để giữ nguyên thứ tự các chiến dịch
        self.cfg["campaigns"] = {
            (new if k == old else k): v for k, v in self.cfg["campaigns"].items()
        }
        self.cfg["active_campaign"] = new
        # Lịch hẹn giờ nhớ chiến dịch theo TÊN. Quên đổi ở đây thì lịch trỏ vào
        # một cái tên không còn tồn tại, và người dùng chỉ phát hiện ra vào lúc
        # tới giờ mà chẳng có gì được đăng.
        doi = [l for l in self.danh_sach_lich() if l.get("campaign") == old]
        for l in doi:
            l["campaign"] = new
        cfg_module.save_config(self.cfg)
        self.refresh_campaign_box()
        self.refresh_lich()
        if doi:
            self.append_log(f"→ Đã cập nhật {len(doi)} lịch hẹn sang tên mới '{new}'.")

    def delete_campaign(self):
        if self.dang_chay():
            return
        if len(self.cfg["campaigns"]) == 1:
            messagebox.showwarning("Không xóa được", "Phải còn ít nhất 1 chiến dịch.")
            return
        name = self.cfg["active_campaign"]
        lich_lien_quan = [l for l in self.danh_sach_lich() if l.get("campaign") == name]
        canh_bao_lich = (
            f"\n\n⚠ Có {len(lich_lien_quan)} lịch hẹn giờ đang dùng chiến dịch này, "
            "sẽ bị xóa theo." if lich_lien_quan else "")
        if not messagebox.askyesno(
            "Xóa chiến dịch",
            f"Xóa chiến dịch '{name}' khỏi cấu hình?\n\n"
            "File Excel, ảnh và log trên ổ đĩa vẫn giữ nguyên, chỉ mất phần "
            "khai báo đường dẫn và danh sách nhóm." + canh_bao_lich,
        ):
            return
        self.cfg["campaigns"].pop(name)
        # Lịch mồ côi không chạy được nữa, để lại chỉ tổ gây hiểu nhầm là còn hẹn
        for l in lich_lien_quan:
            self.danh_sach_lich().remove(l)
        self.cfg["active_campaign"] = next(iter(self.cfg["campaigns"]))
        cfg_module.save_config(self.cfg)
        self.refresh_campaign_box()
        self.load_campaign_into_views()
        self.refresh_lich()

    # ==================== TAB: TÀI KHOẢN & CHẠY ====================

    def _build_run_tab(self):
        f = self.tab_run

        left = ttk.LabelFrame(f, text="Nick Facebook (dùng chung mọi chiến dịch)")
        left.pack(side="left", fill="y", padx=(0, 8), pady=4)

        self.acc_list = tk.Listbox(left, width=26, exportselection=False)
        self.acc_list.pack(fill="y", expand=True, padx=6, pady=6)

        ttk.Button(left, text="+ Thêm nick mới", command=self.add_account).pack(fill="x", padx=6, pady=2)
        ttk.Button(left, text="Đăng nhập lại nick này", command=self.relogin_account).pack(fill="x", padx=6, pady=2)
        ttk.Button(left, text="Xóa nick", command=self.delete_account).pack(fill="x", padx=6, pady=(2, 8))

        right = ttk.Frame(f)
        right.pack(side="left", fill="both", expand=True, pady=4)

        bar = ttk.Frame(right)
        bar.pack(fill="x", pady=(0, 6))
        self.btn_start = ttk.Button(bar, text="▶ Bắt đầu đăng", command=self.start_posting)
        self.btn_start.pack(side="left")
        self.btn_stop = ttk.Button(bar, text="■ Dừng", command=self.request_stop, state="disabled")
        self.btn_stop.pack(side="left", padx=6)
        self.btn_check = ttk.Button(bar, text="🔍 Kiểm tra dữ liệu", command=self.kiem_tra)
        self.btn_check.pack(side="left", padx=(0, 6))
        # Tên nút phải nói rõ nó xóa cái gì: một nút chỉ dọn chữ trên màn hình,
        # nút kia xóa file chống đăng trùng — nhầm cái thứ hai là cả loạt bài cũ
        # lên Facebook lần nữa.
        ttk.Button(bar, text="Xóa màn hình log",
                   command=self.xoa_man_hinh_log).pack(side="left")
        ttk.Button(bar, text="🗑 Xóa lịch sử đã đăng",
                   command=self.xoa_lich_su_dang).pack(side="left", padx=6)

        self.status = ttk.Label(right, text="Sẵn sàng.")
        self.status.pack(anchor="w", pady=(0, 4))

        self.log_box = tk.Text(right, wrap="word", height=25, font=("Menlo", 11))
        scroll = ttk.Scrollbar(right, command=self.log_box.yview)
        self.log_box.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log_box.pack(fill="both", expand=True)

    def refresh_accounts(self):
        """Nạp lại danh sách nick từ thư mục profiles hiện tại."""
        bot.reload_config()
        self.acc_list.delete(0, "end")
        for name, _ in bot.list_sessions():
            self.acc_list.insert("end", name)

    def selected_account(self):
        """Trả về (tên, đường dẫn profile) của nick đang chọn, hoặc None."""
        sel = self.acc_list.curselection()
        if not sel:
            return None
        name = self.acc_list.get(sel[0])
        return name, os.path.join(self.cfg["profiles_dir"], name)

    def add_account(self):
        name = SimplePrompt(self, "Tên nick mới", "Đặt tên cho nick (vd: acc_phu_1):").result
        if not name:
            return
        safe = bot.sanitize_account_name(name)
        if not safe:
            messagebox.showerror("Lỗi", "Tên không hợp lệ.")
            return
        path = os.path.join(self.cfg["profiles_dir"], safe)
        if os.path.isdir(path):
            messagebox.showinfo("Đã có", f"Nick '{safe}' đã tồn tại, sẽ mở lại profile này.")
        self._open_login_window(safe, path)

    def relogin_account(self):
        acc = self.selected_account()
        if not acc:
            messagebox.showwarning("Chưa chọn", "Hãy chọn một nick trong danh sách.")
            return
        self._open_login_window(*acc)

    def _open_login_window(self, name, path):
        """Mở Chrome cho người dùng đăng nhập thủ công, chờ họ bấm Xong."""
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("Đang bận", "Đang có tiến trình chạy, hãy dừng trước.")
            return

        done = threading.Event()

        win = tk.Toplevel(self)
        win.title(f"Đăng nhập: {name}")
        win.geometry("420x160")
        ttk.Label(
            win,
            text=f"Chrome đang mở cho nick '{name}'.\n"
                 "Hãy đăng nhập Facebook trong cửa sổ Chrome đó,\n"
                 "xong rồi quay lại đây bấm nút bên dưới.",
            justify="left",
        ).pack(padx=16, pady=16)
        ttk.Button(win, text="Tôi đã đăng nhập xong", command=done.set).pack(pady=4)
        win.protocol("WM_DELETE_WINDOW", done.set)

        def work():
            from playwright.sync_api import sync_playwright
            try:
                with sync_playwright() as p:
                    context = bot.open_profile(p, path, slow_mo=200)
                    page = bot.get_page(context)
                    page.goto("https://www.facebook.com/login", wait_until="domcontentloaded")
                    done.wait()
                    ok = bot.is_logged_in(page)
                    context.close()
                self.log_queue.put(
                    f"✓ Đã lưu đăng nhập cho '{name}'." if ok
                    else f"⚠ '{name}' có vẻ chưa đăng nhập xong — thử lại nếu cần."
                )
            except Exception as e:
                self.log_queue.put(f"✗ Lỗi mở Chrome: {e}")
            finally:
                self.log_queue.put(("__login_done__", win))

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def delete_account(self):
        acc = self.selected_account()
        if not acc:
            messagebox.showwarning("Chưa chọn", "Hãy chọn một nick trong danh sách.")
            return
        name, path = acc
        lich_lien_quan = [l for l in self.danh_sach_lich() if l.get("nick") == name]
        canh_bao_lich = (
            f"\n\n⚠ Có {len(lich_lien_quan)} lịch hẹn giờ đang dùng nick này — "
            "xóa xong các lịch đó sẽ không chạy được nữa, hãy sửa lại nick cho "
            "chúng ở tab '⏰ Hẹn giờ đăng'." if lich_lien_quan else "")
        if not messagebox.askyesno(
            "Xóa nick",
            f"Xóa hẳn profile '{name}'?\nSẽ mất cookie đăng nhập của nick này."
            + canh_bao_lich):
            return
        shutil.rmtree(path, ignore_errors=True)
        self.refresh_accounts()
        self.refresh_join_nicks()
        self.append_log(f"Đã xóa nick '{name}'.")

    # ==================== CHẠY ĐĂNG BÀI ====================

    def start_posting(self):
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("Đang chạy", "Tiến trình đang chạy rồi.")
            return
        acc = self.selected_account()
        if not acc:
            messagebox.showwarning("Chưa chọn nick", "Hãy chọn nick Facebook để đăng bài.")
            return
        nick, path = acc

        camp_name = self.cfg["active_campaign"]
        camp = self.campaign()
        if not camp["groups"]:
            messagebox.showerror(
                "Chưa có nhóm",
                f"Chiến dịch '{camp_name}' chưa khai báo nhóm nào.\n"
                "Vào tab 'Nhóm theo phân loại' để thêm.",
            )
            return

        if not os.path.exists(camp["excel_path"]):
            messagebox.showerror(
                "Thiếu file",
                f"Chiến dịch '{camp_name}' chưa có file Excel hợp lệ:\n{camp['excel_path']}\n\n"
                "Vào tab 'Dữ liệu chiến dịch' để chọn file.",
            )
            return

        # Soát dữ liệu trước; có lỗi thì không cho chạy
        van_de = bot.kiem_tra_du_lieu(camp_name)
        loi = [v for v in van_de if v["muc"] == "loi"]
        canh_bao = [v for v in van_de if v["muc"] == "canh_bao"]
        if loi:
            self.hien_van_de(van_de, camp_name)
            messagebox.showerror(
                "Có lỗi trong dữ liệu",
                f"Tìm thấy {len(loi)} lỗi — không thể chạy.\n\n"
                "Xem chi tiết ở ô log, sửa xong bấm 🔍 Kiểm tra dữ liệu lại.")
            return
        if canh_bao:
            self.hien_van_de(van_de, camp_name)
            if not messagebox.askyesno(
                "Có cảnh báo",
                f"Không có lỗi chặn, nhưng có {len(canh_bao)} cảnh báo "
                "(xem chi tiết ở ô log).\n\nVẫn chạy?"):
                return

        self._khoi_dong_worker(nick, path, camp_name)

    def _khoi_dong_worker(self, nick, path, camp_name, tu_lich=False):
        """Nhả worker thread đăng bài. Dùng chung cho nút '▶ Bắt đầu đăng' và
        cho lịch hẹn giờ tự kích hoạt — hai đường vào, một đường chạy."""
        self.chay_tu_lich = tu_lich
        self.stop_event.clear()
        bot.reload_config(camp_name)
        bot.LOG_FN = self.log_queue.put
        bot.SHOULD_STOP = self.stop_event.is_set

        self.set_ui_locked(True)
        self.status.config(text=f"⏵ Đang chạy — chiến dịch '{camp_name}', nick: {nick}")

        def work():
            stats = None
            try:
                stats = bot.run_posting(path, campaign=camp_name)
            except Exception as e:
                self.log_queue.put(f"✗ LỖI: {e}")
            finally:
                self.log_queue.put(("__run_done__", stats))

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def xoa_man_hinh_log(self):
        """Dọn chữ trong ô log. Không đụng dữ liệu, nhưng vẫn hỏi: log của lần
        chạy vừa rồi là thứ duy nhất cho biết bài nào lỗi ở nhóm nào, xóa nhầm
        là mất luôn, không xem lại được."""
        if not self.log_box.get("1.0", "end").strip():
            return
        if not messagebox.askyesno(
            "Xóa màn hình log",
            "Xóa toàn bộ chữ đang hiện trong ô log?\n\n"
            "Chỉ dọn màn hình, KHÔNG đụng tới lịch sử đã đăng. Nhưng nội dung "
            "log của lần chạy vừa rồi sẽ không xem lại được."):
            return
        self.log_box.delete("1.0", "end")

    def xoa_lich_su_dang(self):
        """Mở cửa sổ chọn log của chiến dịch nào để xóa.

        Mỗi chiến dịch có file chống đăng trùng riêng, nên phải cho chọn chứ
        không mặc định xóa của chiến dịch đang mở — người dùng hay đứng ở chiến
        dịch này mà muốn dọn log của chiến dịch kia.
        """
        if self.dang_chay():
            return
        XoaLichSuDang(self)

    @staticmethod
    def dem_luot_da_dang(duong_dan):
        """Số lượt đã đăng ghi trong một file log; -1 nếu không đọc được."""
        if not duong_dan or not os.path.exists(duong_dan):
            return 0
        try:
            with open(duong_dan, "r", encoding="utf-8") as f:
                return len(json.load(f))
        except Exception:
            return -1

    def request_stop(self):
        """Dừng việc đang chạy — dùng chung cho cả đăng bài lẫn tham gia nhóm,
        vì hai việc không bao giờ chạy cùng lúc (chung một worker thread)."""
        self.stop_event.set()
        self.btn_stop.config(state="disabled")
        self.btn_join_stop.config(state="disabled")
        self.status.config(text="⏸ Đang dừng... chờ việc hiện tại xong rồi mới nhả khoá.")

    def set_ui_locked(self, khoa):
        """Khoá/mở toàn bộ phần sửa cấu hình.

        Trong lúc đăng, worker thread đang đọc các biến toàn cục của bot; sửa
        cấu hình lúc đó làm bài đang chạy ghi log sai chiến dịch và tra nhầm
        bản đồ nhóm. Nên khoá hẳn thay vì chỉ hiện cảnh báo.
        """
        self.dang_khoa = khoa
        trang_thai = "disabled" if khoa else "normal"

        def duyet(widget):
            for con in widget.winfo_children():
                # Nút Dừng phải luôn bấm được, nếu không sẽ không dừng nổi
                if con in (self.btn_stop, self.btn_join_stop):
                    continue
                if isinstance(con, (ttk.Button, ttk.Entry, tk.Listbox, tk.Text, ttk.Combobox)):
                    try:
                        con.config(state=trang_thai)
                    except tk.TclError:
                        pass
                duyet(con)

        duyet(self)
        # Combobox chiến dịch vốn là readonly, trả lại đúng trạng thái đó
        if not khoa:
            self.campaign_box.config(state="readonly")
        self.btn_start.config(state="disabled" if khoa else "normal")
        self.btn_stop.config(state="normal" if khoa else "disabled")
        # Hai ô log luôn phải mở. Không chỉ để đọc: Text đang bị khoá thì
        # insert() im lặng không ghi được gì — log của lượt chạy sẽ mất sạch mà
        # không báo lỗi, nhìn như tool đứng im trong khi nó vẫn đang chạy.
        self.log_box.config(state="normal")
        self.join_log_box.config(state="normal")

    # ==================== KIỂM TRA DỮ LIỆU ====================

    def kiem_tra(self):
        """Soát dữ liệu chiến dịch hiện tại và in kết quả ra ô log."""
        camp_name = self.cfg["active_campaign"]
        try:
            van_de = bot.kiem_tra_du_lieu(camp_name)
        except Exception as e:
            messagebox.showerror("Lỗi kiểm tra", f"Không kiểm tra được:\n{e}")
            return
        self.hien_van_de(van_de, camp_name)

        loi = sum(1 for v in van_de if v["muc"] == "loi")
        canh_bao = len(van_de) - loi
        if not van_de:
            messagebox.showinfo("Không có vấn đề", f"Chiến dịch '{camp_name}': dữ liệu hợp lệ, chạy được.")
        elif loi:
            messagebox.showerror("Có lỗi", f"{loi} lỗi và {canh_bao} cảnh báo.\nChi tiết ở ô log.")
        else:
            messagebox.showwarning("Có cảnh báo", f"{canh_bao} cảnh báo, không có lỗi chặn.\nChi tiết ở ô log.")

    def hien_van_de(self, van_de, camp_name):
        """In danh sách vấn đề ra ô log, lỗi trước cảnh báo sau."""
        self.append_log("")
        self.append_log("=" * 60)
        self.append_log(f"KIỂM TRA DỮ LIỆU — chiến dịch '{camp_name}'")
        if not van_de:
            self.append_log("✓ Không phát hiện vấn đề nào.")
            self.append_log("=" * 60)
            return
        for muc, nhan in (("loi", "✗ LỖI"), ("canh_bao", "⚠ CẢNH BÁO")):
            nhom = [v for v in van_de if v["muc"] == muc]
            if not nhom:
                continue
            self.append_log(f"{nhan} ({len(nhom)}):")
            for v in nhom:
                self.append_log(f"   • [{v['o']}] {v['chi_tiet']}")
        self.append_log("=" * 60)

    # ==================== LOG ====================

    def _bao_da_dang_hom_nay(self, stats):
        """Chạy xong mà có lượt bị chặn vì trùng thì nói rõ cho người dùng.

        Chỉ hiện với lần chạy bấm tay. Lần chạy do lịch hẹn kích hoạt có thể
        diễn ra lúc 3h sáng, dựng hộp thoại ở đó thì nó đứng nguyên tới sáng —
        mà hộp thoại đang mở lại khiến các lịch sau bị hoãn theo.
        """
        if not stats or not stats.get("dup"):
            return
        if getattr(self, "chay_tu_lich", False):
            return
        messagebox.showinfo(
            "Có bài không đăng vì trùng",
            f"{stats['dup']} lượt không đăng vì HÔM NAY đã đăng bài đó lên "
            "nhóm đó rồi — tính năng chống đăng trùng đang bật.\n\n"
            "Muốn đăng lại ngay trong hôm nay, chọn một trong hai:\n"
            "  • Bấm '🗑 Xóa lịch sử đã đăng' rồi chạy lại\n"
            "  • Tắt ô 'Chống đăng trùng' ở tab 'Dữ liệu chiến dịch'\n\n"
            "Để nguyên thì sang ngày mai các bài này lại đăng được bình thường.")

    def append_log(self, text):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")

    def _drain_log_queue(self):
        """Đọc log từ worker thread và đổ vào ô log — chỉ chạy ở main thread."""
        while True:
            try:
                item = self.log_queue.get_nowait()
            except queue.Empty:
                break

            if isinstance(item, tuple):
                kind, payload = item
                if kind == "__run_done__":
                    self.set_ui_locked(False)
                    self.status.config(text="Đã dừng — giờ sửa được cấu hình.")
                    self._bao_da_dang_hom_nay(payload)
                elif kind == "__login_done__":
                    payload.destroy()
                    self.refresh_accounts()
                    self.refresh_join_nicks()
                continue

            self.append_log(str(item))

        self._drain_join_log_queue()
        self.after(100, self._drain_log_queue)

    # ==================== TAB: NỘI DUNG BÀI ĐĂNG ====================

    def _build_data_tab(self):
        """Bảng sửa trực tiếp nội dung file Excel của chiến dịch đang chọn."""
        f = self.tab_data
        self.excel_rows = []        # [{stt, noi_dung, phan_loai, folder}, ...]

        bar = ttk.Frame(f)
        bar.pack(fill="x", pady=6)
        ttk.Button(bar, text="+ Thêm bài", command=self.add_row).pack(side="left")
        ttk.Button(bar, text="Sửa bài", command=self.edit_row).pack(side="left", padx=4)
        ttk.Button(bar, text="Xóa bài", command=self.delete_row).pack(side="left")
        ttk.Button(bar, text="🖼 Ảnh của bài này", command=self.quan_ly_anh).pack(side="left", padx=12)
        ttk.Button(bar, text="⟳ Tải lại từ Excel", command=self.load_excel).pack(side="left", padx=12)
        ttk.Button(bar, text="💾 Ghi vào file Excel", command=self.save_excel).pack(side="left")

        ttk.Label(
            f,
            text="STT là mã định danh của bài — file chống đăng trùng ghi theo STT. "
                 "Đổi STT của bài cũ sẽ làm tool đăng lại hoặc bỏ sót bài.",
            foreground="#a05000",
        ).pack(anchor="w", padx=2, pady=(4, 0))

        self.excel_status = ttk.Label(f, text="", foreground="#555")
        self.excel_status.pack(anchor="w", padx=2, pady=(0, 4))

        cols = ("stt", "noi_dung", "phan_loai", "folder")
        self.data_tree = ttk.Treeview(f, columns=cols, show="headings", height=18)
        for col, title, width in [
            ("stt", "STT", 50),
            ("noi_dung", "Nội dung bài", 520),
            ("phan_loai", "Phân loại", 150),
            ("folder", "Thư mục ảnh", 140),
        ]:
            self.data_tree.heading(col, text=title)
            self.data_tree.column(col, width=width, anchor="w")

        vs = ttk.Scrollbar(f, orient="vertical", command=self.data_tree.yview)
        self.data_tree.configure(yscrollcommand=vs.set)
        vs.pack(side="right", fill="y")
        self.data_tree.pack(fill="both", expand=True)
        self.data_tree.bind("<Double-1>", lambda _e: self.edit_row())

    def refresh_data_tree(self):
        """Vẽ lại bảng từ self.excel_rows.

        KHÔNG đánh số lại STT: posted_log định danh bài bằng STT, nên đổi STT
        của bài cũ sẽ khiến tool tưởng nhầm đã đăng (bỏ sót) hoặc đăng lại.
        Bảng sắp xếp theo STT vì bot cũng đăng theo thứ tự STT.
        """
        self.excel_rows.sort(key=lambda r: (r["stt"] is None, r["stt"] or 0))
        self.data_tree.delete(*self.data_tree.get_children())
        for row in self.excel_rows:
            preview = row["noi_dung"].replace("\n", " ⏎ ")
            if len(preview) > 90:
                preview = preview[:90] + "..."
            self.data_tree.insert(
                "", "end", values=(row["stt"], preview, row["phan_loai"], row["folder"]))

    def quan_ly_anh(self):
        """Mở cửa sổ xem/xóa/thêm ảnh cho thư mục ảnh của bài đang chọn."""
        i = self.selected_row_index()
        if i is None:
            messagebox.showwarning("Chưa chọn", "Hãy chọn một bài trong bảng.")
            return
        row = self.excel_rows[i]
        ten_thu_muc = (row.get("folder") or "").strip()
        if not ten_thu_muc:
            messagebox.showinfo(
                "Bài này chưa có thư mục ảnh",
                "Cột 'Thư mục ảnh' của bài đang trống nên không có ảnh nào để "
                "quản lý.\n\nBấm 'Sửa bài' để đặt tên thư mục ảnh trước (ví dụ "
                "'phong_501'), rồi quay lại đây.")
            return

        goc = self.campaign().get("images_dir") or ""
        duong_dan = os.path.join(goc, ten_thu_muc)
        if not os.path.isdir(duong_dan):
            if not messagebox.askyesno(
                "Chưa có thư mục",
                f"Chưa có thư mục ảnh cho bài này:\n{duong_dan}\n\nTạo bây giờ?"):
                return
            try:
                os.makedirs(duong_dan, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Không tạo được", f"Không tạo được thư mục:\n{e}")
                return

        QuanLyAnh(self, duong_dan, f"STT {row['stt']} — {ten_thu_muc}")

    def next_stt(self):
        """STT trống tiếp theo — không đụng vào STT nào đã dùng."""
        used = [r["stt"] for r in self.excel_rows if isinstance(r["stt"], int)]
        return max(used) + 1 if used else 1

    def load_excel(self):
        """Đọc file Excel của chiến dịch hiện tại vào bảng."""
        camp = self.campaign()
        path, sheet = camp["excel_path"], camp["sheet_name"]
        self.excel_rows = []
        self.excel_loaded = False       # chưa đọc được thì cấm ghi đè

        if not os.path.exists(path):
            self.refresh_data_tree()
            self.excel_status.config(text=f"⚠ Chưa có file: {path}")
            return
        try:
            wb = openpyxl.load_workbook(path)
            if sheet not in wb.sheetnames:
                self.refresh_data_tree()
                self.excel_status.config(
                    text=f"⚠ File không có sheet '{sheet}'. Sheet đang có: {', '.join(wb.sheetnames)}")
                return
            ws = wb[sheet]
            for r in range(2, ws.max_row + 1):
                vals = [ws.cell(row=r, column=c).value for c in range(1, 5)]
                if all(v is None or str(v).strip() == "" for v in vals):
                    continue
                try:
                    stt = int(float(vals[0])) if vals[0] is not None else None
                except (TypeError, ValueError):
                    stt = None
                self.excel_rows.append({
                    "stt": stt,
                    "noi_dung": str(vals[1] or "").strip(),
                    "phan_loai": str(vals[2] or "").strip(),
                    "folder": str(vals[3] or "").strip(),
                })
        except Exception as e:
            self.excel_status.config(text=f"⚠ Không đọc được file: {e}")
            self.refresh_data_tree()
            return

        self.excel_loaded = True
        self.refresh_data_tree()
        self.excel_snapshot = [dict(r) for r in self.excel_rows]
        self.excel_status.config(
            text=f"{len(self.excel_rows)} bài — {os.path.basename(path)} / sheet '{sheet}'")

    def selected_row_index(self):
        sel = self.data_tree.selection()
        if not sel:
            return None
        return self.data_tree.index(sel[0])

    def add_row(self):
        cats = list(self.campaign()["groups"])
        row = RowEditor(self, "Thêm bài đăng", None, cats, self.next_stt()).result
        if row:
            self.excel_rows.append(row)
            self.refresh_data_tree()
            self.excel_status.config(text="Đã thêm — nhớ bấm 💾 Ghi vào file Excel.")

    def edit_row(self):
        idx = self.selected_row_index()
        if idx is None:
            messagebox.showwarning("Chưa chọn", "Hãy chọn một bài trong bảng.")
            return
        cats = list(self.campaign()["groups"])
        row = RowEditor(self, "Sửa bài đăng", self.excel_rows[idx], cats).result
        if row:
            self.excel_rows[idx] = row
            self.refresh_data_tree()
            self.excel_status.config(text="Đã sửa — nhớ bấm 💾 Ghi vào file Excel.")

    def delete_row(self):
        idx = self.selected_row_index()
        if idx is None:
            messagebox.showwarning("Chưa chọn", "Hãy chọn một bài trong bảng.")
            return
        if not messagebox.askyesno("Xóa bài", f"Xóa bài STT {self.excel_rows[idx]['stt']} khỏi bảng?"):
            return
        self.excel_rows.pop(idx)
        self.refresh_data_tree()
        self.excel_status.config(text="Đã xóa — nhớ bấm 💾 Ghi vào file Excel.")

    def save_excel(self):
        """Ghi bảng xuống file Excel gốc. Sao lưu bản cũ thành .bak trước khi ghi."""
        camp = self.campaign()
        path, sheet = camp["excel_path"], camp["sheet_name"]

        if not self.excel_loaded:
            messagebox.showerror(
                "Chưa đọc được file",
                "Bảng này chưa nạp được nội dung từ Excel, nên ghi xuống sẽ "
                "xóa trắng file.\n\nBấm '⟳ Tải lại từ Excel' và xử lý lỗi hiện "
                "ở dòng trạng thái trước đã.")
            return
        if not os.path.exists(path):
            messagebox.showerror("Thiếu file", f"Chưa có file Excel:\n{path}")
            return
        # File ~$ nghĩa là Excel đang mở file này — ghi đè sẽ mất phần đang gõ dở
        lock = os.path.join(os.path.dirname(path), "~$" + os.path.basename(path))
        if os.path.exists(lock) and not messagebox.askyesno(
            "File đang mở",
            f"Có vẻ '{os.path.basename(path)}' đang mở trong Excel.\n\n"
            "Ghi đè bây giờ có thể mất phần bạn đang gõ dở bên Excel, và Excel "
            "có thể ghi đè ngược lại khi bạn bấm Save bên đó.\n\nVẫn ghi?",
        ):
            return
        if not messagebox.askyesno(
            "Ghi vào Excel",
            f"Ghi {len(self.excel_rows)} bài vào:\n{path}\n(sheet '{sheet}')\n\n"
            "Bản cũ được sao lưu thành file .bak cạnh đó.",
        ):
            return

        try:
            shutil.copy2(path, path + ".bak")
            wb = openpyxl.load_workbook(path)
            ws = wb[sheet]

            for i, row in enumerate(self.excel_rows, start=2):
                ws.cell(row=i, column=1).value = row["stt"]
                ws.cell(row=i, column=2).value = row["noi_dung"]
                ws.cell(row=i, column=3).value = row["phan_loai"]
                ws.cell(row=i, column=4).value = row["folder"] or None

            # Xóa sạch các dòng thừa còn sót lại từ lần trước
            for r in range(len(self.excel_rows) + 2, ws.max_row + 1):
                for c in range(1, 5):
                    ws.cell(row=r, column=c).value = None

            wb.save(path)
        except Exception as e:
            messagebox.showerror("Lỗi ghi file", f"Không ghi được:\n{e}")
            return

        self.excel_snapshot = [dict(r) for r in self.excel_rows]
        self.excel_status.config(
            text=f"✓ Đã ghi {len(self.excel_rows)} bài vào {os.path.basename(path)} "
                 f"(bản cũ: {os.path.basename(path)}.bak)")
        messagebox.showinfo(
            "Đã ghi",
            f"Đã ghi {len(self.excel_rows)} bài vào file Excel.\n\n"
            "Lần chạy tới tool sẽ đọc luôn nội dung mới này.",
        )

    # ==================== TAB: DỮ LIỆU CHIẾN DỊCH ====================

    def _build_paths_tab(self):
        f = self.tab_paths
        self.path_vars = {}

        hint = ttk.Label(
            f,
            text="File Excel cần 4 cột:  A = STT  |  B = Nội dung bài  |  "
                 "C = Phân loại  |  D = Tên thư mục ảnh (để trống nếu không đăng ảnh)",
            foreground="#555",
        )
        hint.grid(row=0, column=0, columnspan=3, sticky="w", padx=6, pady=(8, 12))

        rows = [
            ("excel_path", "File Excel dữ liệu", "file"),
            ("sheet_name", "Tên sheet trong Excel", "text"),
            ("images_dir", "Thư mục ảnh", "dir"),
            ("posted_log", "File log bài đã đăng", "file"),
        ]
        for i, (key, label, kind) in enumerate(rows):
            r = i + 1
            ttk.Label(f, text=label).grid(row=r, column=0, sticky="w", padx=6, pady=6)
            var = tk.StringVar()
            self.path_vars[key] = var
            ttk.Entry(f, textvariable=var, width=62).grid(row=r, column=1, sticky="we", padx=6)
            if kind != "text":
                ttk.Button(f, text="Chọn...", command=lambda v=var, k=kind: self._pick_path(v, k)) \
                    .grid(row=r, column=2, padx=6)

        ttk.Separator(f, orient="horizontal").grid(row=9, column=0, columnspan=3, sticky="we", pady=12)

        ttk.Label(f, text="Thư mục profile Chrome (chung mọi chiến dịch)") \
            .grid(row=10, column=0, sticky="w", padx=6, pady=6)
        self.profiles_var = tk.StringVar()
        ttk.Entry(f, textvariable=self.profiles_var, width=62).grid(row=10, column=1, sticky="we", padx=6)
        ttk.Button(f, text="Chọn...", command=lambda: self._pick_path(self.profiles_var, "dir")) \
            .grid(row=10, column=2, padx=6)

        ttk.Separator(f, orient="horizontal").grid(row=11, column=0, columnspan=3, sticky="we", pady=12)

        nums = [
            ("max_posts", "Số mục tối đa mỗi lần chạy (0 = chạy hết)"),
            ("delay_between_groups", "Nghỉ giữa 2 group (giây)"),
            ("delay_between_rooms", "Nghỉ giữa 2 mục (giây)"),
            ("max_retries", "Số lần thử lại khi lỗi"),
        ]
        self.num_vars = {}
        for i, (key, label) in enumerate(nums):
            r = 12 + i
            ttk.Label(f, text=label).grid(row=r, column=0, sticky="w", padx=6, pady=4)
            var = tk.StringVar()
            self.num_vars[key] = var
            ttk.Entry(f, textvariable=var, width=10).grid(row=r, column=1, sticky="w", padx=6)

        self.chong_trung_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            f, text="Chống đăng trùng — không đăng lại bài đã lên nhóm đó trong cùng ngày",
            variable=self.chong_trung_var,
        ).grid(row=16, column=0, columnspan=3, sticky="w", padx=6, pady=(12, 2))
        ttk.Label(
            f,
            text="Tắt ô này thì tool đăng bất chấp lịch sử — cùng bài có thể lên "
                 "cùng nhóm nhiều lần trong ngày, Facebook rất dễ đánh dấu spam.",
            foreground="#a05000", wraplength=560, justify="left",
        ).grid(row=17, column=0, columnspan=3, sticky="w", padx=26, pady=(0, 4))

        ttk.Button(f, text="💾 Lưu cấu hình chiến dịch", command=self.save_paths) \
            .grid(row=20, column=0, sticky="w", padx=6, pady=16)
        f.columnconfigure(1, weight=1)

    def _pick_path(self, var, kind):
        current = var.get()
        start = current if os.path.isdir(current) else os.path.dirname(current)
        picked = (filedialog.askdirectory(initialdir=start) if kind == "dir"
                  else filedialog.askopenfilename(initialdir=start))
        if picked:
            var.set(picked)

    def save_paths(self):
        if self.dang_chay():
            return
        camp = self.campaign()

        # Hai chiến dịch dùng chung file log = chiến dịch này coi bài của chiến
        # dịch kia là đã đăng (khóa chống trùng chỉ gồm STT + link nhóm, không
        # có tên chiến dịch). Hậu quả im lặng: bài trùng STT sẽ KHÔNG BAO GIỜ
        # được đăng, mà log chỉ ghi "đã đăng hôm nay".
        log_moi = self.path_vars["posted_log"].get().strip()
        ten_hien_tai = self.cfg["active_campaign"]
        for ten, c in self.cfg["campaigns"].items():
            if ten != ten_hien_tai and log_moi and c.get("posted_log") == log_moi:
                messagebox.showerror(
                    "Trùng file lịch sử đăng",
                    f"Chiến dịch '{ten}' đang dùng đúng file này làm 'File log "
                    f"bài đã đăng':\n{log_moi}\n\n"
                    "Hai chiến dịch dùng chung một file sẽ coi bài của nhau là "
                    "đã đăng — bài trùng STT sẽ không bao giờ được đăng lên.\n\n"
                    "Hãy chọn một file khác, ví dụ thêm đuôi tên chiến dịch.")
                return

        for key, var in self.path_vars.items():
            camp[key] = var.get().strip()
        # Delay quá thấp là con đường nhanh nhất để bị Facebook khoá nick, nên
        # kẹp sàn đúng bằng giá trị đã chạy ổn định trước đây.
        SAN = {"delay_between_groups": 10, "delay_between_rooms": 10, "max_retries": 1, "max_posts": 0}
        # Trần cho max_retries: mỗi lần thử lại là một lần thao tác thật lên
        # Facebook. Nếu bài thực ra đã đăng được mà tool tưởng hỏng (hay gặp khi
        # dialog không đóng kịp), retry cao sẽ rải nhiều bài trùng lên cùng group.
        TRAN = {"max_retries": 3}
        for key, var in self.num_vars.items():
            try:
                gia_tri = int(var.get().strip())
            except ValueError:
                messagebox.showerror("Sai định dạng", f"'{var.get()}' không phải số nguyên.")
                return
            if gia_tri < SAN[key]:
                messagebox.showwarning(
                    "Giá trị quá thấp",
                    f"'{key}' tối thiểu là {SAN[key]}, đã tự nâng lên.\n\n"
                    "Đặt delay thấp hơn rất dễ bị Facebook đánh dấu spam và khoá nick.",
                )
                gia_tri = SAN[key]
                var.set(str(gia_tri))
            if key in TRAN and gia_tri > TRAN[key]:
                messagebox.showwarning(
                    "Giá trị quá cao",
                    f"'{key}' tối đa là {TRAN[key]}, đã tự hạ xuống.\n\n"
                    "Thử lại nhiều lần dễ làm cùng một bài bị đăng lặp lên một "
                    "group khi tool hiểu nhầm là đăng hỏng.",
                )
                gia_tri = TRAN[key]
                var.set(str(gia_tri))
            camp[key] = gia_tri
        camp["chong_trung"] = self.chong_trung_var.get()
        if not camp["chong_trung"]:
            messagebox.showwarning(
                "Đã tắt chống đăng trùng",
                f"Chiến dịch '{self.cfg['active_campaign']}' sẽ đăng bất chấp "
                "lịch sử: chạy bao nhiêu lần thì cùng một bài lên cùng một nhóm "
                "bấy nhiêu lần.\n\nChỉ nên tắt tạm khi cần đăng lại gấp, xong "
                "nhớ bật lại.")
        self.cfg["profiles_dir"] = self.profiles_var.get().strip()
        cfg_module.save_config(self.cfg)
        self.refresh_accounts()
        # Đường dẫn Excel có thể vừa đổi — phải nạp lại bảng, nếu không lần
        # bấm Ghi sau sẽ đổ nội dung file cũ đè lên file mới.
        self.load_excel()
        messagebox.showinfo("Đã lưu", f"Đã lưu cấu hình chiến dịch '{self.cfg['active_campaign']}'.")

    # ==================== TAB: NHÓM THEO PHÂN LOẠI ====================

    def _build_groups_tab(self):
        f = self.tab_groups

        left = ttk.LabelFrame(f, text="Phân loại")
        left.pack(side="left", fill="y", padx=(0, 8), pady=4)

        self.cat_list = tk.Listbox(left, width=24, exportselection=False)
        self.cat_list.pack(fill="y", expand=True, padx=6, pady=6)
        self.cat_list.bind("<<ListboxSelect>>", self.on_select_category)

        ttk.Button(left, text="+ Thêm phân loại", command=self.add_category).pack(fill="x", padx=6, pady=2)
        ttk.Button(left, text="Đổi tên", command=self.rename_category).pack(fill="x", padx=6, pady=2)
        ttk.Button(left, text="Xóa phân loại", command=self.delete_category).pack(fill="x", padx=6, pady=(2, 8))

        right = ttk.LabelFrame(f, text="Link group (mỗi dòng 1 link)")
        right.pack(side="left", fill="both", expand=True, pady=4)

        ttk.Label(
            right,
            text="Tên phân loại phải khớp với cột C trong Excel (so khớp bỏ dấu, "
                 "không phân biệt hoa thường).",
            foreground="#555",
        ).pack(anchor="w", padx=6, pady=(6, 0))

        self.url_box = tk.Text(right, wrap="none", font=("Menlo", 11))
        self.url_box.pack(fill="both", expand=True, padx=6, pady=6)

        ttk.Button(right, text="💾 Lưu danh sách group", command=self.save_groups) \
            .pack(anchor="w", padx=6, pady=(0, 8))

    def refresh_categories(self):
        self.cat_list.delete(0, "end")
        for cat in self.campaign()["groups"]:
            self.cat_list.insert("end", cat)

    def on_select_category(self, _event=None):
        sel = self.cat_list.curselection()
        if not sel:
            return
        cat = self.cat_list.get(sel[0])
        self.current_cat = cat
        self.url_box.delete("1.0", "end")
        self.url_box.insert("1.0", "\n".join(self.campaign()["groups"][cat]))

    def add_category(self):
        if self.dang_chay():
            return
        name = SimplePrompt(self, "Thêm phân loại",
                            "Tên phân loại (vd: Cầu Giấy, Kế toán, Chủ đề A):").result
        if not name:
            return
        groups = self.campaign()["groups"]
        if name in groups:
            messagebox.showinfo("Đã có", f"Phân loại '{name}' đã tồn tại.")
            return
        groups[name] = []
        cfg_module.save_config(self.cfg)
        self.refresh_categories()

    def rename_category(self):
        if self.dang_chay():
            return
        if not self.current_cat:
            messagebox.showwarning("Chưa chọn", "Hãy chọn một phân loại.")
            return
        new = SimplePrompt(self, "Đổi tên", f"Tên mới cho '{self.current_cat}':", self.current_cat).result
        if not new or new == self.current_cat:
            return
        camp = self.campaign()
        # Dựng lại dict để giữ nguyên thứ tự các phân loại
        camp["groups"] = {
            (new if k == self.current_cat else k): v for k, v in camp["groups"].items()
        }
        cfg_module.save_config(self.cfg)
        self.current_cat = new
        self.refresh_categories()

    def delete_category(self):
        if self.dang_chay():
            return
        if not self.current_cat:
            messagebox.showwarning("Chưa chọn", "Hãy chọn một phân loại.")
            return
        if not messagebox.askyesno("Xóa", f"Xóa phân loại '{self.current_cat}' và toàn bộ link group của nó?"):
            return
        self.campaign()["groups"].pop(self.current_cat, None)
        cfg_module.save_config(self.cfg)
        self.current_cat = None
        self.url_box.delete("1.0", "end")
        self.refresh_categories()

    def save_groups(self):
        if self.dang_chay():
            return
        if not self.current_cat:
            messagebox.showwarning("Chưa chọn", "Hãy chọn một phân loại trước khi lưu.")
            return
        urls = [u.strip() for u in self.url_box.get("1.0", "end").split("\n") if u.strip()]
        self.campaign()["groups"][self.current_cat] = urls
        cfg_module.save_config(self.cfg)
        messagebox.showinfo("Đã lưu", f"Phân loại '{self.current_cat}': {len(urls)} group.")

    # ==================== TAB: HẸN GIỜ ĐĂNG ====================

    def _build_lich_tab(self):
        f = self.tab_lich

        note = ttk.LabelFrame(f, text="Cách hoạt động")
        note.pack(fill="x", padx=4, pady=(6, 4))
        ttk.Label(
            note,
            text="• Giờ hẹn tính theo giờ Việt Nam (UTC+7), không phụ thuộc múi giờ máy.\n"
                 "• Đến giờ, máy phải đang bật và tool đang mở thì bài mới đăng được.\n"
                 "• Mở tool muộn hơn giờ hẹn: tool sẽ hỏi lại bạn có chạy bù không.",
            justify="left",
        ).pack(anchor="w", padx=10, pady=8)

        dong_ho = ttk.Frame(f)
        dong_ho.pack(fill="x", padx=4)
        self.lbl_dong_ho = ttk.Label(dong_ho, text="", font=("Menlo", 12))
        self.lbl_dong_ho.pack(side="left", pady=4)

        self.tu_mo_var = tk.BooleanVar(value=lich_hen.dang_tu_mo())
        self.chk_tu_mo = ttk.Checkbutton(
            dong_ho, text="Tự mở tool khi bật máy (thu nhỏ sẵn)",
            variable=self.tu_mo_var, command=self.doi_tu_mo)
        self.chk_tu_mo.pack(side="right", pady=4)
        if not lich_hen.ho_tro_tu_mo():
            self.chk_tu_mo.config(state="disabled")

        cot = ("bat", "lich", "camp", "nick", "ke_tiep", "chay_cuoi")
        self.lich_tree = ttk.Treeview(f, columns=cot, show="headings", height=12)
        for ma, ten, rong in (
            ("bat", "Bật", 50), ("lich", "Lịch", 200), ("camp", "Chiến dịch", 150),
            ("nick", "Nick", 130), ("ke_tiep", "Lần chạy tới", 140),
            ("chay_cuoi", "Chạy gần nhất", 140),
        ):
            self.lich_tree.heading(ma, text=ten)
            self.lich_tree.column(ma, width=rong, anchor="w")
        self.lich_tree.pack(fill="both", expand=True, padx=4, pady=6)
        self.lich_tree.bind("<Double-1>", lambda _e: self.sua_lich())

        bar = ttk.Frame(f)
        bar.pack(fill="x", padx=4, pady=(0, 8))
        ttk.Button(bar, text="+ Thêm lịch", command=self.them_lich).pack(side="left")
        ttk.Button(bar, text="Sửa", command=self.sua_lich).pack(side="left", padx=4)
        ttk.Button(bar, text="Bật / Tắt", command=self.doi_bat_lich).pack(side="left", padx=4)
        ttk.Button(bar, text="Xóa", command=self.xoa_lich).pack(side="left", padx=4)

    def _nhip_dong_ho(self):
        """Đồng hồ giờ Việt Nam, đập mỗi giây — để người dùng đối chiếu ngay
        được giờ hẹn với giờ thật, khỏi phải đoán máy mình lệch múi giờ hay
        không."""
        bg = lich_hen.bay_gio()
        self.lbl_dong_ho.config(
            text=f"🕐 Giờ Việt Nam bây giờ: {bg.strftime('%H:%M:%S — %d/%m/%Y')}")
        self.after(1000, self._nhip_dong_ho)

    def danh_sach_lich(self):
        """Danh sách lịch hẹn, đã chuẩn hóa và bảo đảm id không trùng.

        Người dùng copy nguyên một khối lịch trong config.json để nhân đôi là
        chuyện có thật; hai lịch cùng id sẽ làm bảng lịch ném TclError và hỏng
        cả tab.
        """
        ds = self.cfg.setdefault("schedules", [])
        da_gap = set()
        for l in ds:
            lich_hen.chuan_hoa(l)
            if l["id"] in da_gap:
                l["id"] = lich_hen.lich_moi()["id"]
            da_gap.add(l["id"])
        return ds

    def refresh_lich(self):
        """Vẽ lại bảng lịch từ cấu hình."""
        if not hasattr(self, "lich_tree"):
            return
        self.lich_tree.delete(*self.lich_tree.get_children())
        for l in self.danh_sach_lich():
            ke_tiep = lich_hen.lan_ke_tiep(l) if l.get("bat") else None
            self.lich_tree.insert("", "end", iid=l["id"], values=(
                "✓" if l.get("bat") else "—",
                lich_hen.mo_ta(l),
                l.get("campaign") or "?",
                l.get("nick") or "?",
                lich_hen.mo_ta_moc(ke_tiep),
                lich_hen.mo_ta_moc(l.get("lan_chay_cuoi")),
            ))

    def _lich_dang_chon(self):
        sel = self.lich_tree.selection()
        if not sel:
            messagebox.showwarning("Chưa chọn", "Hãy chọn một lịch trong bảng.")
            return None
        return next((l for l in self.danh_sach_lich() if l["id"] == sel[0]), None)

    def _luu_lich(self):
        cfg_module.save_config(self.cfg)
        self.refresh_lich()

    def them_lich(self):
        if self.dang_chay():
            return
        moi = LichEditor(self, "Thêm lịch hẹn", lich_hen.lich_moi(),
                         list(self.cfg["campaigns"]), self.ten_cac_nick()).result
        if moi:
            self.danh_sach_lich().append(moi)
            self._luu_lich()
            self.append_log(f"⏰ Đã thêm lịch: {lich_hen.mo_ta(moi)} — {moi['campaign']} / {moi['nick']}")

    def sua_lich(self):
        if self.dang_chay():
            return
        l = self._lich_dang_chon()
        if not l:
            return
        sua = LichEditor(self, "Sửa lịch hẹn", dict(l),
                         list(self.cfg["campaigns"]), self.ten_cac_nick()).result
        if sua:
            # Đổi giờ giấc thì mốc đã xử lý của lịch cũ không còn ý nghĩa
            sua["moc_da_xu_ly"] = ""
            l.clear()
            l.update(sua)
            self._luu_lich()

    def doi_bat_lich(self):
        l = self._lich_dang_chon()
        if not l:
            return
        l["bat"] = not l.get("bat")
        # Bật lại một lịch cũ thì xóa dấu mốc đã xử lý, nếu không lịch sẽ im
        # lặng bỏ qua đúng lần chạy gần nhất mà người dùng vừa mong đợi.
        if l["bat"]:
            l["moc_da_xu_ly"] = ""
        self._luu_lich()

    def xoa_lich(self):
        l = self._lich_dang_chon()
        if not l:
            return
        if not messagebox.askyesno("Xóa lịch", f"Xóa lịch '{lich_hen.mo_ta(l)}'?"):
            return
        self.danh_sach_lich().remove(l)
        self._luu_lich()

    def ten_cac_nick(self):
        return [ten for ten, _ in bot.list_sessions()]

    def doi_tu_mo(self):
        try:
            if self.tu_mo_var.get():
                lich_hen.bat_tu_mo()
                self.append_log("⏰ Đã bật: Windows sẽ tự mở tool khi khởi động máy.")
            else:
                lich_hen.tat_tu_mo()
                self.append_log("⏰ Đã tắt tự mở tool khi bật máy.")
        except Exception as e:
            self.tu_mo_var.set(lich_hen.dang_tu_mo())
            messagebox.showerror("Không đặt được", f"Không đổi được cài đặt tự mở:\n{e}")

    # ==================== VÒNG LẶP KIỂM TRA LỊCH ====================

    def _vong_kiem_tra_lich(self):
        """Cứ 20 giây soát một lượt xem có lịch nào tới giờ chưa.

        Dùng after() của Tk chứ không phải thread riêng: mọi thao tác đọc/ghi
        cấu hình và dựng hộp thoại đều phải nằm ở main thread.
        """
        try:
            self._soat_lich()
        except Exception as e:
            self.append_log(f"⏰ Lỗi khi soát lịch hẹn: {e}")
        self.after(20000, self._vong_kiem_tra_lich)

    def _soat_lich(self):
        # Hộp thoại "chạy bù?" chạy vòng lặp sự kiện lồng nhau, after() vẫn nổ
        # trong lúc nó mở — không có cờ này thì mỗi 20 giây lại chồng thêm một
        # hộp thoại nữa lên màn hình.
        if getattr(self, "_dang_hoi_lich", False):
            return

        # Đang bận thì hoãn lượt chứ không đánh dấu đã xử lý — soát lại sau 20
        # giây. Người dùng đang gõ dở trong hộp thoại sửa bài mà lịch nổ thì
        # set_ui_locked() khoá luôn nút Lưu của hộp thoại đó, họ kẹt giữa chừng
        # không thoát ra được.
        ban = None
        if self.dang_khoa or (self.worker and self.worker.is_alive()):
            ban = "đang có tiến trình chạy"
        elif self._co_hop_thoai_dang_mo():
            ban = "đang mở một hộp thoại"

        bg = lich_hen.bay_gio()

        if ban:
            # Báo một lần cho mỗi mốc, nếu không thì cứ 20 giây một dòng log
            da_bao = self.__dict__.setdefault("_da_bao_ban", set())
            for l in self.danh_sach_lich():
                lich_hen.chuan_hoa(l)
                hanh_dong, moc = lich_hen.trang_thai(l, bg)
                if hanh_dong in ("chay", "hoi") and (l["id"], moc) not in da_bao:
                    da_bao.add((l["id"], moc))
                    self.append_log(
                        f"⏰ Lịch {lich_hen.mo_ta_moc(moc)} ({l['campaign']}) phải "
                        f"chờ vì {ban} — sẽ chạy ngay khi xong.")
            return
        thay_doi = False
        for l in list(self.danh_sach_lich()):
            lich_hen.chuan_hoa(l)
            hanh_dong, moc = lich_hen.trang_thai(l, bg)
            if not hanh_dong:
                continue

            if hanh_dong == "qua_han":
                lich_hen.danh_dau_da_xu_ly(l, moc)
                l["bat"] = False        # chỉ xảy ra với lịch một lần, đã lỡ hẳn
                self.append_log(
                    f"⏰ Bỏ qua lịch {lich_hen.mo_ta_moc(moc)} ({l['campaign']}) — "
                    "đã quá hạn quá lâu, không chạy bù. Lịch được tắt.")
                thay_doi = True
                continue

            if hanh_dong == "hoi":
                self._dang_hoi_lich = True
                try:
                    self.deiconify()
                    self.lift()
                    dong_y = messagebox.askyesno(
                        "Lịch hẹn đã lỡ giờ",
                        f"Lịch {lich_hen.mo_ta_moc(moc)} (chiến dịch "
                        f"'{l['campaign']}', nick '{l['nick']}') đã tới giờ lúc "
                        "tool chưa mở.\n\nChạy bù ngay bây giờ?")
                finally:
                    self._dang_hoi_lich = False
                lich_hen.danh_dau_da_xu_ly(l, moc)
                thay_doi = True
                if not dong_y:
                    self.append_log(f"⏰ Bạn đã bỏ qua lịch {lich_hen.mo_ta_moc(moc)}.")
                    continue

            # Tới đây là chạy: đúng giờ, hoặc người dùng đồng ý chạy bù.
            # Chỉ ghi "đã chạy" khi thật sự khởi động được — bỏ lượt vì thiếu
            # file hay sai nick mà vẫn ghi thì bảng lịch báo "chạy lúc 19:30"
            # trong khi chẳng có bài nào lên.
            chay_duoc = self._chay_theo_lich(l, moc)
            lich_hen.danh_dau_da_xu_ly(l, moc, da_chay=chay_duoc)
            # Lịch một lần thì mốc của nó đã trôi qua, có chạy được hay không
            # cũng không bao giờ tới lượt nữa
            if l.get("kieu") == "mot_lan":
                l["bat"] = False
            cfg_module.save_config(self.cfg)
            self.refresh_lich()
            return      # mỗi lượt chỉ khởi động một lần chạy

        if thay_doi:
            cfg_module.save_config(self.cfg)
            self.refresh_lich()

    def _co_hop_thoai_dang_mo(self):
        """Có cửa sổ con nào (sửa bài, đăng nhập nick, sửa lịch) đang mở không."""
        return any(
            isinstance(w, tk.Toplevel) and w.winfo_exists() and w.winfo_viewable()
            for w in self.winfo_children()
        )

    def _chay_theo_lich(self, l, moc):
        """Khởi động một lần đăng do lịch hẹn kích hoạt.

        Khác nút bấm tay ở chỗ: không dựng hộp thoại hỏi han. Đến giờ hẹn có
        thể chẳng có ai ngồi trước máy, một hộp thoại chờ bấm OK sẽ treo cả
        lượt chạy tới sáng hôm sau. Vướng gì thì ghi log rồi bỏ lượt.

        Trả về True nếu đã khởi động được lần chạy.
        """
        camp_name, nick = l.get("campaign"), l.get("nick")

        def bo_qua(ly_do):
            self.append_log(f"⏰ Không chạy được lịch {lich_hen.mo_ta_moc(moc)}: {ly_do}")
            return False

        if camp_name not in self.cfg["campaigns"]:
            return bo_qua(f"chiến dịch '{camp_name}' không còn tồn tại.")
        path = os.path.join(self.cfg["profiles_dir"], nick or "")
        if not nick or not os.path.isdir(path):
            return bo_qua(f"không tìm thấy nick '{nick}' trong thư mục profile.")

        camp = self.cfg["campaigns"][camp_name]
        if not camp["groups"]:
            return bo_qua(f"chiến dịch '{camp_name}' chưa khai báo nhóm nào.")

        if not os.path.exists(camp["excel_path"]):
            return bo_qua(f"thiếu file Excel: {camp['excel_path']}")

        van_de = bot.kiem_tra_du_lieu(camp_name)
        loi = [v for v in van_de if v["muc"] == "loi"]
        if loi:
            self.hien_van_de(van_de, camp_name)
            return bo_qua(f"dữ liệu có {len(loi)} lỗi (xem chi tiết ở trên).")

        # Kéo giao diện về đúng chiến dịch của lịch, để log và các tab đang
        # hiện không nói một đằng còn worker chạy một nẻo.
        #
        # Trừ khi tab 'Nội dung bài đăng' đang có sửa chưa ghi vào Excel: nạp
        # lại bảng sẽ nuốt sạch phần người dùng vừa gõ, mà họ không hề bấm gì
        # cả. Bỏ việc chuyển giao diện đi thì lượt chạy vẫn đúng — worker đọc
        # cấu hình theo tên chiến dịch từ file, không đọc từ các ô trên màn
        # hình.
        if self.cfg["active_campaign"] != camp_name:
            if self.co_thay_doi_chua_luu():
                self.append_log(
                    f"⏰ Vẫn chạy chiến dịch '{camp_name}' theo lịch, nhưng giữ "
                    "nguyên giao diện vì tab 'Nội dung bài đăng' đang có sửa "
                    "chưa lưu — nhớ bấm 💾 Ghi vào file Excel.")
            else:
                self.cfg["active_campaign"] = camp_name
                cfg_module.save_config(self.cfg)
                self.refresh_campaign_box()
                self.load_campaign_into_views()

        ds_nick = self.acc_list.get(0, "end")
        if nick in ds_nick:
            self.acc_list.selection_clear(0, "end")
            self.acc_list.selection_set(ds_nick.index(nick))

        self.append_log(
            f"⏰ Tới giờ hẹn {lich_hen.mo_ta_moc(moc)} — chạy chiến dịch "
            f"'{camp_name}' bằng nick '{nick}'.")
        self._khoi_dong_worker(nick, path, camp_name, tu_lich=True)
        return True

    # ==================== TAB: THAM GIA NHÓM ====================
    # Tách hẳn khỏi việc đăng bài: danh sách nhóm riêng, nick chọn riêng, log
    # riêng, nút chạy riêng. Đăng bài không bao giờ tự đi vào nhóm và ngược lại.

    def _build_join_tab(self):
        f = self.tab_join

        # --- Cột trái: chọn nick ---
        left = ttk.LabelFrame(f, text="Nick sẽ đi tham gia nhóm")
        left.pack(side="left", fill="y", padx=(0, 8), pady=4)

        self.join_nick_list = tk.Listbox(left, width=24, exportselection=False)
        self.join_nick_list.pack(fill="y", expand=True, padx=6, pady=6)
        self.join_nick_list.bind("<<ListboxSelect>>", lambda _e: self.refresh_join_tree())

        ttk.Label(
            left,
            text="Lịch sử tham gia ghi theo từng\nnick, nằm trong profile của\nnick đó.",
            foreground="#555", justify="left",
        ).pack(anchor="w", padx=6, pady=(0, 8))

        right = ttk.Frame(f)
        right.pack(side="left", fill="both", expand=True, pady=4)

        # --- Danh sách nhóm + trạng thái theo nick đang chọn ---
        khung_ds = ttk.LabelFrame(right, text="Danh sách nhóm muốn tham gia")
        khung_ds.pack(fill="both", expand=True)

        thanh = ttk.Frame(khung_ds)
        thanh.pack(fill="x", padx=6, pady=6)
        ttk.Button(thanh, text="+ Dán link nhóm", command=self.them_link_nhom).pack(side="left")
        ttk.Button(thanh, text="Nạp từ chiến dịch...", command=self.nap_link_tu_chien_dich) \
            .pack(side="left", padx=4)
        ttk.Button(thanh, text="📋 Sao chép link", command=self.sao_chep_link).pack(side="left")
        ttk.Button(thanh, text="☑ Tick hết", command=lambda: self.tick_tat_ca(True)) \
            .pack(side="left", padx=(4, 0))
        ttk.Button(thanh, text="☐ Bỏ tick", command=lambda: self.tick_tat_ca(False)) \
            .pack(side="left", padx=4)
        ttk.Button(thanh, text="Xóa nhóm đã tick", command=self.xoa_link_da_chon).pack(side="left")
        ttk.Button(thanh, text="Xóa hết", command=self.xoa_het_link).pack(side="left", padx=4)
        self.join_dem = ttk.Label(thanh, text="", foreground="#555")
        self.join_dem.pack(side="left", padx=8)

        # Treeview không có checkbox thật, nên cột đầu là ô chữ ☐/☑ và bắt
        # click vào đúng cột đó. Dùng ô tick thay vì "dòng đang bôi đen" vì bôi
        # đen mất ngay khi bảng vẽ lại hoặc khi bấm chỗ khác — người dùng tick
        # 20 nhóm xong quay ra bấm nút thì mất sạch.
        cot = ("chon", "link", "trang_thai", "thoi_gian")
        self.join_tree = ttk.Treeview(khung_ds, columns=cot, show="headings",
                                      height=9, selectmode="extended")
        self.join_tree.heading("chon", text="✓", command=self.doi_tick_tat_ca)
        self.join_tree.heading("link", text="Link nhóm")
        self.join_tree.heading("trang_thai", text="Trạng thái với nick đang chọn")
        self.join_tree.heading("thoi_gian", text="Lần gần nhất")
        self.join_tree.column("chon", width=34, anchor="center", stretch=False)
        self.join_tree.column("link", width=360)
        self.join_tree.column("trang_thai", width=180)
        self.join_tree.column("thoi_gian", width=130)
        self.join_tree.bind("<Button-1>", self._bam_vao_bang)
        self.join_tree.bind("<space>", lambda _e: self._doi_tick(self.join_tree.selection()))
        self.join_tree.bind("<Control-c>", lambda _e: self.sao_chep_link())
        self.join_tree.bind("<Command-c>", lambda _e: self.sao_chep_link())
        # Chuột phải trên Windows là Button-3, trên Mac có máy ra Button-2
        for phim in ("<Button-3>", "<Button-2>"):
            self.join_tree.bind(phim, self._menu_chuot_phai)

        self.menu_nhom = tk.Menu(self, tearoff=0)
        self.menu_nhom.add_command(label="Sao chép link", command=self.sao_chep_link)
        self.menu_nhom.add_command(label="Mở nhóm trong trình duyệt", command=self.mo_nhom_tren_web)
        cuon = ttk.Scrollbar(khung_ds, command=self.join_tree.yview)
        self.join_tree.configure(yscrollcommand=cuon.set)
        cuon.pack(side="right", fill="y")
        self.join_tree.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        # --- Cài đặt né ban ---
        khung_cd = ttk.LabelFrame(right, text="Cài đặt an toàn (dùng chung mọi nick)")
        khung_cd.pack(fill="x", pady=6)

        o_so = [
            ("max_moi_lan", "Tối đa mỗi lần chạy"),
            ("max_moi_ngay", "Tối đa mỗi nick mỗi ngày"),
            ("nghi_min", "Nghỉ giữa 2 nhóm — ít nhất (giây)"),
            ("nghi_max", "— nhiều nhất (giây)"),
            ("nghi_dai_sau", "Nghỉ dài sau mỗi (nhóm)"),
            ("nghi_dai_phut", "Nghỉ dài bao lâu (phút)"),
            ("gio_bat_dau", "Chỉ chạy từ (giờ VN)"),
            ("gio_ket_thuc", "đến (giờ VN)"),
            ("gio_tam_nghi", "Bị chặn thì nghỉ (giờ)"),
        ]
        self.join_vars = {}
        for i, (key, nhan) in enumerate(o_so):
            hang, cot_i = divmod(i, 2)
            ttk.Label(khung_cd, text=nhan).grid(row=hang, column=cot_i * 2,
                                                sticky="w", padx=6, pady=3)
            var = tk.StringVar()
            self.join_vars[key] = var
            ttk.Entry(khung_cd, textvariable=var, width=8) \
                .grid(row=hang, column=cot_i * 2 + 1, sticky="w", padx=6)
        self.nick_moi_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            khung_cd, text="Nick mới lập — ép trần xuống rất thấp (2 nhóm/lần, 3 nhóm/ngày)",
            variable=self.nick_moi_var,
        ).grid(row=5, column=0, columnspan=4, sticky="w", padx=6, pady=(6, 2))

        ttk.Button(khung_cd, text="💾 Lưu cài đặt", command=self.luu_cai_dat_join) \
            .grid(row=0, column=4, rowspan=3, padx=12)
        ttk.Button(khung_cd, text="Bỏ tạm nghỉ của nick", command=self.bo_tam_nghi_nick) \
            .grid(row=3, column=4, padx=12)
        ttk.Button(khung_cd, text="🔄 Soát lại trạng thái", command=self.soat_lai_trang_thai) \
            .grid(row=4, column=4, padx=12)

        ttk.Label(
            khung_cd,
            text="Tool xáo thứ tự nhóm, xem trang vài giây rồi mới bấm, thỉnh thoảng "
                 "ghé bảng tin, nghỉ ngẫu nhiên giữa các nhóm và nghỉ dài sau mỗi vài "
                 "nhóm. Bấm xong luôn kiểm chứng nút có đổi trạng thái không — 2 lần "
                 "bấm không ăn thua là coi như đang bị chặn ngầm: dừng và khoá nick "
                 "lại vài chục tiếng, không cho chạy tiếp.",
            foreground="#a05000", wraplength=640, justify="left",
        ).grid(row=6, column=0, columnspan=5, sticky="w", padx=6, pady=(4, 6))

        # --- Nút chạy + log riêng ---
        thanh2 = ttk.Frame(right)
        thanh2.pack(fill="x", pady=(0, 4))
        self.btn_join_chon = ttk.Button(thanh2, text="▶ Tham gia nhóm đã tick",
                                        command=lambda: self.bat_dau_join(chi_chon=True))
        self.btn_join_chon.pack(side="left")
        self.btn_join_all = ttk.Button(thanh2, text="▶ Tham gia tất cả nhóm chưa vào",
                                       command=lambda: self.bat_dau_join(chi_chon=False))
        self.btn_join_all.pack(side="left", padx=6)
        self.btn_join_stop = ttk.Button(thanh2, text="■ Dừng",
                                        command=self.request_stop, state="disabled")
        self.btn_join_stop.pack(side="left")
        ttk.Button(thanh2, text="Xóa màn hình log",
                   command=lambda: self.join_log_box.delete("1.0", "end")).pack(side="left", padx=6)

        self.join_log_box = tk.Text(right, wrap="word", height=10, font=("Menlo", 11))
        cuon2 = ttk.Scrollbar(right, command=self.join_log_box.yview)
        self.join_log_box.configure(yscrollcommand=cuon2.set)
        cuon2.pack(side="right", fill="y")
        self.join_log_box.pack(fill="both", expand=True)

        self.nap_cai_dat_join()

    # ---------- dữ liệu của tab ----------

    def nap_cai_dat_join(self):
        """Đổ cài đặt tham gia nhóm từ config lên các ô."""
        cai_dat = joiner.load_cai_dat()
        for key, var in self.join_vars.items():
            var.set(str(cai_dat[key]))
        self.nick_moi_var.set(bool(cai_dat["nick_moi"]))
        self.refresh_join_tree()

    def refresh_join_nicks(self):
        """Nạp lại danh sách nick cho tab tham gia nhóm."""
        self.join_nick_list.delete(0, "end")
        for name, _ in bot.list_sessions():
            self.join_nick_list.insert("end", name)
        self.refresh_join_tree()

    def join_nick_dang_chon(self):
        """(tên, đường dẫn profile) của nick đang chọn ở tab này, hoặc None."""
        sel = self.join_nick_list.curselection()
        if not sel:
            return None
        name = self.join_nick_list.get(sel[0])
        return name, os.path.join(self.cfg["profiles_dir"], name)

    def refresh_join_tree(self):
        """Vẽ lại bảng nhóm kèm trạng thái theo nick đang chọn."""
        if not hasattr(self, "join_tree"):
            return
        self.join_tree.delete(*self.join_tree.get_children())

        cai_dat = joiner.load_cai_dat()
        links = cai_dat["links"]
        acc = self.join_nick_dang_chon()
        lich_su = joiner.load_log(acc[1]) if acc else joiner.log_rong()

        # Bỏ khỏi danh sách tick những link không còn trong bảng nữa
        self.join_da_tick &= {joiner.clean_group_url(u) for u in links}

        for url in links:
            muc = joiner.muc_nhom(lich_su, url) if acc else {}
            mo_ta = joiner.mo_ta_trang_thai(muc) if acc else "—"
            danh_dau = "☑" if joiner.clean_group_url(url) in self.join_da_tick else "☐"
            self.join_tree.insert("", "end",
                                  values=(danh_dau, url, mo_ta, muc.get("time", "")))

        if not acc:
            self.join_dem.config(
                text=f"{len(links)} nhóm — đã tick {len(self.join_da_tick)} | "
                     "chọn nick để xem trạng thái")
            return

        can_vao = len(joiner.loc_nhom_can_vao(links, lich_su))
        hom_nay = joiner.dem_da_xin_hom_nay(lich_su)
        dang_nghi, den, ly_do = joiner.dang_tam_nghi(lich_su)

        chu = (f"{len(links)} nhóm — {can_vao} cần xử lý | đã tick "
               f"{len(self.join_da_tick)} | nick '{acc[0]}' hôm nay đã bấm "
               f"{hom_nay}/{cai_dat['max_moi_ngay']}")
        if dang_nghi:
            chu += f"  ⛔ ĐANG TẠM NGHỈ tới {den.strftime('%d/%m %H:%M')} ({ly_do})"
        elif not joiner.trong_gio_hoat_dong(cai_dat):
            chu += (f"  ⏰ ngoài khung giờ {cai_dat['gio_bat_dau']}h–"
                    f"{cai_dat['gio_ket_thuc']}h")
        self.join_dem.config(text=chu)

    def _link_cua_dong(self, dong):
        return self.join_tree.item(dong, "values")[1]

    def _bam_vao_bang(self, event):
        """Bấm vào bất kỳ đâu trên một dòng là đổi tick dòng đó.

        Không trả về "break": để Treeview vẫn bôi đen dòng vừa bấm như thường,
        nhờ vậy nút 'Sao chép link' và Ctrl+C biết đang nói tới dòng nào.
        """
        if self.dang_khoa:
            return None
        if self.join_tree.identify_region(event.x, event.y) not in ("cell", "tree"):
            return None
        dong = self.join_tree.identify_row(event.y)
        if not dong:
            return None
        self._doi_tick([dong])
        return None

    def _doi_tick(self, cac_dong):
        """Đảo trạng thái tick của các dòng (đang tick → bỏ, chưa → tick)."""
        if self.dang_khoa:
            return
        for dong in cac_dong:
            url = joiner.clean_group_url(self._link_cua_dong(dong))
            if url in self.join_da_tick:
                self.join_da_tick.discard(url)
                self.join_tree.set(dong, "chon", "☐")
            else:
                self.join_da_tick.add(url)
                self.join_tree.set(dong, "chon", "☑")
        self._cap_nhat_dem_tick()

    def tick_tat_ca(self, bat):
        """Tick hết hoặc bỏ tick hết mọi dòng đang hiện."""
        if self.dang_khoa:
            return
        for dong in self.join_tree.get_children():
            url = joiner.clean_group_url(self._link_cua_dong(dong))
            if bat:
                self.join_da_tick.add(url)
            else:
                self.join_da_tick.discard(url)
            self.join_tree.set(dong, "chon", "☑" if bat else "☐")
        self._cap_nhat_dem_tick()

    def doi_tick_tat_ca(self):
        """Bấm vào tiêu đề cột ✓: chưa tick hết thì tick hết, tick hết rồi thì bỏ."""
        tong = len(self.join_tree.get_children())
        self.tick_tat_ca(len(self.join_da_tick) < tong)

    def _cap_nhat_dem_tick(self):
        """Sửa mỗi con số 'đã tick' trong dòng chữ, khỏi vẽ lại cả bảng."""
        chu = self.join_dem["text"]
        if "đã tick" in chu:
            dau = chu.index("đã tick")
            cuoi = chu.index("|", dau) if "|" in chu[dau:] else len(chu)
            self.join_dem.config(
                text=chu[:dau] + f"đã tick {len(self.join_da_tick)} " + chu[cuoi:])

    def _link_dang_nham(self):
        """Link của các dòng đang bôi đen; không có dòng nào thì lấy dòng đã tick."""
        dang_boi_den = [self._link_cua_dong(d) for d in self.join_tree.selection()]
        return dang_boi_den or self.link_da_tick()

    def sao_chep_link(self):
        """Chép link của dòng đang bôi đen (hoặc các dòng đã tick) vào clipboard."""
        links = self._link_dang_nham()
        if not links:
            messagebox.showinfo("Chưa chọn dòng nào",
                                "Bấm vào một dòng trong bảng rồi sao chép.")
            return
        self.clipboard_clear()
        self.clipboard_append("\n".join(links))
        self.update()       # không gọi thì clipboard trống sau khi tool đóng
        self.append_join_log(
            f"📋 Đã sao chép {len(links)} link." if len(links) > 1
            else f"📋 Đã sao chép: {links[0]}")

    def mo_nhom_tren_web(self):
        """Mở nhóm đang chọn bằng trình duyệt mặc định để xem thử bằng tay."""
        links = self._link_dang_nham()
        if not links:
            return
        if len(links) > 5 and not messagebox.askyesno(
            "Mở nhiều tab", f"Sẽ mở {len(links)} tab trình duyệt. Tiếp tục?"):
            return
        for u in links[:20]:
            webbrowser.open(u)

    def _menu_chuot_phai(self, event):
        """Menu chuột phải trên bảng nhóm. Không đổi tick — chỉ chọn dòng."""
        dong = self.join_tree.identify_row(event.y)
        if not dong:
            return
        if dong not in self.join_tree.selection():
            self.join_tree.selection_set(dong)
        try:
            self.menu_nhom.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu_nhom.grab_release()

    def link_da_tick(self):
        """Danh sách link đang tick, theo đúng thứ tự hiện trong bảng."""
        return [self._link_cua_dong(d) for d in self.join_tree.get_children()
                if joiner.clean_group_url(self._link_cua_dong(d)) in self.join_da_tick]

    def _ghi_links(self, links):
        """Lưu danh sách nhóm mới (đã bỏ trùng) rồi vẽ lại bảng."""
        sach = []
        da_gap = set()
        for u in links:
            cu = joiner.clean_group_url(u)
            if cu and cu not in da_gap:
                da_gap.add(cu)
                sach.append(cu)
        cai_dat = joiner.load_cai_dat()
        cai_dat["links"] = sach
        # Đồng bộ luôn vào self.cfg đang giữ trong bộ nhớ: joiner ghi thẳng
        # xuống file, còn các nút 'Lưu cấu hình' khác lại ghi đè cả file bằng
        # self.cfg — không đồng bộ là danh sách nhóm vừa thêm biến mất.
        self.cfg["auto_join"] = joiner.save_cai_dat(cai_dat)
        self.refresh_join_tree()
        return sach

    def them_link_nhom(self):
        if self.dang_chay():
            return
        text = DanNhieuDong(
            self, "Dán link nhóm",
            "Mỗi dòng một link nhóm Facebook:").result
        if not text:
            return
        cu = joiner.load_cai_dat()["links"]
        moi = [d for d in text.splitlines() if d.strip()]
        sach = self._ghi_links(cu + moi)
        self.append_join_log(f"Đã thêm {len(moi)} dòng — danh sách còn {len(sach)} nhóm (đã bỏ trùng).")

    def nap_link_tu_chien_dich(self):
        """Mượn danh sách nhóm đã khai ở một chiến dịch cho khỏi gõ lại.

        Chỉ CHÉP sang một lần, không dùng chung: sau đó sửa bên nào là việc của
        bên đó, tab này không ăn theo chiến dịch nào cả.
        """
        if self.dang_chay():
            return
        ten = ChonMotMuc(self, "Nạp link từ chiến dịch",
                         "Chép link nhóm của chiến dịch nào sang đây?",
                         list(self.cfg["campaigns"])).result
        if not ten:
            return
        nguon = [u for urls in self.cfg["campaigns"][ten]["groups"].values()
                 for u in urls if u.strip()]
        if not nguon:
            messagebox.showinfo("Không có gì để nạp", f"Chiến dịch '{ten}' chưa khai link nhóm nào.")
            return
        truoc = len(joiner.load_cai_dat()["links"])
        sach = self._ghi_links(joiner.load_cai_dat()["links"] + nguon)
        self.append_join_log(
            f"Đã chép {len(nguon)} link từ chiến dịch '{ten}' — thêm được "
            f"{len(sach) - truoc} nhóm mới, tổng {len(sach)} nhóm.")

    def xoa_link_da_chon(self):
        if self.dang_chay():
            return
        bo = set(self.link_da_tick())
        if not bo:
            messagebox.showwarning("Chưa tick nhóm nào",
                                   "Hãy tick vào ô ✓ ở đầu các dòng cần xóa.")
            return
        if not messagebox.askyesno("Xóa nhóm đã tick",
                                   f"Bỏ {len(bo)} nhóm khỏi danh sách?\n\n"
                                   "Chỉ xóa khỏi danh sách này, không rời nhóm trên Facebook."):
            return
        con = [u for u in joiner.load_cai_dat()["links"]
               if joiner.clean_group_url(u) not in {joiner.clean_group_url(x) for x in bo}]
        self._ghi_links(con)
        self.append_join_log(f"Đã bỏ {len(bo)} nhóm khỏi danh sách.")

    def xoa_het_link(self):
        if self.dang_chay():
            return
        if not joiner.load_cai_dat()["links"]:
            return
        if not messagebox.askyesno("Xóa hết", "Xóa toàn bộ danh sách nhóm ở tab này?\n\n"
                                              "Lịch sử đã tham gia của các nick vẫn giữ nguyên."):
            return
        self._ghi_links([])
        self.append_join_log("Đã xóa toàn bộ danh sách nhóm.")

    def luu_cai_dat_join(self):
        if self.dang_chay():
            return
        cai_dat = joiner.load_cai_dat()
        cai_dat["nick_moi"] = self.nick_moi_var.get()
        for key, var in self.join_vars.items():
            try:
                cai_dat[key] = int(var.get().strip())
            except ValueError:
                messagebox.showerror("Sai định dạng", f"'{var.get()}' không phải số nguyên.")
                return

        # save_cai_dat tự kẹp về khoảng an toàn; đọc lại rồi đổ lên ô để người
        # dùng thấy ngay con số thật sự được dùng, không tưởng mình đặt được 5s.
        da_luu = joiner.save_cai_dat(cai_dat)
        self.cfg["auto_join"] = da_luu
        khac = {k: (cai_dat[k], da_luu[k]) for k in self.join_vars if cai_dat[k] != da_luu[k]}
        for key, var in self.join_vars.items():
            var.set(str(da_luu[key]))
        self.nick_moi_var.set(bool(da_luu["nick_moi"]))
        self.refresh_join_tree()

        if khac:
            chi_tiet = "\n".join(f"  • {k}: {a} → {b}" for k, (a, b) in khac.items())
            messagebox.showwarning(
                "Đã chỉnh về mức an toàn",
                "Vài giá trị nằm ngoài khoảng an toàn nên tool tự sửa:\n\n" + chi_tiet +
                "\n\nVào nhóm dồn dập là hành vi Facebook chặn nhanh nhất.")
        else:
            messagebox.showinfo("Đã lưu", "Đã lưu cài đặt tham gia nhóm.")

    # ---------- chạy ----------

    def bat_dau_join(self, chi_chon=False):
        """Chạy tham gia nhóm. `chi_chon` = chỉ các dòng đang chọn trong bảng."""
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("Đang chạy", "Đang có tiến trình chạy, hãy dừng trước.")
            return

        acc = self.join_nick_dang_chon()
        if not acc:
            messagebox.showwarning("Chưa chọn nick", "Hãy chọn nick sẽ đi tham gia nhóm.")
            return
        nick, path = acc

        cai_dat = joiner.load_cai_dat()
        if chi_chon:
            links = self.link_da_tick()
            if not links:
                messagebox.showwarning(
                    "Chưa tick nhóm nào",
                    "Hãy tick vào ô ✓ ở đầu mỗi dòng nhóm muốn chạy.\n\n"
                    "Hoặc bấm '▶ Tham gia tất cả nhóm chưa vào' để chạy cả danh sách.")
                return
        else:
            links = cai_dat["links"]

        if not links:
            messagebox.showinfo("Chưa có nhóm", "Danh sách nhóm đang trống. Bấm '+ Dán link nhóm'.")
            return

        lich_su = joiner.load_log(path)

        # Nick đang bị khoá vì có dấu hiệu bị chặn — chặn ngay tại giao diện,
        # đừng để mở Chrome ra rồi mới báo.
        dang_nghi, den, ly_do = joiner.dang_tam_nghi(lich_su)
        if dang_nghi:
            messagebox.showerror(
                "Nick đang tạm nghỉ",
                f"Lần chạy trước tool thấy dấu hiệu Facebook chặn nick '{nick}' "
                f"({ly_do}), nên đã khoá tới {den.strftime('%H:%M ngày %d/%m')}.\n\n"
                "Chạy tiếp lúc này là cách nhanh nhất để mất nick. Hãy chờ hết hạn, "
                "hoặc dùng nick khác.\n\nThật sự cần thì bấm 'Bỏ tạm nghỉ của nick' "
                "— bạn tự chịu rủi ro.")
            return

        if not joiner.trong_gio_hoat_dong(cai_dat):
            messagebox.showwarning(
                "Ngoài khung giờ",
                f"Tool chỉ tham gia nhóm trong khung {cai_dat['gio_bat_dau']}h–"
                f"{cai_dat['gio_ket_thuc']}h giờ Việt Nam.\n\n"
                "Xin vào nhóm lúc nửa đêm là dấu vết máy móc rất rõ — người thật "
                "thì đang ngủ. Đổi khung giờ ở ô cài đặt nếu bạn thật sự cần.")
            return

        can_vao = joiner.loc_nhom_can_vao(links, lich_su)
        da_hom_nay = joiner.dem_da_xin_hom_nay(lich_su)
        quota = joiner.tinh_quota(cai_dat, lich_su)

        if not can_vao:
            messagebox.showinfo(
                "Không có gì để làm",
                f"Mọi nhóm đang chọn đều đã xong với nick '{nick}', hoặc đang "
                "chờ tới hạn thử lại.")
            return
        if quota <= 0:
            messagebox.showwarning(
                "Chạm trần trong ngày",
                f"Nick '{nick}' hôm nay đã bấm xin vào {da_hom_nay} nhóm, chạm trần "
                f"{cai_dat['max_moi_ngay']} nhóm/ngày.\n\nMai chạy tiếp. Nâng trần "
                "lên được, nhưng số càng cao thì nhịp thao tác càng dày — nick ngày "
                "nào cũng vào hàng chục nhóm là thứ Facebook nhận ra rất nhanh.")
            return

        uoc_phut = round(quota * (cai_dat["nghi_min"] + cai_dat["nghi_max"]) / 2 / 60
                         + quota // max(cai_dat["nghi_dai_sau"], 1) * cai_dat["nghi_dai_phut"])
        if not messagebox.askyesno(
            "Bắt đầu tham gia nhóm",
            f"Nick    : {nick}\n"
            f"Đang chọn: {len(links)} nhóm, trong đó {len(can_vao)} nhóm chưa xử lý\n"
            f"Lần này  : tối đa {quota} nhóm (hôm nay đã {da_hom_nay}/{cai_dat['max_moi_ngay']})\n"
            f"Ước tính : khoảng {uoc_phut} phút\n\n"
            "Chrome sẽ mở và lần lượt vào từng nhóm, có nghỉ dài giữa chừng nên "
            "đừng sốt ruột. Đừng dùng cửa sổ Chrome đó trong lúc chạy.\n\n"
            "Chạy luôn?"):
            return

        self.stop_event.clear()
        joiner.LOG_FN = self.join_log_queue.put
        joiner.SHOULD_STOP = self.stop_event.is_set

        self.set_ui_locked(True)
        self.btn_join_stop.config(state="normal")
        self.status.config(text=f"⏵ Đang tham gia nhóm bằng nick: {nick}")

        def work():
            stats = None
            try:
                stats = joiner.chay_join(path, links=links, cai_dat=cai_dat)
            except Exception as e:
                self.join_log_queue.put(f"✗ LỖI: {e}")
            finally:
                self.join_log_queue.put(("__join_done__", stats))

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def bo_tam_nghi_nick(self):
        """Gỡ khoá tạm nghỉ cho nick đang chọn — chỉ dùng khi chắc chắn là báo nhầm."""
        if self.dang_chay():
            return
        acc = self.join_nick_dang_chon()
        if not acc:
            messagebox.showwarning("Chưa chọn nick", "Hãy chọn nick trong danh sách.")
            return
        nick, path = acc
        lich_su = joiner.load_log(path)
        dang_nghi, den, ly_do = joiner.dang_tam_nghi(lich_su)
        if not dang_nghi:
            messagebox.showinfo("Không có gì để bỏ", f"Nick '{nick}' đang không bị tạm nghỉ.")
            return
        if not messagebox.askyesno(
            "Bỏ tạm nghỉ",
            f"Tool khoá nick '{nick}' tới {den.strftime('%H:%M ngày %d/%m')} vì "
            f"{ly_do}.\n\nĐây là cảnh báo thật, không phải lỗi vặt: Facebook vừa "
            "chặn thao tác vào nhóm của nick này. Bỏ khoá rồi chạy tiếp thì khả "
            "năng mất nick là rất cao.\n\nVẫn bỏ khoá?"):
            return
        joiner.bo_tam_nghi(path, lich_su)
        self.refresh_join_tree()
        self.append_join_log(f"⚠ Đã bỏ tạm nghỉ cho nick '{nick}' theo yêu cầu.")

    def soat_lai_trang_thai(self):
        """Bỏ các bản ghi 'đã vào nhóm' mà tool chưa từng bấm nút Tham gia.

        Dùng khi bạn kiểm tra trên Facebook thấy thật ra chưa vào nhóm đó: xoá
        đi thì nhóm quay về 'chưa xử lý' và lần chạy tới tool mở lại kiểm tra.
        """
        if self.dang_chay():
            return
        acc = self.join_nick_dang_chon()
        if not acc:
            messagebox.showwarning("Chưa chọn nick", "Hãy chọn nick trong danh sách.")
            return
        nick, path = acc
        lich_su = joiner.load_log(path)
        so = sum(1 for m in lich_su.get("nhom", {}).values()
                 if isinstance(m, dict)
                 and m.get("status") in (joiner.DA_LA_THANH_VIEN, joiner.CHO_DUYET)
                 and not m.get("da_bam"))
        if not so:
            messagebox.showinfo(
                "Không có gì để soát",
                f"Mọi nhóm trong lịch sử của '{nick}' đều do tool tự bấm tham gia, "
                "không có bản ghi nào là tool tự kết luận.")
            return
        if not messagebox.askyesno(
            "Soát lại trạng thái",
            f"Có {so} nhóm được ghi là 'đã là thành viên / đang chờ duyệt' mà tool "
            f"KHÔNG hề bấm nút Tham gia — nó tự nhìn trang rồi kết luận.\n\n"
            "Xoá mấy bản ghi đó đi thì các nhóm này quay về 'chưa xử lý', lần chạy "
            "tới tool sẽ mở lại và kiểm tra tử tế.\n\nLàm luôn?"):
            return
        da_bo = joiner.don_log_ghi_khong(path, lich_su, ep_buoc=True)
        self.refresh_join_tree()
        self.append_join_log(f"🔄 Đã bỏ {da_bo} bản ghi tự kết luận của nick '{nick}' "
                             "— các nhóm đó sẽ được kiểm tra lại.")

    def append_join_log(self, text):
        self.join_log_box.insert("end", text + "\n")
        self.join_log_box.see("end")

    def _drain_join_log_queue(self):
        """Đổ log của việc tham gia nhóm vào ô log riêng của tab đó."""
        ve_lai_bang = False
        while True:
            try:
                item = self.join_log_queue.get_nowait()
            except queue.Empty:
                break

            if isinstance(item, tuple):
                _kind, stats = item
                self.set_ui_locked(False)
                self.btn_join_stop.config(state="disabled")
                self.status.config(text="Đã dừng — giờ sửa được cấu hình.")
                self.refresh_join_tree()
                if stats and stats.get("con_lai"):
                    messagebox.showinfo(
                        "Còn nhóm chưa xử lý",
                        f"Xong lượt này: {stats['vao']} nhóm vừa gửi yêu cầu.\n\n"
                        f"Còn {stats['con_lai']} nhóm chưa đụng tới — chạy lại vào "
                        "hôm khác, đừng chạy dồn trong một ngày.")
                continue

            dong = str(item)
            self.append_join_log(dong)
            # Xong một nhóm thì vẽ lại bảng để thấy trạng thái đổi ngay, không
            # phải chờ hết cả lượt chạy mới biết nhóm nào đã vào được.
            if "── Nhóm" in dong:
                ve_lai_bang = True

        if ve_lai_bang:
            self.refresh_join_tree()

    # ==================== ĐÓNG ====================

    def on_close(self):
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno(
                "Đang đăng bài",
                "Đang có bài đăng dở.\n\n"
                "Chọn Yes: dừng an toàn — chờ bài hiện tại gửi xong và ghi vào "
                "file chống trùng, rồi mới thoát (có thể mất một lúc).\n"
                "Chọn No: quay lại cửa sổ, không thoát.\n\n"
                "Tắt ngang giữa chừng dễ để lại bài đã lên Facebook nhưng chưa "
                "kịp ghi log — lần chạy sau sẽ đăng lại bài đó.",
            ):
                return
            self.stop_event.set()
            self.status.config(text="⏸ Đang dừng an toàn, sẽ tự thoát khi xong...")
            self._cho_roi_thoat()
            return

        if self.co_thay_doi_chua_luu():
            if not messagebox.askyesno(
                "Chưa lưu",
                "Tab 'Nội dung bài đăng' có thay đổi chưa ghi vào file Excel.\n\n"
                "Thoát bây giờ sẽ mất phần sửa đó. Vẫn thoát?"):
                return
        self.destroy()

    def _cho_roi_thoat(self):
        """Chờ worker dừng hẳn rồi mới đóng cửa sổ."""
        if self.worker and self.worker.is_alive():
            self.after(300, self._cho_roi_thoat)
            return
        self.destroy()

    def co_thay_doi_chua_luu(self):
        """So bảng hiện tại với lần nạp/ghi Excel gần nhất."""
        return self.excel_loaded and self.excel_rows != self.excel_snapshot


class XoaLichSuDang(tk.Toplevel):
    """Chọn file chống đăng trùng của (các) chiến dịch nào để xóa.

    Xóa là thao tác một chiều và hậu quả nằm trên Facebook chứ không nằm trong
    tool: xóa xong, lần chạy tới đăng lại toàn bộ bài lên mọi nhóm. Nên cửa sổ
    này nói rõ từng chiến dịch có bao nhiêu lượt, và còn một lần hỏi nữa trước
    khi xóa thật.
    """

    def __init__(self, parent):
        super().__init__(parent)
        self.app = parent
        self.title("Xóa lịch sử đã đăng")
        self.geometry("680x420")
        self.transient(parent)
        self.grab_set()

        ttk.Label(
            self,
            text="Mỗi chiến dịch có một file lịch sử riêng. Chọn chiến dịch muốn "
                 "xóa lịch sử (giữ Ctrl để chọn nhiều):",
            wraplength=640, justify="left",
        ).pack(anchor="w", padx=14, pady=(14, 6))

        khung = ttk.Frame(self)
        khung.pack(fill="both", expand=True, padx=14)
        self.danh_sach = tk.Listbox(khung, selectmode="extended", font=("Menlo", 11))
        thanh = ttk.Scrollbar(khung, command=self.danh_sach.yview)
        self.danh_sach.configure(yscrollcommand=thanh.set)
        thanh.pack(side="right", fill="y")
        self.danh_sach.pack(fill="both", expand=True)

        ttk.Label(
            self,
            text="Xóa xong, lần chạy tới tool sẽ đăng LẠI TỪ ĐẦU toàn bộ bài của "
                 "chiến dịch đó lên mọi nhóm — kể cả bài đã lên Facebook rồi.",
            foreground="#a05000", wraplength=640, justify="left",
        ).pack(anchor="w", padx=14, pady=8)

        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=14, pady=(0, 14))
        ttk.Button(bar, text="🗑 Xóa lịch sử đã chọn", command=self.xoa).pack(side="left")
        ttk.Button(bar, text="Đóng", command=self.destroy).pack(side="right")

        self.nap_lai()
        parent.wait_window(self)

    def nap_lai(self):
        self.ten_chien_dich = list(self.app.cfg["campaigns"])
        dang_chon = self.app.cfg["active_campaign"]
        self.danh_sach.delete(0, "end")
        for i, ten in enumerate(self.ten_chien_dich):
            duong_dan = self.app.cfg["campaigns"][ten].get("posted_log") or ""
            n = self.app.dem_luot_da_dang(duong_dan)
            mo_ta = "không đọc được file" if n < 0 else f"{n} lượt đã đăng"
            self.danh_sach.insert(
                "end", f"{ten}  —  {mo_ta}  —  {os.path.basename(duong_dan) or '(chưa đặt)'}")
            if ten == dang_chon:
                self.danh_sach.selection_set(i)

    def xoa(self):
        chon = [self.ten_chien_dich[i] for i in self.danh_sach.curselection()]
        if not chon:
            messagebox.showwarning(
                "Chưa chọn", "Hãy chọn ít nhất một chiến dịch.", parent=self)
            return

        dong = []
        tong = 0
        for ten in chon:
            duong_dan = self.app.cfg["campaigns"][ten].get("posted_log") or ""
            n = self.app.dem_luot_da_dang(duong_dan)
            tong += max(n, 0)
            dong.append(f"  • {ten}: {max(n, 0)} lượt")

        if not tong:
            messagebox.showinfo(
                "Chưa có gì để xóa",
                "Các chiến dịch đã chọn chưa ghi nhận lượt đăng nào.", parent=self)
            return

        if not messagebox.askyesno(
            "Xác nhận xóa",
            f"Xóa lịch sử đã đăng của {len(chon)} chiến dịch?\n\n"
            + "\n".join(dong)
            + f"\n\nTổng cộng {tong} lượt sẽ bị xóa.\n\n"
            "Lần chạy tới các bài này sẽ được đăng LẠI lên Facebook.\n"
            "Không hoàn tác được.", parent=self):
            return

        for ten in chon:
            duong_dan = self.app.cfg["campaigns"][ten].get("posted_log") or ""
            try:
                with open(duong_dan, "w", encoding="utf-8") as f:
                    json.dump({}, f)
            except Exception as e:
                messagebox.showerror(
                    "Không xóa được", f"Chiến dịch '{ten}':\n{e}", parent=self)
                continue
            self.app.append_log(f"🗑 Đã xóa lịch sử đã đăng của chiến dịch '{ten}'.")

        self.nap_lai()


class QuanLyAnh(tk.Toplevel):
    """Xem, xóa, thêm ảnh trong thư mục ảnh của một bài.

    Sinh ra để khỏi phải mở Explorer mỗi lần thay ảnh: chọn ảnh cũ xóa đi, thư
    mục vẫn còn nguyên để thả ảnh mới vào. Không bao giờ xóa chính thư mục —
    xóa mất là cột 'Thư mục ảnh' trong Excel trỏ vào chỗ trống.
    """

    DUOI_ANH = (".jpg", ".jpeg", ".png", ".webp")

    def __init__(self, parent, thu_muc, nhan):
        super().__init__(parent)
        self.thu_muc = thu_muc
        self.title(f"Ảnh: {nhan}")
        self.geometry("640x520")
        self.transient(parent)
        self.grab_set()

        ttk.Label(self, text=thu_muc, foreground="#555").pack(
            anchor="w", padx=14, pady=(12, 2))

        self.tom_tat = ttk.Label(self, text="")
        self.tom_tat.pack(anchor="w", padx=14)
        self.canh_bao = ttk.Label(self, text="", foreground="#a05000",
                                  wraplength=600, justify="left")
        self.canh_bao.pack(anchor="w", padx=14, pady=(2, 6))

        khung = ttk.Frame(self)
        khung.pack(fill="both", expand=True, padx=14)
        self.danh_sach = tk.Listbox(khung, selectmode="extended", font=("Menlo", 11))
        thanh = ttk.Scrollbar(khung, command=self.danh_sach.yview)
        self.danh_sach.configure(yscrollcommand=thanh.set)
        thanh.pack(side="right", fill="y")
        self.danh_sach.pack(fill="both", expand=True)
        self.danh_sach.bind("<Double-1>", lambda _e: self.xem_anh())

        ttk.Label(
            self,
            text="Giữ Ctrl (hoặc Shift) để chọn nhiều ảnh. Bấm đúp để xem ảnh.",
            foreground="#777",
        ).pack(anchor="w", padx=14, pady=(4, 0))

        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=14, pady=12)
        ttk.Button(bar, text="+ Thêm ảnh...", command=self.them_anh).pack(side="left")
        ttk.Button(bar, text="Xóa ảnh đã chọn", command=self.xoa_da_chon).pack(side="left", padx=4)
        ttk.Button(bar, text="Xóa tất cả ảnh", command=self.xoa_tat_ca).pack(side="left")
        ttk.Button(bar, text="Mở thư mục", command=self.mo_thu_muc).pack(side="left", padx=12)
        ttk.Button(bar, text="Đóng", command=self.destroy).pack(side="right")

        self.nap_lai()
        parent.wait_window(self)

    # ---------- đọc thư mục ----------

    def _quet(self):
        """Trả về (ảnh dùng được, file không dùng được) trong thư mục."""
        try:
            ten = sorted(os.listdir(self.thu_muc))
        except OSError:
            return [], []
        anh, khac = [], []
        for t in ten:
            if not os.path.isfile(os.path.join(self.thu_muc, t)):
                continue
            (anh if t.lower().endswith(self.DUOI_ANH) else khac).append(t)
        return anh, khac

    @staticmethod
    def _co(duong_dan):
        try:
            n = os.path.getsize(duong_dan)
        except OSError:
            return "?"
        return f"{n / 1024:.0f} KB" if n < 1024 * 1024 else f"{n / 1024 / 1024:.1f} MB"

    def nap_lai(self):
        anh, khac = self._quet()
        self.anh = anh
        self.danh_sach.delete(0, "end")
        tong = 0
        for ten in anh:
            dd = os.path.join(self.thu_muc, ten)
            try:
                tong += os.path.getsize(dd)
            except OSError:
                pass
            self.danh_sach.insert("end", f"{ten}    ({self._co(dd)})")

        if anh:
            self.tom_tat.config(
                text=f"{len(anh)} ảnh — tổng {tong / 1024 / 1024:.1f} MB")
        else:
            self.tom_tat.config(text="Thư mục trống — chưa có ảnh nào.")

        # File .heic của iPhone hay lọt vào qua Zalo mà Facebook không nhận;
        # không báo thì người dùng tưởng đã có ảnh, tới lúc đăng mới thấy trống.
        if khac:
            self.canh_bao.config(
                text=f"⚠ {len(khac)} file KHÔNG dùng để đăng được "
                     f"({', '.join(khac[:4])}{'...' if len(khac) > 4 else ''}) — "
                     "tool chỉ nhận .jpg .jpeg .png .webp")
        else:
            self.canh_bao.config(text="")

    def _dang_chon(self):
        return [self.anh[i] for i in self.danh_sach.curselection()]

    # ---------- thao tác ----------

    def them_anh(self):
        chon = filedialog.askopenfilenames(
            parent=self, title="Chọn ảnh để thêm vào thư mục này",
            filetypes=[("Ảnh", "*.jpg *.jpeg *.png *.webp"), ("Tất cả", "*.*")])
        if not chon:
            return
        them = bo_qua = 0
        for nguon in chon:
            ten = os.path.basename(nguon)
            dich = os.path.join(self.thu_muc, ten)
            # Trùng tên thì thêm hậu tố chứ không đè — hai ảnh khác nhau từ Zalo
            # rất hay cùng tên kiểu "photo_2026.jpg"
            goc, duoi = os.path.splitext(ten)
            n = 1
            while os.path.exists(dich):
                dich = os.path.join(self.thu_muc, f"{goc}_{n}{duoi}")
                n += 1
            try:
                shutil.copy2(nguon, dich)
                them += 1
            except Exception as e:
                bo_qua += 1
                messagebox.showerror("Không chép được", f"{ten}:\n{e}", parent=self)
        self.nap_lai()
        if them:
            self.master.append_log(f"🖼 Đã thêm {them} ảnh vào {self.thu_muc}")

    def xoa_da_chon(self):
        chon = self._dang_chon()
        if not chon:
            messagebox.showwarning(
                "Chưa chọn", "Hãy chọn ảnh muốn xóa trong danh sách.", parent=self)
            return
        if not messagebox.askyesno(
            "Xóa ảnh",
            f"Xóa hẳn {len(chon)} ảnh khỏi ổ đĩa?\n\n"
            "Không vào Thùng rác, không hoàn tác được.\n"
            "Thư mục vẫn giữ nguyên để bạn bỏ ảnh mới vào.", parent=self):
            return
        self._xoa(chon)

    def xoa_tat_ca(self):
        if not self.anh:
            messagebox.showinfo("Trống", "Thư mục chưa có ảnh nào.", parent=self)
            return
        if not messagebox.askyesno(
            "Xóa tất cả ảnh",
            f"Xóa hẳn toàn bộ {len(self.anh)} ảnh trong thư mục này?\n\n"
            "Không vào Thùng rác, không hoàn tác được.\n"
            "Thư mục vẫn giữ nguyên để bạn bỏ ảnh mới vào.", parent=self):
            return
        self._xoa(list(self.anh))

    def _xoa(self, ten_file):
        xoa = 0
        for ten in ten_file:
            try:
                os.remove(os.path.join(self.thu_muc, ten))
                xoa += 1
            except Exception as e:
                messagebox.showerror("Không xóa được", f"{ten}:\n{e}", parent=self)
        self.nap_lai()
        if xoa:
            self.master.append_log(f"🖼 Đã xóa {xoa} ảnh trong {self.thu_muc}")

    def xem_anh(self):
        chon = self._dang_chon()
        if chon:
            self._mo(os.path.join(self.thu_muc, chon[0]))

    def mo_thu_muc(self):
        self._mo(self.thu_muc)

    def _mo(self, duong_dan):
        """Mở file/thư mục bằng ứng dụng mặc định của hệ điều hành."""
        try:
            if sys.platform == "win32":
                os.startfile(duong_dan)
            else:
                subprocess.run(["open" if sys.platform == "darwin" else "xdg-open",
                                duong_dan], check=False)
        except Exception as e:
            messagebox.showerror("Không mở được", str(e), parent=self)


class _Popup(tk.Toplevel):
    """Phần chung của các popup chọn: bám dưới nút bấm, Esc để đóng."""

    def __init__(self, parent, tieu_de, neo):
        super().__init__(parent)
        self.title(tieu_de)
        self.result = None
        self.transient(parent)
        self.resizable(False, False)
        self.bind("<Escape>", lambda _e: self.destroy())

    def _bam_vao(self, neo, parent):
        """Đặt popup ngay dưới nút vừa bấm cho khỏi che ô đang nhập."""
        self.update_idletasks()
        if neo is not None and neo.winfo_ismapped():
            x, y = neo.winfo_rootx(), neo.winfo_rooty() + neo.winfo_height() + 2
        else:
            x = parent.winfo_rootx() + 60
            y = parent.winfo_rooty() + 60
        # Không cho tràn khỏi màn hình
        x = max(0, min(x, self.winfo_screenwidth() - self.winfo_width() - 8))
        y = max(0, min(y, self.winfo_screenheight() - self.winfo_height() - 8))
        self.geometry(f"+{x}+{y}")
        self.grab_set()
        parent.wait_window(self)


class ChonGio(_Popup):
    """Popup chọn giờ và phút. Chọn xong, self.result là chuỗi HH:MM."""

    def __init__(self, parent, gio_hien_tai="", neo=None):
        super().__init__(parent, "Chọn giờ", neo)

        gio, phut = self._tach(gio_hien_tai)

        khung = ttk.Frame(self)
        khung.pack(padx=12, pady=(10, 6))
        self.ds_gio = self._cot(khung, 0, "Giờ", 24, gio)
        ttk.Label(khung, text=":", font=("", 16)).grid(row=1, column=1, padx=6)
        self.ds_phut = self._cot(khung, 2, "Phút", 60, phut)

        bar = ttk.Frame(self)
        bar.pack(pady=(0, 10))
        ttk.Button(bar, text="Chọn", command=self.ok).pack(side="left", padx=4)
        ttk.Button(bar, text="Hủy", command=self.destroy).pack(side="left", padx=4)

        self._bam_vao(neo, parent)

    @staticmethod
    def _tach(gio_hien_tai):
        """"19:30" → (19, 30). Ô đang trống hoặc gõ sai thì lấy 19:30."""
        try:
            g, p = gio_hien_tai.strip().split(":")
            g, p = int(g), int(p)
            if 0 <= g <= 23 and 0 <= p <= 59:
                return g, p
        except (ValueError, AttributeError):
            pass
        return 19, 30

    def _cot(self, khung, cot, nhan, so_muc, dang_chon):
        ttk.Label(khung, text=nhan).grid(row=0, column=cot)
        hop = ttk.Frame(khung)
        hop.grid(row=1, column=cot)
        thanh = ttk.Scrollbar(hop, orient="vertical")
        ds = tk.Listbox(hop, height=8, width=4, exportselection=False,
                        font=("", 13), yscrollcommand=thanh.set)
        thanh.config(command=ds.yview)
        ds.pack(side="left")
        thanh.pack(side="left", fill="y")
        for i in range(so_muc):
            ds.insert("end", f"{i:02d}")
        ds.selection_set(dang_chon)
        # Kéo mục đang chọn ra giữa khung cho dễ nhìn
        ds.see(min(so_muc - 1, dang_chon + 4))
        ds.see(max(0, dang_chon - 3))
        ds.bind("<Double-Button-1>", lambda _e: self.ok())
        return ds

    @staticmethod
    def _dang_chon(ds):
        chon = ds.curselection()
        return int(ds.get(chon[0])) if chon else 0

    def ok(self):
        self.result = "{:02d}:{:02d}".format(
            self._dang_chon(self.ds_gio), self._dang_chon(self.ds_phut))
        self.destroy()


class ChonNgay(_Popup):
    """Popup lịch tháng. Chọn xong, self.result là chuỗi dd/mm/yyyy."""

    TEN_THU_NGAN = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"]

    def __init__(self, parent, ngay_hien_tai="", neo=None):
        super().__init__(parent, "Chọn ngày", neo)

        self.hom_nay = lich_hen.bay_gio().date()
        self.dang_chon = self._doc(ngay_hien_tai) or self.hom_nay
        self.nam, self.thang = self.dang_chon.year, self.dang_chon.month

        dieu_huong = ttk.Frame(self)
        dieu_huong.pack(fill="x", padx=10, pady=(10, 4))
        ttk.Button(dieu_huong, text="◀", width=3,
                   command=lambda: self._doi_thang(-1)).pack(side="left")
        self.nhan_thang = ttk.Label(dieu_huong, anchor="center", font=("", 13, "bold"))
        self.nhan_thang.pack(side="left", expand=True, fill="x")
        ttk.Button(dieu_huong, text="▶", width=3,
                   command=lambda: self._doi_thang(1)).pack(side="left")

        self.luoi = ttk.Frame(self)
        self.luoi.pack(padx=10)
        for i, ten in enumerate(self.TEN_THU_NGAN):
            ttk.Label(self.luoi, text=ten, width=4, anchor="center",
                      foreground="#c0392b" if i == 6 else "#333").grid(row=0, column=i, pady=2)

        bar = ttk.Frame(self)
        bar.pack(pady=8)
        ttk.Button(bar, text="Hôm nay", command=self._ve_hom_nay).pack(side="left", padx=4)
        ttk.Button(bar, text="Hủy", command=self.destroy).pack(side="left", padx=4)

        self._ve_thang()
        self._bam_vao(neo, parent)

    @staticmethod
    def _doc(chuoi):
        try:
            return datetime.strptime(chuoi.strip(), "%d/%m/%Y").date()
        except (ValueError, AttributeError):
            return None

    def _doi_thang(self, buoc):
        thang = self.thang + buoc
        self.nam += (thang - 1) // 12
        self.thang = (thang - 1) % 12 + 1
        self._ve_thang()

    def _ve_hom_nay(self):
        self._chon(self.hom_nay)

    def _ve_thang(self):
        """Vẽ lại lưới ngày của tháng đang xem."""
        self.nhan_thang.config(text=f"Tháng {self.thang}/{self.nam}")
        for o in self.luoi.grid_slaves():
            if int(o.grid_info()["row"]) > 0:
                o.destroy()

        for dong, tuan in enumerate(calendar.Calendar().monthdayscalendar(self.nam, self.thang), 1):
            for cot, ngay in enumerate(tuan):
                if ngay == 0:
                    continue
                d = date(self.nam, self.thang, ngay)
                nut = tk.Button(self.luoi, text=str(ngay), width=3, relief="flat",
                                command=lambda d=d: self._chon(d))
                if d == self.dang_chon:
                    nut.config(bg="#2d7ff9", fg="white", relief="raised")
                elif d == self.hom_nay:
                    nut.config(fg="#2d7ff9", font=("", 12, "bold"))
                elif d < self.hom_nay:
                    # Ngày đã qua thì làm mờ — chọn vào cũng không chạy được
                    nut.config(fg="#aaa")
                nut.grid(row=dong, column=cot, padx=1, pady=1)

    def _chon(self, d):
        self.result = d.strftime("%d/%m/%Y")
        self.destroy()


class LichEditor(tk.Toplevel):
    """Hộp thoại tạo/sửa một lịch hẹn giờ đăng."""

    def __init__(self, parent, title, lich, campaigns=(), nicks=()):
        super().__init__(parent)
        self.title(title)
        self.geometry("560x460")
        self.result = None
        self.lich = lich_hen.chuan_hoa(dict(lich))
        self.transient(parent)
        self.grab_set()

        bg = lich_hen.bay_gio()

        khung = ttk.Frame(self)
        khung.pack(fill="both", expand=True, padx=16, pady=14)

        ttk.Label(khung, text="Kiểu hẹn:").grid(row=0, column=0, sticky="w", pady=6)
        self.kieu_var = tk.StringVar(value=self.lich["kieu"])
        hop_kieu = ttk.Frame(khung)
        hop_kieu.grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(hop_kieu, text="Một lần", value="mot_lan",
                        variable=self.kieu_var, command=self._doi_kieu).pack(side="left")
        ttk.Radiobutton(hop_kieu, text="Lặp lại theo thứ", value="hang_ngay",
                        variable=self.kieu_var, command=self._doi_kieu).pack(side="left", padx=12)

        ttk.Label(khung, text="Giờ (giờ VN):").grid(row=1, column=0, sticky="w", pady=6)
        self.gio_var = tk.StringVar(value=self.lich["gio"] or "19:30")
        hop_gio = ttk.Frame(khung)
        hop_gio.grid(row=1, column=1, sticky="w")
        ttk.Entry(hop_gio, textvariable=self.gio_var, width=10).pack(side="left")
        self.nut_gio = ttk.Button(hop_gio, text="🕒", width=3, command=self._mo_chon_gio)
        self.nut_gio.pack(side="left", padx=4)
        ttk.Label(khung, text="bấm 🕒 để chọn, hoặc gõ dạng HH:MM",
                  foreground="#777").grid(row=1, column=2, sticky="w", padx=8)

        ttk.Label(khung, text="Ngày:").grid(row=2, column=0, sticky="w", pady=6)
        self.ngay_var = tk.StringVar(
            value=self._ngay_hien_thi(self.lich["ngay"]) or bg.strftime("%d/%m/%Y"))
        hop_ngay = ttk.Frame(khung)
        hop_ngay.grid(row=2, column=1, sticky="w")
        self.o_ngay = ttk.Entry(hop_ngay, textvariable=self.ngay_var, width=10)
        self.o_ngay.pack(side="left")
        self.nut_ngay = ttk.Button(hop_ngay, text="📅", width=3, command=self._mo_chon_ngay)
        self.nut_ngay.pack(side="left", padx=4)
        self.ghi_chu_ngay = ttk.Label(khung, text="bấm 📅 để chọn, hoặc gõ ngày/tháng/năm",
                                      foreground="#777")
        self.ghi_chu_ngay.grid(row=2, column=2, sticky="w", padx=8)

        self.khung_thu = ttk.LabelFrame(khung, text="Lặp vào các thứ (không chọn gì = mọi ngày)")
        self.khung_thu.grid(row=3, column=0, columnspan=3, sticky="ew", pady=10)
        self.thu_vars = []
        for i, ten in enumerate(lich_hen.TEN_THU):
            v = tk.BooleanVar(value=i in (self.lich["thu"] or []))
            ttk.Checkbutton(self.khung_thu, text=ten, variable=v).grid(
                row=i // 4, column=i % 4, sticky="w", padx=8, pady=3)
            self.thu_vars.append(v)

        ttk.Label(khung, text="Chiến dịch:").grid(row=4, column=0, sticky="w", pady=6)
        self.camp_var = tk.StringVar(
            value=self.lich["campaign"] or (campaigns[0] if campaigns else ""))
        ttk.Combobox(khung, textvariable=self.camp_var, values=list(campaigns),
                     state="readonly", width=28).grid(row=4, column=1, columnspan=2, sticky="w")

        ttk.Label(khung, text="Nick Facebook:").grid(row=5, column=0, sticky="w", pady=6)
        self.nick_var = tk.StringVar(
            value=self.lich["nick"] or (nicks[0] if nicks else ""))
        ttk.Combobox(khung, textvariable=self.nick_var, values=list(nicks),
                     state="readonly", width=28).grid(row=5, column=1, columnspan=2, sticky="w")

        self.bat_var = tk.BooleanVar(value=self.lich["bat"])
        ttk.Checkbutton(khung, text="Bật lịch này", variable=self.bat_var).grid(
            row=6, column=1, sticky="w", pady=8)

        bar = ttk.Frame(self)
        bar.pack(pady=(0, 14))
        ttk.Button(bar, text="Lưu", command=self.ok).pack(side="left", padx=4)
        ttk.Button(bar, text="Hủy", command=self.destroy).pack(side="left", padx=4)

        self._doi_kieu()
        parent.wait_window(self)

    @staticmethod
    def _ngay_hien_thi(iso):
        """ISO trong config (YYYY-MM-DD) → dd/mm/yyyy cho người Việt đọc."""
        try:
            return datetime.strptime(iso, "%Y-%m-%d").strftime("%d/%m/%Y")
        except (ValueError, TypeError):
            return ""

    def _mo_chon_gio(self):
        pop = ChonGio(self, self.gio_var.get(), neo=self.nut_gio)
        if pop.result:
            self.gio_var.set(pop.result)

    def _mo_chon_ngay(self):
        pop = ChonNgay(self, self.ngay_var.get(), neo=self.nut_ngay)
        if pop.result:
            self.ngay_var.set(pop.result)

    def _doi_kieu(self):
        """Ẩn/hiện ô Ngày và bảng Thứ theo kiểu hẹn đang chọn."""
        mot_lan = self.kieu_var.get() == "mot_lan"
        for w in (self.o_ngay, self.nut_ngay, self.ghi_chu_ngay):
            w.configure(state="normal" if mot_lan else "disabled")
        for con in self.khung_thu.winfo_children():
            con.configure(state="disabled" if mot_lan else "normal")

    def ok(self):
        gio = self.gio_var.get().strip()
        if not lich_hen.gio_hop_le(gio):
            messagebox.showwarning(
                "Giờ không hợp lệ",
                "Giờ phải theo dạng HH:MM, vd 08:00 hoặc 19:30.", parent=self)
            return

        if not self.camp_var.get():
            messagebox.showwarning("Thiếu chiến dịch", "Hãy chọn chiến dịch.", parent=self)
            return
        if not self.nick_var.get():
            messagebox.showwarning(
                "Thiếu nick",
                "Hãy chọn nick Facebook.\nChưa có nick nào thì thêm ở tab "
                "'Tài khoản & Chạy' trước.", parent=self)
            return

        l = dict(self.lich)
        l.update({
            "kieu": self.kieu_var.get(),
            "gio": gio,
            "campaign": self.camp_var.get(),
            "nick": self.nick_var.get(),
            "bat": self.bat_var.get(),
            "thu": [i for i, v in enumerate(self.thu_vars) if v.get()],
        })

        if l["kieu"] == "mot_lan":
            try:
                ngay = datetime.strptime(self.ngay_var.get().strip(), "%d/%m/%Y")
            except ValueError:
                messagebox.showwarning(
                    "Ngày không hợp lệ",
                    "Ngày phải theo dạng ngày/tháng/năm, vd 25/07/2026.", parent=self)
                return
            l["ngay"] = ngay.strftime("%Y-%m-%d")
            if l["bat"] and lich_hen.lan_ke_tiep(l) is None:
                messagebox.showwarning(
                    "Mốc đã qua",
                    "Ngày giờ bạn chọn nằm trong quá khứ nên lịch sẽ không bao "
                    "giờ chạy. Hãy chọn mốc trong tương lai.", parent=self)
                return
        else:
            l["ngay"] = ""

        self.result = l
        self.destroy()


class RowEditor(tk.Toplevel):
    """Hộp thoại sửa 1 bài đăng: nội dung nhiều dòng, phân loại, thư mục ảnh."""

    def __init__(self, parent, title, row=None, categories=(), goi_y_stt=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("680x520")
        self.result = None
        self.transient(parent)
        self.grab_set()

        self.stt_goc = row["stt"] if row else None
        row = row or {"stt": goi_y_stt, "noi_dung": "", "phan_loai": "", "folder": ""}

        ttk.Label(self, text="Nội dung bài đăng (gõ xuống dòng thoải mái, "
                             "Facebook sẽ giữ nguyên):").pack(anchor="w", padx=14, pady=(14, 4))
        self.text = tk.Text(self, wrap="word", height=14, font=("Menlo", 12))
        self.text.pack(fill="both", expand=True, padx=14)
        self.text.insert("1.0", row["noi_dung"])

        grid = ttk.Frame(self)
        grid.pack(fill="x", padx=14, pady=10)

        ttk.Label(grid, text="STT:").grid(row=0, column=0, sticky="w", pady=4)
        self.stt_var = tk.StringVar(value="" if row["stt"] is None else str(row["stt"]))
        ttk.Entry(grid, textvariable=self.stt_var, width=8).grid(row=0, column=1, sticky="w", padx=8)
        ttk.Label(grid, text="(mã định danh bài — đừng đổi nếu bài đã từng đăng)",
                  foreground="#a05000").grid(row=0, column=2, sticky="w")

        ttk.Label(grid, text="Phân loại:").grid(row=1, column=0, sticky="w", pady=4)
        self.cat_var = tk.StringVar(value=row["phan_loai"])
        # Combobox cho chọn nhanh nhưng vẫn gõ tay được nếu muốn phân loại mới
        ttk.Combobox(grid, textvariable=self.cat_var, values=list(categories), width=28) \
            .grid(row=0, column=1, sticky="w", padx=8)
        ttk.Label(grid, text="(phải khớp tên nhóm đã khai)", foreground="#777") \
            .grid(row=0, column=2, sticky="w")

        ttk.Label(grid, text="Thư mục ảnh:").grid(row=1, column=0, sticky="w", pady=4)
        self.folder_var = tk.StringVar(value=row["folder"])
        ttk.Entry(grid, textvariable=self.folder_var, width=30).grid(row=1, column=1, sticky="w", padx=8)
        ttk.Label(grid, text="(tên thư mục con, để trống nếu không đăng ảnh)", foreground="#777") \
            .grid(row=1, column=2, sticky="w")

        bar = ttk.Frame(self)
        bar.pack(pady=(0, 14))
        ttk.Button(bar, text="Lưu", command=self.ok).pack(side="left", padx=4)
        ttk.Button(bar, text="Hủy", command=self.destroy).pack(side="left", padx=4)

        self.text.focus_set()
        parent.wait_window(self)

    def ok(self):
        noi_dung = self.text.get("1.0", "end").rstrip("\n")
        if not noi_dung.strip():
            messagebox.showwarning("Thiếu nội dung", "Nội dung bài không được để trống.", parent=self)
            return
        try:
            stt = int(self.stt_var.get().strip())
        except ValueError:
            messagebox.showwarning("STT không hợp lệ", "STT phải là số nguyên.", parent=self)
            return
        # Đổi STT của bài cũ = đổi mã định danh trong posted_log → dễ đăng lại
        if self.stt_goc is not None and stt != self.stt_goc and not messagebox.askyesno(
            "Đổi STT?",
            f"Bài này đang có STT {self.stt_goc}, bạn đổi thành {stt}.\n\n"
            "File chống đăng trùng ghi theo STT, nên đổi số có thể khiến bài "
            "được đăng lại lên các group đã đăng, hoặc bị bỏ qua nhầm.\n\nVẫn đổi?",
            parent=self,
        ):
            return
        self.result = {
            "stt": stt,
            "noi_dung": noi_dung,
            "phan_loai": self.cat_var.get().strip(),
            "folder": self.folder_var.get().strip(),
        }
        self.destroy()


class SimplePrompt(tk.Toplevel):
    """Hộp thoại nhập 1 dòng text, gõ được tiếng Việt."""

    def __init__(self, parent, title, label, initial=""):
        super().__init__(parent)
        self.title(title)
        self.result = None
        self.transient(parent)
        self.grab_set()

        ttk.Label(self, text=label).pack(padx=16, pady=(16, 6), anchor="w")
        var = tk.StringVar(value=initial)
        entry = ttk.Entry(self, textvariable=var, width=44)
        entry.pack(padx=16, fill="x")
        entry.focus_set()

        def ok():
            self.result = var.get().strip() or None
            self.destroy()

        bar = ttk.Frame(self)
        bar.pack(pady=12)
        ttk.Button(bar, text="OK", command=ok).pack(side="left", padx=4)
        ttk.Button(bar, text="Hủy", command=self.destroy).pack(side="left", padx=4)
        entry.bind("<Return>", lambda _e: ok())

        parent.wait_window(self)


class DanNhieuDong(tk.Toplevel):
    """Hộp thoại dán nhiều dòng (danh sách link nhóm). Trả về chuỗi hoặc None."""

    def __init__(self, parent, title, label):
        super().__init__(parent)
        self.title(title)
        self.result = None
        self.transient(parent)
        self.grab_set()

        ttk.Label(self, text=label).pack(padx=16, pady=(16, 6), anchor="w")
        box = tk.Text(self, width=70, height=14, wrap="none")
        box.pack(padx=16, fill="both", expand=True)
        box.focus_set()

        def ok():
            self.result = box.get("1.0", "end").strip() or None
            self.destroy()

        bar = ttk.Frame(self)
        bar.pack(pady=12)
        ttk.Button(bar, text="Thêm vào danh sách", command=ok).pack(side="left", padx=4)
        ttk.Button(bar, text="Hủy", command=self.destroy).pack(side="left", padx=4)

        parent.wait_window(self)


class ChonMotMuc(tk.Toplevel):
    """Hộp thoại chọn 1 mục trong danh sách. Trả về mục đã chọn hoặc None."""

    def __init__(self, parent, title, label, cac_muc):
        super().__init__(parent)
        self.title(title)
        self.result = None
        self.transient(parent)
        self.grab_set()

        ttk.Label(self, text=label).pack(padx=16, pady=(16, 6), anchor="w")
        ds = tk.Listbox(self, width=44, height=min(max(len(cac_muc), 3), 12),
                        exportselection=False)
        for m in cac_muc:
            ds.insert("end", m)
        if cac_muc:
            ds.selection_set(0)
        ds.pack(padx=16, fill="both", expand=True)
        ds.focus_set()

        def ok():
            sel = ds.curselection()
            self.result = ds.get(sel[0]) if sel else None
            self.destroy()

        bar = ttk.Frame(self)
        bar.pack(pady=12)
        ttk.Button(bar, text="OK", command=ok).pack(side="left", padx=4)
        ttk.Button(bar, text="Hủy", command=self.destroy).pack(side="left", padx=4)
        ds.bind("<Double-Button-1>", lambda _e: ok())

        parent.wait_window(self)


if __name__ == "__main__":
    App().mainloop()
