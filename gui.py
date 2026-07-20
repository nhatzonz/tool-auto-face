"""
Giao diện cho tool đăng bài Facebook hàng loạt.

Chạy:  source .venv/bin/activate && python gui.py

Làm việc theo "chiến dịch" — phòng trọ, tuyển dụng, seeding website... Mỗi
chiến dịch có file Excel, thư mục ảnh, file log và bản đồ nhóm riêng; nick
Facebook dùng chung cho mọi chiến dịch.

Ba tab:
  - Tài khoản & Chạy   : chọn/thêm nick Facebook, bấm chạy, xem log trực tiếp
  - Dữ liệu chiến dịch : file Excel, thư mục ảnh, thư mục profile, các delay
  - Nhóm theo phân loại: thêm/sửa/xóa phân loại và danh sách link group

Việc đăng bài chạy trong thread riêng để giao diện không bị treo.
"""
import os
import queue
import shutil
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import openpyxl

import config as cfg_module
import post_phong_tro_fb as bot
import tao_file_mau


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tool đăng bài Facebook")
        self.geometry("960x700")

        self.cfg = cfg_module.load_config()
        self.log_queue = queue.Queue()      # worker thread → giao diện
        self.stop_event = threading.Event()
        self.worker = None
        self.current_cat = None             # phân loại đang chọn ở tab Nhóm
        self.dang_khoa = False              # True khi đang đăng → cấm sửa
        self.excel_loaded = False
        self.excel_snapshot = []            # ảnh chụp lần nạp/ghi Excel gần nhất

        self._build_campaign_bar()

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        self.tab_run = ttk.Frame(notebook)
        self.tab_data = ttk.Frame(notebook)
        self.tab_paths = ttk.Frame(notebook)
        self.tab_groups = ttk.Frame(notebook)
        notebook.add(self.tab_run, text="  Tài khoản & Chạy  ")
        notebook.add(self.tab_data, text="  Nội dung bài đăng  ")
        notebook.add(self.tab_paths, text="  Dữ liệu chiến dịch  ")
        notebook.add(self.tab_groups, text="  Nhóm theo phân loại  ")

        self._build_run_tab()
        self._build_data_tab()
        self._build_paths_tab()
        self._build_groups_tab()

        self._tao_du_lieu_mau_lan_dau()
        self.load_campaign_into_views()
        self.refresh_accounts()
        self.after(100, self._drain_log_queue)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(50, self._hich_ve_lai)

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
                "Đang đăng bài. Bấm '■ Dừng' và chờ dừng hẳn rồi mới sửa được "
                "cấu hình — sửa giữa chừng sẽ làm bài đang chạy ghi log sai "
                "chiến dịch.")
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
        import json
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
        cfg_module.save_config(self.cfg)
        self.refresh_campaign_box()

    def delete_campaign(self):
        if self.dang_chay():
            return
        if len(self.cfg["campaigns"]) == 1:
            messagebox.showwarning("Không xóa được", "Phải còn ít nhất 1 chiến dịch.")
            return
        name = self.cfg["active_campaign"]
        if not messagebox.askyesno(
            "Xóa chiến dịch",
            f"Xóa chiến dịch '{name}' khỏi cấu hình?\n\n"
            "File Excel, ảnh và log trên ổ đĩa vẫn giữ nguyên, chỉ mất phần "
            "khai báo đường dẫn và danh sách nhóm.",
        ):
            return
        self.cfg["campaigns"].pop(name)
        self.cfg["active_campaign"] = next(iter(self.cfg["campaigns"]))
        cfg_module.save_config(self.cfg)
        self.refresh_campaign_box()
        self.load_campaign_into_views()

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
        ttk.Button(bar, text="Xóa log", command=lambda: self.log_box.delete("1.0", "end")).pack(side="left")

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
        if not messagebox.askyesno("Xóa nick", f"Xóa hẳn profile '{name}'?\nSẽ mất cookie đăng nhập của nick này."):
            return
        shutil.rmtree(path, ignore_errors=True)
        self.refresh_accounts()
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
        if not os.path.exists(camp["excel_path"]):
            messagebox.showerror(
                "Thiếu file",
                f"Chiến dịch '{camp_name}' chưa có file Excel hợp lệ:\n{camp['excel_path']}\n\n"
                "Vào tab 'Dữ liệu chiến dịch' để chọn file.",
            )
            return
        if not camp["groups"]:
            messagebox.showerror(
                "Chưa có nhóm",
                f"Chiến dịch '{camp_name}' chưa khai báo nhóm nào.\n"
                "Vào tab 'Nhóm theo phân loại' để thêm.",
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

        self.stop_event.clear()
        bot.reload_config(camp_name)
        bot.LOG_FN = self.log_queue.put
        bot.SHOULD_STOP = self.stop_event.is_set

        self.set_ui_locked(True)
        self.status.config(text=f"⏵ Đang chạy — chiến dịch '{camp_name}', nick: {nick}")

        def work():
            try:
                bot.run_posting(path, campaign=camp_name)
            except Exception as e:
                self.log_queue.put(f"✗ LỖI: {e}")
            finally:
                self.log_queue.put(("__run_done__", None))

        self.worker = threading.Thread(target=work, daemon=True)
        self.worker.start()

    def request_stop(self):
        self.stop_event.set()
        self.btn_stop.config(state="disabled")
        self.status.config(text="⏸ Đang dừng... chờ bài hiện tại đăng xong rồi mới nhả khoá.")

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
                if con in (self.btn_stop,):
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
        # Ô log luôn xem được
        self.log_box.config(state="normal")

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
                elif kind == "__login_done__":
                    payload.destroy()
                    self.refresh_accounts()
                continue

            self.append_log(str(item))

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


if __name__ == "__main__":
    App().mainloop()
