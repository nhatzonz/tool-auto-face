"""Bảng màu và style ttk dùng chung cho toàn bộ giao diện.

Gom một chỗ để mọi cửa sổ (cửa sổ chính lẫn các hộp thoại) trông như nhau, và
đổi màu sau này chỉ sửa ở đây chứ không đi dò từng widget.

Vì sao ép theme 'clam' thay vì để theme mặc định của từng hệ điều hành: theme
'aqua' trên macOS và 'vista' trên Windows không cho đổi màu nền nút, nên nút
chính không thể nổi lên được, mà hai máy lại hiện ra hai kiểu khác hẳn nhau —
sửa giao diện trên Mac xong sang Windows thành ra chưa từng sửa gì. 'clam' có
trên mọi bản Tk và ăn đủ các tuỳ chọn màu.
"""
import platform
import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk

NEN = "#f7f8fa"          # nền cửa sổ
THE = "#ffffff"          # nền ô nhập, bảng, danh sách
CHU = "#1f2328"          # chữ chính
CHU_PHU = "#6b7280"      # chú thích, tiêu đề cột
CHU_MO = "#9aa1ab"       # chữ của nút đang bị khoá
VIEN = "#d8dce2"          # viền nhạt: khung nhóm, đường kẻ
VIEN_DAM = "#b9c0ca"
# Viền nút phải đậm hơn viền khung: nút là thứ bấm được, nhìn phải ra ngay ranh
# giới của nó. Để cùng tông nhạt với khung thì trên nền trắng gần như mất hút.
VIEN_NUT = "#a3acb9"
VIEN_NUT_RE = "#7d8794"   # khi rê chuột
NHAN = "#2563eb"         # màu nhấn: nút chính, viền ô đang gõ
NHAN_DAM = "#1d4ed8"
NHAN_NHAT = "#e8efff"    # nền dòng đang chọn
CANH_BAO = "#b45309"
NGUY_HIEM = "#dc2626"
THANH_CONG = "#15803d"
SOC = "#f4f6f9"          # dòng kẻ xen kẽ trong bảng


def _font_he_dieu_hanh():
    """Trả về (font chữ thường, font đều nét) hợp với hệ điều hành đang chạy.

    Code cũ ghi thẳng ("Menlo", 11) — Menlo chỉ có trên Mac, sang Windows Tk âm
    thầm rơi về font mặc định xấu hơn. Chọn theo hệ máy để chỗ nào cũng đúng.
    """
    he = platform.system()
    if he == "Windows":
        return ("Segoe UI", 10), ("Consolas", 10)
    if he == "Darwin":
        return ("SF Pro Text", 13), ("Menlo", 12)
    return ("DejaVu Sans", 10), ("DejaVu Sans Mono", 10)


def ap_dung(root):
    """Áp bảng màu + font cho root và cho mọi widget tạo ra sau lời gọi này."""
    chu, don_cach = _font_he_dieu_hanh()

    for ten in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
        tkfont.nametofont(ten).configure(family=chu[0], size=chu[1])
    tkfont.nametofont("TkFixedFont").configure(family=don_cach[0], size=don_cach[1])

    _option_widget_co_dien(root)
    _style_ttk(root, chu)
    root.configure(background=NEN)


def _option_widget_co_dien(root):
    """Màu cho tk.Text và tk.Listbox — hai widget không theo style của ttk.

    Dùng option database thay vì sửa từng chỗ tạo widget: nó áp cho cả widget
    nằm trong các hộp thoại mở ra sau này, nên thêm hộp thoại mới cũng không
    phải nhớ tô màu lại.
    """
    o = root.option_add
    for lop in ("Text", "Listbox"):
        o(f"*{lop}.background", THE)
        o(f"*{lop}.foreground", CHU)
        o(f"*{lop}.borderWidth", 0)
        o(f"*{lop}.highlightThickness", 1)
        o(f"*{lop}.highlightBackground", VIEN)
        o(f"*{lop}.highlightColor", NHAN)
        o(f"*{lop}.relief", "flat")

    o("*Text.insertBackground", CHU)
    o("*Text.selectBackground", NHAN_NHAT)
    o("*Text.selectForeground", CHU)
    o("*Text.padX", 8)
    o("*Text.padY", 6)
    o("*Text.spacing1", 1)       # thở giữa các dòng cho dễ đọc log dài
    o("*Text.spacing3", 1)

    o("*Listbox.selectBackground", NHAN)
    o("*Listbox.selectForeground", "#ffffff")
    o("*Listbox.activeStyle", "none")

    # Danh sách xổ xuống của Combobox là một Listbox riêng, không dính style ttk
    o("*TCombobox*Listbox.background", THE)
    o("*TCombobox*Listbox.foreground", CHU)
    o("*TCombobox*Listbox.selectBackground", NHAN)
    o("*TCombobox*Listbox.selectForeground", "#ffffff")

    o("*Toplevel.background", NEN)


def _style_ttk(root, chu):
    st = ttk.Style(root)
    st.theme_use("clam")

    st.configure(".", background=NEN, foreground=CHU, fieldbackground=THE,
                 bordercolor=VIEN, focuscolor=NHAN, font=chu)
    st.configure("TFrame", background=NEN)
    st.configure("TLabel", background=NEN, foreground=CHU)
    st.configure("Phu.TLabel", foreground=CHU_PHU)
    st.configure("CanhBao.TLabel", foreground=CANH_BAO)
    st.configure("NguyHiem.TLabel", foreground=NGUY_HIEM)
    st.configure("TieuDe.TLabel", font=(chu[0], chu[1] + 1, "bold"))

    st.configure("TLabelframe", background=NEN, bordercolor=VIEN,
                 relief="solid", borderwidth=1, padding=6)
    st.configure("TLabelframe.Label", background=NEN, foreground=CHU_PHU,
                 font=(chu[0], chu[1], "bold"))

    # Nút thường: nền trắng, viền liền rõ, đậm thêm khi rê chuột.
    # relief phải là "solid" — để "flat" thì clam không vẽ đường viền nào cả,
    # nút chỉ còn là chữ trôi trên nền trắng, không biết đâu là vùng bấm được.
    st.configure("TButton", background=THE, foreground=CHU, bordercolor=VIEN_NUT,
                 borderwidth=1, relief="solid", padding=(12, 6), anchor="center",
                 focuscolor=NEN)
    st.map("TButton",
           background=[("pressed", "#e7ebf1"), ("active", "#f1f4f8"), ("disabled", NEN)],
           foreground=[("disabled", CHU_MO)],
           bordercolor=[("pressed", VIEN_NUT_RE), ("active", VIEN_NUT_RE),
                        ("disabled", VIEN)])

    # Nút chính của mỗi màn hình: chỉ nên có một, để mắt biết đi đâu trước
    st.configure("Nhan.TButton", background=NHAN, foreground="#ffffff", bordercolor=NHAN)
    st.map("Nhan.TButton",
           background=[("pressed", "#1a46c7"), ("active", NHAN_DAM), ("disabled", "#bfcdf0")],
           bordercolor=[("pressed", "#1a46c7"), ("active", NHAN_DAM), ("disabled", "#bfcdf0")],
           foreground=[("disabled", "#eef2ff")])

    # Nút xóa/dừng: chữ đỏ chứ không nền đỏ — nền đỏ nhiều quá thành ra nhức mắt
    st.configure("NguyHiem.TButton", foreground=NGUY_HIEM, bordercolor="#e0a7a7")
    st.map("NguyHiem.TButton",
           background=[("pressed", "#fbdcdc"), ("active", "#fdeeee"), ("disabled", NEN)],
           bordercolor=[("pressed", "#c96b6b"), ("active", "#c96b6b"), ("disabled", VIEN)],
           foreground=[("disabled", CHU_MO)])

    st.configure("TNotebook", background=NEN, bordercolor=VIEN, tabmargins=(2, 6, 2, 0))
    st.configure("TNotebook.Tab", background="#e9edf2", foreground=CHU_PHU,
                 bordercolor=VIEN, padding=(14, 8))
    st.map("TNotebook.Tab",
           background=[("selected", THE), ("active", "#f1f4f8")],
           foreground=[("selected", NHAN)],
           expand=[("selected", (0, 0, 0, 0))])

    st.configure("Treeview", background=THE, fieldbackground=THE, foreground=CHU,
                 bordercolor=VIEN, borderwidth=0, rowheight=28)
    st.configure("Treeview.Heading", background="#eef1f5", foreground=CHU_PHU,
                 relief="flat", padding=(8, 6), font=(chu[0], chu[1], "bold"))
    st.map("Treeview.Heading", background=[("active", "#e2e7ee")])
    st.map("Treeview",
           background=[("selected", NHAN_NHAT)], foreground=[("selected", CHU)])

    for ten in ("TEntry", "TCombobox", "TSpinbox"):
        st.configure(ten, fieldbackground=THE, background=THE, bordercolor=VIEN_NUT,
                     lightcolor=VIEN, darkcolor=VIEN, arrowcolor=CHU_PHU,
                     padding=(6, 5), insertcolor=CHU, relief="solid", borderwidth=1)
        st.map(ten,
               bordercolor=[("focus", NHAN), ("hover", VIEN_DAM)],
               lightcolor=[("focus", NHAN)], darkcolor=[("focus", NHAN)],
               fieldbackground=[("readonly", THE), ("disabled", NEN)])

    st.configure("TCheckbutton", background=NEN, foreground=CHU, focuscolor=NEN)
    st.map("TCheckbutton",
           background=[("active", NEN)],
           indicatorcolor=[("selected", NHAN)])
    st.configure("TRadiobutton", background=NEN, foreground=CHU, focuscolor=NEN)
    st.map("TRadiobutton", background=[("active", NEN)],
           indicatorcolor=[("selected", NHAN)])

    st.configure("Vertical.TScrollbar", background="#dde1e7", troughcolor=NEN,
                 bordercolor=NEN, arrowcolor=CHU_PHU, relief="flat")
    st.configure("Horizontal.TScrollbar", background="#dde1e7", troughcolor=NEN,
                 bordercolor=NEN, arrowcolor=CHU_PHU, relief="flat")
    for huong in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
        st.map(huong, background=[("active", "#c6ccd5")])

    st.configure("TSeparator", background=VIEN)
    st.configure("TProgressbar", background=NHAN, troughcolor="#e3e7ec", bordercolor=NEN)


def dat_kich_thuoc(cua_so, rong_muon, cao_muon, rong_toi_thieu, cao_toi_thieu):
    """Mở cửa sổ vừa với màn hình đang dùng, canh giữa, và chặn thu nhỏ quá đà.

    Không ghi cứng 1040x740: máy bán hàng hay dùng màn 1366x768 hoặc laptop 13"
    chia tỉ lệ, mở cố định cỡ đó là tràn khỏi màn hình, nửa dưới (chỗ để nút
    Lưu) nằm ngoài vùng nhìn thấy mà người dùng không biết kéo lên.

    Cũng chặn chiều ngược lại: co nhỏ quá thì các thanh nút bị cắt cụt, bấm
    không tới. minsize giữ cho mọi nút luôn trong tầm với.
    """
    cua_so.update_idletasks()
    rong_man, cao_man = cua_so.winfo_screenwidth(), cua_so.winfo_screenheight()

    # Chừa chỗ cho thanh taskbar/dock, nên lấy 92% màn hình làm trần
    rong = max(320, min(rong_muon, int(rong_man * 0.92)))
    cao = max(320, min(cao_muon, int(cao_man * 0.92)))
    x = max(0, (rong_man - rong) // 2)
    y = max(0, (cao_man - cao) // 3)      # hơi lệch lên trên, nhìn cân mắt hơn
    cua_so.geometry(f"{rong}x{cao}+{x}+{y}")
    cua_so.minsize(min(rong_toi_thieu, rong), min(cao_toi_thieu, cao))


class HangNut(ttk.Frame):
    """Hàng nút tự xuống dòng khi cửa sổ hẹp lại.

    pack(side="left") không biết xuống dòng: thu nhỏ cửa sổ là mấy nút cuối
    hàng bị đẩy ra ngoài, không thấy và không bấm được — mà toàn là nút hay
    dùng (Ghi vào file Excel, Tải lại từ Excel). Frame này tự xếp lại chỗ mỗi
    lần đổi kích thước, thiếu chỗ thì xuống hàng dưới.

    Dùng place() cho từng nút nên frame không tự biết mình cao bao nhiêu —
    phải tự tính và tự đặt chiều cao, xem _xep().
    """

    def __init__(self, cha, khoang_doc=6, **kw):
        super().__init__(cha, **kw)
        self._muc = []           # [(widget, khoảng cách với nút trước)]
        self._khoang_doc = khoang_doc
        self._cao_hien = 0
        self.bind("<Configure>", lambda _e: self._xep())

    def them(self, widget, khoang_truoc=6):
        """Thêm một widget vào hàng. Trả về chính nó để gán biến cho gọn."""
        self._muc.append((widget, khoang_truoc))
        self._xep()
        return widget

    def _xep(self):
        rong = self.winfo_width()
        if rong <= 1 or not self._muc:
            return

        x = y = cao_hang = 0
        for i, (w, truoc) in enumerate(self._muc):
            rw, rh = w.winfo_reqwidth(), w.winfo_reqheight()
            lui = 0 if x == 0 else truoc
            if x and x + lui + rw > rong:       # không đủ chỗ → xuống hàng
                x, lui = 0, 0
                y += cao_hang + self._khoang_doc
                cao_hang = 0
            w.place(x=x + lui, y=y, width=rw, height=rh)
            x += lui + rw
            cao_hang = max(cao_hang, rh)

        # Đổi chiều cao lại kích hoạt <Configure>, nên chỉ đổi khi thực sự khác
        # — không thì thành vòng lặp vô tận làm treo giao diện.
        can = y + cao_hang
        if can and can != self._cao_hien:
            self._cao_hien = can
            self.configure(height=can)
