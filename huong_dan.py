"""Cửa sổ Hướng dẫn sử dụng: cột mục lục bên trái, nội dung bên phải.

Nội dung đọc thẳng từ HUONGDAN_NGUOI_DUNG.md chứ không chép lại vào code — một
tài liệu, một nguồn. Sửa file .md là cửa sổ này đổi theo, khỏi lo hướng dẫn
trong tool nói một đằng file nói một nẻo.
"""
import os
import re
import textwrap
import unicodedata
import tkinter as tk
from tkinter import ttk

import giao_dien
import paths

TEN_FILE = "HUONGDAN_NGUOI_DUNG.md"


def duong_dan_file():
    """Tìm file hướng dẫn ở cả bản chạy mã nguồn lẫn bản .exe đã đóng gói."""
    for dd in (paths.bundle_path(TEN_FILE),
               paths.data_path(TEN_FILE),
               os.path.join(os.path.dirname(os.path.abspath(__file__)), TEN_FILE)):
        if os.path.isfile(dd):
            return dd
    return None


def doc_cac_muc(noi_dung):
    """Cắt file markdown thành [(tên mục, các dòng), ...] theo tiêu đề '## '.

    Phần đầu file (trước mục đầu tiên) thành mục "Giới thiệu", nhưng bỏ khối
    **Mục lục** đi: cột bên trái đã là mục lục rồi, để lại thành ra có hai cái
    mục lục chồng nhau.
    """
    muc, ten, dong = [], "Giới thiệu", []
    bo_qua_muc_luc = False
    for d in noi_dung.splitlines():
        if d.startswith("## "):
            muc.append((ten, dong))
            ten, dong = d[3:].strip(), []
            bo_qua_muc_luc = False
            continue
        if d.startswith("# "):           # tiêu đề toàn tài liệu, cột trái lo rồi
            continue
        if d.strip() == "**Mục lục**":
            bo_qua_muc_luc = True
            continue
        if bo_qua_muc_luc:
            if d.strip() in ("", "---"):
                bo_qua_muc_luc = d.strip() == ""
            continue
        dong.append(d)
    muc.append((ten, dong))
    return [(t, d) for t, d in muc if any(x.strip() for x in d)]


def so_cua_muc(ten):
    """Lấy số đứng đầu tên mục ("3. Làm quen giao diện" → "3"), không có thì "".

    Dùng để đánh số mục con thành 3.1, 3.2 — tự suy ra từ tiêu đề nên tài liệu
    không phải ghi số tay, đổi thứ tự mục cũng không sợ lệch số.
    """
    m = re.match(r"^(\d+)\.", ten.strip())
    return m.group(1) if m else ""


def _bo_danh_dau(chu):
    """Gỡ cú pháp markdown còn sót lại trong một dòng chữ thường."""
    chu = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", chu)   # [chữ](link) → chữ
    return chu.replace("**", "").replace("`", "")


def _be_rong(chu):
    """Đếm một chuỗi chiếm bao nhiêu ô trên font đều nét.

    Không dùng len() được: emoji và ký tự Á Đông chiếm 2 ô, còn dấu tiếng Việt
    rời (dạng tổ hợp) chiếm 0 ô. Đếm sai là khung bảng lệch hẳn một cột.
    """
    rong = 0
    for k in chu:
        if unicodedata.combining(k):
            continue
        rong += 2 if (unicodedata.east_asian_width(k) in ("W", "F")
                      or ord(k) >= 0x1F300) else 1
    return rong


def _dem_o(chu, rong):
    """Thêm khoảng trắng cho chuỗi đủ `rong` ô hiển thị."""
    return chu + " " * max(0, rong - _be_rong(chu))


def doc_bang(cac_dong):
    """Đổi các dòng |...|...| của markdown thành danh sách hàng, bỏ dòng |---|."""
    hang = []
    for d in cac_dong:
        o = [c.strip() for c in d.strip().strip("|").split("|")]
        if all(set(c) <= set("-: ") for c in o):
            continue
        hang.append([_bo_danh_dau(c) for c in o])
    return hang


def ve_bang(hang, rong_toi_da=92):
    """Vẽ bảng có khung bằng ký tự kẻ ô, tự xuống dòng trong ô nếu chữ dài.

    Trước đây chỉ căn cột bằng khoảng trắng: ô nào chữ dài là cả bảng kéo ngang
    quá khổ, mắt dò từ cột đầu sang cột cuối hay lạc hàng. Có khung và có ngắt
    dòng trong ô thì bảng luôn nằm gọn trong bề ngang đang có.
    """
    if not hang:
        return ""
    so_cot = max(len(h) for h in hang)
    hang = [h + [""] * (so_cot - len(h)) for h in hang]

    # Khung và lề mỗi ô ăn mất: mỗi cột 3 ô ("│ " + " ") cộng 1 ô vạch cuối
    cho_chu = rong_toi_da - (3 * so_cot + 1)
    tu_nhien = [max(_be_rong(h[i]) for h in hang) for i in range(so_cot)]

    rong = tu_nhien[:]
    # Thừa chỗ thì thôi; thiếu chỗ thì bớt dần ở cột đang rộng nhất, cột nào
    # cũng được giữ tối thiểu 10 ô để không bị bẻ thành từng chữ cái một.
    while sum(rong) > cho_chu and max(rong) > 10:
        rong[rong.index(max(rong))] -= 1
    rong = [max(r, 1) for r in rong]

    def ngat(o, r):
        return textwrap.wrap(o, width=r, break_long_words=False) or [""]

    ngang = lambda trai, giua, phai: (
        trai + giua.join("─" * (r + 2) for r in rong) + phai)

    ra = [ngang("┌", "┬", "┐")]
    for chi_so, h in enumerate(hang):
        o_da_ngat = [ngat(h[i], rong[i]) for i in range(so_cot)]
        cao = max(len(x) for x in o_da_ngat)
        for d in range(cao):
            ra.append("│ " + " │ ".join(
                _dem_o(o_da_ngat[i][d] if d < len(o_da_ngat[i]) else "", rong[i])
                for i in range(so_cot)) + " │")
        if chi_so == 0:                     # vạch đậm dưới hàng tiêu đề
            ra.append(ngang("├", "┼", "┤"))
    ra.append(ngang("└", "┴", "┘"))
    return "\n".join(ra)


class CuaSoHuongDan(tk.Toplevel):
    """Một cửa sổ, cột mục bên trái, nội dung bên phải, có ô tìm kiếm."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Hướng dẫn sử dụng")
        giao_dien.dat_kich_thuoc(self, 1000, 700, 720, 480)
        # Cố ý KHÔNG transient(parent): cửa sổ transient trên một số bản Tk của
        # macOS bị dán cứng vào cửa sổ cha, nằm đè lên và không nhận được click.
        # Đây là tài liệu để đọc song song khi thao tác, nên để nó là cửa sổ độc
        # lập, có mục riêng trên Dock, chuyển qua lại bằng Cmd+` được.
        self._tach_khoi_cua_so_cha(parent)

        dd = duong_dan_file()
        if not dd:
            ttk.Label(self, text=f"Không tìm thấy file {TEN_FILE} cạnh tool.",
                      style="CanhBao.TLabel").pack(padx=20, pady=20)
            return
        with open(dd, encoding="utf-8") as fh:
            self.muc = doc_cac_muc(fh.read())

        # Dùng frame thường chứ không PanedWindow: trên Tk 8.6 (bản đi kèm máy
        # build) thanh chia đôi có lúc không được đặt vị trí, cột trái co về 0
        # và biến mất — nhìn như cửa sổ hỏng, không bấm vào mục nào được.
        khung = ttk.Frame(self)
        khung.pack(fill="both", expand=True, padx=10, pady=10)

        trai = ttk.Frame(khung, width=300)
        trai.pack(side="left", fill="y", padx=(0, 12))
        trai.pack_propagate(False)      # giữ đúng 300px, không co theo nội dung
        ttk.Label(trai, text="Nội dung", style="TieuDe.TLabel").pack(anchor="w", pady=(0, 6))
        self.o_tim = tk.StringVar()
        o = ttk.Entry(trai, textvariable=self.o_tim, width=30)
        o.pack(fill="x", pady=(0, 6))
        o.insert(0, "")
        self.o_tim.trace_add("write", lambda *_: self._loc())
        ttk.Label(trai, text="Gõ để lọc mục", style="Phu.TLabel").pack(anchor="w")

        self.ds = tk.Listbox(trai, width=30, exportselection=False, activestyle="none")
        self.ds.pack(fill="both", expand=True, pady=6)
        self.ds.bind("<<ListboxSelect>>", lambda _e: self._hien())

        phai = ttk.Frame(khung)
        phai.pack(side="left", fill="both", expand=True)

        # Thanh vị trí: đang đọc mục nào, phần mấy — cuộn giữa trang vẫn biết
        self.duong_dan_doc = ttk.Label(phai, text="", style="Phu.TLabel")
        self.duong_dan_doc.pack(anchor="w", pady=(0, 6))

        than = ttk.Frame(phai)
        than.pack(fill="both", expand=True)
        self.o_chu = tk.Text(than, wrap="word", padx=24, pady=16, state="disabled",
                             cursor="arrow")
        cuon = ttk.Scrollbar(than, command=self.o_chu.yview)
        self.o_chu.configure(yscrollcommand=cuon.set)
        cuon.pack(side="right", fill="y")
        self.o_chu.pack(fill="both", expand=True)
        self._tao_the()

        # Đọc tài liệu thường là đọc một mạch từ đầu tới cuối, nên có nút sang
        # mục kế tiếp ngay dưới chân trang, khỏi phải rê lên cột trái mỗi lần.
        chan = ttk.Frame(phai)
        chan.pack(fill="x", pady=(8, 0))
        self.nut_truoc = ttk.Button(chan, text="← Mục trước",
                                    command=lambda: self._nhay(-1))
        self.nut_truoc.pack(side="left")
        self.nut_sau = ttk.Button(chan, text="Mục tiếp →", command=lambda: self._nhay(1))
        self.nut_sau.pack(side="left", padx=6)
        ttk.Button(chan, text="Đóng", command=self.destroy).pack(side="right")

        # Ảnh đã thu nhỏ theo bề ngang lúc vẽ, nên kéo rộng cửa sổ phải vẽ lại
        # thì ảnh mới to theo. Chỉ vẽ khi bề ngang đổi thật, không thì mỗi lần
        # nhích chuột là vẽ lại cả trang.
        self._rong_cu = 0
        self.o_chu.bind("<Configure>", self._khi_doi_co)

        self._loc()
        if self.ds.size():
            self.ds.selection_set(0)
            self._hien()
        self.after(50, self._len_truoc)

    def _tach_khoi_cua_so_cha(self, cha):
        """Dời cửa sổ lệch khỏi cửa sổ chính để thấy rõ là hai cửa sổ riêng.

        Cả hai đều mở ở giữa màn hình nên chồng khít lên nhau, nhìn tưởng một
        cửa sổ bị kẹt — bấm vào phần thò ra lại hoá ra đang bấm vào cửa sổ dưới.
        """
        try:
            cha.update_idletasks()
            x = cha.winfo_rootx() + 60
            y = cha.winfo_rooty() + 40
            rong, cao = self.winfo_width(), self.winfo_height()
            x = max(0, min(x, self.winfo_screenwidth() - rong))
            y = max(0, min(y, self.winfo_screenheight() - cao))
            self.geometry(f"+{x}+{y}")
        except tk.TclError:
            pass

    def _len_truoc(self):
        """Đưa cửa sổ lên trên và nhận bàn phím/chuột."""
        try:
            self.lift()
            self.focus_force()
        except tk.TclError:
            pass

    def _nhay(self, buoc):
        """Sang mục kế tiếp/trước đó trong cột trái."""
        chon = self.ds.curselection()
        if not chon:
            return
        moi = max(0, min(self.ds.size() - 1, chon[0] + buoc))
        self.ds.selection_clear(0, "end")
        self.ds.selection_set(moi)
        self.ds.see(moi)
        self._hien()

    def _khi_doi_co(self, su_kien):
        """Vẽ lại khi ô chữ đổi bề ngang, để bảng và ảnh co giãn theo.

        Cờ _dang_ve là bắt buộc: vẽ lại làm đổi bố cục, đổi bố cục lại sinh
        <Configure> mới. Thiếu cờ thì hai việc đó gọi nhau vòng tròn và giao
        diện đứng hình — không báo lỗi gì, chỉ là bấm vào đâu cũng không ăn.
        """
        if getattr(self, "_dang_ve", False):
            return
        if abs(su_kien.width - self._rong_cu) < 40:
            return
        self._rong_cu = su_kien.width
        self._dang_xem = None          # ép vẽ lại ở lần _hien() tới
        self._dang_ve = True
        try:
            self._hien()
        finally:
            self._dang_ve = False

    # ---------- hiển thị ----------

    def _tao_the(self):
        """Khai màu/kiểu chữ cho từng loại nội dung.

        Chữ hướng dẫn là chữ để đọc liền mạch chứ không phải nhãn trên nút, nên
        cỡ to hơn mặc định 1 bậc và giãn dòng rộng hơn — đọc nửa trang mới đỡ mỏi.
        """
        from tkinter import font as tkfont
        nen = tkfont.nametofont("TkDefaultFont")
        ho, co = nen.cget("family"), nen.cget("size") + 1
        self._font_than = (ho, co)
        t = self.o_chu.tag_configure

        t("tieu_de", font=(ho, co + 7, "bold"), spacing3=16, foreground=giao_dien.CHU)
        t("muc", font=(ho, co + 3, "bold"), spacing1=22, spacing3=8,
          foreground=giao_dien.NHAN)
        t("thuong", font=(ho, co), spacing1=2, spacing3=10, lmargin1=2, lmargin2=2)
        t("dam", font=(ho, co, "bold"))

        t("gach_dau", font=(ho, co), lmargin1=16, lmargin2=34, spacing1=2, spacing3=6)
        t("so_thu_tu", font=(ho, co), lmargin1=16, lmargin2=38, spacing1=2, spacing3=6)

        # Khối lưu ý: nền vàng nhạt + lề rộng để nhìn ra ngay là chỗ cần dừng đọc
        t("trich", font=(ho, co), lmargin1=18, lmargin2=18, rmargin=18,
          background="#fdf6e3", foreground="#8a5a00", spacing1=6, spacing3=6)
        t("trich_gach", font=(ho, co), lmargin1=34, lmargin2=50, rmargin=18,
          background="#fdf6e3", foreground="#8a5a00", spacing1=2, spacing3=6)

        t("ma", font="TkFixedFont", background="#eef1f5", lmargin1=18, lmargin2=18,
          spacing1=6, spacing3=6)
        t("bang", font="TkFixedFont", lmargin1=12, lmargin2=12, spacing3=2)

        # Bảng 2 cột đổ thành danh sách: tên đứng riêng một dòng màu nhấn,
        # phần giải thích thụt vào dưới
        t("nhan_bang", font=(ho, co - 2, "bold"), foreground=giao_dien.CHU_PHU,
          spacing1=16, spacing3=6)
        t("khoa", font=(ho, co, "bold"), foreground=giao_dien.NHAN,
          lmargin1=14, lmargin2=14, spacing1=6, spacing3=2)
        t("gia_tri", font=(ho, co), lmargin1=30, lmargin2=30, rmargin=12, spacing3=6)

        t("ngan", spacing1=10, spacing3=10, foreground=giao_dien.VIEN)
        t("chu_thich_anh", font=(ho, co - 1), foreground=giao_dien.CHU_PHU,
          lmargin1=8, lmargin2=8, spacing3=14)
        t("thieu_anh", font=(ho, co - 1), foreground=giao_dien.CHU_PHU,
          background="#f0f2f5", spacing1=6, spacing3=10)

    def _loc(self):
        """Dựng lại cột trái: mỗi mục '## ' một dòng, mỗi '### ' một dòng con."""
        tim = self.o_tim.get().strip().lower()
        self.hien_thi = []          # [(chỉ số mục, tiêu đề con hoặc None)]
        self.ds.delete(0, "end")
        for i, (ten, dong) in enumerate(self.muc):
            if tim and tim not in ten.lower() and tim not in "\n".join(dong).lower():
                continue
            self.ds.insert("end", "  " + _bo_danh_dau(ten))
            self.ds.itemconfig("end", foreground=giao_dien.CHU)
            self.hien_thi.append((i, None))
            so_cha, thu_tu = so_cua_muc(ten), 0
            for d in dong:
                if d.startswith("### "):
                    thu_tu += 1
                    con = _bo_danh_dau(d[4:].strip())
                    danh_so = f"{so_cha}.{thu_tu}  " if so_cha else ""
                    self.ds.insert("end", f"      {danh_so}{con}")
                    self.ds.itemconfig("end", foreground=giao_dien.CHU_PHU)
                    self.hien_thi.append((i, con))
        if self.hien_thi:
            self.ds.selection_set(0)
            self._hien()

    def _hien(self):
        chon = self.ds.curselection()
        if not chon or not self.hien_thi:
            return
        chi_so, tieu_de_con = self.hien_thi[chon[0]]
        ten, cac_dong = self.muc[chi_so]

        # Chọn mục con thì vẫn đổ cả mục ra rồi mới nhảy tới chỗ đó: đọc tiếp
        # xuống phần sau được, chứ cắt lẻ ra thì cụt mạch.
        if chi_so != getattr(self, "_dang_xem", None):
            self._can_le()
            self.o_chu.configure(state="normal")
            self.o_chu.delete("1.0", "end")
            self._anh_dang_giu = []          # giữ tham chiếu kẻo Tk dọn mất ảnh
            self._moc_con = {}
            self.o_chu.insert("end", _bo_danh_dau(ten) + "\n", "tieu_de")
            self._do_chu(cac_dong, so_cua_muc(ten))
            self.o_chu.configure(state="disabled")
            self._dang_xem = chi_so

        vi_tri = self.ds.curselection()[0]
        self.duong_dan_doc.config(
            text=_bo_danh_dau(ten) + (f"   ›   {tieu_de_con}" if tieu_de_con else "")
                 + f"      ({vi_tri + 1}/{self.ds.size()})")
        self.nut_truoc.config(state="normal" if vi_tri else "disabled")
        self.nut_sau.config(
            state="disabled" if vi_tri >= self.ds.size() - 1 else "normal")

        if tieu_de_con and tieu_de_con in self._moc_con:
            self.o_chu.see(self._moc_con[tieu_de_con])
            self.o_chu.yview_scroll(-1, "units")
        else:
            self.o_chu.yview_moveto(0)

    # Dòng chữ dài quá thì mắt nhảy xuống dòng dễ bị lạc hàng; sách báo giữ
    # khoảng 60–90 ký tự một dòng. Cửa sổ rộng ra thì chừa lề hai bên chứ không
    # kéo chữ dài mãi.
    RONG_DOC_TOI_DA = 90        # tính theo số ký tự

    def _can_le(self):
        """Đặt lề trái/phải cho ô chữ sao cho dòng chữ không dài quá mức đọc tốt."""
        from tkinter import font as tkfont
        rong_chu = tkfont.Font(font=self._font_than).measure("x") or 8
        toi_da = self.RONG_DOC_TOI_DA * rong_chu
        thua = self.o_chu.winfo_width() - toi_da
        self._le = 24 if thua <= 48 else min(thua // 2, 160)
        self.o_chu.configure(padx=self._le)

    def _so_o_ngang(self):
        """Ô chữ rộng được bao nhiêu ký tự font đều nét — để kẻ bảng vừa khít."""
        from tkinter import font as tkfont
        rong_mot_ky_tu = tkfont.nametofont("TkFixedFont").measure("0") or 8
        co = (self.o_chu.winfo_width() - 2 * getattr(self, "_le", 24) - 24) // rong_mot_ky_tu
        return max(40, min(co, 110))

    def _do_chu(self, cac_dong, so_muc=""):
        thu_tu_con = 0
        for loai, chu in gom_khoi(cac_dong):
            if loai == "ma":
                self.o_chu.insert("end", chu + "\n\n", "ma")
            elif loai == "bang":
                if chu and max(len(h) for h in chu) == 2:
                    self._chen_danh_sach(chu)
                else:
                    self.o_chu.insert("end", ve_bang(chu, self._so_o_ngang()) + "\n\n", "bang")
            elif loai == "ngan":
                self.o_chu.insert("end", "─" * 60 + "\n", "ngan")
            elif loai == "trong":
                self.o_chu.insert("end", "\n")
            elif loai == "anh":
                self._chen_anh(*chu)
            else:
                if loai == "muc":
                    thu_tu_con += 1
                    self._moc_con[_bo_danh_dau(chu)] = self.o_chu.index("end-1c")
                    if so_muc:
                        chu = f"{so_muc}.{thu_tu_con}  {chu}"
                self._dong_co_dam(_bo_danh_dau_giu_dam(chu), loai)

    def _chen_danh_sach(self, hang):
        """Đổ bảng 2 cột thành danh sách tên → giải thích.

        Gần như mọi bảng trong tài liệu đều dạng "Nút | Làm gì". Kẻ khung cho
        dạng đó vừa tốn chỗ vừa khó đọc: cột phải bị bóp hẹp nên câu nào cũng
        gãy làm ba bốn dòng. Để tên riêng một dòng, giải thích thụt xuống dưới
        thì chữ chạy hết bề ngang, đọc như danh sách.
        """
        tieu_de = hang[0]
        self.o_chu.insert("end", f"{_bo_danh_dau(tieu_de[0])} — {_bo_danh_dau(tieu_de[1])}\n",
                          "nhan_bang")
        for h in hang[1:]:
            self._dong_co_dam(_bo_danh_dau_giu_dam(h[0]), "khoa")
            self._dong_co_dam(_bo_danh_dau_giu_dam(h[1] if len(h) > 1 else ""), "gia_tri")
        self.o_chu.insert("end", "\n")

    def _chen_anh(self, chu_thich, duong_dan_anh):
        """Vẽ ảnh minh hoạ, tự thu nhỏ cho vừa bề ngang ô chữ.

        Thiếu file thì ghi một dòng xám báo thiếu chứ không để trống: người đọc
        biết chỗ đó đáng lẽ có hình, và người sửa tài liệu biết phải bổ sung.
        """
        dd = duong_dan_anh
        if not os.path.isabs(dd):
            goc = duong_dan_file()
            dd = os.path.join(os.path.dirname(goc) if goc else ".", dd)
        if not os.path.isfile(dd):
            self.o_chu.insert("end", f"[chưa có ảnh: {duong_dan_anh}]\n", "thieu_anh")
            return
        try:
            anh = tk.PhotoImage(file=dd)
        except tk.TclError as e:
            self.o_chu.insert("end", f"[không mở được ảnh {duong_dan_anh}: {e}]\n", "thieu_anh")
            return

        # PhotoImage chỉ thu nhỏ được theo số nguyên lần, nên lấy bội số nhỏ
        # nhất đủ để ảnh lọt vào bề ngang hiện có.
        rong_o = max(self.o_chu.winfo_width() - 60, 360)
        if anh.width() > rong_o:
            lan = -(-anh.width() // rong_o)
            anh = anh.subsample(lan, lan)
        self._anh_dang_giu.append(anh)
        self.o_chu.image_create("end", image=anh, padx=4, pady=6)
        self.o_chu.insert("end", "\n")
        if chu_thich:
            self.o_chu.insert("end", chu_thich + "\n\n", "chu_thich_anh")

    def _dong_co_dam(self, chu, the_nen):
        """Chèn một dòng, phần nằm giữa ** ** thì in đậm."""
        for phan, in_dam in tach_dam(chu):
            self.o_chu.insert("end", phan, (the_nen, "dam") if in_dam else the_nen)
        self.o_chu.insert("end", "\n", the_nen)


def _bo_danh_dau_giu_dam(chu):
    """Như _bo_danh_dau nhưng giữ lại ** ** để còn tô đậm."""
    chu = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", chu)
    return chu.replace("`", "")


def tach_dam(chu):
    """Cắt chuỗi thành [(đoạn, có in đậm không), ...] theo cặp dấu **."""
    ra, phan = [], chu.split("**")
    for i, p in enumerate(phan):
        if p:
            ra.append((p, i % 2 == 1))
    return ra


def gom_khoi(cac_dong):
    """Gom các dòng markdown thành khối để hiển thị, trả về [(loại, chữ), ...].

    File .md ngắt dòng cứng ở cột 78 cho dễ đọc trong editor. Nếu đổ y nguyên
    vào ô chữ thì cửa sổ rộng bao nhiêu đoạn văn vẫn cụt ở cột 78, nhìn như thơ.
    Nên nối các dòng cùng một đoạn lại thành một dòng dài, để ô chữ tự xuống
    dòng theo bề ngang thật của cửa sổ.

    Dòng nối tiếp của mục đánh số hay gạch đầu dòng trong file này được thụt
    vào, nên cứ thấy thụt đầu dòng là nối vào mục ngay trước.
    """
    khoi, i = [], 0
    while i < len(cac_dong):
        d = cac_dong[i]
        cat = d.strip()

        if cat.startswith("```"):
            i += 1
            trong = []
            while i < len(cac_dong) and not cac_dong[i].strip().startswith("```"):
                trong.append(cac_dong[i])
                i += 1
            khoi.append(("ma", "\n".join(trong)))
        elif cat.startswith("|"):
            bang = []
            while i < len(cac_dong) and cac_dong[i].strip().startswith("|"):
                bang.append(cac_dong[i])
                i += 1
            khoi.append(("bang", doc_bang(bang)))
            continue
        elif cat in ("---", "***"):
            khoi.append(("ngan", ""))
        elif re.match(r"^!\[[^\]]*\]\([^)]+\)$", cat):      # ![chú thích](đường dẫn)
            m = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)$", cat)
            khoi.append(("anh", (m.group(1), m.group(2))))
        elif cat.startswith("### "):
            khoi.append(("muc", cat[4:]))
        elif cat.startswith(">"):
            # Giữ nguyên khoảng trắng sau dấu ">" chứ không strip: dòng thụt
            # vào là dòng nối tiếp của gạch đầu dòng ngay trên, cắt mất khoảng
            # trắng là không phân biệt được nữa, câu sau rơi ra thành đoạn rời.
            trong = [re.sub(r"^>\s?", "", cat)]
            while i + 1 < len(cac_dong) and cac_dong[i + 1].strip().startswith(">"):
                i += 1
                trong.append(re.sub(r"^\s*>\s?", "", cac_dong[i]))

            loai_dang_gom, dang_gom = None, []

            def xa():
                if dang_gom:
                    khoi.append((loai_dang_gom, " ".join(dang_gom)))

            for d_trich in trong:
                if d_trich.strip().startswith(("- ", "* ")):
                    xa()
                    loai_dang_gom, dang_gom = "trich_gach", ["• " + d_trich.strip()[2:]]
                elif not d_trich.strip():
                    xa()
                    loai_dang_gom, dang_gom = None, []
                elif d_trich[:1] in (" ", "\t") and loai_dang_gom == "trich_gach":
                    dang_gom.append(d_trich.strip())          # nối tiếp gạch trên
                else:
                    if loai_dang_gom != "trich":
                        xa()
                        loai_dang_gom, dang_gom = "trich", []
                    dang_gom.append(d_trich.strip())
            xa()
        elif cat.startswith(("- ", "* ")) or re.match(r"^\d+\. ", cat):
            loai = "gach_dau" if cat.startswith(("- ", "* ")) else "so_thu_tu"
            dau = "• " + cat[2:] if loai == "gach_dau" else cat
            noi = [dau]
            while (i + 1 < len(cac_dong) and cac_dong[i + 1].strip()
                   and cac_dong[i + 1][:1] in (" ", "\t")):
                i += 1
                noi.append(cac_dong[i].strip())
            khoi.append((loai, " ".join(noi)))
        elif not cat:
            khoi.append(("trong", ""))
        else:
            noi = [cat]
            while (i + 1 < len(cac_dong) and cac_dong[i + 1].strip()
                   and not cac_dong[i + 1].strip().startswith(
                       ("|", "- ", "* ", "> ", "#", "```", "---"))
                   and not re.match(r"^\d+\. ", cac_dong[i + 1].strip())):
                i += 1
                noi.append(cac_dong[i].strip())
            khoi.append(("thuong", " ".join(noi)))
        i += 1
    return khoi
