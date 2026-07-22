"""
Hẹn giờ đăng bài — phần tính toán thuần, không dính tới giao diện.

Mọi mốc thời gian trong file này đều theo giờ Việt Nam (UTC+7), cố định cứng
chứ không lấy múi giờ của máy. Lý do: người dùng khai "19:30" là có ý 19:30 giờ
Việt Nam; nếu máy họ lỡ đặt sai múi giờ (máy mua lại, cài Windows để mặc định
Pacific...) thì bài sẽ lên lệch cả nửa ngày mà không ai hiểu vì sao.

Hai kiểu lịch:
  mot_lan   — đúng một ngày giờ cụ thể, chạy xong thì tự tắt
  hang_ngay — như báo thức: một mốc giờ, lặp lại vào các thứ được chọn

Muốn nhiều mốc trong ngày thì tạo nhiều lịch, mỗi lịch một mốc giờ.

Tool là ứng dụng chạy trên máy, không phải server: đến giờ hẹn mà máy tắt hoặc
tool chưa mở thì không có gì chạy cả. Phần "bỏ lỡ" bên dưới sinh ra để xử lý
đúng tình huống đó — mở tool muộn thì hỏi lại người dùng có chạy bù không.
"""
import os
import subprocess
import sys
import uuid
from datetime import datetime, timedelta, timezone

MUI_GIO_VN = timezone(timedelta(hours=7), "VN")

# Tên thứ theo chỉ số của datetime.weekday(): 0 = thứ 2 ... 6 = chủ nhật
TEN_THU = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ nhật"]
TEN_THU_NGAN = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"]

# Vòng lặp kiểm tra chạy mỗi ~20 giây nên không bao giờ trúng đúng giây 00 của
# mốc hẹn. Trễ trong khoảng này vẫn coi là "đến giờ" và chạy thẳng, không hỏi.
DUNG_SAI = timedelta(minutes=2)

# Trễ quá DUNG_SAI thì coi là bỏ lỡ và hỏi lại người dùng — nhưng chỉ hỏi nếu
# còn trong cửa sổ này. Mở tool sau 3 ngày mà vẫn hỏi "chạy bù lịch hôm kia?"
# thì vô nghĩa, và bấm nhầm Yes là ăn nguyên loạt bài sai khung giờ.
CUA_SO_HOI = timedelta(hours=12)

LICH_MAC_DINH = {
    "id": "",
    "bat": True,
    "kieu": "hang_ngay",     # "mot_lan" hoặc "hang_ngay"
    "gio": "19:30",          # "HH:MM" giờ Việt Nam
    "ngay": "",              # "YYYY-MM-DD", chỉ dùng cho kiểu mot_lan
    "thu": [],               # [0..6] cho kiểu hang_ngay; rỗng = mọi ngày
    "campaign": "",
    "nick": "",
    "lan_chay_cuoi": "",     # ISO của lần thực sự chạy gần nhất
    "moc_da_xu_ly": "",      # ISO của mốc gần nhất đã giải quyết xong
}


def bay_gio():
    """Thời điểm hiện tại theo giờ Việt Nam."""
    return datetime.now(MUI_GIO_VN)


def lich_moi():
    """Một lịch rỗng với giá trị mặc định và id riêng."""
    l = dict(LICH_MAC_DINH)
    l["thu"] = []
    l["id"] = uuid.uuid4().hex[:8]
    return l


def chuan_hoa(lich):
    """Bù các key còn thiếu — lịch cũ đọc từ config.json có thể chưa có đủ."""
    for key, val in LICH_MAC_DINH.items():
        lich.setdefault(key, list(val) if isinstance(val, list) else val)
    if not lich["id"]:
        lich["id"] = uuid.uuid4().hex[:8]
    return lich


# ==================== TÍNH MỐC THỜI GIAN ====================

def _gio_phut(lich):
    """Tách "HH:MM" thành (giờ, phút); None nếu người dùng gõ sai định dạng."""
    try:
        phan = str(lich.get("gio", "")).strip().split(":")
        gio, phut = int(phan[0]), int(phan[1])
    except (ValueError, IndexError):
        return None
    if 0 <= gio <= 23 and 0 <= phut <= 59:
        return gio, phut
    return None


def gio_hop_le(chuoi):
    """Chuỗi giờ người dùng gõ có đúng dạng HH:MM không."""
    return _gio_phut({"gio": chuoi}) is not None


def _ghep(ngay, gio, phut):
    return datetime(ngay.year, ngay.month, ngay.day, gio, phut, tzinfo=MUI_GIO_VN)


def _cac_moc_quanh(lich, moc_goc):
    """Các mốc chạy rơi vào khoảng hôm qua → ngày mai, sắp xếp tăng dần.

    Lấy cả hôm qua vì lịch 23:00 mà người dùng mở tool lúc 0h30 hôm sau thì đó
    vẫn là một lần bỏ lỡ cần hỏi, không phải chuyện đã qua từ lâu.
    """
    hp = _gio_phut(lich)
    if hp is None:
        return []
    gio, phut = hp

    if lich.get("kieu") == "mot_lan":
        try:
            ngay = datetime.strptime(lich.get("ngay", ""), "%Y-%m-%d").date()
        except ValueError:
            return []
        return [_ghep(ngay, gio, phut)]

    thu = lich.get("thu") or list(range(7))
    ds = []
    for lech in (-1, 0, 1):
        ngay = (moc_goc + timedelta(days=lech)).date()
        if ngay.weekday() in thu:
            ds.append(_ghep(ngay, gio, phut))
    return sorted(ds)


def lan_ke_tiep(lich, moc_goc=None):
    """Mốc chạy kế tiếp tính từ bây giờ, hoặc None nếu không còn lần nào."""
    moc_goc = moc_goc or bay_gio()
    hp = _gio_phut(lich)
    if hp is None:
        return None
    gio, phut = hp

    if lich.get("kieu") == "mot_lan":
        ds = _cac_moc_quanh(lich, moc_goc)
        return ds[0] if ds and ds[0] > moc_goc else None

    thu = lich.get("thu") or list(range(7))
    for lech in range(0, 8):
        ngay = (moc_goc + timedelta(days=lech)).date()
        if ngay.weekday() not in thu:
            continue
        moc = _ghep(ngay, gio, phut)
        if moc > moc_goc:
            return moc
    return None


def trang_thai(lich, moc_goc=None):
    """Xem lịch này có việc phải làm ngay không.

    Trả về (hành_động, mốc):
      (None, None)        — chưa tới giờ, hoặc mốc gần nhất đã giải quyết rồi
      ("chay", mốc)       — vừa tới giờ, chạy luôn
      ("hoi", mốc)        — đã lỡ giờ nhưng còn mới, hỏi người dùng có chạy bù
      ("qua_han", mốc)    — lỡ quá lâu, bỏ qua và chỉ ghi log
    """
    if not lich.get("bat"):
        return None, None

    moc_goc = moc_goc or bay_gio()
    da_qua = [m for m in _cac_moc_quanh(lich, moc_goc) if m <= moc_goc]
    if not da_qua:
        return None, None

    moc = da_qua[-1]
    if lich.get("moc_da_xu_ly") == moc.isoformat():
        return None, None

    tre = moc_goc - moc
    if tre <= DUNG_SAI:
        return "chay", moc
    if tre <= CUA_SO_HOI:
        return "hoi", moc

    # Lỡ quá lâu. Lịch lặp thì cứ im lặng chờ mốc kế tiếp — 19h hôm nay mà đi
    # càu nhàu về mốc 19h30 hôm qua thì ngày nào cũng có một dòng log vô ích.
    # Lịch một lần thì phải báo, vì nó sẽ không bao giờ tới lượt nữa.
    if lich.get("kieu") == "mot_lan":
        return "qua_han", moc
    return None, None


def danh_dau_da_xu_ly(lich, moc, da_chay=False):
    """Ghi nhận đã giải quyết xong một mốc — chạy, người dùng từ chối, hay bỏ
    qua vì quá hạn đều tính là xong. Không có bước này thì vòng lặp 20 giây sẽ
    hỏi lại người dùng không ngừng nghỉ."""
    lich["moc_da_xu_ly"] = moc.isoformat()
    if da_chay:
        lich["lan_chay_cuoi"] = bay_gio().isoformat()
    return lich


# ==================== HIỂN THỊ ====================

def mo_ta(lich):
    """Câu mô tả ngắn để hiện trong bảng, vd '19:30 · T2, T4, T6'."""
    gio = lich.get("gio") or "??:??"
    if lich.get("kieu") == "mot_lan":
        try:
            ngay = datetime.strptime(lich.get("ngay", ""), "%Y-%m-%d")
            return f"{gio} · ngày {ngay.strftime('%d/%m/%Y')}"
        except ValueError:
            return f"{gio} · (chưa chọn ngày)"

    thu = lich.get("thu") or []
    if not thu or len(thu) == 7:
        return f"{gio} · hằng ngày"
    return f"{gio} · " + ", ".join(TEN_THU_NGAN[t] for t in sorted(thu))


def mo_ta_moc(moc):
    """Mốc thời gian dạng người đọc được; chuỗi rỗng nếu không có."""
    if not moc:
        return "—"
    if isinstance(moc, str):
        try:
            moc = datetime.fromisoformat(moc)
        except ValueError:
            return moc
    return moc.strftime("%d/%m/%Y %H:%M")


# ==================== TỰ MỞ KHI BẬT MÁY (WINDOWS) ====================
#
# Chỉ có tác dụng với bản .exe trên Windows: đặt một shortcut vào thư mục
# Startup, đăng nhập Windows xong là tool tự mở (thu nhỏ sẵn) và nằm chờ tới
# giờ hẹn. Chạy từ mã nguồn bằng python thì không làm gì — shortcut trỏ tới
# python.exe kèm đường dẫn script rất dễ hỏng khi đổi máy.

TEN_SHORTCUT = "ToolDangBaiFacebook.lnk"
CO_THU_NHO = "--thu-nho"


def ho_tro_tu_mo():
    """Máy này có dùng được tính năng tự mở khi bật máy không."""
    return sys.platform == "win32" and getattr(sys, "frozen", False)


def _duong_dan_shortcut():
    thu_muc = os.path.join(
        os.environ.get("APPDATA", ""),
        "Microsoft", "Windows", "Start Menu", "Programs", "Startup",
    )
    return os.path.join(thu_muc, TEN_SHORTCUT)


def dang_tu_mo():
    return ho_tro_tu_mo() and os.path.exists(_duong_dan_shortcut())


def bat_tu_mo():
    """Tạo shortcut trong thư mục Startup. Ném lỗi kèm lý do nếu không được."""
    if not ho_tro_tu_mo():
        raise RuntimeError(
            "Chỉ dùng được ở bản ứng dụng (.exe) trên Windows.")

    exe = os.path.abspath(sys.executable)
    lnk = _duong_dan_shortcut()
    os.makedirs(os.path.dirname(lnk), exist_ok=True)

    # File .lnk là định dạng nhị phân của Windows, không tự ghi tay được; nhờ
    # PowerShell gọi COM WScript.Shell dựng hộ. Không dùng file .bat thay thế
    # vì .bat luôn kèm một cửa sổ dòng lệnh đen bật lên mỗi lần khởi động máy.
    # Chuỗi trong PowerShell bọc bằng nháy đơn; tên người dùng Windows có dấu
    # nháy đơn (C:\Users\O'Brien\...) sẽ cắt đứt chuỗi làm lệnh hỏng. Escape
    # bằng cách nhân đôi dấu nháy, đúng quy tắc của PowerShell.
    ps = lambda s: s.replace("'", "''")
    lenh = (
        "$s = (New-Object -ComObject WScript.Shell).CreateShortcut('{lnk}'); "
        "$s.TargetPath = '{exe}'; "
        "$s.Arguments = '{co}'; "
        "$s.WorkingDirectory = '{thu_muc}'; "
        "$s.Description = 'Tool dang bai Facebook - tu mo de chay lich hen'; "
        "$s.Save()"
    ).format(lnk=ps(lnk), exe=ps(exe), co=CO_THU_NHO,
             thu_muc=ps(os.path.dirname(exe)))

    ket_qua = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", lenh],
        capture_output=True, text=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if ket_qua.returncode != 0 or not os.path.exists(lnk):
        raise RuntimeError(
            (ket_qua.stderr or "").strip() or "PowerShell không tạo được shortcut.")


def tat_tu_mo():
    lnk = _duong_dan_shortcut()
    if os.path.exists(lnk):
        os.remove(lnk)


def mo_thu_nho():
    """Lần chạy này có phải do Windows tự mở lúc khởi động máy không."""
    return CO_THU_NHO in sys.argv
