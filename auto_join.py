"""
Tự động xin vào nhóm Facebook — chạy RIÊNG, không dính gì tới việc đăng bài.

Đây là việc của cái nick, không phải của chiến dịch: nick vào được nhóm nào là
chuyện lâu dài, còn đăng bài là việc từng ngày. Nên module này có danh sách
nhóm riêng, cài đặt riêng, lịch sử riêng (ghi theo từng nick), và chạy bằng
nút riêng ở tab "Tham gia nhóm".

⚠️ RỦI RO: xin vào nhóm hàng loạt là hành vi Facebook chặn NHANH HƠN đăng bài.
Facebook đếm số lượt join theo nick, theo ngày, và nhìn cả nhịp thao tác. Không
có cách nào làm việc này mà an toàn tuyệt đối; mọi thứ dưới đây chỉ là giảm rủi
ro. Vì thế module tự đặt trần thấp, tự nghỉ dài, tự kiểm chứng sau mỗi lượt
bấm, và khi thấy dấu hiệu bị chặn thì KHOÁ luôn nick lại vài chục tiếng chứ
không để người dùng bấm chạy tiếp.

Dùng chung với post_phong_tro_fb.py đúng phần mở Chrome và nhận biết checkpoint
— thứ đó là hạ tầng trình duyệt, không phải logic đăng bài.
"""
import os
import re
import json
import time
import random
from datetime import datetime, timedelta

import config as cfg_module
import lich_hen
import post_phong_tro_fb as fb


# ======================== CẤU HÌNH ========================
# Cài đặt nằm ở gốc config.json (không nằm trong chiến dịch nào) vì việc vào
# nhóm không thuộc về chiến dịch nào cả.
#
# Mặc định cố ý CHẬM. Con số ở đây không phải để chạy cho nhanh mà để nick sống
# lâu: một nick vào được 5 nhóm mỗi ngày trong hai tuần vẫn hơn hẳn một nick
# vào 40 nhóm trong một buổi rồi bị khoá.
CAI_DAT_MAC_DINH = {
    "links": [],            # danh sách link nhóm muốn vào
    "max_moi_lan": 5,       # trần số nhóm BẤM XIN VÀO trong một lần chạy
    "max_moi_ngay": 8,      # trần mỗi nick mỗi ngày, tính cả các lần chạy trước
    "nghi_min": 90,         # nghỉ giữa 2 nhóm (giây)
    "nghi_max": 300,
    "nghi_dai_sau": 3,      # cứ ngần này lượt xin vào thì nghỉ dài một lần
    "nghi_dai_phut": 20,
    "gio_bat_dau": 8,       # chỉ chạy trong khung giờ này (giờ Việt Nam)
    "gio_ket_thuc": 22,
    "nick_moi": False,      # bật khi nick mới lập: ép trần xuống rất thấp
    "gio_tam_nghi": 24,     # thấy dấu hiệu bị chặn thì khoá nick bấy nhiêu giờ
}

# Trần cho nick mới. Nick vừa lập mà đi xin vào nhóm là kiểu dính nhanh nhất,
# nên bật ô "nick mới" là mọi con số bị ép xuống, không cho tự nâng.
NICK_MOI = {"max_moi_lan": 2, "max_moi_ngay": 3, "nghi_min": 180, "nghi_max": 600}

# Sàn/trần an toàn — người dùng đặt ngoài khoảng này thì tool tự kéo về.
SAN = {
    "max_moi_lan": 1,
    "max_moi_ngay": 1,
    "nghi_min": 40,
    "nghi_max": 80,
    "nghi_dai_sau": 2,
    "nghi_dai_phut": 5,
    "gio_bat_dau": 0,
    "gio_ket_thuc": 1,
    "gio_tam_nghi": 6,
}
TRAN = {
    # Trần này là mức người dùng đã cân nhắc và tự chọn, không phải mức an toàn:
    # càng gần trần thì nhịp thao tác càng dày và càng dễ bị chặn. Nick đang
    # dùng để đăng bài kiếm tiền thì nên chạy quanh 10/12, đừng kịch trần.
    "max_moi_lan": 15,
    "max_moi_ngay": 20,
    "gio_bat_dau": 23,
    "gio_ket_thuc": 24,
    "gio_tam_nghi": 168,
}

TEN_FILE_LOG = "joined_groups.json"

# ---- Trạng thái ghi vào lịch sử ----
DA_LA_THANH_VIEN = "da_la_thanh_vien"
CHO_DUYET = "cho_duyet"
KHONG_VAO_DUOC = "khong_vao_duoc"        # link hỏng / nhóm xóa / hỏi câu hỏi
KHONG_XAC_NHAN = "khong_xac_nhan"        # đã bấm nhưng nút không đổi trạng thái

TEN_TRANG_THAI = {
    DA_LA_THANH_VIEN: "Đã là thành viên",
    CHO_DUYET: "Đang chờ duyệt",
    KHONG_VAO_DUOC: "Không vào được",
    KHONG_XAC_NHAN: "Bấm rồi nhưng chưa rõ",
}

# Nhóm không vào được thì tạm quên đi ngần này ngày rồi mới thử lại. Nếu không,
# mỗi lần chạy tool lại mở đúng mấy link hỏng đó — tốn lượt tải trang vô ích mà
# vẫn bị Facebook tính là hoạt động.
NGAY_THU_LAI = 7

# Giao diện gán lại 2 hook này để hứng log và yêu cầu dừng giữa chừng.
LOG_FN = print
SHOULD_STOP = lambda: False


class StopRequested(Exception):
    """Người dùng bấm Dừng."""


class BiChanError(Exception):
    """Facebook chặn tạm thời thao tác vào nhóm, hoặc chặn ngầm.

    Gặp cái này mà vẫn cố là đi thẳng tới khoá nick, nên tool dừng cả lượt chạy
    VÀ khoá nick lại vài chục tiếng, không để bấm chạy tiếp ngay.
    """


def log(msg):
    LOG_FN(f"[{lich_hen.bay_gio().strftime('%H:%M:%S')}] {msg}")


def check_stop():
    if SHOULD_STOP():
        raise StopRequested()


def nghi(giay, nhan=""):
    """Chờ `giay` giây, mỗi giây ngó xem người dùng có bấm Dừng không."""
    moc = max(giay // 4, 30)
    for con in range(giay, 0, -1):
        check_stop()
        if con == giay or con % moc == 0:
            log(f"  ⏳ {nhan}còn {_mo_ta_thoi_luong(con)}...")
        time.sleep(1)


def _mo_ta_thoi_luong(giay):
    if giay < 60:
        return f"{giay}s"
    return f"{giay // 60} phút {giay % 60}s" if giay % 60 else f"{giay // 60} phút"


# ======================== CÀI ĐẶT ========================

def load_cai_dat():
    """Đọc cài đặt vào nhóm từ config.json, bù mặc định cho key còn thiếu."""
    cfg = cfg_module.load_config()
    cai_dat = dict(CAI_DAT_MAC_DINH)
    cai_dat.update(cfg.get("auto_join") or {})
    return ep_gioi_han(cai_dat)


def save_cai_dat(cai_dat):
    """Ghi cài đặt vào nhóm, có kẹp sàn/trần trước khi ghi. Trả về bản đã ghi."""
    cfg = cfg_module.load_config()
    cfg["auto_join"] = ep_gioi_han(cai_dat)
    cfg_module.save_config(cfg)
    return cfg["auto_join"]


def ep_gioi_han(cai_dat):
    """Kẹp mọi con số về khoảng an toàn. Trả về bản đã sửa."""
    ra = dict(CAI_DAT_MAC_DINH)
    ra.update(cai_dat)

    ra["nick_moi"] = bool(ra.get("nick_moi"))
    for key in ("max_moi_lan", "max_moi_ngay", "nghi_min", "nghi_max",
                "nghi_dai_sau", "nghi_dai_phut", "gio_bat_dau",
                "gio_ket_thuc", "gio_tam_nghi"):
        try:
            ra[key] = int(ra.get(key, CAI_DAT_MAC_DINH[key]))
        except (TypeError, ValueError):
            ra[key] = CAI_DAT_MAC_DINH[key]

    # Nick mới thì siết thêm, và không cho nới ra
    if ra["nick_moi"]:
        ra["max_moi_lan"] = min(ra["max_moi_lan"], NICK_MOI["max_moi_lan"])
        ra["max_moi_ngay"] = min(ra["max_moi_ngay"], NICK_MOI["max_moi_ngay"])
        ra["nghi_min"] = max(ra["nghi_min"], NICK_MOI["nghi_min"])
        ra["nghi_max"] = max(ra["nghi_max"], NICK_MOI["nghi_max"])

    for key, san in SAN.items():
        ra[key] = max(ra[key], san)
    for key, tran in TRAN.items():
        ra[key] = min(ra[key], tran)

    # Nghỉ tối đa phải lớn hơn tối thiểu, nếu không random.randint sẽ nổ
    ra["nghi_max"] = max(ra["nghi_max"], ra["nghi_min"] + 30)
    # Trần ngày phải ≥ trần mỗi lần, không thì con số mỗi lần thành vô nghĩa
    ra["max_moi_ngay"] = max(ra["max_moi_ngay"], ra["max_moi_lan"])
    # Khung giờ phải có ít nhất 1 tiếng
    ra["gio_ket_thuc"] = max(ra["gio_ket_thuc"], ra["gio_bat_dau"] + 1)

    ra["links"] = [str(u).strip() for u in ra.get("links") or [] if str(u).strip()]
    return ra


def trong_gio_hoat_dong(cai_dat, luc=None):
    """Bây giờ (giờ Việt Nam) có nằm trong khung giờ được phép chạy không.

    Xin vào chục nhóm lúc 3h sáng là dấu vết máy móc rõ ràng — người thật ngủ.
    """
    gio = (luc or lich_hen.bay_gio()).hour
    return cai_dat["gio_bat_dau"] <= gio < cai_dat["gio_ket_thuc"]


# ======================== LỊCH SỬ THEO NICK ========================

def clean_group_url(url):
    """Chuẩn hóa link nhóm: bỏ web./m., bỏ query và dấu / cuối.

    Cùng một nhóm ghi 3 kiểu là chuyện thường; không chuẩn hóa thì lịch sử
    không khớp được và tool mở lại nhóm cũ lần nữa.
    """
    url = str(url).strip()
    url = url.split("?")[0].split("#")[0]
    url = url.replace("web.facebook.com", "www.facebook.com")
    url = url.replace("m.facebook.com", "www.facebook.com")
    url = url.replace("//facebook.com", "//www.facebook.com")
    return url.rstrip("/")


def duong_dan_log(profile_dir):
    return os.path.join(profile_dir, TEN_FILE_LOG)


def log_rong():
    return {"nhom": {}, "meta": {}}


def load_log(profile_dir):
    """Lịch sử của một nick: {"nhom": {link: {...}}, "meta": {...}}

    "meta" giữ mốc tạm nghỉ khi phát hiện bị chặn — phải nằm cùng file với lịch
    sử nhóm để copy profile sang máy khác là mốc đó đi theo, không bị mất.
    """
    path = duong_dan_log(profile_dir)
    if not os.path.exists(path):
        return log_rong()
    try:
        with open(path, "r", encoding="utf-8") as f:
            du_lieu = json.load(f)
    except (json.JSONDecodeError, OSError):
        # File hỏng thì coi như chưa vào nhóm nào: cùng lắm tool mở lại nhóm cũ
        # và thấy "đã là thành viên", không có hậu quả gì nặng.
        return log_rong()

    if not isinstance(du_lieu, dict):
        return log_rong()
    if "nhom" not in du_lieu:
        # File kiểu cũ: cả file là bảng nhóm, chưa có phần meta
        return {"nhom": du_lieu, "meta": {}}
    du_lieu.setdefault("meta", {})
    if not isinstance(du_lieu.get("nhom"), dict):
        du_lieu["nhom"] = {}
    return du_lieu


def save_log(profile_dir, lich_su):
    """Ghi lịch sử. Ghi ra file tạm rồi đổi tên để không mất khi tắt ngang."""
    path = duong_dan_log(profile_dir)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(lich_su, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def ghi_nhom(profile_dir, lich_su, url, status, **them):
    """Ghi trạng thái một nhóm và lưu xuống đĩa ngay.

    Lưu ngay từng lượt chứ không gom cuối buổi: tool tắt ngang hoặc Chrome chết
    giữa chừng mà chưa lưu thì mấy lượt vừa bấm coi như chưa xảy ra, lần chạy
    sau bấm lại — đúng kiểu hành vi làm Facebook chú ý.
    """
    muc = {"status": status, "time": lich_hen.bay_gio().strftime("%Y-%m-%d %H:%M:%S")}
    muc.update(them)
    lich_su["nhom"][clean_group_url(url)] = muc
    try:
        save_log(profile_dir, lich_su)
    except OSError as e:
        log(f"  ⚠ Không ghi được lịch sử: {e}")


def muc_nhom(lich_su, url):
    """Bản ghi của một nhóm, hoặc {} nếu chưa đụng tới."""
    muc = lich_su.get("nhom", {}).get(clean_group_url(url))
    return muc if isinstance(muc, dict) else {}


def mo_ta_trang_thai(muc):
    """Chữ hiện trong bảng ở giao diện."""
    if not muc:
        return "Chưa xử lý"
    ten = TEN_TRANG_THAI.get(muc.get("status"), muc.get("status") or "?")
    if muc.get("thu_lai_sau"):
        return f"{ten} (thử lại sau {muc['thu_lai_sau'][:10]})"
    return ten


def _sang_ngay(chuoi):
    """Đọc phần ngày 'YYYY-MM-DD' ở đầu chuỗi, None nếu hỏng."""
    try:
        return datetime.strptime(str(chuoi)[:10], "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return None


def dem_da_xin_hom_nay(lich_su):
    """Số lượt BẤM xin vào nhóm của nick này trong hôm nay (giờ Việt Nam).

    Đếm theo trường "da_bam" chứ không theo trạng thái cuối: một lượt bấm mà
    Facebook lờ đi vẫn là một thao tác Facebook đã đếm, phải tính vào trần.
    """
    hom_nay = lich_hen.bay_gio().date()
    return sum(
        1 for muc in lich_su.get("nhom", {}).values()
        if isinstance(muc, dict) and muc.get("da_bam")
        and _sang_ngay(muc.get("da_bam")) == hom_nay
    )


def dat_tam_nghi(profile_dir, lich_su, so_gio, ly_do):
    """Khoá nick không cho chạy tiếp trong `so_gio` tiếng."""
    den = lich_hen.bay_gio() + timedelta(hours=so_gio)
    lich_su.setdefault("meta", {})["tam_nghi_den"] = den.strftime("%Y-%m-%d %H:%M:%S")
    lich_su["meta"]["ly_do_tam_nghi"] = ly_do
    try:
        save_log(profile_dir, lich_su)
    except OSError:
        pass
    return den


def dang_tam_nghi(lich_su):
    """(True, mốc, lý do) nếu nick đang bị khoá tạm; (False, None, "") nếu không."""
    meta = lich_su.get("meta") or {}
    moc = meta.get("tam_nghi_den")
    if not moc:
        return False, None, ""
    try:
        den = datetime.strptime(str(moc), "%Y-%m-%d %H:%M:%S")
        den = den.replace(tzinfo=lich_hen.bay_gio().tzinfo)
    except (ValueError, TypeError):
        return False, None, ""
    if lich_hen.bay_gio() >= den:
        return False, None, ""
    return True, den, meta.get("ly_do_tam_nghi", "")


def bo_tam_nghi(profile_dir, lich_su):
    """Gỡ khoá tạm nghỉ (người dùng tự chịu trách nhiệm)."""
    meta = lich_su.setdefault("meta", {})
    meta.pop("tam_nghi_den", None)
    meta.pop("ly_do_tam_nghi", None)
    try:
        save_log(profile_dir, lich_su)
    except OSError:
        pass


# Bản kiểm tra trạng thái nhóm. Tăng số này khi cách nhận biết "đã ở trong
# nhóm" thay đổi, để lịch sử ghi bằng cách nhận biết cũ bị soát lại.
#
# v1 lấy ô soạn bài "Bạn viết gì đi..." làm dấu hiệu đã tham gia. Nhóm công
# khai hiện ô đó cho cả người ngoài, nên v1 ghi khống "đã là thành viên" cho
# hàng loạt nhóm chưa hề vào — và vì đã có trong lịch sử nên tool không bao giờ
# mở lại để sửa. Phải xoá những bản ghi đó đi.
PHIEN_BAN_KIEM_TRA = 2


def don_log_ghi_khong(profile_dir, lich_su, ep_buoc=False):
    """Xoá các bản ghi 'đã vào nhóm' mà tool chưa từng bấm nút Tham gia.

    Chỉ xoá bản ghi KHÔNG có dấu "da_bam" — tức là tool tự kết luận nick đã ở
    trong nhóm chứ không phải nó bấm rồi thấy thành công. Đó đúng là những bản
    ghi mà cách nhận biết cũ có thể kết luận sai. Xoá đi thì nhóm quay về "chưa
    xử lý" và lần chạy tới tool mở lại kiểm tra tử tế.

    Trả về số bản ghi đã xoá.
    """
    meta = lich_su.setdefault("meta", {})
    if not ep_buoc and int(meta.get("phien_ban_kiem_tra", 1)) >= PHIEN_BAN_KIEM_TRA:
        return 0

    bo = [u for u, m in lich_su.get("nhom", {}).items()
          if isinstance(m, dict)
          and m.get("status") in (DA_LA_THANH_VIEN, CHO_DUYET)
          and not m.get("da_bam")]
    for u in bo:
        del lich_su["nhom"][u]

    meta["phien_ban_kiem_tra"] = PHIEN_BAN_KIEM_TRA
    try:
        save_log(profile_dir, lich_su)
    except OSError as e:
        log(f"⚠ Không ghi được lịch sử sau khi dọn: {e}")
    return len(bo)


def loc_nhom_can_vao(links, lich_su):
    """Bỏ link trùng, bỏ nhóm đã xong, giữ lại nhóm tới hạn thử lại."""
    hom_nay = lich_hen.bay_gio().date()
    ra = []
    da_gap = set()
    for u in links:
        cu = clean_group_url(u)
        if not cu or cu in da_gap:
            continue
        da_gap.add(cu)

        muc = muc_nhom(lich_su, cu)
        if not muc:
            ra.append(cu)
            continue
        # Đã vào hoặc đang chờ duyệt thì thôi
        if muc.get("status") in (DA_LA_THANH_VIEN, CHO_DUYET):
            continue
        # Nhóm hỏng / bấm rồi mà không rõ: chỉ thử lại khi tới hạn
        han = _sang_ngay(muc.get("thu_lai_sau"))
        if han is None or han <= hom_nay:
            ra.append(cu)
    return ra


def tinh_quota(cai_dat, lich_su):
    """Số lượt xin vào nhóm còn được phép trong lần chạy này."""
    con_lai_ngay = max(cai_dat["max_moi_ngay"] - dem_da_xin_hom_nay(lich_su), 0)
    return min(cai_dat["max_moi_lan"], con_lai_ngay)


# ======================== NHẬN BIẾT BỊ CHẶN ========================

# Facebook báo bằng nhiều câu khác nhau tùy phiên bản giao diện; đây là những
# câu hay gặp nhất khi thao tác join bị chặn.
CAU_BI_CHAN = (
    "Tạm thời bị chặn",
    "tạm thời bị chặn",
    "Bạn tạm thời bị chặn",
    "Tính năng này hiện không khả dụng",
    "Bạn đang bị tạm khóa tính năng",
    "You’re Temporarily Blocked",
    "You're Temporarily Blocked",
    "Temporarily Blocked",
    "You can't use this feature right now",
    "This feature is temporarily blocked",
)


def kiem_tra_bi_chan(page):
    """Ném BiChanError nếu Facebook đang chặn tạm thời thao tác này."""
    for cau in CAU_BI_CHAN:
        try:
            o = page.get_by_text(cau, exact=False).first
            if o.count() > 0 and o.is_visible(timeout=800):
                raise BiChanError(f"Facebook báo: \"{cau}\"")
        except BiChanError:
            raise
        except Exception:
            continue        # không thấy câu này thì thử câu kế


# ======================== THAO TÁC TRÊN MỘT NHÓM ========================

# Tìm nút theo TÊN ĐẦY ĐỦ của nút (accessible name), khớp trọn chuỗi chứ không
# phải "có chứa": kiểu tìm chuỗi con sẽ khớp cả cái khung to bao ngoài chứa chữ
# đó, và khớp luôn mấy nút "Tham gia nhóm" của phần nhóm gợi ý bên cạnh.
# \W* hai đầu để bỏ qua dấu cộng, dấu tích, khoảng trắng mà Facebook hay ghép
# vào nhãn nút ("+ Tham gia nhóm", "✓ Đã tham gia"). Chữ tiếng Việt là ký tự
# chữ nên "Tham gia nhóm này để xem thêm" vẫn không khớp.
# Nút chính ở đầu trang nhóm. Tách riêng khỏi dạng ngắn vì mấy thẻ "Nhóm gợi ý"
# ở cột bên phải cũng có nút, nhưng nhãn của chúng là "Tham gia" trống không.
NUT_THAM_GIA = re.compile(r"^\W*(tham gia nhóm|join group)\W*$", re.I)
NUT_THAM_GIA_NGAN = re.compile(r"^\W*(tham gia|join)\W*$", re.I)
NUT_CHO_DUYET = re.compile(r"^\W*(đã yêu cầu|hủy yêu cầu|huỷ yêu cầu|requested|cancel request)\W*$", re.I)
NUT_DA_THAM_GIA = re.compile(r"^\W*(đã tham gia|đã là thành viên|joined|rời nhóm|leave group)\W*$", re.I)

# Ô soạn bài. CHỈ được dùng làm bằng chứng phụ, khi trên trang đã chắc chắn
# không còn nút "Tham gia nhóm" — xem _da_la_thanh_vien().
O_SOAN_BAI = re.compile(r"^\W*(bạn viết gì đi.*|viết gì đó.*|write something.*)$", re.I)


def _tim_nut(page, ten, timeout=2500):
    """Nút hiển thị đầu tiên có tên khớp `ten`, hoặc None."""
    try:
        o = page.get_by_role("button", name=ten).first
        if o.count() > 0 and o.is_visible(timeout=timeout):
            return o
    except Exception:
        pass
    return None


def _nut_cao_nhat(page, ten):
    """Trong các nút khớp `ten`, lấy nút nằm CAO NHẤT trên trang.

    Nút tham gia của chính nhóm này nằm ở đầu trang; mấy nút cùng nhãn ở cột
    "Nhóm gợi ý" bên phải nằm thấp hơn nhiều. Không chọn theo vị trí thì có lúc
    tool bấm nhầm vào nhóm gợi ý — vào một nhóm hoàn toàn khác.
    """
    try:
        cac_nut = page.get_by_role("button", name=ten).all()
    except Exception:
        return None

    tot_nhat, y_nho_nhat = None, None
    for nut in cac_nut[:12]:
        try:
            if not nut.is_visible(timeout=800):
                continue
            hop = nut.bounding_box()
        except Exception:
            continue
        if not hop:
            continue
        if y_nho_nhat is None or hop["y"] < y_nho_nhat:
            tot_nhat, y_nho_nhat = nut, hop["y"]
    return tot_nhat


def _nut_tham_gia(page):
    """Nút tham gia của CHÍNH nhóm đang mở.

    Ưu tiên nhãn đầy đủ "Tham gia nhóm"; chỉ khi không có mới chấp nhận nhãn
    ngắn "Tham gia", và khi đó lấy nút cao nhất trang để khỏi vớ phải nút của
    nhóm gợi ý bên cạnh.
    """
    return _nut_cao_nhat(page, NUT_THAM_GIA) or _nut_cao_nhat(page, NUT_THAM_GIA_NGAN)


def _dang_cho_duyet(page):
    return _tim_nut(page, NUT_CHO_DUYET) is not None


def _da_la_thanh_vien(page):
    """Nick có thật sự ở trong nhóm không.

    KHÔNG dựa vào ô soạn bài "Bạn viết gì đi...": nhóm công khai hiện ô đó cho
    cả người ngoài xem, nên lấy nó làm dấu hiệu là tool sẽ ghi "đã là thành
    viên" cho hàng loạt nhóm chưa hề tham gia. Dấu hiệu đáng tin là nút trạng
    thái "Đã tham gia"/"Rời nhóm", và điều kiện bắt buộc là KHÔNG còn nút
    "Tham gia nhóm" trên trang.
    """
    if _nut_tham_gia(page) is not None:
        return False
    if _tim_nut(page, NUT_DA_THAM_GIA) is not None:
        return True
    # Không còn nút tham gia mà vẫn viết bài được: gần như chắc chắn là thành
    # viên. Chỉ xét tới đây sau khi đã loại trường hợp nhóm công khai — nhóm đó
    # luôn kèm nút "Tham gia nhóm" bên cạnh ô soạn bài.
    return _tim_nut(page, O_SOAN_BAI) is not None


def _xem_trang_nhu_nguoi_that(page):
    """Xem trang một lúc rồi cuộn vài nhịp trước khi bấm.

    Người thật mở nhóm ra là đọc mô tả, kéo xem vài bài rồi mới bấm tham gia.
    Vào thẳng rồi bấm ngay trong một giây là dấu hiệu máy móc rõ nhất.
    """
    page.wait_for_timeout(random.randint(3000, 7000))
    for _ in range(random.randint(2, 4)):
        try:
            page.mouse.move(random.randint(200, 1000), random.randint(200, 700))
        except Exception:
            pass
        page.mouse.wheel(0, random.randint(300, 900))
        page.wait_for_timeout(random.randint(900, 2600))
    page.mouse.wheel(0, -random.randint(200, 600))
    page.wait_for_timeout(random.randint(1000, 2500))


def _ngo_bang_tin(page):
    """Thỉnh thoảng ghé bảng tin giữa hai nhóm.

    Một phiên chỉ toàn mở nhóm rồi bấm tham gia, không xem gì khác, là dấu vết
    rất dễ nhận. Ghé bảng tin cuộn vài nhịp cho giống người dùng thật.
    """
    try:
        log("  🙂 Ghé bảng tin một lát cho tự nhiên")
        page.goto("https://www.facebook.com/", wait_until="domcontentloaded", timeout=45000)
        page.wait_for_timeout(random.randint(2000, 4000))
        for _ in range(random.randint(2, 5)):
            page.mouse.wheel(0, random.randint(400, 1200))
            page.wait_for_timeout(random.randint(1200, 3000))
    except Exception:
        pass        # ghé chơi thôi, hỏng cũng không sao


# Chờ tối đa ngần này giây cho nút đổi trạng thái sau khi bấm. Facebook có lúc
# vài giây mới đổi, có lúc phải tải lại cả khối đầu trang — chờ cứng 3–6s là
# kết luận "bấm không ăn thua" trong khi thật ra đã vào nhóm.
GIAY_CHO_NUT_DOI = 20


def _nut_con_nguyen(nut_da_bam):
    """Cái nút vừa bấm còn nằm đó và vẫn là nút "Tham gia" hay không.

    Bám đúng cái nút đã bấm thay vì quét lại cả trang: sau khi vào nhóm thành
    công, cột "Nhóm gợi ý" bên phải VẪN còn nút "Tham gia" của nhóm khác — quét
    cả trang là lại thấy nút tham gia và kết luận nhầm rằng cú bấm hỏng.
    """
    try:
        if not nut_da_bam.is_visible(timeout=800):
            return False
        chu = (nut_da_bam.text_content(timeout=800) or "").strip()
    except Exception:
        return False        # nút biến mất khỏi trang = đã đổi
    return bool(NUT_THAM_GIA.match(chu) or NUT_THAM_GIA_NGAN.match(chu))


def _cho_nut_doi_trang_thai(page, nut_da_bam):
    """Chờ tới GIAY_CHO_NUT_DOI giây xem cú bấm có ăn thua không.

    Trả về True ngay khi thấy nút đã đổi, hoặc thấy dấu hiệu đã vào / đang chờ
    duyệt. Không đợi hết thời gian nếu đã rõ kết quả.
    """
    for _ in range(GIAY_CHO_NUT_DOI):
        check_stop()
        if not _nut_con_nguyen(nut_da_bam):
            return True
        if _dang_cho_duyet(page) or _tim_nut(page, NUT_DA_THAM_GIA, timeout=500) is not None:
            return True
        time.sleep(1)
    return False


def _co_form_cau_hoi(page):
    """Nhóm bắt trả lời câu hỏi trước khi vào."""
    try:
        o = page.locator(
            "[role='dialog']:has-text('Câu hỏi'), "
            "[role='dialog']:has-text('câu hỏi'), "
            "[role='dialog']:has-text('Answer'), "
            "[role='dialog']:has-text('Questions')"
        ).first
        return o.count() > 0 and o.is_visible(timeout=2000)
    except Exception:
        return False


def vao_mot_nhom(page, url, lich_su, profile_dir):
    """Mở một nhóm và xin vào nếu chưa là thành viên.

    Trả về (trạng thái, đã_bấm). `đã_bấm` = có thực sự bấm nút Tham gia hay
    không — dùng để tính trần, vì Facebook đếm thao tác chứ không đếm kết quả.
    """
    url = clean_group_url(url)
    log(f"  🔗 Mở nhóm: {url}")

    page.goto(url, wait_until="domcontentloaded", timeout=45000)
    page.wait_for_timeout(random.randint(2500, 5000))
    fb.kiem_tra_checkpoint(page)
    kiem_tra_bi_chan(page)

    # THỨ TỰ Ở ĐÂY QUAN TRỌNG. Nút "Tham gia nhóm" còn trên trang là bằng chứng
    # chắc chắn nhất rằng nick CHƯA ở trong nhóm, nên phải xét trước mọi dấu
    # hiệu khác. Xét "đã là thành viên" trước là cách tool từng ghi khống hàng
    # loạt nhóm chưa hề tham gia, chỉ vì nhóm công khai vẫn hiện ô soạn bài cho
    # người ngoài xem.
    nut = _nut_tham_gia(page)

    if nut is None:
        if _dang_cho_duyet(page):
            log("  ✓ Đã gửi yêu cầu từ trước, đang chờ duyệt")
            ghi_nhom(profile_dir, lich_su, url, CHO_DUYET)
            return CHO_DUYET, False

        if _da_la_thanh_vien(page):
            log("  ✓ Đã là thành viên nhóm này")
            ghi_nhom(profile_dir, lich_su, url, DA_LA_THANH_VIEN)
            return DA_LA_THANH_VIEN, False

        # Không thấy nút nào: link hỏng, nhóm đã xóa, nhóm kín không cho xin
        # vào từ đây, hoặc trang chưa tải xong. Không đoán bừa là đã vào.
        log("  ⚠ Không thấy nút tham gia và cũng không rõ có ở trong nhóm không "
            f"(link hỏng, nhóm đã xóa, hoặc nhóm kín) → bỏ qua, "
            f"{NGAY_THU_LAI} ngày nữa mới thử lại")
        ghi_nhom(profile_dir, lich_su, url, KHONG_VAO_DUOC,
                 thu_lai_sau=(lich_hen.bay_gio() + timedelta(days=NGAY_THU_LAI))
                 .strftime("%Y-%m-%d"))
        return KHONG_VAO_DUOC, False

    _xem_trang_nhu_nguoi_that(page)
    check_stop()

    log("  👉 Bấm 'Tham gia nhóm'")
    luc_bam = lich_hen.bay_gio().strftime("%Y-%m-%d %H:%M:%S")
    nut.click(timeout=10000)

    # Ghi dấu ĐÃ BẤM ngay lập tức, trước cả khi biết kết quả: nếu Chrome chết
    # hoặc người dùng tắt tool ngay sau cú bấm này mà chưa ghi, lần chạy sau
    # tool sẽ bấm lại nhóm đó và trần trong ngày cũng đếm thiếu.
    ghi_nhom(profile_dir, lich_su, url, KHONG_XAC_NHAN, da_bam=luc_bam,
             thu_lai_sau=(lich_hen.bay_gio() + timedelta(days=1)).strftime("%Y-%m-%d"))

    page.wait_for_timeout(random.randint(2000, 3500))
    fb.kiem_tra_checkpoint(page)
    kiem_tra_bi_chan(page)

    # Nhóm bật form câu hỏi — tool KHÔNG tự trả lời hộ: trả lời bừa là bị quản
    # trị viên từ chối vĩnh viễn, mà từ chối rồi thì vào tay cũng khó.
    if _co_form_cau_hoi(page):
        log("  ⚠ Nhóm này bắt trả lời câu hỏi — hãy tự trả lời trong Chrome. "
            "Tool bỏ qua nhóm này.")
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
        ghi_nhom(profile_dir, lich_su, url, KHONG_VAO_DUOC, da_bam=luc_bam,
                 thu_lai_sau=(lich_hen.bay_gio() + timedelta(days=NGAY_THU_LAI))
                 .strftime("%Y-%m-%d"))
        return KHONG_VAO_DUOC, True

    # Kiểm chứng: chờ tới 20s xem CHÍNH cái nút vừa bấm có đổi không.
    if not _cho_nut_doi_trang_thai(page, nut):
        log(f"  ⚠ Chờ {GIAY_CHO_NUT_DOI}s mà nút 'Tham gia nhóm' vẫn còn nguyên "
            "— cú bấm không có tác dụng. KHÔNG ghi là đã tham gia.")
        return KHONG_XAC_NHAN, True

    kiem_tra_bi_chan(page)

    if _dang_cho_duyet(page):
        log("  ✓ Đã gửi yêu cầu, đang chờ quản trị viên duyệt")
        ghi_nhom(profile_dir, lich_su, url, CHO_DUYET, da_bam=luc_bam)
        return CHO_DUYET, True

    if _da_la_thanh_vien(page):
        log("  ✓ Vào nhóm thành công (nhóm không cần duyệt)")
        ghi_nhom(profile_dir, lich_su, url, DA_LA_THANH_VIEN, da_bam=luc_bam)
        return DA_LA_THANH_VIEN, True

    # Nút tham gia biến mất nhưng không thấy dấu hiệu nào khẳng định đã vào.
    # Thà ghi "chưa rõ" để mai kiểm tra lại, còn hơn ghi khống là đã tham gia.
    log("  ⚠ Nút tham gia biến mất nhưng chưa thấy dấu hiệu đã vào nhóm — ghi "
        "'chưa rõ', mai kiểm tra lại.")
    return KHONG_XAC_NHAN, True


# ======================== CHẠY CẢ DANH SÁCH ========================

def chay_join(profile_dir, links=None, cai_dat=None, confirm_nick=None):
    """Mở Chrome bằng nick đã chọn và lần lượt xin vào các nhóm.

    `links` None = lấy danh sách trong cài đặt. Trả về dict thống kê, hoặc None
    nếu không chạy được (đang tạm nghỉ, ngoài giờ, chưa đăng nhập, người dùng
    hủy).
    """
    cai_dat = ep_gioi_han(cai_dat or load_cai_dat())
    links = links if links is not None else cai_dat["links"]

    lich_su = load_log(profile_dir)

    # Dọn bản ghi do bản cũ kết luận sai, trước khi lọc danh sách
    da_don = don_log_ghi_khong(profile_dir, lich_su)
    if da_don:
        log(f"⚠ Đã bỏ {da_don} bản ghi 'đã vào nhóm' do bản cũ kết luận sai "
            "(nhóm công khai hiện ô soạn bài cho cả người ngoài). Mấy nhóm đó "
            "quay về 'chưa xử lý' và sẽ được kiểm tra lại.")

    # --- Ba cửa chặn trước khi mở Chrome ---
    dang_nghi, den, ly_do = dang_tam_nghi(lich_su)
    if dang_nghi:
        log(f"⛔ Nick này đang tạm nghỉ tới {den.strftime('%d/%m %H:%M')} "
            f"({ly_do}).")
        log("Chạy tiếp lúc này là cách nhanh nhất để mất nick. Hãy chờ hết hạn.")
        return None

    if not trong_gio_hoat_dong(cai_dat):
        log(f"⛔ Ngoài khung giờ cho phép ({cai_dat['gio_bat_dau']}h–"
            f"{cai_dat['gio_ket_thuc']}h giờ Việt Nam).")
        log("Xin vào nhóm lúc nửa đêm là dấu vết máy móc rõ ràng — người thật ngủ.")
        return None

    can_vao = loc_nhom_can_vao(links, lich_su)
    da_hom_nay = dem_da_xin_hom_nay(lich_su)
    quota = tinh_quota(cai_dat, lich_su)

    log("=" * 60)
    log("THAM GIA NHÓM TỰ ĐỘNG")
    log(f"Nhóm trong danh sách : {len(links)}")
    log(f"Cần xử lý            : {len(can_vao)}")
    log(f"Hôm nay đã xin vào   : {da_hom_nay}/{cai_dat['max_moi_ngay']} nhóm")
    log(f"Lượt chạy này tối đa : {quota} nhóm")
    if cai_dat["nick_moi"]:
        log("Đang bật chế độ NICK MỚI — trần thấp và nghỉ lâu hơn.")
    log("=" * 60)

    if not can_vao:
        log("Không có nhóm nào cần xử lý.")
        return {"vao": 0, "cho_duyet": 0, "da_co": 0, "loi": 0, "con_lai": 0}
    if quota <= 0:
        log(f"⚠ Hôm nay nick này đã xin vào {da_hom_nay} nhóm, chạm trần "
            f"{cai_dat['max_moi_ngay']} nhóm/ngày → dừng, mai chạy tiếp.")
        return {"vao": 0, "cho_duyet": 0, "da_co": 0, "loi": 0, "con_lai": len(can_vao)}

    # Xáo thứ tự: chạy đúng thứ tự khai ngày nào cũng như ngày nào là dấu vết
    # máy móc, mà danh sách nhóm thường xếp theo khu vực nên càng lộ.
    can_vao = list(can_vao)
    random.shuffle(can_vao)

    vao = cho_duyet = da_co = loi = 0
    da_bam_tong = 0
    loi_lien_tiep = 0
    khong_xac_nhan_lien_tiep = 0
    ket_thuc_som = ""

    with fb.sync_playwright() as p:
        context = fb.open_profile(p, profile_dir)
        page = fb.get_page(context)

        try:
            log("Đang kiểm tra tài khoản Facebook...")
            page.goto("https://www.facebook.com/me", wait_until="domcontentloaded", timeout=45000)
            page.wait_for_timeout(3000)
            fb.kiem_tra_checkpoint(page)

            if not fb.is_logged_in(page):
                log("⚠ Nick này chưa đăng nhập Facebook. Hãy đăng nhập rồi chạy lại.")
                return None

            ten_nick = fb.get_profile_name(page)
            log(f"Đang dùng nick: {ten_nick}")
            if confirm_nick and not confirm_nick(ten_nick):
                log("Đã hủy — không đúng nick.")
                return None

            for i, url in enumerate(can_vao):
                check_stop()
                if da_bam_tong >= quota:
                    ket_thuc_som = f"đủ {quota} lượt xin vào nhóm cho lần chạy này"
                    break
                # Khung giờ có thể trôi qua giữa chừng nếu chạy sát giờ cuối
                if not trong_gio_hoat_dong(cai_dat):
                    ket_thuc_som = "hết khung giờ cho phép"
                    break

                log(f"── Nhóm {i + 1}/{len(can_vao)} (đã bấm {da_bam_tong}/{quota})")
                try:
                    kq, da_bam = vao_mot_nhom(page, url, lich_su, profile_dir)
                except (fb.FacebookCheckpointError, BiChanError, StopRequested):
                    raise
                except Exception as e:
                    log(f"  ✗ Lỗi ở nhóm này: {str(e).splitlines()[0][:180]}")
                    kq, da_bam = "loi", False

                if da_bam:
                    da_bam_tong += 1

                # Đếm theo việc lần chạy NÀY có bấm hay không, để người dùng
                # phân biệt "tool vừa xin vào" với "vốn đã ở trong nhóm rồi".
                if kq in (CHO_DUYET, DA_LA_THANH_VIEN):
                    if da_bam:
                        vao += 1
                    elif kq == CHO_DUYET:
                        cho_duyet += 1
                    else:
                        da_co += 1
                else:
                    loi += 1

                # --- Đếm dấu hiệu nguy hiểm ---
                if kq == KHONG_XAC_NHAN:
                    khong_xac_nhan_lien_tiep += 1
                    # Bấm mà nút không đổi, lặp lại 2 lần: gần như chắc chắn
                    # Facebook đang chặn ngầm thao tác join của nick này.
                    if khong_xac_nhan_lien_tiep >= 2:
                        raise BiChanError(
                            "2 lượt bấm liên tiếp không có tác dụng — Facebook "
                            "nhiều khả năng đang chặn ngầm thao tác vào nhóm")
                else:
                    khong_xac_nhan_lien_tiep = 0

                if kq == "loi":
                    loi_lien_tiep += 1
                    if loi_lien_tiep >= 3:
                        ket_thuc_som = ("3 nhóm liền không xử lý được — dừng cho "
                                        "an toàn, hãy mở Chrome kiểm tra bằng tay")
                        break
                else:
                    loi_lien_tiep = 0

                if da_bam_tong >= quota or i == len(can_vao) - 1:
                    continue

                # --- Nghỉ ---
                if da_bam_tong and da_bam_tong % cai_dat["nghi_dai_sau"] == 0 and da_bam:
                    log(f"  ☕ Đã bấm {da_bam_tong} nhóm — nghỉ dài "
                        f"{cai_dat['nghi_dai_phut']} phút")
                    nghi(cai_dat["nghi_dai_phut"] * 60, "Nghỉ dài: ")
                else:
                    if da_bam and random.random() < 0.35:
                        _ngo_bang_tin(page)
                    nghi(random.randint(cai_dat["nghi_min"], cai_dat["nghi_max"]),
                         "Nhóm tiếp: ")

        except StopRequested:
            log("⏹ Đã dừng theo yêu cầu.")
        except fb.FacebookCheckpointError as e:
            log(f"⛔ {e}")
            den = dat_tam_nghi(profile_dir, lich_su, cai_dat["gio_tam_nghi"],
                               "Facebook đòi xác minh")
            log(f"Đã khoá việc tham gia nhóm của nick này tới "
                f"{den.strftime('%d/%m %H:%M')}. Hãy mở Chrome xác minh trước.")
        except BiChanError as e:
            log(f"⛔ {e}")
            den = dat_tam_nghi(profile_dir, lich_su, cai_dat["gio_tam_nghi"],
                               "bị chặn thao tác vào nhóm")
            log(f"Đã khoá việc tham gia nhóm của nick này tới "
                f"{den.strftime('%d/%m %H:%M')} — chạy lại ngay là mất nick.")
        finally:
            context.close()

    if ket_thuc_som:
        log(f"■ Dừng sớm: {ket_thuc_som}.")

    con_lai = len(loc_nhom_can_vao(links, lich_su))
    log("=" * 60)
    log(f"KẾT QUẢ: {vao} nhóm vừa xin vào | {cho_duyet} đang chờ duyệt từ trước | "
        f"{da_co} đã là thành viên | {loi} không xử lý được")
    if con_lai:
        log(f"Còn {con_lai} nhóm chưa xong — để hôm khác, đừng chạy dồn trong ngày.")
    log(f"Hôm nay nick này đã bấm: {dem_da_xin_hom_nay(lich_su)}/{cai_dat['max_moi_ngay']} nhóm.")
    log("=" * 60)

    return {"vao": vao, "cho_duyet": cho_duyet, "da_co": da_co,
            "loi": loi, "con_lai": con_lai}


# ======================== CHẠY BẰNG TERMINAL ========================

def main():
    """python auto_join.py — chọn nick rồi xin vào các nhóm đã khai."""
    cai_dat = load_cai_dat()
    if not cai_dat["links"]:
        print("Chưa khai nhóm nào. Mở giao diện (python gui.py), tab "
              "'👥 Tham gia nhóm' để dán danh sách link.")
        return

    profile_dir = fb.choose_account()

    def confirm_nick(ten):
        return input(f"\n>>> Đúng nick '{ten}'? Enter để chạy, 'q' để hủy: ").strip().lower() != "q"

    chay_join(profile_dir, confirm_nick=confirm_nick)


if __name__ == "__main__":
    main()
