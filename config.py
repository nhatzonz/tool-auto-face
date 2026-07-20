"""
Quản lý cấu hình dùng chung cho script CLI và giao diện GUI.

Cấu hình nằm trong config.json cạnh script, gồm 2 phần:
  - Cài đặt chung: thư mục profile Chrome (nick dùng chung cho mọi chiến dịch)
  - campaigns: mỗi chiến dịch (phòng trọ, tuyển dụng, seeding web...) có file
    Excel, thư mục ảnh, file log và bản đồ nhóm riêng.

Mọi chiến dịch dùng chung cấu trúc Excel 4 cột:
  A: STT | B: Nội dung bài | C: Phân loại | D: Tên thư mục ảnh
Cột "Phân loại" quyết định bài được đăng lên nhóm nào (với phòng trọ đó là khu
vực, với tuyển dụng có thể là ngành nghề, với seeding là chủ đề).
"""
import os
import json

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

# Giá trị mặc định cho một chiến dịch mới tạo
CAMPAIGN_DEFAULTS = {
    "excel_path": "",
    "sheet_name": "Sheet1",
    "images_dir": "",
    "posted_log": "",
    "max_posts": 0,
    "delay_between_groups": 10,
    "delay_between_rooms": 10,
    "max_retries": 2,
    "groups": {},
}

DEFAULT_CONFIG = {
    "profiles_dir": os.path.join(BASE_DIR, "fb_profiles"),
    "active_campaign": "Phòng trọ",
    "campaigns": {
        "Phòng trọ": {
            "excel_path": os.path.join(BASE_DIR, "phong_tro.xlsx"),
            "sheet_name": "PhongTro",
            "images_dir": os.path.join(BASE_DIR, "anh_phong"),
            "posted_log": os.path.join(BASE_DIR, "posted_log.json"),
            "max_posts": 0,
            "delay_between_groups": 10,
            "delay_between_rooms": 10,
            "max_retries": 2,
            "groups": {
                "Hà Đông": [
                    "https://web.facebook.com/groups/phongtro.hanoi.hadong/",
                    "https://web.facebook.com/groups/835892593690478/",
                    "https://web.facebook.com/groups/3555475404499952/",
                    "https://web.facebook.com/groups/1041520932684656/",
                    "https://web.facebook.com/groups/1589501227985413/",
                    "https://web.facebook.com/groups/631650078775924",
                ],
                "Thanh Xuân": [
                    "https://web.facebook.com/groups/nhatrodongdathanhxuan/",
                    "https://web.facebook.com/groups/176362986942358/",
                    "https://web.facebook.com/groups/605109991280427/",
                    "https://web.facebook.com/groups/908726406847516/",
                    "https://web.facebook.com/groups/timphongtrodongdangatusothanhxuanhanoi/",
                    "https://web.facebook.com/groups/1385595868491454/",
                ],
                "Bắc từ liêm": [
                    "https://web.facebook.com/groups/2012063565703273/",
                    "https://web.facebook.com/groups/1069950844149062/",
                ],
                "Cầu Giấy": [
                    "https://web.facebook.com/groups/142775226671894/",
                    "https://web.facebook.com/groups/702443907550431/",
                    "https://web.facebook.com/groups/370974904259405/",
                    "https://web.facebook.com/groups/1041097177406107/",
                    "https://web.facebook.com/groups/6104634336285691/",
                    "https://web.facebook.com/groups/nhatrometrimydinhcaugiay/",
                    "https://web.facebook.com/groups/phongtrocaugiayhn/",
                    "https://web.facebook.com/groups/140397885361011/",
                    "https://web.facebook.com/groups/2237019069763450/",
                    "https://web.facebook.com/groups/2202922693407349/",
                ],
                "Mỹ Đình": [
                    "https://web.facebook.com/groups/1914388365626022/",
                    "https://web.facebook.com/groups/2237019069763450/",
                    "https://web.facebook.com/groups/phongtrocaugiaymydinhmetri/",
                    "https://web.facebook.com/groups/507104870413526/",
                    "https://web.facebook.com/groups/2148539488498466/",
                    "https://web.facebook.com/groups/1542856739594335/",
                ],
                "Ba Đình": [
                    "https://web.facebook.com/groups/phongtrobadinh.giatot/",
                    "https://web.facebook.com/groups/757259302549445/",
                ],
                "Hai Bà Trưng": [
                    "https://web.facebook.com/groups/1747492728936509/",
                    "https://web.facebook.com/groups/494231151747853/",
                    "https://web.facebook.com/groups/724565062266526/",
                    "https://web.facebook.com/groups/647543593374506/",
                    "https://web.facebook.com/groups/2790717834511802/",
                ],
                "Hoàng Mai": [
                    "https://web.facebook.com/groups/649420490421652/",
                    "https://web.facebook.com/groups/900402601097250/",
                    "https://web.facebook.com/groups/583641796650849/",
                    "https://web.facebook.com/groups/778767189186540/",
                ],
                "Thanh Trì": [
                    "https://web.facebook.com/groups/1145700312817923/",
                    "https://web.facebook.com/groups/964760639021647/",
                    "https://web.facebook.com/groups/TimPhongTroThanhTri/",
                    "https://web.facebook.com/groups/1620926588466366/",
                ],
            },
        }
    },
}


def _migrate_flat_config(cfg):
    """Chuyển config kiểu cũ (1 chiến dịch, các key nằm phẳng ở gốc) sang
    kiểu campaigns. Giữ nguyên dữ liệu người dùng đã cấu hình."""
    campaign = dict(CAMPAIGN_DEFAULTS)
    campaign.update({
        "excel_path": cfg.get("excel_path", ""),
        "sheet_name": "PhongTro",
        "images_dir": cfg.get("images_dir", ""),
        "posted_log": cfg.get("posted_log", ""),
        "max_posts": cfg.get("max_posts", 0),
        "delay_between_groups": cfg.get("delay_between_groups", 10),
        "delay_between_rooms": cfg.get("delay_between_rooms", 10),
        "max_retries": cfg.get("max_retries", 2),
        "groups": cfg.get("khu_vuc_groups", {}),
    })
    return {
        "profiles_dir": cfg.get("profiles_dir", DEFAULT_CONFIG["profiles_dir"]),
        "active_campaign": "Phòng trọ",
        "campaigns": {"Phòng trọ": campaign},
    }


def load_config():
    """Đọc config.json, tự tạo/nâng cấp nếu cần. Luôn trả về cấu trúc campaigns."""
    if not os.path.exists(CONFIG_PATH):
        save_config(DEFAULT_CONFIG)
        return json.loads(json.dumps(DEFAULT_CONFIG))

    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    # Config kiểu cũ: có khu_vuc_groups ở gốc, chưa có campaigns
    if "campaigns" not in cfg:
        cfg = _migrate_flat_config(cfg)
        save_config(cfg)

    cfg.setdefault("profiles_dir", DEFAULT_CONFIG["profiles_dir"])
    if not cfg.get("campaigns"):
        cfg["campaigns"] = json.loads(json.dumps(DEFAULT_CONFIG["campaigns"]))

    # Thiếu key nào trong chiến dịch thì bù bằng mặc định
    for camp in cfg["campaigns"].values():
        for key, val in CAMPAIGN_DEFAULTS.items():
            camp.setdefault(key, json.loads(json.dumps(val)))

    if cfg.get("active_campaign") not in cfg["campaigns"]:
        cfg["active_campaign"] = next(iter(cfg["campaigns"]))

    return cfg


def save_config(cfg):
    """Ghi cấu hình xuống config.json.

    Ghi ra file tạm rồi đổi tên đè lên: nếu máy tắt hoặc tiến trình bị kill
    giữa chừng thì config.json cũ vẫn nguyên vẹn, thay vì thành file JSON cụt
    làm tool không khởi động được nữa.
    """
    tmp = CONFIG_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
    os.replace(tmp, CONFIG_PATH)


def get_campaign(cfg, name=None):
    """Lấy dict cấu hình của một chiến dịch (mặc định: chiến dịch đang chọn)."""
    return cfg["campaigns"][name or cfg["active_campaign"]]


def slug(name):
    """Chuyển tên chiến dịch thành chuỗi an toàn để đặt tên file."""
    import re
    import unicodedata

    s = unicodedata.normalize("NFD", name.lower())
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]+", "_", s.replace("đ", "d")).strip("_") or "chien_dich"


def new_campaign(name, dang_dung=()):
    """Tạo cấu hình rỗng cho chiến dịch mới, đặt sẵn đường dẫn log theo tên.

    `dang_dung` là các đường dẫn log đã bị chiến dịch khác chiếm — cần tránh vì
    2 chiến dịch dùng chung file log sẽ coi bài của nhau là đã đăng. Cần thiết
    vì bỏ dấu làm "Tuyển dụng" và "Tuyen dung" ra cùng một slug.
    """
    base = slug(name)
    log_path = os.path.join(BASE_DIR, f"posted_log_{base}.json")
    n = 2
    while log_path in set(dang_dung):
        log_path = os.path.join(BASE_DIR, f"posted_log_{base}_{n}.json")
        n += 1

    camp = json.loads(json.dumps(CAMPAIGN_DEFAULTS))
    camp["posted_log"] = log_path
    camp["images_dir"] = os.path.join(BASE_DIR, f"anh_{base}")
    return camp
