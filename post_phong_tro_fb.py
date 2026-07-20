"""
Đăng bài hàng loạt lên group Facebook, dùng chung cho nhiều chiến dịch
(cho thuê phòng trọ, tuyển dụng, seeding website...).

Mỗi dòng trong Excel là 1 bài; cột "Phân loại" quyết định bài đó được đăng lên
những group nào. Dữ liệu và bản đồ group của từng chiến dịch nằm trong
config.json — sửa bằng giao diện: python gui.py
# source .venv/bin/activate
# python post_phong_tro_fb.py
# .venv313/bin/python gui.py
⚠️ LƯU Ý:
- Facebook chống automation, tài khoản có thể bị checkpoint/khóa.
- Dùng tài khoản phụ, không dùng tài khoản chính.
"""
import os
import json
import time
import random
import unicodedata
import re
from datetime import datetime
import openpyxl
from playwright.sync_api import sync_playwright

import config as cfg_module
import paths

# Trỏ Playwright vào Chromium đóng kèm (chỉ có tác dụng khi chạy từ .exe)
paths.setup_playwright_env()

# ======================== CẤU HÌNH ========================
# Thư mục dữ liệu người dùng — chạy đúng trên Mac, Windows, và cả khi đã đóng
# gói thành .exe (khi đó là thư mục chứa file .exe). Xem paths.py.
BASE_DIR = paths.DATA_DIR

# Toàn bộ cấu hình đọc từ config.json (sửa được bằng giao diện: python gui.py).
# Các biến dưới đây được reload_config() gán lại theo chiến dịch đang chọn.
EXCEL_PATH = SHEET_NAME = IMAGES_DIR = PROFILES_DIR = POSTED_LOG = None
CAMPAIGN_NAME = ""
GROUPS = {}
MAX_POSTS = DELAY_BETWEEN_GROUPS = DELAY_BETWEEN_ROOMS = MAX_RETRIES = 0


def reload_config(campaign=None):
    """Nạp cấu hình của một chiến dịch vào các biến toàn cục.

    `campaign` là tên chiến dịch; None = chiến dịch đang chọn trong config.
    """
    global EXCEL_PATH, SHEET_NAME, IMAGES_DIR, PROFILES_DIR, POSTED_LOG
    global CAMPAIGN_NAME, GROUPS
    global MAX_POSTS, DELAY_BETWEEN_GROUPS, DELAY_BETWEEN_ROOMS, MAX_RETRIES

    cfg = cfg_module.load_config()
    camp = cfg_module.get_campaign(cfg, campaign)

    CAMPAIGN_NAME = campaign or cfg["active_campaign"]
    PROFILES_DIR = cfg["profiles_dir"]        # nick dùng chung mọi chiến dịch
    EXCEL_PATH = camp["excel_path"]
    SHEET_NAME = camp["sheet_name"]
    IMAGES_DIR = camp["images_dir"]
    # posted_log rỗng sẽ làm tắt im lặng toàn bộ chống đăng trùng, nên luôn ép
    # về một đường dẫn hợp lệ theo tên chiến dịch.
    POSTED_LOG = camp["posted_log"] or os.path.join(
        BASE_DIR, f"posted_log_{cfg_module.slug(CAMPAIGN_NAME)}.json")
    GROUPS = camp["groups"]
    MAX_POSTS = camp["max_posts"]
    DELAY_BETWEEN_GROUPS = camp["delay_between_groups"]
    DELAY_BETWEEN_ROOMS = camp["delay_between_rooms"]
    MAX_RETRIES = camp["max_retries"]
    return cfg


reload_config()

# GUI gán lại 2 hook này để hứng log và yêu cầu dừng giữa chừng.
LOG_FN = print          # nhận 1 chuỗi đã kèm timestamp
SHOULD_STOP = lambda: False

# ======================== TRÁNH ĐĂNG TRÙNG ========================

def load_posted_log():
    """Đọc danh sách bài đã đăng thành công. Format: {"stt|group_url": "2026-03-29 10:00"}"""
    if not os.path.exists(POSTED_LOG):
        return {}
    with open(POSTED_LOG, "r", encoding="utf-8") as f:
        return json.load(f)


def save_posted_log(posted):
    """Ghi lại danh sách bài đã đăng."""
    with open(POSTED_LOG, "w", encoding="utf-8") as f:
        json.dump(posted, f, ensure_ascii=False, indent=2)


def is_posted(posted, stt, group_url):
    """Kiểm tra bài này đã đăng lên group này chưa."""
    key = f"{stt}|{group_url}"
    return key in posted


def mark_posted(posted, stt, group_url):
    """Đánh dấu đã đăng thành công."""
    key = f"{stt}|{group_url}"
    posted[key] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_posted_log(posted)


# ======================== ĐỌC DỮ LIỆU ========================

def load_items():
    """
    Đọc danh sách bài cần đăng từ Excel của chiến dịch đang chọn.
    Cấu trúc sheet (tên sheet lấy từ cấu hình chiến dịch):
      A: STT (số thứ tự, dùng để sắp xếp thứ tự đăng)
      B: Nội dung bài (dùng làm caption)
      C: Phân loại (khu vực / ngành nghề / chủ đề — match với key trong GROUPS)
      D: Tên thư mục ảnh trong IMAGES_DIR (để trống nếu đăng bài không ảnh)
    """
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb[SHEET_NAME]
    items = []

    for row in range(2, ws.max_row + 1):
        stt = ws.cell(row=row, column=1).value
        if stt is None:
            continue

        noi_dung = ws.cell(row=row, column=2).value or ""
        phan_loai = ws.cell(row=row, column=3).value or ""
        folder_anh = ws.cell(row=row, column=4).value or ""

        images = find_images(folder_anh)

        items.append({
            "stt": int(float(stt)),
            "noi_dung": str(noi_dung).strip(),
            "phan_loai": str(phan_loai).strip(),
            "images": images,
            "row": row,
        })

    # Sắp xếp theo STT
    items.sort(key=lambda r: r["stt"])
    return items


def find_images(folder_name):
    """Tìm tất cả ảnh trong thư mục con của IMAGES_DIR."""
    if not folder_name:
        return []
    folder = os.path.join(IMAGES_DIR, str(folder_name).strip())
    if not os.path.isdir(folder):
        return []
    exts = (".jpg", ".jpeg", ".png", ".webp")
    return [
        os.path.join(folder, f)
        for f in sorted(os.listdir(folder))
        if f.lower().endswith(exts)
    ]


def format_post_content(item):
    """Trả về nội dung dùng làm caption bài đăng."""
    return item["noi_dung"]


def normalize(text):
    """Bỏ dấu tiếng Việt, chuyển thường, bỏ khoảng trắng + ký tự đặc biệt.
    VD: 'Thanh Xuân' → 'thanhxuan', 'Q. Thủ Đức' → 'qthucduc'
    """
    text = str(text).strip().lower()
    # Bỏ dấu tiếng Việt
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    # Xử lý đ/Đ
    text = text.replace("đ", "d")
    # Bỏ tất cả ký tự không phải chữ/số
    text = re.sub(r"[^a-z0-9]", "", text)
    return text


def tim_phan_loai_khop(phan_loai):
    """Trả về (khớp_chính_xác, [các_khớp_một_phần]) cho một giá trị cột C.

    Tách riêng để tab kiểm tra dữ liệu chỉ ra được chỗ nhập nhằng: 'Kế toán'
    và 'Kế toán trưởng' khớp lẫn nhau theo kiểu chuỗi con, nếu không cảnh báo
    thì bài sẽ âm thầm đăng lên sai tập group.
    """
    kv_norm = normalize(phan_loai)
    if not kv_norm:
        return None, []

    chinh_xac = None
    mot_phan = []
    for key in GROUPS:
        key_norm = normalize(key)
        if kv_norm == key_norm:
            chinh_xac = key
        elif kv_norm in key_norm or key_norm in kv_norm:
            mot_phan.append(key)
    return chinh_xac, mot_phan


def get_group_urls(phan_loai):
    """Tìm danh sách group ứng với phân loại của bài (so khớp bỏ dấu). Trả về list URL hoặc [].

    Khớp chính xác được ưu tiên; chỉ khi không có mới xét khớp một phần như cũ.
    """
    chinh_xac, mot_phan = tim_phan_loai_khop(phan_loai)
    key = chinh_xac or (mot_phan[0] if mot_phan else None)
    if key is None:
        return []
    if chinh_xac is None and len(mot_phan) > 1:
        log(f"  ⚠ '{phan_loai}' khớp nhập nhằng với: {', '.join(mot_phan)} — đang dùng '{key}'")
    # Bỏ qua nếu list rỗng (phân loại chưa gán group nào)
    return [u for u in GROUPS[key] if u.strip()]


# ======================== KIỂM TRA DỮ LIỆU ========================

def kiem_tra_du_lieu(campaign=None):
    """Soát toàn bộ cấu hình + Excel của một chiến dịch TRƯỚC khi đăng.

    Trả về list dict: {"muc": "lỗi"|"canh_bao", "o": mô tả vị trí, "chi_tiet": ...}
      - "lỗi"      : chắc chắn hỏng, không nên chạy
      - "canh_bao" : chạy được nhưng kết quả có thể không như ý
    """
    if campaign:
        reload_config(campaign)

    van_de = []
    loi = lambda o, ct: van_de.append({"muc": "loi", "o": o, "chi_tiet": ct})
    canh_bao = lambda o, ct: van_de.append({"muc": "canh_bao", "o": o, "chi_tiet": ct})

    # --- Cấu hình ---
    if not EXCEL_PATH or not os.path.exists(EXCEL_PATH):
        loi("File Excel", f"Không tìm thấy: {EXCEL_PATH or '(chưa đặt)'}")
        return van_de
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH)
    except Exception as e:
        loi("File Excel", f"Không đọc được: {e}")
        return van_de
    if SHEET_NAME not in wb.sheetnames:
        loi("Sheet", f"File không có sheet '{SHEET_NAME}'. Đang có: {', '.join(wb.sheetnames)}")
        return van_de

    if not GROUPS:
        loi("Nhóm", "Chiến dịch chưa khai báo phân loại nào")
    if not IMAGES_DIR or not os.path.isdir(IMAGES_DIR):
        canh_bao("Thư mục ảnh", f"Không tồn tại: {IMAGES_DIR or '(chưa đặt)'} — mọi bài sẽ đăng không ảnh")

    # --- Phân loại nhập nhằng / rỗng ---
    for key, urls in GROUPS.items():
        sach = [u for u in urls if u.strip()]
        if not sach:
            canh_bao(f"Phân loại '{key}'", "Chưa có link group nào — bài thuộc loại này sẽ bị bỏ qua")
        xau = [u for u in sach if not u.startswith("http")]
        if xau:
            loi(f"Phân loại '{key}'", f"Link không hợp lệ: {', '.join(xau[:3])}")

    ten = list(GROUPS)
    for i, a in enumerate(ten):
        for b in ten[i + 1:]:
            na, nb = normalize(a), normalize(b)
            if na == nb:
                loi("Phân loại trùng", f"'{a}' và '{b}' giống nhau sau khi bỏ dấu → luôn khớp nhầm")
            elif na in nb or nb in na:
                canh_bao("Phân loại lồng nhau",
                         f"'{a}' và '{b}' khớp chuỗi con lẫn nhau — ô Excel phải ghi "
                         f"CHÍNH XÁC một trong hai, nếu không sẽ đăng sai nhóm")

    # --- Từng dòng Excel ---
    ws = wb[SHEET_NAME]
    stt_da_gap = {}
    for row in range(2, ws.max_row + 1):
        stt = ws.cell(row=row, column=1).value
        noi_dung = ws.cell(row=row, column=2).value
        phan_loai = ws.cell(row=row, column=3).value
        folder = ws.cell(row=row, column=4).value

        trong = [v for v in (stt, noi_dung, phan_loai, folder)
                 if v is not None and str(v).strip()]
        if not trong:
            continue

        vi_tri = f"Dòng {row}"
        if stt is None or str(stt).strip() == "":
            loi(vi_tri, "Thiếu STT ở cột A → cả dòng bị bỏ qua im lặng")
            continue
        try:
            stt = int(float(stt))
        except (TypeError, ValueError):
            loi(vi_tri, f"STT '{stt}' không phải số")
            continue

        if stt in stt_da_gap:
            loi(vi_tri, f"STT {stt} trùng với dòng {stt_da_gap[stt]} → dòng sau bị bỏ qua vĩnh viễn")
        else:
            stt_da_gap[stt] = row

        vi_tri = f"Dòng {row} (STT {stt})"
        if not noi_dung or not str(noi_dung).strip():
            loi(vi_tri, "Nội dung bài trống")

        if not phan_loai or not str(phan_loai).strip():
            loi(vi_tri, "Thiếu phân loại ở cột C → không biết đăng lên nhóm nào")
        else:
            chinh_xac, mot_phan = tim_phan_loai_khop(phan_loai)
            if not chinh_xac and not mot_phan:
                loi(vi_tri, f"Phân loại '{str(phan_loai).strip()}' chưa khai trong danh sách nhóm → bài bị bỏ qua")
            elif not chinh_xac and len(mot_phan) > 1:
                canh_bao(vi_tri, f"'{str(phan_loai).strip()}' khớp nhiều phân loại: "
                                 f"{', '.join(mot_phan)} → sẽ dùng '{mot_phan[0]}'")
            elif not chinh_xac:
                canh_bao(vi_tri, f"'{str(phan_loai).strip()}' không khớp chính xác, "
                                 f"đang hiểu là '{mot_phan[0]}'")

        if folder and str(folder).strip():
            duong_dan = os.path.join(IMAGES_DIR or "", str(folder).strip())
            if not os.path.isdir(duong_dan):
                canh_bao(vi_tri, f"Không có thư mục ảnh '{str(folder).strip()}' → "
                                 f"bài sẽ đăng KHÔNG ẢNH và bị ghi nhận là đã đăng")
            elif not find_images(str(folder).strip()):
                canh_bao(vi_tri, f"Thư mục ảnh '{str(folder).strip()}' không có file ảnh nào")

    return van_de


# ======================== ĐĂNG NHẬP ========================

def open_profile(p, profile_dir, slow_mo=150):
    """Mở Chrome thật với profile bền tại `profile_dir`.

    Dùng channel='chrome' (Chrome cài trên máy) thay vì Chromium đóng gói của
    Playwright, và tắt cờ AutomationControlled — nếu không Facebook đọc được
    navigator.webdriver và từ chối cấp session dù mật khẩu đúng.
    Cookie lưu thẳng trong profile nên không cần storage_state.
    """
    os.makedirs(profile_dir, exist_ok=True)
    try:
        context = p.chromium.launch_persistent_context(
            profile_dir,
            channel="chrome",
            headless=False,
            slow_mo=slow_mo,
            viewport={"width": 1280, "height": 900},
            locale="vi-VN",
            args=["--disable-blink-features=AutomationControlled"],
        )
    except Exception as e:
        # Máy chưa cài Chrome là nguyên nhân thường gặp nhất, nhưng Playwright
        # chỉ báo một dòng tiếng Anh về "channel chrome" mà người dùng không
        # hiểu phải làm gì. Dịch sang hướng dẫn cụ thể.
        if "chrome" in str(e).lower() and "executable" in str(e).lower():
            raise RuntimeError(
                "Không tìm thấy Google Chrome trên máy này.\n\n"
                "Tool cần Chrome thật để đăng nhập Facebook. Hãy tải và cài "
                "Chrome tại https://www.google.com/chrome rồi mở lại tool."
            ) from e
        raise
    return context


def get_page(context):
    """Lấy tab đầu tiên của persistent context (Chrome luôn mở sẵn 1 tab)."""
    page = context.pages[0] if context.pages else context.new_page()
    page.set_default_timeout(30000)
    return page


def login_facebook(p, profile_dir):
    """Đăng nhập Facebook thủ công vào profile. Cookie tự lưu trong profile_dir."""
    print("Mở Chrome để đăng nhập Facebook...")
    print("Hãy đăng nhập thủ công, sau đó nhấn Enter trong terminal.")

    context = open_profile(p, profile_dir, slow_mo=200)
    page = get_page(context)
    page.goto("https://www.facebook.com/login", wait_until="domcontentloaded")

    input("\n>>> Đã đăng nhập xong? Nhấn Enter để tiếp tục... ")

    if is_logged_in(page):
        print(f"Đã lưu đăng nhập vào profile {profile_dir}")
        ok = True
    else:
        print("⚠ Chưa đăng nhập được (vẫn ở màn login). Profile giữ nguyên, chạy lại để thử tiếp.")
        ok = False
    context.close()
    return ok


def is_logged_in(page):
    """Session còn sống hay không. Chỉ coi là hết hạn khi Facebook đá về trang
    login — không dựa vào việc render được tên nick, vì trang profile có thể
    load chậm hoặc đổi layout."""
    url = page.url.lower()
    if "login" in url or "checkpoint" in url:
        return False
    if page.locator("input[name='email'], input[name='pass']").count() > 0:
        return False
    cookies = page.context.cookies("https://www.facebook.com")
    return any(c["name"] == "c_user" for c in cookies)


def get_profile_name(page):
    """Lấy tên nick, thử vài selector. Trả về '(không đọc được tên)' nếu thất bại."""
    for sel in ["h1", "[role='main'] h2", "title"]:
        try:
            name = page.locator(sel).first.inner_text(timeout=3000).strip()
            if name and "facebook" not in name.lower():
                return name
        except Exception:
            continue
    return "(không đọc được tên)"


def list_sessions():
    """Liệt kê các profile đã đăng nhập. Trả về list (tên hiển thị, đường dẫn profile)."""
    sessions = []
    if os.path.isdir(PROFILES_DIR):
        for d in sorted(os.listdir(PROFILES_DIR)):
            full = os.path.join(PROFILES_DIR, d)
            if os.path.isdir(full):
                sessions.append((d, full))
    return sessions


def sanitize_account_name(name):
    """Chuẩn hóa tên acc thành tên file an toàn (bỏ ký tự đặc biệt)."""
    name = name.strip()
    name = re.sub(r"[^\w\-. ]", "", name, flags=re.UNICODE)
    name = name.strip().replace(" ", "_")
    return name


def login_new_account():
    """Đăng nhập tài khoản mới vào fb_profiles/<tên>/. Trả về đường dẫn profile."""
    while True:
        name = input("\n>>> Đặt tên cho tài khoản mới (vd: acc_chinh): ").strip()
        safe = sanitize_account_name(name)
        if not safe:
            print("  ⚠ Tên không hợp lệ, nhập lại.")
            continue
        os.makedirs(PROFILES_DIR, exist_ok=True)
        path = os.path.join(PROFILES_DIR, safe)
        if os.path.isdir(path):
            ow = input(f"  Acc '{safe}' đã tồn tại. Dùng lại profile này? (y/n): ").strip().lower()
            if ow != "y":
                continue
        break

    with sync_playwright() as p:
        login_facebook(p, path)
    return path


def choose_account():
    """Menu chọn tài khoản đăng bài. Trả về đường dẫn file session.

    [1] Dùng tài khoản đã đăng nhập → chọn từ danh sách session đã lưu
    [2] Đăng nhập tài khoản mới → đăng nhập rồi lưu state cho acc mới
    """
    while True:
        print()
        log(f"{'=' * 60}")
        log("CHỌN TÀI KHOẢN ĐĂNG BÀI")
        print("   [1] Dùng tài khoản đã đăng nhập")
        print("   [2] Đăng nhập tài khoản mới")
        log(f"{'=' * 60}")

        choice = input("\n>>> Chọn (1/2): ").strip()

        if choice == "1":
            sessions = list_sessions()
            if not sessions:
                print("  ⚠ Chưa có tài khoản nào được lưu. Hãy chọn [2] để đăng nhập mới.")
                continue
            print("\n  Danh sách tài khoản đã đăng nhập:")
            for idx, (name, _) in enumerate(sessions, start=1):
                print(f"   [{idx}] {name}")
            sel = input(f"\n>>> Chọn tài khoản (1-{len(sessions)}): ").strip()
            if sel.isdigit() and 1 <= int(sel) <= len(sessions):
                name, path = sessions[int(sel) - 1]
                log(f"Đã chọn: {name}")
                return path
            print("  ⚠ Lựa chọn không hợp lệ.")
            continue

        if choice == "2":
            return login_new_account()

        print("  ⚠ Lựa chọn không hợp lệ. Nhập 1 hoặc 2.")


# ======================== ĐĂNG BÀI LÊN 1 GROUP ========================

def log(msg):
    """In log kèm timestamp qua LOG_FN (terminal hoặc ô log của GUI)."""
    now = datetime.now().strftime("%H:%M:%S")
    LOG_FN(f"[{now}] {msg}")


class StopRequested(Exception):
    """Người dùng bấm Dừng trên giao diện."""


def check_stop():
    """Ném StopRequested nếu GUI yêu cầu dừng."""
    if SHOULD_STOP():
        raise StopRequested()


def countdown(seconds, label=""):
    """Chờ `seconds` giây, mỗi giây kiểm tra xem có bị yêu cầu dừng không."""
    for i in range(seconds, 0, -1):
        check_stop()
        if LOG_FN is print:
            print(f"\r  ⏳ {label}Chờ {i}s...  ", end="", flush=True)
        time.sleep(1)
    if LOG_FN is print:
        print(f"\r  ✓ {label}Tiếp tục!      ")
    else:
        log(f"  ✓ {label}Đã chờ {seconds}s, tiếp tục!")


def close_popup_if_any(page):
    """Đóng popup 'Add to your post' / 'Thêm vào bài viết' nếu nó xuất hiện (thường gặp trên Windows)."""
    try:
        # Kiểm tra popup có đang hiển thị không (tìm text đặc trưng)
        popup_visible = page.locator(
            "[role='dialog'] :text-is('Add to your post'), "
            "[role='dialog'] :text-is('Thêm vào bài viết')"
        )
        if popup_visible.count() == 0 or not popup_visible.first.is_visible():
            return  # Không có popup, không cần làm gì

        log(f"  Đóng popup 'Add to your post'...")

        # Cách 1: Tìm nút back bằng nhiều selector
        back_selectors = [
            "[role='dialog'] [aria-label='Back']",
            "[role='dialog'] [aria-label='Quay lại']",
            "[role='dialog'] [aria-label='Trở về']",
            "[role='dialog'] [aria-label='Close']",
            "[role='dialog'] [aria-label='Đóng']",
        ]
        for sel in back_selectors:
            btn = page.locator(sel)
            if btn.count() > 0 and btn.first.is_visible():
                btn.first.click()
                page.wait_for_timeout(1000)
                return

        # Cách 2: Tìm nút back bằng JavaScript (icon SVG mũi tên)
        clicked = page.evaluate("""() => {
            const dialog = document.querySelector('[role="dialog"]');
            if (!dialog) return false;
            // Tìm tất cả div/span clickable gần text "Add to your post"
            const heading = dialog.querySelector('span');
            if (!heading) return false;
            // Tìm phần tử trước heading (thường là nút back)
            const parent = heading.closest('div');
            if (parent) {
                const prevSibling = parent.previousElementSibling;
                if (prevSibling) {
                    prevSibling.click();
                    return true;
                }
            }
            return false;
        }""")
        if clicked:
            page.wait_for_timeout(1000)
            return

        # Cách 3: Nhấn Escape để đóng sub-popup
        page.keyboard.press("Escape")
        page.wait_for_timeout(1000)

    except Exception:
        pass


def find_and_click_post_button(page):
    """Tìm và click nút Đăng/Post. Trả về True nếu tìm thấy."""
    # Đóng popup phụ trước nếu có
    close_popup_if_any(page)

    post_btn = page.locator(
        "[role='dialog'] [role='button']:has-text('Đăng'), "
        "[role='dialog'] [role='button']:has-text('Post'), "
        "[role='dialog'] div[role='button'][aria-label='Đăng'], "
        "[role='dialog'] div[role='button'][aria-label='Post']"
    ).first
    post_btn.click(timeout=10000)
    return True


def wait_for_post_complete(page):
    """Chờ bài đăng hoàn tất. Xử lý cả Mac và Windows."""
    try:
        # Cách 1: Chờ dialog đóng (thường hoạt động trên Mac)
        page.locator("[role='dialog']").first.wait_for(state="hidden", timeout=20000)
        page.wait_for_timeout(2000)
        return
    except Exception:
        pass

    # Cách 2: Dialog chưa đóng — thử đóng popup phụ rồi chờ tiếp
    log(f"  Dialog chưa đóng, thử xử lý...")
    close_popup_if_any(page)

    try:
        page.locator("[role='dialog']").first.wait_for(state="hidden", timeout=10000)
        page.wait_for_timeout(2000)
        return
    except Exception:
        pass

    # Cách 3: Nhấn Escape để đóng dialog
    log(f"  Nhấn Escape để đóng...")
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(3000)
    except Exception:
        pass

    # Cách 4: Nếu vẫn không đóng được, navigate đi luôn (bài đã gửi rồi)
    page.wait_for_timeout(2000)


def upload_images(page, images):
    """Upload ảnh vào dialog tạo bài. Xử lý cả Mac và Windows."""
    log(f"  Đang upload {len(images)} ảnh...")
    uploaded = False

    # === Cách 1: Tìm input[multiple] sẵn trong dialog (Mac) ===
    multi_input = page.locator("[role='dialog'] input[type='file'][multiple]")
    if multi_input.count() > 0:
        multi_input.first.set_input_files(images)
        log(f"  Đã chọn {len(images)} ảnh (multiple input)")
        uploaded = True

    # === Cách 2: Click Photo/video rồi upload (Windows) ===
    if not uploaded:
        # Nếu popup "Add to your post" đang hiển thị → click "Photo/video" từ đó
        # Nếu không → tìm nút Photo/video ở footer dialog
        log(f"  Click nút Photo/video...")
        try:
            photo_btn = page.locator(
                "[role='dialog'] [aria-label*='hoto'], "
                "[role='dialog'] [aria-label*='ảnh'], "
                "[role='dialog'] [aria-label*='Ảnh']"
            ).first
            photo_btn.click(timeout=5000)
            page.wait_for_timeout(2000)
        except Exception:
            # Thử click text "Photo/video" trong popup "Add to your post"
            try:
                page.locator("[role='dialog']").first.locator(
                    "text=Photo/video, text=Ảnh/video, text=Ảnh/Video"
                ).first.click(timeout=3000)
                page.wait_for_timeout(2000)
            except Exception:
                log(f"  Không tìm thấy nút Photo/video")

        # Tìm input file sau khi click
        file_input = page.locator("input[type='file'][accept*='image']")
        if file_input.count() > 0:
            try:
                file_input.first.set_input_files(images)
                log(f"  Đã chọn {len(images)} ảnh")
                uploaded = True
            except Exception:
                # Upload từng ảnh nếu input không nhận multiple
                log(f"  Upload từng ảnh...")
                for idx, img_path in enumerate(images):
                    try:
                        file_input.first.set_input_files(img_path)
                        log(f"    Ảnh {idx + 1}/{len(images)}: {os.path.basename(img_path)}")
                        page.wait_for_timeout(2000)
                        uploaded = True
                    except Exception as e:
                        log(f"    ⚠ Lỗi ảnh {idx + 1}: {e}")

    if not uploaded:
        log(f"  ⚠ Không upload được ảnh, đăng bài không có ảnh")
        return False

    # Chờ ảnh preview hiển thị
    try:
        page.locator(
            "[role='dialog'] img[src*='blob:'], "
            "[role='dialog'] img[src*='scontent'], "
            "[role='dialog'] div[data-imagefbid]"
        ).first.wait_for(state="visible", timeout=30000)
    except Exception:
        log(f"  Không thấy preview, chờ thêm...")

    page.wait_for_timeout(2000 + len(images) * 1000)
    log(f"  Upload ảnh xong!")
    return True


def post_to_group(page, group_url, content, images):
    """Đăng 1 bài lên 1 group Facebook. Tương thích cả Mac và Windows."""
    log(f"  Đang mở group: {group_url}")
    page.goto(group_url, wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(3000)

    # Click vào ô "Bạn viết gì đi..." để mở popup tạo bài
    log(f"  Mở ô viết bài...")
    create_post_box = page.locator(
        "[role='button']:has-text('Bạn viết gì đi'), "
        "[role='button']:has-text('Write something'), "
        "[role='button']:has-text('Viết gì đó'), "
        "span:has-text('Bạn viết gì đi'), "
        "span:has-text('Write something')"
    ).first
    create_post_box.click(timeout=10000)
    page.wait_for_timeout(2000)

    # Kiểm tra popup "Add to your post" có hiện không
    add_to_post = page.locator(
        "[role='dialog'] :text-is('Add to your post'), "
        "[role='dialog'] :text-is('Thêm vào bài viết')"
    )
    has_add_popup = add_to_post.count() > 0 and add_to_post.first.is_visible()

    # === NẾU CÓ POPUP "Add to your post" (Windows) ===
    # Upload ảnh TRƯỚC từ popup này, rồi quay lại nhập text
    if has_add_popup and images:
        log(f"  Phát hiện popup 'Add to your post' — upload ảnh từ đây...")
        upload_images(page, images)
        images = []  # Đã upload, không cần upload lại

    # Đóng popup "Add to your post" nếu còn
    close_popup_if_any(page)

    # Nhập nội dung bài viết vào editor
    log(f"  Đang nhập nội dung ({len(content)} ký tự)...")
    editor = page.locator(
        "[role='dialog'] [contenteditable='true'], "
        "[role='dialog'] [role='textbox'], "
        "[role='dialog'] p[data-lexical-text]"
    ).first
    editor.click()
    page.wait_for_timeout(500)

    # Gõ từng dòng
    for i, line in enumerate(content.split("\n")):
        if i > 0:
            page.keyboard.press("Shift+Enter")
            page.wait_for_timeout(random.randint(100, 300))
        if line:
            page.keyboard.insert_text(line)
            page.wait_for_timeout(random.randint(200, 500))

    page.wait_for_timeout(1000)

    # Upload ảnh nếu chưa upload ở trên (Mac — không có popup "Add to your post")
    if images:
        upload_images(page, images)

    # Click nút Đăng
    log(f"  Nhấn nút Đăng...")
    find_and_click_post_button(page)

    # Chờ bài đăng hoàn tất
    log(f"  Đang chờ bài đăng hoàn tất...")
    wait_for_post_complete(page)

    log(f"  ✓ THÀNH CÔNG!")
    return True


# ======================== MAIN ========================

def run_posting(profile_dir, confirm_nick=None, campaign=None):
    """Chạy toàn bộ vòng đăng bài của một chiến dịch với profile đã chọn.

    Không gọi input() ở đâu cả nên dùng được cho cả CLI lẫn GUI.
    `confirm_nick(ten_nick) -> bool`: hỏi lại người dùng có đúng nick không
    (CLI truyền hàm hỏi qua terminal, GUI truyền None vì đã chọn nick sẵn).
    `campaign`: tên chiến dịch; None = dùng chiến dịch đang chọn trong config.
    Trả về dict thống kê, hoặc None nếu dừng trước khi đăng.
    """
    if campaign:
        reload_config(campaign)

    items = load_items()

    log(f"{'=' * 60}")
    log(f"Chiến dịch: {CAMPAIGN_NAME}")
    log(f"Tổng số bài trong Excel: {len(items)}")

    batch = items[:MAX_POSTS] if MAX_POSTS > 0 else items

    # Tính tổng số bài sẽ đăng
    total_posts = 0
    for item in batch:
        urls = get_group_urls(item["phan_loai"])
        total_posts += len(urls)

    log(f"Sẽ đăng {len(batch)} mục → ~{total_posts} lượt đăng (theo thứ tự STT)")
    log(f"{'=' * 60}")

    posted = load_posted_log()

    with sync_playwright() as p:
        context = open_profile(p, profile_dir)
        page = get_page(context)

        # Kiểm tra nick đang đăng nhập
        log("Đang kiểm tra tài khoản Facebook...")
        page.goto("https://www.facebook.com/me", wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(3000)
        if is_logged_in(page):
            profile_name = get_profile_name(page)
            log(f"Đang đăng nhập với nick: {profile_name}")
            if confirm_nick and not confirm_nick(profile_name):
                log("Đã hủy — không đúng nick.")
                context.close()
                return None
        else:
            log("⚠ Profile này chưa đăng nhập Facebook. Hãy đăng nhập cho acc này rồi chạy lại.")
            context.close()
            return None

        try:
            stats = _post_all(page, batch, posted, total_posts)
        except StopRequested:
            log("⏹ Đã dừng theo yêu cầu.")
            stats = None
        finally:
            # Cookie đã tự lưu trong profile, chỉ cần đóng
            context.close()

    return stats


def _post_all(page, batch, posted, total_posts):
    """Vòng lặp đăng bài cho toàn bộ danh sách. Trả về dict thống kê."""
    ok = fail = skip = dup = post_num = 0
    loi_chi_tiet = []       # để cuối buổi liệt kê rõ bài nào hỏng ở group nào

    for i, item in enumerate(batch):
            check_stop()
            log(f"{'─' * 60}")
            log(f"MỤC {i + 1}/{len(batch)} | STT {item['stt']} | Row {item['row']}")
            log(f"  Phân loại: {item['phan_loai']}")
            log(f"  Caption  : {item['noi_dung'][:80]}{'...' if len(item['noi_dung']) > 80 else ''}")
            log(f"  Ảnh      : {len(item['images'])} file")

            # Tìm các group theo địa chỉ
            group_urls = get_group_urls(item["phan_loai"])
            if not group_urls:
                log(f"  ⚠ BỎ QUA — Không tìm thấy group cho '{item['phan_loai']}'")
                loi_chi_tiet.append(
                    f"STT {item['stt']}: phân loại '{item['phan_loai']}' chưa khai nhóm → bỏ qua")
                skip += 1
                continue

            log(f"  Tìm thấy {len(group_urls)} group cho phân loại này")

            content = format_post_content(item)

            for j, group_url in enumerate(group_urls):
                post_num += 1
                log(f"  ── Group {j + 1}/{len(group_urls)} (bài {post_num}/{total_posts})")

                # Kiểm tra đã đăng chưa
                if is_posted(posted, item["stt"], group_url):
                    log(f"  ⏭ ĐÃ ĐĂNG TRƯỚC ĐÓ — bỏ qua")
                    dup += 1
                    continue

                # Đăng bài với retry
                success = False
                loi_cuoi = ""
                for attempt in range(1, MAX_RETRIES + 1):
                    try:
                        post_to_group(page, group_url, content, item["images"])
                    except Exception as e:
                        loi_cuoi = str(e).split(chr(10))[0][:200]
                        log(f"  ✗ LỖI (lần {attempt}/{MAX_RETRIES}): {e}")
                        if attempt < MAX_RETRIES:
                            retry_delay = 10 + random.randint(0, 5)
                            countdown(retry_delay, "Thử lại: ")
                        continue

                    # Bài ĐÃ lên Facebook. Ghi dấu tách riêng khỏi try ở trên:
                    # nếu ghi log lỗi mà vẫn nằm chung try thì sẽ bị coi là đăng
                    # hỏng và retry, làm bài thứ hai y hệt lên cùng group.
                    ok += 1
                    success = True
                    try:
                        mark_posted(posted, item["stt"], group_url)
                    except Exception as e:
                        log(f"  ⚠ ĐÃ ĐĂNG XONG nhưng không ghi được log chống trùng: {e}")
                        log(f"  ⚠ Lần chạy sau bài STT {item['stt']} có thể bị đăng lại lên group này.")
                    break
                if not success:
                    fail += 1
                    loi_chi_tiet.append(
                        f"STT {item['stt']} → {group_url}: {loi_cuoi}")

                # Delay giữa các group để tránh bị Facebook phát hiện spam
                if j < len(group_urls) - 1:
                    delay = DELAY_BETWEEN_GROUPS + random.randint(0, 5)
                    countdown(delay, "Group tiếp: ")

            # Delay giữa các phòng
            if i < len(batch) - 1:
                delay = DELAY_BETWEEN_ROOMS + random.randint(0, 5)
                countdown(delay, "Mục tiếp: ")

    log(f"{'=' * 60}")
    log(f"KẾT QUẢ: {ok} thành công | {fail} lỗi | {skip} bỏ qua | {dup} đã đăng trước đó")
    log(f"{'=' * 60}")

    if loi_chi_tiet:
        log("")
        log(f"CHI TIẾT {len(loi_chi_tiet)} TRƯỜNG HỢP KHÔNG ĐĂNG ĐƯỢC:")
        for dong in loi_chi_tiet:
            log(f"  • {dong}")
        log(f"{'=' * 60}")

    return {"ok": ok, "fail": fail, "skip": skip, "dup": dup, "loi": loi_chi_tiet}


def choose_campaign():
    """Menu chọn chiến dịch. Trả về tên chiến dịch."""
    cfg = cfg_module.load_config()
    names = list(cfg["campaigns"])
    if len(names) == 1:
        return names[0]

    print()
    log(f"{'=' * 60}")
    log("CHỌN CHIẾN DỊCH")
    for idx, name in enumerate(names, start=1):
        mark = " (đang chọn)" if name == cfg["active_campaign"] else ""
        print(f"   [{idx}] {name}{mark}")
    log(f"{'=' * 60}")

    while True:
        sel = input(f"\n>>> Chọn (1-{len(names)}, Enter = đang chọn): ").strip()
        if not sel:
            return cfg["active_campaign"]
        if sel.isdigit() and 1 <= int(sel) <= len(names):
            return names[int(sel) - 1]
        print("  ⚠ Lựa chọn không hợp lệ.")


def main():
    """Chạy bằng terminal: chọn chiến dịch, chọn acc, rồi đăng."""
    campaign = choose_campaign()
    reload_config(campaign)
    profile_dir = choose_account()

    def confirm_nick(name):
        answer = input(f"\n>>> Đúng nick '{name}'? Nhấn Enter để tiếp tục, gõ 'q' để hủy: ").strip()
        return answer.lower() != "q"

    run_posting(profile_dir, confirm_nick=confirm_nick, campaign=campaign)


if __name__ == "__main__":
    main()
