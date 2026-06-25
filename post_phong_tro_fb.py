"""
Script đăng bài cho thuê phòng trọ lên group Facebook theo khu vực.
Mỗi phòng đăng đúng 1 bài lên group tương ứng với khu vực của nó.
Dữ liệu phòng từ file Excel, ảnh từ thư mục images.
# source .venv/bin/activate
# python post_phong_tro_fb.py
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

# ======================== CẤU HÌNH ========================
# Tự detect thư mục chứa script (hoạt động trên cả Mac và Windows)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_PATH = os.path.join(BASE_DIR, "phong_tro.xlsx")
IMAGES_DIR = os.path.join(BASE_DIR, "anh_phong")

# Hỗ trợ lưu nhiều tài khoản Facebook cùng lúc, mỗi acc 1 file session riêng.
# Các session mới được lưu trong thư mục fb_sessions/<tên acc>.json
# File cũ "fb_auth_state.json" (nếu có) vẫn được nhận diện như 1 acc mặc định.
SESSIONS_DIR = os.path.join(BASE_DIR, "fb_sessions")
LEGACY_AUTH_STATE = os.path.join(BASE_DIR, "fb_auth_state.json")

# Mapping khu vực → danh sách link group Facebook
# Mỗi khu vực có thể có NHIỀU group, phòng sẽ được đăng lên TẤT CẢ group của khu vực đó
# Key sẽ được normalize (bỏ dấu, viết thường, bỏ khoảng trắng) khi so khớp
KHU_VUC_GROUPS = {
    "Hà Đông": [
        "https://web.facebook.com/groups/phongtro.hanoi.hadong/",
        "https://web.facebook.com/groups/835892593690478/",
        "https://web.facebook.com/groups/3555475404499952/",
        "https://web.facebook.com/groups/1041520932684656/",
        "https://web.facebook.com/groups/1589501227985413/",
        "https://web.facebook.com/groups/631650078775924"
    ],
    "Thanh Xuân": [
        "https://web.facebook.com/groups/nhatrodongdathanhxuan/",
        "https://web.facebook.com/groups/176362986942358/",
        "https://web.facebook.com/groups/605109991280427/",
        "https://web.facebook.com/groups/908726406847516/",
        "https://web.facebook.com/groups/timphongtrodongdangatusothanhxuanhanoi/",
        "https://web.facebook.com/groups/1385595868491454/"
    ],
    "Bắc từ liêm": [
        "https://web.facebook.com/groups/2012063565703273/",
        "https://web.facebook.com/groups/1069950844149062/",
        # "https://web.facebook.com/groups/1115314632966592/",
        # "https://web.facebook.com/groups/1006836590849323/"
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
        "https://web.facebook.com/groups/2202922693407349/"

    ],
    "Mỹ Đình" : [
        "https://web.facebook.com/groups/1914388365626022/",
        "https://web.facebook.com/groups/2237019069763450/",
        "https://web.facebook.com/groups/phongtrocaugiaymydinhmetri/",
        "https://web.facebook.com/groups/507104870413526/",
        "https://web.facebook.com/groups/2148539488498466/",
        "https://web.facebook.com/groups/1542856739594335/"
    ],
    "Ba Đình": [
        "https://web.facebook.com/groups/phongtrobadinh.giatot/",
        "https://web.facebook.com/groups/757259302549445/",
        
    ],
    "Hai Bà Trưng" : [
        "https://web.facebook.com/groups/1747492728936509/",
        "https://web.facebook.com/groups/494231151747853/",
        "https://web.facebook.com/groups/724565062266526/",
        "https://web.facebook.com/groups/647543593374506/",
        "https://web.facebook.com/groups/2790717834511802/"
    ],
    "Hoàng Mai": [
        "https://web.facebook.com/groups/649420490421652/",
        "https://web.facebook.com/groups/900402601097250/",
        "https://web.facebook.com/groups/583641796650849/",
        "https://web.facebook.com/groups/778767189186540/"
    ],
    "Thanh Trì": [
        "https://web.facebook.com/groups/1145700312817923/",
        "https://web.facebook.com/groups/964760639021647/",
        "https://web.facebook.com/groups/TimPhongTroThanhTri/",
        "https://web.facebook.com/groups/1620926588466366/"
    ]
    # Thêm khu vực khác tại đây, thay bằng ID group thật...
}

MAX_POSTS = 0         # 0 = đăng hết, >0 = giới hạn số phòng
DELAY_BETWEEN_GROUPS = 10   # Delay (giây) giữa các group để tránh bị Facebook phát hiện
DELAY_BETWEEN_ROOMS = 10    # Delay (giây) giữa các phòng khác nhau
MAX_RETRIES = 2             # Số lần thử lại khi đăng lỗi
POSTED_LOG = os.path.join(BASE_DIR, "posted_log.json")  # File ghi nhận bài đã đăng


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
    """Kiểm tra phòng này đã đăng lên group này chưa."""
    key = f"{stt}|{group_url}"
    return key in posted


def mark_posted(posted, stt, group_url):
    """Đánh dấu đã đăng thành công."""
    key = f"{stt}|{group_url}"
    posted[key] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_posted_log(posted)


# ======================== ĐỌC DỮ LIỆU ========================

def load_rooms():
    """
    Đọc danh sách phòng từ Excel.
    Cấu trúc sheet 'PhongTro':
      A: STT (số thứ tự, dùng để sắp xếp thứ tự đăng)
      B: Mô tả chi tiết (dùng làm caption bài đăng)
      C: Địa chỉ (dùng để match với key trong KHU_VUC_GROUPS)
      D: Tên thư mục ảnh (trong IMAGES_DIR)
    """
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb["PhongTro"]
    rooms = []

    for row in range(2, ws.max_row + 1):
        stt = ws.cell(row=row, column=1).value
        if stt is None:
            continue

        mo_ta = ws.cell(row=row, column=2).value or ""
        dia_chi = ws.cell(row=row, column=3).value or ""
        folder_anh = ws.cell(row=row, column=4).value or ""

        images = find_images(folder_anh)

        rooms.append({
            "stt": int(float(stt)),
            "mo_ta": str(mo_ta).strip(),
            "dia_chi": str(dia_chi).strip(),
            "images": images,
            "row": row,
        })

    # Sắp xếp theo STT
    rooms.sort(key=lambda r: r["stt"])
    return rooms


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


def format_post_content(room):
    """Trả về mô tả chi tiết làm caption bài đăng."""
    return room["mo_ta"]


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


def get_group_urls(khu_vuc):
    """Tìm danh sách group Facebook tương ứng với khu vực (so khớp bỏ dấu). Trả về list URL hoặc []."""
    kv_norm = normalize(khu_vuc)
    if not kv_norm:
        return []
    for key, urls in KHU_VUC_GROUPS.items():
        key_norm = normalize(key)
        if kv_norm == key_norm or kv_norm in key_norm or key_norm in kv_norm:
            # Bỏ qua nếu list rỗng (khu vực chưa có group)
            return [u for u in urls if u.strip()]
    return []


# ======================== ĐĂNG NHẬP ========================

def login_facebook(p, auth_state):
    """Đăng nhập Facebook thủ công, lưu auth state vào đường dẫn `auth_state`."""
    print("Mở trình duyệt để đăng nhập Facebook...")
    print("Hãy đăng nhập thủ công, sau đó nhấn Enter trong terminal.")

    browser = p.chromium.launch(headless=False, slow_mo=200)
    context = browser.new_context(
        viewport={"width": 1280, "height": 900},
        locale="vi-VN",
    )
    page = context.new_page()
    page.goto("https://www.facebook.com/login", wait_until="networkidle")

    input("\n>>> Đã đăng nhập xong? Nhấn Enter để lưu session... ")

    context.storage_state(path=auth_state)
    print(f"Đã lưu session vào {auth_state}")
    browser.close()


def list_sessions():
    """Liệt kê các session đã lưu. Trả về list (tên hiển thị, đường dẫn file)."""
    sessions = []
    # File cũ (nếu có) coi như 1 acc mặc định
    if os.path.exists(LEGACY_AUTH_STATE):
        sessions.append(("Tài khoản mặc định (fb_auth_state)", LEGACY_AUTH_STATE))
    # Các acc trong thư mục fb_sessions/
    if os.path.isdir(SESSIONS_DIR):
        for f in sorted(os.listdir(SESSIONS_DIR)):
            if f.endswith(".json"):
                sessions.append((f[:-5], os.path.join(SESSIONS_DIR, f)))
    return sessions


def sanitize_account_name(name):
    """Chuẩn hóa tên acc thành tên file an toàn (bỏ ký tự đặc biệt)."""
    name = name.strip()
    name = re.sub(r"[^\w\-. ]", "", name, flags=re.UNICODE)
    name = name.strip().replace(" ", "_")
    return name


def login_new_account():
    """Đăng nhập tài khoản mới và lưu session vào fb_sessions/<tên>.json. Trả về đường dẫn."""
    while True:
        name = input("\n>>> Đặt tên cho tài khoản mới (vd: acc_chinh): ").strip()
        safe = sanitize_account_name(name)
        if not safe:
            print("  ⚠ Tên không hợp lệ, nhập lại.")
            continue
        os.makedirs(SESSIONS_DIR, exist_ok=True)
        path = os.path.join(SESSIONS_DIR, f"{safe}.json")
        if os.path.exists(path):
            ow = input(f"  Acc '{safe}' đã tồn tại. Ghi đè? (y/n): ").strip().lower()
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
    """In log kèm timestamp."""
    now = datetime.now().strftime("%H:%M:%S")
    print(f"[{now}] {msg}")


def countdown(seconds, label=""):
    """Đếm ngược hiển thị trên terminal."""
    for i in range(seconds, 0, -1):
        print(f"\r  ⏳ {label}Chờ {i}s...  ", end="", flush=True)
        time.sleep(1)
    print(f"\r  ✓ {label}Tiếp tục!      ")


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

def main():
    rooms = load_rooms()

    log(f"{'=' * 60}")
    log(f"Tổng số phòng trong Excel: {len(rooms)}")

    batch = rooms[:MAX_POSTS] if MAX_POSTS > 0 else rooms

    # Tính tổng số bài sẽ đăng
    total_posts = 0
    for room in batch:
        urls = get_group_urls(room["dia_chi"])
        total_posts += len(urls)

    log(f"Sẽ đăng {len(batch)} phòng → ~{total_posts} bài (theo thứ tự STT)")
    log(f"{'=' * 60}")

    # Chọn tài khoản Facebook để đăng bài (acc đã login hoặc đăng nhập mới)
    auth_state = choose_account()

    posted = load_posted_log()
    ok = 0
    fail = 0
    skip = 0
    dup = 0
    post_num = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=150)
        context = browser.new_context(
            storage_state=auth_state,
            viewport={"width": 1280, "height": 900},
            locale="vi-VN",
        )
        page = context.new_page()
        page.set_default_timeout(30000)

        # Kiểm tra nick đang đăng nhập
        log("Đang kiểm tra tài khoản Facebook...")
        page.goto("https://www.facebook.com/me", wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(3000)
        try:
            profile_name = page.locator("h1").first.inner_text(timeout=5000)
            log(f"Đang đăng nhập với nick: {profile_name}")
            confirm = input(f"\n>>> Đúng nick '{profile_name}'? Nhấn Enter để tiếp tục, gõ 'q' để hủy: ").strip()
            if confirm.lower() == 'q':
                log(f"Đã hủy. Xóa {os.path.basename(auth_state)} rồi chạy lại nếu muốn đổi nick.")
                browser.close()
                return
        except Exception:
            log("⚠ Session hết hạn. Xóa session cũ và mở lại để đăng nhập...")
            browser.close()
            os.remove(auth_state)
            # Mở trình duyệt mới để đăng nhập lại
            login_facebook(p, auth_state)
            # Khởi tạo lại browser với session mới
            browser = p.chromium.launch(headless=False, slow_mo=150)
            context = browser.new_context(
                storage_state=auth_state,
                viewport={"width": 1280, "height": 900},
                locale="vi-VN",
            )
            page = context.new_page()
            page.set_default_timeout(30000)

        for i, room in enumerate(batch):
            print()
            log(f"{'─' * 60}")
            log(f"PHÒNG {i + 1}/{len(batch)} | STT {room['stt']} | Row {room['row']}")
            log(f"  Địa chỉ : {room['dia_chi']}")
            log(f"  Caption  : {room['mo_ta'][:80]}{'...' if len(room['mo_ta']) > 80 else ''}")
            log(f"  Ảnh      : {len(room['images'])} file")

            # Tìm các group theo địa chỉ
            group_urls = get_group_urls(room["dia_chi"])
            if not group_urls:
                log(f"  ⚠ BỎ QUA — Không tìm thấy group cho '{room['dia_chi']}'")
                skip += 1
                continue

            log(f"  Tìm thấy {len(group_urls)} group cho khu vực này")

            content = format_post_content(room)

            for j, group_url in enumerate(group_urls):
                post_num += 1
                log(f"  ── Group {j + 1}/{len(group_urls)} (bài {post_num}/{total_posts})")

                # Kiểm tra đã đăng chưa
                if is_posted(posted, room["stt"], group_url):
                    log(f"  ⏭ ĐÃ ĐĂNG TRƯỚC ĐÓ — bỏ qua")
                    dup += 1
                    continue

                # Đăng bài với retry
                success = False
                for attempt in range(1, MAX_RETRIES + 1):
                    try:
                        post_to_group(page, group_url, content, room["images"])
                        mark_posted(posted, room["stt"], group_url)
                        ok += 1
                        success = True
                        break
                    except Exception as e:
                        log(f"  ✗ LỖI (lần {attempt}/{MAX_RETRIES}): {e}")
                        if attempt < MAX_RETRIES:
                            retry_delay = 10 + random.randint(0, 5)
                            countdown(retry_delay, "Thử lại: ")
                if not success:
                    fail += 1

                # Delay giữa các group để tránh bị Facebook phát hiện spam
                if j < len(group_urls) - 1:
                    delay = DELAY_BETWEEN_GROUPS + random.randint(0, 5)
                    countdown(delay, "Group tiếp: ")

            # Delay giữa các phòng
            if i < len(batch) - 1:
                delay = DELAY_BETWEEN_ROOMS + random.randint(0, 5)
                countdown(delay, "Phòng tiếp: ")

        # Lưu lại session
        context.storage_state(path=auth_state)
        browser.close()

    print()
    log(f"{'=' * 60}")
    log(f"KẾT QUẢ: {ok} thành công | {fail} lỗi | {skip} bỏ qua | {dup} đã đăng trước đó")
    log(f"{'=' * 60}")


if __name__ == "__main__":
    main()
