# Hướng dẫn cài đặt và chạy Tool Đăng Bài Phòng Trọ Facebook

## Giới thiệu

Tool tự động đăng bài cho thuê phòng trọ lên các group Facebook theo khu vực.
- Đọc dữ liệu phòng từ file Excel
- Tự động upload ảnh, nhập nội dung, đăng bài
- Mỗi phòng được đăng lên đúng group tương ứng với khu vực của nó
- Hỗ trợ nhiều group cho mỗi khu vực
- Tự động bỏ qua bài đã đăng (tránh đăng trùng)

---

## Bước 1: Cài đặt Python

### Trên Windows

1. Mở trình duyệt, truy cập: https://www.python.org/downloads/
2. Nhấn nút **"Download Python 3.x.x"** (chọn phiên bản mới nhất)
3. Mở file `.exe` vừa tải về
4. **QUAN TRỌNG: Tick vào ô "Add python.exe to PATH"** ở màn hình đầu tiên (nếu không tick, máy sẽ không nhận lệnh `python`)
5. Nhấn **"Install Now"**
6. Chờ cài xong, nhấn **Close**

Kiểm tra đã cài thành công chưa:
- Nhấn tổ hợp phím `Win + R`
- Gõ `cmd` rồi nhấn Enter (sẽ mở cửa sổ đen gọi là Command Prompt)
- Gõ lệnh sau rồi nhấn Enter:

```
python --version
```

Nếu hiện ra `Python 3.x.x` là thành công. Nếu báo lỗi `'python' is not recognized` thì cần cài lại Python và nhớ tick "Add to PATH".

> **Lưu ý:** Trên Windows dùng lệnh `python` (không có số 3). Trên Mac dùng `python3`.

### Trên Mac

1. Nhấn `Cmd + Space`, gõ **Terminal**, nhấn Enter (sẽ mở cửa sổ dòng lệnh)
2. Gõ lệnh sau rồi nhấn Enter:

```bash
brew install python3
```

Nếu báo lỗi `brew: command not found`, cần cài Homebrew trước bằng lệnh:

```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

Cài xong Homebrew, chạy lại `brew install python3`.

Kiểm tra:

```bash
python3 --version
```

---

## Bước 2: Giải nén và cài đặt thư viện

### 2.1. Giải nén file zip

- **Windows:** Click chuột phải vào file zip → chọn **"Extract All..."** → chọn nơi lưu → nhấn **Extract**
- **Mac:** Click đúp vào file zip, nó sẽ tự giải nén

Đặt vào một vị trí dễ nhớ, ví dụ:

- **Windows:** `C:\Users\ten_ban\Desktop\tool_auto\`
- **Mac:** `/Users/ten_ban/Desktop/tool_auto/`

(Thay `ten_ban` bằng tên user máy tính của bạn)

### 2.2. Mở cửa sổ dòng lệnh (Terminal / Command Prompt)

**Windows:**
- Nhấn `Win + R`, gõ `cmd`, nhấn Enter

**Mac:**
- Nhấn `Cmd + Space`, gõ `Terminal`, nhấn Enter

### 2.3. Di chuyển vào thư mục tool

Gõ lệnh `cd` (change directory) để di chuyển vào thư mục vừa giải nén:

**Windows:**

```
cd C:\Users\ten_ban\Desktop\tool_auto
```

**Mac:**

```bash
cd /Users/ten_ban/Desktop/tool_auto
```

> Mẹo: Trên Windows, bạn có thể mở thư mục tool trong File Explorer, gõ `cmd` vào thanh địa chỉ rồi nhấn Enter — sẽ tự mở Command Prompt đúng thư mục.

### 2.4. Tạo môi trường ảo (virtual environment)

Môi trường ảo giúp cài thư viện riêng cho tool, không ảnh hưởng tới máy.

**Windows:**

```
python -m venv .venv
```

**Mac:**

```bash
python3 -m venv .venv
```

### 2.5. Kích hoạt môi trường ảo

**Windows (Command Prompt — cửa sổ đen):**

```
.venv\Scripts\activate
```

**Windows (PowerShell — cửa sổ xanh):**

```
.venv\Scripts\Activate.ps1
```

> Nếu PowerShell báo lỗi "execution policy", gõ lệnh sau rồi thử lại:
> `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned`

**Mac:**

```bash
source .venv/bin/activate
```

**Làm sao biết đã kích hoạt thành công?** Nhìn đầu dòng lệnh, nếu thấy chữ `(.venv)` là OK:

```
(.venv) C:\Users\ten_ban\Desktop\tool_auto>     ← Windows
(.venv) ten_ban@MacBook tool_auto %              ← Mac
```

> **QUAN TRỌNG:** Mỗi lần mở Terminal / Command Prompt mới đều phải kích hoạt lại môi trường ảo (lặp lại bước 2.3 và 2.5) trước khi chạy tool.

### 2.6. Cài thư viện

Gõ lệnh sau (đảm bảo đã kích hoạt môi trường ảo):

```
pip install playwright openpyxl
```

Chờ cài xong (khoảng 1-2 phút).

### 2.7. Cài trình duyệt cho Playwright

```
playwright install chromium
```

Bước này sẽ tải về trình duyệt Chromium (khoảng 150MB). Chỉ cần chạy 1 lần.

**Tới đây đã cài xong. Các bước trên chỉ cần làm 1 lần duy nhất.**

---

## Bước 3: Chuẩn bị dữ liệu

### 3.1. Sửa đường dẫn trong script (QUAN TRỌNG)

Mở file `post_phong_tro_fb.py` bằng **Notepad** (Windows) hoặc **TextEdit** (Mac), tìm phần **CẤU HÌNH** ở gần đầu file (khoảng dòng 20-23), sửa lại đường dẫn cho đúng với máy của bạn:

**Windows** (lưu ý dùng dấu `/` chứ không phải `\`):

```python
EXCEL_PATH = "C:/Users/ten_ban/Desktop/tool_auto/phong_tro.xlsx"
IMAGES_DIR = "C:/Users/ten_ban/Desktop/tool_auto/anh_phong"
AUTH_STATE = "C:/Users/ten_ban/Desktop/tool_auto/fb_auth_state.json"
POSTED_LOG = "C:/Users/ten_ban/Desktop/tool_auto/posted_log.json"
```

**Mac:**

```python
EXCEL_PATH = "/Users/ten_ban/Desktop/tool_auto/phong_tro.xlsx"
IMAGES_DIR = "/Users/ten_ban/Desktop/tool_auto/anh_phong"
AUTH_STATE = "/Users/ten_ban/Desktop/tool_auto/fb_auth_state.json"
POSTED_LOG = "/Users/ten_ban/Desktop/tool_auto/posted_log.json"
```

> Thay `ten_ban` bằng tên user máy tính của bạn. Đường dẫn phải trỏ đúng tới thư mục bạn đã giải nén ở Bước 2.1.

### 3.2. Chuẩn bị file Excel `phong_tro.xlsx`

Mở file `phong_tro.xlsx` bằng Excel hoặc Google Sheets. Sheet phải đặt tên là **PhongTro**. Có 4 cột:

| Cột A | Cột B | Cột C | Cột D |
|-------|-------|-------|-------|
| **STT** | **Mô tả (Caption)** | **Địa chỉ (Khu vực)** | **Folder ảnh** |
| 1 | Phòng đẹp 25m², đầy đủ nội thất... | Hà Đông | phong1 |
| 2 | Cho thuê phòng cao cấp... | Thanh Xuân | phong2 |
| 3 | Phòng giá rẻ gần Lotte... | Cầu Giấy | phong3 |

Giải thích từng cột:
- **Cột A (STT):** Số thứ tự, quyết định đăng bài nào trước (số nhỏ đăng trước)
- **Cột B (Mô tả):** Nội dung nguyên văn sẽ được dùng làm caption bài đăng Facebook. Bạn viết gì ở đây, Facebook sẽ hiện đúng như vậy
- **Cột C (Địa chỉ):** Tên khu vực để tool biết đăng lên group nào. Có thể ghi có dấu hoặc không dấu đều được (ví dụ: "Hà Đông", "Ha Dong", "hadong" đều nhận)
- **Cột D (Folder ảnh):** Tên thư mục con chứa ảnh của phòng đó, đặt trong thư mục `anh_phong/`

> **Hàng 1 là tiêu đề**, dữ liệu bắt đầu từ hàng 2.

### 3.3. Chuẩn bị thư mục ảnh `anh_phong/`

Tạo các thư mục con trong `anh_phong/`, tên phải **trùng với cột D** trong Excel:

```
anh_phong/
  phong1/          ← trùng với cột D hàng 2 trong Excel
    anh1.jpg
    anh2.jpg
    anh3.png
  phong2/          ← trùng với cột D hàng 3 trong Excel
    anh1.jpg
  phong3/
    anh1.jpg
    anh2.webp
```

- Hỗ trợ định dạng ảnh: `.jpg`, `.jpeg`, `.png`, `.webp`
- Tên file ảnh đặt gì cũng được, tool sẽ tự lấy tất cả ảnh trong folder
- Nếu folder không có ảnh hoặc không tồn tại, bài đăng sẽ không có ảnh

### 3.4. Cấu hình group Facebook

Trong file `post_phong_tro_fb.py`, tìm phần `KHU_VUC_GROUPS` và sửa link group thật của bạn:

```python
KHU_VUC_GROUPS = {
    "Hà Đông": [
        "https://www.facebook.com/groups/link_group_1/",
        "https://www.facebook.com/groups/link_group_2/",
    ],
    "Thanh Xuân": [
        "https://www.facebook.com/groups/link_group_3/",
    ],
    # Thêm khu vực khác...
}
```

Cách lấy link group: Vào group Facebook trên trình duyệt → copy URL trên thanh địa chỉ → dán vào.

- Mỗi khu vực có thể có **nhiều group** (phòng sẽ được đăng lên tất cả group của khu vực đó)
- Khu vực nào chưa có group thì để list rỗng `[]`, tool sẽ tự bỏ qua

---

## Bước 4: Chạy tool

### 4.1. Mở Terminal / Command Prompt và kích hoạt môi trường ảo

**Windows:**

```
cd C:\Users\ten_ban\Desktop\tool_auto
.venv\Scripts\activate
```

**Mac:**

```bash
cd /Users/ten_ban/Desktop/tool_auto
source .venv/bin/activate
```

> Nhớ kiểm tra có `(.venv)` ở đầu dòng lệnh trước khi chạy.

### 4.2. Chạy script

**Windows:**

```
python post_phong_tro_fb.py
```

**Mac:**

```bash
python3 post_phong_tro_fb.py
```

### 4.3. Lần đầu chạy — Đăng nhập Facebook

Lần đầu tiên chạy, tool sẽ:
1. Mở ra một cửa sổ trình duyệt Chromium (trình duyệt riêng của tool, **KHÔNG liên quan** đến Chrome, Cốc Cốc hay bất kỳ trình duyệt nào bạn đang dùng)
2. Hiển thị trang đăng nhập Facebook
3. **Bạn tự tay đăng nhập** bằng tài khoản muốn dùng để đăng bài
4. Sau khi đăng nhập xong, quay lại cửa sổ Terminal / Command Prompt và **nhấn Enter**
5. Tool sẽ lưu session, **lần sau không cần đăng nhập lại**

### 4.4. Xác nhận nick

Mỗi lần chạy, tool sẽ hiển thị tên nick đang đăng nhập để bạn kiểm tra:

```
[14:30:04] Đang đăng nhập với nick: Nguyễn Văn A

>>> Đúng nick 'Nguyễn Văn A'? Nhấn Enter để tiếp tục, gõ 'q' để hủy:
```

- **Nhấn Enter** → bắt đầu đăng bài
- **Gõ `q` rồi nhấn Enter** → hủy, không đăng gì cả

### 4.5. Theo dõi tiến trình

Tool sẽ log chi tiết trên Terminal để bạn biết đang làm gì:

```
[14:30:10] PHÒNG 1/3 | STT 1 | Row 2
[14:30:10]   Địa chỉ : Hà Đông
[14:30:10]   Caption  : Phòng đẹp 25m², đầy đủ nội thất...
[14:30:10]   Ảnh      : 3 file
[14:30:10]   Tìm thấy 2 group cho khu vực này
[14:30:10]   ── Group 1/2 (bài 1/5)
[14:30:12]     Đang mở group...
[14:30:15]     Đang nhập nội dung...
[14:30:20]     Đang upload 3 ảnh...
[14:30:28]     Upload ảnh xong!
[14:30:29]     Nhấn nút Đăng...
[14:30:35]     ✓ THÀNH CÔNG!
[14:30:35]   ── Group 2/2 (bài 2/5)
[14:30:37]     Đang mở group...
...

[14:35:00] KẾT QUẢ: 5 thành công | 0 lỗi | 0 bỏ qua | 0 đã đăng trước đó
```

### 4.6. Dừng tool giữa chừng

Nếu muốn dừng tool đang chạy: nhấn `Ctrl + C` trong Terminal / Command Prompt.

Không lo mất dữ liệu — những bài đã đăng thành công đã được ghi vào `posted_log.json`. Khi chạy lại, tool sẽ tự bỏ qua các bài đã đăng và tiếp tục từ bài chưa đăng.

---

## Bước 5 (tùy chọn): Tạo shortcut để chạy nhanh

Thay vì gõ lệnh mỗi lần, bạn có thể tạo file shortcut để click đúp là chạy.

### Windows — Tạo file `chay.bat`

Mở Notepad, dán nội dung sau:

```bat
@echo off
cd /d "%~dp0"
call .venv\Scripts\activate
python post_phong_tro_fb.py
pause
```

Lưu lại với tên `chay.bat` trong thư mục `tool_auto`.

> Lưu ý khi lưu trong Notepad: ở ô **"Save as type"** chọn **All Files (*.*)** để tránh bị lưu thành `chay.bat.txt`.

Lần sau **click đúp vào file `chay.bat`** là chạy được, không cần gõ lệnh.

### Mac — Tạo file `chay.sh`

Mở Terminal, gõ lần lượt:

```bash
cd /Users/ten_ban/Desktop/tool_auto

cat > chay.sh << 'EOF'
#!/bin/bash
cd "$(dirname "$0")"
source .venv/bin/activate
python3 post_phong_tro_fb.py
EOF

chmod +x chay.sh
```

Lần sau gõ `./chay.sh` hoặc click đúp vào file `chay.sh` là chạy được.

---

## Các tính năng khác

### Tránh đăng trùng

Tool tự động ghi nhớ bài đã đăng trong file `posted_log.json`. Nếu chạy lại, những bài đã đăng thành công sẽ được **bỏ qua tự động**.

### Tự động thử lại khi lỗi

Nếu đăng bài bị lỗi (mất mạng, Facebook lag...), tool sẽ tự động thử lại (mặc định 2 lần).

### Đổi nick Facebook

Xóa file `fb_auth_state.json` rồi chạy lại script. Tool sẽ mở trình duyệt để bạn đăng nhập nick mới.

### Giới hạn số bài đăng

Mở file `post_phong_tro_fb.py`, tìm và sửa dòng:

```python
MAX_POSTS = 0         # 0 = đăng hết, 3 = chỉ đăng 3 phòng đầu tiên
```

### Thay đổi tốc độ delay

```python
DELAY_BETWEEN_GROUPS = 10   # Delay (giây) giữa các group
DELAY_BETWEEN_ROOMS = 10    # Delay (giây) giữa các phòng
```

Tăng lên nếu sợ bị Facebook phát hiện, giảm xuống nếu muốn nhanh hơn.

---

## Cấu trúc thư mục

```
tool_auto/
  post_phong_tro_fb.py     # Script chính (code tool)
  phong_tro.xlsx           # File Excel dữ liệu phòng
  anh_phong/               # Thư mục chứa ảnh
    phong1/
      anh1.jpg
    phong2/
      anh1.jpg
  HUONGDAN.md              # File hướng dẫn này
  chay.bat                 # (Tùy chọn) Shortcut chạy nhanh trên Windows
  chay.sh                  # (Tùy chọn) Shortcut chạy nhanh trên Mac
  fb_auth_state.json       # Session Facebook (tự động tạo khi đăng nhập)
  posted_log.json          # Log bài đã đăng (tự động tạo khi chạy)
  .venv/                   # Môi trường ảo Python (tự động tạo khi cài đặt)
```

---

## Xử lý lỗi thường gặp

| Lỗi | Cách xử lý |
|-----|-----------|
| `'python' is not recognized` (Windows) | Cài lại Python, nhớ tick "Add to PATH" |
| `No module named 'playwright'` | Kích hoạt venv trước (`activate`), rồi chạy `pip install playwright` và `playwright install chromium` |
| `No module named 'openpyxl'` | Kích hoạt venv trước, rồi chạy `pip install openpyxl` |
| `No such file or directory` | Kiểm tra đường dẫn trong phần CẤU HÌNH của script có đúng không |
| Trình duyệt mở lên nhưng trang trắng | Kiểm tra kết nối mạng |
| Session hết hạn / bị logout | Xóa `fb_auth_state.json` rồi chạy lại để đăng nhập mới |
| Không tìm thấy group cho khu vực | Kiểm tra cột C trong Excel có khớp với key trong `KHU_VUC_GROUPS` không |
| Facebook checkpoint / khóa tài khoản | Dùng tài khoản phụ, giảm số bài/ngày, tăng `DELAY_BETWEEN_GROUPS` |
| Không tìm thấy ảnh | Kiểm tra tên folder trong cột D có trùng với tên thư mục trong `anh_phong/` không |
| PowerShell báo lỗi "execution policy" | Chạy `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` |
| Tool đăng bài nhưng bị Facebook từ chối | Group có thể yêu cầu duyệt bài — kiểm tra trên Facebook |

---

## Reset / Làm sạch dữ liệu

### Xóa từng file

| File | Khi nào cần xóa | Cách xóa |
|------|-----------------|----------|
| `fb_auth_state.json` | Muốn đổi nick Facebook, hoặc bị logout | Xóa file trong thư mục tool |
| `posted_log.json` | Muốn đăng lại tất cả bài từ đầu | Xóa file trong thư mục tool |
| `.venv/` | Muốn cài lại thư viện từ đầu, hoặc gửi cho người khác | Xóa cả thư mục `.venv` |

### Xóa tất cả để gửi cho người khác

Khi muốn đóng gói tool để gửi zip cho người khác:

**Windows (Command Prompt):**

```
del fb_auth_state.json
del posted_log.json
rmdir /s /q .venv
```

**Mac (Terminal):**

```bash
rm -f fb_auth_state.json posted_log.json
rm -rf .venv
```

Sau khi xóa, chỉ còn lại các file cần thiết:

```
tool_auto/
  post_phong_tro_fb.py     # Script chính
  phong_tro.xlsx           # File Excel dữ liệu phòng
  anh_phong/               # Thư mục chứa ảnh
  HUONGDAN.md              # File hướng dẫn
```

Người nhận chỉ cần làm lại từ **Bước 2.4** (tạo môi trường ảo) trở đi là chạy được.

> **KHÔNG xóa** file `post_phong_tro_fb.py`, `phong_tro.xlsx`, thư mục `anh_phong/`, và `HUONGDAN.md` — đó là code và dữ liệu chính.

---

## Tóm tắt nhanh (dành cho người đã cài xong)

Mỗi lần muốn đăng bài, chỉ cần 3 bước:

**Windows:**

```
cd C:\Users\ten_ban\Desktop\tool_auto
.venv\Scripts\activate
python post_phong_tro_fb.py
```

**Mac:**

```bash
cd /Users/ten_ban/Desktop/tool_auto
source .venv/bin/activate
python3 post_phong_tro_fb.py
```

Hoặc nếu đã tạo shortcut: **click đúp vào `chay.bat` (Windows) hoặc `chay.sh` (Mac)**.

---

## Lưu ý quan trọng

- **Dùng tài khoản phụ** để đăng bài, không dùng tài khoản chính — đề phòng bị checkpoint
- Không nên đăng quá nhiều bài trong 1 ngày (nên dưới 10 bài/ngày)
- Facebook có thể thay đổi giao diện bất cứ lúc nào, khi đó script cần được cập nhật
- Nên **test với 1 phòng + 1 group trước** khi chạy toàn bộ (đặt `MAX_POSTS = 1`)
- Tool dùng trình duyệt Chromium riêng, **không ảnh hưởng** đến Chrome, Cốc Cốc hay trình duyệt bạn đang dùng
