# Hướng dẫn sử dụng — Tool đăng bài Facebook

Dành cho người dùng bản đóng gói sẵn trên Windows. Không cần cài Python,
không cần cài trình duyệt, không cần biết lập trình.

---

## 1. Cài đặt

1. Giải nén file `ToolDangBaiFacebook-Windows.zip` ra một thư mục bất kỳ —
   ví dụ `D:\ToolFacebook`.
2. Mở thư mục vừa giải nén, bấm đúp vào **`ToolDangBaiFacebook.exe`**.

Lần đầu mở sẽ hơi lâu (khoảng 10–30 giây) vì máy quét virus kiểm tra file mới.
Những lần sau nhanh hơn nhiều.

> **Windows hiện bảng xanh "Windows protected your PC"?**
> Đây là cảnh báo mặc định với mọi phần mềm chưa mua chứng chỉ ký số, không
> phải máy bị nhiễm gì. Bấm **More info** → **Run anyway**.

> **Đừng để tool trong thư mục `C:\Program Files`.** Windows chặn ghi file ở
> đó, tool sẽ không lưu được cấu hình và phiên đăng nhập. Để ở ổ D hoặc trong
> Desktop/Documents là tốt nhất.

Lần chạy đầu tiên tool tự tạo sẵn cạnh file .exe:

| Tên | Là gì |
|---|---|
| `config.json` | Cấu hình: đường dẫn, danh sách nhóm, thời gian chờ |
| `phong_tro.xlsx` | File Excel mẫu để bạn điền nội dung bài đăng |
| `anh_phong\` | Thư mục chứa ảnh |
| `fb_profiles\` | Phiên đăng nhập Facebook (mỗi nick một thư mục con) |

---

## 2. Chuẩn bị nội dung

### File Excel

Mở `phong_tro.xlsx`, xóa 2 dòng ví dụ rồi điền dữ liệu thật. Đúng 4 cột:

| Cột | Tên | Nội dung |
|---|---|---|
| A | STT | Số thứ tự: 1, 2, 3... |
| B | Nội dung bài | Caption đăng lên Facebook. Xuống dòng thoải mái, đăng lên giữ nguyên |
| C | Phân loại | Quyết định bài này đăng lên nhóm nào. Phải trùng tên phân loại khai ở tab "Nhóm theo phân loại" |
| D | Folder ảnh | Tên thư mục con trong `anh_phong\`. Để trống nếu bài không cần ảnh |

**Lưu và đóng Excel trước khi bấm chạy.** Excel khóa file khi đang mở, tool sẽ
không đọc được.

### Ảnh

Mỗi bài một thư mục con riêng trong `anh_phong\`. Ví dụ bài ở cột D ghi
`phong_501` thì ảnh để ở `anh_phong\phong_501\`. Bỏ bao nhiêu ảnh cũng được.

---

## 3. Cấu hình trong tool

Giao diện có 4 tab:

- **Tài khoản & Chạy** — thêm nick Facebook, bấm chạy, xem log
- **Nội dung bài đăng** — xem/sửa nội dung Excel ngay trong tool
- **Dữ liệu chiến dịch** — đường dẫn file, thời gian chờ
- **Nhóm theo phân loại** — khai danh sách link nhóm cho từng phân loại

### Thêm nick Facebook

1. Vào tab **Tài khoản & Chạy** → bấm **Thêm nick**
2. Đặt tên gợi nhớ (ví dụ `nick_chinh`) → một cửa sổ trình duyệt mở ra
3. **Đăng nhập Facebook trong cửa sổ đó**, làm hết các bước xác minh nếu có
4. Đăng nhập xong thì đóng cửa sổ

Phiên đăng nhập được lưu lại, lần sau không phải đăng nhập lại nữa.

### Khai danh sách nhóm

Sang tab **Nhóm theo phân loại**. Xóa 2 phân loại mẫu đi, thêm phân loại của
bạn, rồi dán link nhóm vào — mỗi dòng một link, dạng
`https://web.facebook.com/groups/...`.

Tên phân loại ở đây **phải trùng chính xác** với cột C trong Excel.

---

## 4. Chạy

Tab **Tài khoản & Chạy** → chọn nick → bấm **Bắt đầu đăng**.

Log hiện trực tiếp trong tool. Muốn dừng giữa chừng thì bấm **Dừng** — tool
đăng nốt bài đang làm rồi mới dừng, không cắt ngang.

Bài nào đã đăng được ghi vào `posted_log.json` nên chạy lại sẽ không đăng
trùng. Muốn đăng lại từ đầu thì xóa file đó đi.

---

## 5. Hạn chế rủi ro khóa tài khoản

Facebook chống công cụ tự động. Đây là rủi ro có thật, không thể loại bỏ hoàn
toàn — chỉ giảm được.

**Dùng tài khoản phụ, không dùng tài khoản chính.**

Những điều quan trọng nhất, xếp theo mức ảnh hưởng:

1. **Tốc độ đăng** — đây là nguyên nhân khóa nick nhiều nhất. Để thời gian chờ
   ít nhất 60–120 giây giữa các nhóm trong tab "Dữ liệu chiến dịch". Đăng 20
   nhóm trong 10 phút gần như chắc chắn dính.
2. **Nội dung lặp lại** — đừng copy-paste y hệt sang mọi nhóm. Đổi câu chữ,
   đổi ảnh, đổi thứ tự.
3. **Nick mới cần làm nóng** — nick vừa lập đừng đăng ngay. Vài ngày đầu chỉ
   like, bình luận, kết bạn, đăng 1–2 bài. Tăng dần.
4. **Nhiều nick trên cùng một máy** — mỗi nick phải là một profile riêng trong
   tool (thêm bằng nút "Thêm nick", tool tự tách). Đừng đăng 2 nick cùng lúc,
   giãn ra vài giờ. Nếu làm nghiêm túc thì cho mỗi nick một mạng khác nhau —
   một nick wifi, một nick 4G điện thoại.

---

## 6. Lỗi thường gặp

| Hiện tượng | Cách xử lý |
|---|---|
| `⚠ Chưa có file: ...xlsx` | File Excel bị xóa hoặc đổi chỗ. Vào tab "Dữ liệu chiến dịch" chỉ lại đường dẫn |
| Không đọc được Excel | Đang mở file đó trong Excel. Đóng lại rồi bấm nạp lại |
| Mở tool báo nhóm không đăng được | Nick chưa vào nhóm đó, hoặc nhóm bắt duyệt bài. Kiểm tra bằng tay |
| Bị đòi đăng nhập lại liên tục | Thư mục `fb_profiles\` bị xóa, hoặc tool đang nằm trong `C:\Program Files` (không ghi được) |
| Đăng vài bài rồi dừng, báo checkpoint | Facebook đã cảnh báo nick. **Dừng ngay**, đăng nhập bằng tay giải quyết, nghỉ vài ngày rồi tăng thời gian chờ lên |

### Chuyển tool sang máy khác

Copy nguyên cả thư mục. Trong đó đã có sẵn cấu hình, Excel, ảnh, và cả phiên
đăng nhập — sang máy mới chạy được luôn.

Nhưng lưu ý phần 5: chuyển nick sang máy khác là đổi thiết bị đột ngột dưới
mắt Facebook. Nếu đổi cả máy lẫn mạng cùng lúc thì khả năng bị hỏi xác minh
khá cao.
