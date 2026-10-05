# Hướng dẫn sử dụng — Tool đăng bài Facebook

Tài liệu dành cho người dùng bản đóng gói sẵn (`ToolDangBaiFacebook.exe`) trên
Windows. Không cần cài Python, không cần biết lập trình.

**Mục lục**
1. [Cần chuẩn bị](#1-cần-chuẩn-bị)
2. [Mở tool lần đầu](#2-mở-tool-lần-đầu)
3. [Làm quen giao diện](#3-làm-quen-giao-diện)
4. [Bước 1 — Thêm nick Facebook](#4-bước-1--thêm-nick-facebook)
5. [Bước 2 — Khai danh sách nhóm](#5-bước-2--khai-danh-sách-nhóm)
6. [Bước 3 — Soạn nội dung bài](#6-bước-3--soạn-nội-dung-bài)
7. [Bước 4 — Chuẩn bị ảnh](#7-bước-4--chuẩn-bị-ảnh)
8. [Bước 5 — Kiểm tra và chạy](#8-bước-5--kiểm-tra-và-chạy)
9. [Tham gia nhóm tự động](#9-tham-gia-nhóm-tự-động)
10. [Hẹn giờ đăng tự động](#10-hẹn-giờ-đăng-tự-động)
11. [Hạn chế rủi ro khóa tài khoản](#11-hạn-chế-rủi-ro-khóa-tài-khoản)
12. [Lỗi thường gặp](#12-lỗi-thường-gặp)
13. [Tách dữ liệu ra thư mục riêng](#13-tách-dữ-liệu-ra-thư-mục-riêng--nên-làm-ngay-từ-đầu)
14. [Nhận bản cập nhật](#14-nhận-bản-cập-nhật)
15. [Sao lưu và chuyển máy](#15-sao-lưu-và-chuyển-máy)

---

## 1. Cần chuẩn bị

**Google Chrome** — máy phải có sẵn. Tool điều khiển Chrome thật để đăng bài;
dùng trình duyệt khác không được, Facebook sẽ chặn đăng nhập.

Phần lớn máy đã có Chrome. Nếu chưa, tải miễn phí tại
https://www.google.com/chrome, cài xong là được, không cần chỉnh gì.

Ngoài Chrome ra **không cần cài thêm gì** — mọi thứ khác đã nằm trong gói.

---

## 2. Mở tool lần đầu

1. Giải nén file zip ra một thư mục. **Nên để ở ổ D hoặc trong Documents.**
   > ⚠ Đừng để trong `C:\Program Files` — Windows chặn ghi ở đó, tool sẽ không
   > lưu được cấu hình và phiên đăng nhập.
2. Mở thư mục, bấm đúp **`ToolDangBaiFacebook.exe`**.

Lần đầu mở chờ 10–30 giây (Windows quét file mới), những lần sau nhanh hơn.

> **Hiện bảng xanh "Windows protected your PC"?**
> Cảnh báo mặc định với phần mềm chưa mua chứng chỉ ký số, không phải máy nhiễm
> virus. Bấm **More info** → **Run anyway**.

Ngay lần chạy đầu, tool tự tạo sẵn cạnh file .exe những thứ sau — **không cần
đụng tay**:

| Tên | Là gì |
|---|---|
| `config.json` | Toàn bộ cấu hình (đường dẫn, nhóm, thời gian chờ) |
| `phong_tro.xlsx` | File Excel mẫu để điền nội dung bài |
| `anh_phong\` | Thư mục chứa ảnh |
| `fb_profiles\` | Phiên đăng nhập từng nick |

> 💡 Bốn thứ dưới cùng bảng là **dữ liệu của bạn**, không phải của tool. Nên
> chuyển chúng sang một thư mục riêng để sau này nhận bản cập nhật không bị mất
> — xem [mục 13](#13-tách-dữ-liệu-ra-thư-mục-riêng--nên-làm-ngay-từ-đầu). Làm
> ngay bây giờ cũng được, mà dùng quen rồi làm cũng không sao.

---

## 3. Làm quen giao diện

Tool có 6 tab ở trên cùng:

| Tab | Dùng để |
|---|---|
| **Tài khoản & Chạy** | Thêm nick, bấm chạy, xem log trực tiếp |
| **Nội dung bài đăng** | Xem và sửa nội dung bài ngay trong tool |
| **Dữ liệu chiến dịch** | Đường dẫn file, chỉnh thời gian chờ giữa các bài/nhóm |
| **Nhóm theo phân loại** | Khai danh sách link nhóm cho từng phân loại |
| **⏰ Hẹn giờ đăng** | Đặt lịch để tool tự đăng đúng giờ (xem mục 10) |
| **👥 Tham gia nhóm** | Cho nick vào nhóm — chạy riêng, không dính việc đăng bài (xem mục 9) |

Góc trên có ô **Chiến dịch** — mỗi chiến dịch (phòng trọ, tuyển dụng,
seeding...) có nội dung, ảnh, nhóm riêng, nhưng dùng chung các nick Facebook.
Mới bắt đầu thì cứ dùng chiến dịch "Phòng trọ" có sẵn.

Cạnh ô đó là **+ Thêm** (tạo chiến dịch mới), **Đổi tên**, **Xóa**. Ngoài cùng
bên phải là **📖 Hướng dẫn sử dụng** — mở đúng tài liệu bạn đang đọc, ngay
trong tool, không phải đóng tool để đi tìm file.

Dưới đây đi qua 3 tab dùng nhiều nhất. Ba tab còn lại có mục riêng: Dữ liệu
chiến dịch nằm rải trong các bước, hẹn giờ ở mục 10, tham gia nhóm ở mục 9.

### Tab Tài khoản & Chạy

Đây là nơi bấm chạy và theo dõi tool làm việc.

![Tab Tài khoản & Chạy](anh_huong_dan/tai_khoan_va_chay.png)

**Cột trái — danh sách nick.** Mỗi dòng là một tài khoản Facebook đã đăng nhập.
Nick dùng chung cho mọi chiến dịch, khai một lần là xong.

- **+ Thêm nick mới** — mở Chrome cho bạn đăng nhập bằng tay, xong bấm nút xác
  nhận là tool nhớ luôn, lần sau không phải đăng nhập lại. **Nick phải bật xác
  thực 2 lớp trước đã** — xem mục 4.
- **Đăng nhập lại nick này** — dùng khi nick bị đòi xác minh hoặc hết phiên.
- **Xóa nick** — bỏ nick khỏi danh sách và xóa profile Chrome của nó.

**Thanh nút trên cùng.**

| Nút | Làm gì |
|---|---|
| **▶ Bắt đầu đăng** | Chạy chiến dịch đang chọn bằng nick đang chọn |
| **■ Dừng** | Dừng sau khi xong việc đang làm dở, không cắt ngang giữa chừng |
| **🔍 Kiểm tra dữ liệu** | Soát trước khi chạy: thiếu file, thiếu ảnh, phân loại sai tên |
| **Xóa màn hình log** | Chỉ dọn chữ đang hiện, không đụng dữ liệu |
| **🗑 Xóa lịch sử đã đăng** | Xóa file chống đăng trùng — **cẩn thận**, xóa xong lần chạy tới đăng lại từ đầu |

**Ô log bên dưới** chạy chữ theo thời gian thực. Màu chữ cho biết loại: đỏ là
lỗi, cam là cảnh báo, xanh lá là xong việc, xám là thao tác phụ.

> ⚠ Luôn bấm **🔍 Kiểm tra dữ liệu** trước khi chạy lần đầu trong ngày. Nó bắt
> hết lỗi khai báo trong vài giây, đỡ phải phát hiện giữa chừng khi tool đã
> đăng được nửa chừng.

### Tab Nội dung bài đăng

Bảng này chính là file Excel, mở ngay trong tool. Mỗi dòng một bài.

![Tab Nội dung bài đăng](anh_huong_dan/noi_dung_bai_dang.png)

Bốn cột: **STT** (mã định danh bài), **Nội dung bài**, **Phân loại** (bài đăng
lên nhóm nào), **Thư mục ảnh**.

| Nút | Làm gì |
|---|---|
| **+ Thêm bài** | Thêm một dòng mới |
| **Xóa bài** | Bỏ dòng khỏi bảng (không xóa ảnh) |
| **🗑 Xóa toàn bộ ảnh** | Dọn sạch ảnh **và video .mp4** trong thư mục ảnh của bài đang chọn |
| **🖼 Ảnh của bài này** | Mở cửa sổ xem/thêm/xóa từng ảnh |
| **⟳ Tải lại từ Excel** | Đọc lại file, bỏ mọi sửa chưa ghi |
| **💾 Ghi vào file Excel** | **Bắt buộc bấm sau khi sửa**, không thì mất |

**Sửa một bài: bấm đúp vào dòng đó.** Cửa sổ soạn bài mở ra, gõ xuống dòng
thoải mái — Facebook giữ nguyên cách xuống dòng. Trong đó có nút
**. Chấm câu cuối dòng** thêm dấu chấm vào cuối mọi dòng còn thiếu.

> ⚠ **STT là mã định danh bài.** File chống đăng trùng ghi theo STT, nên đổi
> STT của một bài cũ sẽ làm tool tưởng là bài khác và đăng lại, hoặc tưởng đã
> đăng rồi mà bỏ qua.

### Tab Nhóm theo phân loại

Nơi khai link group. Chia hai cột.

![Tab Nhóm theo phân loại](anh_huong_dan/nhom_theo_phan_loai.png)

**Cột trái — danh sách phân loại.** "Phân loại" là cách gom nhóm lại: với phòng
trọ thì mỗi quận một phân loại. Có **+ Thêm phân loại**, **Đổi tên**,
**Xóa phân loại**, và mục **Chuyển sang máy khác** với 2 nút xuất/nhập (xem
mục 5).

**Cột phải — link group.** Mỗi dòng một link, dạng
`https://www.facebook.com/groups/...`. Sửa xong bấm **💾 Lưu danh sách group**.

Cách nó ăn khớp với bảng nội dung: bài nào có cột **Phân loại** ghi `Cầu Giấy`
thì được đăng lên đúng những link nằm trong phân loại `Cầu Giấy`. Tên phải
trùng nhau — tool có bỏ dấu và không phân biệt hoa thường khi so, nhưng viết
sai chữ là bài đó không đăng đi đâu cả.

---

## 4. Bước 1 — Thêm nick Facebook

> ⚠ **BẮT BUỘC: nick phải bật xác thực 2 lớp trước khi thêm vào tool.**
> Chưa bật thì đừng thêm nick đó vào.

**Vì sao bắt buộc:** nick chạy tool là nick Facebook để ý tới nhiều hơn bình
thường — đăng đều, đăng nhiều nhóm, đăng từ máy chạy tự động. Chỉ có mật khẩu
thì mất nick là mất luôn: người lấy được mật khẩu đăng nhập từ bất kỳ đâu cũng
vào được. Bật 2 lớp rồi thì dù lộ mật khẩu, không có mã trên điện thoại bạn vẫn
không ai vào được.

Thêm một lý do thực tế: Facebook tin nick có 2 lớp hơn. Nick bật 2 lớp ít bị
hỏi xác minh vặt hơn, và khi bị khoá tạm thì mở lại dễ hơn hẳn.

### Bật xác thực 2 lớp trước khi làm tiếp

1. Mở Facebook trên điện thoại hoặc máy tính, vào **Cài đặt & quyền riêng tư →
   Cài đặt → Trung tâm tài khoản → Mật khẩu và bảo mật → Xác thực 2 yếu tố**.
2. Chọn nick cần bật, chọn cách nhận mã.
3. **Nên dùng ứng dụng tạo mã** (Google Authenticator, Microsoft Authenticator)
   thay vì tin nhắn SMS — SMS bị tráo SIM là mất, mà đi nước ngoài hay đổi số
   là cũng không nhận được mã.
4. **Lưu lại mã khôi phục** Facebook đưa cho, cất chỗ khác máy chạy tool. Mất
   điện thoại mà không có mã khôi phục là mất nick.

Bật xong mới làm tiếp phần dưới.

### Các bước thêm nick vào tool

1. Vào tab **Tài khoản & Chạy**, bấm **+ Thêm nick mới**.
2. Đặt tên gợi nhớ cho nick (ví dụ `nick_phu_1`) rồi OK.
3. Một cửa sổ **Chrome** tự mở ra ở trang đăng nhập Facebook.
4. **Đăng nhập Facebook ngay trong cửa sổ Chrome đó** — nhập tài khoản, mật
   khẩu, làm hết các bước xác minh (mã OTP, v.v.) nếu Facebook yêu cầu.
5. Facebook hỏi mã xác thực 2 lớp — mở ứng dụng tạo mã, nhập vào. Có ô
   **"Ghi nhớ trình duyệt này"** thì tick, để lần sau không bị hỏi lại.
6. Đăng nhập xong, **quay lại tool** và bấm nút **"Tôi đã đăng nhập xong"**.

> Điểm hay bị nhầm: bấm nút "Tôi đã đăng nhập xong" **trong tool**, chứ không
> phải chỉ đóng cửa sổ Chrome. Đóng Chrome mà chưa bấm nút thì phiên đăng nhập
> có thể chưa được lưu.

Xong, phiên được lưu vào `fb_profiles\`. **Lần sau không phải đăng nhập lại** —
mã 2 lớp chỉ phải nhập lần đầu này thôi, không phải mỗi lần chạy tool.

Thêm nhiều nick thì lặp lại các bước trên. Nick nào lỡ bị đăng xuất thì chọn
nick đó trong danh sách rồi bấm **"Đăng nhập lại nick này"**.

---

## 5. Bước 2 — Khai danh sách nhóm

Sang tab **Nhóm theo phân loại**.

"Phân loại" là cách nhóm các group lại với nhau. Ví dụ với phòng trọ, mỗi quận
là một phân loại ("Cầu Giấy", "Đống Đa"...), mỗi phân loại chứa danh sách link
group của quận đó.

1. Xóa 2 phân loại mẫu đi (chọn rồi bấm **Xóa phân loại**).
2. Bấm **+ Thêm phân loại**, đặt tên (ví dụ `Cầu Giấy`).
3. Ở ô **Link group (mỗi dòng 1 link)** bên phải, dán link vào — **mỗi dòng
   một link**, dạng `https://web.facebook.com/groups/...`.
4. Bấm **💾 Lưu danh sách group**.

> ⚠ Nhớ tên phân loại — bước sau phải điền **trùng chính xác** tên này vào
> Excel thì bài mới biết đăng lên đúng nhóm.

### Mang danh sách nhóm sang máy khác

Khai được vài chục nhóm rồi mà phải gõ lại trên máy mới thì rất mất công. Cột
bên trái có mục **Chuyển sang máy khác** với 2 nút:

- **⬆ Xuất ra file** — ghi toàn bộ phân loại + link group của chiến dịch đang
  chọn ra một file `.json` nhỏ. Chép file đó sang máy mới bằng USB, Zalo, Drive
  đều được.
- **⬇ Nhập từ file** — trên máy mới, chọn file đó. Tool hỏi **Gộp thêm** (giữ
  nguyên nhóm đang có, chỉ thêm cái mới, bỏ link trùng) hay **Thay hẳn** (xóa
  sạch phân loại hiện tại rồi dùng y hệt file).

File xuất ra **chỉ chứa nhóm theo phân loại**. Cố ý không mang theo:

| Không có trong file | Vì sao |
|---|---|
| Đường dẫn file Excel, thư mục ảnh | Là đường dẫn đầy đủ của máy cũ, sang máy mới là sai hết. Máy mới giữ đường dẫn của nó |
| Nick Facebook | Nick là cả một profile Chrome vài GB. Chép sang máy khác Facebook thấy đổi thiết bị, rất dễ bắt xác minh — nên đăng nhập lại bằng tay |
| Lịch hẹn giờ | Mỗi lịch trỏ tới nick theo tên; máy mới chưa có nick đó thì tới giờ chạy sẽ lỗi |

> ⚠ Những thứ phải **chép tay** khi đổi máy, không nút nào làm thay: file
> `.xlsx`, các thư mục ảnh, và **các file `posted_log*.json`**. Quên file
> posted_log là tool coi như chưa đăng gì, đăng lại từ đầu lên toàn bộ nhóm đã
> đăng rồi.

---

## 6. Bước 3 — Soạn nội dung bài

Có 2 cách, chọn 1:

**Cách A — sửa ngay trong tool** (dễ hơn): vào tab **Nội dung bài đăng**, bấm
**+ Thêm bài** để thêm, **bấm đúp vào dòng** để sửa bài đó. Sửa xong **bắt buộc
bấm 💾 Ghi vào file Excel**, không thì mất.

Trong cửa sổ soạn bài có nút **. Chấm câu cuối dòng**: bấm một phát là thêm dấu
chấm vào cuối mọi dòng còn thiếu. Nó bỏ qua dòng trống, dòng đã có dấu câu, và
dòng không có chữ (hàng emoji, gạch phân cách). Bấm nhầm thì `Ctrl+Z` một lần
là về như cũ.

**Cách B — mở file Excel:** mở `phong_tro.xlsx`, điền theo đúng 4 cột dưới,
lưu lại. Trong tool bấm **⟳ Tải lại từ Excel** để cập nhật.

| Cột | Tên | Nội dung |
|---|---|---|
| A | STT | Số thứ tự: 1, 2, 3... |
| B | Nội dung bài | Nội dung đăng lên Facebook. Xuống dòng thoải mái, đăng lên giữ nguyên |
| C | Phân loại | Bài này đăng lên nhóm nào. **Phải trùng tên phân loại** đã khai ở Bước 2 |
| D | Folder ảnh | Tên thư mục ảnh của bài (xem Bước 4). Để trống nếu bài không cần ảnh |

> ⚠ Nếu dùng cách B: **lưu và đóng Excel trước khi chạy.** Excel đang mở sẽ
> khóa file, tool không đọc được.

---

## 7. Bước 4 — Chuẩn bị ảnh

Mỗi bài để ảnh trong một thư mục con riêng bên trong `anh_phong\`.

Ví dụ: bài có cột D ghi `phong_501` → tạo thư mục `anh_phong\phong_501\` rồi
bỏ ảnh vào đó. Bao nhiêu ảnh cũng được.

Bài không cần ảnh thì để trống cột D, bỏ qua bước này.

### Thay ảnh nhanh ngay trong tool

Không phải mở Explorer đi tìm thư mục nữa. Vào tab **Nội dung bài đăng**, chọn
bài, bấm **🖼 Ảnh của bài này**:

| Nút | Làm gì |
|---|---|
| **+ Thêm ảnh...** | Chọn ảnh từ máy, tool tự chép vào đúng thư mục |
| **Xóa ảnh đã chọn** | Giữ Ctrl chọn nhiều ảnh rồi xóa. Bấm đúp để xem trước |
| **Xóa tất cả ảnh** | Dọn sạch thư mục cho đợt ảnh mới |
| **Mở thư mục** | Mở Explorer đúng thư mục đó, để kéo thả ảnh từ Zalo vào |

### Dọn sạch ảnh của một bài trong 1 nút

Ngay trên thanh của tab **Nội dung bài đăng** có **🗑 Xóa toàn bộ ảnh**: chọn
bài rồi bấm, nó xóa sạch ảnh **và cả video `.mp4`** trong thư mục ảnh của bài
đó — không phải mở cửa sổ quản lý ảnh. Tiện cho kiểu thay ảnh theo đợt: xóa hết
→ bỏ ảnh mới vào → chạy.

Khác với **Xóa tất cả ảnh** trong cửa sổ quản lý ảnh ở chỗ nút kia chỉ xóa ảnh,
để lại video. Video nằm lại chỉ chiếm chỗ vì tool không đăng được video.

Trước khi xóa, tool hỏi lại và ghi rõ sẽ mất bao nhiêu ảnh, bao nhiêu video.

**Xóa ảnh không bao giờ xóa thư mục** — thư mục vẫn nằm đó chờ bạn bỏ ảnh mới
vào, nên cột D trong Excel không cần sửa gì cả. Đúng quy trình thay ảnh theo
đợt: xóa tất cả → lưu ảnh mới từ Zalo vào → chạy.

> ⚠ Hai điều lưu ý:
> - Xóa ảnh ở đây là **xóa hẳn khỏi ổ đĩa**, không vào Thùng rác, không hoàn tác.
> - Ảnh `.heic` (ảnh iPhone, hay lọt vào khi lưu từ Zalo) **không đăng được**.
>   Cửa sổ này sẽ báo màu cam nếu thư mục có file như vậy. Chỉ `.jpg`, `.jpeg`,
>   `.png`, `.webp` mới dùng được — đổi đuôi ảnh trước khi bỏ vào.

---

## 8. Bước 5 — Kiểm tra và chạy

1. Về tab **Tài khoản & Chạy**.
2. Bấm **🔍 Kiểm tra dữ liệu** trước — tool soát xem có bài nào trỏ tới phân
   loại chưa khai, thiếu ảnh, v.v. Có lỗi thì xem ô log, sửa rồi kiểm tra lại.
3. **Chọn nick** muốn đăng trong danh sách.
4. Bấm **▶ Bắt đầu đăng**.

Log chạy hiện trực tiếp trong tool. Muốn dừng thì bấm **■ Dừng** — tool đăng
nốt bài đang làm rồi mới dừng, không cắt ngang giữa chừng.

### Chạy lần 2 trở đi — tool tự tránh đăng trùng **trong ngày**

Mỗi lượt đăng thành công được ghi nhớ theo cặp **STT của bài + link nhóm** kèm
ngày đăng. Chạy lại **trong cùng ngày**, tool gặp lại cặp đó sẽ bỏ qua ngay,
không mở nhóm ra, cũng không mất thời gian chờ:

```
  ── Group 3/6 (bài 15/40)
  ⏭ ĐÃ ĐĂNG HÔM NAY — bỏ qua (chống đăng trùng đang bật)
```

Cuối buổi có dòng tổng kết đếm riêng, kèm hộp thông báo nhắc cách xử lý:

```
KẾT QUẢ: 12 thành công | 0 lỗi | 1 bỏ qua | 27 đã đăng hôm nay
```

**Sang ngày mới thì mọi bài lại đăng được bình thường** — không cần xóa gì cả.
Đây chính là cách để đăng lại định kỳ cho tin nổi lên: cứ hẹn giờ chạy hằng
ngày, tool tự lo phần còn lại.

Ngày tính theo **giờ Việt Nam**, chuyển sang ngày mới lúc 0h00, không phải "đủ
24 tiếng". Bài đăng 23h50 tối nay thì 0h10 sáng mai đã đăng lại được.

**Hai chỗ hay làm người dùng ngạc nhiên:**

- **Sửa nội dung mà giữ nguyên STT** → trong ngày hôm đó tool vẫn coi là bài cũ,
  **không đăng lại**. Cần đăng bản mới ngay thì cho nó một STT mới.
- **Đổi link nhóm** dù chỉ thêm/bớt dấu `/` ở cuối, hay đổi `facebook.com` thành
  `web.facebook.com` → thành link khác trong mắt tool → nó **đăng lại lên đúng
  nhóm cũ**. Khai link nhóm rồi thì đừng sửa vặt.

### Muốn đăng lại ngay trong hôm nay

Hai cách, chọn một:

**Cách 1 — Bấm 🗑 Xóa lịch sử đã đăng.** Cửa sổ hiện ra liệt kê **từng chiến
dịch** kèm số lượt đã đăng của nó; chọn chiến dịch muốn xóa (giữ Ctrl để chọn
nhiều), tool hỏi xác nhận lần nữa có nêu rõ tổng số lượt rồi mới xóa.

Mỗi chiến dịch có một file lịch sử **riêng biệt** — xóa của chiến dịch này không
ảnh hưởng chiến dịch kia. Sau khi xóa, lần chạy tới đăng lại toàn bộ bài của
chiến dịch đó lên mọi nhóm, **kể cả bài vừa lên sáng nay**. Không hoàn tác được.

**Cách 2 — Tắt ô "Chống đăng trùng"** ở tab **Dữ liệu chiến dịch**. Tool sẽ đăng
bất chấp lịch sử: chạy bao nhiêu lần thì cùng một bài lên cùng một nhóm bấy
nhiêu lần.

> ⚠ Cách 2 **chỉ nên tắt tạm**, xong nhớ bật lại. Đăng lặp cùng nội dung lên
> cùng nhóm là hành vi Facebook đánh dấu spam rất nhanh, và nhiều nhóm cũng có
> luật cấm. Để nguyên chống đăng trùng thì bạn được bảo vệ khỏi việc lỡ bấm chạy
> hai lần.

> Đừng nhầm với nút **Xóa màn hình log** ngay bên cạnh — nút đó chỉ dọn chữ hiển
> thị, không đụng gì tới lịch sử đăng.

---

## 9. Tham gia nhóm tự động

Nick chỉ đăng được vào nhóm mà nó **đã là thành viên**. Việc cho nick vào nhóm
nằm ở tab riêng **👥 Tham gia nhóm**, chạy tách hẳn khỏi việc đăng bài: nút
riêng, danh sách nhóm riêng, log riêng. Bấm đăng bài thì tool không bao giờ tự
đi xin vào nhóm, và ngược lại.

**Các bước:**

1. Chọn **nick** ở cột trái — lịch sử tham gia nhóm ghi theo từng nick, nên
   bảng bên phải hiện đúng trạng thái của nick đó.
2. Bấm **+ Dán link nhóm** rồi dán danh sách, mỗi dòng một link. Đã khai nhóm ở
   tab "Nhóm theo phân loại" rồi thì bấm **Nạp từ chiến dịch...** cho khỏi gõ
   lại — nó chỉ **chép sang một lần**, sau đó hai bên độc lập.
3. Bấm **▶ Tham gia tất cả nhóm chưa vào** để chạy cả danh sách — tool tự bỏ
   qua nhóm đã vào rồi.

Muốn nhắm riêng vài nhóm thì **tick chúng lại** rồi bấm **▶ Tham gia nhóm đã
tick**:

| Thao tác | Kết quả |
|---|---|
| Bấm vào **bất kỳ đâu trên dòng** | Tick / bỏ tick dòng đó |
| Bôi đen nhiều dòng rồi nhấn **Space** | Tick cả loạt |
| Bấm chữ **✓ trên tiêu đề cột** | Tick hết, bấm lần nữa thì bỏ hết |
| Nút **☑ Tick hết** / **☐ Bỏ tick** | Như trên |

Ô tick giữ nguyên khi bạn đổi nick hay khi bảng vẽ lại, chỉ mất khi tắt tool.

**Lấy link ra dùng:** bấm chọn dòng rồi bấm **📋 Sao chép link** (hoặc Ctrl+C /
Cmd+C, hoặc chuột phải → *Sao chép link*). Chọn nhiều dòng thì chép nhiều link,
mỗi link một dòng. Chuột phải còn có *Mở nhóm trong trình duyệt* để xem thử
nhóm đó bằng tay.

Nút **Xóa nhóm đã tick** chạy theo mấy ô ✓ này (chỉ xóa khỏi danh sách, không
rời nhóm trên Facebook).

Cột "Trạng thái" cho biết với nick đang chọn thì nhóm đó *Chưa xử lý / Đã là
thành viên / Đang chờ duyệt / Không vào được / Bấm rồi nhưng chưa rõ*. Nhóm đã
xong thì lần chạy sau tool không mở lại nữa; nhóm hỏng thì hẹn ngày thử lại.

Dòng chữ cạnh các nút cho biết còn bao nhiêu nhóm cần xử lý, hôm nay nick đã
bấm mấy lượt, và báo đỏ nếu nick đang bị tạm nghỉ hoặc đang ngoài khung giờ.

### ⚠ Việc này rủi ro hơn đăng bài

Nói thẳng: **tự động tham gia nhóm dễ bị Facebook chặn hơn tự động đăng bài.**
Facebook đếm số lượt xin vào nhóm theo từng nick, theo ngày, và nhìn cả nhịp
thao tác. Không có cách nào làm việc này mà an toàn tuyệt đối — chỉ giảm rủi ro
được thôi.

Vì thế tool cố tình chạy **chậm**, và mặc định đã là mức nên dùng:

| Cài đặt | Mặc định | Chỉnh được tới |
|---|---|---|
| Tối đa mỗi lần chạy | 5 nhóm | 15 |
| Tối đa mỗi nick mỗi ngày | 8 nhóm | 20 |
| Nghỉ giữa 2 nhóm | 90–300 giây | không xuống dưới 40s |
| Nghỉ dài | 20 phút sau mỗi 3 nhóm | ít nhất 5 phút |
| Khung giờ chạy | 8h–22h giờ VN | 0h–24h |

> Cột bên phải là **giới hạn kỹ thuật, không phải mức an toàn.** Càng chạy gần
> mức đó thì nhịp thao tác càng dày và càng dễ bị chặn. Nick đang dùng để đăng
> bài thật thì nên giữ quanh 10 nhóm/lần, 12 nhóm/ngày, nghỉ từ 60s trở lên.

Trần ngày tính **cả các lần chạy trước trong ngày**, nên bấm chạy lại nhiều lần
không lách được. Đặt số ngoài khoảng an toàn thì tool tự kéo về và báo cho bạn.

Tick ô **"Nick mới lập"** cho nick vừa tạo: mọi con số bị ép xuống còn 2
nhóm/lần, 3 nhóm/ngày, nghỉ 3–10 phút — và không nâng lên được cho tới khi bạn
bỏ tick.

**Ngoài mấy con số trên, tool còn:**

- **Xáo thứ tự nhóm** mỗi lần chạy, không đi tuần tự theo danh sách.
- **Xem trang 3–7 giây, cuộn vài nhịp** rồi mới bấm Tham gia.
- **Thỉnh thoảng ghé bảng tin** giữa hai nhóm — một phiên chỉ toàn mở nhóm rồi
  bấm là dấu vết rất dễ nhận.
- **Không chạy ngoài khung giờ** — xin vào nhóm lúc 3h sáng thì người thật đang ngủ.
- **Kiểm chứng sau mỗi cú bấm**: nút phải đổi sang "Đã yêu cầu" hoặc vào hẳn
  nhóm. Nếu bấm mà nút không đổi **2 lần liên tiếp** thì gần như chắc chắn
  Facebook đang chặn ngầm → dừng ngay.

### Khi tool khoá nick lại

Gặp checkpoint, gặp thông báo "tạm thời bị chặn", hoặc bấm 2 lần không ăn thua,
tool **dừng và khoá việc tham gia nhóm của nick đó 24 tiếng**. Trong 24 tiếng
đó bấm chạy là bị từ chối ngay, kèm mốc giờ hết hạn.

Đây là bảo vệ, không phải lỗi. **Chạy tiếp ngay sau khi bị chặn là cách nhanh
nhất để mất nick.** Nút *"Bỏ tạm nghỉ của nick"* có sẵn nếu bạn chắc chắn là
báo nhầm — nhưng bỏ rồi thì tự chịu.

### Vài điều nữa

Phần lớn nhóm **phải chờ quản trị viên duyệt**, có nhóm còn bắt trả lời câu
hỏi. Tool **không tự trả lời câu hỏi hộ** — trả lời bừa là bị từ chối vĩnh
viễn, sau đó vào bằng tay cũng khó.

Nhóm không vào được (link hỏng, nhóm đã xóa, nhóm hỏi câu hỏi) được ghi lại và
**7 ngày sau mới thử lại**, để mỗi lần chạy tool không mở đi mở lại mấy link
chết.

**Thấy trạng thái không khớp với thực tế?** Ví dụ bảng ghi "Đã là thành viên"
nhưng bạn vào Facebook kiểm tra thì chưa hề tham gia. Bấm **🔄 Soát lại trạng
thái** — nút này bỏ đi những bản ghi mà tool *tự nhìn trang rồi kết luận* (giữ
lại các nhóm tool thật sự đã bấm tham gia), đưa chúng về "chưa xử lý" để lần
chạy tới kiểm tra lại từ đầu.

---

## 10. Hẹn giờ đăng tự động

Vào tab **⏰ Hẹn giờ đăng**, bấm **+ Thêm lịch**. Có 2 kiểu:

- **Một lần** — chọn ngày + giờ cụ thể, chạy xong lịch tự tắt.
- **Lặp lại theo thứ** — như báo thức: chọn giờ và các thứ trong tuần. Không
  tick thứ nào thì hiểu là mọi ngày.

Mỗi lịch tự chọn **chiến dịch** và **nick** của riêng nó, nên đặt được kiểu
"7h sáng đăng phòng trọ bằng nick A, 20h đăng tuyển dụng bằng nick B". Muốn
nhiều mốc trong cùng một ngày thì tạo nhiều lịch, mỗi lịch một mốc giờ.

Giờ hẹn luôn tính theo **giờ Việt Nam (UTC+7)**, kể cả khi máy bạn đặt sai múi
giờ. Đồng hồ ngay trên bảng lịch hiện giờ Việt Nam để bạn đối chiếu.

**Ba điều cần nhớ:**

1. **Máy phải đang bật và tool đang mở** thì lịch mới chạy. Tool không phải
   dịch vụ chạy trên mạng — tắt máy là không có gì đăng cả.
2. Tick ô **"Tự mở tool khi bật máy"** để Windows tự mở tool (thu nhỏ sẵn, không
   chắn màn hình) mỗi lần bạn đăng nhập máy. Đây là cách để lịch chạy đều mà
   không phải nhớ mở tool thủ công.
3. Mở tool **muộn hơn giờ hẹn**, tool sẽ hỏi *"Chạy bù ngay bây giờ?"* — bạn tự
   quyết. Lỡ quá 12 tiếng thì tool bỏ luôn, không hỏi, tránh chuyện mở máy lúc
   nửa đêm rồi bài lũ lượt lên vào khung giờ không ai đọc.

Đang có bài chạy dở thì lịch tới giờ sẽ được bỏ qua, không chạy chồng lên nhau.
Nếu đến giờ mà thiếu file Excel, sai nick hoặc dữ liệu có lỗi, tool ghi lý do
vào ô log rồi bỏ lượt đó chứ không hiện hộp thoại chờ bạn bấm — vì lúc đó có
thể chẳng có ai ngồi trước máy.

---

## 11. Hạn chế rủi ro khóa tài khoản

Facebook chống công cụ tự động. Rủi ro là có thật, **không thể loại bỏ hoàn
toàn, chỉ giảm được.** Dùng **tài khoản phụ, không dùng tài khoản chính.**

Xếp theo mức độ quan trọng:

1. **Tốc độ đăng — nguyên nhân khóa nick nhiều nhất.** Vào tab "Dữ liệu chiến
   dịch", để thời gian chờ giữa các nhóm **ít nhất 60–120 giây**. Đăng 20 nhóm
   trong 10 phút gần như chắc chắn dính.
2. **Nội dung lặp lại.** Đừng để y hệt một nội dung đăng khắp mọi nhóm. Đổi câu
   chữ, đổi ảnh, đổi thứ tự.
3. **Nick mới phải "làm nóng".** Nick vừa lập đừng đăng ngay. Vài ngày đầu chỉ
   like, bình luận, kết bạn, đăng 1–2 bài. Tăng dần lên.
4. **Xin vào nhóm hàng loạt.** Vào vài chục nhóm trong một buổi là dấu hiệu
   Facebook bắt rất nhanh — nick mới càng dễ dính. Xem mục
   [Tham gia nhóm tự động](#9-tham-gia-nhóm-tự-động) và giữ trần mỗi ngày thấp.
5. **Nhiều nick trên một máy.** Mỗi nick là một profile riêng (tool tự tách khi
   bấm "+ Thêm nick mới"). Đừng chạy 2 nick cùng lúc — giãn nhau vài giờ. Làm
   nghiêm túc thì mỗi nick một mạng: một nick wifi, một nick 4G điện thoại.

> Đăng 2 nick riêng biệt **không** an toàn hơn nếu chúng chung máy, chung mạng
> và đăng cùng kiểu — Facebook nhìn ra là cùng một người. Cái quyết định bị
> khóa là **tốc độ và nội dung lặp**, không phải số lượng nick.

---

## 12. Lỗi thường gặp

| Hiện tượng | Cách xử lý |
|---|---|
| `Không tìm thấy Google Chrome trên máy này` | Cài Chrome tại https://www.google.com/chrome rồi mở lại tool |
| `Không chạy được trình điều khiển trình duyệt` / `[WinError 2]` | Thiếu file `node.exe` trong thư mục tool — **không phải lỗi Chrome**. Xem dòng "Cần file" ngay dưới thông báo: thường do phần mềm diệt virus xóa mất (vào mục Cách ly khôi phục lại, rồi thêm thư mục tool vào danh sách loại trừ), hoặc giải nén thiếu file |
| `⚠ Chưa có file: ...xlsx` | File Excel bị xóa/đổi chỗ. Vào tab "Dữ liệu chiến dịch" chỉ lại đường dẫn |
| Không đọc được Excel | Đang mở file đó trong Excel. Đóng lại rồi bấm ⟳ Tải lại từ Excel |
| Sửa nội dung rồi mà chạy vẫn ra bài cũ | Chưa bấm **💾 Ghi vào file Excel** sau khi sửa |
| Báo nhóm không đăng được | Nick chưa tham gia nhóm đó, hoặc nhóm bắt duyệt bài. Dùng tab **👥 Tham gia nhóm** hoặc kiểm tra bằng tay |
| Tab Tham gia nhóm báo `không thấy nút tham gia` | Link nhóm hỏng/nhóm đã xóa, hoặc nhóm bắt trả lời câu hỏi trước khi vào — phải vào bằng tay |
| Tham gia nhóm dừng giữa chừng, báo bị chặn tạm thời | **Nghỉ vài ngày**, đừng chạy lại ngay. Sau đó hạ 'Tối đa mỗi nick mỗi ngày' xuống |
| Bị đòi đăng nhập lại liên tục | Thư mục `fb_profiles\` bị xóa, hoặc tool đang nằm trong `C:\Program Files` (không ghi được) |
| Đăng vài bài rồi dừng, báo checkpoint | Facebook đã cảnh báo nick. **Dừng ngay**, mở Facebook đăng nhập bằng tay giải quyết, nghỉ vài ngày rồi tăng thời gian chờ lên |

---

## 13. Tách dữ liệu ra thư mục riêng — **nên làm ngay từ đầu**

Mặc định mọi thứ nằm chung trong thư mục tool: cả code lẫn dữ liệu của bạn.

Vấn đề nằm ở lúc nhận bản cập nhật. Bản mới là một thư mục mới hoàn toàn, bạn
sẽ thay thư mục cũ bằng nó — **và cuốn theo toàn bộ nhóm, nội dung, ảnh, phiên
đăng nhập.** Cấu hình lại từ đầu, đăng nhập lại từng nick.

Chỉ cần làm **một lần** dưới đây là hết lo chuyện đó mãi mãi.

### Bước 1 — Tạo thư mục dữ liệu

Tạo một thư mục **nằm ngoài thư mục tool**, ví dụ `D:\DuLieuTool\`.

> Ổ nào, tên gì cũng được, miễn là:
> - **không** nằm trong thư mục tool
> - **không** nằm trong `C:\Program Files` (Windows chặn ghi ở đó)

### Bước 2 — Kéo dữ liệu sang

**Đóng tool trước.** Rồi kéo những thứ này từ thư mục tool sang `D:\DuLieuTool\`:

| Kéo ra ngoài | Là gì |
|---|---|
| `fb_profiles\` | Phiên đăng nhập các nick — **quan trọng nhất**, mất là phải đăng nhập lại hết |
| `phong_tro.xlsx` *(và mọi file `.xlsx` khác)* | Nội dung bài của bạn |
| `anh_phong\` *(và mọi thư mục ảnh khác)* | Ảnh đăng kèm |
| `posted_log.json` *(và mọi file `posted_log_*.json`)* | Ghi nhớ bài nào đã đăng, để không đăng trùng |

**Giữ nguyên tại chỗ, đừng đụng vào:** `ToolDangBaiFacebook.exe`, thư mục
`_internal\`, và `config.json` (xem Bước 4).

### Bước 3 — Chỉ lại đường dẫn cho tool

Mở tool, vào tab **Dữ liệu chiến dịch**, sửa 4 ô cho trỏ tới chỗ mới (bấm nút
**Chọn...** cho nhanh):

- File Excel dữ liệu → `D:\DuLieuTool\phong_tro.xlsx`
- Thư mục ảnh → `D:\DuLieuTool\anh_phong`
- File log bài đã đăng → `D:\DuLieuTool\posted_log.json`
- Thư mục profile Chrome → `D:\DuLieuTool\fb_profiles`

Bấm **💾 Lưu cấu hình chiến dịch**.

> ⚠ Ba ô đầu là **của riêng từng chiến dịch**. Có nhiều chiến dịch thì phải đổi
> ô **Chiến dịch** ở góc trên rồi làm lại cho từng cái. Riêng ô "Thư mục profile
> Chrome" dùng chung, chỉ cần đặt một lần.

Xong bấm **🔍 Kiểm tra dữ liệu** ở tab Tài khoản & Chạy — không báo lỗi là đã
trỏ đúng.

### Bước 4 — Riêng `config.json` thì sao

`config.json` là ngoại lệ duy nhất: **nó bắt buộc phải nằm cạnh file .exe**,
không trỏ đi nơi khác được. Nhưng nó chỉ là một file nhỏ, và giờ nó chẳng chứa
gì ngoài đường dẫn với danh sách nhóm — copy tay một cái là xong (mục 14).

Cho chắc, copy sẵn một bản dự phòng vào `D:\DuLieuTool\` luôn.

---

## 14. Nhận bản cập nhật

Khi được gửi bản mới:

1. **Đóng tool.**
2. Copy `config.json` từ thư mục tool cũ sang `D:\DuLieuTool\`.
3. Đổi tên thư mục tool cũ thành `ToolDangBaiFacebook_cu` (đừng xóa vội).
4. Giải nén bản mới ra.
5. Copy `config.json` ở bước 2 vào thư mục mới, **cạnh file .exe**.
6. Mở tool. Kiểm tra tab **Dữ liệu chiến dịch** thấy đường dẫn vẫn trỏ về
   `D:\DuLieuTool\`, danh sách nick vẫn còn → xong, xóa `ToolDangBaiFacebook_cu`.

Toàn bộ nhóm, nội dung, ảnh, phiên đăng nhập giữ nguyên vì chúng nằm ở
`D:\DuLieuTool\`, bản cập nhật không hề đụng tới.

> **Chưa làm mục 13?** Vậy thì bước 2 phải copy ra **tất cả**: `config.json`,
> `fb_profiles\`, các file `.xlsx`, các thư mục ảnh, các file `posted_log*.json`
> — rồi bước 5 copy ngược lại hết. Sót một thứ là mất thứ đó. Đây đúng là lý do
> nên làm mục 13.

---

## 15. Sao lưu và chuyển máy

**Sao lưu:** copy thư mục `D:\DuLieuTool\` sang USB hoặc Google Drive, thỉnh
thoảng làm lại. Máy hỏng là mất hết nếu không có bản sao. Thư mục tool thì
không cần sao lưu — hỏng thì xin lại file zip là xong.

**Chuyển sang máy khác:** copy cả thư mục tool lẫn thư mục dữ liệu sang máy mới
(máy đó cũng cần có Chrome). Mở tool, vào tab **Dữ liệu chiến dịch** kiểm tra
lại đường dẫn — nếu để thư mục dữ liệu ở vị trí khác trên máy mới thì phải trỏ
lại, vì tool lưu đường dẫn đầy đủ chứ không tự dò.

> Lưu ý mục 11: chuyển nick sang máy khác là đổi thiết bị đột ngột dưới mắt
> Facebook. Đổi cả máy lẫn mạng cùng lúc thì khả năng bị hỏi xác minh khá cao —
> nên đăng nhập lại bằng tay một lần trên máy mới trước khi chạy tool.
