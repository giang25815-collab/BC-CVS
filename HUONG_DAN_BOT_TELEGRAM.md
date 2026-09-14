# HƯỚNG DẪN THIẾT LẬP VÀ CHẠY BOT TELEGRAM BÁO CÁO DOANH SỐ (CVS & BHX)

Hệ thống cho phép bạn gửi **2 file Excel doanh số** (File Siêu Thị BHX + WinMart và File CVS & NPP) trực tiếp vào Bot Telegram. Bot sẽ tự động phân tích dữ liệu, tính toán KPI theo nhân viên và trả về:
1. 📎 **File Báo Cáo Excel (.xlsx)** chuẩn định dạng Team Cẩm Giang.
2. 🔗 **Đường dẫn xem trực tiếp Google Sheet Online** trên Google Drive.
3. 💬 **Tin nhắn tóm tắt KPI Dashboard** (Target, Thực hiện, % Đạt, tình trạng tiến độ, xếp hạng nhân viên).

---

## 📌 BƯỚC 1: TẠO BOT TELEGRAM & LẤY BOT TOKEN (MẤT 1 PHÚT)

1. Mở ứng dụng **Telegram** trên điện thoại hoặc máy tính.
2. Tìm kiếm người dùng chính thức: **`@BotFather`** (có dấu tích xanh).
3. Bấm **Start** hoặc gửi lệnh: `/newbot`
4. Đặt tên hiển thị cho Bot, ví dụ: `Báo Cáo Doanh Số Team`
5. Đặt username cho Bot (phải kết thúc bằng chữ `bot`, viết liền không dấu), ví dụ: `cvs_bhx_report_bot` hoặc `team_camgiang_bot`.
6. BotFather sẽ gửi cho bạn một đoạn mã Token có định dạng như sau:
   ```text
   7891234567:AAFn_XXXXXXX-YYYYYYY_ZZZZZZZZZZZ
   ```
   👉 **Hãy copy và lưu lại đoạn Token này** để dùng ở Bước 3.

---

## 📌 BƯỚC 2: TẢI HOẶC COPY MÃ NGUỒN LÊN GOOGLE APPS SCRIPT

Nếu bạn đang dùng trực tiếp trang quản trị của Google Apps Script ([script.google.com](https://script.google.com)):

1. Tạo một dự án mới hoặc mở dự án hiện tại của bạn.
2. Đảm bảo dự án có đầy đủ các file sau (mở các file tương ứng trong thư mục này để copy nội dung):
   - **`Code.gs`**: Chứa logic tạo 6 sheet Google Spreadsheet.
   - **`MasterData.gs`**: Chứa danh sách nhân viên, target mặc định, 35 cửa hàng FamilyMart và SKU.
   - **`XlsxBundle.gs`**: Thư viện SheetJS đọc file Excel trên backend.
   - **`ReportEngine.gs`**: Bộ máy bóc tách dữ liệu 2 file.
   - **`TelegramBot.gs`**: Xử lý Webhook Telegram, nhận 2 file và gửi trả báo cáo.
   - **`Index.html`**: Giao diện Web App (nếu muốn dùng song song trên web).
   - **`appsscript.json`**: Cấu hình múi giờ `Asia/Ho_Chi_Minh`.

---

## 📌 BƯỚC 3: CÀI ĐẶT BOT TOKEN VÀO GOOGLE APPS SCRIPT

Có 2 cách đơn giản để lưu Bot Token vào hệ thống:

### Cách 1: Chạy hàm `setTelegramBotToken` (Khuyên dùng)
1. Trong màn hình soạn thảo Apps Script, mở file **`TelegramBot.gs`**.
2. Tìm hàm `setTelegramBotToken`:
   ```javascript
   function setBotTokenQuickly() {
     setTelegramBotToken('DÁN_TOKEN_CỦA_BẠN_VÀO_ĐÂY');
   }
   ```
3. Thay thế chuỗi bằng Token bạn vừa lấy từ `@BotFather`.
4. Trên thanh công cụ, chọn hàm `setBotTokenQuickly` và bấm nút **Chạy (Run)**.

### Cách 2: Nhập trong Cài đặt Dự án (Project Settings)
1. Bấm vào biểu tượng **Bánh răng (Project Settings)** ở thanh menu bên trái.
2. Kéo xuống mục **Script Properties (Thuộc tính tập lệnh)**.
3. Bấm **Add script property**:
   - **Property:** `TELEGRAM_BOT_TOKEN`
   - **Value:** Dán đoạn Token lấy từ BotFather vào.
4. Bấm **Save script properties**.

---

## 📌 BƯỚC 4: TRIỂN KHAI ỨNG DỤNG WEB (DEPLOY WEB APP)

Để Telegram có thể gửi file và tin nhắn cho Apps Script, bạn cần kích hoạt Web App:

1. Ở góc trên cùng bên phải, bấm nút xanh **Deploy (Triển khai)** > **New deployment (Triển khai mới)**.
2. Bấm vào biểu tượng bánh răng bên cạnh "Select type" > chọn **Web app (Ứng dụng web)**.
3. Điền thông tin:
   - **Description:** `Telegram Bot Report v1.0`
   - **Execute as:** **Me (Tôi - email của bạn)**
   - **Who has access:** **Anyone (Bất kỳ ai)** *(Rất quan trọng: Bắt buộc chọn Anyone để Telegram gửi được Webhook)*.
4. Bấm nút **Deploy**.
5. Google sẽ yêu cầu cấp quyền truy cập (**Authorize access**):
   - Chọn tài khoản Google của bạn.
   - Nếu thấy cảnh báo "Google hasn't verified this app" -> Bấm **Advanced (Nâng cao)** -> Chọn **Go to ... (unsafe)** -> Bấm **Allow (Cho phép)**.
6. Sau khi hoàn tất, hệ thống sẽ hiện ra một đường dẫn **Web app URL** (dạng: `https://script.google.com/macros/s/AKfycb.../exec`).

---

## 📌 BƯỚC 5: KÍCH HOẠT WEBHOOK CHO BOT (1-CLICK)

1. Quay lại màn hình soạn thảo code Apps Script, mở file **`TelegramBot.gs`**.
2. Trên thanh chọn hàm cần chạy, chọn hàm: **`setupTelegramWebhook`**.
3. Bấm nút **Run (Chạy)**.
4. Mở cửa sổ **Execution log (Nhật ký thực thi)** phía dưới:
   - Thấy hiện: `✅ Đăng ký Webhook thành công!` nghĩa là Bot của bạn đã chính thức hoạt động 24/7!

*(Hoặc bạn có thể mở tab mới trên trình duyệt và gõ: `https://api.telegram.org/bot<TOKEN_CỦA_BẠN>/setWebhook?url=<WEB_APP_URL>`)*.

---

## 📌 BƯỚC 6: HƯỚNG DẪN SỬ DỤNG TRÊN TELEGRAM

1. Mở Telegram, tìm tên bot của bạn và bấm **Start** (hoặc gửi `/start`).
2. Gửi **File 1 (File Siêu Thị BHX + WinMart+)**:
   - Bot sẽ tự động nhận diện và phản hồi:
     > ✅ **Đã nhận diện:** 🏬 File Siêu Thị (BHX & WinMart)
     > 📄 *SO_ST_KD6_T09.xlsx*
     >
     > ⏳ **Còn thiếu:** 🏪 File CVS & NPP
     > 👉 *Vui lòng gửi tiếp file thứ 2 để Bot hoàn thiện báo cáo!*
3. Gửi tiếp **File 2 (File CVS & NPP)**:
   - Bot lập tức thông báo:
     > ✨ **ĐÃ NHẬN ĐỦ 2 FILE DỮ LIỆU!**
     > ⏳ *Hệ thống đang tiến hành bóc tách số liệu và tạo báo cáo chuẩn Team Cẩm Giang... Vui lòng đợi trong khoảng 10-15 giây!*
4. Sau khoảng 10 - 15 giây, bạn sẽ nhận được:
   - 📎 **File Excel hoàn chỉnh** (`Team_Tran_Thi_Cam_Giang_Report_Thang_9_2026.xlsx`) tải về mở được ngay trên Excel máy tính hoặc điện thoại.
   - 📊 **Tin nhắn tóm tắt KPI Dashboard**: Target, Thực hiện BHX, Thực hiện CVS, % Đạt, Đánh giá tiến độ (Vượt/Cận/Chậm), Bảng xếp hạng Top nhân viên.
   - 🔗 **Link Google Sheets** để xem online hoặc chia sẻ ngay cho cấp trên và team.

---

## 🛠️ CÁC LỆNH HỖ TRỢ TRONG BOT

- `/start` hoặc `/help`: Xem hướng dẫn sử dụng.
- `/status`: Kiểm tra bot đang lưu những file nào, còn thiếu file nào.
- `/reset`: Xóa dữ liệu phiên hiện tại để nạp lại từ đầu.

---

## ⚠️ MỘT SỐ LƯU Ý KHI CHỈNH SỬA CODE SAU NÀY

Mỗi khi bạn sửa bất kỳ dòng code nào trong Google Apps Script, để phiên bản mới có hiệu lực với Telegram Bot:
1. Bấm **Deploy (Triển khai)** > **Manage deployments (Quản lý các bản triển khai)**.
2. Bấm vào biểu tượng chiếc bút chì (Chỉnh sửa).
3. Ở mục **Version (Phiên bản)**: Chọn **New version (Phiên bản mới)**.
4. Bấm **Deploy**.
*(Không cần phải cài đặt lại Webhook nếu dùng cách này)*.
