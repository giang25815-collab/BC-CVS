# 🌐 HƯỚNG DẪN ĐỂ TELEGRAM BOT CHẠY 24/7 (KHI ĐÃ TẮT MÁY TÍNH)

## ❓ VÌ SAO HIỆN TẠI TẮT MÁY THÌ BOT KHÔNG CHẠY ĐƯỢC?
Hiện tại, script `telegram_bot.py` đang chạy trực tiếp trên máy tính cá nhân của bạn. Khi bạn **Shutdown / Tắt máy tính**, chương trình Python sẽ bị tắt theo, nên Bot Telegram không còn ai kết nối để nhận lệnh và xử lý file.

---

## 🎯 GIẢI PHÁP: ĐƯA BOT LÊN "MÁY CHỦ ĐÁM MÂY" (CLOUD / VPS 24/7)
Khi đưa code lên máy chủ đám mây, bot sẽ luôn luôn thức 24/24. 
👉 **Bạn chỉ cần dùng điện thoại gửi 2 file Excel vào Telegram bất kỳ lúc nào, bot trên đám mây sẽ tự tính toán và trả file `Team_CamGiang_Report.xlsx` về điện thoại của bạn ngay lập tức!**

---

### 🌟 CÁCH 1: DÙNG RENDER.COM + UPTIMEROBOT (MIỄN PHÍ 100% - KHUYÊN DÙNG)
Render.com cung cấp gói **Web Service miễn phí**. Kết hợp với **Flask Web Server** vừa được tích hợp vào Bot và dịch vụ ping tự động **UptimeRobot**, Bot sẽ thức 24/24 mà không bao giờ ngủ!

#### Bước 1: Đưa mã nguồn lên GitHub
1. Tạo tài khoản tại [github.com](https://github.com).
2. Tạo một kho lưu trữ mới (Repository) ở chế độ **Private** (Riêng tư).
3. Tải toàn bộ các file trong thư mục này lên Repo đó (hoặc giải nén từ file `bot_cvs_deploy.zip` rồi đẩy lên).

#### Bước 2: Tạo Web Service trên Render
1. Vào [render.com](https://render.com) đăng ký/đăng nhập bằng GitHub.
2. Chọn **New +** -> **Web Service** (chọn Web Service để Render cấp URL công khai).
3. Chọn Repository GitHub vừa tạo.
4. Điền các thông số:
   - **Name**: `cvs-bhx-telegram-bot`
   - **Language / Runtime**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `python telegram_bot.py`
   - **Instance Type**: Chọn **Free** ($0/month)
5. Mục **Environment Variables** (Thêm 2 biến):
   - `BOT_TOKEN`: `<Token bot của bạn>`
   - `ALLOWED_CHAT_ID`: `<ID chat Telegram của bạn>`
6. Bấm **Create Web Service**.
7. Đợi 1-2 phút Render build xong, bạn sẽ thấy link web dạng:  
   `https://cvs-bhx-telegram-bot.onrender.com`  
   (Nhấp vào link thấy màn hình báo *ONLINE 24/7* là web server đã hoạt động).

#### Bước 3: Cài đặt UptimeRobot để Bot KHÔNG BAO GIỜ NGỦ
Gói Free của Render sẽ tạm dừng (sleep) sau 15 phút nếu không có lượt truy cập web. Để bot thức liên tục:
1. Vào trang miễn phí [uptimerobot.com](https://uptimerobot.com) tạo tài khoản.
2. Bấm **Add New Monitor**:
   - **Monitor Type**: `HTTP(s)`
   - **Friendly Name**: `Bot Doanh So Render`
   - **URL (or IP)**: Điền link web Render của bạn kèm `/ping` (Ví dụ: `https://cvs-bhx-telegram-bot.onrender.com/ping`)
   - **Monitoring Interval**: `5 minutes` (hoặc `10 minutes`)
3. Bấm **Create Monitor**.

👉 **XONG!** Cứ mỗi 5 phút UptimeRobot sẽ "đánh thức" server một lần qua cổng Flask web. Bot Telegram và Web Server sẽ luôn thức 24/7 để bạn gửi file bất kỳ lúc nào!

---

### 🌟 CÁCH 2: DÙNG PYTHONANYWHERE (MIỄN PHÍ - DỄ LÀM KHÔNG CẦN GITHUB)
Nếu bạn không quen dùng GitHub, PythonAnywhere có giao diện kéo thả file trực tiếp trên web:

1. Vào [pythonanywhere.com](https://www.pythonanywhere.com/) đăng ký tài khoản **Free Beginner**.
2. Vào tab **Files**: Nén toàn bộ thư mục này thành file `.zip` và upload lên PythonAnywhere, sau đó mở Bash console gõ: `unzip <ten_file>.zip`.
3. Trong Bash console, gõ lệnh cài thư viện:
   ```bash
   pip install -r requirements.txt
   ```
4. Chạy bot:
   ```bash
   python3 telegram_bot.py
   ```
*(Lưu ý: Tài khoản miễn phí của PythonAnywhere cần bấm "Extend" 3 tháng 1 lần).*

---

### 🌟 CÁCH 3: THUÊ MỘT CLOUD VPS LINUX (ỔN ĐỊNH VĨNH VIỄN 100%)
Nếu bạn muốn dùng lâu dài và chuyên nghiệp cho công việc:
- Thuê 1 Cloud VPS nhỏ (giá chỉ khoảng 30.000đ - 60.000đ/tháng tại Vietnix, CloudFly, BKNS, TinoHost...).
- Đăng nhập vào VPS bằng SSH và gõ lệnh chạy nền:
  ```bash
  nohup python3 telegram_bot.py > bot.log 2>&1 &
  ```
  hoặc chạy bằng Docker:
  ```bash
  docker build -t telegram-bot .
  docker run -d --restart always --name cvs_bot telegram-bot
  ```
- Bot sẽ chạy vĩnh viễn không bao giờ tắt.

---

### 🌟 CÁCH 4: DÙNG MÁY TÍNH PHỤ / MÁY BÀN CÔNG TY TREO 24/24
Nếu ở văn phòng hoặc ở nhà có máy tính bàn cắm điện và wifi liên tục:
- Chỉ cần copy thư mục này sang máy đó.
- Nhấp đúp chuột vào file `run_telegram_bot.bat` để máy phụ đó chạy.
- Còn laptop/máy tính cá nhân của bạn thì cứ tắt máy mang về thoải mái.
