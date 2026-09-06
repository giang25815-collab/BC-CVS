# 🌐 HƯỚNG DẪN ĐỂ TELEGRAM BOT CHẠY 24/7 (KHI ĐÃ TẮT MÁY TÍNH)

## ❓ VÌ SAO HIỆN TẠI TẮT MÁY THÌ BOT KHÔNG CHẠY ĐƯỢC?
Hiện tại, script `telegram_bot.py` đang chạy trực tiếp trên máy tính cá nhân của bạn. Khi bạn **Shutdown / Tắt máy tính**, chương trình Python sẽ bị tắt theo, nên Bot Telegram không còn ai kết nối để nhận lệnh và xử lý file.

---

## 🎯 GIẢI PHÁP: ĐƯA BOT LÊN "MÁY CHỦ ĐÁM MÂY" (CLOUD / VPS 24/7)
Khi đưa code lên máy chủ đám mây, bot sẽ luôn luôn thức 24/24. 
👉 **Bạn chỉ cần dùng điện thoại gửi 2 file Excel vào Telegram bất kỳ lúc nào, bot trên đám mây sẽ tự tính toán và trả file `Team_CamGiang_Report.xlsx` về điện thoại của bạn ngay lập tức!**

---

### 🌟 CÁCH 1: DÙNG RENDER.COM (KHUYÊN DÙNG - CỰC NHANH VÀ TIỆN)
Render.com là nền tảng điện toán đám mây cho phép chạy ứng dụng Python.

1. **Đưa mã nguồn lên GitHub:**
   - Tạo tài khoản tại [github.com](https://github.com).
   - Tạo một kho lưu trữ mới (Repository) ở chế độ **Private** (Riêng tư).
   - Tải toàn bộ các file trong thư mục này lên Repo đó (bao gồm: `telegram_bot.py`, `team_leader_report_v4.py`, `update_data_and_config.py`, `requirements.txt`, `Procfile`, `team_config.csv`, `team_cvs_stores.csv`, `team_employees.csv`).

2. **Kết nối với Render:**
   - Vào [render.com](https://render.com) đăng ký tài khoản (đăng nhập nhanh bằng GitHub).
   - Chọn **New +** -> **Background Worker**.
   - Chọn Repository GitHub bạn vừa tạo ở Bước 1.
   - Cấu hình:
     - **Name**: `cvs-bhx-telegram-bot`
     - **Runtime**: `Python 3`
     - **Build Command**: `pip install -r requirements.txt`
     - **Start Command**: `python telegram_bot.py`
   - Mục **Environment Variables** (Thêm 2 biến):
     - `BOT_TOKEN`: `<Token bot của bạn>`
     - `ALLOWED_CHAT_ID`: `<ID chat Telegram của bạn>`
   - Bấm **Create Background Worker**.

👉 **Xong!** Từ giờ Render sẽ giữ bot chạy 24/7 trên mạng. Bạn có thể tắt máy tính thoải mái.

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
