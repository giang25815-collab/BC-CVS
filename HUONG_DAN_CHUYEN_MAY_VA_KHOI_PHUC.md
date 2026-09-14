# 📦 HƯỚNG DẪN CHUYỂN DỰ ÁN QUA MÁY MỚI VÀ LOAD LẠI TOÀN BỘ NÃO BỘ

Tài liệu này hướng dẫn chi tiết cách chuyển gói dự án này sang máy tính mới và kích hoạt toàn bộ trí nhớ (Não bộ / Ngữ cảnh / Context) để AI tiếp tục làm việc chính xác 100% như trên máy hiện tại.

---

## 🚀 1. CÁCH GIẢI NÉN VÀ MỞ DỰ ÁN TRÊN MÁY TÍNH MỚI

1. **Sao chép tệp nén:**
   - Tệp nén đầy đủ: `Team_CamGiang_Full_Project_And_Brain_Backup.zip`
   - Copy tệp này qua USB hoặc tải lên Google Drive / Zalo / OneDrive để tải về máy mới.
2. **Giải nén:**
   - Giải nén tệp zip vào ổ đĩa trên máy mới, ví dụ: `D:\BC-CVS-main` hoặc `C:\Projects\BC-CVS-main`.
3. **Mở thư mục trên IDE:**
   - Mở IDE (Khuyến nghị **Antigravity IDE** hoặc Cursor / VS Code).
   - Chọn **Open Folder** và trỏ thẳng vào thư mục `gas_app` (hoặc thư mục gốc dự án).

---

## 🧠 2. CƠ CHẾ TỰ ĐỘNG LOAD NÃO BỘ CHO AI TRÊN MÁY MỚI

Hệ thống đã được thiết kế sẵn để **TỰ ĐỘNG NẠP NÃO** khi bạn mở thư mục:

### Khi dùng Antigravity IDE / Gemini:
- Antigravity sẽ **tự động quét và nạp toàn bộ**:
  - `GEMINI.md`: Chỉ dẫn nghiệp vụ cốt lõi, danh sách 138 cửa hàng, số liệu chuẩn.
  - `AGENTS.md`: Chỉ dẫn vận hành kỹ thuật.
  - `.agents/rules/01-project-context.md`: Các luật bất biến (Hòa không làm WinMart+, đồng bộ Android assets...).
  - `.agents/skills/cvs-report-expert/`: Kỹ năng chuyên sâu bóc tách sheet NPP, build APK, đồng bộ Excel.
- **Bạn chỉ cần chat:**
  > *"Tiếp tục dự án báo cáo doanh số Team Cẩm Giang"* hoặc *"Đọc file PROJECT_BRAIN.md và hỗ trợ tôi..."*  
  AI trên máy mới sẽ lập tức nắm rõ 100% mọi chi tiết và lịch sử làm việc!

### Khi dùng Cursor IDE / VS Code / Claude:
- Tệp `.cursorrules` và `PROJECT_BRAIN.md` sẽ lập tức định hướng cho AI.
- Bạn có thể tag: `@PROJECT_BRAIN.md` vào khung chat.

---

## 💻 3. CÁCH CHẠY DỰ ÁN TRÊN MÁY MỚI

### A. Chạy ứng dụng Web (Báo cáo & Xuất Excel):
- Cách 1 (Tiện nhất): Nhấp đúp vào tệp `KHOI_DONG_WEB.bat`.
- Cách 2: Mở trực tiếp file `index.html` bằng trình duyệt web (Chrome, Edge, Cốc Cốc).
- Cách 3 (Dùng server): Mở terminal trong thư mục `gas_app` và gõ:
  ```bash
  python -m http.server 8080
  ```
  sau đó truy cập: `http://localhost:8080/index.html`.

### B. Sử dụng ứng dụng Android (APK):
- Trong gói nén đã có sẵn 2 tệp APK bản phát hành hoàn thiện mới nhất:
  - `BaoCaoDoanhSo_TeamCamGiang.apk`
  - `BaoCaoThucDat.apk`
- Bạn có thể gửi thẳng 2 file này vào điện thoại Android để cài đặt sử dụng ngay.
- **Nếu muốn build lại APK trên máy mới:**
  1. Mở file `android/local.properties`.
  2. Sửa lại đường dẫn Android SDK cho đúng với máy mới:
     ```properties
     sdk.dir=C\:\\Users\\<TenUserCuaMayMoi>\\AppData\\Local\\Android\\Sdk
     ```
  3. Chạy file `build_apk.bat`.

---

## 📊 4. TÓM TẮT DỮ LIỆU ĐÃ ĐỒNG BỘ CHUẨN XÁC
- **Tổng số cửa hàng CVS:** **138 cửa hàng** (GS25: 70, FamilyMart: 34, Circle K: 27, 7-Eleven: 5, Hoàng Đức: 2, WinMart+: 0).
- **Nhân sự Nguyễn Đức Hòa:** **48 cửa hàng** (35 GS25, 13 Circle K, 0 WinMart+), doanh thu CVS: `114,438,802 đ`.
- **Toàn Team:** Target `10,826,339,729 đ`, Thực đạt `3,911,174,471 đ` (36.1%).
- **Cơ chế SKU NPP:** Tự động nhận diện mọi mã SKU mới ngoài 46 SKU cũ và gom vào danh mục phát sinh, không bỏ sót bất kỳ đồng doanh thu nào.
- **Hình ảnh đối soát gốc:** Đã được lưu trữ an toàn trong `.brain_backup/images/`.
