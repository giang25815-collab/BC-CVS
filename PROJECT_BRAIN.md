# 🧠 HỆ THỐNG NÃO BỘ VÀ NGỮ CẢNH DỰ ÁN (PROJECT BRAIN & MEMORY)
**Dự án:** Hệ thống Báo Cáo Doanh Số CVS & BHX - Team Cẩm Giang (Sales Report Intelligence App)  
**Phiên bản cập nhật:** 14/09/2026 (Đã đồng bộ toàn diện logic, số liệu và tự động nhận diện SKU)  
**Tác giả / Quản trị:** Team Cẩm Giang & AI Pair Programmer (Antigravity / Gemini)

---

## 📌 1. TỔNG QUAN HỆ THỐNG VÀ KIẾN TRÚC ĐA NỀN TẢNG
Dự án là một giải pháp báo cáo doanh số đa kênh tích hợp, kết hợp tự động hoá trích xuất dữ liệu, tính toán phần trăm hoàn thành chỉ tiêu (% Đạt) và hiển thị trực quan trên 4 nền tảng:
1. **Web App độc lập (Single Page App - Vanilla JS/CSS):** `index.html`, `view.html`, `master_data.js`.
2. **Ứng dụng Android (APK):** Gói WebView tự động chạy offline/online từ assets (`android/app/src/main/assets/www/`). Tự động build ra 2 bản APK phát hành:
   - `BaoCaoDoanhSo_TeamCamGiang.apk`
   - `BaoCaoThucDat.apk`
3. **Google Apps Script (GAS WebApp & Telegram Bot):** `Code.gs`, `MasterData.gs`, `ReportEngine.gs`, `TelegramBot.gs`, `XlsxBundle.gs`.
4. **Báo cáo Excel chuẩn hóa (Dynamic Excel Engine):** `Team_CamGiang_Report.xlsx` với các sheet đồng bộ hoàn hảo:
   - `Dashboard`: Tổng quan chỉ tiêu, thực đạt và tỷ lệ % của toàn team.
   - `Tiến độ Team`: Bảng tổng hợp theo từng nhân viên và từng chuỗi bán lẻ.
   - `Chi tiết CVS`: Danh sách 138 cửa hàng CVS phân công theo từng nhân viên.
   - `Chi tiết BHX`: Báo cáo chi tiết chuỗi Bách Hóa Xanh.
   - `Config`: Cấu hình danh sách cửa hàng, nhân sự và chuỗi.

---

## 👥 2. DANH SÁCH NHÂN SỰ VÀ PHÂN BỔ 138 CỬA HÀNG CVS

> [!IMPORTANT]
> **Quy tắc phân bổ mới nhất:**
> - Nhân viên **Nguyễn Đức Hòa KHÔNG phụ trách chuỗi WinMart+**. 4 cửa hàng WinMart+ trước đây đã được gỡ bỏ hoàn toàn khỏi danh sách của Hòa.
> - Tổng số cửa hàng CVS toàn team hiện tại là **138 cửa hàng** (Trước đây là 142 cửa hàng).
> - Chuỗi WinMart+ (WMP) của cả team hiện có **0 cửa hàng** phân công, nhưng vẫn duy trì cột tiêu đề trong báo cáo để bảo đảm tính chuẩn hóa cấu trúc 7 kênh CVS.

### Bảng phân bổ 138 cửa hàng theo nhân sự:
| STT | Nhân viên | GS25 | FamilyMart | Circle K | 7-Eleven | Hoàng Đức | WinMart+ | Tổng CH CVS | Doanh thu CVS (VNĐ) |
|---|---|:---:|:---:|:---:|:---:|:---:|:---:|:---:|---:|
| 1 | **Trần Thị Cẩm Giang** | 0 | 14 | 14 | 0 | 0 | 0 | **28** | 56,128,420 |
| 2 | **Phan Vũ Đình Duy** | 20 | 0 | 0 | 0 | 0 | 0 | **20** | 45,408,440 |
| 3 | **Trương Thanh Thảo** | 15 | 0 | 0 | 0 | 0 | 0 | **15** | 34,056,330 |
| 4 | **Nguyễn Đức Hòa** | 35 | 0 | 13 | 0 | 0 | 0 | **48** | 114,438,802 |
| 5 | **Kim Hoàng Khang** | 0 | 20 | 0 | 0 | 0 | 0 | **20** | 71,940,350 |
| 6 | **Nguyễn Thị Hướng Dương** | 0 | 0 | 0 | 5 | 2 | 0 | **7** | 37,274,686 |
| | **TỔNG CỘNG TEAM** | **70** | **34** | **27** | **5** | **2** | **0** | **138** | **359,247,028** |

### Số liệu Chỉ tiêu & Thực đạt toàn Team:
- **Chỉ tiêu Toàn Team:** `10,826,339,729 đ`
- **Thực đạt BHX:** `3,551,927,443 đ`
- **Thực đạt CVS (138 CH):** `359,247,028 đ`
- **TỔNG THỰC ĐẠT TOÀN TEAM:** `3,911,174,471 đ`
- **TỶ LỆ HOÀN THÀNH (% ĐẠT):** `36.1%` (Chênh lệch còn thiếu: `6,915,165,258 đ`)

---

## 🔍 3. NGUYÊN LÝ BÓC TÁCH VÀ LỊCH SỬ GIẢI QUYẾT LỆCH DỮ LIỆU

### 1. Tại sao có sự lệch số giữa Ảnh 1 (29,947,350 đ) và Ảnh 2 (84,178,050 đ)?
- **Ảnh 1 (Web app ban đầu):** Hiển thị số fallback mặc định trong `master_data.js` được ghi nhận từ một ảnh chụp nhanh (snapshot) đơn lẻ trước khi tải tệp.
- **Ảnh 2 (Excel report bóc tách từ sheet NPP):** Trích xuất toàn bộ đơn hàng thực tế phát sinh trong sheet `NPP` của chuỗi FamilyMart cho toàn bộ 34-35 cửa hàng trong kỳ. Tổng thực đạt bóc tách chi tiết là `84,178,050 đ`.
- **Kết luận:** Số bóc tách từng dòng chi tiết theo mã sản phẩm (`icode`), số lượng (`iqty`), thành tiền (`iamt`) trong sheet `NPP` của Ảnh 2 mới là số thực tế đầy đủ và chính xác 100%.

### 2. Nguyên nhân chênh lệch ~100 triệu CVS với các dự án/phiên bản cũ:
- Trong phiên bản cũ, chuỗi **GS25** được áp dụng đơn giá trung bình cũ là `873,680 đ/cửa hàng` (Tổng 70 CH = `61,157,600 đ`).
- Trong phiên bản mới nhất, chuỗi **GS25** được tính theo đơn giá điều chỉnh chuẩn hóa là `2,270,422 đ/cửa hàng` (Tổng 70 CH = `158,929,540 đ`).
- Khoản chênh lệch: `158,929,540 - 61,157,600 = 97,771,940 đ` (~98 - 100 triệu đồng). Đây là lý do cốt lõi tạo ra độ lệch khi so sánh với dự án cũ.

### 3. Cơ chế tự động nhận diện SKU mới phát sinh (Dynamic SKU Auto-Discovery):
- **Bản chất trước đây:** Hệ thống chỉ duyệt qua 46 SKU FamilyMart cố định (`FM_CATEGORIES`). Nếu NPP phát sinh SKU mới ngoài 46 SKU này, tuy tiền tổng vẫn được cộng nhưng bảng ma trận SKU không có cột để thể hiện.
- **Cơ chế mới hoàn thiện:**
  - Thuật toán tự động quét toàn bộ `icode` xuất hiện trong sheet `NPP`.
  - Bất kỳ mã `icode` nào chưa có trong 46 SKU chuẩn sẽ được tự động gom vào nhóm: **"SẢN PHẨM MỚI PHÁT SINH / KHÁC"**.
  - Tên sản phẩm được lấy trực tiếp từ cột `INAME`.
  - Tự động sinh thêm cột trong ma trận hiển thị (Web app, Link online `view.html`, và xuất Excel).
  - Đóng gói danh mục động vào tham số `cats` trong URL rút gọn `compactObj` để khi nhân viên mở link chia sẻ, giao diện vẫn hiển thị đầy đủ các SKU mới mà không bị lỗi.

### 4. Công thức chuẩn tính Doanh số Circle K (Trích xuất từ sheet SO):
- **Nguồn dữ liệu:** Sheet `SO` trong file `HNTRINH_KD6-Doanh so Nhap - Ban theo Mien Kenh Nhom hang_sent...`.
- **Mã khách hàng chuỗi Circle K:** `VT4050`, `VT3013`, `VT3014`, `VT3015`.
- **Phân bổ Kho Khô:**
  - Kho Khô định vị tại: `Lô G1-9, Đường N3, N4, D2, KCN Nam Tân Uyên...` với mã `VT4050`. Doanh số Kho Khô toàn hệ thống: `430,620,084 đ`.
  - Số lượng cửa hàng Circle K phát sinh đơn hàng mát trong toàn hệ thống miền: `230 cửa hàng`.
  - Đơn giá chia bình quân Kho Khô cho mỗi cửa hàng có phát sinh đơn:  
    $$tCk = \text{round}(430,620,084 / 230) = 1,872,261 \text{ đ/CH}$$
- **Quy tắc tính doanh số từng cửa hàng:**
  - Cửa hàng có phát sinh đơn hàng mát (`Hàng Mát > 0`):  
    $$\text{Doanh số Thực đạt} = 1,872,261 \text{ đ} + \text{Thành tiền Hàng Mát của CH}$$
  - Cửa hàng không phát sinh đơn (`Hàng Mát = 0`):  
    $$\text{Doanh số Thực đạt} = 0 \text{ đ}$$
- **Chi tiết Thực đạt 27 Cửa hàng Circle K Team Cẩm Giang:**
  - **Lê Thị Thùy Châu (2 CH có đơn):** 7,687,126 đ (`CVS_CIRCLEK.75.4325`: 3,771,969 đ; `CVS_CIRCLEK.75.4326`: 3,915,157 đ).
  - **Não Thị Anh Đào (4 CH - 3 có đơn, 1 không đơn):** 12,650,375 đ (`CVS_CIRCLEK.79.4052`: 6,229,361 đ; `CVS_CIRCLEK.75.4159`: 2,547,861 đ; `CVS_CIRCLEK.75.4160`: 3,873,153 đ; `CVS_CIRCLEK.75.4097`: 0 đ).
  - **Nguyễn Đức Hoà (3 CH - 2 có đơn, 1 không đơn):** 6,324,882 đ (`CVS_CIRCLEK.79.4098`: 2,990,541 đ; `CVS_CIRCLEK.79.4099`: 3,334,341 đ; `CVS_CIRCLEK.79.4100`: 0 đ).
  - **Nguyễn Thị Thanh Thủy (18 CH - 15 có đơn, 3 không đơn):** 42,795,255 đ.
  - **TỔNG CỘNG CIRCLE K TOÀN TEAM:** **69,457,638 đ** (Đồng bộ chuẩn xác 100% giữa sheet `SO`, `Team_CamGiang_Report.xlsx`, `master_data.js`, App Web và APK).

---

## ⚡ 4. NGUYÊN TẮC BẤT BIẾN KHI ĐỒNG BỘ DỰ ÁN (CRITICAL RULES)

1. **Đồng bộ song song Web & Android Assets:**
   - Khi chỉnh sửa bất kỳ nội dung nào trong:
     - `gas_app/index.html`
     - `gas_app/view.html`
     - `gas_app/master_data.js`
   - **BẮT BUỘC** phải sao chép ghi đè ngay lập tức sang:
     - `gas_app/android/app/src/main/assets/www/index.html`
     - `gas_app/android/app/src/main/assets/www/view.html`
     - `gas_app/android/app/src/main/assets/www/master_data.js`
2. **Build APK:**
   - Chạy file `build_apk.bat` trong thư mục `gas_app/` để kích hoạt `gradlew assembleRelease` và tự động copy file ra thư mục gốc thành:
     - `BaoCaoDoanhSo_TeamCamGiang.apk`
     - `BaoCaoThucDat.apk`
3. **Đồng bộ File Excel Báo Cáo:**
   - Mọi thay đổi về phân bổ nhân viên (như việc bỏ WinMart+ của Nguyễn Đức Hòa) phải được cập nhật đồng thời trên:
     - `gas_app/Team_CamGiang_Report.xlsx`
     - `gas_app/MasterData.gs`
     - `gas_app/master_data.js`

---

## 📂 5. DANH MỤC TÀI LIỆU VÀ TỆP TIN QUAN TRỌNG

| Tên tệp tin | Chức năng chính |
|---|---|
| `PROJECT_BRAIN.md` | Tài liệu não bộ tổng hợp toàn bộ tri thức và logic dự án. |
| `GEMINI.md` & `AGENTS.md` | Tệp chỉ dẫn cho AI Agent tự động đọc khi mở workspace trên bất kỳ máy nào. |
| `.agents/rules/` | Tập hợp các luật chuẩn hóa cho Antigravity IDE. |
| `.agents/skills/` | Skill chuyên dụng xử lý trích xuất và đồng bộ báo cáo CVS. |
| `.brain_backup/images/` | Bản lưu trữ các ảnh chụp màn hình bằng chứng số liệu (`media_...png`). |
| `index.html` | Ứng dụng trung tâm xử lý upload file, tính toán và xuất báo cáo. |
| `view.html` | Trang web nhẹ xem báo cáo online chia sẻ qua link / QR code. |
| `master_data.js` | Cấu hình cửa hàng, nhân viên, mục tiêu và dữ liệu dự phòng. |
| `Team_CamGiang_Report.xlsx` | Báo cáo Excel chuẩn định dạng cho toàn team. |
| `build_apk.bat` | Script 1-click build ứng dụng Android APK. |
