# 🤖 ANTIGRAVITY / GEMINI SYSTEM INSTRUCTIONS - TEAM CẨM GIANG SALES REPORT

## 🎯 VAI TRÒ & NGỮ CẢNH HỆ THỐNG
Bạn là AI Kỹ sư phần mềm cao cấp và Chuyên gia Phân tích Dữ liệu Bán lẻ phụ trách hệ thống Báo cáo Doanh số CVS & BHX cho **Team Cẩm Giang**.

Khi thư mục này được mở trên bất kỳ máy tính nào:
1. **Toàn bộ não bộ và ngữ cảnh dự án** được lưu trữ chi tiết tại: [PROJECT_BRAIN.md](file:///d:/v/BC-CVS-main/gas_app/PROJECT_BRAIN.md).
2. Hãy đọc kỹ tài liệu này để nắm bắt toàn bộ kiến trúc, dữ liệu lịch sử và logic nghiệp vụ.

---

## 🔑 CÁC QUY TẮC NGHIỆP VỤ CỐT LÕI (BẮT BUỘC TUÂN THỦ)

1. **Phân bổ nhân sự và cửa hàng CVS (Tổng: 138 Cửa hàng):**
   - **Trần Thị Cẩm Giang:** 28 CH (Circle K: 14, FamilyMart: 14) - Doanh thu: 56,128,420 đ
   - **Phan Vũ Đình Duy:** 20 CH (GS25: 20) - Doanh thu: 45,408,440 đ
   - **Trương Thanh Thảo:** 15 CH (GS25: 15) - Doanh thu: 34,056,330 đ
   - **Nguyễn Đức Hòa:** 48 CH (GS25: 35, Circle K: 13, **WinMart+: 0**) - Doanh thu CVS: 114,438,802 đ
   - **Kim Hoàng Khang:** 20 CH (FamilyMart: 20) - Doanh thu: 71,940,350 đ
   - **Nguyễn Thị Hướng Dương:** 7 CH (7-Eleven: 5, Hoàng Đức: 2) - Doanh thu: 37,274,686 đ
   - **LƯU Ý:** Nhân viên Nguyễn Đức Hòa **KHÔNG phụ trách WinMart+**. 4 cửa hàng WinMart+ đã được xóa bỏ khỏi danh sách của Hòa. Cả team hiện có 0 cửa hàng WinMart+.

2. **Số liệu Toàn Team:**
   - Chỉ tiêu toàn team: `10,826,339,729 đ`
   - Thực đạt BHX: `3,551,927,443 đ`
   - Thực đạt CVS (138 CH): `359,247,028 đ`
   - Tổng thực đạt: `3,911,174,471 đ` (Đạt 36.1%)

3. **Cơ chế SKU Động (Dynamic SKU Auto-Discovery):**
   - Dữ liệu sheet `NPP` FamilyMart: Nếu phát sinh mã SKU (`icode`) mới ngoài danh mục 46 SKU tĩnh, hệ thống tự động nhận diện, gom vào nhóm `"SẢN PHẨM MỚI PHÁT SINH / KHÁC"` và hiển thị cột tương ứng trong bảng ma trận, không bị mất doanh thu và không bị lỗi hiển thị.

4. **Nguyên tắc Đồng bộ Tài sản (Sync Rules):**
   - Bất kỳ thay đổi nào trong `gas_app/index.html`, `gas_app/view.html`, `gas_app/master_data.js` **PHẢI** được sao chép ghi đè vào `gas_app/android/app/src/main/assets/www/`.
   - File Excel `gas_app/Team_CamGiang_Report.xlsx` và Google Apps Script `MasterData.gs` phải luôn đồng bộ danh sách 138 cửa hàng.
   - Khi build ứng dụng Android, chạy `build_apk.bat` để tạo ra `BaoCaoDoanhSo_TeamCamGiang.apk` và `BaoCaoThucDat.apk`.

5. **Công thức chuẩn tính Circle K (Sheet SO):**
   - Doanh số Circle K lấy từ sheet `SO` trong file `HNTRINH_KD6...`.
   - Kho Khô (`VT4050` Nam Tân Uyên = `430,620,084 đ`) chia đều cho số cửa hàng phát sinh đơn hàng mát toàn hệ thống (`230 cửa hàng`) = `1,872,261 đ/CH`.
   - Mỗi cửa hàng có phát sinh đơn trong team nhận: `1,872,261 đ` + `Thành tiền hàng mát của cửa hàng đó`. Cửa hàng không có đơn nhận `0 đ`.
   - Tổng Circle K của toàn team: **`69,457,638 đ`** (22 CH có đơn, 5 CH không có đơn).
