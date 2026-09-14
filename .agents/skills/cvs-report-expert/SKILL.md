---
name: cvs-report-expert
description: Chuyên gia xử lý báo cáo doanh số CVS & BHX Team Cẩm Giang, bóc tách dữ liệu sheet NPP FamilyMart, đồng bộ báo cáo Excel và build APK.
---

# CVS Report Expert Skill

Kỹ năng này hướng dẫn thực hiện các tác vụ thường xuyên trên hệ thống Báo cáo Doanh số Team Cẩm Giang.

## 1. Kiểm tra và đồng bộ phân bổ cửa hàng CVS
- Tổng số cửa hàng chuẩn: **138 cửa hàng**.
- Khi có thay đổi nhân sự phụ trách (ví dụ gỡ bỏ WinMart+ khỏi Nguyễn Đức Hòa hoặc phân công cho người khác):
  1. Cập nhật mảng cửa hàng trong `master_data.js` và `MasterData.gs`.
  2. Cập nhật hàm `getInitialStores()` và các badge hiển thị trong `index.html` và `view.html`.
  3. Cập nhật sheet `Tiến độ Team`, `Chi tiết CVS` và `Config` trong `Team_CamGiang_Report.xlsx`.
  4. Đồng bộ ngay sang thư mục `android/app/src/main/assets/www/`.

## 2. Bóc tách dữ liệu NPP FamilyMart (Dynamic SKU)
- Khi đọc file Excel chứa sheet `NPP`:
  - Quét qua từng dòng dữ liệu từ dòng 10 trở đi.
  - Lấy mã cửa hàng (`store_id` hoặc `CUST_ID`), mã sản phẩm (`icode`), tên sản phẩm (`iname`), số lượng (`iqty`), thành tiền (`iamt`).
  - Nếu `icode` chưa có trong `FM_CATEGORIES`, tự động gom vào category `"SẢN PHẨM MỚI PHÁT SINH / KHÁC"`.
  - Không bao giờ được bỏ qua doanh thu của SKU mới.

## 3. Build ứng dụng Android APK
- Để đóng gói file APK mới nhất:
  ```cmd
  cd gas_app\android
  .\gradlew assembleRelease
  ```
  hoặc chạy trực tiếp script:
  ```cmd
  build_apk.bat
  ```
- File APK đầu ra nằm tại thư mục gốc:
  - `BaoCaoDoanhSo_TeamCamGiang.apk`
  - `BaoCaoThucDat.apk`
