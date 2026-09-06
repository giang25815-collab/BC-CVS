---
name: team-cvs-bhx-tracker
display_name: Team CVS + BHX Sales Tracker
description: Skill theo dõi doanh số hàng tháng cho Team Lead bán lẻ, tự động chia BHX (5 Hubs/232 CH), GS25 (2 Kho DC * 20.46% / 72 CH), 7-Eleven (2 Kho DC * 3.62% / 5 CH), Circle K (Kho khô/233 CH + Doanh số từng CH) và cộng dồn CVS để tính % Đạt vs % Timegone.
version: 4.4.0
author: Trần Thị Cẩm Giang (Team Lead)
created: 2026-08-20
category: sales-analytics
tags:
  - retail
  - sales-tracking
  - bhx
  - cvs
  - circle-k
  - gs25
  - 7-eleven
  - team-management
  - excel-automation
runtime:
  language: python
  version: ">=3.8"
  dependencies:
    - openpyxl>=3.1.0
    - pyxlsb>=1.0.10
entry_point: team_leader_report_v4.py
---

# 🎯 SKILL: Team CVS + BHX Sales Tracker

## 📌 Mục đích
Skill này giúp **Team Lead** theo dõi doanh số hàng tháng của cả team bán lẻ, hỗ trợ:
- ✅ Tự động **tính doanh số BHX (Bách Hóa Xanh)** từ tổng 5 Hubs trong file Siêu Thị chia cho 232 cửa hàng hệ thống.
- ✅ Tự động **tính doanh số GS25** từ tổng 2 kho DC (Long Hậu + Bình Điền) × 20.46% chia cho 72 cửa hàng.
- ✅ Tự động **tính doanh số 7-Eleven (7E)** từ (tổng 2 kho DC: BW Tân Phú Trung + Lô II-3 KCN Tân Bình × 3.62%) chia cho 5 cửa hàng phụ trách.
- ✅ Tự động **tính doanh số Circle K (CK)**: Lấy Doanh số Kho khô Tân Uyên ÷ 233 CH ($= T$) + Đơn hàng thực tế (`SumOfAMOUNT`) của từng CH trong tháng.
- ✅ Tự động **tổng hợp doanh số CVS** chi tiết từng cửa hàng (GS25, 7-Eleven, FamilyMart, Circle K, Hoàng Đức).
- ✅ Tính **% Đạt** = (BHX + CVS) ÷ Target × 100%.
- ✅ **So sánh với % Timegone** tự động theo thời gian thực để biết team đang vượt hay chậm tiến độ.
- ✅ Xuất báo cáo Excel 6 sheets chuyên nghiệp có Dashboard, xếp hạng NV, chi tiết CVS và Matrix SKU FamilyMart.

---

## 🧮 Công thức cốt lõi & Quy tắc phân bổ

```
┌─────────────────────────────────────────────────────────────────────────────┐
│ 1. BÁCH HÓA XANH (BHX) — TỔNG 5 HUBS ÷ 232 CỬA HÀNG                         │
│    Tổng 5 Hubs = Châu Pha + Bùi Văn Hòa + 2 Hubs + Long Khánh               │
│    BHX/CH = Tổng 5 Hubs ÷ 232 CH (~ 57.827.395 VNĐ/CH)                      │
│    BHX_NV = BHX/CH × Số CH NV phụ trách                                     │
│                                                                             │
│ 2. GS25 — (TỔNG 2 KHO DC × 20,46%) ÷ 72 CỬA HÀNG                            │
│    Tổng 2 DC = DC Long Hậu (H.04) + DC Bình Điền (IIIB2)                    │
│    GS25/CH = (Tổng 2 DC × 20,46%) ÷ 72 CH (~ 6.315.784 VNĐ/CH)              │
│    Áp dụng đồng đều cho tất cả CH GS25 trong danh sách team                 │
│                                                                             │
│ 3. 7-ELEVEN (7E) — (TỔNG 2 KHO DC × 3,62%) ÷ 5 CỬA HÀNG                     │
│    Tổng 2 DC = DC BW Tân Phú Trung (Lô D2) + DC KCN Tân Bình (Lô II-3)      │
│    7E/CH = (Tổng 2 DC × 3,62%) ÷ 5 CH (~ 8.874.940 VNĐ/CH)                  │
│    Phụ trách: Não Thị Anh Đào (3 CH), Nguyễn Đức Hoà (2 CH)                 │
│                                                                             │
│ 4. CIRCLE K (CK) — KHO KHÔ TÂN UYÊN ÷ 233 CH + DOANH SỐ TỪNG CỬA HÀNG      │
│    T = Doanh số Kho khô Nam Tân Uyên ÷ 233 CH (~ 5.520.869 VNĐ/CH)          │
│    Doanh số CH CK = T + SumOfAMOUNT (đơn hàng thực tế của CH trong SO)      │
│                                                                             │
│ 5. TỔNG CVS & TỔNG THỰC HIỆN                                                │
│    CVS_NV = GS25 + 7-Eleven + FamilyMart + Circle K + Hoàng Đức             │
│    Total_NV = BHX_NV + CVS_NV                                               │
│                                                                             │
│ 6. % ĐẠT & TIẾN ĐỘ                                                          │
│    Percent_NV = Total_NV ÷ Target_NV × 100%                                 │
│    Nếu %Đạt ≥ %Timegone → 🟢 VƯỢT tiến độ                                   │
│    Nếu %Đạt < %Timegone → 🔴 CHẬM tiến độ                                   │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## 📂 Danh sách các file cấu hình & dữ liệu nguồn

### 1. File nguồn tự động (đặt trong thư mục tải về hoặc cùng cấp):
- `HNTRINH_ST_KD6...xlsb`: Nguồn dữ liệu Siêu Thị (Bách Hóa Xanh 5 Hubs).
- `HNTRINH_KD6...xlsb`: Nguồn dữ liệu CVS & NPP (GS25 2 DCs, 7-Eleven 2 DCs, Circle K Kho khô + Đơn hàng CH, FamilyMart NPP).
- `DOANH SỐ T8.xlsx`: Danh sách phân công cửa hàng, Target và danh mục SKU FamilyMart.

### 2. Các file CSV trung gian (tự động cập nhật qua `update_data_and_config.py`):
- `team_config.csv`: Cấu hình tổng BHX, số CH BHX (232), % Timegone realtime.
- `team_employees.csv`: Danh sách 10 NV, số CH BHX phụ trách, Target tháng.
- `team_cvs_stores.csv`: Chi tiết 138 cửa hàng CVS (70 GS25, 5 7-Eleven, 34 FamilyMart, 27 Circle K, 2 Hoàng Đức) kèm doanh số thực tế đã tính toán.

---

## 📊 Output — File Excel 6 sheets

1. **Dashboard** 🎯: KPI Cards (Target, BHX, CVS, Tổng TH, % Đạt), Status bar VƯỢT/CHẬM, Bảng xếp hạng NV.
2. **Tiến độ Team** ⭐: Bảng tổng hợp STT, NV, Target, BHX, GS25, 7E, FM, CK, Hoàng Đức, Tổng TH, % Đạt, Color Scale.
3. **Chi tiết CVS** 🏪: Danh sách chi tiết từng cửa hàng CVS phân nhóm theo từng NV kèm Subtotal.
4. **Data BHX** 📦: Thông số 5 Hubs BHX, đơn giá/CH (57.827.395 VNĐ) và phân bổ cho từng NV.
5. **Config** ⚙️: Toàn bộ thông số cấu hình hệ thống để kiểm tra audit.
6. **Chi tiết SKU FamilyMart** 📋: Matrix 35 cửa hàng FamilyMart của tất cả nhân viên (Kim Hoàng Khang, Não Thị Anh Đào, Nguyễn Đức Hoà, Nguyễn Thị Thanh Thủy) và 46 SKU thực tế đặt hàng kèm tính năng AutoFilter lọc theo từng nhân viên.

---

## 🚀 Quy trình chạy báo cáo tự động (1 bước)

Chỉ cần chạy lệnh cập nhật và xuất báo cáo:
```bash
python update_data_and_config.py
python team_leader_report_v4.py
```
File Excel `Team_CamGiang_Report.xlsx` sẽ được tạo mới với đầy đủ số liệu chính xác 100%.
