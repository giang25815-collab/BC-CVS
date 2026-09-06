# 📊 Team CVS + BHX Sales Tracker

**Team Lead:** Trần Thị Cẩm Giang
**Version:** 4.0.0

---

## 🚀 Quick Start

```bash
# 1. Cài dependency
pip install openpyxl

# 2. Chạy báo cáo
python3 team_leader_report_v4.py
```

Kết quả: `Team_CamGiang_Report.xlsx`

---

## 📂 Cấu trúc bộ file

```
team-cvs-bhx-tracker/
├── team-cvs-bhx-tracker.md      # Skill definition cho Antigravity
├── team_leader_report_v4.py     # Script Python chính
├── team_config.csv              # Cấu hình tháng, target hệ thống
├── team_employees.csv           # Danh sách NV team (8 NV)
├── team_cvs_stores.csv          # Chi tiết CVS 35 cửa hàng
└── Team_CamGiang_Report.xlsx    # Output báo cáo
```

---

## 📝 Workflow hàng ngày (5 phút)

### Bước 1: Cập nhật config
Mở `team_config.csv`, sửa các dòng:
- `total_actual_bhx` → tổng BHX mới nhất từ hệ thống
- `time_percentage` → % ngày đã qua trong tháng

### Bước 2: Cập nhật CVS
Mở `team_cvs_stores.csv`, sửa cột `actual` cho các cửa hàng CVS

### Bước 3: Chạy script
```bash
python3 team_leader_report_v4.py
```

### Bước 4: Kiểm tra file Excel output

---

## 🧮 Công thức tính

```
BHX_NV = (Tổng BHX ÷ 232) × Số CH BHX phụ trách   [cố định]
CVS_NV = GS25 + 7-Eleven + FamilyMart + Circle K   [thực tế]
Total  = BHX_NV + CVS_NV
% Đạt  = Total ÷ Target × 100%

So sánh với % Timegone:
- % Đạt ≥ % Timegone → 🟢 VƯỢT tiến độ
- % Đạt < % Timegone → 🔴 CHẬM tiến độ
```

---

## 📊 File Excel Output

5 sheets:
1. **Dashboard** — KPI + Xếp hạng NV
2. **Tiến độ Team** — Bảng theo dõi chính
3. **Chi tiết CVS** — Từng cửa hàng CVS
4. **Data BHX** — Phân bổ BHX
5. **Config** — Cấu hình audit

---

## 💬 Câu lệnh Antigravity

```
/team-cvs-bhx-tracker run
/team-cvs-bhx-tracker update --month 8 --time-percentage 65
/team-cvs-bhx-tracker view --month 7
```
