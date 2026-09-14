# AGENTS.md - CONTEXT & INSTRUCTIONS FOR AI AGENTS

This workspace contains the CVS & BHX Sales Reporting Application for **Team Cẩm Giang**.

## Critical Project Brain Reference
For complete business logic, store distribution, SKU discovery engine, and historical discrepancy explanations, refer to:
- [PROJECT_BRAIN.md](file:///d:/v/BC-CVS-main/gas_app/PROJECT_BRAIN.md)
- [GEMINI.md](file:///d:/v/BC-CVS-main/gas_app/GEMINI.md)

## Key Technical & Business Truths
- Total CVS stores: **138** across 6 representatives.
- Representative **Nguyễn Đức Hòa** does NOT manage WinMart+ (4 stores removed; WMP count = 0).
- Team CVS Target: 10,826,339,729 đ. Team Total Attainment: 3,911,174,471 đ (36.1%).
- FamilyMart NPP sheet parsing: Dynamic auto-discovery of new SKUs outside static 46 list under category "SẢN PHẨM MỚI PHÁT SINH / KHÁC".
- Sync constraint: Any edits to `index.html`, `view.html`, `master_data.js` must be copied to `android/app/src/main/assets/www/`.
- Release APKs: Built via `build_apk.bat` -> `BaoCaoDoanhSo_TeamCamGiang.apk` and `BaoCaoThucDat.apk`.
- Circle K formula (sheet SO): Kho Khô (VT4050 / 230 CH có đơn = 1,872,261 đ/CH) + Hàng Mát của từng cửa hàng có đơn. Cửa hàng không đơn = 0 đ. Tổng Circle K: 69,457,638 đ.
