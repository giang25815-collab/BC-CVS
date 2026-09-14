# RULES: TEAM CẨM GIANG SALES REPORT CONTEXT & CONSTRAINTS

## Business Constraints
1. **Total CVS Stores**: Exactly 138 stores across 6 reps.
2. **Nguyễn Đức Hòa**: Responsible for 48 CVS stores (35 GS25, 13 Circle K, 0 WinMart+). Under no circumstances should WinMart+ stores be assigned to Nguyễn Đức Hòa unless explicitly requested by the user.
3. **Dynamic SKU Discovery**: FamilyMart NPP sheet parsing must dynamically discover all new SKUs and include them in table rendering, online links (`cats`), and Excel export.

## Technical Constraints
1. **Android Assets Synchronization**: Any edit to `gas_app/index.html`, `gas_app/view.html`, or `gas_app/master_data.js` must be immediately mirrored into `gas_app/android/app/src/main/assets/www/`.
2. **Dual APK Generation**: `build_apk.bat` must generate both `BaoCaoDoanhSo_TeamCamGiang.apk` and `BaoCaoThucDat.apk`.
3. **Clean Git Tree**: Keep commits descriptive with Vietnamese commit messages conforming to conventional commits (`feat:`, `fix:`).
