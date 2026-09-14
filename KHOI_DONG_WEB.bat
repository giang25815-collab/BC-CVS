@echo off
chcp 65001 >nul
title Khởi động Hệ thống Báo cáo Doanh số - Team Cẩm Giang
echo =====================================================================
echo  HỆ THỐNG BÁO CÁO DOANH SỐ CVS & BHX - TEAM CẨM GIANG
echo =====================================================================
echo [1] Đang kiểm tra môi trường...
echo.

where python >nul 2>nul
if %errorlevel% equ 0 (
    echo [OK] Đã tìm thấy Python. Khởi động Web Server tại cổng 8080...
    start "" http://localhost:8080/index.html
    python -m http.server 8080
    goto end
)

where npx >nul 2>nul
if %errorlevel% equ 0 (
    echo [OK] Đã tìm thấy Node/NPX. Khởi động serve...
    start "" http://localhost:3000/index.html
    npx -y serve -p 3000 .
    goto end
)

echo [INFO] Mở trực tiếp file index.html trên trình duyệt mặc định...
start "" index.html

:end
pause
