@echo off
chcp 65001 >nul
title DAY TOAN BO CODE MOI LEN GITHUB VA VERCEL
color 0B
echo ======================================================================
echo   ĐANG TIẾN HÀNH ĐẨY CODE LÊN GITHUB: giang25815-collab/BC-CVS
echo   Sau khi đẩy xong, Vercel sẽ tự động build và cập nhật bcgiang.vercel.app
echo ======================================================================
echo.

where git >nul 2>nul
if %errorlevel% equ 0 (
    git push -u origin main
) else (
    "%LOCALAPPDATA%\MinGit\cmd\git.exe" push -u origin main
)

echo.
if %errorlevel% equ 0 (
    color 0A
    echo ======================================================================
    echo   [THÀNH CÔNG] Đã đẩy code lên GitHub thành công!
    echo   Vercel đang tự động cập nhật web tại https://bcgiang.vercel.app
    echo ======================================================================
) else (
    color 0C
    echo ======================================================================
    echo   [LỖI] Chưa đẩy được lên GitHub. Vui lòng kiểm tra đăng nhập GitHub.
    echo ======================================================================
)
echo.
echo Nhấn phím bất kỳ để đóng cửa sổ này...
pause >nul
