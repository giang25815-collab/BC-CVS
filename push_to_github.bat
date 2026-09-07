@echo off
chcp 65001 >nul
title ĐẨY CODE LÊN GITHUB BC-CVS
echo ======================================================================
echo   ĐANG TIẾN HÀNH ĐẨY CODE MỚI LÊN GITHUB: giang25815-collab/BC-CVS
echo ======================================================================
echo.
"%LOCALAPPDATA%\MinGit\cmd\git.exe" push -u origin main
echo.
echo ======================================================================
echo   HOÀN TẤT! NHẤN PHÍM BẤT KỲ ĐỂ ĐÓNG CỬA SỔ...
echo ======================================================================
pause >nul
