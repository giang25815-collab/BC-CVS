@echo off
cd /d "%~dp0"
set "JAVA_HOME=C:\Users\PT COMPUTER\.jdks\ms-21.0.9"
set "PATH=%JAVA_HOME%\bin;%PATH%"
set "ANDROID_HOME=C:\Users\PT COMPUTER\AppData\Local\Android\Sdk"
call "C:\Users\PT COMPUTER\.gradle\wrapper\dists\gradle-8.13-bin\5xuhj0ry160q40clulazy9h7d\gradle-8.13\bin\gradle.bat" assembleRelease

if exist "app\build\outputs\apk\release\app-release.apk" (
    copy /y "app\build\outputs\apk\release\app-release.apk" "..\BaoCaoThucDat.apk"
    copy /y "app\build\outputs\apk\release\app-release.apk" "..\BaoCaoDoanhSo_TeamCamGiang.apk"
    echo ========================================================
    echo  DA TAO THANH CONG APK MOI TAI:
    echo  ..\BaoCaoThucDat.apk
    echo  ..\BaoCaoDoanhSo_TeamCamGiang.apk
    echo ========================================================
)
