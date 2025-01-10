@echo off
:: Kiểm tra quyền Administrator
openfiles >nul 2>&1
if %errorlevel% NEQ 0 (
    echo This script requires Administrator privileges. Please run as Administrator.
    pause
    exit /b
)

:: Xóa Task trong Task Scheduler
echo Deleting Task "TimeKeeperTask"...
schtasks /Delete /TN "TimeKeeperTask" /F >nul 2>&1

:: Kiểm tra và xóa file TimeKeeper.vbs
if exist "C:\TimeKeeper\TimeKeeper.vbs" (
    echo Deleting file "TimeKeeper.vbs"...
    del /f "C:\TimeKeeper\TimeKeeper.vbs"
)

:: Kiểm tra và xóa thư mục C:\TimeKeeper nếu trống
if exist "C:\TimeKeeper" (
    echo Deleting directory "C:\TimeKeeper"...
    rmdir /s /q "C:\TimeKeeper"
)

echo Task and files have been successfully deleted.
pause
